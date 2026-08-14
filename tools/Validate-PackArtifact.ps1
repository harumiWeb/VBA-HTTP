[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$ArtifactPath,
    [string]$ManifestPath = "",
    [string]$RootPath = ""
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$projectRoot = if ([string]::IsNullOrWhiteSpace($RootPath)) {
    [IO.Path]::GetFullPath((Join-Path $PSScriptRoot ".."))
}
else {
    [IO.Path]::GetFullPath($RootPath)
}
$resolvedArtifact = if ([IO.Path]::IsPathRooted($ArtifactPath)) { [IO.Path]::GetFullPath($ArtifactPath) } else { [IO.Path]::GetFullPath((Join-Path $projectRoot $ArtifactPath)) }
$resolvedManifest = if ([string]::IsNullOrWhiteSpace($ManifestPath)) { "$resolvedArtifact.pack.json" } elseif ([IO.Path]::IsPathRooted($ManifestPath)) { [IO.Path]::GetFullPath($ManifestPath) } else { [IO.Path]::GetFullPath((Join-Path $projectRoot $ManifestPath)) }

function Get-RequiredProperty($Object, [string]$Name) {
    if ($null -eq $Object) { throw "Pack manifest record is null while reading '$Name'." }
    $property = $Object.PSObject.Properties[$Name]
    if ($null -eq $property) { throw "Pack manifest is missing '$Name'." }
    return $property.Value
}

function Assert-AllowedPropertySet($Object, [string[]]$Allowed, [string]$Label) {
    if ($null -eq $Object) { throw "$Label is null." }
    $actual = @($Object.PSObject.Properties.Name | Sort-Object)
    $expected = @($Allowed | Sort-Object)
    if ($null -ne (Compare-Object $expected $actual)) { throw "$Label contains an unexpected or missing property." }
}

function Get-Sha256Hex([string]$Path) {
    $sha = [Security.Cryptography.SHA256]::Create()
    $stream = [IO.File]::OpenRead($Path)
    try {
        return ([BitConverter]::ToString($sha.ComputeHash($stream))).Replace('-', '').ToLowerInvariant()
    }
    finally {
        $stream.Dispose()
        $sha.Dispose()
    }
}

function ConvertTo-RepoPath([string]$Path) {
    return $Path.Replace('\', '/')
}

function Resolve-RepoPath([string]$Path) {
    $normalized = ConvertTo-RepoPath $Path
    if ($normalized.StartsWith('/') -or $normalized -match '^[A-Za-z]:') {
        throw "Pack manifest source path must be relative: $Path"
    }
    if ($normalized -match '(^|/)\.\.?(/|$)' -or $normalized.Contains('//')) {
        throw "Pack manifest source path contains traversal: $Path"
    }
    $resolved = [IO.Path]::GetFullPath((Join-Path $projectRoot $normalized.Replace('/', [IO.Path]::DirectorySeparatorChar)))
    $root = $projectRoot.TrimEnd('\', '/') + [IO.Path]::DirectorySeparatorChar
    if (-not $resolved.StartsWith($root, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Pack manifest source path escapes the project root: $Path"
    }
    return $resolved
}

if (-not (Test-Path -LiteralPath $resolvedArtifact -PathType Leaf)) { throw "Pack artifact does not exist: $resolvedArtifact" }
if (-not (Test-Path -LiteralPath $resolvedManifest -PathType Leaf)) { throw "Pack manifest does not exist: $resolvedManifest" }
$policy = Get-Content -LiteralPath (Join-Path $projectRoot 'tools/build-component-policy.json') -Raw | ConvertFrom-Json
$manifest = Get-Content -LiteralPath $resolvedManifest -Raw | ConvertFrom-Json
Assert-AllowedPropertySet $manifest @('schema_version', 'package_id', 'release_tag', 'source_revision', 'target', 'artifact', 'template', 'template_sha256', 'xlflow_version', 'pack_backend', 'experimental', 'vbe_validation', 'included_components', 'excluded_components', 'module_counts') 'Pack manifest'

if ([int](Get-RequiredProperty $manifest 'schema_version') -ne 1 -or
    [string](Get-RequiredProperty $manifest 'package_id') -ne 'VBA-HTTP' -or
    [string](Get-RequiredProperty $manifest 'target') -ne 'windows-x64-office' -or
    [string](Get-RequiredProperty $manifest 'pack_backend') -ne 'pure-go' -or
    [bool](Get-RequiredProperty $manifest 'experimental') -ne $true -or
    [string](Get-RequiredProperty $manifest 'vbe_validation') -ne 'not_performed') {
    throw 'Pack manifest identity or provenance is invalid.'
}
$manifestRevision = [string](Get-RequiredProperty $manifest 'source_revision')
$manifestTag = [string](Get-RequiredProperty $manifest 'release_tag')
if ($manifestTag -eq 'local') {
    if ($manifestRevision -notmatch '^(working-tree|[0-9a-fA-F]{40,64}(-dirty)?)$') {
        throw 'Local pack manifest source revision is invalid.'
    }
}
elseif ($manifestRevision -notmatch '^[0-9a-fA-F]{40,64}$') {
    throw 'Release pack manifest source revision is not a full hexadecimal SHA.'
}
if ([string](Get-RequiredProperty $manifest 'xlflow_version') -notmatch '^[0-9A-Za-z][0-9A-Za-z._-]*$') {
    throw 'Pack manifest xlflow version is invalid.'
}
if ([string](Get-RequiredProperty $manifest 'artifact') -ne [IO.Path]::GetFileName($resolvedArtifact)) {
    throw 'Pack manifest artifact name does not match the supplied artifact.'
}

$included = @((Get-RequiredProperty $manifest 'included_components'))
$excluded = @((Get-RequiredProperty $manifest 'excluded_components'))
$includedNames = @($included | ForEach-Object { [string](Get-RequiredProperty $_ 'name') } | Sort-Object)
$excludedNames = @($excluded | ForEach-Object { [string](Get-RequiredProperty $_ 'name') } | Sort-Object)
$expectedIncluded = @($policy.included | ForEach-Object { [string]$_ } | Sort-Object)
$expectedExcluded = @($policy.excluded | ForEach-Object { [string]$_ } | Sort-Object)
if ($null -ne (Compare-Object $expectedIncluded $includedNames)) { throw 'Pack manifest included components differ from the production policy.' }
if ($null -ne (Compare-Object $expectedExcluded $excludedNames)) { throw 'Pack manifest excluded components differ from the production policy.' }

$seen = @{}
foreach ($component in $included) {
    Assert-AllowedPropertySet $component @('name', 'type', 'source_path', 'sha256') 'Included pack component'
    $name = [string](Get-RequiredProperty $component 'name')
    if ($seen.ContainsKey($name)) { throw "Pack manifest contains duplicate component: $name" }
    $seen[$name] = $true
    $sourcePath = ConvertTo-RepoPath ([string](Get-RequiredProperty $component 'source_path'))
    if ($sourcePath -match '(?i)(^|/)(Tests|Benchmarks|Xlflow|Dev)(/|$)') { throw "Pack manifest includes a development source path: $sourcePath" }
    $sourceFile = Resolve-RepoPath $sourcePath
    if (-not (Test-Path -LiteralPath $sourceFile -PathType Leaf)) { throw "Pack source file is missing: $sourcePath" }
    if ([IO.Path]::GetFileNameWithoutExtension($sourceFile) -ne $name) { throw "Pack source name does not match component: $name" }
    $expectedType = if ($sourcePath -match '(?i)^src/classes/') { 'class' } elseif ($sourcePath -match '(?i)^src/workbook/') { 'document' } elseif ($sourcePath -match '(?i)^src/forms/') { 'form' } else { 'standard' }
    if ([string](Get-RequiredProperty $component 'type') -ne $expectedType) { throw "Pack component type does not match source path: $name" }
    $hash = [string](Get-RequiredProperty $component 'sha256')
    if ($hash -notmatch '^[0-9a-f]{64}$' -or $hash -ne (Get-Sha256Hex $sourceFile)) { throw "Pack source hash mismatch: $name" }
}

$excludedSeen = @{}
foreach ($component in $excluded) {
    Assert-AllowedPropertySet $component @('name', 'type', 'source_path') 'Excluded pack component'
    $name = [string](Get-RequiredProperty $component 'name')
    if ($excludedSeen.ContainsKey($name)) { throw "Pack manifest contains duplicate excluded component: $name" }
    $excludedSeen[$name] = $true
    $sourcePath = ConvertTo-RepoPath ([string](Get-RequiredProperty $component 'source_path'))
    $sourceFile = Resolve-RepoPath $sourcePath
    if ([IO.Path]::GetFileNameWithoutExtension($sourceFile) -ne $name) { throw "Excluded source name does not match component: $name" }
}

$counts = Get-RequiredProperty $manifest 'module_counts'
foreach ($key in @('class', 'standard', 'document', 'form')) {
    if ($null -eq $counts.PSObject.Properties[$key]) { throw "Pack manifest module_counts is missing '$key'." }
}
$actualCounts = [ordered]@{
    class = @($included | Where-Object { $_.type -eq 'class' }).Count
    standard = @($included | Where-Object { $_.type -eq 'standard' }).Count
    document = @($included | Where-Object { $_.type -eq 'document' }).Count
    form = @($included | Where-Object { $_.type -eq 'form' }).Count
}
foreach ($key in @('class', 'standard', 'document', 'form')) {
    if ([int]$counts.$key -ne [int]$actualCounts[$key]) { throw "Pack manifest module count is inconsistent for $key." }
}

Write-Output "Pack artifact valid: $($included.Count) production components, pure-Go experimental pack, VBE validation not performed."
