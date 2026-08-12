[CmdletBinding()]
param(
    [string]$ArtifactPath = "build/Release/VBA-HTTP.xlsm",
    [string]$PolicyPath = "tools/build-component-policy.json",
    [string]$ReportPath = "",
    [string]$RootPath = "",
    [string]$RiskRegisterPath = ""
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$scriptRoot = Split-Path -Parent $PSScriptRoot
$projectRoot = if ([string]::IsNullOrWhiteSpace($RootPath)) {
    [IO.Path]::GetFullPath($scriptRoot)
}
else {
    [IO.Path]::GetFullPath($RootPath)
}

if ([string]::IsNullOrWhiteSpace($RiskRegisterPath) -and [string]::IsNullOrWhiteSpace($RootPath)) {
    $RiskRegisterPath = "docs/security/risk-register.json"
}
if (-not [string]::IsNullOrWhiteSpace($RiskRegisterPath)) {
    & (Join-Path $PSScriptRoot "Validate-SecurityRiskRegister.ps1") -Path $RiskRegisterPath | Out-Null
}

function Resolve-InputPath([string]$Path) {
    if ([string]::IsNullOrWhiteSpace($Path)) {
        throw "A security validation path cannot be empty."
    }
    if ([IO.Path]::IsPathRooted($Path)) {
        return [IO.Path]::GetFullPath($Path)
    }
    return [IO.Path]::GetFullPath((Join-Path $projectRoot $Path))
}

function Normalize-RepoPath([string]$Path) {
    if ([string]::IsNullOrWhiteSpace($Path)) {
        throw "A manifest path cannot be empty."
    }
    $normalized = $Path.Replace('\', '/').Trim()
    if ($normalized.StartsWith('/') -or $normalized -match '^[A-Za-z]:') {
        throw "Manifest paths must be repository-relative: $Path"
    }
    while ($normalized.StartsWith('./', [StringComparison]::Ordinal)) {
        $normalized = $normalized.Substring(2)
    }
    if ($normalized -match '(^|/)\.\.?(/|$)' -or $normalized.Contains('//')) {
        throw "Manifest path contains traversal or an empty segment: $Path"
    }
    return $normalized
}

function Get-RepoRelativePath([string]$Path) {
    $resolved = [IO.Path]::GetFullPath($Path)
    $rootWithSeparator = $projectRoot.TrimEnd('\', '/') + [IO.Path]::DirectorySeparatorChar
    if (-not $resolved.StartsWith($rootWithSeparator, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Path is outside the validation root: $Path"
    }
    return Normalize-RepoPath $resolved.Substring($rootWithSeparator.Length)
}

function Require-Property($Object, [string]$Name) {
    if ($null -eq $Object) {
        throw "Manifest record is null while reading '$Name'."
    }
    $property = $Object.PSObject.Properties[$Name]
    if ($null -eq $property) {
        throw "Manifest record is missing '$Name'."
    }
    return $property.Value
}

function Get-StringList($Value) {
    if ($null -eq $Value) { return @() }
    return @($Value | ForEach-Object { [string]$_ })
}

function Assert-ExactSet([string[]]$Expected, [string[]]$Actual, [string]$Description) {
    $expectedSorted = @($Expected | Sort-Object)
    $actualSorted = @($Actual | Sort-Object)
    if ($null -ne (Compare-Object -ReferenceObject $expectedSorted -DifferenceObject $actualSorted)) {
        throw "$Description differs from the checked-in policy."
    }
}

function Assert-Unique([string[]]$Values, [string]$Description) {
    $duplicate = @($Values | Group-Object | Where-Object { $_.Count -gt 1 })
    if ($duplicate.Count -ne 0) {
        throw "$Description contains duplicate component names or paths."
    }
}

function Assert-SourcePath([string]$Path, [bool]$AllowDevelopmentSegment, [string]$Description) {
    $normalized = Normalize-RepoPath $Path
    $allowedRoot = $normalized.StartsWith('src/classes/', [StringComparison]::OrdinalIgnoreCase) -or
        $normalized.StartsWith('src/modules/', [StringComparison]::OrdinalIgnoreCase) -or
        $normalized.StartsWith('src/workbook/', [StringComparison]::OrdinalIgnoreCase)
    if (-not $allowedRoot) {
        throw "$Description is outside the approved VBA source roots: $Path"
    }

    $segments = @($normalized.ToLowerInvariant().Split('/'))
    $developmentSegments = @('tests', 'benchmarks', 'xlflow', 'dev')
    $hasDevelopmentSegment = $false
    foreach ($segment in $segments) {
        if ($developmentSegments -contains $segment) {
            $hasDevelopmentSegment = $true
            break
        }
    }
    if (-not $AllowDevelopmentSegment -and $hasDevelopmentSegment) {
        throw "$Description points at a development-only source path: $Path"
    }
    if ($AllowDevelopmentSegment -and -not $hasDevelopmentSegment) {
        throw "$Description is excluded but is not under a development-only path: $Path"
    }
    return $normalized
}

function Assert-ComponentRecord($Component, [bool]$Excluded, [hashtable]$SeenNames, [hashtable]$SeenPaths) {
    $name = [string](Require-Property $Component 'name')
    if ($name -notmatch '^[A-Za-z][A-Za-z0-9_]*$') {
        throw "Component name is not a safe VBA identifier: $name"
    }
    if ($SeenNames.ContainsKey($name)) {
        throw "Component name is duplicated: $name"
    }
    $SeenNames[$name] = $true

    $type = [string](Require-Property $Component 'type')
    if (@('class', 'standard', 'document') -notcontains $type) {
        throw "Component '$name' has an unsupported type: $type"
    }
    $sourcePath = Assert-SourcePath ([string](Require-Property $Component 'source_path')) $Excluded "Component '$name' source_path"
    if ($SeenPaths.ContainsKey($sourcePath)) {
        throw "Component source_path is duplicated: $sourcePath"
    }
    $SeenPaths[$sourcePath] = $true

    $extension = [IO.Path]::GetExtension($sourcePath).ToLowerInvariant()
    if ($type -eq 'class' -and $extension -ne '.cls') {
        throw "Class component '$name' does not use a .cls source_path."
    }
    if ($type -ne 'class' -and $extension -ne '.bas') {
        throw "Non-class component '$name' does not use a .bas source_path."
    }
    if (-not $Excluded -and $name -match '(?i)(^|[_-])(test|tests|benchmark|benchmarks|stress|xlflow|dev|mock|fake)([_-]|$)') {
        throw "Included component has a development-like name: $name"
    }

    $relatedProperty = $Component.PSObject.Properties['related_paths']
    if ($null -eq $relatedProperty) {
        throw "Component '$name' is missing related_paths."
    }
    $relatedPaths = if ($null -eq $relatedProperty.Value) { @() } else { @($relatedProperty.Value) }
    foreach ($relatedPathValue in $relatedPaths) {
        [void](Assert-SourcePath ([string]$relatedPathValue) $Excluded "Component '$name' related_paths entry")
    }
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

function Publish-Report([string]$Path, [string]$Content) {
    $directory = [IO.Path]::GetDirectoryName($Path)
    if (-not (Test-Path -LiteralPath $directory -PathType Container)) {
        [void](New-Item -ItemType Directory -Path $directory -Force)
    }
    $fileName = [IO.Path]::GetFileName($Path)
    $stagingPath = Join-Path $directory ('.{0}.{1}.tmp' -f $fileName, [guid]::NewGuid().ToString('N'))
    $backupPath = Join-Path $directory ('.{0}.{1}.bak' -f $fileName, [guid]::NewGuid().ToString('N'))
    try {
        [IO.File]::WriteAllText($stagingPath, $Content, [Text.UTF8Encoding]::new($false))
        if (Test-Path -LiteralPath $Path -PathType Leaf) {
            [IO.File]::Replace($stagingPath, $Path, $backupPath)
            if (Test-Path -LiteralPath $backupPath -PathType Leaf) {
                Remove-Item -LiteralPath $backupPath -Force
            }
        }
        else {
            Move-Item -LiteralPath $stagingPath -Destination $Path
        }
    }
    finally {
        if (Test-Path -LiteralPath $stagingPath -PathType Leaf) { Remove-Item -LiteralPath $stagingPath -Force }
        if (Test-Path -LiteralPath $backupPath -PathType Leaf) { Remove-Item -LiteralPath $backupPath -Force }
    }
}

$resolvedArtifact = Resolve-InputPath $ArtifactPath
$resolvedManifest = "$resolvedArtifact.build.json"
$resolvedChecksum = "$resolvedArtifact.checksum.json"
$resolvedPolicy = Resolve-InputPath $PolicyPath
$artifactRelative = Get-RepoRelativePath $resolvedArtifact
$manifestRelative = "$artifactRelative.build.json"
$checksumRelative = "$artifactRelative.checksum.json"

if ($artifactRelative -ne 'build/Release/VBA-HTTP.xlsm') {
    throw "Security validation only accepts the canonical release output build/Release/VBA-HTTP.xlsm."
}
foreach ($path in @($resolvedArtifact, $resolvedManifest, $resolvedChecksum, $resolvedPolicy)) {
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
        throw "Release security input does not exist: $path"
    }
}

$manifest = Get-Content -LiteralPath $resolvedManifest -Raw | ConvertFrom-Json
$policy = Get-Content -LiteralPath $resolvedPolicy -Raw | ConvertFrom-Json

if ([int](Require-Property $manifest 'schema_version') -ne 1 -or
    [string](Require-Property $manifest 'command') -ne 'build' -or
    [string](Require-Property $manifest 'backend') -ne 'excel') {
    throw 'Release manifest schema or build identity is invalid.'
}
$manifestBase = Normalize-RepoPath ([string](Require-Property $manifest 'base'))
if ($manifestBase -ne 'build/VBA-HTTP.xlsm') {
    throw 'Release manifest base is not the tracked development workbook.'
}
$manifestOutput = Normalize-RepoPath ([string](Require-Property $manifest 'output'))
if ($manifestOutput -ne $artifactRelative) {
    throw 'Release manifest output does not match the canonical release artifact.'
}

$validation = Require-Property $manifest 'validation'
if ($validation.source_applied -ne $true -or
    [string]$validation.vbe_compile -ne 'passed' -or
    $validation.workbook_saved -ne $true -or
    $validation.workbook_closed -ne $true -or
    [string]$validation.excel_cleanup -ne 'clean') {
    throw 'Release manifest does not prove source application, VBE compile, save, close, and Excel cleanup.'
}
if ([string](Require-Property (Require-Property $manifest 'publication') 'method') -ne 'atomic_replace') {
    throw 'Release publication was not atomic_replace.'
}

$included = @($manifest.included_components)
$excluded = @($manifest.excluded_components)
$expectedIncluded = Get-StringList $policy.included
$expectedExcluded = Get-StringList $policy.excluded
Assert-ExactSet $expectedIncluded @($included | ForEach-Object { [string](Require-Property $_ 'name') }) 'Included component names'
Assert-ExactSet $expectedExcluded @($excluded | ForEach-Object { [string](Require-Property $_ 'name') }) 'Excluded component names'
if ([int](Require-Property $validation 'components_applied') -ne $included.Count) {
    throw 'Manifest components_applied does not match included component count.'
}

$seenIncludedNames = @{}
$seenIncludedPaths = @{}
foreach ($component in $included) {
    Assert-ComponentRecord $component $false $seenIncludedNames $seenIncludedPaths
}
$seenExcludedNames = @{}
$seenExcludedPaths = @{}
foreach ($component in $excluded) {
    Assert-ComponentRecord $component $true $seenExcludedNames $seenExcludedPaths
}
Assert-Unique $expectedIncluded 'Included policy'
Assert-Unique $expectedExcluded 'Excluded policy'

& (Join-Path $PSScriptRoot 'Verify-ReleaseChecksums.ps1') `
    -ArtifactPath $resolvedArtifact `
    -ManifestPath $resolvedManifest `
    -ChecksumPath $resolvedChecksum | Out-Null

$resolvedReportPath = $null
if (-not [string]::IsNullOrWhiteSpace($ReportPath)) {
    $resolvedReportPath = Resolve-InputPath $ReportPath
}
if ($null -ne $resolvedReportPath) {
    $report = [ordered]@{
        schema_version = 1
        status = 'passed'
        scope = 'release-manifest-integrity'
        artifact = $artifactRelative
        manifest = $manifestRelative
        checksum = $checksumRelative
        policy = Get-RepoRelativePath $resolvedPolicy
        artifact_sha256 = Get-Sha256Hex $resolvedArtifact
        manifest_sha256 = Get-Sha256Hex $resolvedManifest
        policy_sha256 = Get-Sha256Hex $resolvedPolicy
        included_component_count = $included.Count
        excluded_component_count = $excluded.Count
        checks = @(
            'canonical_artifact_path',
            'manifest_identity_and_paths',
            'vbe_compile_save_close_cleanup',
            'atomic_publication',
            'component_allowlist_and_exclusions',
            'source_path_boundary',
            'checksum_sidecar'
        )
        deferred_security_review = @(
            'http2_http3_tls_matrix',
            'integrated_and_proxy_challenge_authentication'
        )
    }
    Publish-Report $resolvedReportPath (($report | ConvertTo-Json -Depth 5 -Compress) + [Environment]::NewLine)
}

Write-Output "Release security validation passed: $($included.Count) included, $($excluded.Count) excluded; manifest paths, validation evidence, publication, and checksum verified."
