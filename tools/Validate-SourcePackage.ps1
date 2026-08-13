[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$PackageRoot
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$resolvedRoot = [IO.Path]::GetFullPath($PackageRoot)
if (-not (Test-Path -LiteralPath $resolvedRoot -PathType Container)) {
    throw "Source package root does not exist: $resolvedRoot"
}

function Get-RequiredProperty($Object, [string]$Name) {
    if ($null -eq $Object) { throw "Package record is null while reading '$Name'." }
    $property = $Object.PSObject.Properties[$Name]
    if ($null -eq $property) { throw "Package record is missing '$Name'." }
    return $property.Value
}

function ConvertTo-PackagePath([string]$Path) {
    if ([string]::IsNullOrWhiteSpace($Path)) { throw "Package path cannot be empty." }
    $normalized = $Path.Replace('\', '/').Trim()
    if ($normalized.StartsWith('/') -or $normalized -match '^[A-Za-z]:') {
        throw "Package paths must be relative: $Path"
    }
    if ($normalized -match '(^|/)\.\.?(/|$)' -or $normalized.Contains('//')) {
        throw "Package path contains traversal or an empty segment: $Path"
    }
    return $normalized
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

function Get-PackageFile([string]$RelativePath) {
    $normalized = ConvertTo-PackagePath $RelativePath
    $path = Join-Path $resolvedRoot ($normalized.Replace('/', [IO.Path]::DirectorySeparatorChar))
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
        throw "Source package file is missing: $RelativePath"
    }
    return $path
}

$manifestPath = Get-PackageFile 'manifest.json'
$manifest = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json
if ([int](Get-RequiredProperty $manifest 'schema_version') -ne 1 -or
    [string](Get-RequiredProperty $manifest 'package_id') -ne 'VBA-HTTP' -or
    [string](Get-RequiredProperty $manifest 'target') -ne 'windows-x64-office') {
    throw 'Source package manifest identity or support target is invalid.'
}

$install = Get-RequiredProperty $manifest 'install'
if ([string](Get-RequiredProperty $install 'script') -ne 'Install-VBAHttp.ps1' -or
    [string](Get-RequiredProperty $install 'uninstall_script') -ne 'Uninstall-VBAHttp.ps1' -or
    [bool](Get-RequiredProperty $install 'requires_closed_workbook') -ne $true -or
    [bool](Get-RequiredProperty $install 'requires_vbproject_access') -ne $true) {
    throw 'Source package install contract is invalid.'
}

$requiredFiles = @('README.md', 'docs/API.md', 'docs/specs/source-package.md', 'docs/specs/distribution.md', 'docs/specs/licensing.md', 'benchmarks/schema/source-package-manifest.schema.json', 'LICENSE', 'THIRD_PARTY_NOTICES.md', 'Install-VBAHttp.ps1', 'Uninstall-VBAHttp.ps1', 'Validate-SourcePackage.ps1')
foreach ($requiredFile in $requiredFiles) { [void](Get-PackageFile $requiredFile) }

$fileRecords = @((Get-RequiredProperty $manifest 'files'))
$seenFiles = @{}
foreach ($record in $fileRecords) {
    $relativePath = ConvertTo-PackagePath ([string](Get-RequiredProperty $record 'path'))
    if ($seenFiles.ContainsKey($relativePath)) { throw "Source package file is duplicated: $relativePath" }
    $seenFiles[$relativePath] = $true
    $path = Get-PackageFile $relativePath
    $expectedHash = [string](Get-RequiredProperty $record 'sha256')
    if ($expectedHash -notmatch '^[0-9a-f]{64}$' -or $expectedHash -ne (Get-Sha256Hex $path)) {
        throw "Source package file hash mismatch: $relativePath"
    }
}

$components = @((Get-RequiredProperty $manifest 'components'))
if ($components.Count -eq 0) { throw 'Source package contains no production components.' }
$seenNames = @{}
$seenPaths = @{}
foreach ($component in $components) {
    $name = [string](Get-RequiredProperty $component 'name')
    $type = [string](Get-RequiredProperty $component 'type')
    $sourcePath = ConvertTo-PackagePath ([string](Get-RequiredProperty $component 'source_path'))
    $packagePath = ConvertTo-PackagePath ([string](Get-RequiredProperty $component 'package_path'))
    if ($name -notmatch '^[A-Za-z][A-Za-z0-9_]*$') { throw "Unsafe VBA component name: $name" }
    if ($seenNames.ContainsKey($name)) { throw "Duplicate component name: $name" }
    if ($seenPaths.ContainsKey($packagePath)) { throw "Duplicate component path: $packagePath" }
    if ($name -in @('App', 'Main', 'Ui')) { throw "Scaffold entry component is not allowed in a source package: $name" }
    if ($sourcePath -match '(?i)(^|/)(Tests|Benchmarks|Xlflow|Dev)(/|$)') {
        throw "Development-only source path is not allowed in a source package: $sourcePath"
    }
    if ($packagePath -notmatch '^components/(classes|modules)/[^/]+\.(cls|bas)$') {
        throw "Component package_path is outside the source package boundary: $packagePath"
    }
    if (($type -eq 'class' -and $packagePath -notmatch '\.cls$') -or
        ($type -eq 'standard' -and $packagePath -notmatch '\.bas$')) {
        throw "Component type and package extension disagree: $name"
    }
    $componentPath = Get-PackageFile $packagePath
    $expectedHash = [string](Get-RequiredProperty $component 'sha256')
    if ($expectedHash -notmatch '^[0-9a-f]{64}$' -or $expectedHash -ne (Get-Sha256Hex $componentPath)) {
        throw "Component hash mismatch: $name"
    }
    if (-not $seenFiles.ContainsKey($packagePath)) { throw "Component is missing from manifest files: $packagePath" }
    $seenNames[$name] = $true
    $seenPaths[$packagePath] = $true
}

Write-Output "Source package valid: $($components.Count) components, $($fileRecords.Count) hashed files, target windows-x64-office."
