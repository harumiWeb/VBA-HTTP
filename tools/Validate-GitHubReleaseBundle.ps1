[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$Tag,
    [Parameter(Mandatory = $true)]
    [string]$BundleDirectory,
    [string]$ExpectedCommit = "",
    [string]$RootPath = ""
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$projectRoot = if ([string]::IsNullOrWhiteSpace($RootPath)) { [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..')) } else { [IO.Path]::GetFullPath($RootPath) }
$resolvedBundle = if ([IO.Path]::IsPathRooted($BundleDirectory)) { [IO.Path]::GetFullPath($BundleDirectory) } else { [IO.Path]::GetFullPath((Join-Path $projectRoot $BundleDirectory)) }
& (Join-Path $PSScriptRoot 'Validate-ReleaseTag.ps1') -Tag $Tag -ExpectedCommit $ExpectedCommit | Out-Null
if (-not (Test-Path -LiteralPath $resolvedBundle -PathType Container)) { throw "Release bundle directory does not exist: $resolvedBundle" }

function Get-RequiredProperty($Object, [string]$Name) {
    if ($null -eq $Object) { throw "Release manifest record is null while reading '$Name'." }
    $property = $Object.PSObject.Properties[$Name]
    if ($null -eq $property) { throw "Release manifest is missing '$Name'." }
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

function Get-AssetPath([string]$Name) {
    if ($Name -notmatch '^[A-Za-z0-9._-]+$') { throw "Release asset name is unsafe: $Name" }
    $path = Join-Path $resolvedBundle $Name
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { throw "Release asset is missing: $Name" }
    return $path
}

$releaseManifestPath = Get-AssetPath "VBA-HTTP-$Tag.release.json"
$checksumPath = Get-AssetPath "VBA-HTTP-$Tag.SHA256SUMS.txt"
$release = Get-Content -LiteralPath $releaseManifestPath -Raw | ConvertFrom-Json
Assert-AllowedPropertySet $release @('schema_version', 'package_id', 'release_tag', 'version', 'source_revision', 'target', 'toolchain', 'support', 'pack', 'assets', 'checksum_file') 'Release manifest'
Assert-AllowedPropertySet (Get-RequiredProperty $release 'support') @('office_bitness', 'office_32_bit', 'http3_quic') 'Release support boundary'
Assert-AllowedPropertySet (Get-RequiredProperty $release 'toolchain') @('xlflow', 'task', 'go', 'psscriptanalyzer') 'Release toolchain'
Assert-AllowedPropertySet (Get-RequiredProperty $release 'pack') @('backend', 'experimental', 'vbe_validation', 'manifest') 'Release pack provenance'
if ([int](Get-RequiredProperty $release 'schema_version') -ne 1 -or
    [string](Get-RequiredProperty $release 'package_id') -ne 'VBA-HTTP' -or
    [string](Get-RequiredProperty $release 'release_tag') -ne $Tag -or
    [string](Get-RequiredProperty $release 'target') -ne 'windows-x64-office') {
    throw 'Release manifest identity or target is invalid.'
}
if ([string](Get-RequiredProperty (Get-RequiredProperty $release 'pack') 'vbe_validation') -ne 'not_performed') {
    throw 'Release manifest must explicitly state that VBE validation was not performed.'
}
$releaseRevision = [string](Get-RequiredProperty $release 'source_revision')
if ($releaseRevision -notmatch '^[0-9a-fA-F]{40,64}$') { throw 'Release manifest source revision is not a full hexadecimal SHA.' }
if ([string](Get-RequiredProperty $release 'version') -ne $Tag.Substring(1)) {
    throw 'Release manifest version does not match the tag.'
}
$support = Get-RequiredProperty $release 'support'
if ([string](Get-RequiredProperty $support 'office_bitness') -ne 'x64' -or
    [string](Get-RequiredProperty $support 'office_32_bit') -ne 'unverified' -or
    [string](Get-RequiredProperty $support 'http3_quic') -ne 'unsupported-by-policy') {
    throw 'Release manifest support boundary is invalid.'
}
$toolchain = Get-RequiredProperty $release 'toolchain'
foreach ($tool in @('xlflow', 'task', 'go', 'psscriptanalyzer')) {
    if ([string]::IsNullOrWhiteSpace([string](Get-RequiredProperty $toolchain $tool))) {
        throw "Release manifest toolchain is missing '$tool'."
    }
}
$lock = Get-Content -LiteralPath (Join-Path $projectRoot 'tools/release-toolchain.json') -Raw | ConvertFrom-Json
if ([string]$toolchain.task -ne [string]$lock.task.version -or
    [string]$toolchain.go -ne [string]$lock.go.version -or
    [string]$toolchain.psscriptanalyzer -ne [string]$lock.psscriptanalyzer.version) {
    throw 'Release manifest toolchain does not match the repository lock.'
}
if (-not [string]::IsNullOrWhiteSpace($ExpectedCommit) -and
    [string]$toolchain.xlflow -ne ([string]$lock.xlflow.version).TrimStart('v')) {
    throw 'Release manifest xlflow version does not match the pinned release toolchain.'
}
if (-not [string]::IsNullOrWhiteSpace($ExpectedCommit) -and $releaseRevision -ne $ExpectedCommit.Trim()) {
    throw 'Release manifest source revision does not match the expected tag commit.'
}

$sourceName = "VBA-HTTP-$Tag-source.zip"
$packName = "VBA-HTTP-$Tag.xlsm"
$packManifestName = "VBA-HTTP-$Tag.pack.json"
$licenseName = 'LICENSE'
$noticeName = 'THIRD_PARTY_NOTICES.md'
$releaseName = "VBA-HTTP-$Tag.release.json"
$checksumName = "VBA-HTTP-$Tag.SHA256SUMS.txt"
$expectedAssetNames = @($sourceName, $packName, $packManifestName, $licenseName, $noticeName)
$declaredChecksumName = [string](Get-RequiredProperty $release 'checksum_file')
if ($declaredChecksumName -ne $checksumName) { throw 'Release manifest checksum_file does not match the tagged asset name.' }
$expectedAllNames = @($expectedAssetNames + $releaseName + $checksumName) | Sort-Object
$actualNames = @(Get-ChildItem -LiteralPath $resolvedBundle -File | Select-Object -ExpandProperty Name | Sort-Object)
if ($null -ne (Compare-Object $expectedAllNames $actualNames)) { throw 'Release bundle contains an unexpected or missing file.' }

$releaseAssets = @((Get-RequiredProperty $release 'assets'))
if ($null -ne (Compare-Object ($expectedAssetNames | Sort-Object) (@($releaseAssets | ForEach-Object { [string](Get-RequiredProperty $_ 'name') } | Sort-Object)))) {
    throw 'Release manifest asset names do not match the required asset set.'
}
foreach ($asset in $releaseAssets) {
    Assert-AllowedPropertySet $asset @('name', 'sha256', 'bytes') 'Release asset record'
    $name = [string](Get-RequiredProperty $asset 'name')
    $path = Get-AssetPath $name
    $expectedHash = [string](Get-RequiredProperty $asset 'sha256')
    if ($expectedHash -notmatch '^[0-9a-f]{64}$' -or $expectedHash -ne (Get-Sha256Hex $path)) { throw "Release asset hash mismatch: $name" }
    if ([int64](Get-RequiredProperty $asset 'bytes') -ne ([IO.FileInfo]$path).Length) { throw "Release asset byte count mismatch: $name" }
}

$checksumLines = @(Get-Content -LiteralPath $checksumPath)
$checksums = @{}
foreach ($line in $checksumLines) {
    if ([string]::IsNullOrWhiteSpace($line)) { continue }
    if ($line -notmatch '^([0-9a-f]{64}) \*(.+)$') { throw 'Release checksum file contains an invalid line.' }
    $name = $Matches[2]
    if ($checksums.ContainsKey($name)) { throw "Release checksum file contains duplicate asset: $name" }
    $checksums[$name] = $Matches[1]
}
$checksumNames = @($checksums.Keys | Sort-Object)
$expectedChecksumNames = @($expectedAssetNames + $releaseName) | Sort-Object
if ($null -ne (Compare-Object $expectedChecksumNames $checksumNames)) { throw 'Release checksum file asset names are incomplete or unexpected.' }
foreach ($name in $checksumNames) {
    $path = Get-AssetPath $name
    if ($checksums[$name] -ne (Get-Sha256Hex $path)) { throw "Release checksum mismatch: $name" }
}

$sourceZip = Get-AssetPath $sourceName
$packArtifact = Get-AssetPath $packName
$packManifest = Get-AssetPath $packManifestName
& (Join-Path $PSScriptRoot 'Validate-SourceArchive.ps1') -ArchivePath $sourceZip | Out-Null
& (Join-Path $PSScriptRoot 'Validate-PackArtifact.ps1') -ArtifactPath $packArtifact -ManifestPath $packManifest -RootPath $projectRoot | Out-Null
$pack = Get-Content -LiteralPath $packManifest -Raw | ConvertFrom-Json
if ([string](Get-RequiredProperty $pack 'release_tag') -ne $Tag -or
    [string](Get-RequiredProperty $pack 'source_revision') -ne $releaseRevision -or
    [string](Get-RequiredProperty $pack 'artifact') -ne $packName -or
    [string](Get-RequiredProperty (Get-RequiredProperty $release 'pack') 'manifest') -ne $packManifestName -or
    [string]$toolchain.xlflow -ne [string]$pack.xlflow_version) {
    throw 'Pack manifest identity does not match the release manifest.'
}

$extractRoot = Join-Path ([IO.Path]::GetTempPath()) ('vba-http-release-source-' + [guid]::NewGuid().ToString('N'))
try {
    Expand-Archive -LiteralPath $sourceZip -DestinationPath $extractRoot -Force
    $sourceManifest = Get-Content -LiteralPath (Join-Path $extractRoot 'manifest.json') -Raw | ConvertFrom-Json
    if ([string](Get-RequiredProperty $sourceManifest 'package_version') -ne $Tag -or
        [string](Get-RequiredProperty $sourceManifest 'source_revision') -ne [string](Get-RequiredProperty $release 'source_revision')) {
        throw 'Source package version or source revision does not match the release manifest.'
    }
}
finally {
    if (Test-Path -LiteralPath $extractRoot -PathType Container) { Remove-Item -LiteralPath $extractRoot -Recurse -Force }
}

Write-Output "GitHub release bundle valid: $Tag, $($expectedAssetNames.Count) primary assets, pack VBE validation not performed."
