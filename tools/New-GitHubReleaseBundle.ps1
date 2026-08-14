[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$Tag,
    [string]$SourceRevision = "",
    [string]$OutputDirectory = "",
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
& (Join-Path $PSScriptRoot 'Validate-ReleaseTag.ps1') -Tag $Tag | Out-Null

if ([string]::IsNullOrWhiteSpace($SourceRevision)) {
    $gitOutput = & git -C $projectRoot rev-parse HEAD 2>$null
    $gitExitCode = $LASTEXITCODE
    $SourceRevision = ($gitOutput | Select-Object -First 1)
    if ($gitExitCode -ne 0 -or [string]::IsNullOrWhiteSpace([string]$SourceRevision)) {
        throw 'Could not determine the source revision for the release bundle.'
    }
    $SourceRevision = ([string]$SourceRevision).Trim()
}
if ($SourceRevision -notmatch '^[0-9a-fA-F]{40,64}$') {
    throw 'SourceRevision must be a full hexadecimal commit SHA.'
}

$resolvedOutputDirectory = if ([string]::IsNullOrWhiteSpace($OutputDirectory)) {
    [IO.Path]::GetFullPath((Join-Path $projectRoot (Join-Path '.release' $Tag)))
}
elseif ([IO.Path]::IsPathRooted($OutputDirectory)) {
    [IO.Path]::GetFullPath($OutputDirectory)
}
else {
    [IO.Path]::GetFullPath((Join-Path $projectRoot $OutputDirectory))
}
$rootWithSeparator = $projectRoot.TrimEnd('\', '/') + [IO.Path]::DirectorySeparatorChar
if (-not $resolvedOutputDirectory.StartsWith($rootWithSeparator, [StringComparison]::OrdinalIgnoreCase)) {
    throw 'Release bundle output must be inside the project root.'
}
if (Test-Path -LiteralPath $resolvedOutputDirectory) {
    $existing = @(Get-ChildItem -LiteralPath $resolvedOutputDirectory -Force)
    if ($existing.Count -gt 0) {
        throw "Release bundle output directory is not empty: $resolvedOutputDirectory"
    }
}
[void](New-Item -ItemType Directory -Path $resolvedOutputDirectory -Force)

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

function Publish-TextAtomic([string]$Path, [string]$Content) {
    $directory = [IO.Path]::GetDirectoryName($Path)
    $staging = Join-Path $directory ('.{0}.{1}.tmp' -f [IO.Path]::GetFileName($Path), [guid]::NewGuid().ToString('N'))
    $backup = Join-Path $directory ('.{0}.{1}.bak' -f [IO.Path]::GetFileName($Path), [guid]::NewGuid().ToString('N'))
    try {
        [IO.File]::WriteAllText($staging, $Content, [Text.UTF8Encoding]::new($false))
        if (Test-Path -LiteralPath $Path -PathType Leaf) {
            [IO.File]::Replace($staging, $Path, $backup)
            if (Test-Path -LiteralPath $backup -PathType Leaf) { Remove-Item -LiteralPath $backup -Force }
        }
        else {
            Move-Item -LiteralPath $staging -Destination $Path
        }
    }
    finally {
        foreach ($candidate in @($staging, $backup)) {
            if (Test-Path -LiteralPath $candidate -PathType Leaf) { Remove-Item -LiteralPath $candidate -Force }
        }
    }
}

$sourceZip = Join-Path $resolvedOutputDirectory "VBA-HTTP-$Tag-source.zip"
$packArtifact = Join-Path $resolvedOutputDirectory "VBA-HTTP-$Tag.xlsm"
$packManifest = Join-Path $resolvedOutputDirectory "VBA-HTTP-$Tag.pack.json"
$releaseManifest = Join-Path $resolvedOutputDirectory "VBA-HTTP-$Tag.release.json"
$checksumFile = Join-Path $resolvedOutputDirectory "VBA-HTTP-$Tag.SHA256SUMS.txt"
$licensePath = Join-Path $resolvedOutputDirectory 'LICENSE'
$noticePath = Join-Path $resolvedOutputDirectory 'THIRD_PARTY_NOTICES.md'

& (Join-Path $PSScriptRoot 'New-SourcePackage.ps1') -OutputPath $sourceZip -PackageVersion $Tag -SourceRevision $SourceRevision
& (Join-Path $PSScriptRoot 'Validate-SourceArchive.ps1') -ArchivePath $sourceZip
& (Join-Path $PSScriptRoot 'New-PackArtifact.ps1') -ReleaseTag $Tag -SourceRevision $SourceRevision -OutputPath $packArtifact -ManifestPath $packManifest
& (Join-Path $PSScriptRoot 'Validate-PackArtifact.ps1') -ArtifactPath $packArtifact -ManifestPath $packManifest
$pack = Get-Content -LiteralPath $packManifest -Raw | ConvertFrom-Json
$toolchain = Get-Content -LiteralPath (Join-Path $projectRoot 'tools/release-toolchain.json') -Raw | ConvertFrom-Json
Copy-Item -LiteralPath (Join-Path $projectRoot 'LICENSE') -Destination $licensePath -Force
Copy-Item -LiteralPath (Join-Path $projectRoot 'THIRD_PARTY_NOTICES.md') -Destination $noticePath -Force

$assetPaths = @($sourceZip, $packArtifact, $packManifest, $licensePath, $noticePath)
$assets = @()
foreach ($assetPath in $assetPaths) {
    $assets += [ordered]@{
        name = [IO.Path]::GetFileName($assetPath)
        sha256 = Get-Sha256Hex $assetPath
        bytes = ([IO.FileInfo]$assetPath).Length
    }
}
$release = [ordered]@{
    schema_version = 1
    package_id = 'VBA-HTTP'
    release_tag = $Tag
    version = $Tag.Substring(1)
    source_revision = $SourceRevision
    target = 'windows-x64-office'
    toolchain = [ordered]@{
        xlflow = [string]$pack.xlflow_version
        task = [string]$toolchain.task.version
        go = [string]$toolchain.go.version
        psscriptanalyzer = [string]$toolchain.psscriptanalyzer.version
    }
    support = [ordered]@{
        office_bitness = 'x64'
        office_32_bit = 'unsupported-by-policy'
        http3_quic = 'unsupported-by-policy'
    }
    pack = [ordered]@{
        backend = 'pure-go'
        experimental = $true
        vbe_validation = 'not_performed'
        manifest = [IO.Path]::GetFileName($packManifest)
    }
    assets = @($assets | Sort-Object name)
    checksum_file = [IO.Path]::GetFileName($checksumFile)
}
Publish-TextAtomic $releaseManifest (($release | ConvertTo-Json -Depth 8) + [Environment]::NewLine)

$checksumLines = @()
foreach ($asset in @($assets | Sort-Object name)) {
    $checksumLines += ("{0} *{1}" -f $asset.sha256, $asset.name)
}
$checksumLines += ("{0} *{1}" -f (Get-Sha256Hex $releaseManifest), [IO.Path]::GetFileName($releaseManifest))
Publish-TextAtomic $checksumFile (($checksumLines -join [Environment]::NewLine) + [Environment]::NewLine)

Write-Output "GitHub release bundle created: $resolvedOutputDirectory ($($assets.Count) primary assets plus manifest and checksums)."
