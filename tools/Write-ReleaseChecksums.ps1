[CmdletBinding()]
param(
    [string]$ArtifactPath = "build/Release/VBA-HTTP.xlsm",
    [string]$ManifestPath = "",
    [string]$OutputPath = ""
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$projectRoot = Split-Path -Parent $PSScriptRoot

function Resolve-ProjectPath([string]$Path) {
    if ([string]::IsNullOrWhiteSpace($Path)) {
        throw "A path cannot be empty."
    }
    if ([IO.Path]::IsPathRooted($Path)) {
        return [IO.Path]::GetFullPath($Path)
    }
    return [IO.Path]::GetFullPath((Join-Path $projectRoot $Path))
}

function Get-Sha256Hex([string]$Path) {
    $sha = [Security.Cryptography.SHA256]::Create()
    $stream = [IO.File]::OpenRead($Path)
    try {
        return ([BitConverter]::ToString($sha.ComputeHash($stream))).Replace("-", "").ToLowerInvariant()
    }
    finally {
        $stream.Dispose()
        $sha.Dispose()
    }
}

function Publish-TextAtomically([string]$Path, [string]$Content) {
    $directory = [IO.Path]::GetDirectoryName($Path)
    if (-not (Test-Path -LiteralPath $directory -PathType Container)) {
        [void](New-Item -ItemType Directory -Path $directory -Force)
    }

    $fileName = [IO.Path]::GetFileName($Path)
    $stagingPath = Join-Path $directory (".{0}.{1}.tmp" -f $fileName, [guid]::NewGuid().ToString("N"))
    $backupPath = Join-Path $directory (".{0}.{1}.bak" -f $fileName, [guid]::NewGuid().ToString("N"))
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
        if (Test-Path -LiteralPath $stagingPath -PathType Leaf) {
            Remove-Item -LiteralPath $stagingPath -Force
        }
        if (Test-Path -LiteralPath $backupPath -PathType Leaf) {
            Remove-Item -LiteralPath $backupPath -Force
        }
    }
}

$resolvedArtifact = Resolve-ProjectPath $ArtifactPath
$resolvedManifest = if ([string]::IsNullOrWhiteSpace($ManifestPath)) {
    "$resolvedArtifact.build.json"
}
else {
    Resolve-ProjectPath $ManifestPath
}
$resolvedOutput = if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    "$resolvedArtifact.checksum.json"
}
else {
    Resolve-ProjectPath $OutputPath
}

if (-not (Test-Path -LiteralPath $resolvedArtifact -PathType Leaf)) {
    throw "Release artifact does not exist: $resolvedArtifact"
}
if (-not (Test-Path -LiteralPath $resolvedManifest -PathType Leaf)) {
    throw "Release manifest does not exist: $resolvedManifest"
}

$artifactDirectory = [IO.Path]::GetFullPath([IO.Path]::GetDirectoryName($resolvedArtifact))
$outputDirectory = [IO.Path]::GetFullPath([IO.Path]::GetDirectoryName($resolvedOutput))
if (-not $outputDirectory.Equals($artifactDirectory, [StringComparison]::OrdinalIgnoreCase)) {
    throw "Checksum sidecar must be published beside the release artifact."
}
if ($resolvedOutput.Equals($resolvedArtifact, [StringComparison]::OrdinalIgnoreCase) -or
    $resolvedOutput.Equals($resolvedManifest, [StringComparison]::OrdinalIgnoreCase)) {
    throw "Checksum sidecar cannot replace the artifact or build manifest."
}

$record = [ordered]@{
    schema_version = 1
    algorithm = "SHA-256"
    artifact = [IO.Path]::GetFileName($resolvedArtifact)
    artifact_sha256 = Get-Sha256Hex $resolvedArtifact
    manifest = [IO.Path]::GetFileName($resolvedManifest)
    manifest_sha256 = Get-Sha256Hex $resolvedManifest
}
$json = ($record | ConvertTo-Json -Depth 3 -Compress) + [Environment]::NewLine
Publish-TextAtomically $resolvedOutput $json

Write-Output "Release checksums written: $([IO.Path]::GetFileName($resolvedOutput)) (artifact $($record.artifact_sha256), manifest $($record.manifest_sha256))."
