[CmdletBinding()]
param(
    [string]$ArtifactPath = "build/Release/VBA-HTTP.xlsm",
    [string]$ManifestPath = "",
    [string]$ChecksumPath = ""
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

function Get-RequiredProperty($Record, [string]$Name) {
    $property = $Record.PSObject.Properties[$Name]
    if ($null -eq $property) {
        throw "Release checksum sidecar is missing '$Name'."
    }
    return $property.Value
}

$resolvedArtifact = Resolve-ProjectPath $ArtifactPath
$resolvedManifest = if ([string]::IsNullOrWhiteSpace($ManifestPath)) {
    "$resolvedArtifact.build.json"
}
else {
    Resolve-ProjectPath $ManifestPath
}
$resolvedChecksum = if ([string]::IsNullOrWhiteSpace($ChecksumPath)) {
    "$resolvedArtifact.checksum.json"
}
else {
    Resolve-ProjectPath $ChecksumPath
}

foreach ($path in @($resolvedArtifact, $resolvedManifest, $resolvedChecksum)) {
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
        throw "Release checksum input does not exist: $path"
    }
}

$record = Get-Content -LiteralPath $resolvedChecksum -Raw | ConvertFrom-Json
$schemaVersion = Get-RequiredProperty $record "schema_version"
$algorithm = Get-RequiredProperty $record "algorithm"
$artifactName = Get-RequiredProperty $record "artifact"
$artifactHash = Get-RequiredProperty $record "artifact_sha256"
$manifestName = Get-RequiredProperty $record "manifest"
$manifestHash = Get-RequiredProperty $record "manifest_sha256"

if ([int]$schemaVersion -ne 1 -or [string]$algorithm -ne "SHA-256") {
    throw "Release checksum sidecar has an unsupported schema or algorithm."
}
if ([string]$artifactName -ne [IO.Path]::GetFileName($resolvedArtifact) -or
    [string]$manifestName -ne [IO.Path]::GetFileName($resolvedManifest)) {
    throw "Release checksum sidecar names do not match the artifact and manifest."
}

$actualArtifactHash = Get-Sha256Hex $resolvedArtifact
$actualManifestHash = Get-Sha256Hex $resolvedManifest
if ([string]$artifactHash -ne $actualArtifactHash) {
    throw "Release artifact SHA-256 does not match its checksum sidecar."
}
if ([string]$manifestHash -ne $actualManifestHash) {
    throw "Release manifest SHA-256 does not match its checksum sidecar."
}

Write-Output "Release checksums verified: $([IO.Path]::GetFileName($resolvedArtifact)) ($actualArtifactHash), $([IO.Path]::GetFileName($resolvedManifest)) ($actualManifestHash)."
