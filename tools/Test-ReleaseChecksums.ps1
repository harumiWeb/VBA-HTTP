[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$projectRoot = Split-Path -Parent $PSScriptRoot
$testRoot = [IO.Path]::GetFullPath((Join-Path $projectRoot ".xlflow\release-checksum-test"))
$testDirectory = Join-Path $testRoot ([guid]::NewGuid().ToString("N"))
$artifactPath = Join-Path $testDirectory "fixture.xlsm"
$manifestPath = "$artifactPath.build.json"
$checksumPath = "$artifactPath.checksum.json"
$writerPath = Join-Path $PSScriptRoot "Write-ReleaseChecksums.ps1"
$verifierPath = Join-Path $PSScriptRoot "Verify-ReleaseChecksums.ps1"

function Assert-Throws([scriptblock]$Action, [string]$Message) {
    $threw = $false
    try {
        & $Action | Out-Null
    }
    catch {
        $threw = $true
    }
    if (-not $threw) {
        throw $Message
    }
}

try {
    [void](New-Item -ItemType Directory -Path $testDirectory -Force)
    [IO.File]::WriteAllBytes($artifactPath, [byte[]](0..255))
    [IO.File]::WriteAllText($manifestPath, '{"schema_version":1,"validation":{"vbe_compile":"passed"}}' + [Environment]::NewLine, [Text.UTF8Encoding]::new($false))

    & $writerPath -ArtifactPath $artifactPath -ManifestPath $manifestPath -OutputPath $checksumPath | Out-Null
    if (-not (Test-Path -LiteralPath $checksumPath -PathType Leaf)) {
        throw "Checksum writer did not publish a sidecar."
    }
    $firstSidecar = Get-Content -LiteralPath $checksumPath -Raw
    & $verifierPath -ArtifactPath $artifactPath -ManifestPath $manifestPath -ChecksumPath $checksumPath | Out-Null
    & $writerPath -ArtifactPath $artifactPath -ManifestPath $manifestPath -OutputPath $checksumPath | Out-Null
    if ($firstSidecar -ne (Get-Content -LiteralPath $checksumPath -Raw)) {
        throw "Checksum sidecar is not deterministic for unchanged inputs."
    }

    [IO.File]::WriteAllBytes($artifactPath, [byte[]](255..0))
    Assert-Throws { & $verifierPath -ArtifactPath $artifactPath -ManifestPath $manifestPath -ChecksumPath $checksumPath } "Verifier accepted a modified release artifact."
    [IO.File]::WriteAllBytes($artifactPath, [byte[]](0..255))
    & $writerPath -ArtifactPath $artifactPath -ManifestPath $manifestPath -OutputPath $checksumPath | Out-Null

    $sidecarBeforeMissingInput = Get-Content -LiteralPath $checksumPath -Raw
    Remove-Item -LiteralPath $manifestPath -Force
    Assert-Throws { & $writerPath -ArtifactPath $artifactPath -ManifestPath $manifestPath -OutputPath $checksumPath } "Checksum writer accepted a missing build manifest."
    if ($sidecarBeforeMissingInput -ne (Get-Content -LiteralPath $checksumPath -Raw)) {
        throw "A failed checksum generation replaced the previous sidecar."
    }

    Write-Output "Release checksum generation and validation tests passed."
}
finally {
    if (Test-Path -LiteralPath $testDirectory) {
        Remove-Item -LiteralPath $testDirectory -Recurse -Force
    }
    if ((Test-Path -LiteralPath $testRoot -PathType Container) -and
        -not (Get-ChildItem -LiteralPath $testRoot -Force | Select-Object -First 1)) {
        Remove-Item -LiteralPath $testRoot -Force
    }
}
