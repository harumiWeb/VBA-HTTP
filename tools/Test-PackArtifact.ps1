[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$fixtureRoot = Join-Path ([IO.Path]::GetTempPath()) ('vba-http-pack-test-' + [guid]::NewGuid().ToString('N'))
$artifactPath = Join-Path $fixtureRoot 'VBA-HTTP-test.xlsm'
$manifestPath = Join-Path $fixtureRoot 'VBA-HTTP-test.pack.json'
$tamperedManifest = Join-Path $fixtureRoot 'tampered.pack.json'

function Test-ThrowsError([scriptblock]$Script, [string]$Description) {
    $threw = $false
    try { & $Script }
    catch { $threw = $true }
    if (-not $threw) { throw "Pack artifact test expected failure: $Description" }
}

try {
    [void](New-Item -ItemType Directory -Path $fixtureRoot -Force)
    & (Join-Path $PSScriptRoot 'New-PackArtifact.ps1') -ReleaseTag 'v0.0.0-test' -SourceRevision '0123456789abcdef0123456789abcdef01234567' -OutputPath $artifactPath -ManifestPath $manifestPath
    & (Join-Path $PSScriptRoot 'Validate-PackArtifact.ps1') -ArtifactPath $artifactPath -ManifestPath $manifestPath
    $manifest = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json
    if ($manifest.vbe_validation -ne 'not_performed' -or $manifest.pack_backend -ne 'pure-go' -or $manifest.experimental -ne $true) {
        throw 'Pack artifact provenance was not recorded correctly.'
    }
    foreach ($forbidden in @('App', 'Main', 'Ui', 'XlflowAssert', 'WinHttpNativeTransportTests', 'RawWinHttpBenchmark')) {
        if (@($manifest.included_components.name) -contains $forbidden) { throw "Pack artifact includes forbidden component: $forbidden" }
    }

    Copy-Item -LiteralPath $manifestPath -Destination $tamperedManifest -Force
    $tampered = Get-Content -LiteralPath $tamperedManifest -Raw | ConvertFrom-Json
    $tampered.included_components[0].sha256 = ('0' * 64)
    [IO.File]::WriteAllText($tamperedManifest, (($tampered | ConvertTo-Json -Depth 8) + [Environment]::NewLine), [Text.UTF8Encoding]::new($false))
    Test-ThrowsError { & (Join-Path $PSScriptRoot 'Validate-PackArtifact.ps1') -ArtifactPath $artifactPath -ManifestPath $tamperedManifest } 'tampered pack manifest'
    Write-Output "Pack artifact tests passed: $(@($manifest.included_components).Count) production components, no development components, and tamper rejection."
}
finally {
    if (Test-Path -LiteralPath $fixtureRoot -PathType Container) { Remove-Item -LiteralPath $fixtureRoot -Recurse -Force }
}
