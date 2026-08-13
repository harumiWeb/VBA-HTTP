[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$fixtureRoot = Join-Path ([IO.Path]::GetTempPath()) ("vba-http-source-package-test-" + [guid]::NewGuid().ToString('N'))
$zipPath = Join-Path $fixtureRoot 'VBA-HTTP-source.zip'
$extractRoot = Join-Path $fixtureRoot 'extracted'
$tamperedRoot = Join-Path $fixtureRoot 'tampered'
$whatIfWorkbook = Join-Path $fixtureRoot 'consumer.xlsm'

function Test-ThrowsError([scriptblock]$Script, [string]$Description) {
    $threw = $false
    try { & $Script }
    catch { $threw = $true }
    if (-not $threw) { throw "Source package test expected failure: $Description" }
}

try {
    [void](New-Item -ItemType Directory -Path $fixtureRoot -Force)
    & (Join-Path $PSScriptRoot 'New-SourcePackage.ps1') -OutputPath $zipPath -PackageVersion 'test'
    if (-not (Test-Path -LiteralPath $zipPath -PathType Leaf)) {
        throw 'Source package test did not produce a ZIP file.'
    }
    & (Join-Path $PSScriptRoot 'New-SourcePackage.ps1') -OutputPath $zipPath -PackageVersion 'test'

    Expand-Archive -LiteralPath $zipPath -DestinationPath $extractRoot -Force
    & (Join-Path $PSScriptRoot 'Validate-SourcePackage.ps1') -PackageRoot $extractRoot
    [IO.File]::WriteAllBytes($whatIfWorkbook, [byte[]](0..15))
    & (Join-Path $extractRoot 'Install-VBAHttp.ps1') -Workbook $whatIfWorkbook -PackageRoot $extractRoot -WhatIf -Confirm:$false
    if (Test-Path -LiteralPath "$whatIfWorkbook.vba-http.bak" -PathType Leaf) {
        throw 'Source package installer WhatIf created a backup or touched the target.'
    }
    $manifest = Get-Content -LiteralPath (Join-Path $extractRoot 'manifest.json') -Raw | ConvertFrom-Json
    $names = @($manifest.components.name)
    foreach ($forbidden in @('App', 'Main', 'Ui', 'Sheet1', 'ThisWorkbook')) {
        if ($names -contains $forbidden) { throw "Source package contains forbidden scaffold/document component: $forbidden" }
    }
    if (@($manifest.components).Count -le 10) { throw 'Source package unexpectedly contains too few production components.' }

    Copy-Item -LiteralPath $extractRoot -Destination $tamperedRoot -Recurse -Force
    $tamperedComponent = Join-Path $tamperedRoot ([string]$manifest.components[0].package_path).Replace('/', [IO.Path]::DirectorySeparatorChar)
    Add-Content -LiteralPath $tamperedComponent -Value "' tampered" -Encoding UTF8
    Test-ThrowsError { & (Join-Path $PSScriptRoot 'Validate-SourcePackage.ps1') -PackageRoot $tamperedRoot } 'tampered component'
    Write-Output "Source package tests passed: $($manifest.components.Count) components, manifest hashes, scaffold exclusion, and tamper rejection."
}
finally {
    if (Test-Path -LiteralPath $fixtureRoot) {
        Remove-Item -LiteralPath $fixtureRoot -Recurse -Force
    }
}
