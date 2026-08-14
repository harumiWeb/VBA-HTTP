[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$runner = Join-Path $PSScriptRoot "Run-OfficeBitnessValidation.ps1"
$validator = Join-Path $PSScriptRoot "Validate-OfficeBitnessResult.ps1"
$fixtureRoot = Join-Path (Split-Path -Parent $PSScriptRoot) ".xlflow\office-bitness-policy-test"

# This is intentionally Excel-free. The policy guard must reject a normal X86
# promotion invocation before xlflow or COM can create/attach to a workbook.
$output = @()
$childExit = 0
try {
    $output = @(& powershell -NoProfile -ExecutionPolicy Bypass -File $runner -ExpectedArchitecture X86 2>&1)
    $childExit = $LASTEXITCODE
}
catch {
    $output += $_
    $childExit = 1
}
if ($childExit -eq 0) {
    throw "Office bitness runner accepted an X86 promotion invocation."
}
if (($output -join "`n") -notmatch "unverified") {
    throw "Office bitness runner failed without the expected unverified diagnostic."
}

try {
    [void](New-Item -ItemType Directory -Path $fixtureRoot -Force)
    $fixturePath = Join-Path $fixtureRoot "x86-diagnostic.json"
    $fixture = Get-Content -LiteralPath (Join-Path (Split-Path -Parent $PSScriptRoot) "benchmarks\results\office-bitness-x64.json") -Raw | ConvertFrom-Json
    $fixture.architecture = "X86"
    $fixture.support_status = "unverified"
    $fixture.status = "diagnostic"
    $fixture.consumer_smoke = "deferred-to-release-harness"
    $fixture | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $fixturePath -Encoding UTF8
    try {
        & $validator -Path $fixturePath | Out-Null
    }
    catch {
        throw "X86 diagnostic evidence was rejected by its validator: $($_.Exception.Message)"
    }

    $fixture.status = "passed"
    $fixture | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $fixturePath -Encoding UTF8
    $accepted = $true
    try { & $validator -Path $fixturePath | Out-Null } catch { $accepted = $false }
    if ($accepted) { throw "Validator accepted X86 evidence with promotion status passed." }
}
finally {
    if (Test-Path -LiteralPath $fixtureRoot) {
        Remove-Item -LiteralPath $fixtureRoot -Recurse -Force
    }
}

Write-Output "Office bitness policy guard passed: normal X86 promotion is rejected before Excel work, while diagnostic evidence remains available."
