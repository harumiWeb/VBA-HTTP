[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"
$projectRoot = Split-Path -Parent $PSScriptRoot
$validator = Join-Path $PSScriptRoot "Validate-SecurityRiskRegister.ps1"
$sourcePath = Join-Path $projectRoot "docs/security/risk-register.json"
$testRoot = [IO.Path]::GetFullPath((Join-Path $projectRoot ".xlflow\security-risk-register-test"))
$validPath = Join-Path $testRoot "valid.json"
$invalidPath = Join-Path $testRoot "invalid.json"

try {
    [void](New-Item -ItemType Directory -Path $testRoot -Force)
    Copy-Item -LiteralPath $sourcePath -Destination $validPath -Force
    try {
        & $validator -Path $validPath | Out-Null
    }
    catch {
        throw "Valid security risk register was rejected: $($_.Exception.Message)"
    }

    $fixture = Get-Content -LiteralPath $validPath -Raw | ConvertFrom-Json
    $fixture.current_release_blockers = @('compat-http2-http3-32bit')
    [IO.File]::WriteAllText($invalidPath, ($fixture | ConvertTo-Json -Depth 10), [Text.UTF8Encoding]::new($false))
    $accepted = $true
    try { & $validator -Path $invalidPath | Out-Null } catch { $accepted = $false }
    if ($accepted) { throw "Validator accepted a current release blocker." }

    $fixture.current_release_blockers = @()
    $fixture.issues[0].id = $fixture.issues[1].id
    [IO.File]::WriteAllText($invalidPath, ($fixture | ConvertTo-Json -Depth 10), [Text.UTF8Encoding]::new($false))
    $accepted = $true
    try { & $validator -Path $invalidPath | Out-Null } catch { $accepted = $false }
    if ($accepted) { throw "Validator accepted duplicate risk IDs." }
}
finally {
    if (Test-Path -LiteralPath $testRoot) {
        Remove-Item -LiteralPath $testRoot -Recurse -Force
    }
}

Write-Output "Security risk register validation tests passed."
