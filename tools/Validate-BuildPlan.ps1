[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"
$projectRoot = Split-Path -Parent $PSScriptRoot
$policyPath = Join-Path $PSScriptRoot "build-component-policy.json"
$outputPath = Join-Path $projectRoot "build\Release\VBA-HTTP.xlsm"

& (Join-Path $PSScriptRoot "Ensure-XlflowDirectories.ps1")

$policy = Get-Content -LiteralPath $policyPath -Raw | ConvertFrom-Json
$json = & xlflow build --dry-run --json --out $outputPath
if ($LASTEXITCODE -ne 0) {
    throw "xlflow build --dry-run failed with exit code $LASTEXITCODE."
}

$result = $json | Out-String | ConvertFrom-Json
if ($result.status -ne "ok") {
    throw "xlflow returned build status '$($result.status)'."
}

if (@($result.build.warnings).Count -ne 0) {
    throw "Build plan contains warnings: $($result.build.warnings | ConvertTo-Json -Compress)"
}

$actualIncluded = @($result.build.included_components.name | Sort-Object)
$actualExcluded = @($result.build.excluded_components.name | Sort-Object)
$expectedIncluded = @($policy.included | Sort-Object)
$expectedExcluded = @($policy.excluded | Sort-Object)

$includedDifference = Compare-Object $expectedIncluded $actualIncluded
if ($null -ne $includedDifference) {
    throw "Included components differ from policy: $($includedDifference | ConvertTo-Json -Compress)"
}

$excludedDifference = Compare-Object $expectedExcluded $actualExcluded
if ($null -ne $excludedDifference) {
    throw "Excluded components differ from policy: $($excludedDifference | ConvertTo-Json -Compress)"
}

Write-Output "Release build plan is valid: $($actualIncluded.Count) included, $($actualExcluded.Count) excluded."
