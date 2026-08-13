[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$runner = Join-Path $PSScriptRoot "Run-OfficeBitnessValidation.ps1"

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
if (($output -join "`n") -notmatch "unsupported by policy") {
    throw "Office bitness runner failed without the expected unsupported-by-policy diagnostic."
}

Write-Output "Office bitness policy guard passed: normal X86 validation is rejected before Excel work."
