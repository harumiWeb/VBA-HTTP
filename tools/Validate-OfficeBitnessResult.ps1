[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$Path
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw "Office bitness result does not exist: $Path" }
$result = Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json
if ([int]$result.schema_version -ne 1 -or [string]$result.benchmark -ne "office-bitness-validation") {
    throw "Office bitness result schema identity is invalid."
}
if ([string]$result.architecture -notin @("X86", "X64") -or [string]$result.status -notin @("passed", "diagnostic") -or
    [string]$result.consumer_smoke -ne "passed" -or [bool]$result.external_network -or
    [string]$result.support_status -notin @("supported", "unsupported-by-policy")) {
    throw "Office bitness result status or architecture is invalid."
}
if (([string]$result.architecture -eq "X64" -and ([string]$result.support_status -ne "supported" -or [string]$result.status -ne "passed")) -or
    ([string]$result.architecture -eq "X86" -and ([string]$result.support_status -ne "unsupported-by-policy" -or [string]$result.status -ne "diagnostic"))) {
    throw "Office bitness support_status does not match the architecture policy."
}
if ([string]$result.build.vbe_compile -ne "passed" -or
    [bool]$result.build.source_applied -ne $true -or
    [bool]$result.build.workbook_saved -ne $true -or
    [bool]$result.build.workbook_closed -ne $true -or
    [string]$result.build.excel_cleanup -ne "clean") {
    throw "Office bitness result does not prove a clean VBE build."
}
if ([int]$result.tests.passed -le 0 -or [int]$result.tests.failed -ne 0 -or
    [int]$result.integration.total -le 0 -or [int]$result.integration.failed -ne 0 -or
    [int]$result.integration.passed -ne [int]$result.integration.total) {
    throw "Office bitness result contains failed or incomplete test evidence."
}
if ([string]$result.office.version -eq "unknown" -or [string]$result.office.build -eq "unknown") {
    throw "Office version/build metadata is missing from the bitness result."
}
Write-Output "Office bitness result valid: $($result.architecture), $($result.tests.passed) unit tests, $($result.integration.passed) integration tests."
