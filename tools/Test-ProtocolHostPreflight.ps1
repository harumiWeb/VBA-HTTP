[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$runnerPath = Join-Path $PSScriptRoot "Run-ProtocolHostValidation.ps1"
$validatorPath = Join-Path $PSScriptRoot "Validate-ProtocolHostEvidence.ps1"

$runner = [IO.File]::ReadAllText($runnerPath)
$preflightCall = '$capabilityPreflight = Invoke-ProtocolCapabilityPreflight'
$excelStart = '$excel = New-Object -ComObject Excel.Application'
$preflightIndex = $runner.IndexOf($preflightCall, [StringComparison]::Ordinal)
$excelIndex = $runner.IndexOf($excelStart, [StringComparison]::Ordinal)

if ($preflightIndex -lt 0) { throw "Protocol host runner does not invoke the capability preflight." }
if ($excelIndex -lt 0) { throw "Protocol host runner Excel start contract is missing." }
if ($preflightIndex -ge $excelIndex) { throw "Protocol capability preflight must run before Excel COM creation." }
if ($runner.IndexOf("Probe-WinHttpProtocol.ps1", [StringComparison]::Ordinal) -lt 0) {
    throw "Protocol host runner does not use the Excel-free WinHTTP probe."
}
if ($runner.IndexOf("-Required", [StringComparison]::Ordinal) -lt 0) {
    throw "Protocol capability preflight must use required mode."
}
if ($runner.IndexOf('$preflightTimeoutMilliseconds', [StringComparison]::Ordinal) -lt 0) {
    throw "Protocol capability preflight must have a bounded timeout."
}
if ($runner -match '(?im)Stop-Process\s+-Name\s+EXCEL') {
    throw "Protocol host runner must not stop Excel by process name."
}
if ($runner.IndexOf('preflight = $capabilityPreflight', [StringComparison]::Ordinal) -lt 0) {
    throw "Protocol host runner does not archive preflight metadata."
}
if ($runner -match 'HTTP/3') {
    throw "Protocol host promotion runner must not accept HTTP/3 after the support-boundary decision."
}

$validator = [IO.File]::ReadAllText($validatorPath)
if ($validator.IndexOf('environment.winhttp.preflight', [StringComparison]::Ordinal) -lt 0) {
    throw "Protocol host evidence validator does not validate preflight metadata."
}
if ($validator -match 'HTTP/3') {
    throw "Protocol host evidence validator must not accept HTTP/3 promotion records."
}

Write-Output "Protocol host preflight safety contract passed."
