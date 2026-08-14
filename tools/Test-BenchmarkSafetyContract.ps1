[CmdletBinding()]
param(
    [string]$ProjectRoot
)

$ErrorActionPreference = "Stop"
if ([string]::IsNullOrWhiteSpace($ProjectRoot)) {
    $ProjectRoot = Split-Path -Parent $PSScriptRoot
}

$runnerPath = Join-Path $ProjectRoot "tools\Run-RawBenchmark.ps1"
if (-not (Test-Path -LiteralPath $runnerPath -PathType Leaf)) {
    throw "Benchmark runner was not found: $runnerPath"
}

$runner = Get-Content -LiteralPath $runnerPath -Raw
$requiredTokens = @(
    'function Get-ExcelProcessId',
    '$baselineExcelProcessIds = @(Get-ExcelProcessId)',
    'Benchmark requires an Excel-exclusive window.',
    'Wait-ForBenchmarkExcelCleanup -BaselineProcessIds $baselineExcelProcessIds',
    'No Excel process was touched.'
)
foreach ($token in $requiredTokens) {
    if ($runner.IndexOf($token, [StringComparison]::Ordinal) -lt 0) {
        throw "Benchmark safety contract is missing token: $token"
    }
}

if ($runner -match '(?im)\bStop-Process[^\r\n]*\bEXCEL\b') {
    throw "Benchmark runner must not terminate Excel processes."
}

Write-Output "Benchmark safety contract valid: existing Excel is rejected and no Excel process is terminated."
