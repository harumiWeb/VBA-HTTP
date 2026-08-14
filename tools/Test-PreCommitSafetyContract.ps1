[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$scriptPath = Join-Path $PSScriptRoot 'Run-PreCommitCompile.ps1'
$script = Get-Content -LiteralPath $scriptPath -Raw
foreach ($token in @(
        'Get-Process -Name EXCEL',
        'baselineExcel',
        '.xlflow',
        'recovery_required',
        'workbook_open',
        'VBA-HTTP.xlsm',
        'architecture',
        'X64',
        'vbe_compile',
        'workbook_saved',
        'workbook_closed',
        'excel_cleanup',
        'atomic_replace',
        'No process was terminated'
    )) {
    if ($script.IndexOf($token, [StringComparison]::OrdinalIgnoreCase) -lt 0) {
        throw "Pre-commit safety contract is missing '$token'."
    }
}
if ($script -match '(?im)\bStop-Process\b') { throw 'Pre-commit compile must not terminate arbitrary Excel processes.' }
if ($script -match '(?im)\btask\s+test\b') { throw 'Pre-commit compile must not invoke the Excel test suite.' }
Write-Output 'Pre-commit safety contract valid: x64, lock/recovery checks, PID baseline, atomic temporary build, and no process termination.'
