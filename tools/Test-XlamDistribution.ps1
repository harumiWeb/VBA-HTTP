[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"
$planScript = Join-Path $PSScriptRoot "Validate-XlamBuildPlan.ps1"
& $planScript

$basePath = Join-Path (Split-Path -Parent $PSScriptRoot) "build\VBA-HTTP.xlam"
if (-not (Test-Path -LiteralPath $basePath -PathType Leaf)) {
    throw "Tracked XLAM base is missing: $basePath"
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$workbook = $null
try {
    $workbook = $excel.Workbooks.Open([IO.Path]::GetFullPath($basePath), 0, $true)
    if (-not [bool]$workbook.IsAddin) { throw "Tracked XLAM base is not marked as an add-in." }
}
finally {
    if ($null -ne $workbook) { $workbook.Close($false); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($workbook) }
    $excel.Quit()
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
}

Write-Output "XLAM distribution base and dry-run contract passed."
