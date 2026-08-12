[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"
$projectRoot = Split-Path -Parent $PSScriptRoot
$excludedPath = Join-Path $projectRoot "src\modules\Xlflow\XlflowAssert.bas"
$sourceRoots = @(
    (Join-Path $projectRoot "src\modules"),
    (Join-Path $projectRoot "src\classes"),
    (Join-Path $projectRoot "src\workbook")
)

$existingSourceRoots = @($sourceRoots | Where-Object { Test-Path -LiteralPath $_ -PathType Container })
$sourceFiles = @(
    Get-ChildItem -LiteralPath $existingSourceRoots -Recurse -File |
        Where-Object { $_.Extension -in ".bas", ".cls" -and $_.FullName -ne $excludedPath } |
        ForEach-Object { $_.FullName }
)

if ($sourceFiles.Count -eq 0) {
    throw "No VBA source files were found for the format check."
}

& xlflow fmt --check @sourceFiles
if ($LASTEXITCODE -ne 0) {
    throw "xlflow fmt --check failed with exit code $LASTEXITCODE."
}
