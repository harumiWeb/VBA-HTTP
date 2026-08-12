[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"
$projectRoot = Split-Path -Parent $PSScriptRoot
$excludedPath = Join-Path $projectRoot "src\modules\Xlflow\XlflowAssert.bas"
$classRoot = Join-Path $projectRoot "src\classes"
$formatCheckRoot = Join-Path $projectRoot ".xlflow\format-check"
$projectionRoot = Join-Path $formatCheckRoot ([Guid]::NewGuid().ToString("N"))
$sourceRoots = @(
    (Join-Path $projectRoot "src\modules"),
    (Join-Path $projectRoot "src\workbook")
)

$existingSourceRoots = @($sourceRoots | Where-Object { Test-Path -LiteralPath $_ -PathType Container })
$sourceFiles = [System.Collections.Generic.List[string]]::new()
@(
    Get-ChildItem -LiteralPath $existingSourceRoots -Recurse -File |
        Where-Object { $_.Extension -in ".bas", ".cls" -and $_.FullName -ne $excludedPath } |
        ForEach-Object { $_.FullName }
) | ForEach-Object { $sourceFiles.Add($_) }

try {
    if (Test-Path -LiteralPath $classRoot -PathType Container) {
        foreach ($file in Get-ChildItem -LiteralPath $classRoot -Recurse -File -Filter "*.cls") {
            $relativePath = $file.FullName.Substring($classRoot.Length).TrimStart([char[]]@("\", "/"))
            $projectionPath = Join-Path $projectionRoot $relativePath
            $projectionParent = Split-Path -Parent $projectionPath
            [void](New-Item -ItemType Directory -Path $projectionParent -Force)

            # xlflow fmt renders LF, while clean VBIDE class import requires CRLF.
            # Check a content-equivalent LF projection so neither invariant is skipped.
            $text = [IO.File]::ReadAllText($file.FullName, [Text.Encoding]::UTF8)
            $lfText = $text.Replace("`r`n", "`n").Replace("`r", "`n")
            [IO.File]::WriteAllText($projectionPath, $lfText, [Text.UTF8Encoding]::new($false))
            $sourceFiles.Add($projectionPath)
        }
    }

    if ($sourceFiles.Count -eq 0) {
        throw "No VBA source files were found for the format check."
    }

    & xlflow fmt --check @sourceFiles
    if ($LASTEXITCODE -ne 0) {
        throw "xlflow fmt --check failed with exit code $LASTEXITCODE."
    }
}
finally {
    $resolvedFormatRoot = [IO.Path]::GetFullPath($formatCheckRoot)
    $resolvedProjection = [IO.Path]::GetFullPath($projectionRoot)
    if ($resolvedProjection.StartsWith($resolvedFormatRoot + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase) -and
        (Test-Path -LiteralPath $resolvedProjection)) {
        Remove-Item -LiteralPath $resolvedProjection -Recurse -Force
    }
}
