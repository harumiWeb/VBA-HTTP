[CmdletBinding()]
param(
    [switch]$FullRelease
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$projectRoot = [IO.Path]::GetFullPath((Split-Path -Parent $PSScriptRoot))
$tempRoot = Join-Path ([IO.Path]::GetTempPath()) ("vba-http-clean-" + [guid]::NewGuid().ToString("N"))
$archivePath = Join-Path ([IO.Path]::GetTempPath()) ("vba-http-clean-" + [guid]::NewGuid().ToString("N") + ".zip")

function Invoke-Checked([string]$FilePath, [string[]]$Arguments) {
    & $FilePath @Arguments
    if ($LASTEXITCODE -ne 0) {
        throw "$FilePath failed with exit code $LASTEXITCODE."
    }
}

try {
    Invoke-Checked "git" @("-C", $projectRoot, "archive", "--format=zip", "--output=$archivePath", "HEAD")
    Expand-Archive -LiteralPath $archivePath -DestinationPath $tempRoot -Force
    if (-not (Test-Path -LiteralPath (Join-Path $tempRoot "build\VBA-HTTP.xlsm") -PathType Leaf) -or
        -not (Test-Path -LiteralPath (Join-Path $tempRoot "build\VBA-HTTP.xlam") -PathType Leaf)) {
        throw "The clean archive does not contain both tracked development workbook bases."
    }

    Push-Location $tempRoot
    try {
        Invoke-Checked "task" @("check")
        Invoke-Checked "task" @("test:docs")
        Invoke-Checked "task" @("test:xlam")
        Invoke-Checked "task" @("build:plan")
        Invoke-Checked "task" @("build:plan:xlam")
        if ($FullRelease) {
            Invoke-Checked "task" @("release:build")
            Invoke-Checked "task" @("release:xlam:build")
        }
    }
    finally {
        Pop-Location
    }

    if ($FullRelease) {
        Write-Output "Clean checkout contract and both release build targets passed."
    }
    else {
        Write-Output "Clean checkout contract, source gates, and both dry-run build plans passed."
    }
}
finally {
    if (Test-Path -LiteralPath $archivePath -PathType Leaf) {
        Remove-Item -LiteralPath $archivePath -Force
    }
    $tempRootFull = [IO.Path]::GetFullPath($tempRoot)
    $tempPrefix = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\', '/') + [IO.Path]::DirectorySeparatorChar
    if ($tempRootFull.StartsWith($tempPrefix, [StringComparison]::OrdinalIgnoreCase) -and
        (Test-Path -LiteralPath $tempRootFull)) {
        Remove-Item -LiteralPath $tempRootFull -Recurse -Force
    }
}
