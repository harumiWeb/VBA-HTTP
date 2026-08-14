[CmdletBinding()]
param(
    [switch]$FullRelease,
    [switch]$ExcelFree
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

if ($FullRelease -and $ExcelFree) {
    throw '-FullRelease cannot be combined with -ExcelFree because release builds require Excel/VBIDE.'
}

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
    $sourceRevision = ((& git -C $projectRoot rev-parse HEAD 2>$null) | Out-String).Trim()
    if ($LASTEXITCODE -ne 0 -or $sourceRevision -notmatch '^[0-9a-fA-F]{40,64}$') {
        throw 'Could not resolve the clean-checkout source revision.'
    }
    Invoke-Checked "git" @("-C", $projectRoot, "archive", "--format=zip", "--output=$archivePath", "HEAD")
    Expand-Archive -LiteralPath $archivePath -DestinationPath $tempRoot -Force
    if (-not (Test-Path -LiteralPath (Join-Path $tempRoot "build\VBA-HTTP.xlsm") -PathType Leaf) -or
        -not (Test-Path -LiteralPath (Join-Path $tempRoot "build\VBA-HTTP.xlam") -PathType Leaf)) {
        throw "The clean archive does not contain both tracked development workbook bases."
    }

    $previousArchiveRevision = [Environment]::GetEnvironmentVariable('VBA_HTTP_SOURCE_REVISION')
    $env:VBA_HTTP_SOURCE_REVISION = $sourceRevision.ToLowerInvariant()
    Push-Location $tempRoot
    try {
        Invoke-Checked "task" @("check")
        Invoke-Checked "task" @("test:docs")
        if (-not $ExcelFree) {
            Invoke-Checked "task" @("test:xlam")
        }
        Invoke-Checked "task" @("build:plan")
        Invoke-Checked "task" @("build:plan:xlam")
        if ($FullRelease) {
            Invoke-Checked "task" @("release:build")
            Invoke-Checked "task" @("release:xlam:build")
        }
    }
    finally {
        Pop-Location
        if ($null -eq $previousArchiveRevision) {
            Remove-Item Env:VBA_HTTP_SOURCE_REVISION -ErrorAction SilentlyContinue
        }
        else {
            $env:VBA_HTTP_SOURCE_REVISION = $previousArchiveRevision
        }
    }

    if ($FullRelease) {
        Write-Output "Clean checkout contract and both release build targets passed."
    }
    elseif ($ExcelFree) {
        Write-Output "Excel-free clean checkout contract, source gates, and both dry-run build plans passed."
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
