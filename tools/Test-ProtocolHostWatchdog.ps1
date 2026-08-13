[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$watchdog = Join-Path $PSScriptRoot "Watch-ProtocolHostExcel.ps1"

function Get-ExcelId {
    return @(
        Get-Process -Name EXCEL -ErrorAction SilentlyContinue |
            ForEach-Object { [int]$_.Id }
    )
}

function Assert-EqualId([int[]]$Expected, [int[]]$Actual, [string]$Message) {
    $expectedText = (($Expected | Sort-Object) -join ",")
    $actualText = (($Actual | Sort-Object) -join ",")
    if ($expectedText -ne $actualText) {
        throw "$Message Expected [$expectedText], actual [$actualText]."
    }
}

$beforeExcel = Get-ExcelId
$sleeper = $null
try {
    $sleeper = Start-Process -FilePath "powershell.exe" -WindowStyle Hidden -PassThru -ArgumentList @(
        "-NoLogo",
        "-NoProfile",
        "-NonInteractive",
        "-Command",
        "Start-Sleep -Seconds 1"
    )
    $watchdogArgs = @("-NoLogo", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File", $watchdog, "-ProcessIds", $sleeper.Id, "-TimeoutSeconds", "30")
    & powershell.exe @watchdogArgs
    if ($LASTEXITCODE -ne 0) {
        throw "Watchdog rejected a valid non-Excel process unexpectedly."
    }
    if ($sleeper.HasExited) {
        throw "Watchdog terminated a non-Excel process."
    }

    $invalidRejected = $false
    try {
        & powershell.exe @("-NoLogo", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File", $watchdog, "-ProcessIds", "0", "-TimeoutSeconds", "30") 2>&1 | Out-Null
        $invalidRejected = ($LASTEXITCODE -ne 0)
    }
    catch {
        $invalidRejected = $true
    }
    if (-not $invalidRejected) {
        throw "Watchdog accepted an invalid PID unexpectedly."
    }
}
finally {
    if ($null -ne $sleeper) {
        $process = Get-Process -Id $sleeper.Id -ErrorAction SilentlyContinue
        if ($null -ne $process -and -not $process.HasExited) {
            Wait-Process -Id $sleeper.Id -Timeout 10
        }
    }
}

$afterExcel = Get-ExcelId
Assert-EqualId $beforeExcel $afterExcel "Watchdog changed the existing Excel PID set."
Write-Output "Protocol host watchdog safety tests passed."
