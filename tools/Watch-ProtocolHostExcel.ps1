[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$ProcessIds,
    [ValidateRange(30, 3600)]
    [int]$TimeoutSeconds = 90
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$ownedIds = @(
    $ProcessIds -split "," |
        ForEach-Object {
            $parsed = 0
            if (-not [int]::TryParse($_.Trim(), [ref]$parsed) -or $parsed -le 0) {
                throw "Invalid owned Excel PID: $_"
            }
            $parsed
        }
)
if ($ownedIds.Count -eq 0) { throw "At least one owned Excel PID is required." }

$deadline = [DateTime]::UtcNow.AddSeconds($TimeoutSeconds)
while ([DateTime]::UtcNow -lt $deadline) {
    $remaining = @(
        foreach ($id in $ownedIds) {
            $process = Get-Process -Id $id -ErrorAction SilentlyContinue
            if ($null -ne $process -and -not $process.HasExited -and $process.ProcessName -ieq "EXCEL") {
                $process
            }
        }
    )
    if ($remaining.Count -eq 0) { exit 0 }
    Start-Sleep -Milliseconds 250
}

foreach ($id in $ownedIds) {
    $process = Get-Process -Id $id -ErrorAction SilentlyContinue
    if ($null -ne $process -and -not $process.HasExited -and $process.ProcessName -ieq "EXCEL") {
        Stop-Process -Id $id -Force
    }
}
