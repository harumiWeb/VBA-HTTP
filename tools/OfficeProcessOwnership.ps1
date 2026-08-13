Set-StrictMode -Version Latest

if ($null -eq ("VBAHttpOfficeWindow" -as [type])) {
    Add-Type @"
using System;
using System.Runtime.InteropServices;
public static class VBAHttpOfficeWindow {
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
}
"@
}

function Get-ExcelProcessIds {
    return @(
        Get-Process -Name EXCEL -ErrorAction SilentlyContinue |
            ForEach-Object { [int]$_.Id }
    )
}

function Get-OwnedExcelProcessId($Excel, [int[]]$BaselineIds, [string]$Purpose) {
    $windowHandle = [IntPtr]$Excel.Hwnd
    if ($windowHandle -eq [IntPtr]::Zero) {
        throw "Could not obtain the Excel automation window handle for $Purpose; refusing to touch any Excel process."
    }
    [uint32]$processId = 0
    [void][VBAHttpOfficeWindow]::GetWindowThreadProcessId($windowHandle, [ref]$processId)
    if ($processId -eq 0 -or $BaselineIds -contains [int]$processId) {
        throw "Excel automation did not prove a new process for $Purpose; refusing to touch an existing instance."
    }
    $process = Get-Process -Id ([int]$processId) -ErrorAction SilentlyContinue
    if ($null -eq $process -or $process.ProcessName -ne "EXCEL") {
        throw "Excel automation window did not map to a live Excel process for $Purpose."
    }
    return @([int]$processId)
}

function Stop-OwnedExcelProcesses([int[]]$OwnedIds, [string]$Purpose, [int]$GraceSeconds = 5) {
    $deadline = [DateTime]::UtcNow.AddSeconds($GraceSeconds)
    do {
        $liveOwned = @($OwnedIds | ForEach-Object {
                $candidate = Get-Process -Id $_ -ErrorAction SilentlyContinue
                if ($null -ne $candidate -and $candidate.ProcessName -eq "EXCEL") { $candidate }
            })
        if ($liveOwned.Count -eq 0) { return }
        Start-Sleep -Milliseconds 100
    } while ([DateTime]::UtcNow -lt $deadline)

    foreach ($processId in @($OwnedIds)) {
        $process = Get-Process -Id $processId -ErrorAction SilentlyContinue
        if ($null -ne $process -and $process.ProcessName -eq "EXCEL") {
            try {
                Stop-Process -Id $processId -Force -ErrorAction Stop
                Write-Warning "Stopped owned Excel PID $processId after $Purpose cleanup did not complete."
            }
            catch {
                Write-Warning "Could not stop owned Excel PID $processId after $Purpose cleanup: $_"
            }
        }
    }
}
