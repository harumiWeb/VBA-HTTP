[CmdletBinding()]
param(
    [ValidateSet("all", "com-cancel", "com-deadline", "com-timeout", "native-download-cancel")]
    [string]$Scenario = "all",
    [ValidateRange(1, 1000)]
    [int]$Iterations = 25,
    [string]$OutputPath = "benchmarks/results/phase9-cancellation-stress.json"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$projectRoot = Split-Path -Parent $PSScriptRoot
$artifactDirectory = [IO.Path]::GetFullPath((Join-Path $projectRoot ".xlflow\cancellation-stress"))
$serverExecutable = Join-Path $artifactDirectory "testserver.exe"
$stdoutPath = Join-Path $artifactDirectory "server.stdout.jsonl"
$stderrPath = Join-Path $artifactDirectory "server.stderr.log"
$resultsDirectory = [IO.Path]::GetFullPath((Join-Path $projectRoot "benchmarks\results"))
$resolvedOutput = [IO.Path]::GetFullPath((Join-Path $projectRoot $OutputPath))
$stagingOutput = "$resolvedOutput.$([guid]::NewGuid().ToString('N')).tmp"
$backupOutput = "$resolvedOutput.$([guid]::NewGuid().ToString('N')).bak"
$serverProcess = $null
$serverReady = $null
$httpClient = $null
$previousBaseUrl = [Environment]::GetEnvironmentVariable("VBA_HTTP_TEST_BASE_URL", "Process")

if (-not $resolvedOutput.StartsWith($resultsDirectory + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) {
    throw "OutputPath must remain under benchmarks/results."
}

function Get-ExcelProcessIds {
    @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | ForEach-Object { [int]$_.Id })
}

function Get-ExcelSnapshot([int[]]$BaselineProcessIds) {
    $processes = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | Where-Object { $BaselineProcessIds -notcontains $_.Id })
    $handles = [long]0
    $workingSet = [long]0
    $privateBytes = [long]0
    $observed = $false
    foreach ($process in $processes) {
        try {
            $handles += [long]$process.HandleCount
            $workingSet += [long]$process.WorkingSet64
            $privateBytes += [long]$process.PrivateMemorySize64
            $observed = $true
        }
        catch {
            # A process can exit between enumeration and sampling.
        }
    }
    [ordered]@{
        observed = $observed
        process_count = $processes.Count
        handles = $handles
        working_set_bytes = $workingSet
        private_bytes = $privateBytes
    }
}

function Copy-Snapshot([object]$Snapshot) {
    [ordered]@{
        observed = [bool]$Snapshot.observed
        process_count = [int]$Snapshot.process_count
        handles = [long]$Snapshot.handles
        working_set_bytes = [long]$Snapshot.working_set_bytes
        private_bytes = [long]$Snapshot.private_bytes
    }
}

function Wait-ScenarioExcelCleanup([int[]]$BaselineProcessIds, [string]$ScenarioName) {
    $cleanupDeadline = [DateTime]::UtcNow.AddSeconds(30)
    while ([DateTime]::UtcNow -lt $cleanupDeadline) {
        $remaining = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | Where-Object { $BaselineProcessIds -notcontains $_.Id })
        if ($remaining.Count -eq 0) { return }
        Start-Sleep -Milliseconds 250
    }
    $remainingIds = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | Where-Object { $BaselineProcessIds -notcontains $_.Id } | ForEach-Object { $_.Id })
    if ($remainingIds.Count -gt 0) {
        throw "$ScenarioName left Excel test processes running after release: $($remainingIds -join ',')."
    }
}

function Write-ReleaseMarker([string]$Path) {
    [IO.File]::WriteAllText($Path, "release`n", [Text.UTF8Encoding]::new($false))
}

function Run-CancellationScenario([object]$Definition) {
    $scenarioRoot = Join-Path $artifactDirectory $Definition.Name
    [void](New-Item -ItemType Directory -Path $scenarioRoot -Force)
    $startPath = Join-Path $scenarioRoot "start.marker"
    $donePath = Join-Path $scenarioRoot "done.marker"
    $releasePath = Join-Path $scenarioRoot "release.marker"
    foreach ($path in @($startPath, $donePath, $releasePath)) {
        if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path -Force }
    }

    $previousIterations = [Environment]::GetEnvironmentVariable("VBA_HTTP_CANCELLATION_ITERATIONS", "Process")
    $previousStart = [Environment]::GetEnvironmentVariable("VBA_HTTP_CANCELLATION_START_PATH", "Process")
    $previousDone = [Environment]::GetEnvironmentVariable("VBA_HTTP_CANCELLATION_DONE_PATH", "Process")
    $previousRelease = [Environment]::GetEnvironmentVariable("VBA_HTTP_CANCELLATION_RELEASE_PATH", "Process")
    $testProcess = $null
    $before = $null
    $after = $null
    $idleAfter = $null
    $peak = $null
    $startedAt = $null
    $finishedAt = $null
    $baselineProcessIds = Get-ExcelProcessIds

    try {
        [Environment]::SetEnvironmentVariable("VBA_HTTP_CANCELLATION_ITERATIONS", [string]$Iterations, "Process")
        [Environment]::SetEnvironmentVariable("VBA_HTTP_CANCELLATION_START_PATH", $startPath, "Process")
        [Environment]::SetEnvironmentVariable("VBA_HTTP_CANCELLATION_DONE_PATH", $donePath, "Process")
        [Environment]::SetEnvironmentVariable("VBA_HTTP_CANCELLATION_RELEASE_PATH", $releasePath, "Process")

        $startInfo = New-Object Diagnostics.ProcessStartInfo
        $startInfo.FileName = (Get-Command xlflow).Source
        $startInfo.Arguments = "test --filter $($Definition.Filter) --isolation module --no-save --json"
        $startInfo.WorkingDirectory = $projectRoot
        $startInfo.UseShellExecute = $false
        $startInfo.RedirectStandardOutput = $true
        $startInfo.RedirectStandardError = $true
        $testProcess = [Diagnostics.Process]::Start($startInfo)
        $deadline = [DateTime]::UtcNow.AddMinutes(30)

        while (-not $testProcess.HasExited) {
            if ([DateTime]::UtcNow -ge $deadline) { throw "$($Definition.Name) cancellation stress exceeded the 30 minute timeout." }
            $snapshot = Get-ExcelSnapshot $baselineProcessIds
            if ($null -ne $startedAt -and $snapshot.observed) {
                if ($snapshot.handles -gt $peak.handles) { $peak.handles = $snapshot.handles }
                if ($snapshot.working_set_bytes -gt $peak.working_set_bytes) { $peak.working_set_bytes = $snapshot.working_set_bytes }
                if ($snapshot.private_bytes -gt $peak.private_bytes) { $peak.private_bytes = $snapshot.private_bytes }
            }
            if ($null -eq $before -and (Test-Path -LiteralPath $startPath -PathType Leaf) -and $snapshot.observed) {
                $before = Copy-Snapshot $snapshot
                $peak = Copy-Snapshot $snapshot
                $startedAt = [DateTime]::UtcNow
            }
            if ($null -ne $before -and $null -eq $after -and (Test-Path -LiteralPath $donePath -PathType Leaf)) {
                Start-Sleep -Milliseconds 1000
                $afterSnapshot = Get-ExcelSnapshot $baselineProcessIds
                if (-not $afterSnapshot.observed) { throw "$($Definition.Name) lost its Excel process before the after snapshot." }
                $after = Copy-Snapshot $afterSnapshot
                $idleAfter = $null
                # COM may release an aborted request asynchronously after the
                # release marker.  Keep the strict handle gate, but allow the
                # host up to 30 seconds to settle before declaring growth.
                for ($idleAttempt = 1; $idleAttempt -le 30; $idleAttempt++) {
                    Start-Sleep -Milliseconds 1000
                    $idleSnapshot = Get-ExcelSnapshot $baselineProcessIds
                    if (-not $idleSnapshot.observed) { throw "$($Definition.Name) lost its Excel process before the idle snapshot." }
                    $idleAfter = Copy-Snapshot $idleSnapshot
                    if (([long]$idleAfter.handles - [long]$before.handles) -le [long]$Definition.HandleDeltaLimit) { break }
                }
                $finishedAt = [DateTime]::UtcNow
                Write-ReleaseMarker $releasePath
            }
            Start-Sleep -Milliseconds 100
        }

        $testStdout = $testProcess.StandardOutput.ReadToEnd()
        $testStderr = $testProcess.StandardError.ReadToEnd()
        $testProcess.WaitForExit()
        if ($testProcess.ExitCode -ne 0) {
            throw "$($Definition.Name) cancellation stress failed with exit code $($testProcess.ExitCode): $testStderr$testStdout"
        }
        if ($null -eq $before -or $null -eq $after -or $null -eq $idleAfter -or $null -eq $peak) {
            throw "$($Definition.Name) did not publish complete process snapshots: $testStderr$testStdout"
        }
        $testResult = $testStdout | ConvertFrom-Json
        $tests = @($testResult.tests)
        if ($testResult.status -ne "ok" -or $tests.Count -ne 1 -or @($tests | Where-Object status -ne "passed").Count -ne 0) {
            throw "$($Definition.Name) did not return one passing test: $testStdout"
        }
        $handleDelta = [long]$after.handles - [long]$before.handles
        $idleHandleDelta = [long]$idleAfter.handles - [long]$before.handles
        if ($idleHandleDelta -gt [long]$Definition.HandleDeltaLimit) {
            throw "$($Definition.Name) persistent handle growth exceeded $($Definition.HandleDeltaLimit): $idleHandleDelta (before handles=$($before.handles), after handles=$($after.handles), idle handles=$($idleAfter.handles), before processes=$($before.process_count), after processes=$($after.process_count), idle processes=$($idleAfter.process_count))."
        }
        $scenarioResult = [ordered]@{
            scenario = $Definition.Name
            transport = $Definition.Transport
            iterations = $Iterations
            status = "passed"
            elapsed_ms = [math]::Round(([DateTime]$finishedAt - [DateTime]$startedAt).TotalMilliseconds, 3)
            handle_delta = $handleDelta
            idle_handle_delta = $idleHandleDelta
            handle_delta_limit = [int]$Definition.HandleDeltaLimit
            invariant = $Definition.Invariant
            process_before = $before
            process_peak = $peak
            process_after = $after
            process_idle_after = $idleAfter
        }
        Wait-ScenarioExcelCleanup $baselineProcessIds $Definition.Name
        $scenarioResult
    }
    finally {
        if ($null -ne $testProcess) {
            if (-not $testProcess.HasExited) {
                try { Write-ReleaseMarker $releasePath } catch { }
                Start-Sleep -Milliseconds 500
                if (-not $testProcess.HasExited) { $testProcess.Kill(); [void]$testProcess.WaitForExit(5000) }
            }
            $testProcess.Dispose()
        }
        foreach ($path in @($startPath, $donePath, $releasePath)) {
            if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path -Force }
        }
        [Environment]::SetEnvironmentVariable("VBA_HTTP_CANCELLATION_ITERATIONS", $previousIterations, "Process")
        [Environment]::SetEnvironmentVariable("VBA_HTTP_CANCELLATION_START_PATH", $previousStart, "Process")
        [Environment]::SetEnvironmentVariable("VBA_HTTP_CANCELLATION_DONE_PATH", $previousDone, "Process")
        [Environment]::SetEnvironmentVariable("VBA_HTTP_CANCELLATION_RELEASE_PATH", $previousRelease, "Process")
    }
}

try {
    [void](New-Item -ItemType Directory -Path $artifactDirectory -Force)
    [void](New-Item -ItemType Directory -Path $resultsDirectory -Force)
    Push-Location (Join-Path $projectRoot "tools\testserver")
    try {
        & go build -o $serverExecutable .
        if ($LASTEXITCODE -ne 0) { throw "go build failed with exit code $LASTEXITCODE." }
    }
    finally {
        Pop-Location
    }

    $serverProcess = Start-Process -FilePath $serverExecutable -ArgumentList "-listen", "127.0.0.1:0" -RedirectStandardOutput $stdoutPath -RedirectStandardError $stderrPath -WindowStyle Hidden -PassThru
    $readyDeadline = [DateTime]::UtcNow.AddSeconds(10)
    while ([DateTime]::UtcNow -lt $readyDeadline) {
        if ($serverProcess.HasExited) {
            $stderr = if (Test-Path -LiteralPath $stderrPath) { Get-Content -LiteralPath $stderrPath -Raw } else { "" }
            throw "Test server exited before readiness: $stderr"
        }
        if (Test-Path -LiteralPath $stdoutPath) {
            $line = Get-Content -LiteralPath $stdoutPath -TotalCount 1
            if ($line) { $serverReady = $line | ConvertFrom-Json; break }
        }
        Start-Sleep -Milliseconds 50
    }
    if ($null -eq $serverReady -or $serverReady.event -ne "ready" -or -not $serverReady.url) { throw "Test server did not publish readiness." }
    [Environment]::SetEnvironmentVariable("VBA_HTTP_TEST_BASE_URL", [string]$serverReady.url, "Process")

    $definitions = @(
        [pscustomobject]@{ Name = "com_active_cancellation"; Filter = "WinHttpCancellationStressTests.Test_CancellationStress_ComActiveCancellation"; Transport = "WinHttpComTransport"; HandleDeltaLimit = 32; Invariant = "four delayed COM requests become cancelled on every iteration, followed by a 204 recovery request" },
        [pscustomobject]@{ Name = "com_request_deadline"; Filter = "WinHttpCancellationStressTests.Test_CancellationStress_ComDeadline"; Transport = "WinHttpComTransport"; HandleDeltaLimit = 32; Invariant = "four delayed COM requests become timeout failures on every iteration, followed by a 204 recovery request" },
        [pscustomobject]@{ Name = "com_receive_timeout"; Filter = "WinHttpCancellationStressTests.Test_CancellationStress_ComReceiveTimeout"; Transport = "WinHttpComTransport"; HandleDeltaLimit = 32; Invariant = "repeated COM receive timeouts preserve a usable client and stay within the idle handle budget" },
        [pscustomobject]@{ Name = "native_download_cancellation"; Filter = "WinHttpCancellationStressTests.Test_CancellationStress_NativeDownloadCancellation"; Transport = "WinHttpNativeTransport"; HandleDeltaLimit = 8; Invariant = "64 KiB streaming download cancellation preserves the sentinel destination and temporary-file count" }
    )
    if ($Scenario -eq "all") {
        $selectedDefinitions = @($definitions)
    }
    else {
        $scenarioName = switch ($Scenario) {
            "com-cancel" { "com_active_cancellation" }
            "com-deadline" { "com_request_deadline" }
            "com-timeout" { "com_receive_timeout" }
            "native-download-cancel" { "native_download_cancellation" }
            default { "" }
        }
        $selectedDefinitions = @($definitions | Where-Object { $_.Name -eq $scenarioName })
    }
    if ($selectedDefinitions.Count -eq 0) { throw "No cancellation stress scenario was selected." }
    if ($Scenario -eq "all" -and $selectedDefinitions.Count -ne 4) { throw "The complete cancellation stress gate requires all four scenarios." }

    $scenarioResults = [System.Collections.Generic.List[object]]::new()
    foreach ($definition in $selectedDefinitions) {
        $scenarioResults.Add((Run-CancellationScenario $definition))
    }

    $sourceCommit = ((& git rev-parse HEAD) | Out-String).Trim()
    $result = [ordered]@{
        schema_version = 1
        benchmark = "phase9-cancellation-stress"
        source_commit = $sourceCommit
        environment = [ordered]@{
            platform = "Windows"
            powershell = $PSVersionTable.PSVersion.ToString()
            processor_architecture = [Environment]::GetEnvironmentVariable("PROCESSOR_ARCHITECTURE")
        }
        server = [ordered]@{ base_url = [string]$serverReady.url; external_network = $false }
        parameters = [ordered]@{
            iterations = $Iterations
            com_cancellation_requests = 4
            com_deadline_requests = 4
            com_timeout_receive_milliseconds = 1000
            com_timeout_delay_milliseconds = 10000
            native_download_bytes = 65536
            native_download_cancel_after_bytes = 65536
            idle_wait_ms = 1000
            idle_settle_max_ms = 30000
            handle_delta_limits = [ordered]@{ native = 8; com = 32 }
        }
        results = @($scenarioResults)
    }
    $json = $result | ConvertTo-Json -Depth 20
    [IO.File]::WriteAllText($stagingOutput, $json + [Environment]::NewLine, [Text.UTF8Encoding]::new($false))
    if (Test-Path -LiteralPath $resolvedOutput -PathType Leaf) {
        [IO.File]::Replace($stagingOutput, $resolvedOutput, $backupOutput)
        if (Test-Path -LiteralPath $backupOutput) { Remove-Item -LiteralPath $backupOutput -Force }
    }
    else {
        Move-Item -LiteralPath $stagingOutput -Destination $resolvedOutput
    }
    Write-Output "Phase 9 cancellation stress passed: $($scenarioResults.Count) scenario(s), $Iterations iterations each."
}
finally {
    [Environment]::SetEnvironmentVariable("VBA_HTTP_TEST_BASE_URL", $previousBaseUrl, "Process")
    if ($null -ne $serverReady -and $serverReady.url) {
        try {
            Add-Type -AssemblyName System.Net.Http
            $httpClient = New-Object System.Net.Http.HttpClient
            $content = New-Object System.Net.Http.StringContent("")
            $response = $httpClient.PostAsync("$($serverReady.url)/__admin/shutdown", $content).GetAwaiter().GetResult()
            $response.Dispose()
            $content.Dispose()
        }
        catch { Write-Warning "Could not request cancellation stress server shutdown: $_" }
    }
    if ($null -ne $httpClient) { $httpClient.Dispose() }
    if ($null -ne $serverProcess -and -not $serverProcess.HasExited) {
        if (-not $serverProcess.WaitForExit(5000)) { Stop-Process -Id $serverProcess.Id -Force; [void]$serverProcess.WaitForExit(5000) }
    }
    if ($null -ne $serverProcess) { $serverProcess.Dispose() }
    if (Test-Path -LiteralPath $stagingOutput) { Remove-Item -LiteralPath $stagingOutput -Force }
    if (Test-Path -LiteralPath $backupOutput) { Remove-Item -LiteralPath $backupOutput -Force }
    $resolvedXlflowRoot = [IO.Path]::GetFullPath((Join-Path $projectRoot ".xlflow"))
    if ($artifactDirectory.StartsWith($resolvedXlflowRoot + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase) -and (Test-Path -LiteralPath $artifactDirectory)) {
        Remove-Item -LiteralPath $artifactDirectory -Recurse -Force
    }
}
