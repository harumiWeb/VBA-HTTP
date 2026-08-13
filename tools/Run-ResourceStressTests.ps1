[CmdletBinding()]
param(
    [ValidateSet("all", "sequential", "scheduled")]
    [string]$Scenario = "all",
    [ValidateRange(1, 100000)]
    [int]$Iterations = 10000,
    [ValidateRange(1, 64)]
    [int]$Concurrency = 16,
    [string]$OutputPath = "benchmarks/results/phase9-resource-stress.json"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$projectRoot = Split-Path -Parent $PSScriptRoot
$artifactDirectory = [IO.Path]::GetFullPath((Join-Path $projectRoot ".xlflow\resource-stress"))
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

function Get-ExcelProcessId {
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
            Write-Debug "An Excel process exited during the resource-stress sample. The next sample will retry."
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

function Write-ReleaseMarker([string]$Path) {
    [IO.File]::WriteAllText($Path, "release`n", [Text.UTF8Encoding]::new($false))
}

function Invoke-ResourceScenario([string]$Name, [string]$Filter, [int]$HandleDeltaLimit) {
    $scenarioRoot = Join-Path $artifactDirectory $Name
    [void](New-Item -ItemType Directory -Path $scenarioRoot -Force)
    $startPath = Join-Path $scenarioRoot "start.marker"
    $donePath = Join-Path $scenarioRoot "done.marker"
    $releasePath = Join-Path $scenarioRoot "release.marker"
    foreach ($path in @($startPath, $donePath, $releasePath)) {
        if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path -Force }
    }

    $previousIterations = [Environment]::GetEnvironmentVariable("VBA_HTTP_RESOURCE_ITERATIONS", "Process")
    $previousConcurrency = [Environment]::GetEnvironmentVariable("VBA_HTTP_RESOURCE_CONCURRENCY", "Process")
    $previousStart = [Environment]::GetEnvironmentVariable("VBA_HTTP_RESOURCE_START_PATH", "Process")
    $previousDone = [Environment]::GetEnvironmentVariable("VBA_HTTP_RESOURCE_DONE_PATH", "Process")
    $previousRelease = [Environment]::GetEnvironmentVariable("VBA_HTTP_RESOURCE_RELEASE_PATH", "Process")
    $testProcess = $null
    $before = $null
    $after = $null
    $idleAfter = $null
    $peak = $null
    $startedAt = $null
    $finishedAt = $null
    $baselineProcessIds = Get-ExcelProcessId

    try {
        [Environment]::SetEnvironmentVariable("VBA_HTTP_RESOURCE_ITERATIONS", [string]$Iterations, "Process")
        [Environment]::SetEnvironmentVariable("VBA_HTTP_RESOURCE_CONCURRENCY", [string]$Concurrency, "Process")
        [Environment]::SetEnvironmentVariable("VBA_HTTP_RESOURCE_START_PATH", $startPath, "Process")
        [Environment]::SetEnvironmentVariable("VBA_HTTP_RESOURCE_DONE_PATH", $donePath, "Process")
        [Environment]::SetEnvironmentVariable("VBA_HTTP_RESOURCE_RELEASE_PATH", $releasePath, "Process")

        $startInfo = New-Object Diagnostics.ProcessStartInfo
        $startInfo.FileName = (Get-Command xlflow).Source
        $startInfo.Arguments = "test --filter $Filter --isolation module --no-save --json"
        $startInfo.WorkingDirectory = $projectRoot
        $startInfo.UseShellExecute = $false
        $startInfo.RedirectStandardOutput = $true
        $startInfo.RedirectStandardError = $true
        $testProcess = [Diagnostics.Process]::Start($startInfo)
        $deadline = [DateTime]::UtcNow.AddMinutes(30)

        while (-not $testProcess.HasExited) {
            if ([DateTime]::UtcNow -ge $deadline) { throw "$Name resource stress exceeded the 30 minute timeout." }
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
                if (-not $afterSnapshot.observed) { throw "$Name resource stress lost its Excel process before the after snapshot." }
                $after = Copy-Snapshot $afterSnapshot
                Start-Sleep -Milliseconds 1000
                $idleSnapshot = Get-ExcelSnapshot $baselineProcessIds
                if (-not $idleSnapshot.observed) { throw "$Name resource stress lost its Excel process before the idle snapshot." }
                $idleAfter = Copy-Snapshot $idleSnapshot
                $finishedAt = [DateTime]::UtcNow
                Write-ReleaseMarker $releasePath
            }
            Start-Sleep -Milliseconds 100
        }

        $testStdout = $testProcess.StandardOutput.ReadToEnd()
        $testStderr = $testProcess.StandardError.ReadToEnd()
        $testProcess.WaitForExit()
        if ($testProcess.ExitCode -ne 0) {
            throw "$Name resource stress failed with exit code $($testProcess.ExitCode): $testStderr$testStdout"
        }
        if ($null -eq $before -or $null -eq $after -or $null -eq $idleAfter -or $null -eq $peak) {
            throw "$Name resource stress did not publish complete process snapshots: $testStderr$testStdout"
        }
        $testResult = $testStdout | ConvertFrom-Json
        $tests = @($testResult.tests)
        if ($testResult.status -ne "ok" -or $tests.Count -ne 1 -or @($tests | Where-Object status -ne "passed").Count -ne 0) {
            throw "$Name resource stress did not return one passing test: $testStdout"
        }
        $handleDelta = [long]$after.handles - [long]$before.handles
        $idleHandleDelta = [long]$idleAfter.handles - [long]$before.handles
        if ($idleHandleDelta -gt $HandleDeltaLimit) {
            throw "$Name persistent handle growth exceeded ${HandleDeltaLimit}: $idleHandleDelta."
        }
        [ordered]@{
            scenario = $Name
            transport = if ($Name -eq "sequential_native") { "WinHttpNativeTransport" } else { "WinHttpComTransport" }
            iterations = $Iterations
            concurrency = if ($Name -eq "sequential_native") { 1 } else { $Concurrency }
            status = 204
            elapsed_ms = [math]::Round(([DateTime]$finishedAt - [DateTime]$startedAt).TotalMilliseconds, 3)
            handle_delta = $handleDelta
            idle_handle_delta = $idleHandleDelta
            handle_delta_limit = $HandleDeltaLimit
            process_before = $before
            process_peak = $peak
            process_after = $after
            process_idle_after = $idleAfter
        }
    }
    finally {
        if ($null -ne $testProcess) {
            if (-not $testProcess.HasExited) {
                try { Write-ReleaseMarker $releasePath } catch { Write-Debug "Could not publish the resource release marker during cleanup: $_" }
                Start-Sleep -Milliseconds 500
                if (-not $testProcess.HasExited) { $testProcess.Kill(); [void]$testProcess.WaitForExit(5000) }
            }
            $testProcess.Dispose()
        }
        foreach ($path in @($startPath, $donePath, $releasePath)) {
            if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path -Force }
        }
        [Environment]::SetEnvironmentVariable("VBA_HTTP_RESOURCE_ITERATIONS", $previousIterations, "Process")
        [Environment]::SetEnvironmentVariable("VBA_HTTP_RESOURCE_CONCURRENCY", $previousConcurrency, "Process")
        [Environment]::SetEnvironmentVariable("VBA_HTTP_RESOURCE_START_PATH", $previousStart, "Process")
        [Environment]::SetEnvironmentVariable("VBA_HTTP_RESOURCE_DONE_PATH", $previousDone, "Process")
        [Environment]::SetEnvironmentVariable("VBA_HTTP_RESOURCE_RELEASE_PATH", $previousRelease, "Process")
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

    $scenarioResults = [System.Collections.Generic.List[object]]::new()
    if ($Scenario -in @("all", "sequential")) {
        $scenarioResults.Add((Invoke-ResourceScenario "sequential_native" "WinHttpResourceStressTests.Test_ResourceStress_SequentialNative" 8))
    }
    if ($Scenario -in @("all", "scheduled")) {
        $scenarioResults.Add((Invoke-ResourceScenario "scheduled_com" "WinHttpResourceStressTests.Test_ResourceStress_ScheduledCom" 32))
    }
    if ($scenarioResults.Count -eq 0) { throw "No resource stress scenario was selected." }
    if ($Scenario -eq "all" -and $scenarioResults.Count -ne 2) { throw "The complete resource stress gate requires both scenarios." }

    $sourceCommit = ((& git rev-parse HEAD) | Out-String).Trim()
    $result = [ordered]@{
        schema_version = 1
        benchmark = "phase9-resource-stress"
        source_commit = $sourceCommit
        environment = [ordered]@{
            platform = "Windows"
            powershell = $PSVersionTable.PSVersion.ToString()
            processor_architecture = [Environment]::GetEnvironmentVariable("PROCESSOR_ARCHITECTURE")
        }
        server = [ordered]@{ base_url = [string]$serverReady.url; external_network = $false }
        parameters = [ordered]@{ iterations = $Iterations; scheduled_concurrency = $Concurrency; status_endpoint = "/status/204"; handle_delta_limits = [ordered]@{ sequential_native = 8; scheduled_com = 32 }; idle_wait_ms = 1000 }
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
    Write-Output "Phase 9 resource stress passed: $($scenarioResults.Count) scenario(s), $Iterations requests each."
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
        catch { Write-Warning "Could not request resource stress server shutdown: $_" }
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
