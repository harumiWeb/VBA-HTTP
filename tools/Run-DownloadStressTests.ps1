[CmdletBinding()]
param(
    [long]$PayloadBytes = 1GB,
    [string]$OutputPath = "benchmarks/results/phase6-download-stress.json"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$projectRoot = Split-Path -Parent $PSScriptRoot
$artifactDirectory = [IO.Path]::GetFullPath((Join-Path $projectRoot ".xlflow\download-stress"))
$executablePath = Join-Path $artifactDirectory "testserver.exe"
$stdoutPath = Join-Path $artifactDirectory "stdout.jsonl"
$stderrPath = Join-Path $artifactDirectory "stderr.log"
$destinationPath = Join-Path $artifactDirectory "download.bin"
$goPath = Join-Path $artifactDirectory "go.marker"
$resultsDirectory = [IO.Path]::GetFullPath((Join-Path $projectRoot "benchmarks\results"))
$resolvedOutput = [IO.Path]::GetFullPath((Join-Path $projectRoot $OutputPath))
$stagingOutput = "$resolvedOutput.$([guid]::NewGuid().ToString('N')).tmp"
$backupOutput = "$resolvedOutput.$([guid]::NewGuid().ToString('N')).bak"
$process = $null
$testProcess = $null
$client = $null
$ready = $null
$previousBaseUrl = [Environment]::GetEnvironmentVariable("VBA_HTTP_TEST_BASE_URL", "Process")
$previousStressPath = [Environment]::GetEnvironmentVariable("VBA_HTTP_DOWNLOAD_STRESS_PATH", "Process")
$previousGoPath = [Environment]::GetEnvironmentVariable("VBA_HTTP_DOWNLOAD_STRESS_GO_PATH", "Process")
$stopwatch = [Diagnostics.Stopwatch]::StartNew()
$workingSetBefore = [long]0
$workingSetPeak = [long]0
$workingSetAfter = [long]0
$privateBytesBefore = [long]0
$privateBytesPeak = [long]0
$privateBytesAfter = [long]0
$observedExcel = $false
$goSignaled = $false
$excelSeenAt = $null

if ($PayloadBytes -lt 1GB -or $PayloadBytes -gt 2GB) {
    throw "PayloadBytes must be between 1 GiB and 2 GiB."
}
if (-not $resolvedOutput.StartsWith($resultsDirectory + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) {
    throw "OutputPath must remain under benchmarks/results."
}

function Get-ExcelMemorySnapshot {
    $processes = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue)
    $working = [long]0
    $private = [long]0
    foreach ($excel in $processes) {
        $working += [long]$excel.WorkingSet64
        $private += [long]$excel.PrivateMemorySize64
    }
    [pscustomobject]@{
        Count = $processes.Count
        WorkingSetBytes = $working
        PrivateBytes = $private
    }
}

function Get-Sha256Hex([string]$Path) {
    $sha = [Security.Cryptography.SHA256]::Create()
    $stream = [IO.File]::OpenRead($Path)
    try {
        return ([BitConverter]::ToString($sha.ComputeHash($stream))).Replace("-", "").ToLowerInvariant()
    }
    finally {
        $stream.Dispose()
        $sha.Dispose()
    }
}

try {
    [void](New-Item -ItemType Directory -Path $artifactDirectory -Force)
    [void](New-Item -ItemType Directory -Path $resultsDirectory -Force)
    Push-Location (Join-Path $projectRoot "tools\testserver")
    try {
        & go build -o $executablePath .
        if ($LASTEXITCODE -ne 0) { throw "go build failed with exit code $LASTEXITCODE." }
    }
    finally {
        Pop-Location
    }

    $process = Start-Process `
        -FilePath $executablePath `
        -ArgumentList "-listen", "127.0.0.1:0" `
        -RedirectStandardOutput $stdoutPath `
        -RedirectStandardError $stderrPath `
        -WindowStyle Hidden `
        -PassThru

    $deadline = [DateTime]::UtcNow.AddSeconds(10)
    while ([DateTime]::UtcNow -lt $deadline) {
        if ($process.HasExited) {
            $stderr = if (Test-Path -LiteralPath $stderrPath) { Get-Content -LiteralPath $stderrPath -Raw } else { "" }
            throw "Test server exited before readiness: $stderr"
        }
        if (Test-Path -LiteralPath $stdoutPath) {
            $line = Get-Content -LiteralPath $stdoutPath -TotalCount 1
            if ($line) {
                $ready = $line | ConvertFrom-Json
                break
            }
        }
        Start-Sleep -Milliseconds 50
    }
    if ($null -eq $ready -or $ready.event -ne "ready" -or -not $ready.url) {
        throw "Test server did not publish a valid readiness event."
    }

    [Environment]::SetEnvironmentVariable("VBA_HTTP_TEST_BASE_URL", [string]$ready.url, "Process")
    [Environment]::SetEnvironmentVariable("VBA_HTTP_DOWNLOAD_STRESS_PATH", $destinationPath, "Process")
    [Environment]::SetEnvironmentVariable("VBA_HTTP_DOWNLOAD_STRESS_GO_PATH", $goPath, "Process")
    if (Test-Path -LiteralPath $goPath) { Remove-Item -LiteralPath $goPath -Force }
    $workingSetPeak = [long]0
    $privateBytesPeak = [long]0

    $xlflowPath = (Get-Command xlflow).Source
    $startInfo = New-Object Diagnostics.ProcessStartInfo
    $startInfo.FileName = $xlflowPath
    $startInfo.Arguments = "test --tag stress --isolation module --json"
    $startInfo.WorkingDirectory = $projectRoot
    $startInfo.UseShellExecute = $false
    $startInfo.RedirectStandardOutput = $true
    $startInfo.RedirectStandardError = $true
    $testProcess = [Diagnostics.Process]::Start($startInfo)

    while (-not $testProcess.HasExited) {
        $snapshot = Get-ExcelMemorySnapshot
        if ($snapshot.Count -gt 0) {
            $observedExcel = $true
            if ($null -eq $excelSeenAt) { $excelSeenAt = [DateTime]::UtcNow }
        }
        if (-not $goSignaled -and $snapshot.Count -gt 0 -and $null -ne $excelSeenAt -and ([DateTime]::UtcNow - $excelSeenAt).TotalSeconds -ge 3) {
            $workingSetBefore = $snapshot.WorkingSetBytes
            $workingSetPeak = $workingSetBefore
            $privateBytesBefore = $snapshot.PrivateBytes
            $privateBytesPeak = $snapshot.PrivateBytes
            [IO.File]::WriteAllText($goPath, "go", [Text.UTF8Encoding]::new($false))
            $goSignaled = $true
        }
        if ($snapshot.WorkingSetBytes -gt $workingSetPeak) { $workingSetPeak = $snapshot.WorkingSetBytes }
        if ($snapshot.PrivateBytes -gt $privateBytesPeak) { $privateBytesPeak = $snapshot.PrivateBytes }
        Start-Sleep -Milliseconds 100
    }
    $testStdout = $testProcess.StandardOutput.ReadToEnd()
    $testStderr = $testProcess.StandardError.ReadToEnd()
    $testProcess.WaitForExit()
    if (-not $goSignaled) { throw "Stress test never reached the memory gate." }
    if ($testProcess.ExitCode -ne 0) {
        throw "xlflow stress tests failed with exit code $($testProcess.ExitCode): $testStderr$testStdout"
    }
    $testResult = $testStdout | ConvertFrom-Json
    $stressTests = @($testResult.tests)
    if ($testResult.status -ne "ok" -or $stressTests.Count -ne 1 -or @($stressTests | Where-Object status -ne "passed").Count -ne 0) {
        throw "Stress suite did not return one passing test: $testStdout"
    }
    $afterSnapshot = Get-ExcelMemorySnapshot
    $workingSetAfter = $afterSnapshot.WorkingSetBytes
    $privateBytesAfter = $afterSnapshot.PrivateBytes
    if (-not (Test-Path -LiteralPath $destinationPath -PathType Leaf)) {
        throw "Stress download did not create its destination file."
    }

    Add-Type -AssemblyName System.Net.Http
    $client = New-Object System.Net.Http.HttpClient
    $expected = $client.GetStringAsync("$($ready.url)/sha256/$PayloadBytes").GetAwaiter().GetResult() | ConvertFrom-Json
    $actualBytes = (Get-Item -LiteralPath $destinationPath).Length
    $actualHash = Get-Sha256Hex $destinationPath
    if ([long]$expected.bytes -ne $actualBytes) { throw "Stress download size mismatch: expected $($expected.bytes), got $actualBytes." }
    if ($expected.digest -ne $actualHash) { throw "Stress download SHA-256 mismatch." }

    $stopwatch.Stop()
    $result = [ordered]@{
        schema_version = 1
        benchmark = "vba-http-phase6-download-stress"
        server = [ordered]@{ base_url = [string]$ready.url; external_network = $false }
        parameters = [ordered]@{ payload_bytes = $PayloadBytes; endpoint = "/stream/$PayloadBytes"; chunk_bytes = 65536 }
        result = [ordered]@{
            status = 200
            bytes = $actualBytes
            sha256 = $actualHash
            elapsed_ms = [math]::Round($stopwatch.Elapsed.TotalMilliseconds, 3)
            process_memory = [ordered]@{
                observed_excel = $observedExcel
                working_set_before = $workingSetBefore
                working_set_peak = $workingSetPeak
                working_set_after = $workingSetAfter
                private_bytes_before = $privateBytesBefore
                private_bytes_peak = $privateBytesPeak
                private_bytes_after = $privateBytesAfter
                peak_delta = ($workingSetPeak - $workingSetBefore)
                private_peak_delta = ($privateBytesPeak - $privateBytesBefore)
            }
        }
    }
    [IO.File]::WriteAllText($stagingOutput, ($result | ConvertTo-Json -Depth 10) + [Environment]::NewLine, [Text.UTF8Encoding]::new($false))
    if (Test-Path -LiteralPath $resolvedOutput -PathType Leaf) {
        [IO.File]::Replace($stagingOutput, $resolvedOutput, $backupOutput)
        if (Test-Path -LiteralPath $backupOutput) { Remove-Item -LiteralPath $backupOutput -Force }
    }
    else {
        Move-Item -LiteralPath $stagingOutput -Destination $resolvedOutput
    }
    Write-Output "Phase 6 download stress passed: $actualBytes bytes, SHA-256 $actualHash, peak working-set delta $($workingSetPeak - $workingSetBefore) bytes."
}
finally {
    [Environment]::SetEnvironmentVariable("VBA_HTTP_TEST_BASE_URL", $previousBaseUrl, "Process")
    [Environment]::SetEnvironmentVariable("VBA_HTTP_DOWNLOAD_STRESS_PATH", $previousStressPath, "Process")
    [Environment]::SetEnvironmentVariable("VBA_HTTP_DOWNLOAD_STRESS_GO_PATH", $previousGoPath, "Process")
    if ($null -ne $ready -and $ready.url) {
        try {
            if ($null -eq $client) {
                Add-Type -AssemblyName System.Net.Http
                $client = New-Object System.Net.Http.HttpClient
            }
            $content = New-Object System.Net.Http.StringContent("")
            $response = $client.PostAsync("$($ready.url)/__admin/shutdown", $content).GetAwaiter().GetResult()
            $response.Dispose()
            $content.Dispose()
        }
        catch {
            Write-Warning "Could not request graceful test-server shutdown: $_"
        }
    }
    if ($null -ne $client) { $client.Dispose() }
    if ($null -ne $testProcess) {
        if (-not $testProcess.HasExited) { $testProcess.Kill(); [void]$testProcess.WaitForExit(5000) }
        $testProcess.Dispose()
    }
    if ($null -ne $process -and -not $process.HasExited) {
        if (-not $process.WaitForExit(5000)) { Stop-Process -Id $process.Id -Force; [void]$process.WaitForExit(5000) }
    }
    if ($null -ne $process) { $process.Dispose() }
    if (Test-Path -LiteralPath $stagingOutput) { Remove-Item -LiteralPath $stagingOutput -Force }
    if (Test-Path -LiteralPath $backupOutput) { Remove-Item -LiteralPath $backupOutput -Force }

    $resolvedRoot = [IO.Path]::GetFullPath((Join-Path $projectRoot ".xlflow"))
    if ($artifactDirectory.StartsWith($resolvedRoot + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase) -and (Test-Path -LiteralPath $artifactDirectory)) {
        Remove-Item -LiteralPath $artifactDirectory -Recurse -Force
    }
}
