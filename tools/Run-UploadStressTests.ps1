[CmdletBinding()]
param(
    [long]$PayloadBytes = 1GB,
    [string]$OutputPath = "benchmarks/results/phase7-upload-stress.json"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$projectRoot = Split-Path -Parent $PSScriptRoot
$artifactDirectory = [IO.Path]::GetFullPath((Join-Path $projectRoot ".xlflow\upload-stress"))
$executablePath = Join-Path $artifactDirectory "testserver.exe"
$stdoutPath = Join-Path $artifactDirectory "stdout.jsonl"
$stderrPath = Join-Path $artifactDirectory "stderr.log"
$sourcePath = Join-Path $artifactDirectory "upload.bin"
$goPath = Join-Path $artifactDirectory "go.marker"
$resultsDirectory = [IO.Path]::GetFullPath((Join-Path $projectRoot "benchmarks\results"))
$resolvedOutput = [IO.Path]::GetFullPath((Join-Path $projectRoot $OutputPath))
$stagingOutput = "$resolvedOutput.$([guid]::NewGuid().ToString('N')).tmp"
$backupOutput = "$resolvedOutput.$([guid]::NewGuid().ToString('N')).bak"
$process = $null
$testProcess = $null
$ready = $null
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
$previousBaseUrl = [Environment]::GetEnvironmentVariable("VBA_HTTP_TEST_BASE_URL", "Process")
$previousStressPath = [Environment]::GetEnvironmentVariable("VBA_HTTP_UPLOAD_STRESS_PATH", "Process")
$previousGoPath = [Environment]::GetEnvironmentVariable("VBA_HTTP_UPLOAD_STRESS_GO_PATH", "Process")
$previousExpectedHash = [Environment]::GetEnvironmentVariable("VBA_HTTP_UPLOAD_STRESS_EXPECTED_HASH", "Process")

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

function New-PatternFile {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [string]$Path,
        [long]$Length
    )

    if (-not $PSCmdlet.ShouldProcess($Path, "Create deterministic $Length-byte upload source")) { return }
    $buffer = New-Object byte[] 65536
    for ($index = 0; $index -lt $buffer.Length; $index++) { $buffer[$index] = [byte]($index % 251) }
    $stream = [IO.File]::Create($Path)
    try {
        $remaining = $Length
        while ($remaining -gt 0) {
            $writeCount = [Math]::Min([long]$buffer.Length, $remaining)
            $stream.Write($buffer, 0, [int]$writeCount)
            $remaining -= $writeCount
        }
    }
    finally {
        $stream.Dispose()
    }
}

try {
    [void](New-Item -ItemType Directory -Path $artifactDirectory -Force)
    [void](New-Item -ItemType Directory -Path $resultsDirectory -Force)
    New-PatternFile $sourcePath $PayloadBytes
    $expectedHash = Get-Sha256Hex $sourcePath

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
            if ($line) { $ready = $line | ConvertFrom-Json; break }
        }
        Start-Sleep -Milliseconds 50
    }
    if ($null -eq $ready -or $ready.event -ne "ready" -or -not $ready.url) {
        throw "Test server did not publish a valid readiness event."
    }

    [Environment]::SetEnvironmentVariable("VBA_HTTP_TEST_BASE_URL", [string]$ready.url, "Process")
    [Environment]::SetEnvironmentVariable("VBA_HTTP_UPLOAD_STRESS_PATH", $sourcePath, "Process")
    [Environment]::SetEnvironmentVariable("VBA_HTTP_UPLOAD_STRESS_GO_PATH", $goPath, "Process")
    [Environment]::SetEnvironmentVariable("VBA_HTTP_UPLOAD_STRESS_EXPECTED_HASH", $expectedHash, "Process")
    if (Test-Path -LiteralPath $goPath) { Remove-Item -LiteralPath $goPath -Force }

    $xlflowPath = (Get-Command xlflow).Source
    $startInfo = New-Object Diagnostics.ProcessStartInfo
    $startInfo.FileName = $xlflowPath
    $startInfo.Arguments = "test --module WinHttpUploadStressTests --isolation module --json"
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
    if ($testProcess.ExitCode -ne 0) { throw "xlflow stress tests failed with exit code $($testProcess.ExitCode): $testStderr$testStdout" }
    $testResult = $testStdout | ConvertFrom-Json
    $stressTests = @($testResult.tests)
    if ($testResult.status -ne "ok" -or $stressTests.Count -ne 1 -or @($stressTests | Where-Object status -ne "passed").Count -ne 0) {
        throw "Stress suite did not return one passing test: $testStdout"
    }
    $afterSnapshot = Get-ExcelMemorySnapshot
    $workingSetAfter = $afterSnapshot.WorkingSetBytes
    $privateBytesAfter = $afterSnapshot.PrivateBytes
    $actualBytes = (Get-Item -LiteralPath $sourcePath).Length
    $actualHash = Get-Sha256Hex $sourcePath
    if ($actualBytes -ne $PayloadBytes) { throw "Source size changed: expected $PayloadBytes, got $actualBytes." }
    if ($actualHash -ne $expectedHash) { throw "Source SHA-256 changed during upload." }

    $stopwatch.Stop()
    $result = [ordered]@{
        schema_version = 1
        benchmark = "vba-http-phase7-upload-stress"
        server = [ordered]@{ base_url = [string]$ready.url; external_network = $false }
        parameters = [ordered]@{ payload_bytes = $PayloadBytes; endpoint = "/upload/hash"; chunk_bytes = 65536 }
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
    Write-Output "Phase 7 upload stress passed: $actualBytes bytes, SHA-256 $actualHash, peak working-set delta $($workingSetPeak - $workingSetBefore) bytes."
}
finally {
    [Environment]::SetEnvironmentVariable("VBA_HTTP_TEST_BASE_URL", $previousBaseUrl, "Process")
    [Environment]::SetEnvironmentVariable("VBA_HTTP_UPLOAD_STRESS_PATH", $previousStressPath, "Process")
    [Environment]::SetEnvironmentVariable("VBA_HTTP_UPLOAD_STRESS_GO_PATH", $previousGoPath, "Process")
    [Environment]::SetEnvironmentVariable("VBA_HTTP_UPLOAD_STRESS_EXPECTED_HASH", $previousExpectedHash, "Process")
    if ($null -ne $ready -and $ready.url) {
        try {
            Add-Type -AssemblyName System.Net.Http
            $httpClient = New-Object System.Net.Http.HttpClient
            $content = New-Object System.Net.Http.StringContent("")
            $response = $httpClient.PostAsync("$($ready.url)/__admin/shutdown", $content).GetAwaiter().GetResult()
            $response.Dispose()
            $content.Dispose()
            $httpClient.Dispose()
        }
        catch { Write-Warning "Could not request graceful test-server shutdown: $_" }
    }
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
