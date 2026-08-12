[CmdletBinding()]
param(
    [string]$OutputPath = "benchmarks/results/vba-http-concurrency.json"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$projectRoot = Split-Path -Parent $PSScriptRoot
$artifactDirectory = [IO.Path]::GetFullPath((Join-Path $projectRoot ".xlflow\concurrency-benchmark"))
$serverExecutable = Join-Path $artifactDirectory "testserver.exe"
$benchmarkWorkbook = Join-Path $artifactDirectory "VBA-HTTP-Concurrency-Benchmark.xlsm"
$stdoutPath = Join-Path $artifactDirectory "stdout.jsonl"
$stderrPath = Join-Path $artifactDirectory "stderr.log"
$resultsDirectory = [IO.Path]::GetFullPath((Join-Path $projectRoot "benchmarks\results"))
$resolvedOutput = [IO.Path]::GetFullPath((Join-Path $projectRoot $OutputPath))
if (-not $resolvedOutput.StartsWith($resultsDirectory + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) {
    throw "Benchmark output must be inside $resultsDirectory."
}
$stagingOutput = Join-Path $resultsDirectory (".{0}.{1}.tmp" -f [IO.Path]::GetFileName($resolvedOutput), [guid]::NewGuid().ToString("N"))
$backupOutput = Join-Path $resultsDirectory (".{0}.{1}.bak" -f [IO.Path]::GetFileName($resolvedOutput), [guid]::NewGuid().ToString("N"))
$serverProcess = $null
$client = $null
$ready = $null

try {
    [void](New-Item -ItemType Directory -Path $artifactDirectory -Force)
    $status = (& xlflow status --json) | Out-String | ConvertFrom-Json
    if ($LASTEXITCODE -ne 0 -or $status.status -ne "ok" -or $status.coordination.recovery_required -or $status.state.src_newer_than_workbook) {
        throw "The development workbook is not a safe synchronized benchmark base."
    }
    Copy-Item -LiteralPath (Join-Path $projectRoot "build\VBA-HTTP.xlsm") -Destination $benchmarkWorkbook

    Push-Location (Join-Path $projectRoot "tools\testserver")
    try {
        & go build -o $serverExecutable .
        if ($LASTEXITCODE -ne 0) { throw "go build failed with exit code $LASTEXITCODE." }
    }
    finally {
        Pop-Location
    }
    $serverProcess = Start-Process -FilePath $serverExecutable -ArgumentList "-listen", "127.0.0.1:0" -RedirectStandardOutput $stdoutPath -RedirectStandardError $stderrPath -WindowStyle Hidden -PassThru

    $deadline = [DateTime]::UtcNow.AddSeconds(10)
    while ([DateTime]::UtcNow -lt $deadline) {
        if ($serverProcess.HasExited) {
            $stderr = if (Test-Path -LiteralPath $stderrPath) { Get-Content -LiteralPath $stderrPath -Raw } else { "" }
            throw "Test server exited before readiness: $stderr"
        }
        if (Test-Path -LiteralPath $stdoutPath) {
            $line = Get-Content -LiteralPath $stdoutPath -TotalCount 1
            if ($line) { $ready = $line | ConvertFrom-Json; break }
        }
        Start-Sleep -Milliseconds 50
    }
    if ($null -eq $ready -or $ready.event -ne "ready" -or -not $ready.url) { throw "Test server did not publish readiness." }

    # Do not pass --headless: xlflow intentionally flags the scheduler's opt-in DoEvents boundary
    # statically even though this benchmark sets YieldToHost=False and remains unattended.
    $runJson = & xlflow run ConcurrencyBenchmark.RunConcurrencyBaseline --input $benchmarkWorkbook --no-save --timeout 2m --json --arg "string:$($ready.url)" --arg "string:$stagingOutput"
    if ($LASTEXITCODE -ne 0) { throw "Concurrency benchmark macro failed: $($runJson | Out-String)" }
    $runResult = $runJson | Out-String | ConvertFrom-Json
    if ($runResult.status -ne "ok" -or -not (Test-Path -LiteralPath $stagingOutput -PathType Leaf)) { throw "Concurrency benchmark did not produce a result." }

    $result = Get-Content -LiteralPath $stagingOutput -Raw | ConvertFrom-Json
    if ($result.schema_version -ne 1 -or $result.benchmark -ne "bounded-concurrency" -or
        $result.parameters.requests -ne 100 -or $result.parameters.delay_ms -ne 100 -or
        $result.parameters.concurrent_concurrency -ne 16 -or $result.results.speedup -lt 6 -or
        $result.results.server_max_in_flight -gt 16) {
        throw "Concurrency benchmark result does not satisfy the Phase 3 gate."
    }

    if (Test-Path -LiteralPath $resolvedOutput -PathType Leaf) {
        [IO.File]::Replace($stagingOutput, $resolvedOutput, $backupOutput)
        Remove-Item -LiteralPath $backupOutput -Force
    }
    else {
        Move-Item -LiteralPath $stagingOutput -Destination $resolvedOutput
    }
    Write-Output "Concurrency benchmark completed: $resolvedOutput"
}
finally {
    if ($null -ne $ready -and $ready.url) {
        try {
            $client = New-Object System.Net.WebClient
            [void]$client.UploadString("$($ready.url)/__admin/shutdown", "POST", "")
        }
        catch { Write-Warning "Could not request benchmark server shutdown: $_" }
    }
    if ($null -ne $client) { $client.Dispose() }
    if ($null -ne $serverProcess -and -not $serverProcess.HasExited) {
        if (-not $serverProcess.WaitForExit(5000)) { Stop-Process -Id $serverProcess.Id -Force }
    }
    if ($null -ne $serverProcess) { $serverProcess.Dispose() }
    $resolvedXlflowRoot = [IO.Path]::GetFullPath((Join-Path $projectRoot ".xlflow"))
    if ($artifactDirectory.StartsWith($resolvedXlflowRoot + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase) -and (Test-Path -LiteralPath $artifactDirectory)) {
        Remove-Item -LiteralPath $artifactDirectory -Recurse -Force
    }
    if (Test-Path -LiteralPath $stagingOutput) { Remove-Item -LiteralPath $stagingOutput -Force }
    if (Test-Path -LiteralPath $backupOutput) { Remove-Item -LiteralPath $backupOutput -Force }
}
