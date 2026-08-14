[CmdletBinding()]
param(
    [string]$OutputPath = "",
    [ValidateSet("Raw", "VBA-HTTP")]
    [string]$Implementation = "Raw"
)

$ErrorActionPreference = "Stop"
$projectRoot = Split-Path -Parent $PSScriptRoot
$implementationKey = if ($Implementation -eq "Raw") { "raw" } else { "vba-http" }
$macroName = if ($Implementation -eq "Raw") { "RawWinHttpBenchmark.RunRawBaseline" } else { "RawWinHttpBenchmark.RunVbaHttpBaseline" }
$expectedImplementationName = if ($Implementation -eq "Raw") { "Raw WinHttpRequest" } else { "VBA-HTTP" }
if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $OutputPath = if ($Implementation -eq "Raw") {
        "benchmarks/results/raw-winhttp-baseline.json"
    }
    else {
        "benchmarks/results/vba-http-buffered.json"
    }
}
$artifactDirectory = [System.IO.Path]::GetFullPath((Join-Path $projectRoot ".xlflow\$implementationKey-benchmark"))
$executablePath = Join-Path $artifactDirectory "testserver.exe"
$benchmarkWorkbook = Join-Path $artifactDirectory "$implementationKey-Benchmark.xlsm"
$stdoutPath = Join-Path $artifactDirectory "stdout.jsonl"
$stderrPath = Join-Path $artifactDirectory "stderr.log"
$resolvedOutput = [System.IO.Path]::GetFullPath((Join-Path $projectRoot $OutputPath))
$resultsDirectory = [System.IO.Path]::GetFullPath((Join-Path $projectRoot "benchmarks\results"))
if (-not $resolvedOutput.StartsWith($resultsDirectory + [System.IO.Path]::DirectorySeparatorChar, [System.StringComparison]::OrdinalIgnoreCase)) {
    throw "Benchmark output must be inside $resultsDirectory."
}
$stagingOutput = Join-Path $resultsDirectory (".{0}.{1}.tmp" -f [System.IO.Path]::GetFileName($resolvedOutput), [guid]::NewGuid().ToString("N"))
$backupOutput = Join-Path $resultsDirectory (".{0}.{1}.bak" -f [System.IO.Path]::GetFileName($resolvedOutput), [guid]::NewGuid().ToString("N"))
$process = $null
$client = $null
$ready = $null
$baselineExcelProcessIds = @()

function Get-ExcelProcessId {
    @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | ForEach-Object { [int]$_.Id })
}

function Wait-ForBenchmarkExcelCleanup {
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][int[]]$BaselineProcessIds
    )

    $deadline = [DateTime]::UtcNow.AddSeconds(30)
    while ([DateTime]::UtcNow -lt $deadline) {
        $remaining = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue |
                Where-Object { $BaselineProcessIds -notcontains $_.Id })
        if ($remaining.Count -eq 0) { return }
        Start-Sleep -Milliseconds 250
    }

    $remainingIds = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue |
            Where-Object { $BaselineProcessIds -notcontains $_.Id } |
            ForEach-Object { [int]$_.Id })
    if ($remainingIds.Count -gt 0) {
        throw "Benchmark left runner-created Excel process(es) running: $($remainingIds -join ','). No Excel process was terminated."
    }
}

try {
    $baselineExcelProcessIds = @(Get-ExcelProcessId)
    if ($baselineExcelProcessIds.Count -gt 0) {
        throw "Benchmark requires an Excel-exclusive window. Existing Excel PID(s): $($baselineExcelProcessIds -join ','). No Excel process was touched."
    }

    [void](New-Item -ItemType Directory -Path $artifactDirectory -Force)
    [void](New-Item -ItemType Directory -Path (Split-Path -Parent $resolvedOutput) -Force)
    $statusJson = & xlflow status --json
    if ($LASTEXITCODE -ne 0) {
        throw "xlflow status failed before the $Implementation benchmark."
    }
    $xlflowStatus = $statusJson | Out-String | ConvertFrom-Json
    if ($xlflowStatus.status -ne "ok" -or $xlflowStatus.coordination.recovery_required -or $xlflowStatus.state.src_newer_than_workbook) {
        throw "The development workbook is not a safe synchronized benchmark base. Push and verify source before benchmarking."
    }
    Copy-Item -LiteralPath (Join-Path $projectRoot "build\VBA-HTTP.xlsm") -Destination $benchmarkWorkbook

    Push-Location (Join-Path $projectRoot "tools\testserver")
    try {
        & go build -o $executablePath .
        if ($LASTEXITCODE -ne 0) {
            throw "go build failed with exit code $LASTEXITCODE."
        }
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
    $ready = $null
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

    Add-Type -AssemblyName System.Net.Http
    $client = New-Object System.Net.Http.HttpClient
    $resetContent = New-Object System.Net.Http.StringContent("")
    $reset = $client.PostAsync("$($ready.url)/__admin/reset", $resetContent).GetAwaiter().GetResult()
    if ([int]$reset.StatusCode -ne 204) {
        throw "Test server reset failed with status $([int]$reset.StatusCode)."
    }
    $reset.Dispose()
    $resetContent.Dispose()

    $runJson = & xlflow run $macroName `
        --input $benchmarkWorkbook `
        --no-save `
        --json `
        --arg "string:$($ready.url)" `
        --arg "string:$stagingOutput"
    if ($LASTEXITCODE -ne 0) {
        throw "$Implementation benchmark macro failed with exit code $LASTEXITCODE`: $($runJson | Out-String)"
    }
    Wait-ForBenchmarkExcelCleanup -BaselineProcessIds $baselineExcelProcessIds
    $runResult = $runJson | Out-String | ConvertFrom-Json
    if ($runResult.status -ne "ok") {
        throw "$Implementation benchmark macro returned status '$($runResult.status)'."
    }
    if (-not (Test-Path -LiteralPath $stagingOutput -PathType Leaf)) {
        throw "$Implementation benchmark did not create its result file."
    }

    $result = Get-Content -LiteralPath $stagingOutput -Raw | ConvertFrom-Json
    if ($result.schema_version -ne 1 -or
        $result.implementation.name -ne $expectedImplementationName -or
        $result.server.external_network -ne $false -or
        $result.parameters.download_bytes -lt 104857600 -or
        @($result.results).Count -ne 2) {
        throw "$Implementation benchmark result does not satisfy the baseline contract."
    }
    $latency = @($result.results | Where-Object scenario -eq "sequential_get")
    $download = @($result.results | Where-Object scenario -eq "buffered_download")
    if ($latency.Count -ne 1 -or $latency[0].status -ne 204 -or
        $download.Count -ne 1 -or $download[0].status -ne 200 -or
        $download[0].bytes -ne 104857600) {
        throw "$Implementation benchmark scenarios returned unexpected results."
    }

    if (Test-Path -LiteralPath $resolvedOutput -PathType Leaf) {
        [System.IO.File]::Replace($stagingOutput, $resolvedOutput, $backupOutput)
        Remove-Item -LiteralPath $backupOutput -Force
    }
    else {
        Move-Item -LiteralPath $stagingOutput -Destination $resolvedOutput
    }

    Write-Output "$Implementation benchmark completed: $resolvedOutput"
}
finally {
    try {
        if ($null -ne $baselineExcelProcessIds) {
            Wait-ForBenchmarkExcelCleanup -BaselineProcessIds $baselineExcelProcessIds
        }
    }
    catch {
        Write-Warning $_.Exception.Message
    }
    if ($null -ne $client -and $null -ne $ready -and $ready.url) {
        try {
            $shutdownContent = New-Object System.Net.Http.StringContent("")
            $shutdown = $client.PostAsync("$($ready.url)/__admin/shutdown", $shutdownContent).GetAwaiter().GetResult()
            $shutdown.Dispose()
            $shutdownContent.Dispose()
        }
        catch {
            Write-Warning "Could not request graceful test-server shutdown: $_"
        }
        $client.Dispose()
    }
    if ($null -ne $process -and -not $process.HasExited) {
        if (-not $process.WaitForExit(5000)) {
            Stop-Process -Id $process.Id -Force
            [void]$process.WaitForExit(5000)
        }
    }
    if ($null -ne $process) {
        $process.Dispose()
        $process = $null
    }
    if (Test-Path -LiteralPath $artifactDirectory) {
        $removed = $false
        for ($attempt = 1; $attempt -le 10 -and -not $removed; $attempt++) {
            try {
                Remove-Item -LiteralPath $artifactDirectory -Recurse -Force
                $removed = $true
            }
            catch {
                if ($attempt -eq 10) { throw }
                Start-Sleep -Milliseconds 100
            }
        }
    }
    if (Test-Path -LiteralPath $stagingOutput -PathType Leaf) {
        Remove-Item -LiteralPath $stagingOutput -Force
    }
    if (Test-Path -LiteralPath $backupOutput -PathType Leaf) {
        Remove-Item -LiteralPath $backupOutput -Force
    }
}
