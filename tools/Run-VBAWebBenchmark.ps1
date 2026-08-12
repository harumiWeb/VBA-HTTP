[CmdletBinding()]
param(
    [string]$OutputPath = "benchmarks/results/vba-web-baseline.json",
    [string]$ReferencePath = "references/VBA-Web-v4.1.6"
)

$ErrorActionPreference = "Stop"
$projectRoot = Split-Path -Parent $PSScriptRoot
$artifactDirectory = [System.IO.Path]::GetFullPath((Join-Path $projectRoot ".xlflow\vba-web-benchmark"))
$serverExecutable = Join-Path $artifactDirectory "testserver.exe"
$benchmarkWorkbook = Join-Path $artifactDirectory "VBA-Web-Benchmark.xlsm"
$stdoutPath = Join-Path $artifactDirectory "stdout.jsonl"
$stderrPath = Join-Path $artifactDirectory "stderr.log"
$adapterPath = Join-Path $projectRoot "benchmarks\adapters\VBAWebBenchmark.bas"
$resolvedOutput = [System.IO.Path]::GetFullPath((Join-Path $projectRoot $OutputPath))
$resultsDirectory = [System.IO.Path]::GetFullPath((Join-Path $projectRoot "benchmarks\results"))
if (-not $resolvedOutput.StartsWith($resultsDirectory + [System.IO.Path]::DirectorySeparatorChar, [System.StringComparison]::OrdinalIgnoreCase)) {
    throw "Benchmark output must be inside $resultsDirectory."
}
$stagingOutput = Join-Path $resultsDirectory (".{0}.{1}.tmp" -f [System.IO.Path]::GetFileName($resolvedOutput), [guid]::NewGuid().ToString("N"))
$backupOutput = Join-Path $resultsDirectory (".{0}.{1}.bak" -f [System.IO.Path]::GetFileName($resolvedOutput), [guid]::NewGuid().ToString("N"))
$serverProcess = $null
$client = $null
$ready = $null

function Add-BenchmarkAdapter {
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$SourcePath
    )

    $excelApp = $null
    $workbook = $null
    $component = $null
    try {
        $excelApp = New-Object -ComObject Excel.Application
        $excelApp.Visible = $false
        $excelApp.DisplayAlerts = $false
        $excelApp.EnableEvents = $false
        $excelApp.AutomationSecurity = 3
        $workbook = $excelApp.Workbooks.Open($WorkbookPath, 0, $false)
        $component = $workbook.VBProject.VBComponents.Import($SourcePath)
        if ($component.Name -ne "VBAWebBenchmark") {
            throw "Imported benchmark adapter has unexpected component name '$($component.Name)'."
        }
        $workbook.Save()
        $workbook.Close($false)
    }
    finally {
        if ($null -ne $component) {
            [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($component)
            $component = $null
        }
        if ($null -ne $workbook) {
            try { $workbook.Close($false) } catch { Write-Verbose "Workbook was already closed." }
            [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($workbook)
            $workbook = $null
        }
        if ($null -ne $excelApp) {
            $excelApp.Quit()
            [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excelApp)
            $excelApp = $null
        }
    }
}

try {
    [void](New-Item -ItemType Directory -Path $artifactDirectory -Force)
    [void](New-Item -ItemType Directory -Path (Split-Path -Parent $resolvedOutput) -Force)

    $setupJson = & powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot "Setup-VBAWeb.ps1") -ReferencePath $ReferencePath -VerifyOnly
    if ($LASTEXITCODE -ne 0) {
        throw "VBA-Web setup failed with exit code $LASTEXITCODE."
    }
    $setup = $setupJson | Out-String | ConvertFrom-Json
    if ($setup.commit -ne "cefc320acc5372e0b86eed1d20eb3f31b331d598") {
        throw "VBA-Web setup returned an unexpected commit."
    }
    Copy-Item -LiteralPath $setup.workbook -Destination $benchmarkWorkbook
    Add-BenchmarkAdapter -WorkbookPath $benchmarkWorkbook -SourcePath $adapterPath

    Push-Location (Join-Path $projectRoot "tools\testserver")
    try {
        & go build -o $serverExecutable .
        if ($LASTEXITCODE -ne 0) {
            throw "go build failed with exit code $LASTEXITCODE."
        }
    }
    finally {
        Pop-Location
    }

    $serverProcess = Start-Process `
        -FilePath $serverExecutable `
        -ArgumentList "-listen", "127.0.0.1:0" `
        -RedirectStandardOutput $stdoutPath `
        -RedirectStandardError $stderrPath `
        -WindowStyle Hidden `
        -PassThru

    $deadline = [DateTime]::UtcNow.AddSeconds(10)
    while ([DateTime]::UtcNow -lt $deadline) {
        if ($serverProcess.HasExited) {
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

    $runJson = & xlflow run VBAWebBenchmark.RunVBAWebBaseline `
        --input $benchmarkWorkbook `
        --headless `
        --no-save `
        --timeout 10m `
        --json `
        --arg "string:$($ready.url)" `
        --arg "string:$stagingOutput"
    if ($LASTEXITCODE -ne 0) {
        throw "VBA-Web benchmark macro failed with exit code $LASTEXITCODE`: $($runJson | Out-String)"
    }
    $runResult = $runJson | Out-String | ConvertFrom-Json
    if ($runResult.status -ne "ok") {
        throw "VBA-Web benchmark macro returned status '$($runResult.status)'."
    }
    if (-not (Test-Path -LiteralPath $stagingOutput -PathType Leaf)) {
        throw "VBA-Web benchmark did not create its result file."
    }

    $result = Get-Content -LiteralPath $stagingOutput -Raw | ConvertFrom-Json
    if ($result.schema_version -ne 1 -or
        $result.implementation.name -ne "VBA-Web" -or
        $result.implementation.version -ne "4.1.6" -or
        $result.implementation.source_commit -ne $setup.commit -or
        $result.server.external_network -ne $false -or
        $result.parameters.download_bytes -lt 104857600 -or
        @($result.results).Count -ne 2) {
        throw "VBA-Web benchmark result does not satisfy the baseline contract."
    }
    $latency = @($result.results | Where-Object scenario -eq "sequential_get")
    $download = @($result.results | Where-Object scenario -eq "buffered_download")
    if ($latency.Count -ne 1 -or $latency[0].status -ne 204 -or
        $download.Count -ne 1 -or $download[0].status -ne 200 -or
        $download[0].bytes -ne 104857600) {
        throw "VBA-Web benchmark scenarios returned unexpected results."
    }

    if (Test-Path -LiteralPath $resolvedOutput -PathType Leaf) {
        [System.IO.File]::Replace($stagingOutput, $resolvedOutput, $backupOutput)
        Remove-Item -LiteralPath $backupOutput -Force
    }
    else {
        Move-Item -LiteralPath $stagingOutput -Destination $resolvedOutput
    }

    Write-Output "VBA-Web benchmark completed: $resolvedOutput"
}
finally {
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
    if ($null -ne $serverProcess -and -not $serverProcess.HasExited) {
        if (-not $serverProcess.WaitForExit(5000)) {
            Stop-Process -Id $serverProcess.Id -Force
            [void]$serverProcess.WaitForExit(5000)
        }
    }
    if ($null -ne $serverProcess) {
        $serverProcess.Dispose()
        $serverProcess = $null
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
