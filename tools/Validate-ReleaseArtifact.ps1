[CmdletBinding()]
param(
    [string]$ArtifactPath = "build/Release/VBA-HTTP.xlsm"
)

$ErrorActionPreference = "Stop"
$projectRoot = Split-Path -Parent $PSScriptRoot
$resolvedArtifact = [System.IO.Path]::GetFullPath((Join-Path $projectRoot $ArtifactPath))
$manifestPath = "$resolvedArtifact.build.json"
$policyPath = Join-Path $PSScriptRoot "build-component-policy.json"

if (-not (Test-Path -LiteralPath $resolvedArtifact -PathType Leaf)) {
    throw "Release artifact does not exist: $resolvedArtifact"
}
if (-not (Test-Path -LiteralPath $manifestPath -PathType Leaf)) {
    throw "Release manifest does not exist: $manifestPath"
}

$policy = Get-Content -LiteralPath $policyPath -Raw | ConvertFrom-Json
$manifest = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json

if ($manifest.validation.source_applied -ne $true -or
    $manifest.validation.vbe_compile -ne "passed" -or
    $manifest.validation.workbook_saved -ne $true -or
    $manifest.validation.workbook_closed -ne $true -or
    $manifest.validation.excel_cleanup -ne "clean") {
    throw "Release manifest does not prove a clean compiled build."
}
if ($manifest.publication.method -ne "atomic_replace") {
    throw "Release publication was not atomic."
}

$expectedIncluded = @($policy.included | Sort-Object)
$expectedExcluded = @($policy.excluded | Sort-Object)
$manifestIncluded = @($manifest.included_components.name | Sort-Object)
$manifestExcluded = @($manifest.excluded_components.name | Sort-Object)

if ($null -ne (Compare-Object $expectedIncluded $manifestIncluded)) {
    throw "Release manifest included components differ from policy."
}
if ($null -ne (Compare-Object $expectedExcluded $manifestExcluded)) {
    throw "Release manifest excluded components differ from policy."
}

$excel = $null
$workbooks = $null
$workbook = $null
$components = $null
$component = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 3
    $workbooks = $excel.Workbooks
    $workbook = $workbooks.Open($resolvedArtifact, 0, $true)
    $components = $workbook.VBProject.VBComponents

    $actualComponents = @()
    for ($index = 1; $index -le $components.Count; $index++) {
        $component = $components.Item($index)
        $actualComponents += $component.Name
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($component)
        $component = $null
    }

    $actualComponents = @($actualComponents | Sort-Object)
    if ($null -ne (Compare-Object $expectedIncluded $actualComponents)) {
        throw "Release workbook components differ from policy."
    }
}
finally {
    if ($null -ne $component) {
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($component)
    }
    if ($null -ne $components) {
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($components)
    }
    if ($null -ne $workbook) {
        try { $workbook.Close($false) } catch { Write-Warning "Could not close release workbook: $_" }
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($workbook)
    }
    if ($null -ne $workbooks) {
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($workbooks)
    }
    if ($null -ne $excel) {
        try { $excel.Quit() } catch { Write-Warning "Could not quit Excel inspection process: $_" }
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
    }
}

# The production transport contains an intentional, controlled DoEvents checkpoint.
# xlflow's conservative headless preflight rejects any workbook containing DoEvents,
# even though Main.Run does not execute that path. Excel still runs invisibly here.
$runJson = & xlflow run Main.Run --input $resolvedArtifact --no-save --direct --json
if ($LASTEXITCODE -ne 0) {
    throw "Release consumer smoke failed with exit code $LASTEXITCODE."
}
$runResult = $runJson | Out-String | ConvertFrom-Json
if ($runResult.status -ne "ok" -or $runResult.macro.name -ne "Main.Run") {
    throw "Release consumer smoke returned an unexpected result."
}

$smokeDirectory = [IO.Path]::GetFullPath((Join-Path $projectRoot ".xlflow\release-consumer-smoke"))
$serverExecutable = Join-Path $smokeDirectory "testserver.exe"
$serverStdout = Join-Path $smokeDirectory "stdout.jsonl"
$serverStderr = Join-Path $smokeDirectory "stderr.log"
$serverProcess = $null
$consumerExcel = $null
$consumerWorkbooks = $null
$consumerWorkbook = $null
$harnessWorkbook = $null
$harnessComponent = $null
$client = $null
$response = $null
$ready = $null
try {
    [void](New-Item -ItemType Directory -Path $smokeDirectory -Force)
    Push-Location (Join-Path $projectRoot "tools\testserver")
    try {
        & go build -o $serverExecutable .
        if ($LASTEXITCODE -ne 0) { throw "Release smoke test-server build failed with exit code $LASTEXITCODE." }
    }
    finally {
        Pop-Location
    }

    $serverProcess = Start-Process `
        -FilePath $serverExecutable `
        -ArgumentList "-listen", "127.0.0.1:0" `
        -RedirectStandardOutput $serverStdout `
        -RedirectStandardError $serverStderr `
        -WindowStyle Hidden `
        -PassThru
    $deadline = [DateTime]::UtcNow.AddSeconds(10)
    while ([DateTime]::UtcNow -lt $deadline) {
        if ($serverProcess.HasExited) {
            $stderr = if (Test-Path -LiteralPath $serverStderr) { Get-Content -LiteralPath $serverStderr -Raw } else { "" }
            throw "Release smoke test server exited before readiness: $stderr"
        }
        if (Test-Path -LiteralPath $serverStdout) {
            $line = Get-Content -LiteralPath $serverStdout -TotalCount 1
            if ($line) {
                $ready = $line | ConvertFrom-Json
                break
            }
        }
        Start-Sleep -Milliseconds 50
    }
    if ($null -eq $ready -or $ready.event -ne "ready" -or -not $ready.url) {
        throw "Release smoke test server did not publish readiness."
    }

    $consumerExcel = New-Object -ComObject Excel.Application
    $consumerExcel.Visible = $false
    $consumerExcel.DisplayAlerts = $false
    $consumerExcel.EnableEvents = $false
    $consumerExcel.AutomationSecurity = 1
    $consumerWorkbooks = $consumerExcel.Workbooks
    $consumerWorkbook = $consumerWorkbooks.Open($resolvedArtifact, 0, $true)
    $factoryMacro = "'$($consumerWorkbook.Name)'!VBAHttp.CreateClient"
    $client = $consumerExcel.Run($factoryMacro)
    if ($null -eq $client) { throw "Release factory returned no HttpClient." }
    $client.BaseUrl = [string]$ready.url
    $response = $client.GetResponse("/status/204")
    if ($null -eq $response -or $response.StatusCode -ne 204 -or -not $response.IsSuccess) {
        throw "Release HttpClient GET returned an unexpected response."
    }

    $harnessWorkbook = $consumerWorkbooks.Add()
    $harnessComponent = $harnessWorkbook.VBProject.VBComponents.Import((Join-Path $PSScriptRoot "consumer\ReleaseBatchSmoke.bas"))
    if ($harnessComponent.Name -ne "ReleaseBatchSmoke") { throw "External batch smoke module import failed." }
    $batchMacro = "'$($harnessWorkbook.Name)'!ReleaseBatchSmoke.RunBatchSmoke"
    [void]$consumerExcel.Run($batchMacro, $consumerWorkbook.Name, [string]$ready.url)
    $reliabilityMacro = "'$($harnessWorkbook.Name)'!ReleaseBatchSmoke.RunReliabilitySmoke"
    [void]$consumerExcel.Run($reliabilityMacro, $consumerWorkbook.Name, [string]$ready.url)
}
finally {
    if ($null -ne $response -and [Runtime.InteropServices.Marshal]::IsComObject($response)) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($response)
    }
    if ($null -ne $client -and [Runtime.InteropServices.Marshal]::IsComObject($client)) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($client)
    }
    if ($null -ne $harnessComponent) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($harnessComponent) }
    if ($null -ne $harnessWorkbook) {
        try { $harnessWorkbook.Close($false) } catch { Write-Warning "Could not close external batch smoke workbook: $_" }
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($harnessWorkbook)
    }
    if ($null -ne $consumerWorkbook) {
        try { $consumerWorkbook.Close($false) } catch { Write-Warning "Could not close consumer smoke workbook: $_" }
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($consumerWorkbook)
    }
    if ($null -ne $consumerWorkbooks) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($consumerWorkbooks) }
    if ($null -ne $consumerExcel) {
        try { $consumerExcel.Quit() } catch { Write-Warning "Could not quit consumer smoke Excel: $_" }
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($consumerExcel)
    }
    if ($null -ne $ready -and $ready.url) {
        try { Invoke-WebRequest -Method Post -Uri "$($ready.url)/__admin/shutdown" -UseBasicParsing | Out-Null } catch { Write-Warning "Could not request smoke server shutdown: $_" }
    }
    if ($null -ne $serverProcess -and -not $serverProcess.HasExited) {
        if (-not $serverProcess.WaitForExit(5000)) { Stop-Process -Id $serverProcess.Id -Force }
    }
    if ($null -ne $serverProcess) { $serverProcess.Dispose() }
    $resolvedXlflowRoot = [IO.Path]::GetFullPath((Join-Path $projectRoot ".xlflow"))
    if ($smokeDirectory.StartsWith($resolvedXlflowRoot + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase) -and
        (Test-Path -LiteralPath $smokeDirectory)) {
        Remove-Item -LiteralPath $smokeDirectory -Recurse -Force
    }
}

Write-Output "Release artifact is valid: $($actualComponents.Count) components; external GET, batch, retry, and deadline smoke passed."
