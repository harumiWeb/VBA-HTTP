[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$projectRoot = Split-Path -Parent $PSScriptRoot
$artifactDirectory = [IO.Path]::GetFullPath((Join-Path $projectRoot ".xlflow\integration-tests"))
$executablePath = Join-Path $artifactDirectory "testserver.exe"
$stdoutPath = Join-Path $artifactDirectory "stdout.jsonl"
$stderrPath = Join-Path $artifactDirectory "stderr.log"
$process = $null
$client = $null
$ready = $null
$proxyUrl = $null
$previousBaseUrl = [Environment]::GetEnvironmentVariable("VBA_HTTP_TEST_BASE_URL", "Process")
$previousProxyUrl = [Environment]::GetEnvironmentVariable("VBA_HTTP_TEST_PROXY_URL", "Process")
$previousProxyTargetUrl = [Environment]::GetEnvironmentVariable("VBA_HTTP_TEST_PROXY_TARGET_URL", "Process")

try {
    [void](New-Item -ItemType Directory -Path $artifactDirectory -Force)
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
        -ArgumentList "-listen", "127.0.0.1:0", "-proxy-listen", "127.0.0.1:0" `
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
    if ($null -eq $ready -or $ready.event -ne "ready" -or -not $ready.url -or -not $ready.proxy_url -or -not $ready.proxy_target_url) {
        throw "Test server did not publish valid HTTP and proxy readiness URLs."
    }

    [Environment]::SetEnvironmentVariable("VBA_HTTP_TEST_BASE_URL", [string]$ready.url, "Process")
    [Environment]::SetEnvironmentVariable("VBA_HTTP_TEST_PROXY_URL", [string]$ready.proxy_url, "Process")
    [Environment]::SetEnvironmentVariable("VBA_HTTP_TEST_PROXY_TARGET_URL", [string]$ready.proxy_target_url, "Process")
    $testJson = & xlflow test --tag integration --json
    $testExitCode = $LASTEXITCODE
    $testJson | Write-Output
    if ($testExitCode -ne 0) {
        throw "xlflow integration tests failed with exit code $testExitCode."
    }
    $result = $testJson | Out-String | ConvertFrom-Json
    $tests = @($result.tests)
    if ($result.status -ne "ok" -or $tests.Count -ne 59 -or @($tests | Where-Object status -ne "passed").Count -ne 0) {
        throw "Integration suite did not return fifty-nine passing tests."
    }
}
finally {
    [Environment]::SetEnvironmentVariable("VBA_HTTP_TEST_BASE_URL", $previousBaseUrl, "Process")
    [Environment]::SetEnvironmentVariable("VBA_HTTP_TEST_PROXY_URL", $previousProxyUrl, "Process")
    [Environment]::SetEnvironmentVariable("VBA_HTTP_TEST_PROXY_TARGET_URL", $previousProxyTargetUrl, "Process")
    if ($null -ne $ready -and $ready.url) {
        try {
            Add-Type -AssemblyName System.Net.Http
            $client = New-Object System.Net.Http.HttpClient
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
    if ($null -ne $process -and -not $process.HasExited) {
        if (-not $process.WaitForExit(5000)) {
            Stop-Process -Id $process.Id -Force
            [void]$process.WaitForExit(5000)
        }
    }
    if ($null -ne $process) { $process.Dispose() }

    $resolvedRoot = [IO.Path]::GetFullPath((Join-Path $projectRoot ".xlflow"))
    if ($artifactDirectory.StartsWith($resolvedRoot + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase) -and
        (Test-Path -LiteralPath $artifactDirectory)) {
        Remove-Item -LiteralPath $artifactDirectory -Recurse -Force
    }
}
