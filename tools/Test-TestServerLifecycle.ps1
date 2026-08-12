[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"
$projectRoot = Split-Path -Parent $PSScriptRoot
$artifactDirectory = [System.IO.Path]::GetFullPath((Join-Path $projectRoot ".xlflow\testserver-smoke"))
$expectedDirectory = [System.IO.Path]::GetFullPath((Join-Path $projectRoot ".xlflow\testserver-smoke"))
if ($artifactDirectory -ne $expectedDirectory) {
    throw "Unexpected test server artifact path."
}

$executablePath = Join-Path $artifactDirectory "testserver.exe"
$stdoutPath = Join-Path $artifactDirectory "stdout.jsonl"
$stderrPath = Join-Path $artifactDirectory "stderr.log"
$process = $null
$client = $null
$health = $null
$hashResponse = $null
$streamResponse = $null
$stream = $null
$sha256 = $null
$emptyContent = $null
$reset = $null
$shutdownContent = $null
$shutdown = $null

try {
    [void](New-Item -ItemType Directory -Path $artifactDirectory -Force)
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
    $health = $client.GetAsync("$($ready.url)/healthz").GetAwaiter().GetResult()
    if (-not $health.IsSuccessStatusCode) {
        throw "Health check failed with status $([int]$health.StatusCode)."
    }

    $payloadSize = 100 * 1024 * 1024
    $hashResponse = $client.GetAsync("$($ready.url)/sha256/$payloadSize").GetAwaiter().GetResult()
    if (-not $hashResponse.IsSuccessStatusCode) {
        throw "Payload hash request failed with status $([int]$hashResponse.StatusCode)."
    }
    $expectedHash = ($hashResponse.Content.ReadAsStringAsync().GetAwaiter().GetResult() | ConvertFrom-Json).digest
    $streamResponse = $client.GetAsync(
        "$($ready.url)/stream/$payloadSize",
        [System.Net.Http.HttpCompletionOption]::ResponseHeadersRead
    ).GetAwaiter().GetResult()
    if (-not $streamResponse.IsSuccessStatusCode) {
        throw "100 MiB stream failed with status $([int]$streamResponse.StatusCode)."
    }
    $stream = $streamResponse.Content.ReadAsStreamAsync().GetAwaiter().GetResult()
    $sha256 = [System.Security.Cryptography.SHA256]::Create()
    $actualHash = [System.BitConverter]::ToString($sha256.ComputeHash($stream)).Replace("-", "").ToLowerInvariant()
    if ($actualHash -ne $expectedHash) {
        throw "100 MiB stream hash mismatch."
    }

    $emptyContent = New-Object System.Net.Http.StringContent("")
    $reset = $client.PostAsync("$($ready.url)/__admin/reset", $emptyContent).GetAwaiter().GetResult()
    if ([int]$reset.StatusCode -ne 204) {
        throw "State reset failed with status $([int]$reset.StatusCode)."
    }

    $shutdownContent = New-Object System.Net.Http.StringContent("")
    $shutdown = $client.PostAsync("$($ready.url)/__admin/shutdown", $shutdownContent).GetAwaiter().GetResult()
    if ([int]$shutdown.StatusCode -ne 202) {
        throw "Graceful shutdown request failed with status $([int]$shutdown.StatusCode)."
    }
    if (-not $process.WaitForExit(5000)) {
        throw "Test server did not exit after graceful shutdown."
    }

    Write-Output "Test server lifecycle is valid: $($ready.url)"
}
finally {
    foreach ($disposable in @($shutdown, $shutdownContent, $reset, $emptyContent, $sha256, $stream, $streamResponse, $hashResponse, $health)) {
        if ($null -ne $disposable) {
            $disposable.Dispose()
        }
    }
    if ($null -ne $client) {
        $client.Dispose()
    }
    if ($null -ne $process -and -not $process.HasExited) {
        Stop-Process -Id $process.Id -Force
        [void]$process.WaitForExit(5000)
    }
    if (Test-Path -LiteralPath $artifactDirectory) {
        Remove-Item -LiteralPath $artifactDirectory -Recurse -Force
    }
}
