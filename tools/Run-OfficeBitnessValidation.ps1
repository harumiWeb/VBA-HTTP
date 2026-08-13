[CmdletBinding()]
param(
    [ValidateSet("Auto", "X86", "X64")]
    [string]$ExpectedArchitecture = "Auto",
    [string]$OutputPath = ""
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$projectRoot = [IO.Path]::GetFullPath((Split-Path -Parent $PSScriptRoot))
$validationRoot = [IO.Path]::GetFullPath((Join-Path $projectRoot ".xlflow\bitness-validation"))
$buildRoot = Join-Path $validationRoot "build"
$artifactPath = Join-Path $buildRoot "VBA-HTTP.xlsm"

function Invoke-JsonCommand([string]$FilePath, [string[]]$Arguments) {
    $jsonLines = & $FilePath @Arguments
    if ($LASTEXITCODE -ne 0) { throw "$FilePath failed with exit code $LASTEXITCODE." }
    return (($jsonLines | Out-String) | ConvertFrom-Json)
}

function Publish-Json([string]$Path, $Document) {
    $directory = [IO.Path]::GetDirectoryName($Path)
    if (-not (Test-Path -LiteralPath $directory -PathType Container)) {
        [void](New-Item -ItemType Directory -Path $directory -Force)
    }
    $staging = Join-Path $directory ("." + [IO.Path]::GetFileName($Path) + "." + [guid]::NewGuid().ToString("N") + ".tmp")
    try {
        [IO.File]::WriteAllText($staging, (($Document | ConvertTo-Json -Depth 10) + [Environment]::NewLine), [Text.UTF8Encoding]::new($false))
        Move-Item -LiteralPath $staging -Destination $Path -Force
    }
    finally {
        if (Test-Path -LiteralPath $staging -PathType Leaf) { Remove-Item -LiteralPath $staging -Force }
    }
}

function Get-ExcelProcessIds {
    @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | ForEach-Object { [int]$_.Id })
}

function Get-OfficeInfo {
    $excel = $null
    $baselineIds = @(Get-ExcelProcessIds)
    $ownedIds = @()
    try {
        $excel = New-Object -ComObject Excel.Application
        $deadline = [DateTime]::UtcNow.AddSeconds(5)
        do {
            $afterIds = @(Get-ExcelProcessIds)
            $ownedIds = @($afterIds | Where-Object { $baselineIds -notcontains $_ })
            if ($ownedIds.Count -gt 0) { break }
            Start-Sleep -Milliseconds 50
        } while ([DateTime]::UtcNow -lt $deadline)
        if ($ownedIds.Count -eq 0) {
            throw "Could not prove ownership of a new Excel process; refusing to quit a pre-existing Excel instance."
        }
        return [ordered]@{
            version = [string]$excel.Version
            build = [string]$excel.Build
            operating_system = [string]$excel.OperatingSystem
        }
    }
    finally {
        if ($null -ne $excel) {
            if ($ownedIds.Count -gt 0) {
                try { $excel.Quit() } catch {}
            }
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
        }
    }
}

try {
    if (Test-Path -LiteralPath $validationRoot) { Remove-Item -LiteralPath $validationRoot -Recurse -Force }
    [void](New-Item -ItemType Directory -Path $buildRoot -Force)

    Push-Location $projectRoot
    try {
        $status = Invoke-JsonCommand "xlflow" @("status", "--json")
        if ($status.status -ne "ok" -or $status.state.src_newer_than_workbook -eq $true -or $status.coordination.recovery_required -eq $true) {
            throw "Source/workbook state is not a synchronized, recovery-free baseline."
        }

        $doctor = Invoke-JsonCommand "xlflow" @("doctor", "--json")
        $architecture = [string]$doctor.bridge.architecture
        if ($architecture -notin @("X86", "X64")) { throw "xlflow did not report X86 or X64 bridge architecture." }
        if ($ExpectedArchitecture -ne "Auto" -and $architecture -ne $ExpectedArchitecture) {
            throw "Expected $ExpectedArchitecture Office bridge, but xlflow selected $architecture."
        }

        $tests = Invoke-JsonCommand "xlflow" @("test", "--isolation", "module", "--json")
        if ($tests.status -ne "ok") { throw "Office $architecture test suite did not return status ok." }
        $testRecords = @($tests.tests)
        $passed = @($testRecords | Where-Object { $_.status -eq "passed" }).Count
        if ($passed -le 0) { throw "Office $architecture test suite produced no passing tests." }

        $integrationScript = Join-Path $projectRoot "tools\Run-IntegrationTests.ps1"
        $integration = Invoke-JsonCommand "powershell" @("-NoProfile", "-ExecutionPolicy", "Bypass", "-File", $integrationScript)
        $integrationRecords = @($integration.tests)
        if ($integration.status -ne "ok" -or $integrationRecords.Count -eq 0 -or
            @($integrationRecords | Where-Object { $_.status -ne "passed" }).Count -ne 0) {
            throw "Office $architecture integration suite did not return only passing tests."
        }

        $build = Invoke-JsonCommand "xlflow" @("build", "--json", "--base", (Join-Path $projectRoot "build\VBA-HTTP.xlsm"), "--out", $artifactPath)
        $validation = $build.build.validation
        if ($build.status -ne "ok" -or $build.bridge.architecture -ne $architecture -or
            $validation.source_applied -ne $true -or $validation.vbe_compile -ne "passed" -or
            $validation.workbook_saved -ne $true -or $validation.workbook_closed -ne $true -or
            $validation.excel_cleanup -ne "clean") {
            throw "Office $architecture build did not provide complete VBE evidence."
        }

        $smoke = Invoke-JsonCommand "xlflow" @("run", "Main.Run", "--input", $artifactPath, "--no-save", "--direct", "--json")
        if ($smoke.status -ne "ok" -or $smoke.macro.name -ne "Main.Run") {
            throw "Office $architecture consumer smoke did not pass."
        }

        $relativeOutput = if ([string]::IsNullOrWhiteSpace($OutputPath)) {
            "benchmarks/results/office-bitness-$($architecture.ToLowerInvariant()).json"
        }
        else { $OutputPath }
        $resolvedOutput = if ([IO.Path]::IsPathRooted($relativeOutput)) { [IO.Path]::GetFullPath($relativeOutput) } else { [IO.Path]::GetFullPath((Join-Path $projectRoot $relativeOutput)) }
        $evidence = [ordered]@{
            schema_version = 1
            benchmark = "office-bitness-validation"
            architecture = $architecture
            bridge = [ordered]@{ name = [string]$doctor.bridge.name; version = [string]$doctor.bridge.version; runtime = [string]$doctor.bridge.runtime }
            office = Get-OfficeInfo
            tests = [ordered]@{ total = $testRecords.Count; passed = $passed; inconclusive = @($testRecords | Where-Object { $_.status -eq "inconclusive" }).Count; failed = @($testRecords | Where-Object { $_.status -eq "failed" }).Count }
            integration = [ordered]@{ total = $integrationRecords.Count; passed = @($integrationRecords | Where-Object { $_.status -eq "passed" }).Count; failed = @($integrationRecords | Where-Object { $_.status -ne "passed" }).Count }
            build = [ordered]@{ vbe_compile = [string]$validation.vbe_compile; source_applied = [bool]$validation.source_applied; workbook_saved = [bool]$validation.workbook_saved; workbook_closed = [bool]$validation.workbook_closed; excel_cleanup = [string]$validation.excel_cleanup }
            consumer_smoke = "passed"
            external_network = $false
            status = "passed"
        }
        Publish-Json $resolvedOutput $evidence
        & (Join-Path $PSScriptRoot "Validate-OfficeBitnessResult.ps1") -Path $resolvedOutput
        Write-Output "Office bitness validation passed: $architecture ($passed passing unit tests, $($integrationRecords.Count) passing integration tests, VBE compile, and consumer smoke). Evidence: $resolvedOutput"
    }
    finally {
        Pop-Location
    }
}
finally {
    $resolvedRoot = [IO.Path]::GetFullPath($validationRoot)
    if ($resolvedRoot.StartsWith([IO.Path]::GetFullPath($projectRoot) + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase) -and
        (Test-Path -LiteralPath $resolvedRoot)) {
        Remove-Item -LiteralPath $resolvedRoot -Recurse -Force
    }
}
