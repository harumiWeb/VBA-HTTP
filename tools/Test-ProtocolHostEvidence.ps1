[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$root = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot ".."))
$testRoot = Join-Path $root ".xlflow\protocol-host-evidence-test"
$validator = Join-Path $PSScriptRoot "Validate-ProtocolHostEvidence.ps1"

function Write-Record([string]$Path, $Record) {
    $Record | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $Path -Encoding UTF8
}

function New-ValidRecord {
    return [ordered]@{
        schema_version = 1
        benchmark = "protocol-host-validation"
        status = "passed"
        run_utc = "2026-01-01T00:00:00Z"
        source_revision = "0123456789abcdef0123456789abcdef01234567"
        requested_protocol = "HTTP/2"
        observed_protocol = "HTTP/2"
        mode = "required"
        external_network = $true
        target = [ordered]@{ scheme = "https"; host = "h2.example.test"; port = 443 }
        bridge = [ordered]@{ name = "xlflow-excel-bridge"; version = "1.0.0"; runtime = ".NET 8"; architecture = "X64" }
        office = [ordered]@{ version = "16.0"; build = "12345"; operating_system = "Windows" }
        environment = [ordered]@{ windows = "Microsoft Windows"; winhttp = [ordered]@{ enabled_protocol_option = 133; used_protocol_option = 134; required_protocol_option = 145; observed_protocol = "HTTP/2" } }
        artifact = [ordered]@{ name = "VBA-HTTP.xlsm"; sha256 = ("a" * 64); manifest_sha256 = ("b" * 64) }
        build = [ordered]@{ vbe_compile = "passed"; source_applied = $true; workbook_saved = $true; workbook_closed = $true; excel_cleanup = "clean" }
    }
}

try {
    if (Test-Path -LiteralPath $testRoot) { Remove-Item -LiteralPath $testRoot -Recurse -Force }
    [void](New-Item -ItemType Directory -Path $testRoot -Force)
    $validPath = Join-Path $testRoot "valid.json"
    Write-Record $validPath (New-ValidRecord)
    & powershell -NoProfile -ExecutionPolicy Bypass -File $validator -Path $validPath -ExpectedProtocol "HTTP/2"
    if ($LASTEXITCODE -ne 0) { throw "Valid protocol host evidence was rejected." }

    $cases = @(
        @{ Name = "mismatch"; Change = { param($r) $r.observed_protocol = "HTTP/3" } },
        @{ Name = "http"; Change = { param($r) $r.target.scheme = "http" } },
        @{ Name = "secret"; Change = { param($r) $r.authorization = "Basic hidden" } },
        @{ Name = "bad-hash"; Change = { param($r) $r.artifact.sha256 = "bad" } },
        @{ Name = "x86"; Change = { param($r) $r.bridge.architecture = "X86" } },
        @{ Name = "unknown"; Change = { param($r) $r.unexpected = "not allowed" } }
    )
    foreach ($case in $cases) {
        $record = New-ValidRecord
        & $case.Change $record
        $casePath = Join-Path $testRoot ($case.Name + ".json")
        Write-Record $casePath $record
        $failed = $false
        try {
            & powershell -NoProfile -ExecutionPolicy Bypass -File $validator -Path $casePath -ExpectedProtocol "HTTP/2" 2>$null | Out-Null
            if ($LASTEXITCODE -ne 0) { $failed = $true }
        }
        catch { $failed = $true }
        if (-not $failed) { throw "Tamper case '$($case.Name)' was accepted." }
    }
    Write-Output "Protocol host evidence validator tests passed: 1 valid record and $($cases.Count) fail-closed tamper cases."
}
finally {
    if (Test-Path -LiteralPath $testRoot) { Remove-Item -LiteralPath $testRoot -Recurse -Force }
}
