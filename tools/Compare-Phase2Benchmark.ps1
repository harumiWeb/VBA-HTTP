[CmdletBinding()]
param(
    [string]$RawPath = "benchmarks/results/phase2-raw-winhttp.json",
    [string]$CandidatePath = "benchmarks/results/vba-http-buffered.json",
    [string]$OutputPath = "benchmarks/results/phase2-buffered-overhead.json"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$projectRoot = Split-Path -Parent $PSScriptRoot
$resultsRoot = [IO.Path]::GetFullPath((Join-Path $projectRoot "benchmarks\results"))

function Resolve-ResultPath([string]$Path) {
    $resolved = [IO.Path]::GetFullPath((Join-Path $projectRoot $Path))
    if (-not $resolved.StartsWith($resultsRoot + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Benchmark paths must be inside $resultsRoot."
    }
    return $resolved
}

function Get-SequentialResult($Document, [string]$ExpectedImplementation) {
    if ($Document.schema_version -ne 1 -or $Document.benchmark -ne "http-client" -or
        $Document.implementation.name -ne $ExpectedImplementation -or
        $Document.server.external_network -ne $false) {
        throw "Benchmark document for $ExpectedImplementation does not satisfy the comparison contract."
    }
    $result = @($Document.results | Where-Object scenario -eq "sequential_get")
    if ($result.Count -ne 1 -or $result[0].status -ne 204 -or $result[0].iterations -lt 1 -or $result[0].mean_ms -le 0) {
        throw "Benchmark document for $ExpectedImplementation lacks a valid sequential_get result."
    }
    return $result[0]
}

$resolvedRaw = Resolve-ResultPath $RawPath
$resolvedCandidate = Resolve-ResultPath $CandidatePath
$resolvedOutput = Resolve-ResultPath $OutputPath
if (-not (Test-Path -LiteralPath $resolvedRaw -PathType Leaf) -or
    -not (Test-Path -LiteralPath $resolvedCandidate -PathType Leaf)) {
    throw "Both Phase 2 benchmark inputs must exist."
}

$raw = Get-Content -LiteralPath $resolvedRaw -Raw | ConvertFrom-Json
$candidate = Get-Content -LiteralPath $resolvedCandidate -Raw | ConvertFrom-Json
$rawSequential = Get-SequentialResult $raw "Raw WinHttpRequest"
$candidateSequential = Get-SequentialResult $candidate "VBA-HTTP"

$rawParameters = $raw.parameters | ConvertTo-Json -Depth 10 -Compress
$candidateParameters = $candidate.parameters | ConvertTo-Json -Depth 10 -Compress
if ($rawParameters -ne $candidateParameters -or
    $raw.environment.office_bitness -ne $candidate.environment.office_bitness -or
    $raw.environment.excel_version -ne $candidate.environment.excel_version -or
    $raw.environment.platform -ne $candidate.environment.platform) {
    throw "Raw and VBA-HTTP benchmark conditions differ."
}

$ratio = [double]$candidateSequential.mean_ms / [double]$rawSequential.mean_ms
$overheadPercent = ($ratio - 1.0) * 100.0
$targetPercent = 15.0
$comparison = [ordered]@{
    schema_version = 1
    benchmark = "phase2-buffered-com-overhead"
    baseline = [ordered]@{
        implementation = "Raw WinHttpRequest"
        result_file = [IO.Path]::GetFileName($resolvedRaw)
        mean_ms = [Math]::Round([double]$rawSequential.mean_ms, 3)
    }
    candidate = [ordered]@{
        implementation = "VBA-HTTP"
        result_file = [IO.Path]::GetFileName($resolvedCandidate)
        mean_ms = [Math]::Round([double]$candidateSequential.mean_ms, 3)
    }
    metric = "sequential_get_mean_ms"
    ratio = [Math]::Round($ratio, 3)
    overhead_percent = [Math]::Round($overheadPercent, 3)
    target_max_percent = $targetPercent
    within_target = ($overheadPercent -le $targetPercent)
    conditions = [ordered]@{
        office_bitness = $raw.environment.office_bitness
        excel_version = $raw.environment.excel_version
        platform = $raw.environment.platform
        warmup_iterations = $raw.parameters.warmup_iterations
        latency_iterations = $raw.parameters.latency_iterations
        endpoint = $rawSequential.path
        external_network = $false
    }
}

$temporaryPath = "$resolvedOutput.$([guid]::NewGuid().ToString('N')).tmp"
$backupPath = "$resolvedOutput.$([guid]::NewGuid().ToString('N')).bak"
try {
    [IO.File]::WriteAllText($temporaryPath, ($comparison | ConvertTo-Json -Depth 10) + [Environment]::NewLine, [Text.UTF8Encoding]::new($false))
    if (Test-Path -LiteralPath $resolvedOutput -PathType Leaf) {
        [IO.File]::Replace($temporaryPath, $resolvedOutput, $backupPath)
        Remove-Item -LiteralPath $backupPath -Force
    }
    else {
        Move-Item -LiteralPath $temporaryPath -Destination $resolvedOutput
    }
}
finally {
    if (Test-Path -LiteralPath $temporaryPath -PathType Leaf) { Remove-Item -LiteralPath $temporaryPath -Force }
    if (Test-Path -LiteralPath $backupPath -PathType Leaf) { Remove-Item -LiteralPath $backupPath -Force }
}

Write-Output ("Phase 2 overhead: {0:N3}% (target <= {1:N1}%); within target: {2}" -f $overheadPercent, $targetPercent, $comparison.within_target)
