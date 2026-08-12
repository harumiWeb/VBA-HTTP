[CmdletBinding()]
param(
    [string]$Path = "benchmarks/results/phase9-cancellation-stress.json"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$projectRoot = Split-Path -Parent $PSScriptRoot
$resultsRoot = [IO.Path]::GetFullPath((Join-Path $projectRoot "benchmarks\results"))
$resolvedPath = [IO.Path]::GetFullPath((Join-Path $projectRoot $Path))
if (-not $resolvedPath.StartsWith($resultsRoot + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) {
    throw "Cancellation result path must remain under benchmarks/results."
}
if (-not (Test-Path -LiteralPath $resolvedPath -PathType Leaf)) { throw "Cancellation result does not exist: $resolvedPath" }

$result = Get-Content -LiteralPath $resolvedPath -Raw | ConvertFrom-Json
if ($result.schema_version -ne 1 -or $result.benchmark -ne "phase9-cancellation-stress") { throw "Cancellation result schema identity is invalid." }
if ($result.source_commit -notmatch "^[0-9a-f]{40}$") { throw "Cancellation result source_commit is invalid." }
if ($result.server.external_network -ne $false -or [string]$result.server.base_url -notmatch "^http://(127\.0\.0\.1|localhost)(:[0-9]+)?$") {
    throw "Cancellation result server scope is invalid."
}
if ([int]$result.parameters.iterations -lt 1 -or [int]$result.parameters.iterations -gt 1000) { throw "Cancellation result iteration parameter is invalid." }
if ([int]$result.parameters.com_cancellation_requests -ne 4 -or [int]$result.parameters.com_deadline_requests -ne 4) { throw "Cancellation result COM workload is invalid." }
if ([int]$result.parameters.native_download_bytes -ne 4194304 -or [int]$result.parameters.native_download_cancel_after_bytes -ne 65536) { throw "Cancellation result download workload is invalid." }
if ([int]$result.parameters.idle_wait_ms -ne 1000) { throw "Cancellation result timeout/sampling parameters are invalid." }
if ([int]$result.parameters.handle_delta_limits.native -ne 8 -or [int]$result.parameters.handle_delta_limits.com -ne 32) { throw "Cancellation result handle limits are invalid." }

$scenarios = @($result.results)
if ($scenarios.Count -lt 1) { throw "Cancellation result contains no scenarios." }
$known = @("com_active_cancellation", "com_request_deadline", "native_download_cancellation")
foreach ($scenario in $scenarios) {
    if ($scenario.scenario -notin $known) { throw "Unknown cancellation scenario: $($scenario.scenario)" }
    if ($scenario.status -ne "passed" -or [int]$scenario.iterations -ne [int]$result.parameters.iterations) { throw "Cancellation scenario workload/status is invalid: $($scenario.scenario)" }
    if ([string]$scenario.invariant -notmatch "\S") { throw "Cancellation scenario invariant is missing: $($scenario.scenario)" }
    if ($scenario.process_before.observed -ne $true -or $scenario.process_after.observed -ne $true -or $scenario.process_idle_after.observed -ne $true) { throw "Cancellation scenario is missing PID-scoped process observations: $($scenario.scenario)" }
    if ([int]$scenario.process_before.process_count -lt 1 -or [int]$scenario.process_after.process_count -lt 1 -or [int]$scenario.process_idle_after.process_count -lt 1) { throw "Cancellation scenario process counts are invalid: $($scenario.scenario)" }
    if ([long]$scenario.idle_handle_delta -gt [long]$scenario.handle_delta_limit) { throw "Cancellation scenario exceeded its idle handle limit: $($scenario.scenario)" }
    if ($scenario.scenario -like "com_*") {
        if ($scenario.transport -ne "WinHttpComTransport" -or [int]$scenario.handle_delta_limit -ne 32) { throw "COM cancellation scenario transport/limit is invalid: $($scenario.scenario)" }
    }
    else {
        if ($scenario.transport -ne "WinHttpNativeTransport" -or [int]$scenario.handle_delta_limit -ne 8) { throw "Native cancellation scenario transport/limit is invalid: $($scenario.scenario)" }
    }
}
if ($scenarios.Count -eq 3 -and (@($scenarios | Select-Object -ExpandProperty scenario -Unique).Count -ne 3)) { throw "Complete cancellation result contains duplicate scenarios." }
Write-Output "Cancellation stress result is valid: $resolvedPath"
