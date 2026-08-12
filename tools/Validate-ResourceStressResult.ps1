[CmdletBinding()]
param(
    [string]$Path = "benchmarks/results/phase9-resource-stress.json"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$projectRoot = Split-Path -Parent $PSScriptRoot
$resultsRoot = [IO.Path]::GetFullPath((Join-Path $projectRoot "benchmarks\results"))
$resolvedPath = [IO.Path]::GetFullPath((Join-Path $projectRoot $Path))
if (-not $resolvedPath.StartsWith($resultsRoot + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) {
    throw "Resource result path must remain under benchmarks/results."
}
if (-not (Test-Path -LiteralPath $resolvedPath -PathType Leaf)) { throw "Resource result does not exist: $resolvedPath" }

$result = Get-Content -LiteralPath $resolvedPath -Raw | ConvertFrom-Json
if ($result.schema_version -ne 1 -or $result.benchmark -ne "phase9-resource-stress") { throw "Resource result schema identity is invalid." }
if ($result.source_commit -notmatch "^[0-9a-f]{40}$") { throw "Resource result source_commit is invalid." }
if ($result.server.external_network -ne $false -or [string]$result.server.base_url -notmatch "^http://(127\.0\.0\.1|localhost)(:[0-9]+)?$") {
    throw "Resource result server scope is invalid."
}
if ([int]$result.parameters.iterations -lt 1 -or [int]$result.parameters.scheduled_concurrency -lt 1) {
    throw "Resource result parameters are invalid."
}
$scenarios = @($result.results)
if ($scenarios.Count -lt 1) { throw "Resource result contains no scenarios." }
foreach ($scenario in $scenarios) {
    if ($scenario.scenario -notin @("sequential_native", "scheduled_com")) { throw "Unknown resource scenario: $($scenario.scenario)" }
    if ($scenario.status -ne 204 -or [int]$scenario.iterations -ne [int]$result.parameters.iterations) { throw "Resource scenario workload/status is invalid." }
    if ($scenario.process_before.observed -ne $true -or $scenario.process_after.observed -ne $true -or $scenario.process_idle_after.observed -ne $true) { throw "Resource scenario is missing PID-scoped process observations." }
    if ([int]$scenario.process_before.process_count -lt 1 -or [int]$scenario.process_after.process_count -lt 1 -or [int]$scenario.process_idle_after.process_count -lt 1) { throw "Resource scenario process counts are invalid." }
    if ([long]$scenario.idle_handle_delta -gt [long]$scenario.handle_delta_limit) { throw "Resource scenario exceeded its idle handle limit: $($scenario.scenario)." }
}
if ([int]$result.parameters.iterations -ge 10000 -and $scenarios.Count -ne 2) { throw "The 10,000-request gate requires sequential and scheduled scenarios." }
Write-Output "Resource stress result is valid: $resolvedPath"
