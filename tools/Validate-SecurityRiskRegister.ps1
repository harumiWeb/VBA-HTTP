[CmdletBinding()]
param(
    [string]$Path = "docs/security/risk-register.json"
)

$ErrorActionPreference = "Stop"
$projectRoot = Split-Path -Parent $PSScriptRoot
$resolvedPath = if ([IO.Path]::IsPathRooted($Path)) {
    [IO.Path]::GetFullPath($Path)
}
else {
    [IO.Path]::GetFullPath((Join-Path $projectRoot $Path))
}

if (-not (Test-Path -LiteralPath $resolvedPath -PathType Leaf)) {
    throw "Security risk register does not exist: $Path"
}

$register = Get-Content -LiteralPath $resolvedPath -Raw | ConvertFrom-Json
if ([int]$register.schema_version -ne 1) {
    throw "Security risk register schema_version must be 1."
}
if ([string]::IsNullOrWhiteSpace([string]$register.review_id)) {
    throw "Security risk register review_id is required."
}

$allowedStatuses = @('open', 'known', 'deferred', 'mitigated')
$allowedSeverities = @('low', 'medium', 'high', 'critical')
$issues = @($register.issues)
$blockers = @($register.current_release_blockers)
$ids = @{}

foreach ($issue in $issues) {
    $id = [string]$issue.id
    if ([string]::IsNullOrWhiteSpace($id)) { throw "Security risk issue id is required." }
    if ($ids.ContainsKey($id)) { throw "Duplicate security risk issue id: $id" }
    $ids[$id] = $true
    if ($allowedStatuses -notcontains [string]$issue.status) {
        throw "Unsupported security risk status for $($id): $($issue.status)"
    }
    if ($allowedSeverities -notcontains [string]$issue.severity) {
        throw "Unsupported security risk severity for $($id): $($issue.severity)"
    }
    if ([string]::IsNullOrWhiteSpace([string]$issue.required_evidence) -or
        [string]::IsNullOrWhiteSpace([string]$issue.evidence)) {
        throw "Security risk evidence fields are required for $id."
    }
    if ([bool]$issue.blocks_current_release -and [string]$issue.status -eq 'mitigated') {
        throw "Mitigated issue cannot block the current release: $id"
    }
}

foreach ($blocker in $blockers) {
    $blockerId = [string]$blocker
    if (-not $ids.ContainsKey($blockerId)) {
        throw "Current release blocker is not present in issues: $blockerId"
    }
    $issue = $issues | Where-Object { [string]$_.id -eq $blockerId }
    if (-not [bool]$issue.blocks_current_release) {
        throw "Current release blocker is not marked blocks_current_release: $blockerId"
    }
    if ([string]$issue.status -eq 'mitigated') {
        throw "Current release blocker is marked mitigated: $blockerId"
    }
}

$actualCurrentBlockers = @($issues | Where-Object { [bool]$_.blocks_current_release -and [string]$_.status -ne 'mitigated' } | ForEach-Object { [string]$_.id })
if ($actualCurrentBlockers.Count -ne $blockers.Count -or
    $null -ne (Compare-Object (@($actualCurrentBlockers | Sort-Object) ) (@($blockers | Sort-Object)))) {
    throw "current_release_blockers does not match active blocks_current_release issues."
}
if ($blockers.Count -gt 0) {
    throw "Security risk register contains current release blockers: $($blockers -join ', ')"
}

Write-Output "Security risk register valid: $($issues.Count) issues, 0 current release blockers."
