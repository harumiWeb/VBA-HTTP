[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$root = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$workflowPath = Join-Path $root '.github/workflows/ci.yml'
$lockPath = Join-Path $root 'tools/release-toolchain.json'
if (-not (Test-Path -LiteralPath $workflowPath -PathType Leaf)) {
    throw "CI workflow is missing: $workflowPath"
}
if (-not (Test-Path -LiteralPath $lockPath -PathType Leaf)) {
    throw "Release toolchain lock is missing: $lockPath"
}

$workflow = Get-Content -LiteralPath $workflowPath -Raw
$lock = Get-Content -LiteralPath $lockPath -Raw | ConvertFrom-Json

function Assert-ContainsToken([string]$Text, [string]$Token) {
    if ($Text.IndexOf($Token, [StringComparison]::OrdinalIgnoreCase) -lt 0) {
        throw "CI workflow contract is missing '$Token'."
    }
}

foreach ($token in @(
        'name: Excel-free CI',
        'pull_request:',
        'push:',
        'schedule:',
        'workflow_dispatch:',
        'branches:',
        '- main',
        'permissions:',
        'contents: read',
        'concurrency:',
        'cancel-in-progress: true',
        'runs-on: windows-2022',
        'timeout-minutes: 30',
        'actions/checkout@v4',
        'fetch-depth: 0',
        'actions/setup-go@v5',
        'cache-dependency-path: tools/testserver/go.mod',
        'tools/release-toolchain.json',
        '$lock.task.version',
        '$lock.psscriptanalyzer.version',
        'Install-ReleaseToolchain.ps1',
        "'check'",
        'test:docs',
        'test:security-risks',
        'testserver:unit',
        'test:release-checksum',
        'test:release-security',
        'build:plan',
        'build:plan:xlam',
        'test:xlam',
        'test:pack-release',
        'test:github-release',
        'task test:clean-checkout',
        "github.event_name == 'schedule'",
        "github.event_name == 'workflow_dispatch'"
    )) {
    Assert-ContainsToken $workflow $token
}

foreach ($token in @([string]$lock.go.version)) {
    Assert-ContainsToken $workflow $token
}

foreach ($forbidden in @(
        'contents: write',
        'gh release create',
        'xlflow build',
        'precommit:compile',
        'runs-on: self-hosted'
    )) {
    if ($workflow.IndexOf($forbidden, [StringComparison]::OrdinalIgnoreCase) -ge 0) {
        throw "Excel-free CI must not contain '$forbidden'."
    }
}

Write-Output 'CI workflow contract valid: PR/main/scheduled triggers, read-only hosted runner, pinned tools, Excel-free gates, and clean-checkout boundary.'
