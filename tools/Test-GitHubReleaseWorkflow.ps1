[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$root = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$workflowPath = Join-Path $root '.github/workflows/release.yml'
if (-not (Test-Path -LiteralPath $workflowPath -PathType Leaf)) { throw "Release workflow is missing: $workflowPath" }
$workflow = Get-Content -LiteralPath $workflowPath -Raw
$lock = Get-Content -LiteralPath (Join-Path $root 'tools/release-toolchain.json') -Raw | ConvertFrom-Json

function Assert-ContainsToken([string]$Text, [string]$Token) {
    if ($Text.IndexOf($Token, [StringComparison]::OrdinalIgnoreCase) -lt 0) {
        throw "Release workflow contract is missing '$Token'."
    }
}

foreach ($token in @(
        'tags:',
        '"v*"',
        'runs-on: windows-2022',
        'permissions:',
        'contents: write',
        'fetch-depth: 0',
        'Install-ReleaseToolchain.ps1',
        "@('check', 'test:docs', 'test:security-risks', 'testserver:unit', 'build:plan', 'test:pack-release', 'test:github-release')",
        'Invoke-RequiredTask',
        'gh release list --limit 1000 --json tagName,isDraft',
        'New-GitHubReleaseBundle.ps1',
        'Validate-GitHubReleaseBundle.ps1',
        'gh release create',
        '--verify-tag',
        '--notes-file',
        '--latest=false',
        'VBE validation not performed',
        'VBA-HTTP-$($env:RELEASE_TAG)-source.zip',
        'VBA-HTTP-$($env:RELEASE_TAG).xlsm',
        'VBA-HTTP-$($env:RELEASE_TAG).pack.json',
        'VBA-HTTP-$($env:RELEASE_TAG).release.json',
        'VBA-HTTP-$($env:RELEASE_TAG).SHA256SUMS.txt',
        'LICENSE',
        'THIRD_PARTY_NOTICES.md'
    )) {
    Assert-ContainsToken $workflow $token
}
foreach ($token in @(
        [string]$lock.go.version,
        [string]$lock.task.version,
        [string]$lock.psscriptanalyzer.version
    )) {
    Assert-ContainsToken $workflow $token
}
if ($workflow -match '(?im)^\s*uses:\s+actions/upload-artifact') { throw 'Release workflow must not upload an intermediate artifact.' }
if ($workflow -match '(?im)^\s*runs-on:\s+self-hosted') { throw 'Release workflow must use GitHub-hosted Windows runners only.' }
if ($workflow -match '(?im)--clobber') { throw 'Release workflow must not use --clobber.' }
if ($workflow -match '(?im)xlflow\s+build') { throw 'GitHub release workflow must not invoke the Excel-backed xlflow build.' }
if ($workflow -match '(?im)precommit:compile') { throw 'GitHub release workflow must not invoke the local VBE compile gate.' }
Write-Output 'GitHub release workflow contract valid: tag filter, x64 hosted runner, least privilege, gates, verify-tag publication, and asset set.'
