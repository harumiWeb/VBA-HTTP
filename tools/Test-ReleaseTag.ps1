[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$scriptPath = Join-Path $PSScriptRoot 'Validate-ReleaseTag.ps1'
$testRoot = Join-Path ([IO.Path]::GetTempPath()) ('vba-http-release-tag-' + [guid]::NewGuid().ToString('N'))

function Invoke-Tag([string]$Tag, [string]$ExpectedCommit = '') {
    Push-Location $testRoot
    try {
        $arguments = @{ Tag = $Tag }
        if (-not [string]::IsNullOrWhiteSpace($ExpectedCommit)) { $arguments.ExpectedCommit = $ExpectedCommit }
        $output = & $scriptPath @arguments
        if ($LASTEXITCODE -ne 0) { throw "Tag validator failed unexpectedly for $Tag." }
        return (($output | Out-String) | ConvertFrom-Json)
    }
    finally {
        Pop-Location
    }
}

function Assert-TagFailure([string]$Tag, [string]$ExpectedCommit = '') {
    $failed = $false
    try {
        Invoke-Tag $Tag $ExpectedCommit | Out-Null
    }
    catch { $failed = $true }
    if (-not $failed) { throw "Tag validator accepted an invalid case: $Tag" }
}

try {
    [void](New-Item -ItemType Directory -Path $testRoot -Force)
    & git init --quiet $testRoot
    & git -C $testRoot -c user.name='VBA-HTTP test' -c user.email='test@example.invalid' commit --allow-empty --quiet -m 'tag test'
    $commit = (& git -C $testRoot rev-parse HEAD).Trim()
    & git -C $testRoot tag v1.2.3

    $stable = Invoke-Tag 'v1.2.3' $commit
    if ($stable.prerelease -ne $false -or $stable.commit -ne $commit) { throw 'Stable tag classification is invalid.' }

    $prerelease = Invoke-Tag 'v1.2.3-rc.1'
    if ($prerelease.prerelease -ne $true) { throw 'Prerelease tag classification is invalid.' }
    Assert-TagFailure 'v1.2.3+build'
    Assert-TagFailure 'v01.2.3'
    Assert-TagFailure 'v1.2'
    Assert-TagFailure 'v1.2.3' ('0' * 40)
    Write-Output 'Release tag tests passed: strict SemVer, stable/prerelease classification, commit match, and rejection cases.'
}
finally {
    if (Test-Path -LiteralPath $testRoot -PathType Container) { Remove-Item -LiteralPath $testRoot -Recurse -Force }
}
