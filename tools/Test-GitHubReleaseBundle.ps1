[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$root = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$testRoot = Join-Path $root (Join-Path '.release' ('test-github-release-' + [guid]::NewGuid().ToString('N')))
$bundle = Join-Path $testRoot 'bundle'
$tag = 'v0.0.0-test'
$revision = '0123456789abcdef0123456789abcdef01234567'

function Invoke-Script([string]$ScriptName, [hashtable]$Arguments) {
    & (Join-Path $PSScriptRoot $ScriptName) @Arguments | Out-Null
}

try {
    [void](New-Item -ItemType Directory -Path $testRoot -Force)
    Invoke-Script 'New-GitHubReleaseBundle.ps1' @{
        Tag = $tag
        SourceRevision = $revision
        OutputDirectory = $bundle
        RootPath = $root
    }
    Invoke-Script 'Validate-GitHubReleaseBundle.ps1' @{
        Tag = $tag
        BundleDirectory = $bundle
        RootPath = $root
    }

    $expected = @(
        "VBA-HTTP-$tag-source.zip",
        "VBA-HTTP-$tag.xlsm",
        "VBA-HTTP-$tag.pack.json",
        "VBA-HTTP-$tag.release.json",
        "VBA-HTTP-$tag.SHA256SUMS.txt",
        'LICENSE',
        'THIRD_PARTY_NOTICES.md'
    ) | Sort-Object
    $actual = @(Get-ChildItem -LiteralPath $bundle -File | Select-Object -ExpandProperty Name | Sort-Object)
    if ($null -ne (Compare-Object $expected $actual)) { throw 'GitHub release bundle test produced the wrong asset set.' }

    $release = Get-Content -LiteralPath (Join-Path $bundle "VBA-HTTP-$tag.release.json") -Raw | ConvertFrom-Json
    if ($release.pack.vbe_validation -ne 'not_performed' -or $release.support.office_bitness -ne 'x64' -or $release.support.http3_quic -ne 'unsupported-by-policy') {
        throw 'GitHub release bundle test produced invalid support/provenance metadata.'
    }

    $tampered = Join-Path $bundle "VBA-HTTP-$tag.xlsm"
    $original = [IO.File]::ReadAllBytes($tampered)
    try {
        [IO.File]::AppendAllText($tampered, 'tamper')
        $failed = $false
        try {
            Invoke-Script 'Validate-GitHubReleaseBundle.ps1' @{ Tag = $tag; BundleDirectory = $bundle; RootPath = $root }
        }
        catch { $failed = $true }
        if (-not $failed) { throw 'GitHub release bundle validator accepted a tampered asset.' }
    }
    finally {
        [IO.File]::WriteAllBytes($tampered, $original)
    }

    $releasePath = Join-Path $bundle "VBA-HTTP-$tag.release.json"
    $releaseOriginal = [IO.File]::ReadAllBytes($releasePath)
    try {
        [IO.File]::AppendAllText($releasePath, 'tamper')
        $failed = $false
        try {
            Invoke-Script 'Validate-GitHubReleaseBundle.ps1' @{ Tag = $tag; BundleDirectory = $bundle; RootPath = $root }
        }
        catch { $failed = $true }
        if (-not $failed) { throw 'GitHub release bundle validator accepted a tampered release manifest.' }
    }
    finally {
        [IO.File]::WriteAllBytes($releasePath, $releaseOriginal)
    }

    $checksumPath = Join-Path $bundle "VBA-HTTP-$tag.SHA256SUMS.txt"
    $checksumOriginal = [IO.File]::ReadAllBytes($checksumPath)
    try {
        [IO.File]::AppendAllText($checksumPath, "0000000000000000000000000000000000000000000000000000000000000000 *tampered.txt`r`n")
        $failed = $false
        try {
            Invoke-Script 'Validate-GitHubReleaseBundle.ps1' @{ Tag = $tag; BundleDirectory = $bundle; RootPath = $root }
        }
        catch { $failed = $true }
        if (-not $failed) { throw 'GitHub release bundle validator accepted a tampered checksum file.' }
    }
    finally {
        [IO.File]::WriteAllBytes($checksumPath, $checksumOriginal)
    }

    Write-Output 'GitHub release bundle test passed: asset set, provenance, source/package validation, and tamper rejection.'
}
finally {
    if (Test-Path -LiteralPath $testRoot -PathType Container) {
        Remove-Item -LiteralPath $testRoot -Recurse -Force
    }
}
