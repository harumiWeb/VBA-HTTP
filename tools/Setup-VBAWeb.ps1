[CmdletBinding()]
param(
    [string]$ReferencePath = "references/VBA-Web-v4.1.6",
    [switch]$VerifyOnly
)

$ErrorActionPreference = "Stop"
$upstreamRepository = "https://github.com/VBA-tools/VBA-Web.git"
$upstreamTag = "v4.1.6"
$expectedCommit = "cefc320acc5372e0b86eed1d20eb3f31b331d598"
$projectRoot = Split-Path -Parent $PSScriptRoot
$referencesRoot = [System.IO.Path]::GetFullPath((Join-Path $projectRoot "references"))
$resolvedReference = [System.IO.Path]::GetFullPath((Join-Path $projectRoot $ReferencePath))

if (-not $resolvedReference.StartsWith($referencesRoot + [System.IO.Path]::DirectorySeparatorChar, [System.StringComparison]::OrdinalIgnoreCase)) {
    throw "VBA-Web reference must be stored below $referencesRoot."
}

function Invoke-GitCapture {
    param([Parameter(Mandatory)][string[]]$Arguments)

    $output = & git @Arguments
    if ($LASTEXITCODE -ne 0) {
        throw "git $($Arguments -join ' ') failed: $($output | Out-String)"
    }
    return ($output | Out-String).Trim()
}

function Assert-PinnedCheckout {
    param([Parameter(Mandatory)][string]$Path)

    $gitDirectory = Join-Path $Path ".git"
    if (-not (Test-Path -LiteralPath $gitDirectory -PathType Container)) {
        throw "VBA-Web reference is not a Git checkout: $Path"
    }

    $actualCommit = Invoke-GitCapture -Arguments @("-C", $Path, "rev-parse", "HEAD")
    if ($actualCommit -ne $expectedCommit) {
        throw "VBA-Web checkout is at $actualCommit; expected $expectedCommit ($upstreamTag)."
    }

    $status = Invoke-GitCapture -Arguments @("-C", $Path, "status", "--porcelain")
    if ($status) {
        throw "VBA-Web checkout contains local changes; benchmark inputs must be pristine."
    }

    $origin = Invoke-GitCapture -Arguments @("-C", $Path, "remote", "get-url", "origin")
    if ($origin.TrimEnd("/") -ne $upstreamRepository.TrimEnd("/")) {
        throw "VBA-Web origin is '$origin'; expected '$upstreamRepository'."
    }

    $workbookPath = Join-Path $Path "VBA-Web - Blank.xlsm"
    $licensePath = Join-Path $Path "LICENSE"
    if (-not (Test-Path -LiteralPath $workbookPath -PathType Leaf)) {
        throw "Pinned checkout does not contain VBA-Web - Blank.xlsm."
    }
    if (-not (Test-Path -LiteralPath $licensePath -PathType Leaf)) {
        throw "Pinned checkout does not contain its MIT LICENSE."
    }

    return $workbookPath
}

[void](New-Item -ItemType Directory -Path $referencesRoot -Force)

if (-not (Test-Path -LiteralPath $resolvedReference)) {
    if ($VerifyOnly) {
        throw "Pinned VBA-Web checkout is missing. Run 'task benchmark:vba-web:setup' before the offline benchmark."
    }
    $stagingPath = "$resolvedReference.staging.$([guid]::NewGuid().ToString('N'))"
    try {
        [void](Invoke-GitCapture -Arguments @(
            "clone",
            "--branch", $upstreamTag,
            "--depth", "1",
            "--no-tags",
            $upstreamRepository,
            $stagingPath
        ))
        [void](Assert-PinnedCheckout -Path $stagingPath)
        Move-Item -LiteralPath $stagingPath -Destination $resolvedReference
    }
    finally {
        if (Test-Path -LiteralPath $stagingPath) {
            Remove-Item -LiteralPath $stagingPath -Recurse -Force
        }
    }
}

$resolvedWorkbook = Assert-PinnedCheckout -Path $resolvedReference
$sha256 = [System.Security.Cryptography.SHA256]::Create()
$stream = $null
try {
    $stream = [System.IO.File]::OpenRead($resolvedWorkbook)
    $workbookHash = ([System.BitConverter]::ToString($sha256.ComputeHash($stream))).Replace("-", "").ToLowerInvariant()
}
finally {
    if ($null -ne $stream) { $stream.Dispose() }
    $sha256.Dispose()
}

[pscustomobject]@{
    repository = $upstreamRepository
    tag = $upstreamTag
    commit = $expectedCommit
    checkout = $resolvedReference
    workbook = $resolvedWorkbook
    workbook_sha256 = $workbookHash
} | ConvertTo-Json
