[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$root = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path

function Read-ProjectFile([string]$RelativePath) {
    $path = Join-Path $root $RelativePath
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
        throw "License contract failed: missing $RelativePath"
    }
    return [IO.File]::ReadAllText($path)
}

$license = Read-ProjectFile 'LICENSE'
foreach ($marker in @(
        'MIT License',
        'Copyright (c) 2026 harumiWeb',
        'Permission is hereby granted, free of charge',
        'The above copyright notice and this permission notice shall be included',
        'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND'
    )) {
    if ($license.IndexOf($marker, [StringComparison]::Ordinal) -lt 0) {
        throw "License contract failed: LICENSE is missing '$marker'"
    }
}

$readme = Read-ProjectFile 'README.md'
if ($readme.IndexOf('[MIT License](LICENSE)', [StringComparison]::Ordinal) -lt 0) {
    throw 'License contract failed: README.md must link the canonical LICENSE.'
}

$spec = Read-ProjectFile 'docs/specs/licensing.md'
foreach ($marker in @('MIT', 'LICENSE', 'THIRD_PARTY_NOTICES.md', 'task test:license')) {
    if ($spec.IndexOf($marker, [StringComparison]::OrdinalIgnoreCase) -lt 0) {
        throw "License contract failed: docs/specs/licensing.md is missing '$marker'"
    }
}

$notices = Read-ProjectFile 'THIRD_PARTY_NOTICES.md'
foreach ($marker in @('VBA-Web', 'MIT', 'not a VBA-HTTP dependency')) {
    if ($notices.IndexOf($marker, [StringComparison]::OrdinalIgnoreCase) -lt 0) {
        throw "License contract failed: THIRD_PARTY_NOTICES.md is missing '$marker'"
    }
}

$taskfile = Read-ProjectFile 'Taskfile.yml'
if ($taskfile.IndexOf('test:license', [StringComparison]::Ordinal) -lt 0) {
    throw 'License contract failed: Taskfile.yml must expose test:license.'
}

Write-Output 'License contract valid: MIT text, attribution, notices, and package gates are present.'
