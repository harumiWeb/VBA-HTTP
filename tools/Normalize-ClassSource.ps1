Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $PSScriptRoot
$classRoot = Join-Path $root "src/classes"
$encoding = [Text.UTF8Encoding]::new($false)
$updated = 0

if (Test-Path -LiteralPath $classRoot) {
    foreach ($file in Get-ChildItem -LiteralPath $classRoot -Recurse -File -Filter "*.cls") {
        $text = [IO.File]::ReadAllText($file.FullName, [Text.Encoding]::UTF8)
        $normalized = $text.Replace("`r`n", "`n").Replace("`r", "`n").Replace("`n", "`r`n")
        [IO.File]::WriteAllText($file.FullName, $normalized, $encoding)
        $updated++
    }
}

Write-Output "Normalized $updated class source file(s) to UTF-8 CRLF."
