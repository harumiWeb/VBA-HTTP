Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $PSScriptRoot
$classRoot = Join-Path $root "src/classes"
$invalid = [System.Collections.Generic.List[string]]::new()

if (Test-Path -LiteralPath $classRoot) {
    foreach ($file in Get-ChildItem -LiteralPath $classRoot -Recurse -File -Filter "*.cls") {
        $bytes = [IO.File]::ReadAllBytes($file.FullName)
        $hasBareLf = $false
        for ($index = 0; $index -lt $bytes.Length; $index++) {
            if ($bytes[$index] -eq 10 -and ($index -eq 0 -or $bytes[$index - 1] -ne 13)) {
                $hasBareLf = $true
                break
            }
        }

        $text = [Text.UTF8Encoding]::new($false, $true).GetString($bytes)
        if ($hasBareLf -or -not $text.StartsWith("VERSION 1.0 CLASS`r`n", [StringComparison]::Ordinal)) {
            $relativePath = $file.FullName.Substring($root.Length).TrimStart([char[]]@("\", "/"))
            $invalid.Add($relativePath.Replace("\", "/"))
        }
    }
}

if ($invalid.Count -gt 0) {
    throw "VBE-importable class source requires UTF-8 CRLF and a VERSION 1.0 CLASS header. Run 'task class-source:normalize'. Invalid: $($invalid -join ', ')"
}

Write-Output "Class source format OK"
