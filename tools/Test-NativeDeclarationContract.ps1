[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$projectRoot = [IO.Path]::GetFullPath((Split-Path -Parent $PSScriptRoot))
$sourcePath = Join-Path $projectRoot "src/modules/WinHttpNativeApi.bas"
$validator = Join-Path $PSScriptRoot "Validate-NativeDeclarationContract.ps1"
$testRoot = Join-Path $projectRoot ".xlflow/native-declaration-contract-test"

if (Test-Path -LiteralPath $testRoot) { Remove-Item -LiteralPath $testRoot -Recurse -Force }
[void](New-Item -ItemType Directory -Path $testRoot -Force)

function Invoke-Validator([string]$Path) {
    $oldErrorActionPreference = $ErrorActionPreference
    try {
        $ErrorActionPreference = "Continue"
        & powershell -NoProfile -ExecutionPolicy Bypass -File $validator -Path $Path 1>$null 2>$null
        return $LASTEXITCODE
    }
    finally {
        $ErrorActionPreference = $oldErrorActionPreference
    }
}

function Convert-Text([string]$Value, [string]$OldValue, [string]$NewValue) {
    return $Value.Replace($OldValue, $NewValue)
}

try {
    if ((Invoke-Validator $sourcePath) -ne 0) {
        throw "The checked-in declaration source did not pass its own contract."
    }

    $source = [IO.File]::ReadAllText($sourcePath)
    $optionBufferAssignments = ([regex]::Matches($source, '(?m)^\s*optionBuffer = OptionValue\s*$')).Count
    if ($optionBufferAssignments -ne 2) {
        throw "SetOptionLong must copy OptionValue into a Long buffer in both VBA7 and legacy branches."
    }
    $optionBufferPointers = ([regex]::Matches($source, '(?m)^\s*SetOptionLong = \(WinHttpSetOption\(HandleValue, OptionCode, VarPtr\(optionBuffer\), 4\) <> 0\)\s*$')).Count
    if ($optionBufferPointers -ne 2) {
        throw "SetOptionLong must pass the explicit DWORD buffer address in both VBA7 and legacy branches."
    }
    $mutations = @{
        "missing-vba7-guard.bas" = Convert-Text $source "#If VBA7 Then" "#If VBA6 Then"
        "legacy-longptr.bas" = Convert-Text $source "ByVal pwszUserAgent As Long, ByVal dwAccessType" "ByVal pwszUserAgent As LongPtr, ByVal dwAccessType"
        "wrong-upload-sentinel.bas" = Convert-Text $source "WinHttpIgnoreRequestTotalLength As Long = 0" "WinHttpIgnoreRequestTotalLength As Long = -1"
    }

    foreach ($entry in $mutations.GetEnumerator()) {
        $fixture = Join-Path $testRoot $entry.Key
        [IO.File]::WriteAllText($fixture, $entry.Value, [Text.UTF8Encoding]::new($false))
        if ((Invoke-Validator $fixture) -eq 0) {
            throw "Validator accepted a known-invalid fixture: $($entry.Key)"
        }
    }

    Write-Output "Native declaration contract tests passed: valid source plus $($mutations.Count) fail-closed fixtures."
}
finally {
    if (Test-Path -LiteralPath $testRoot) { Remove-Item -LiteralPath $testRoot -Recurse -Force }
}
