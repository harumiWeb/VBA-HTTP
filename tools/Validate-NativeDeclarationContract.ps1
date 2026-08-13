[CmdletBinding()]
param(
    [string]$Path = "src/modules/WinHttpNativeApi.bas"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$projectRoot = [IO.Path]::GetFullPath((Split-Path -Parent $PSScriptRoot))
$resolvedPath = if ([IO.Path]::IsPathRooted($Path)) {
    [IO.Path]::GetFullPath($Path)
}
else {
    [IO.Path]::GetFullPath((Join-Path $projectRoot $Path))
}

function Assert-Contract([bool]$Condition, [string]$Message) {
    if (-not $Condition) { throw "Native declaration contract failed: $Message" }
}

if (-not (Test-Path -LiteralPath $resolvedPath -PathType Leaf)) {
    throw "Native declaration source does not exist: $resolvedPath"
}

$source = [IO.File]::ReadAllText($resolvedPath)
$source = $source -replace "`r`n?", "`n"
$declarationMatch = [regex]::Match($source, "(?ms)^#If\s+VBA7\s+Then\s*$.*?^#End\s+If\s*$")
Assert-Contract $declarationMatch.Success "the first declaration block must contain #If VBA7 Then / #Else / #End If."
$declarationParts = $declarationMatch.Value -split "(?m)^#Else\s*$", 2
Assert-Contract ($declarationParts.Count -eq 2) "the declaration block must contain one #Else branch."
$vba7Block = $declarationParts[0]
$legacyBlock = $declarationParts[1]
Assert-Contract ($vba7Block -match "(?im)Declare\s+PtrSafe\s+(Function|Sub)") "VBA7 declarations must be PtrSafe."
Assert-Contract ($legacyBlock -notmatch "(?i)PtrSafe|LongPtr") "legacy declarations must not contain PtrSafe or LongPtr."
Assert-Contract ($vba7Block -match "(?i)LongPtr") "VBA7 declarations must use pointer-sized types."

$pointerDeclarations = @(
    "WinHttpOpen",
    "WinHttpCloseHandle",
    "WinHttpConnect",
    "WinHttpOpenRequest",
    "WinHttpAddRequestHeaders",
    "WinHttpSetTimeouts",
    "WinHttpSendRequest",
    "WinHttpWriteData",
    "WinHttpReceiveResponse",
    "WinHttpQueryAuthSchemes",
    "WinHttpSetCredentials",
    "WinHttpQueryHeaders",
    "WinHttpQueryDataAvailable",
    "WinHttpReadData",
    "WinHttpQueryOption",
    "WinHttpSetOption",
    "GetCurrentProcess",
    "GetProcessHandleCount",
    "CopyMemory"
)

foreach ($name in $pointerDeclarations) {
    $escapedName = [Regex]::Escape($name)
    Assert-Contract ($vba7Block -match "(?im)^\s*Private\s+Declare\s+PtrSafe\s+(Function|Sub)\s+$escapedName\b[^\r\n]*LongPtr") "$name must expose pointer-sized handles or pointers in VBA7."
    Assert-Contract ($legacyBlock -match "(?im)^\s*Private\s+Declare\s+(Function|Sub)\s+$escapedName\b[^\r\n]*\bLong\b") "$name must use Long handles or pointers in legacy VBA."
}

Assert-Contract ($source -match "(?im)^Public\s+Const\s+WinHttpIgnoreRequestTotalLength\s+As\s+Long\s*=\s*0\s*$") "WINHTTP_IGNORE_REQUEST_TOTAL_LENGTH must remain the DWORD sentinel 0."
Write-Output "Native declaration contract passed: VBA7 PtrSafe/LongPtr and legacy Long branches are present for $($pointerDeclarations.Count) WinHTTP/Kernel32 declarations."
