[CmdletBinding()]
param(
    [string]$ProjectRoot
)

$ErrorActionPreference = "Stop"
if ([string]::IsNullOrWhiteSpace($ProjectRoot)) {
    $ProjectRoot = Split-Path -Parent $PSScriptRoot
}

function Read-ProjectFile {
    param([Parameter(Mandatory)][string]$RelativePath)

    $path = Join-Path $ProjectRoot $RelativePath
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
        throw "Required source file was not found: $RelativePath"
    }
    return Get-Content -LiteralPath $path -Raw
}

function Get-SourceSection {
    param(
        [Parameter(Mandatory)][string]$Source,
        [Parameter(Mandatory)][string]$StartMarker,
        [Parameter(Mandatory)][string]$EndMarker
    )

    $start = $Source.IndexOf($StartMarker, [StringComparison]::Ordinal)
    if ($start -lt 0) {
        throw "Source marker was not found: $StartMarker"
    }
    $end = $Source.IndexOf($EndMarker, $start + $StartMarker.Length, [StringComparison]::Ordinal)
    if ($end -lt 0) {
        throw "Source marker was not found: $EndMarker"
    }
    return $Source.Substring($start, $end - $start)
}

function Assert-ContainsToken {
    param(
        [Parameter(Mandatory)][string]$Source,
        [Parameter(Mandatory)][string]$Needle,
        [Parameter(Mandatory)][string]$Description
    )

    if ($Source.IndexOf($Needle, [StringComparison]::Ordinal) -lt 0) {
        throw "Native performance contract is missing $Description."
    }
}

function Assert-NotContainsToken {
    param(
        [Parameter(Mandatory)][string]$Source,
        [Parameter(Mandatory)][string]$Needle,
        [Parameter(Mandatory)][string]$Description
    )

    if ($Source.IndexOf($Needle, [StringComparison]::Ordinal) -ge 0) {
        throw "Native performance contract still contains $Description."
    }
}

$transport = Read-ProjectFile "src\classes\WinHttpNativeTransport.cls"
$download = Get-SourceSection $transport "Private Function DownloadCore" "Private Function UploadFileCore"
$buffered = Get-SourceSection $transport "Private Function ReadBufferedBody" "Private Function TryGetInitialBodyCapacity"
$reader = Read-ProjectFile "src\classes\WinHttpUploadFileReader.cls"

Assert-ContainsToken $download "WinHttpNativeApi.ReadData(nativeRequest.Value, chunk, ReadChunkBytes, readBytes)" "direct streaming WinHttpReadData"
Assert-NotContainsToken $download "WinHttpNativeApi.QueryDataAvailable" "a streaming QueryDataAvailable hot-path call"
Assert-ContainsToken $download "WinHttpNativeApi.CopyByteRange finalChunk, 0, chunk, readBytes" "bounded short-chunk copy"

Assert-ContainsToken $buffered "WinHttpNativeApi.ReadData(nativeRequest.Value, chunk, ReadChunkBytes, readBytes)" "direct buffered WinHttpReadData"
Assert-NotContainsToken $buffered "WinHttpNativeApi.QueryDataAvailable" "a buffered QueryDataAvailable hot-path call"
Assert-ContainsToken $buffered "TryGetInitialBodyCapacity(responseHeaders, capacity)" "known-length output pre-sizing"

$initialCapacity = Get-SourceSection $transport "Private Function TryGetInitialBodyCapacity" "Private Sub EnsureCapacity"
Assert-ContainsToken $initialCapacity "MaxInitialBodyCapacity" "bounded initial allocation for known-length responses"

Assert-ContainsToken $reader "Private mBufferCapacity As Long" "upload buffer capacity state"
Assert-ContainsToken $reader "If requested <> mBufferCapacity Then" "full-chunk buffer reuse"
Assert-ContainsToken $reader "Private Const UploadChunkBytes As Long = 65536" "the bounded upload chunk size"

Write-Output "Native performance source contract valid: direct reads, bounded pre-sizing, short-copy, and upload buffer reuse are present."
