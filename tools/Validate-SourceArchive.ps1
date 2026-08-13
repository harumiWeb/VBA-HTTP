[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$ArchivePath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$resolvedArchive = [IO.Path]::GetFullPath($ArchivePath)
if (-not (Test-Path -LiteralPath $resolvedArchive -PathType Leaf)) {
    throw "Source package archive does not exist: $resolvedArchive"
}
$extractRoot = Join-Path ([IO.Path]::GetTempPath()) ("vba-http-source-archive-" + [guid]::NewGuid().ToString('N'))
try {
    Expand-Archive -LiteralPath $resolvedArchive -DestinationPath $extractRoot -Force
    & (Join-Path $PSScriptRoot 'Validate-SourcePackage.ps1') -PackageRoot $extractRoot
}
finally {
    if (Test-Path -LiteralPath $extractRoot) {
        Remove-Item -LiteralPath $extractRoot -Recurse -Force
    }
}
