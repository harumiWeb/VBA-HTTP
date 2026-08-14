[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$Tag,
    [string]$ExpectedCommit = ""
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$pattern = '^v(0|[1-9][0-9]*)\.(0|[1-9][0-9]*)\.(0|[1-9][0-9]*)(-[0-9A-Za-z.-]+)?$'
if ($Tag -notmatch $pattern) {
    throw "Release tag must match vX.Y.Z or vX.Y.Z-prerelease: $Tag"
}

$version = $Tag.Substring(1)
$isPrerelease = $version.Contains('-')
$resolvedCommit = ""
if (-not [string]::IsNullOrWhiteSpace($ExpectedCommit)) {
    if ($ExpectedCommit -notmatch '^[0-9a-fA-F]{40,64}$') {
        throw 'ExpectedCommit must be a full hexadecimal commit SHA.'
    }
    $gitOutput = & git rev-parse "$Tag^{commit}" 2>$null
    $gitExitCode = $LASTEXITCODE
    $resolvedCommit = ($gitOutput | Select-Object -First 1)
    if ($gitExitCode -ne 0 -or [string]::IsNullOrWhiteSpace([string]$resolvedCommit)) {
        throw "Release tag does not resolve to a commit: $Tag"
    }
    $resolvedCommit = ([string]$resolvedCommit).Trim()
    $expected = $ExpectedCommit.Trim()
    if (-not $resolvedCommit.Equals($expected, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Release tag commit does not match the expected commit: tag=$resolvedCommit expected=$expected"
    }
}

$result = [ordered]@{
    schema_version = 1
    tag = $Tag
    version = $version
    prerelease = $isPrerelease
    commit = if ([string]::IsNullOrWhiteSpace($resolvedCommit)) { $null } else { $resolvedCommit }
}
Write-Output (($result | ConvertTo-Json -Depth 4 -Compress))
