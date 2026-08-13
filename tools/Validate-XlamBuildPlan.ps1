[CmdletBinding()]
param(
    [string]$BasePath = "build/VBA-HTTP.xlam",
    [string]$OutputPath = "build/Release/VBA-HTTP.xlam"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$projectRoot = Split-Path -Parent $PSScriptRoot
$policyPath = Join-Path $PSScriptRoot "build-component-policy.json"

function Resolve-ProjectPath([string]$Path) {
    if ([string]::IsNullOrWhiteSpace($Path)) { throw "A path cannot be empty." }
    if ([IO.Path]::IsPathRooted($Path)) { return [IO.Path]::GetFullPath($Path) }
    return [IO.Path]::GetFullPath((Join-Path $projectRoot $Path))
}

function To-RepoPath([string]$Path) {
    $resolved = [IO.Path]::GetFullPath($Path)
    $root = [IO.Path]::GetFullPath($projectRoot).TrimEnd('\', '/') + [IO.Path]::DirectorySeparatorChar
    if (-not $resolved.StartsWith($root, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Path must stay under the repository root: $Path"
    }
    return $resolved.Substring($root.Length).Replace('\', '/')
}

$resolvedBase = Resolve-ProjectPath $BasePath
$resolvedOutput = Resolve-ProjectPath $OutputPath
if ([IO.Path]::GetExtension($resolvedBase).ToLowerInvariant() -ne ".xlam" -or
    [IO.Path]::GetExtension($resolvedOutput).ToLowerInvariant() -ne ".xlam") {
    throw "The XLAM build target requires .xlam base and output paths."
}
if (-not (Test-Path -LiteralPath $resolvedBase -PathType Leaf)) {
    throw "XLAM base workbook does not exist: $resolvedBase"
}

$expectedBase = To-RepoPath $resolvedBase
$expectedOutput = To-RepoPath $resolvedOutput
if ($expectedBase -ne "build/VBA-HTTP.xlam" -or $expectedOutput -ne "build/Release/VBA-HTTP.xlam") {
    throw "XLAM build target must use build/VBA-HTTP.xlam -> build/Release/VBA-HTTP.xlam."
}

& (Join-Path $PSScriptRoot "Ensure-XlflowDirectories.ps1")
$json = & xlflow build --dry-run --json --base $resolvedBase --out $resolvedOutput
if ($LASTEXITCODE -ne 0) { throw "xlflow XLAM dry-run failed with exit code $LASTEXITCODE." }
$result = $json | Out-String | ConvertFrom-Json
if ($result.status -ne "ok") { throw "xlflow returned XLAM build status '$($result.status)'." }
if (@($result.build.warnings).Count -ne 0) {
    throw "XLAM build plan contains warnings: $($result.build.warnings | ConvertTo-Json -Compress)"
}
if ([string]$result.build.base -ne "build/VBA-HTTP.xlam" -or
    [string]$result.build.output -ne "build/Release/VBA-HTTP.xlam") {
    throw "XLAM dry-run resolved unexpected base/output paths."
}

$policy = Get-Content -LiteralPath $policyPath -Raw | ConvertFrom-Json
$actualIncluded = @($result.build.included_components.name | Sort-Object)
$actualExcluded = @($result.build.excluded_components.name | Sort-Object)
if ($null -ne (Compare-Object @($policy.included | Sort-Object) $actualIncluded) -or
    $null -ne (Compare-Object @($policy.excluded | Sort-Object) $actualExcluded)) {
    throw "XLAM build plan component sets differ from the production policy."
}

Write-Output "XLAM build plan is valid: $($actualIncluded.Count) included, $($actualExcluded.Count) excluded."
