[CmdletBinding()]
param(
    [string]$Path = "tools"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$projectRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot ".."))
$resolvedPath = [IO.Path]::GetFullPath((Join-Path $projectRoot $Path))
$settingsPath = Join-Path $projectRoot "PSScriptAnalyzerSettings.psd1"
$rootWithSeparator = $projectRoot.TrimEnd('\', '/') + [IO.Path]::DirectorySeparatorChar

if (-not (Test-Path -LiteralPath $settingsPath -PathType Leaf)) {
    throw "PSScriptAnalyzer settings file is missing: $settingsPath"
}

Import-Module PSScriptAnalyzer -ErrorAction Stop
$findings = @(Invoke-ScriptAnalyzer -Path $resolvedPath -Recurse -Settings $settingsPath -Severity Error, Warning)
if ($findings.Count -gt 0) {
    foreach ($finding in $findings) {
        $scriptPath = [IO.Path]::GetFullPath([string]$finding.ScriptPath)
        $relativeScript = if ($scriptPath.StartsWith($rootWithSeparator, [StringComparison]::OrdinalIgnoreCase)) {
            $scriptPath.Substring($rootWithSeparator.Length).Replace('\', '/')
        }
        else {
            $scriptPath
        }
        Write-Error ("{0}:{1}:{2} {3} [{4}] {5}" -f $relativeScript, $finding.Line, $finding.Column, $finding.RuleName, $finding.Severity, $finding.Message)
    }
    exit 1
}

Write-Output "PSScriptAnalyzer clean: $resolvedPath"
