[CmdletBinding()]
param(
    [string]$ArtifactPath = "build/Release/VBA-HTTP.xlam"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$projectRoot = Split-Path -Parent $PSScriptRoot

function Resolve-ProjectPath([string]$Path) {
    if ([IO.Path]::IsPathRooted($Path)) { return [IO.Path]::GetFullPath($Path) }
    return [IO.Path]::GetFullPath((Join-Path $projectRoot $Path))
}

function ConvertTo-RepoPath([string]$Path) {
    $resolved = [IO.Path]::GetFullPath($Path)
    $root = [IO.Path]::GetFullPath($projectRoot).TrimEnd('\', '/') + [IO.Path]::DirectorySeparatorChar
    if (-not $resolved.StartsWith($root, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Path must stay under the repository root: $Path"
    }
    return $resolved.Substring($root.Length).Replace('\', '/')
}

$resolvedArtifact = Resolve-ProjectPath $ArtifactPath
$manifestPath = "$resolvedArtifact.build.json"
if ([IO.Path]::GetExtension($resolvedArtifact).ToLowerInvariant() -ne ".xlam") {
    throw "XLAM artifact validation requires a .xlam artifact."
}
if ((ConvertTo-RepoPath $resolvedArtifact) -ne "build/Release/VBA-HTTP.xlam") {
    throw "XLAM artifact must be build/Release/VBA-HTTP.xlam."
}
foreach ($path in @($resolvedArtifact, $manifestPath)) {
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { throw "XLAM release input does not exist: $path" }
}

$manifest = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json
if ([int]$manifest.schema_version -ne 1 -or [string]$manifest.command -ne "build" -or
    [string]$manifest.backend -ne "excel" -or [string]$manifest.base -ne "build/VBA-HTTP.xlam" -or
    [string]$manifest.output -ne "build/Release/VBA-HTTP.xlam") {
    throw "XLAM build manifest identity is invalid."
}
if ($manifest.validation.source_applied -ne $true -or
    $manifest.validation.vbe_compile -ne "passed" -or
    $manifest.validation.workbook_saved -ne $true -or
    $manifest.validation.workbook_closed -ne $true -or
    $manifest.validation.excel_cleanup -ne "clean" -or
    @("atomic_create", "atomic_replace") -notcontains [string]$manifest.publication.method) {
    throw "XLAM manifest does not prove a clean compiled atomic build."
}

& (Join-Path $PSScriptRoot "Verify-ReleaseChecksums.ps1") -ArtifactPath $resolvedArtifact -ManifestPath $manifestPath
& (Join-Path $PSScriptRoot "Validate-SecurityRiskRegister.ps1")

$policy = Get-Content -LiteralPath (Join-Path $PSScriptRoot "build-component-policy.json") -Raw | ConvertFrom-Json
if ($null -ne (Compare-Object @($policy.included | Sort-Object) @($manifest.included_components.name | Sort-Object)) -or
    $null -ne (Compare-Object @($policy.excluded | Sort-Object) @($manifest.excluded_components.name | Sort-Object))) {
    throw "XLAM manifest component sets differ from the production policy."
}

$excel = $null
$workbooks = $null
$workbook = $null
$components = $null
$component = $null
$actualComponents = @()
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 3
    $workbooks = $excel.Workbooks
    $workbook = $workbooks.Open($resolvedArtifact, 0, $true)
    if (-not [bool]$workbook.IsAddin) { throw "XLAM artifact is not marked as an Excel add-in." }
    $components = $workbook.VBProject.VBComponents
    for ($index = 1; $index -le $components.Count; $index++) {
        $component = $components.Item($index)
        $actualComponents += $component.Name
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($component)
        $component = $null
    }
}
finally {
    if ($null -ne $component) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($component) }
    if ($null -ne $components) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($components) }
    if ($null -ne $workbook) { try { $workbook.Close($false) } catch { Write-Debug "XLAM workbook was already closed: $_" }; [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($workbook) }
    if ($null -ne $workbooks) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($workbooks) }
    if ($null -ne $excel) { try { $excel.Quit() } catch { Write-Debug "XLAM Excel instance was already closed: $_" }; [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }
}

if ($null -ne (Compare-Object @($policy.included | Sort-Object) @($actualComponents | Sort-Object))) {
    throw "XLAM workbook components differ from the production policy."
}
$runJson = & xlflow run Main.Run --input $resolvedArtifact --no-save --direct --json
if ($LASTEXITCODE -ne 0) { throw "XLAM consumer macro failed with exit code $LASTEXITCODE." }
$runResult = $runJson | Out-String | ConvertFrom-Json
if ($runResult.status -ne "ok" -or $runResult.macro.name -ne "Main.Run") {
    throw "XLAM consumer macro returned an unexpected result."
}

Write-Output "XLAM artifact is valid: $($actualComponents.Count) production components; add-in identity, manifest, checksum, policy, and consumer smoke passed."
