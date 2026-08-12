[CmdletBinding()]
param(
    [string]$ArtifactPath = "build/Release/VBA-HTTP.xlsm"
)

$ErrorActionPreference = "Stop"
$projectRoot = Split-Path -Parent $PSScriptRoot
$resolvedArtifact = [System.IO.Path]::GetFullPath((Join-Path $projectRoot $ArtifactPath))
$manifestPath = "$resolvedArtifact.build.json"
$policyPath = Join-Path $PSScriptRoot "build-component-policy.json"

if (-not (Test-Path -LiteralPath $resolvedArtifact -PathType Leaf)) {
    throw "Release artifact does not exist: $resolvedArtifact"
}
if (-not (Test-Path -LiteralPath $manifestPath -PathType Leaf)) {
    throw "Release manifest does not exist: $manifestPath"
}

$policy = Get-Content -LiteralPath $policyPath -Raw | ConvertFrom-Json
$manifest = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json

if ($manifest.validation.source_applied -ne $true -or
    $manifest.validation.vbe_compile -ne "passed" -or
    $manifest.validation.workbook_saved -ne $true -or
    $manifest.validation.workbook_closed -ne $true -or
    $manifest.validation.excel_cleanup -ne "clean") {
    throw "Release manifest does not prove a clean compiled build."
}
if ($manifest.publication.method -ne "atomic_replace") {
    throw "Release publication was not atomic."
}

$expectedIncluded = @($policy.included | Sort-Object)
$expectedExcluded = @($policy.excluded | Sort-Object)
$manifestIncluded = @($manifest.included_components.name | Sort-Object)
$manifestExcluded = @($manifest.excluded_components.name | Sort-Object)

if ($null -ne (Compare-Object $expectedIncluded $manifestIncluded)) {
    throw "Release manifest included components differ from policy."
}
if ($null -ne (Compare-Object $expectedExcluded $manifestExcluded)) {
    throw "Release manifest excluded components differ from policy."
}

$excel = $null
$workbooks = $null
$workbook = $null
$components = $null
$component = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 3
    $workbooks = $excel.Workbooks
    $workbook = $workbooks.Open($resolvedArtifact, 0, $true)
    $components = $workbook.VBProject.VBComponents

    $actualComponents = @()
    for ($index = 1; $index -le $components.Count; $index++) {
        $component = $components.Item($index)
        $actualComponents += $component.Name
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($component)
        $component = $null
    }

    $actualComponents = @($actualComponents | Sort-Object)
    if ($null -ne (Compare-Object $expectedIncluded $actualComponents)) {
        throw "Release workbook components differ from policy."
    }
}
finally {
    if ($null -ne $component) {
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($component)
    }
    if ($null -ne $components) {
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($components)
    }
    if ($null -ne $workbook) {
        try { $workbook.Close($false) } catch { Write-Warning "Could not close release workbook: $_" }
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($workbook)
    }
    if ($null -ne $workbooks) {
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($workbooks)
    }
    if ($null -ne $excel) {
        try { $excel.Quit() } catch { Write-Warning "Could not quit Excel inspection process: $_" }
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
    }
}

$runJson = & xlflow run Main.Run --input $resolvedArtifact --headless --no-save --direct --json
if ($LASTEXITCODE -ne 0) {
    throw "Release consumer smoke failed with exit code $LASTEXITCODE."
}
$runResult = $runJson | Out-String | ConvertFrom-Json
if ($runResult.status -ne "ok" -or $runResult.macro.name -ne "Main.Run") {
    throw "Release consumer smoke returned an unexpected result."
}

Write-Output "Release artifact is valid: $($actualComponents.Count) components; Main.Run passed."
