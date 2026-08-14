[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
    [Parameter(Mandatory = $true)]
    [string]$Workbook,
    [string]$PackageRoot = "",
    [string]$BackupPath = "",
    [switch]$Force
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$resolvedPackageRoot = if ([string]::IsNullOrWhiteSpace($PackageRoot)) {
    [IO.Path]::GetFullPath($PSScriptRoot)
}
else {
    [IO.Path]::GetFullPath($PackageRoot)
}
$resolvedWorkbook = [IO.Path]::GetFullPath($Workbook)
if (-not (Test-Path -LiteralPath $resolvedWorkbook -PathType Leaf)) {
    throw "Target workbook does not exist: $resolvedWorkbook"
}
if ([IO.Path]::GetExtension($resolvedWorkbook).ToLowerInvariant() -notin @('.xlsm', '.xlam', '.xlsb')) {
    throw "Target workbook must be .xlsm, .xlam, or .xlsb: $resolvedWorkbook"
}

function Clear-ComObject([object]$Object) {
    if ($null -eq $Object -or -not [Runtime.InteropServices.Marshal]::IsComObject($Object)) {
        return
    }
    try {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Object)
    }
    catch {
        Write-Debug "Could not release a COM object: $($_.Exception.Message)"
    }
}

function Test-TargetWorkbookClosed([string]$Path) {
    $stream = $null
    try {
        $stream = [IO.File]::Open($Path, [IO.FileMode]::Open, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
        return $true
    }
    catch {
        throw "Target workbook must be closed and writable before uninstallation: $Path"
    }
    finally {
        if ($null -ne $stream) { $stream.Dispose() }
    }
}

[void](Test-TargetWorkbookClosed $resolvedWorkbook)

$validatorPath = Join-Path $PSScriptRoot 'Validate-SourcePackage.ps1'
if (-not (Test-Path -LiteralPath $validatorPath -PathType Leaf)) {
    throw 'The package validator is not available beside the uninstaller.'
}
& $validatorPath -PackageRoot $resolvedPackageRoot | Out-Null
$manifest = Get-Content -LiteralPath (Join-Path $resolvedPackageRoot 'manifest.json') -Raw | ConvertFrom-Json

$backupFile = if ([string]::IsNullOrWhiteSpace($BackupPath)) {
    "$resolvedWorkbook.vba-http-uninstall.bak"
}
elseif ([IO.Path]::IsPathRooted($BackupPath)) {
    [IO.Path]::GetFullPath($BackupPath)
}
else {
    [IO.Path]::GetFullPath((Join-Path (Get-Location) $BackupPath))
}

if (-not $PSCmdlet.ShouldProcess($resolvedWorkbook, "Uninstall VBA-HTTP package components")) {
    return
}
if ((Test-Path -LiteralPath $backupFile -PathType Leaf) -and -not $Force) {
    throw "Backup already exists. Choose another -BackupPath or pass -Force: $backupFile"
}
if (Test-Path -LiteralPath $backupFile -PathType Leaf) {
    Remove-Item -LiteralPath $backupFile -Force
}
Copy-Item -LiteralPath $resolvedWorkbook -Destination $backupFile -Force

$excel = $null
$targetWorkbook = $null
$components = $null
$component = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 3
    $targetWorkbook = $excel.Workbooks.Open($resolvedWorkbook, 0, $false)
    $components = $targetWorkbook.VBProject.VBComponents

    foreach ($manifestComponent in @($manifest.components)) {
        $name = [string]$manifestComponent.name
        $component = $null
        try {
            $component = $components.Item($name)
        }
        catch {
            Write-Debug "Component '$name' is not installed: $($_.Exception.Message)"
        }
        if ($null -ne $component) {
            if (-not $Force) {
                throw "Component '$name' exists in the target workbook. Pass -Force to remove package components."
            }
            $components.Remove($component)
            Clear-ComObject $component
            $component = $null
        }
    }
    $targetWorkbook.Save()
    Write-Output "VBA-HTTP uninstalled: backup: $backupFile"
}
finally {
    Clear-ComObject $component
    Clear-ComObject $components
    if ($null -ne $targetWorkbook) {
        try { $targetWorkbook.Close($false) } catch { Write-Debug "Target workbook was already closed: $_" }
        Clear-ComObject $targetWorkbook
    }
    if ($null -ne $excel) {
        try { $excel.Quit() } catch { Write-Debug "Uninstaller Excel instance was already closed: $_" }
        Clear-ComObject $excel
    }
}
