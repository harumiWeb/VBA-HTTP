[CmdletBinding()]
param(
    [string]$RootPath = "",
    [string]$OutputRoot = ".xlflow/precommit-compile"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$projectRoot = if ([string]::IsNullOrWhiteSpace($RootPath)) {
    [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
}
else {
    [IO.Path]::GetFullPath($RootPath)
}
$outputBase = if ([IO.Path]::IsPathRooted($OutputRoot)) {
    [IO.Path]::GetFullPath($OutputRoot)
}
else {
    [IO.Path]::GetFullPath((Join-Path $projectRoot $OutputRoot))
}
$xlflowRoot = [IO.Path]::GetFullPath((Join-Path $projectRoot '.xlflow'))
$xlflowPrefix = $xlflowRoot.TrimEnd('\', '/') + [IO.Path]::DirectorySeparatorChar
if (-not $outputBase.StartsWith($xlflowPrefix, [StringComparison]::OrdinalIgnoreCase)) {
    throw 'Pre-commit compile output must remain under .xlflow.'
}

function Get-ExcelProcessId {
    return @(
        Get-Process -Name EXCEL -ErrorAction SilentlyContinue |
            ForEach-Object { [int]$_.Id }
    )
}

function Invoke-JsonCommand([string]$FilePath, [string[]]$Arguments, [string]$WorkingDirectory) {
    Push-Location $WorkingDirectory
    try {
        $lines = & $FilePath @Arguments
        if ($LASTEXITCODE -ne 0) { throw "$FilePath failed with exit code $LASTEXITCODE." }
        $text = ($lines | Out-String).Trim()
        if ([string]::IsNullOrWhiteSpace($text)) { throw "$FilePath returned no JSON output." }
        return $text | ConvertFrom-Json
    }
    finally {
        Pop-Location
    }
}

function Get-RequiredProperty($Object, [string]$Name) {
    if ($null -eq $Object) { throw "JSON record is null while reading '$Name'." }
    $property = $Object.PSObject.Properties[$Name]
    if ($null -eq $property) { throw "JSON record is missing '$Name'." }
    return $property.Value
}

$baselineExcel = @(Get-ExcelProcessId)
$targetWorkbook = [IO.Path]::GetFullPath((Join-Path $projectRoot 'build/VBA-HTTP.xlsm'))
if (-not (Test-Path -LiteralPath $targetWorkbook -PathType Leaf)) { throw "Target workbook is missing: $targetWorkbook" }
$lockFile = Join-Path ([IO.Path]::GetDirectoryName($targetWorkbook)) ('~$' + [IO.Path]::GetFileName($targetWorkbook))
if (Test-Path -LiteralPath $lockFile -PathType Leaf) {
    throw 'The target workbook has an Excel owner lock; no Excel process was changed.'
}

$status = Invoke-JsonCommand 'xlflow' @('status', '--json') $projectRoot
if ([bool](Get-RequiredProperty (Get-RequiredProperty $status 'coordination') 'busy') -or
    [bool](Get-RequiredProperty (Get-RequiredProperty $status 'coordination') 'recovery_required')) {
    throw 'xlflow coordination is busy or recovery-required; refusing to start Excel.'
}
$session = Get-RequiredProperty $status 'session'
if ([bool](Get-RequiredProperty $session 'active') -or [bool](Get-RequiredProperty $session 'workbook_open')) {
    throw 'The target workbook is already open in an xlflow session; close it and retry.'
}

$lock = $null
try {
    $lock = [IO.File]::Open($targetWorkbook, [IO.FileMode]::Open, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
}
catch {
    throw "The target workbook is locked or open by another process; no Excel process was changed."
}
finally {
    if ($null -ne $lock) { $lock.Dispose() }
}

$doctor = Invoke-JsonCommand 'xlflow' @('doctor', '--json') $projectRoot
$bridge = Get-RequiredProperty $doctor 'bridge'
if ([string](Get-RequiredProperty $bridge 'architecture') -ne 'X64') {
    throw 'Pre-commit VBE compile requires the supported x64 xlflow bridge.'
}

$runRoot = Join-Path $outputBase ([guid]::NewGuid().ToString('N'))
$outputWorkbook = Join-Path $runRoot 'VBA-HTTP.xlsm'
[void](New-Item -ItemType Directory -Path $runRoot -Force)
try {
    $build = Invoke-JsonCommand 'xlflow' @('build', '--json', '--base', $targetWorkbook, '--out', $outputWorkbook) $projectRoot
    $buildRecord = Get-RequiredProperty $build 'build'
    $validation = Get-RequiredProperty $buildRecord 'validation'
    foreach ($property in @('source_applied', 'workbook_saved', 'workbook_closed')) {
        if (-not [bool](Get-RequiredProperty $validation $property)) { throw "Pre-commit compile validation failed: $property." }
    }
    if ([string](Get-RequiredProperty $validation 'vbe_compile') -ne 'passed' -or
        [string](Get-RequiredProperty $validation 'excel_cleanup') -ne 'clean') {
        throw 'Pre-commit compile validation did not report VBE compile passed and Excel cleanup clean.'
    }
    $publication = Get-RequiredProperty $buildRecord 'publication'
    $publicationMethod = [string](Get-RequiredProperty $publication 'method')
    if ($publicationMethod -notin @('atomic_create', 'atomic_replace')) {
        throw 'Pre-commit compile did not report atomic publication.'
    }
    if (-not (Test-Path -LiteralPath $outputWorkbook -PathType Leaf)) { throw 'Pre-commit compile did not produce its temporary workbook.' }

    $newExcel = @(Get-ExcelProcessId | Where-Object { $baselineExcel -notcontains $_ })
    if ($newExcel.Count -gt 0) {
        throw "xlflow left Excel process(es) alive after cleanup: $($newExcel -join ', '). No process was terminated."
    }
    Write-Output 'Pre-commit VBE compile passed on the x64 bridge; temporary artifact was verified and will be removed.'
}
finally {
    if (Test-Path -LiteralPath $runRoot -PathType Container) {
        Remove-Item -LiteralPath $runRoot -Recurse -Force
    }
}
