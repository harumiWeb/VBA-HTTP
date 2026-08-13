[CmdletBinding()]
param(
    [string]$Url = [Environment]::GetEnvironmentVariable("VBA_HTTP_PROTOCOL_HOST_URL", "Process"),
    [ValidateSet("HTTP/2", "HTTP/3")]
    [string]$ExpectedProtocol = [Environment]::GetEnvironmentVariable("VBA_HTTP_PROTOCOL_EXPECTED", "Process"),
    [string]$ArtifactPath = "build/Release/VBA-HTTP.xlsm",
    [string]$OutputPath = ""
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$projectRoot = [IO.Path]::GetFullPath((Split-Path -Parent $PSScriptRoot))

function Resolve-ProjectPath([string]$Path) {
    if ([string]::IsNullOrWhiteSpace($Path)) { throw "A path cannot be empty." }
    if ([IO.Path]::IsPathRooted($Path)) { return [IO.Path]::GetFullPath($Path) }
    return [IO.Path]::GetFullPath((Join-Path $projectRoot $Path))
}

function Invoke-JsonCommand([string]$FilePath, [string[]]$Arguments) {
    $jsonLines = & $FilePath @Arguments
    if ($LASTEXITCODE -ne 0) { throw "$FilePath failed with exit code $LASTEXITCODE." }
    return (($jsonLines | Out-String) | ConvertFrom-Json)
}

function Get-Sha256Hex([string]$Path) {
    $sha256 = [Security.Cryptography.SHA256]::Create()
    $stream = $null
    try {
        $stream = [IO.File]::OpenRead($Path)
        return ([BitConverter]::ToString($sha256.ComputeHash($stream))).Replace("-", "").ToLowerInvariant()
    }
    finally {
        if ($null -ne $stream) { $stream.Dispose() }
        $sha256.Dispose()
    }
}

function Publish-JsonAtomically([string]$Path, $Document) {
    $directory = [IO.Path]::GetDirectoryName($Path)
    if (-not (Test-Path -LiteralPath $directory -PathType Container)) {
        [void](New-Item -ItemType Directory -Path $directory -Force)
    }
    $staging = Join-Path $directory ("." + [IO.Path]::GetFileName($Path) + "." + [guid]::NewGuid().ToString("N") + ".tmp")
    $json = ($Document | ConvertTo-Json -Depth 12) + [Environment]::NewLine
    try {
        [IO.File]::WriteAllText($staging, $json, [Text.UTF8Encoding]::new($false))
        Move-Item -LiteralPath $staging -Destination $Path -Force
    }
    finally {
        if (Test-Path -LiteralPath $staging -PathType Leaf) { Remove-Item -LiteralPath $staging -Force }
    }
}

function Get-ExcelProcessIds {
    $processes = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue)
    return @($processes | ForEach-Object { [int]$_.Id })
}

function Get-RelativeProjectPath([string]$Path) {
    $fullPath = [IO.Path]::GetFullPath($Path)
    $root = $projectRoot.TrimEnd([char[]]@([IO.Path]::DirectorySeparatorChar, [IO.Path]::AltDirectorySeparatorChar))
    if ([string]::Equals($fullPath, $root, [StringComparison]::OrdinalIgnoreCase)) {
        return "."
    }
    $rootPrefix = $root + [IO.Path]::DirectorySeparatorChar
    if (-not $fullPath.StartsWith($rootPrefix, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Path must remain inside the project root: $Path"
    }
    return $fullPath.Substring($rootPrefix.Length).Replace("\", "/")
}

function Assert-CleanWorktreeExcept([string]$AllowedPath) {
    $relativeAllowed = Get-RelativeProjectPath $AllowedPath
    $entries = @(& git status --porcelain 2>$null)
    foreach ($entry in $entries) {
        if ([string]::IsNullOrWhiteSpace($entry)) { continue }
        $path = $entry.Substring(3).Trim().Replace("\", "/")
        if ($path -ne $relativeAllowed) {
            throw "Protocol host evidence requires a clean source worktree; unexpected change: $path"
        }
    }
}

if ([string]::IsNullOrWhiteSpace($Url)) { throw "Set -Url or VBA_HTTP_PROTOCOL_HOST_URL to a trusted HTTPS endpoint." }
if ([string]::IsNullOrWhiteSpace($ExpectedProtocol)) { throw "Set -ExpectedProtocol or VBA_HTTP_PROTOCOL_EXPECTED to HTTP/2 or HTTP/3." }
$ExpectedProtocol = $ExpectedProtocol.ToUpperInvariant()

$targetUri = $null
if (-not [Uri]::TryCreate($Url, [UriKind]::Absolute, [ref]$targetUri) -or
    $targetUri.Scheme -ne "https" -or
    [string]::IsNullOrWhiteSpace($targetUri.Host) -or
    -not [string]::IsNullOrWhiteSpace($targetUri.UserInfo) -or
    -not [string]::IsNullOrWhiteSpace($targetUri.Fragment)) {
    throw "Protocol host validation requires an absolute HTTPS URL without user-info or fragments."
}

$resolvedArtifact = Resolve-ProjectPath $ArtifactPath
$resolvedManifest = "$resolvedArtifact.build.json"
$resolvedChecksum = "$resolvedArtifact.checksum.json"
$resolvedOutput = if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    Join-Path $projectRoot ("benchmarks\results\protocol-host-{0}.json" -f $ExpectedProtocol.ToLowerInvariant().Replace("/", ""))
}
else { Resolve-ProjectPath $OutputPath }
Assert-CleanWorktreeExcept $resolvedOutput
foreach ($requiredPath in @($resolvedArtifact, $resolvedManifest, $resolvedChecksum)) {
    if (-not (Test-Path -LiteralPath $requiredPath -PathType Leaf)) { throw "Required release input is missing: $requiredPath" }
}

& (Join-Path $PSScriptRoot "Validate-ReleaseSecurity.ps1") `
    -ArtifactPath $resolvedArtifact `
    -ReportPath (Join-Path $projectRoot ".xlflow\release-security\release-security.json")
& (Join-Path $PSScriptRoot "Verify-ReleaseChecksums.ps1") -ArtifactPath $resolvedArtifact -ManifestPath $resolvedManifest
Write-Verbose "Release artifact and checksum gates passed."

$manifest = Get-Content -LiteralPath $resolvedManifest -Raw | ConvertFrom-Json
if ($manifest.validation.vbe_compile -ne "passed" -or
    $manifest.validation.source_applied -ne $true -or
    $manifest.validation.workbook_saved -ne $true -or
    $manifest.validation.workbook_closed -ne $true -or
    $manifest.validation.excel_cleanup -ne "clean") {
    throw "Release manifest does not prove a clean compiled artifact."
}
Write-Verbose "Running xlflow doctor."
$doctor = Invoke-JsonCommand "xlflow" @("doctor", "--json")
Write-Verbose "xlflow doctor completed."
$architecture = [string]$doctor.bridge.architecture
if ($architecture -notin @("X86", "X64")) { throw "xlflow did not report an X86 or X64 bridge." }
$sourceRevision = ((& git rev-parse HEAD 2>$null) | Out-String).Trim()
if ($sourceRevision -notmatch '^[0-9a-fA-F]{40,64}$') { throw "Could not determine the source revision." }

$baselineExcelIds = @(Get-ExcelProcessIds)
$ownedExcelIds = @()
$ownsExcel = $false
$excel = $null
$workbooks = $null
$consumerWorkbook = $null
$harnessWorkbook = $null
$harnessComponent = $null
$observedProtocol = $null
$office = $null
try {
    Write-Verbose "Starting an owned Excel automation instance."
    $excel = New-Object -ComObject Excel.Application
    $afterExcelIds = @(Get-ExcelProcessIds)
    $ownedExcelIds = @($afterExcelIds | Where-Object { $baselineExcelIds -notcontains $_ })
    if ($ownedExcelIds.Count -eq 0) {
        throw "Could not prove ownership of a new Excel process; refusing to touch an existing instance."
    }
    $ownsExcel = $true
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.AutomationSecurity = 1
    Write-Verbose ("Owned Excel PID(s): " + ($ownedExcelIds -join ","))
    $office = [ordered]@{
        version = [string]$excel.Version
        build = [string]$excel.Build
        operating_system = [string]$excel.OperatingSystem
    }
    $workbooks = $excel.Workbooks
    Write-Verbose "Opening the release artifact read-only."
    $consumerWorkbook = $workbooks.Open($resolvedArtifact, 0, $true)
    Write-Verbose "Release artifact opened."
    $harnessWorkbook = $workbooks.Add()
    $harnessComponent = $harnessWorkbook.VBProject.VBComponents.Import((Join-Path $PSScriptRoot "consumer\ReleaseBatchSmoke.bas"))
    if ($harnessComponent.Name -ne "ReleaseBatchSmoke") { throw "Protocol host consumer harness import failed." }
    Write-Verbose "Consumer harness imported."
    $macro = "'$($harnessWorkbook.Name)'!ReleaseBatchSmoke.RunProtocolHostSmoke"
    Write-Verbose "Calling the required protocol consumer smoke."
    $observedProtocol = [string]$excel.Run($macro, $consumerWorkbook.Name, $Url, $ExpectedProtocol)
    Write-Verbose ("Protocol consumer returned: " + $observedProtocol)
    if ($observedProtocol -ne $ExpectedProtocol) { throw "Protocol host consumer returned an unexpected protocol." }
}
finally {
    if ($null -ne $harnessComponent) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($harnessComponent) }
    if ($null -ne $harnessWorkbook) {
        try { $harnessWorkbook.Close($false) } catch { Write-Warning "Could not close protocol harness workbook: $_" }
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($harnessWorkbook)
    }
    if ($null -ne $consumerWorkbook) {
        try { $consumerWorkbook.Close($false) } catch { Write-Warning "Could not close protocol consumer workbook: $_" }
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($consumerWorkbook)
    }
    if ($null -ne $workbooks) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($workbooks) }
    if ($null -ne $excel) {
        if ($ownsExcel) { try { $excel.Quit() } catch { Write-Warning "Could not quit owned protocol validation Excel: $_" } }
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
    }
    foreach ($processId in $ownedExcelIds) {
        $process = Get-Process -Id $processId -ErrorAction SilentlyContinue
        if ($null -ne $process -and -not $process.HasExited) {
            Stop-Process -Id $processId -Force
        }
    }
}

$record = [ordered]@{
    schema_version = 1
    benchmark = "protocol-host-validation"
    status = "passed"
    run_utc = [DateTime]::UtcNow.ToString("o")
    source_revision = $sourceRevision
    requested_protocol = $ExpectedProtocol
    observed_protocol = $observedProtocol
    mode = "required"
    external_network = $true
    target = [ordered]@{
        scheme = "https"
        host = $targetUri.IdnHost
        port = if ($targetUri.IsDefaultPort) { 443 } else { $targetUri.Port }
    }
    bridge = [ordered]@{
        name = [string]$doctor.bridge.name
        version = [string]$doctor.bridge.version
        runtime = [string]$doctor.bridge.runtime
        architecture = $architecture
    }
    office = $office
    environment = [ordered]@{
        windows = [Environment]::OSVersion.VersionString
        winhttp = [ordered]@{
            enabled_protocol_option = 133
            used_protocol_option = 134
            required_protocol_option = 145
            observed_protocol = $observedProtocol
        }
    }
    artifact = [ordered]@{
        name = [IO.Path]::GetFileName($resolvedArtifact)
        sha256 = Get-Sha256Hex $resolvedArtifact
        manifest_sha256 = Get-Sha256Hex $resolvedManifest
    }
    build = [ordered]@{
        vbe_compile = [string]$manifest.validation.vbe_compile
        source_applied = [bool]$manifest.validation.source_applied
        workbook_saved = [bool]$manifest.validation.workbook_saved
        workbook_closed = [bool]$manifest.validation.workbook_closed
        excel_cleanup = [string]$manifest.validation.excel_cleanup
    }
}
Publish-JsonAtomically $resolvedOutput $record
& (Join-Path $PSScriptRoot "Validate-ProtocolHostEvidence.ps1") -Path $resolvedOutput -ExpectedProtocol $ExpectedProtocol
Write-Output "Protocol host validation passed: $ExpectedProtocol negotiated on $($targetUri.IdnHost) ($architecture). Evidence: $resolvedOutput"
