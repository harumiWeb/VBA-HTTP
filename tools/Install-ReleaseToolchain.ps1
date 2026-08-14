[CmdletBinding()]
param(
    [string]$RootPath = "",
    [string]$InstallPath = ""
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$projectRoot = if ([string]::IsNullOrWhiteSpace($RootPath)) {
    [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
}
else {
    [IO.Path]::GetFullPath($RootPath)
}
$lockPath = Join-Path $projectRoot 'tools/release-toolchain.json'
if (-not (Test-Path -LiteralPath $lockPath -PathType Leaf)) {
    throw "Release toolchain lock is missing: $lockPath"
}
$lock = Get-Content -LiteralPath $lockPath -Raw | ConvertFrom-Json
$installRoot = if ([string]::IsNullOrWhiteSpace($InstallPath)) {
    Join-Path $projectRoot '.release-tools'
}
elseif ([IO.Path]::IsPathRooted($InstallPath)) {
    [IO.Path]::GetFullPath($InstallPath)
}
else {
    [IO.Path]::GetFullPath((Join-Path $projectRoot $InstallPath))
}

function Get-Sha256Hex([string]$Path) {
    $sha = [Security.Cryptography.SHA256]::Create()
    $stream = [IO.File]::OpenRead($Path)
    try {
        return ([BitConverter]::ToString($sha.ComputeHash($stream))).Replace('-', '').ToLowerInvariant()
    }
    finally {
        $stream.Dispose()
        $sha.Dispose()
    }
}

function Get-RequiredProperty($Object, [string]$Name) {
    if ($null -eq $Object) { throw "Toolchain record is null while reading '$Name'." }
    $property = $Object.PSObject.Properties[$Name]
    if ($null -eq $property) { throw "Toolchain lock is missing '$Name'." }
    return $property.Value
}

function Get-CommandPath([string]$Name) {
    $command = Get-Command $Name -ErrorAction SilentlyContinue
    if ($null -eq $command) { return '' }
    return [string]$command.Source
}

if ([int](Get-RequiredProperty $lock 'schema_version') -ne 1) {
    throw 'Unsupported release toolchain lock schema.'
}
$xlflow = Get-RequiredProperty $lock 'xlflow'
$version = [string](Get-RequiredProperty $xlflow 'version')
$asset = [string](Get-RequiredProperty $xlflow 'asset')
$expectedHash = [string](Get-RequiredProperty $xlflow 'sha256')
$repository = [string](Get-RequiredProperty $xlflow 'repository')
if ($version -notmatch '^v[0-9]+\.[0-9]+\.[0-9]+$' -or
    $asset -notmatch '^xlflow_windows_x86_64\.zip$' -or
    $expectedHash -notmatch '^[0-9a-f]{64}$' -or
    $repository -notmatch '^[A-Za-z0-9_.-]+/[A-Za-z0-9_.-]+$') {
    throw 'Release toolchain xlflow record is invalid.'
}

if (-not (Test-Path -LiteralPath $installRoot -PathType Container)) {
    [void](New-Item -ItemType Directory -Path $installRoot -Force)
}
$archivePath = Join-Path $installRoot $asset
$executablePath = Join-Path $installRoot 'xlflow.exe'
$downloadUrl = "https://github.com/$repository/releases/download/$version/$asset"
if (-not (Test-Path -LiteralPath $archivePath -PathType Leaf) -or
    (Get-Sha256Hex $archivePath) -ne $expectedHash) {
    $temporaryArchive = Join-Path $installRoot ('.{0}.{1}.download' -f $asset, [guid]::NewGuid().ToString('N'))
    try {
        Invoke-WebRequest -UseBasicParsing -Uri $downloadUrl -OutFile $temporaryArchive
        if ((Get-Sha256Hex $temporaryArchive) -ne $expectedHash) {
            throw "xlflow archive hash mismatch for $version."
        }
        Move-Item -LiteralPath $temporaryArchive -Destination $archivePath -Force
    }
    finally {
        if (Test-Path -LiteralPath $temporaryArchive -PathType Leaf) {
            Remove-Item -LiteralPath $temporaryArchive -Force
        }
    }

}

$extractRoot = Join-Path $installRoot ('.extract-' + [guid]::NewGuid().ToString('N'))
try {
    Expand-Archive -LiteralPath $archivePath -DestinationPath $extractRoot -Force
    $candidate = Get-ChildItem -LiteralPath $extractRoot -Recurse -Filter 'xlflow.exe' -File | Select-Object -First 1
    if ($null -eq $candidate) { throw 'The xlflow archive did not contain xlflow.exe.' }
    Copy-Item -LiteralPath $candidate.FullName -Destination $executablePath -Force
}
finally {
    if (Test-Path -LiteralPath $extractRoot -PathType Container) {
        Remove-Item -LiteralPath $extractRoot -Recurse -Force
    }
}

$current = Get-CommandPath 'xlflow'
if ($current -ne $executablePath) {
    $env:PATH = "$installRoot$([IO.Path]::PathSeparator)$env:PATH"
}
$versionOutput = & $executablePath version --json
if ($LASTEXITCODE -ne 0) { throw 'Installed xlflow failed its version probe.' }
$versionJson = ($versionOutput | Out-String | ConvertFrom-Json)
if ([string]$versionJson.version.version -ne $version.TrimStart('v')) {
    throw "Installed xlflow version does not match the lock: expected $version."
}

$task = Get-RequiredProperty $lock 'task'
$go = Get-RequiredProperty $lock 'go'
$psa = Get-RequiredProperty $lock 'psscriptanalyzer'
Write-Output ([ordered]@{
        xlflow = [ordered]@{ version = $version; path = $executablePath; sha256 = $expectedHash }
        task = [string](Get-RequiredProperty $task 'version')
        go = [string](Get-RequiredProperty $go 'version')
        psscriptanalyzer = [string](Get-RequiredProperty $psa 'version')
    } | ConvertTo-Json -Compress)
