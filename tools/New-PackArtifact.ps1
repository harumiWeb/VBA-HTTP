[CmdletBinding()]
param(
    [string]$ReleaseTag = "local",
    [string]$SourceRevision = "",
    [string]$TemplatePath = "build/VBA-HTTP.xlsm",
    [string]$OutputPath = "dist/VBA-HTTP-pack.xlsm",
    [string]$ManifestPath = "",
    [string]$RootPath = ""
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$projectRoot = if ([string]::IsNullOrWhiteSpace($RootPath)) {
    [IO.Path]::GetFullPath((Join-Path $PSScriptRoot ".."))
}
else {
    [IO.Path]::GetFullPath($RootPath)
}

function ConvertTo-RepoPath([string]$Path) {
    return $Path.Replace('\', '/')
}

function Get-RepoRelativePath([string]$Path) {
    $resolved = [IO.Path]::GetFullPath($Path)
    $root = $projectRoot.TrimEnd('\', '/') + [IO.Path]::DirectorySeparatorChar
    if (-not $resolved.StartsWith($root, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Path is outside the project root: $Path"
    }
    return ConvertTo-RepoPath $resolved.Substring($root.Length)
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

function Publish-FileAtomic([string]$SourcePath, [string]$DestinationPath) {
    $destinationDirectory = [IO.Path]::GetDirectoryName($DestinationPath)
    if (-not (Test-Path -LiteralPath $destinationDirectory -PathType Container)) {
        [void](New-Item -ItemType Directory -Path $destinationDirectory -Force)
    }
    $stagingPath = Join-Path $destinationDirectory ('.{0}.{1}.tmp' -f [IO.Path]::GetFileName($DestinationPath), [guid]::NewGuid().ToString('N'))
    $backupPath = Join-Path $destinationDirectory ('.{0}.{1}.bak' -f [IO.Path]::GetFileName($DestinationPath), [guid]::NewGuid().ToString('N'))
    try {
        Copy-Item -LiteralPath $SourcePath -Destination $stagingPath -Force
        if (Test-Path -LiteralPath $DestinationPath -PathType Leaf) {
            [IO.File]::Replace($stagingPath, $DestinationPath, $backupPath)
            if (Test-Path -LiteralPath $backupPath -PathType Leaf) {
                Remove-Item -LiteralPath $backupPath -Force
            }
        }
        else {
            Move-Item -LiteralPath $stagingPath -Destination $DestinationPath
        }
    }
    finally {
        foreach ($path in @($stagingPath, $backupPath)) {
            if (Test-Path -LiteralPath $path -PathType Leaf) {
                Remove-Item -LiteralPath $path -Force
            }
        }
    }
}

function Get-SourceRevision([string]$RequestedRevision) {
    if (-not [string]::IsNullOrWhiteSpace($RequestedRevision)) {
        return $RequestedRevision.Trim()
    }
    $gitOutput = & git -C $projectRoot rev-parse HEAD 2>$null
    $gitExitCode = $LASTEXITCODE
    $revision = ($gitOutput | Select-Object -First 1)
    if ($gitExitCode -ne 0 -or [string]::IsNullOrWhiteSpace([string]$revision)) {
        return 'working-tree'
    }
    $revisionText = ([string]$revision).Trim()
    $status = & git -C $projectRoot status --porcelain 2>$null
    if ($LASTEXITCODE -eq 0 -and -not [string]::IsNullOrWhiteSpace(($status | Out-String))) {
        return "$revisionText-dirty"
    }
    return $revisionText
}

function Invoke-JsonCommand([string]$FilePath, [string[]]$Arguments, [string]$WorkingDirectory) {
    Push-Location $WorkingDirectory
    try {
        $lines = & $FilePath @Arguments
        if ($LASTEXITCODE -ne 0) {
            throw "$FilePath failed with exit code $LASTEXITCODE."
        }
        $text = ($lines | Out-String).Trim()
        if ([string]::IsNullOrWhiteSpace($text)) {
            throw "$FilePath returned no JSON output."
        }
        return $text | ConvertFrom-Json
    }
    finally {
        Pop-Location
    }
}

function Get-XlflowVersion {
    $version = Invoke-JsonCommand 'xlflow' @('version', '--json') $projectRoot
    if ($null -eq $version.version -or [string]::IsNullOrWhiteSpace([string]$version.version.version)) {
        throw 'xlflow version output did not contain a version.'
    }
    return [string]$version.version.version
}

if ($ReleaseTag -ne 'local') {
    if ([string]::IsNullOrWhiteSpace($SourceRevision)) {
        throw 'A tagged pack artifact requires an explicit source commit SHA.'
    }
    if ($SourceRevision -notmatch '^[0-9a-fA-F]{40,64}$') {
        throw 'A tagged pack artifact requires a full hexadecimal source commit SHA.'
    }
    & (Join-Path $PSScriptRoot 'Validate-ReleaseTag.ps1') -Tag $ReleaseTag | Out-Null
}

$policyPath = Join-Path $projectRoot 'tools/build-component-policy.json'
$policy = Get-Content -LiteralPath $policyPath -Raw | ConvertFrom-Json
$resolvedTemplate = if ([IO.Path]::IsPathRooted($TemplatePath)) { [IO.Path]::GetFullPath($TemplatePath) } else { [IO.Path]::GetFullPath((Join-Path $projectRoot $TemplatePath)) }
$resolvedOutput = if ([IO.Path]::IsPathRooted($OutputPath)) { [IO.Path]::GetFullPath($OutputPath) } else { [IO.Path]::GetFullPath((Join-Path $projectRoot $OutputPath)) }
$resolvedManifest = if ([string]::IsNullOrWhiteSpace($ManifestPath)) { "$resolvedOutput.pack.json" } elseif ([IO.Path]::IsPathRooted($ManifestPath)) { [IO.Path]::GetFullPath($ManifestPath) } else { [IO.Path]::GetFullPath((Join-Path $projectRoot $ManifestPath)) }

foreach ($path in @($resolvedTemplate, $policyPath)) {
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
        throw "Pack input does not exist: $path"
    }
}

$stagingRoot = Join-Path ([IO.Path]::GetTempPath()) ('vba-http-pack-' + [guid]::NewGuid().ToString('N'))
$stagingTemplate = Join-Path $stagingRoot 'template.xlsm'
$stagingOutput = Join-Path $stagingRoot 'packed.xlsm'
$componentRecords = @()
$excludedRecords = @()
try {
    foreach ($directory in @('src/classes', 'src/modules', 'src/workbook', 'src/forms')) {
        [void](New-Item -ItemType Directory -Path (Join-Path $stagingRoot $directory) -Force)
    }
    Copy-Item -LiteralPath $resolvedTemplate -Destination $stagingTemplate -Force
    $config = @"
[project]
name = "VBA-HTTP"

[src]
modules = "src/modules"
classes = "src/classes"
forms = "src/forms"
workbook = "src/workbook"
"@
    [IO.File]::WriteAllText((Join-Path $stagingRoot 'xlflow.toml'), $config, [Text.UTF8Encoding]::new($false))

    $allFiles = @()
    foreach ($sourceRoot in @('src/classes', 'src/modules', 'src/workbook', 'src/forms')) {
        $root = Join-Path $projectRoot $sourceRoot
        if (Test-Path -LiteralPath $root -PathType Container) {
            $allFiles += @(Get-ChildItem -LiteralPath $root -Recurse -File | Where-Object { $_.Extension.ToLowerInvariant() -in @('.bas', '.cls', '.frm') })
        }
    }
    $filesByName = @{}
    foreach ($file in $allFiles) {
        if ($filesByName.ContainsKey($file.BaseName)) {
            throw "Duplicate VBA source component name: $($file.BaseName)"
        }
        $filesByName[$file.BaseName] = $file
    }

    foreach ($name in @($policy.included)) {
        if (-not $filesByName.ContainsKey([string]$name)) {
            throw "Included component is missing from source: $name"
        }
        $file = $filesByName[[string]$name]
        $relative = Get-RepoRelativePath $file.FullName
        $destination = Join-Path $stagingRoot ($relative.Replace('/', [IO.Path]::DirectorySeparatorChar))
        [void](New-Item -ItemType Directory -Path ([IO.Path]::GetDirectoryName($destination)) -Force)
        Copy-Item -LiteralPath $file.FullName -Destination $destination -Force
        $type = if ($relative.StartsWith('src/classes/', [StringComparison]::OrdinalIgnoreCase)) { 'class' } elseif ($relative.StartsWith('src/workbook/', [StringComparison]::OrdinalIgnoreCase)) { 'document' } elseif ($relative.StartsWith('src/forms/', [StringComparison]::OrdinalIgnoreCase)) { 'form' } else { 'standard' }
        $componentRecords += [ordered]@{ name = [string]$name; type = $type; source_path = $relative; sha256 = Get-Sha256Hex $file.FullName }
    }
    foreach ($name in @($policy.excluded)) {
        if (-not $filesByName.ContainsKey([string]$name)) {
            throw "Excluded component is missing from source: $name"
        }
        $file = $filesByName[[string]$name]
        $relative = Get-RepoRelativePath $file.FullName
        $type = if ($relative.StartsWith('src/classes/', [StringComparison]::OrdinalIgnoreCase)) { 'class' } elseif ($relative.StartsWith('src/workbook/', [StringComparison]::OrdinalIgnoreCase)) { 'document' } elseif ($relative.StartsWith('src/forms/', [StringComparison]::OrdinalIgnoreCase)) { 'form' } else { 'standard' }
        $excludedRecords += [ordered]@{ name = [string]$name; type = $type; source_path = $relative }
    }

    $pack = Invoke-JsonCommand 'xlflow' @('pack', '--experimental', '--template', $stagingTemplate, '--out', $stagingOutput, '--json') $stagingRoot
    if ($pack.status -ne 'ok' -or $pack.pack.backend -ne 'pure-go' -or $pack.pack.experimental -ne $true -or $pack.pack.vbe_validation -ne 'not_performed') {
        throw 'xlflow pack did not return the required pure-Go experimental provenance.'
    }
    if (-not (Test-Path -LiteralPath $stagingOutput -PathType Leaf)) {
        throw 'xlflow pack did not produce an artifact.'
    }

    $expectedCounts = [ordered]@{
        class = @($componentRecords | Where-Object { $_.type -eq 'class' }).Count
        standard = @($componentRecords | Where-Object { $_.type -eq 'standard' }).Count
        document = @($componentRecords | Where-Object { $_.type -eq 'document' }).Count
        form = @($componentRecords | Where-Object { $_.type -eq 'form' }).Count
    }
    foreach ($key in @('class', 'standard', 'document', 'form')) {
        if ([int]$pack.pack.modules.$key -ne [int]$expectedCounts[$key]) {
            throw "xlflow pack module count mismatch for ${key}: expected $($expectedCounts[$key]), actual $($pack.pack.modules.$key)."
        }
    }

    Publish-FileAtomic $stagingOutput $resolvedOutput
    $manifest = [ordered]@{
        schema_version = 1
        package_id = 'VBA-HTTP'
        release_tag = $ReleaseTag
        source_revision = Get-SourceRevision $SourceRevision
        target = 'windows-x64-office'
        artifact = [IO.Path]::GetFileName($resolvedOutput)
        template = 'build/VBA-HTTP.xlsm'
        template_sha256 = Get-Sha256Hex $resolvedTemplate
        xlflow_version = Get-XlflowVersion
        pack_backend = 'pure-go'
        experimental = $true
        vbe_validation = 'not_performed'
        included_components = @($componentRecords | Sort-Object name)
        excluded_components = @($excludedRecords | Sort-Object name)
        module_counts = $expectedCounts
    }
    $manifestDirectory = [IO.Path]::GetDirectoryName($resolvedManifest)
    if (-not (Test-Path -LiteralPath $manifestDirectory -PathType Container)) {
        [void](New-Item -ItemType Directory -Path $manifestDirectory -Force)
    }
    $manifestTemp = Join-Path $manifestDirectory ('.{0}.{1}.tmp' -f [IO.Path]::GetFileName($resolvedManifest), [guid]::NewGuid().ToString('N'))
    $manifestText = ($manifest | ConvertTo-Json -Depth 8) + [Environment]::NewLine
    [IO.File]::WriteAllText($manifestTemp, $manifestText, [Text.UTF8Encoding]::new($false))
    try {
        Publish-FileAtomic $manifestTemp $resolvedManifest
    }
    finally {
        if (Test-Path -LiteralPath $manifestTemp -PathType Leaf) { Remove-Item -LiteralPath $manifestTemp -Force }
    }
    Write-Output "Pack artifact created: $resolvedOutput ($($componentRecords.Count) included components, VBE validation not performed)."
}
finally {
    if (Test-Path -LiteralPath $stagingRoot -PathType Container) {
        Remove-Item -LiteralPath $stagingRoot -Recurse -Force
    }
}
