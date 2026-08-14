[CmdletBinding()]
param(
    [string]$OutputPath = "dist/VBA-HTTP-source.zip",
    [string]$PackageVersion = "source",
    [string]$RootPath = "",
    [string]$SourceRevision = ""
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
    $rootWithSeparator = $projectRoot.TrimEnd('\', '/') + [IO.Path]::DirectorySeparatorChar
    if (-not $resolved.StartsWith($rootWithSeparator, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Path is outside the project root: $Path"
    }
    return ConvertTo-RepoPath $resolved.Substring($rootWithSeparator.Length)
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

function Get-SourceRevision([string]$RequestedRevision) {
    if (-not [string]::IsNullOrWhiteSpace($RequestedRevision)) {
        return $RequestedRevision.Trim()
    }
    $revision = & git -C $projectRoot rev-parse HEAD 2>$null
    if ($LASTEXITCODE -eq 0 -and -not [string]::IsNullOrWhiteSpace(($revision | Out-String))) {
        $revisionText = ([string]($revision | Select-Object -First 1)).Trim()
        $status = & git -C $projectRoot status --porcelain 2>$null
        if ($LASTEXITCODE -eq 0 -and -not [string]::IsNullOrWhiteSpace(($status | Out-String))) {
            return "$revisionText-dirty"
        }
        return $revisionText
    }
    return "working-tree"
}

function Copy-PackageFile([string]$SourcePath, [string]$PackagePath, [string]$StagingRoot) {
    $destination = Join-Path $StagingRoot ($PackagePath.Replace('/', [IO.Path]::DirectorySeparatorChar))
    $destinationDirectory = [IO.Path]::GetDirectoryName($destination)
    if (-not (Test-Path -LiteralPath $destinationDirectory -PathType Container)) {
        [void](New-Item -ItemType Directory -Path $destinationDirectory -Force)
    }
    Copy-Item -LiteralPath $SourcePath -Destination $destination -Force
    return $destination
}

function Publish-Zip([string]$StagingRoot, [string]$OutputFile) {
    $outputDirectory = [IO.Path]::GetDirectoryName($OutputFile)
    if (-not (Test-Path -LiteralPath $outputDirectory -PathType Container)) {
        [void](New-Item -ItemType Directory -Path $outputDirectory -Force)
    }
    $temporaryZip = Join-Path $outputDirectory ('.{0}.{1}.tmp.zip' -f [IO.Path]::GetFileNameWithoutExtension($OutputFile), [guid]::NewGuid().ToString('N'))
    $backupZip = Join-Path $outputDirectory ('.{0}.{1}.bak.zip' -f [IO.Path]::GetFileNameWithoutExtension($OutputFile), [guid]::NewGuid().ToString('N'))
    try {
        Compress-Archive -Path (Join-Path $StagingRoot '*') -DestinationPath $temporaryZip -CompressionLevel Optimal
        if (Test-Path -LiteralPath $OutputFile -PathType Leaf) {
            [IO.File]::Replace($temporaryZip, $OutputFile, $backupZip)
            if (Test-Path -LiteralPath $backupZip -PathType Leaf) {
                Remove-Item -LiteralPath $backupZip -Force
            }
        }
        else {
            Move-Item -LiteralPath $temporaryZip -Destination $OutputFile
        }
    }
    finally {
        if (Test-Path -LiteralPath $temporaryZip -PathType Leaf) {
            Remove-Item -LiteralPath $temporaryZip -Force
        }
        if (Test-Path -LiteralPath $backupZip -PathType Leaf) {
            Remove-Item -LiteralPath $backupZip -Force
        }
    }
}

$policyPath = Join-Path $projectRoot "tools/build-component-policy.json"
$policy = Get-Content -LiteralPath $policyPath -Raw | ConvertFrom-Json
$outputFile = if ([IO.Path]::IsPathRooted($OutputPath)) {
    [IO.Path]::GetFullPath($OutputPath)
}
else {
    [IO.Path]::GetFullPath((Join-Path $projectRoot $OutputPath))
}

if ([string]::IsNullOrWhiteSpace($PackageVersion) -or $PackageVersion -notmatch '^[A-Za-z0-9][A-Za-z0-9._-]*$') {
    throw "PackageVersion must contain only letters, digits, '.', '_' or '-'."
}

$stagingRoot = Join-Path ([IO.Path]::GetTempPath()) ("vba-http-source-" + [guid]::NewGuid().ToString('N'))
[void](New-Item -ItemType Directory -Path $stagingRoot -Force)
try {
    $components = @()
    $seenNames = @{}
    $seenPackagePaths = @{}
    foreach ($sourceRoot in @('src/classes', 'src/modules')) {
        $resolvedRoot = Join-Path $projectRoot $sourceRoot
        $files = @(Get-ChildItem -LiteralPath $resolvedRoot -Recurse -File | Where-Object { $_.Extension -in @('.bas', '.cls') })
        foreach ($file in $files) {
            $sourcePath = Get-RepoRelativePath $file.FullName
            if ($sourcePath -match '(?i)(^|/)(Tests|Benchmarks|Xlflow|Dev)(/|$)') {
                continue
            }
            if (@($policy.included) -notcontains $file.BaseName) {
                continue
            }
            if (@('App', 'Main', 'Ui') -contains $file.BaseName) {
                continue
            }
            if ($seenNames.ContainsKey($file.BaseName)) {
                throw "Duplicate source component name: $($file.BaseName)"
            }

            $packageDirectory = if ($file.Extension -eq '.cls') { 'components/classes' } else { 'components/modules' }
            $packagePath = ConvertTo-RepoPath (Join-Path $packageDirectory $file.Name)
            if ($seenPackagePaths.ContainsKey($packagePath)) {
                throw "Duplicate package path: $packagePath"
            }
            [void](Copy-PackageFile $file.FullName $packagePath $stagingRoot)
            $seenNames[$file.BaseName] = $true
            $seenPackagePaths[$packagePath] = $true
            $components += [ordered]@{
                name = $file.BaseName
                type = if ($file.Extension -eq '.cls') { 'class' } else { 'standard' }
                source_path = $sourcePath
                package_path = $packagePath
                sha256 = Get-Sha256Hex (Join-Path $stagingRoot ($packagePath.Replace('/', [IO.Path]::DirectorySeparatorChar)))
            }
        }
    }

    if ($components.Count -eq 0) {
        throw "No production VBA components were selected for the source package."
    }

    foreach ($fileName in @('README.md', 'CONTRIBUTING.md', 'CHANGELOG.md', 'docs/API.md', 'docs/RELEASE_CHECKLIST.md', 'docs/BENCHMARKS_BASELINE.md', 'benchmarks/schema/source-package-manifest.schema.json', 'LICENSE', 'THIRD_PARTY_NOTICES.md')) {
        [void](Copy-PackageFile (Join-Path $projectRoot $fileName) $fileName $stagingRoot)
    }
    foreach ($specFile in @(Get-ChildItem -LiteralPath (Join-Path $projectRoot 'docs/specs') -Filter '*.md' -File)) {
        [void](Copy-PackageFile $specFile.FullName (ConvertTo-RepoPath (Join-Path 'docs/specs' $specFile.Name)) $stagingRoot)
    }
    foreach ($adrFileName in @('ADR-0030-32-bit-office-support-boundary.md', 'ADR-0035-http3-support-boundary.md', 'ADR-0037-vba-web-style-source-package.md')) {
        [void](Copy-PackageFile (Join-Path $projectRoot (Join-Path 'docs/adr' $adrFileName)) (ConvertTo-RepoPath (Join-Path 'docs/adr' $adrFileName)) $stagingRoot)
    }
    foreach ($fileName in @('Install-VBAHttp.ps1', 'Uninstall-VBAHttp.ps1', 'Validate-SourcePackage.ps1')) {
        [void](Copy-PackageFile (Join-Path $PSScriptRoot $fileName) $fileName $stagingRoot)
    }

    $fileRecords = @()
    foreach ($file in @(Get-ChildItem -LiteralPath $stagingRoot -Recurse -File | Where-Object { $_.Name -ne 'manifest.json' })) {
        $packageRelativePath = $file.FullName.Substring($stagingRoot.Length).TrimStart([char[]]@('\', '/'))
        $fileRecords += [ordered]@{
            path = ConvertTo-RepoPath $packageRelativePath
            sha256 = Get-Sha256Hex $file.FullName
        }
    }
    $fileRecords = @($fileRecords | Sort-Object path)
    $manifest = [ordered]@{
        schema_version = 1
        package_id = 'VBA-HTTP'
        package_version = $PackageVersion
        source_revision = Get-SourceRevision $SourceRevision
        target = 'windows-x64-office'
        install = [ordered]@{
            script = 'Install-VBAHttp.ps1'
            uninstall_script = 'Uninstall-VBAHttp.ps1'
            requires_closed_workbook = $true
            requires_vbproject_access = $true
        }
        components = @($components | Sort-Object name)
        files = $fileRecords
    }
    $manifestText = ($manifest | ConvertTo-Json -Depth 8) + [Environment]::NewLine
    [IO.File]::WriteAllText((Join-Path $stagingRoot 'manifest.json'), $manifestText, [Text.UTF8Encoding]::new($false))

    Publish-Zip $stagingRoot $outputFile
    Write-Output "VBA-HTTP source package created: $outputFile ($($components.Count) components)."
}
finally {
    if (Test-Path -LiteralPath $stagingRoot) {
        Remove-Item -LiteralPath $stagingRoot -Recurse -Force
    }
}
