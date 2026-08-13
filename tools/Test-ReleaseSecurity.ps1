[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$projectRoot = Split-Path -Parent $PSScriptRoot
$validatorPath = Join-Path $PSScriptRoot 'Validate-ReleaseSecurity.ps1'
$writerPath = Join-Path $PSScriptRoot 'Write-ReleaseChecksums.ps1'
$fixtureRoot = Join-Path $projectRoot ('.xlflow\release-security-test-{0}' -f [guid]::NewGuid().ToString('N'))

function Write-FixtureJson([string]$Path, $Value) {
    [IO.File]::WriteAllText($Path, (($Value | ConvertTo-Json -Depth 12 -Compress) + [Environment]::NewLine), [Text.UTF8Encoding]::new($false))
}

function Get-RepoRelativeSourcePath([IO.FileInfo]$File) {
    $rootWithSeparator = $projectRoot.TrimEnd('\', '/') + [IO.Path]::DirectorySeparatorChar
    return $File.FullName.Substring($rootWithSeparator.Length).Replace('\', '/')
}

function Get-ComponentRecord([string]$Name, [hashtable]$ByName) {
    if (-not $ByName.ContainsKey($Name)) {
        throw "Fixture source map is missing component '$Name'."
    }
    $file = $ByName[$Name]
    $type = if ($file.FullName -match '[\\/]src[\\/]classes[\\/]') {
        'class'
    }
    elseif ($file.FullName -match '[\\/]src[\\/]workbook[\\/]') {
        'document'
    }
    else {
        'standard'
    }
    [ordered]@{
        name = $Name
        type = $type
        source_path = Get-RepoRelativeSourcePath $file
        related_paths = @()
    }
}

function Initialize-FixtureRoot {
    $artifactDirectory = Join-Path $fixtureRoot 'build\Release'
    $toolsDirectory = Join-Path $fixtureRoot 'tools'
    [void](New-Item -ItemType Directory -Path $artifactDirectory -Force)
    [void](New-Item -ItemType Directory -Path $toolsDirectory -Force)
    [IO.File]::WriteAllBytes((Join-Path $artifactDirectory 'VBA-HTTP.xlsm'), [byte[]](0..255))
    [IO.File]::Copy((Join-Path $projectRoot 'tools\build-component-policy.json'), (Join-Path $toolsDirectory 'build-component-policy.json'), $true)

    $policy = Get-Content -LiteralPath (Join-Path $projectRoot 'tools\build-component-policy.json') -Raw | ConvertFrom-Json
    $sourceMap = @{}
    foreach ($file in @(Get-ChildItem -LiteralPath (Join-Path $projectRoot 'src') -Recurse -File | Where-Object { $_.Extension -in @('.bas', '.cls') })) {
        $sourceMap[$file.BaseName] = $file
    }
    $included = @($policy.included | ForEach-Object { Get-ComponentRecord ([string]$_) $sourceMap })
    $excluded = @($policy.excluded | ForEach-Object { Get-ComponentRecord ([string]$_) $sourceMap })
    $manifest = [ordered]@{
        schema_version = 1
        backend = 'excel'
        command = 'build'
        base = 'build/VBA-HTTP.xlsm'
        output = 'build/Release/VBA-HTTP.xlsm'
        included_components = $included
        excluded_components = $excluded
        publication = [ordered]@{ method = 'atomic_create'; replaced_existing = $false }
        validation = [ordered]@{
            components_applied = $included.Count
            source_applied = $true
            vbe_compile = 'passed'
            workbook_saved = $true
            workbook_closed = $true
            excel_cleanup = 'clean'
        }
    }
    $manifestPath = Join-Path $artifactDirectory 'VBA-HTTP.xlsm.build.json'
    Write-FixtureJson $manifestPath $manifest
    & powershell -NoProfile -ExecutionPolicy Bypass -File $writerPath `
        -ArtifactPath (Join-Path $artifactDirectory 'VBA-HTTP.xlsm') `
        -ManifestPath $manifestPath | Out-Null
    if ($LASTEXITCODE -ne 0) { throw 'Could not write the fixture checksum sidecar.' }
    return $manifestPath
}

function Invoke-Validator([string]$Root, [bool]$ShouldPass) {
    $outputPath = Join-Path $Root 'validator.stdout.txt'
    $errorPath = Join-Path $Root 'validator.stderr.txt'
    $process = $null
    try {
        $quotedValidator = '"{0}"' -f $validatorPath
        $quotedRoot = '"{0}"' -f $Root
        $arguments = @(
            '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $quotedValidator,
            '-RootPath', $quotedRoot,
            '-ArtifactPath', 'build/Release/VBA-HTTP.xlsm',
            '-PolicyPath', 'tools/build-component-policy.json',
            '-ReportPath', 'security-report.json'
        )
        $process = Start-Process -FilePath (Get-Command powershell).Source `
            -ArgumentList $arguments `
            -RedirectStandardOutput $outputPath `
            -RedirectStandardError $errorPath `
            -WindowStyle Hidden -Wait -PassThru
        $exitCode = $process.ExitCode
        $output = if (Test-Path -LiteralPath $outputPath -PathType Leaf) { @(Get-Content -LiteralPath $outputPath) } else { @() }
        $errorText = if (Test-Path -LiteralPath $errorPath -PathType Leaf) { Get-Content -LiteralPath $errorPath -Raw } else { '' }
        if ($ShouldPass -and $exitCode -ne 0) {
            throw "Expected release security validation to pass: $($output -join [Environment]::NewLine)$errorText"
        }
        if (-not $ShouldPass -and $exitCode -eq 0) {
            throw 'Expected tampered release security fixture to fail closed.'
        }
        return $output
    }
    finally {
        if ($null -ne $process) { $process.Dispose() }
        if (Test-Path -LiteralPath $outputPath -PathType Leaf) { Remove-Item -LiteralPath $outputPath -Force }
        if (Test-Path -LiteralPath $errorPath -PathType Leaf) { Remove-Item -LiteralPath $errorPath -Force }
    }
}

function Invoke-Case([string]$Name, [scriptblock]$Mutator, [bool]$ShouldPass) {
    $caseRoot = Join-Path $fixtureRoot $Name
    [void](New-Item -ItemType Directory -Path $caseRoot -Force)
    $oldRoot = $fixtureRoot
    try {
        $script:fixtureRoot = $caseRoot
        $manifestPath = Initialize-FixtureRoot
        if ($null -ne $Mutator) {
            $manifest = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json
            & $Mutator $manifest (Join-Path $caseRoot 'build\Release\VBA-HTTP.xlsm')
            Write-FixtureJson $manifestPath $manifest
        }
        Invoke-Validator $caseRoot $ShouldPass | Out-Null
    }
    finally {
        $script:fixtureRoot = $oldRoot
    }
}

try {
    [void](New-Item -ItemType Directory -Path $fixtureRoot -Force)
    Invoke-Case 'valid' $null $true
    Invoke-Case 'bad-base' {
        param($Manifest)
        $Manifest.base = 'build/Other.xlsm'
    } $false
    Invoke-Case 'included-test-path' {
        param($Manifest)
        $Manifest.included_components[0].source_path = 'src/modules/Tests/Injected.bas'
    } $false
    Invoke-Case 'missing-excluded' {
        param($Manifest)
        $Manifest.excluded_components = @($Manifest.excluded_components | Select-Object -Skip 1)
    } $false
    Invoke-Case 'compile-not-proven' {
        param($Manifest)
        $Manifest.validation.vbe_compile = 'failed'
    } $false
    Invoke-Case 'non-atomic-publication' {
        param($Manifest)
        $Manifest.publication.method = 'copy'
    } $false
    Invoke-Case 'related-path-escape' {
        param($Manifest)
        $Manifest.included_components[0].related_paths = @('docs/secret.txt')
    } $false
    Invoke-Case 'artifact-tamper' {
        param($Manifest, $Artifact)
        [void]$Manifest
        [IO.File]::WriteAllBytes($Artifact, [byte[]](255..0))
    } $false
    Write-Output 'Release security validator tests passed: valid manifest plus seven fail-closed tamper cases.'
}
finally {
    if (Test-Path -LiteralPath $fixtureRoot) {
        Remove-Item -LiteralPath $fixtureRoot -Recurse -Force
    }
}
