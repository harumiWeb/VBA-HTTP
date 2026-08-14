[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$root = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$required = @(
    'README.md',
    'LICENSE',
    'THIRD_PARTY_NOTICES.md',
    'CONTRIBUTING.md',
    'docs/API.md',
    'docs/guides/README.md',
    'docs/guides/getting-started.md',
    'docs/guides/api-reference.md',
    'docs/guides/requests-and-responses.md',
    'docs/guides/reliability-and-batches.md',
    'docs/guides/streaming.md',
    'docs/guides/transport-capabilities.md',
    'docs/guides/security-and-state.md',
    'docs/guides/distribution.md',
    'docs/guides/examples.md',
    'docs/guides/compatibility.md',
    'docs/RELEASE_CHECKLIST.md',
    'docs/specs/distribution.md',
    'docs/specs/github-release.md',
    'docs/specs/source-package.md',
    'docs/specs/licensing.md',
    'docs/specs/powershell-quality.md',
    'docs/specs/development-and-release-workflow.md',
    'docs/specs/github-ci.md',
    'docs/adr/ADR-0021-xlam-distribution-target.md',
    'docs/adr/ADR-0022-com-timeout-failure-cleanup.md',
    'docs/adr/ADR-0030-32-bit-office-support-boundary.md',
    'docs/adr/ADR-0039-32-bit-office-unverified-community-validation.md',
    'docs/adr/ADR-0035-http3-support-boundary.md',
    'docs/adr/ADR-0036-mit-license-and-distribution-boundary.md',
    'docs/adr/ADR-0037-vba-web-style-source-package.md',
    'docs/adr/ADR-0038-github-tag-release-and-pack-provenance.md',
    'docs/adr/ADR-0040-pull-request-and-scheduled-excel-free-ci.md',
    'tools/Test-LicenseContract.ps1',
    'tools/Invoke-PSScriptAnalyzer.ps1',
    'tools/New-SourcePackage.ps1',
    'tools/Validate-SourcePackage.ps1',
    'tools/Validate-SourceArchive.ps1',
    'tools/Install-VBAHttp.ps1',
    'tools/Uninstall-VBAHttp.ps1',
    'tools/Test-SourcePackage.ps1',
    'tools/Validate-ReleaseTag.ps1',
    'tools/Test-ReleaseTag.ps1',
    'tools/New-PackArtifact.ps1',
    'tools/Validate-PackArtifact.ps1',
    'tools/New-GitHubReleaseBundle.ps1',
    'tools/Validate-GitHubReleaseBundle.ps1',
    'tools/Install-ReleaseToolchain.ps1',
    'tools/Test-PackArtifact.ps1',
    'tools/Test-GitHubReleaseBundle.ps1',
    'tools/Test-GitHubReleaseWorkflow.ps1',
    'tools/Test-CIWorkflow.ps1',
    'tools/Test-PreCommitSafetyContract.ps1',
    'tools/Run-PreCommitCompile.ps1',
    'tools/release-toolchain.json',
    '.github/workflows/ci.yml',
    '.github/workflows/release.yml',
    'benchmarks/schema/github-release-manifest.schema.json',
    'benchmarks/schema/source-package-manifest.schema.json',
    'PSScriptAnalyzerSettings.psd1',
    'tools/Validate-XlamBuildPlan.ps1',
    'tools/Validate-XlamArtifact.ps1',
    'tools/Test-XlamDistribution.ps1',
    'tools/Test-CleanCheckout.ps1',
    'tools/Run-OfficeBitnessValidation.ps1',
    'tools/Test-OfficeBitnessPolicy.ps1',
    'tools/Validate-OfficeBitnessResult.ps1',
    'benchmarks/schema/office-bitness-result.schema.json',
    'docs/specs/office-bitness-validation.md',
    'tools/Run-ProtocolHostValidation.ps1',
    'tools/Validate-ProtocolHostEvidence.ps1',
    'tools/Test-ProtocolHostEvidence.ps1',
    'benchmarks/schema/protocol-host-evidence.schema.json',
    'docs/specs/protocol-host-validation.md',
    'docs/adr/ADR-0025-protocol-host-evidence-boundary.md',
    'build/VBA-HTTP.xlam',
    'examples/ConsumerBasic.bas',
    'examples/ConsumerDiagnostics.bas'
)

foreach ($relativePath in $required) {
    $path = Join-Path $root $relativePath
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
        throw "Documentation contract failed: missing $relativePath"
    }
}

$tokens = @{
    'README.md' = @('docs/API.md', 'docs/guides/README.md', 'CONTRIBUTING.md', 'docs/specs/distribution.md', 'docs/specs/github-release.md', 'LICENSE', 'MIT', 'task verify', 'VBE validation not performed')
    'CONTRIBUTING.md' = @('task check', 'task precommit', 'xlflow push', 'task release:build', 'github-hosted', 'pull request', 'Excel-free CI', 'LICENSE', 'DiagnosticOnly', 'ADR-0039', 'community-validated')
    'docs/API.md' = @('CreateClient', 'HttpResponse', 'HttpError')
    'docs/guides/README.md' = @('api-reference.md', 'distribution.md', 'compatibility.md', '../specs/')
    'docs/guides/api-reference.md' = @('CreateClient', 'HttpClient', 'HttpResponse', 'HttpErrorCategory', 'IHttpTransport', 'HttpCookieJar', 'HttpDiagnostics')
    'docs/guides/getting-started.md' = @('Install-VBAHttp.ps1', 'CreateClient', 'CreateNativeClient', 'x64')
    'docs/guides/distribution.md' = @('source package', 'vbe_validation=not_performed', 'SHA-256', 'LICENSE', 'THIRD_PARTY_NOTICES')
    'docs/specs/distribution.md' = @('xlflow build', 'source package', 'Install-VBAHttp.ps1', 'checksum', 'rollback', 'xlam', 'release:xlam:build', 'Workbook.IsAddin', 'LICENSE', 'THIRD_PARTY_NOTICES', 'github-release.md', 'not performed')
    'docs/specs/development-and-release-workflow.md' = @('github-ci.md', 'ci.yml', 'test:clean-checkout', 'VBA_HTTP_SOURCE_REVISION', 'VBE compilation')
    'docs/RELEASE_CHECKLIST.md' = @('task release:build', 'task release:security', 'task release:xlam:build', 'task test:github-release', 'risk-register', 'task test:license', 'THIRD_PARTY_NOTICES', 'not VBE-validated')
    'docs/specs/office-bitness-validation.md' = @('ADR-0039', 'unverified', 'DiagnosticOnly', 'community')
    'docs/specs/github-ci.md' = @('pull requests', 'windows-2022', 'contents: read', 'task check', 'test:clean-checkout', 'VBE', 'Test-CIWorkflow.ps1')
}

foreach ($relativePath in $tokens.Keys) {
    $content = Get-Content -LiteralPath (Join-Path $root $relativePath) -Raw
    foreach ($token in $tokens[$relativePath]) {
        if ($content.IndexOf($token, [StringComparison]::OrdinalIgnoreCase) -lt 0) {
            throw "Documentation contract failed: $relativePath does not contain '$token'"
        }
    }
}

Write-Output "Documentation contract valid: $($required.Count) required files and $($tokens.Count) token sets."
