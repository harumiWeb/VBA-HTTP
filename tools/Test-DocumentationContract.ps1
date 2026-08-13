[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$root = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$required = @(
    'README.md',
    'CONTRIBUTING.md',
    'docs/API.md',
    'docs/RELEASE_CHECKLIST.md',
    'docs/specs/distribution.md',
    'docs/specs/development-and-release-workflow.md',
    'docs/adr/ADR-0021-xlam-distribution-target.md',
    'docs/adr/ADR-0022-com-timeout-failure-cleanup.md',
    'docs/adr/ADR-0030-32-bit-office-support-boundary.md',
    'docs/adr/ADR-0035-http3-support-boundary.md',
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
    'README.md' = @('docs/API.md', 'CONTRIBUTING.md', 'docs/specs/distribution.md', 'task verify')
    'CONTRIBUTING.md' = @('task check', 'xlflow push', 'task release:build')
    'docs/API.md' = @('CreateClient', 'HttpResponse', 'HttpError')
    'docs/specs/distribution.md' = @('xlflow build', 'checksum', 'rollback', 'xlam', 'release:xlam:build', 'Workbook.IsAddin')
    'docs/RELEASE_CHECKLIST.md' = @('task release:build', 'task release:security', 'task release:xlam:build', 'risk-register')
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
