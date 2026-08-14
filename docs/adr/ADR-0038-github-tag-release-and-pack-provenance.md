# ADR-0038: GitHub tag releases and pack provenance

## Status

`accepted`

## Context

The source ZIP is the primary library distribution, but consumers also benefit
from a production-only XLSM. GitHub-hosted runners cannot safely provide the
local Excel/VBIDE proof loop, and `xlflow pack --experimental` deliberately does
not compile the workbook in VBE. A release must therefore distinguish source
and pack provenance instead of presenting a pack artifact as VBE-verified.

## Decision

- A push of a strict `vX.Y.Z` or prerelease SemVer tag starts
  `.github/workflows/release.yml` on a GitHub-hosted `windows-2022` runner.
- The workflow validates the tag, its commit (`GITHUB_SHA`), the clean checkout,
  and all Excel-free source, documentation, security-risk, Go, pack, and bundle
  gates before publication.
- `tools/New-PackArtifact.ps1` stages only the production component allowlist
  from `tools/build-component-policy.json` and invokes
  `xlflow pack --experimental` with the staged template. Tests, benchmarks,
  Xlflow helpers, Dev modules, and legacy `App`/`Main`/`Ui` components cannot
  enter the staging tree.
- Pack manifests always contain `pack_backend=pure-go`, `experimental=true`,
  and `vbe_validation=not_performed`. The GitHub release notes repeat that
  VBE validation was not performed in Actions.
- The exact release asset set is a VBA-Web-style source ZIP, production-only
  pack XLSM, pack/release manifests, SHA-256 sums, `LICENSE`, and
  `THIRD_PARTY_NOTICES.md`. Existing Releases are never overwritten and
  `gh release create --verify-tag` is the only publication path.
- Local `task precommit` runs `task check` followed by an ownership-safe,
  temporary `xlflow build` VBE compile. The GitHub workflow never starts Excel.
- The support target is Windows x64 Office. 32-bit Office is recorded as
  `unverified` and is not validated by the tag workflow; HTTP/3/QUIC and XLAM
  remain outside this tag-release contract.

## Consequences

- Consumers can download a small, auditable source package or a production-only
  XLSM while seeing the exact compile boundary.
- A pack artifact is reproducible and inspectable but is not evidence of VBE
  compilation; local pre-commit/build evidence remains a separate artifact.
- Release reruns with an existing Release fail closed and require operator
  review of any draft or partial publication on GitHub.
- The workflow requires network access to pinned tool archives and GitHub's
  release API, while all project tests remain loopback/offline deterministic.

## Rationale and evidence

- Current contract: `docs/specs/github-release.md`,
  `docs/specs/distribution.md`, and
  `docs/specs/development-and-release-workflow.md`.
- Implementation: `tools/New-PackArtifact.ps1`,
  `tools/New-GitHubReleaseBundle.ps1`,
  `tools/Validate-GitHubReleaseBundle.ps1`, and `.github/workflows/release.yml`.
- Excel-free regression: `tools/Test-PackArtifact.ps1`,
  `tools/Test-GitHubReleaseBundle.ps1`, and
  `tools/Test-GitHubReleaseWorkflow.ps1`.
- External references: [GitHub workflow syntax](https://docs.github.com/en/actions/reference/workflows-and-actions/workflow-syntax),
  [GitHub token permissions](https://docs.github.com/en/actions/reference/workflows-and-actions/workflow-syntax#permissions),
  [gh release create](https://cli.github.com/manual/gh_release_create), and
  [xlflow pack](https://github.com/harumiWeb/xlflow/blob/main/vitepress/commands/pack.md).

## Supersedes

- None

## Superseded by

- None
