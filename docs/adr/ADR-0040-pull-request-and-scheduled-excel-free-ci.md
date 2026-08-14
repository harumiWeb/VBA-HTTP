# ADR-0040: Pull-request and scheduled Excel-free CI

## Status

`superseded`

## Background

The repository previously had a GitHub Actions workflow only for strict
version-tag releases. Local Lefthook runs `task precommit`, but a pull request
or a normal push did not receive an equivalent hosted source-quality result.
The project also cannot treat a GitHub-hosted Windows runner as evidence of
Excel/VBIDE compilation, so the hosted path must remain separate from the
local x64 VBE proof loop recorded by ADR-0038.

## Decision

- Add `.github/workflows/ci.yml` for pull requests targeting `main`, pushes to
  `main`, a weekly schedule, and explicit manual runs.
- Run the Excel-free source, documentation, security, deterministic Go,
  release-plan, pack, XLAM, and release-contract gates on a hosted
  `windows-2022` runner.
- Install and probe the versions recorded by `tools/release-toolchain.json`.
- Grant the workflow `contents: read` only, cancel superseded runs for the
  same ref, and fail if validation changes the checkout.
- Run the clean-checkout contract only on scheduled and manual executions
  because it repeats the source gates in a temporary archive.
- Do not start Excel, publish releases, mutate tags, or use a self-hosted
  runner in this workflow. Local `task precommit` and release-host validation
  remain the Excel/VBE evidence boundaries.

## Consequences

- Pull requests receive early feedback for source, tooling, documentation,
  release-plan, and workflow-contract regressions before a tag is created.
- The hosted path remains deterministic and safe for untrusted pull-request
  code because it has no release-write permission and no Excel process scope.
- CI duplicates some setup and gate execution from the tag-release workflow;
  the workflow contract tests and the shared toolchain lock limit drift.
- Excel-backed VBA tests, VBE compilation, protocol-host evidence, stress
  workloads, and Office-bitness promotion remain outside ordinary hosted CI.

## Rationale and evidence

- Current contract: `docs/specs/github-ci.md`,
  `docs/specs/development-and-release-workflow.md`, and
  `docs/specs/github-release.md`.
- Implementation: `.github/workflows/ci.yml`, `Taskfile.yml`, and
  `tools/Test-CIWorkflow.ps1`.
- Regression contract: `task test:ci-workflow` is included in `task check`.
- Related decision: ADR-0038 defines the separate tag-release and VBE
  provenance boundary.

## Supersedes

- None

## Superseded by

- ADR-0041-hosted-excel-free-ci-boundary.md
