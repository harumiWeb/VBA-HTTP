# ADR-0041: Hosted Excel-free CI boundary

## Status

`accepted`

## Background

ADR-0040 established a hosted workflow intended to be Excel-free, but the
workflow also invoked `task test:xlam`. That target opens the tracked XLAM
through Excel COM to verify `Workbook.IsAddin`. GitHub-hosted Windows runners
do not provide an installed Excel COM server, so the hosted job failed before
the remaining release-plan gates could complete. The clean-checkout helper had
the same hidden dependency because it called `task test:xlam` internally.

## Decision

- Keep `task test:xlam` as a local or Excel-host validation target; it remains
  the authoritative tracked XLAM identity check.
- Hosted `.github/workflows/ci.yml` runs the Excel-free `build:plan:xlam`
  dry-run, but never invokes `task test:xlam`.
- Add `-ExcelFree` to `tools/Test-CleanCheckout.ps1` and expose it through
  `task test:clean-checkout:excel-free`. The scheduled and manual hosted path
  uses this target, while local `task test:clean-checkout` retains its
  host-bound XLAM identity check.
- Make `tools/Test-CIWorkflow.ps1` require the Excel-free clean-checkout target
  and reject `test:xlam` in the hosted workflow.

## Consequences

- Pull requests, main pushes, and scheduled/manual hosted runs no longer
  depend on an unavailable Excel COM registration.
- Hosted CI still validates the XLAM inclusion/exclusion plan and source
  package/release contracts, but it cannot establish add-in identity or VBE
  compilation evidence.
- Local contributors and an Excel-capable validation host must continue to run
  `task test:xlam` (or the full `task test:clean-checkout`) for that evidence.
- The workflow and clean-checkout contracts have an explicit Excel-free target,
  reducing the chance that a host-bound gate is reintroduced accidentally.

## Rationale and evidence

- Failure evidence: GitHub Actions run `31792963243` failed at
  `New-Object -ComObject Excel.Application` with `REGDB_E_CLASSNOTREG` on the
  hosted `windows-2022` runner.
- Implementation: `.github/workflows/ci.yml`, `Taskfile.yml`,
  `tools/Test-CleanCheckout.ps1`, and `tools/Test-CIWorkflow.ps1`.
- Current contracts: `docs/specs/github-ci.md`,
  `docs/specs/development-and-release-workflow.md`, and
  `docs/specs/distribution.md`.
- Regression gate: `task test:ci-workflow` runs as part of `task check`.

## Supersedes

- ADR-0040-pull-request-and-scheduled-excel-free-ci.md

## Superseded by

- None
