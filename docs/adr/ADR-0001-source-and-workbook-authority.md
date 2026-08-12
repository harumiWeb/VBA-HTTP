# ADR-0001: Source and workbook authority

## Status

`accepted`

## Background

VBA-HTTP needs a workbook for Excel execution, but embedded VBA is difficult to review and merge. Development also requires tests and xlflow helpers that must not ship in the release workbook. A reproducible build therefore needs an explicit authority and artifact boundary.

## Decision

- Files under `src/` are the authoritative editable VBA source.
- `build/VBA-HTTP.xlsm` is the tracked development workbook used for synchronization and Excel execution. Production VBA is not edited directly in the VBE.
- The development workbook contains production, test, benchmark, xlflow, and development-only components.
- Release workbooks are generated only with `xlflow build` from the tracked development workbook and authoritative source. They are separate, generated artifacts.
- `[build].exclude` removes development-only components. Production components must not depend on excluded components.
- A release build plan is checked against an explicit component policy before Excel creates an artifact.
- `xlflow pack` is not part of the release pipeline because it does not provide the Excel/VBE compile validation required by this project.

If the workbook is changed directly during exceptional recovery, contributors must pull and reconcile it before resuming source edits. A release workbook is never used as an editable source.

## Consequences

- Source changes remain reviewable and mergeable while the tracked workbook makes a fresh clone executable.
- The repository stores one binary development workbook, so source/workbook synchronization must be verified before commits that change VBA.
- Release composition is deterministic and auditable, but adding a production component requires an intentional component-policy update.
- External consumer smoke tests are needed because release workbooks do not contain their own test helpers.

## Rationale

- Tests: `src/modules/Tests/BootstrapTests.bas`
- Code: `xlflow.toml`, `tools/Validate-BuildPlan.ps1`
- Related specs: `docs/specs/development-and-release-workflow.md`

## Supersedes

- None

## Superseded by

- None
