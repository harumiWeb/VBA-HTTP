# ADR-0001: Source and workbook authority

## Status

`accepted`

## Background

VBA-HTTP needs a workbook for Excel execution, but embedded VBA is difficult to review and merge. Development also requires tests and xlflow helpers that must not ship in the release workbook. A reproducible build therefore needs an explicit authority and artifact boundary.

## Decision

- Files under `src/` are the authoritative editable VBA source.
- `build/VBA-HTTP.xlsm` is the tracked development workbook used for synchronization and Excel execution. Production VBA is not edited directly in the VBE.
- The development workbook contains production, test, benchmark, xlflow, and development-only components.
- `dist/VBA-HTTP-source.zip` is the primary module distribution generated from
  authoritative source. It excludes workbook document modules and development
  components and is installed into a consumer workbook by its manifest-based
  installer.
- Release workbooks are optional compiled artifacts generated only with
  `xlflow build` from the tracked development workbook and authoritative source.
  They are separate, generated artifacts.
- `[build].exclude` removes development-only components. Production components must not depend on excluded components.
- A release build plan is checked against an explicit component policy before Excel creates an artifact.
- The GitHub tag-release pipeline uses `xlflow pack --experimental` only after
  production-only staging. Its manifest always says `vbe_validation` is
  `not_performed`; it is not a substitute for the local `xlflow build` VBE
  gate. The source ZIP remains the primary distribution.

If the workbook is changed directly during exceptional recovery, contributors must pull and reconcile it before resuming source edits. A release workbook is never used as an editable source.

## Consequences

- Source changes remain reviewable and mergeable while the tracked workbook makes a fresh clone executable.
- The repository stores one binary development workbook, so source/workbook synchronization must be verified before commits that change VBA.
- Release composition is deterministic and auditable, but adding a production component requires an intentional component-policy update.
- External consumer smoke tests are needed because release workbooks do not contain their own test helpers.
- Source-package validation is Excel-free; workbook compile and runtime smoke
  remain separate gates for optional compiled artifacts.

## Rationale

- Tests: `src/modules/Tests/BootstrapTests.bas`
- Code: `xlflow.toml`, `tools/Validate-BuildPlan.ps1`
- Related specs: `docs/specs/development-and-release-workflow.md`,
  `docs/specs/source-package.md`, and `docs/specs/distribution.md`

## Supersedes

- None

## Superseded by

- None
