# Phase -1 baseline evidence

Recorded: 2026-08-12

## Environment

- Excel bridge architecture: X64
- Excel bridge runtime: .NET 8.0.30
- Development workbook: `build/VBA-HTTP.xlsm`

## Source and workbook proof

- `xlflow lint --json`: 0 issues
- `xlflow analyze --json`: 0 findings, 0 warnings
- `tools/Check-Format.ps1`: 11 checked files clean; the single documented `XlflowAssert.bas` formatter exception remains
- Source was pushed to the managed development workbook and the saved workbook is not older than source.
- `xlflow test --json`: 3 passed, 0 failed, 0 skipped, 0 todo
- A separate `git clone --no-local` of commit `a69cd89` resolved the tracked workbook and passed `task verify` without using the original worktree's ignored build or xlflow state.
- Fresh-clone issues found and fixed during this proof: absent empty source roots and Git line-ending conversion.

## Release build proof

- `tools/Validate-BuildPlan.ps1`: 5 included, 7 excluded
- `xlflow build --json --out build/Release/VBA-HTTP.xlsm`: success
- Source application: passed
- VBE compile: passed
- Workbook save: passed
- Workbook close: passed
- Excel cleanup: clean
- Publication: atomic replace
- Included components: `App`, `Main`, `Ui`, `Sheet1`, `ThisWorkbook`
- Excluded components: `BenchmarkSupport`, `DevSupport`, `BootstrapTests`, `XlflowAssert`, `XlflowDebug`, `XlflowRuntime`, `XlflowUI`
- `task release:smoke`: the actual workbook exposed exactly the 5 allowed VBA components and external `Main.Run` execution passed without source injection

Release artifacts and manifests are generated evidence and remain ignored. This document records the baseline result; future release milestones must retain their generated manifest and checksum as release evidence.

## Office support boundary

The supported runtime target is Windows x64 Office. 32-bit Office is
`unsupported-by-policy` under ADR-0030, so no x86 compile, integration, or
release artifact is required for the current release gate. The Excel-free
VBA7/legacy declaration check remains a source ABI regression guard. A future
support proposal must supersede ADR-0030 and include a dedicated real-host
evidence bundle; it must not be inferred from this x64 baseline.
