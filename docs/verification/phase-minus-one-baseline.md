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

Release artifacts and manifests are generated evidence and remain ignored. This document records the baseline result; future release milestones must retain their generated manifest and checksum as release evidence.

## 32-bit validation route

32-bit Office evidence cannot be produced by the current X64 environment. Before a milestone that requires 32-bit support, run the same `task verify`, source push/test, and release build sequence on a dedicated Windows host with 32-bit Excel. Record Office version, Windows version, xlflow version, command results, manifest component sets, and artifact checksum under `docs/verification/`.
