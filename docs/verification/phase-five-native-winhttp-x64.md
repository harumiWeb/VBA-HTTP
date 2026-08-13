# Phase 5 native WinHTTP evidence (x64)

Recorded: 2026-08-13 (refreshed against the current source revision)

## Host and build

- Excel bridge architecture: X64
- Excel bridge runtime: .NET 8.0.30
- Development workbook: `build/VBA-HTTP.xlsm`
- Native declarations compiled through the VBA7/Win64 branch.

## Proof results

- `xlflow lint --json`: 0 issues
- `xlflow analyze --json`: 0 findings, 0 warnings
- `task format`: clean
- `task class-source:check`: clean
- `task test:integration`: 78/78 test cases passed as reported by xlflow
- `task test:office-bitness`: 239 discovered tests, 154 passed, 85
  inconclusive (environment-gated suites without the loopback server), 0
  failed; the dedicated loopback integration run above passed all 78 tests.
- Native contract coverage includes GET headers/query, text methods, binary body,
  redirect/status behavior, connection error mapping, negotiated protocol query,
  and 20 repeated requests with process handle-count regression.
- Native source was pushed to the development workbook, compiled by VBE, saved,
  and the source/workbook status was clean afterward.
- `task release:build`: production-only artifact compiled, published atomically,
  passed component inspection, and passed external COM/native consumer smoke.

## Bitness boundary

The supported runtime target is Windows x64 Office. 32-bit Office is outside
the support matrix under ADR-0030; no x86 execution evidence is required for
the Phase 5 release gate. The VBA7/legacy declaration contract remains an
Excel-free ABI regression guard only. A future x86 support proposal would need
to supersede ADR-0030 and provide a dedicated compile, integration, release,
and consumer-smoke evidence bundle.
