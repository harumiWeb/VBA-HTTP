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
- `task test:office-bitness`: 234 discovered tests, 149 passed, 85
  inconclusive (environment-gated suites without the loopback server), 0
  failed; the dedicated loopback integration run above passed all 78 tests.
- Native contract coverage includes GET headers/query, text methods, binary body,
  redirect/status behavior, connection error mapping, negotiated protocol query,
  and 20 repeated requests with process handle-count regression.
- Native source was pushed to the development workbook, compiled by VBE, saved,
  and the source/workbook status was clean afterward.
- `task release:build`: production-only artifact compiled, published atomically,
  passed component inspection, and passed external COM/native consumer smoke.

## Bitness limitation

This host cannot produce 32-bit Office evidence. Before the Phase 5 exit gate is
marked complete, repeat the same source push, VBE compile, focused integration
suite, full verification, and production release build on a dedicated 32-bit
Office host and record its Office/Windows/xlflow versions and artifact checksum.
