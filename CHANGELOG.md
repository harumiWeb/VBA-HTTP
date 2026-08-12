## Unreleased

### Added

- Defined the initial synchronous HTTP core API, dual-transport boundary, buffered body ownership, and stable error model.
- Added the default late-bound WinHTTP COM transport contract, including redirect controls and stable transport-error mapping.
- Added `VBAHttp.CreateClient` for consumers that reference the release workbook and an external release-artifact GET smoke test.
- Added matched Raw WinHttpRequest/VBA-HTTP Phase 2 benchmark evidence and optimized defensive Byte-array copies without weakening ownership.
- Added bounded `ExecuteMany` and `GetMany` execution with ordered per-item results, deadlines, cancellation, partial-failure isolation, and same-client reentrancy protection.
- Added deterministic concurrency integration tests, a server-observed in-flight bound, Phase 3 benchmark evidence, and an external release-artifact batch smoke test.
- Added client-owned retry policy and per-call execution options with idempotency-aware defaults, capped exponential backoff and jitter, `Retry-After`, total deadlines, and active WinHTTP cancellation.

### Fixed

- Made exported class sources clean-importable during `xlflow build` and added a preflight check for the required CRLF representation.
