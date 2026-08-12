# ADR-0006: Native WinHTTP callback and handle ownership boundary

## Status

`accepted`

## Background

Native WinHTTP exposes an explicit `HINTERNET` hierarchy (session, connection,
and request handles) and an asynchronous status-callback API. VBA does not have
an ownership type or a safe asynchronous callback lifetime: a callback can run
after a class has been released, re-enter Excel while a public call is active,
or retain a stale `AddressOf` target. A callback that invokes VBA application
logic would therefore make cancellation, cleanup, and release-build behavior
non-deterministic.

## Decision

- Phase 5 uses synchronous WinHTTP calls and registers no status callback.
- Production code never passes `AddressOf`, a VBA object reference, or an
  application callback to `WinHttpSetStatusCallback`.
- A native request is owned by exactly one `WinHttpRequestHandle`; that wrapper
  retains its `WinHttpConnectionHandle`, which retains its
  `WinHttpSessionHandle`. There is no handle cloning or transfer operation.
- Each wrapper exposes idempotent `ReleaseHandle` and also closes its handle from
  `Class_Terminate` as a last-resort safety net. Transport code closes request,
  connection, and session explicitly on both success and failure.
- A failed native call is mapped immediately to the stable `HttpErrors`
  category. Cleanup must not replace the original failure; a cleanup failure is
  diagnostic-only.
- Native buffered transport implements the existing `IHttpTransport` contract.
  Active cancellation and total-deadline enforcement require
  `IHttpAttemptTransport` and remain unavailable until a later native streaming
  design proves a safe polling/callback boundary.
- TLS certificate validation uses WinHTTP's default OS validation. There is no
  ignore-certificate option in this transport.

## Consequences

- The Phase 5 transport is deterministic and safe to call from Excel's VBA
  thread, at the cost of no mid-call cancellation.
- Handle leaks are testable by repeating complete request lifecycles and
  observing the process handle count.
- Later streaming may use native callbacks only for a primitive event queue if
  a new ADR proves callback lifetime and reentrancy safety; this ADR does not
  authorize callbacks by implication.
- Consumers receive the same request, response, and sanitized error contract
  as the COM backend. Backend-specific capability failures are explicit.

## Rationale

- Code: `src/modules/WinHttpNativeApi.bas`,
  `src/classes/WinHttpNativeTransport.cls`, and the three handle wrappers.
- Tests: native loopback integration and repeated-handle regression tests.
- Related specs: `docs/specs/native-winhttp-transport.md` and
  `docs/adr/ADR-0002-dual-transport-boundary.md`.
- API references: [HINTERNET handles in WinHTTP](https://learn.microsoft.com/en-us/windows/win32/winhttp/hinternet-handles-in-winhttp),
  [WinHttpOpenRequest](https://learn.microsoft.com/en-us/windows/win32/api/winhttp/nf-winhttp-winhttpopenrequest),
  and [WinHTTP option flags](https://learn.microsoft.com/en-us/windows/win32/winhttp/option-flags).

## Supersedes

- None

## Superseded by

- None
