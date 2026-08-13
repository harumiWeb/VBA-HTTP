# ADR-0022: COM transport timeout failure cleanup

## Status

`accepted`

## Context

The buffered COM transport owns a late-bound `WinHttp.WinHttpRequest.5.1`
object for each synchronous request and for each asynchronous attempt. When
`Open`, `Send`, or a response operation raises a WinHTTP error, simply dropping
the VBA reference does not provide an explicit cancellation boundary. Repeated
receive-timeout measurements on the current host show handle growth, so the
failure path must make the best cleanup operation available without obscuring
the original transport error.

## Decision

- On every COM backend-construction or synchronous execution failure, attempt
  `Abort` before releasing the backend reference.
- After a successful synchronous-failure `Abort`, wait 250 ms for WinHTTP's
  asynchronous COM-handle teardown to drain. The asynchronous cancellation
  path uses the same 250 ms bounded drain for the same reason; neither delay is
  added to successful requests.
- If `Abort` itself fails, release the reference anyway and append only the
  stable cleanup note to the already-captured error description. The original
  WinHTTP code and public error category remain authoritative.
- Apply the same best-effort cleanup when asynchronous operation creation fails;
  operation cancellation paths continue to use their existing bounded abort
  helper.
- Do not pool or reuse a failed COM request object. A future mitigation for
  repeated receive-timeout handle growth must be measured separately and must
  not change request isolation or retry semantics implicitly.

## Consequences

- Timeout, connection, and malformed-response failures release the COM request
  through an explicit cancellation boundary before the reference is dropped.
- The bounded drain adds up to 250 ms to either failed or cancelled requests,
  trading abort-path latency for lower persistent-handle risk on the supported
  x64 host.
- Cleanup failure is observable only as a sanitized suffix; credentials, URLs,
  headers, response bodies, and raw COM descriptions remain excluded.
- This reduces avoidable failure-path lifetime but does not claim that WinHTTP
  or the COM wrapper has zero repeated receive-timeout growth. The dedicated
  stress/risk gate remains required before v1.0 promotion.

## Evidence

- Code: `src/classes/WinHttpComTransport.cls`
- Contract: `docs/specs/buffered-com-transport.md`
- Existing mapping test: `WinHttpComTransportTests.Test_ComTransport_MapsReceiveTimeout`
- Residual risk: `docs/security/risk-register.json` issue
  `com-receive-timeout-abort-growth`

## Supersedes

- None

## Superseded by

- None
