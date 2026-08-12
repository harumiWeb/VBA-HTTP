# ADR-0004: Cooperative bounded batch scheduler

## Status

`accepted`

## Background

Excel VBA does not provide safe general-purpose multithreading, but WinHTTP COM can run multiple requests asynchronously. VBA-HTTP needs bounded concurrency without allowing native or COM callbacks to enter arbitrary VBA application code. It also needs deterministic result ordering, partial-failure isolation, cancellation, deadlines, and a clear rule for `DoEvents` reentrancy.

Putting a scheduler directly in `HttpClient` would make the domain layer depend on WinHTTP COM details. Treating every `IHttpTransport` as batch-capable through an implicit sequential fallback would silently violate the concurrency contract and make performance failures difficult to diagnose.

## Decision

- Batch execution remains single-threaded VBA coordination over multiple asynchronous WinHTTP COM operations. It does not create VBA worker threads.
- `IHttpBatchTransport` is a capability interface separate from synchronous `IHttpTransport`. `HttpClient.ExecuteMany` requires this capability and fails before I/O when an injected transport does not implement it.
- `WinHttpComTransport` owns the bounded scheduler and all asynchronous COM request objects until response materialization or abort completes.
- Completion is discovered by cooperative polling. COM events, native callbacks, and callbacks into consumer VBA are not used.
- The scheduler never exceeds `HttpBatchOptions.MaxConcurrency`. It preserves input order in `HttpBatchResult` even when requests complete out of order.
- A request failure is captured in its `HttpBatchItem`; it does not abort unrelated requests. Batch setup validation may still fail the whole call before any I/O.
- `HttpBatchOptions.RequestDeadlineMilliseconds` is measured independently from each operation's start. Deadline expiry becomes a timeout item. A batch cancellation token has priority and marks running and not-yet-started items cancelled.
- Host yielding is controlled by `YieldToHost` and `YieldIntervalMilliseconds`. `DoEvents` is never called more frequently than the configured interval.
- `HttpClient` rejects any synchronous or batch reentrant execution on the same instance while a call is active. Separate clients remain independent.

## Consequences

- Concurrency is bounded and works in ordinary VBA without external dependencies or unsafe callback entry.
- Result order and partial failures are deterministic, while completion time remains nondeterministic by nature.
- Cooperative polling adds small scheduler latency and must balance responsiveness against CPU use.
- `DoEvents` can run unrelated host code; the same-client busy guard prevents state corruption but cannot prevent unrelated workbook macros from running.
- Custom transports must explicitly implement `IHttpBatchTransport` to support batch calls. This is intentional capability discovery rather than a silent sequential fallback.
- Retry and total-deadline policy can later compose above the same item/result model without changing the scheduler's ownership rule.

## Rationale

- Related architecture: `docs/adr/ADR-0002-dual-transport-boundary.md`, `docs/adr/ADR-0003-http-error-model.md`
- Current contract: `docs/specs/bounded-concurrency.md`
- Implemented code: `src/classes/IHttpBatchTransport.cls`, `src/classes/WinHttpComTransport.cls`, `src/classes/HttpBatchResult.cls`
- Regression tests: `src/modules/Tests/Unit/HttpBatchTests.bas`, `src/modules/Tests/Integration/WinHttpConcurrencyTests.bas`

## Supersedes

- None

## Superseded by

- None
