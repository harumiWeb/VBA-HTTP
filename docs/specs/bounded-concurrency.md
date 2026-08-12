# Bounded concurrency

## Public API

`HttpClient.ExecuteMany(Requests As Collection, Optional Options As HttpBatchOptions = Nothing) As HttpBatchResult` executes `HttpRequest` objects. `GetMany(Urls As Collection, Optional Options As HttpBatchOptions = Nothing)` creates GET requests in input order and delegates to `ExecuteMany`.

The input collection is validated completely before I/O: it must be non-`Nothing`, every member must be an `HttpRequest` for `ExecuteMany` or a non-empty `String` for `GetMany`, and the configured transport must implement `IHttpBatchTransport`. An empty collection returns an empty result.

`HttpBatchOptions` defaults are:

| Property | Default | Contract |
| --- | ---: | --- |
| `MaxConcurrency` | 8 | 1 through 64 |
| `PollIntervalMilliseconds` | 5 | 1 through 1,000 |
| `YieldToHost` | `True` | Enables controlled `DoEvents` |
| `YieldIntervalMilliseconds` | 20 | 1 through 1,000 |
| `RequestDeadlineMilliseconds` | 0 | 0 disables the scheduler deadline; otherwise any positive `Long` |
| `CancellationToken` | `Nothing` | Optional shared `HttpCancellationToken` |

The options and all requests are cloned before execution. Mutating caller objects after the call starts cannot alter active operations. The cancellation token is intentionally shared rather than deep-cloned, so a caller can request cancellation while the batch is running.

## Result model

`HttpBatchResult.Count` equals the input count and `ItemAt(1..Count)` follows input order. Every `HttpBatchItem` reaches exactly one terminal status:

- `HttpBatchSucceeded`: contains a non-`Nothing` `Response` and no error.
- `HttpBatchFailed`: contains a stable error number, category, source, and sanitized description. A completed 4xx or 5xx exchange remains succeeded because it produced an `HttpResponse`.
- `HttpBatchCancelled`: contains `HttpErrCancelled` and no response.

`SuccessCount`, `FailureCount`, and `CancelledCount` summarize terminal items. An individual transport failure never raises from a running batch. Pre-I/O validation errors and scheduler invariant failures raise normally.

## Scheduling and deadlines

The COM scheduler opens requests asynchronously and fills no more than `MaxConcurrency` slots. It polls completion without COM event callbacks, materializes completed responses, and fills newly available slots. Results are reordered to the original input positions.

Each started operation receives its own monotonic deadline. Deadline expiry aborts that operation and records `HttpErrorTimeout`. WinHTTP resolve/connect/send/receive timeouts continue to apply independently; whichever condition is observed first wins.

Cancellation is checked before filling slots and during every poll cycle. Once requested, running operations are aborted and all not-yet-started items become `HttpBatchCancelled`. Responses already materialized remain succeeded or failed.

## Host yielding and reentrancy

The scheduler sleeps for `PollIntervalMilliseconds` between incomplete poll cycles. When host yielding is enabled, it calls `DoEvents` no more frequently than `YieldIntervalMilliseconds`.

xlflow correctly identifies the explicit `DoEvents` statement as a GUI boundary during static preflight. Unattended tests and benchmarks set `YieldToHost=False`; their runners omit `--headless` only to bypass that conservative static rejection and still run Excel invisibly with alerts disabled. Human interaction is neither required nor permitted by those runners.

Any `Execute`, `ExecuteMany`, or `GetMany` call made reentrantly on the same `HttpClient` raises validation before I/O. A different `HttpClient` may execute independently. The guard is released on success and on every error path.

## Evidence

- Architectural decision: `docs/adr/ADR-0004-cooperative-batch-scheduler.md`
- Unit tests: `src/modules/Tests/Unit/HttpBatchTests.bas`
- Loopback integration: `src/modules/Tests/Integration/WinHttpConcurrencyTests.bas`
