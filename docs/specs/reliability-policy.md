# Reliability policy

## Public configuration

Every `HttpClient` owns a cloneable `RetryPolicy`. Its defaults are:

| Property | Default | Contract |
| --- | ---: | --- |
| `MaxAttempts` | 3 | 1 through 100, including the initial attempt |
| `BaseDelayMilliseconds` | 200 | 0 through 2,147,483,647 |
| `MaxDelayMilliseconds` | 30,000 | 0 through 2,147,483,647 and not less than base delay |
| `JitterRatio` | 0.2 | 0 through 1 |
| `RespectRetryAfter` | `True` | Use a valid server delay before exponential backoff |
| `RetryNonIdempotentMethods` | `False` | Explicit opt-in for `POST` and `PATCH` |

`HttpExecutionOptions` supplies a per-call retry-policy override, `TotalDeadlineMilliseconds`, `CancellationToken`, poll interval, and controlled host-yield settings. A missing override clones `HttpClient.RetryPolicy`. Caller mutations after execution starts do not affect the active call, except that the cancellation token identity is intentionally shared.

`HttpBatchOptions` supplies the same retry-policy override and total deadline in addition to its concurrency and per-request scheduler deadline. The total deadline covers the complete batch; the request deadline restarts for each attempt and cannot extend the total deadline.

## Retry classification

Default retryable methods are `GET`, `HEAD`, `PUT`, `DELETE`, `OPTIONS`, and `TRACE`. `POST` and `PATCH` are never retried unless `RetryNonIdempotentMethods=True`. Unknown extension methods are treated as non-idempotent.

Default retryable response statuses are 408, 429, 500, 502, 503, and 504. A retryable status remains an ordinary `HttpResponse`: the client retries while eligible and returns the final response when attempts are exhausted. It does not convert the final status to a VBA error.

Default retryable transport categories are `HttpErrorDns`, `HttpErrorConnection`, `HttpErrorTimeout`, and `HttpErrorIo`. All other stable categories and unknown VBA errors are final. Cancellation is never retried.

## Delay calculation

After failed attempt `n`, exponential delay is `BaseDelayMilliseconds * 2^(n-1)`, saturated at `MaxDelayMilliseconds`. Jitter multiplies it by a value uniformly selected from `1-JitterRatio` through `1+JitterRatio`, then clamps it to 0 through the maximum.

On a retryable response, a valid `Retry-After` header replaces exponential delay when enabled. Decimal delta-seconds and HTTP-date are supported. Negative or past dates become zero delay; malformed values fall back to exponential delay. Server delay is capped at `MaxDelayMilliseconds` and is not jittered.

## Deadline and cancellation precedence

Checkpoints occur before each attempt, while an async attempt is polled, throughout retry waits, and after each attempt. At every checkpoint:

1. cancellation raises `HttpErrCancelled`;
2. an expired total deadline raises `HttpErrTimeout`;
3. otherwise active transport outcome or retry policy is evaluated.

The total deadline starts before the first attempt and includes transport time and retry waits. A retry is not started when its delay or next attempt cannot fit within the remaining deadline. WinHTTP phase timeouts remain per attempt; the earliest observed limit wins.

## Transport capability

`IHttpAttemptTransport.ExecuteAttempt` performs one attempt and supports polling cancellation plus an attempt deadline. `WinHttpComTransport` implements this capability using async WinHTTP without callbacks into consumer VBA.

A custom transport that implements only `IHttpTransport` remains valid for calls without an execution cancellation token or total deadline. Supplying either control with such a transport fails validation before any I/O. This prevents a false guarantee that a blocking custom transport can be interrupted.

## Batch behavior

Each batch item tracks its own attempt number and next eligible start time. Retry waiting consumes no concurrency slot. The scheduler admits eligible items in stable input order, never exceeds `MaxConcurrency`, and preserves final results in input order. A final per-item transport failure remains isolated; cancellation or total batch deadline terminates all unfinished items with the corresponding stable category.

## Deterministic test seams

The reliability engine obtains monotonic time, UTC wall time, jitter samples, and waits from an internal runtime abstraction. Production uses Windows clocks and cooperative sleep. Unit tests use a fake runtime that advances virtual time and supplies fixed jitter values; integration tests use the real runtime and loopback server. No unit test depends on wall-clock sleeps or external network access.

## Evidence

- Decision: `docs/adr/ADR-0005-client-owned-reliability-policy.md`
- Planned unit tests: `src/modules/Tests/Unit/HttpRetryPolicyTests.bas`
- Planned integration tests: `src/modules/Tests/Integration/WinHttpReliabilityTests.bas`
