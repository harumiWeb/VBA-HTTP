# ADR-0005: Client-owned reliability policy

## Status

`proposed`

## Background

Retry is an HTTP policy decision, while cancellation during an active send or receive is a transport capability. Putting both concerns inside WinHTTP would couple public behavior to one backend. Putting all waiting and cancellation in `HttpClient` would leave a blocking transport attempt uninterruptible. Batch execution adds a further constraint: one retrying item must not occupy a concurrency slot while it is waiting, and it must not fail unrelated items.

The policy must also resolve observable conflicts among cancellation, total deadline, per-attempt timeouts, `Retry-After`, idempotency, and the existing rule that HTTP error statuses are returned as responses.

## Decision

- `HttpClient` owns retry classification, attempt counting, delay calculation, and the total operation deadline. The same policy applies to synchronous and batch calls.
- `IHttpAttemptTransport` is an optional capability beside `IHttpTransport` and `IHttpBatchTransport`. It executes exactly one attempt with cooperative cancellation and a caller-supplied attempt deadline. WinHTTP COM implements it with async polling; it does not implement retry policy.
- A custom transport without `IHttpAttemptTransport` can participate in retries, but cannot be used when an active cancellation token or total deadline is requested. That combination fails validation before I/O rather than pretending a blocking attempt is interruptible.
- The default `HttpRetryPolicy` makes at most three attempts. Only `GET`, `HEAD`, `PUT`, `DELETE`, `OPTIONS`, and `TRACE` retry by default. `POST` and `PATCH` require explicit non-idempotent opt-in.
- Default retry statuses are 408, 429, 500, 502, 503, and 504. Default retryable transport categories are DNS, connection, timeout, and I/O. Validation, invalid URL, TLS, cancellation, protocol, and explicit HTTP-status errors are not retried.
- Exponential backoff is capped and receives symmetric jitter. A valid `Retry-After` delta-seconds or HTTP-date on a retryable response replaces exponential delay and is capped by the configured maximum delay; server-directed delay is not jittered.
- Priority is cancellation, then total deadline, then the active attempt's timeout/failure, then retry eligibility and delay. Cancellation wins when cancellation and deadline are observed at the same checkpoint.
- A final retryable HTTP response is returned normally after attempts are exhausted. A final transport failure is raised through the existing stable error model.
- Batch retries are scheduled per item. Waiting retries consume no in-flight slot, results remain in input order, and cancellation stops running, waiting, and not-yet-started work.

## Consequences

- Retry semantics are backend-independent and testable without network access.
- The WinHTTP COM transport must expose a cancellable single-attempt path in addition to its existing synchronous and batch capabilities.
- Enabling non-idempotent retries can duplicate side effects after ambiguous network failures; the opt-in is deliberately explicit and documented.
- Respecting a capped `Retry-After` avoids unbounded sleeps, but may retry earlier than a server requested when the configured maximum is lower.
- Total deadline enforcement for custom transports requires the new capability; compatibility is preserved for ordinary calls without cancellation or total deadline.
- Cooperative waits may run `DoEvents` when host yielding is enabled, so the existing same-client reentrancy guard remains mandatory.

## Rationale

- Related decisions: `docs/adr/ADR-0003-http-error-model.md`, `docs/adr/ADR-0004-cooperative-batch-scheduler.md`
- Planned contract: `docs/specs/reliability-policy.md`
- Planned code: `src/classes/HttpRetryPolicy.cls`, `src/classes/HttpExecutionOptions.cls`, `src/classes/IHttpAttemptTransport.cls`
- Planned tests: `src/modules/Tests/Unit/HttpRetryPolicyTests.bas`, `src/modules/Tests/Integration/WinHttpReliabilityTests.bas`

## Supersedes

- None

## Superseded by

- None
