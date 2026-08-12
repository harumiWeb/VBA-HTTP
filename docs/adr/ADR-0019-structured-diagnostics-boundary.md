# ADR-0019: Structured diagnostics boundary and secret redaction

## Status

`accepted`

## Background

`Err` is intentionally small and process-global. It cannot provide a stable,
machine-readable account of completed HTTP operations, while transport-native
messages and request headers may contain credentials, cookies, or query data.
The existing `HttpSecurity.RedactHeaderValue` helper prevents accidental header
logging, but there was no public diagnostics object that applied the rule
consistently.

## Decision

- Expose an opt-in, caller-owned `HttpDiagnostics` collector on `HttpClient`.
  It is disabled by default and is not a global logger or a transport callback.
- Record one top-level event for each completed client operation (`execute`,
  `batch`, `download`, or `upload`). Retry attempts and native handle details
  remain internal; callers that need wire-level tracing must instrument their
  test server or transport explicitly.
- Events contain only stable fields: sequence, operation, method, a target with
  query/fragment and URL user-info removed, status, negotiated protocol,
  elapsed milliseconds, stable library error number/category, and ordered
  request/response headers.
- Header values are redacted at event construction for all sensitive names
  (`Authorization`, `Proxy-Authorization`, `Cookie`, `Set-Cookie`,
  `WWW-Authenticate`, and `Proxy-Authenticate`). Request/response bodies,
  reason phrases, error descriptions, query values, and backend-native error
  text are never included.
- The collector is bounded by `MaxEvents` (default 100); adding an event drops
  the oldest event. `ToJson` emits schema version 1 with deterministic field
  order and no wall-clock timestamp.
- Diagnostics failures are fail-open: a collector must never replace the
  operation result or expose a secret. A malformed custom transport error is
  represented only by the stable number/category when available.

## Consequences

- Consumers can persist or inspect safe operation evidence without parsing
  localized VBA error descriptions.
- The event stream is deliberately not a packet trace and does not prove every
  retry attempt; benchmark and security evidence continues to use dedicated
  machine-readable artifacts.
- The collector is mutable caller-owned state. Applications should provide a
  separate collector per client or clear it at an operation boundary.
- Adding a new sensitive header requires updating `HttpSecurity` and the
  diagnostics redaction tests together.

## Rationale and evidence

- Code: `src/classes/HttpDiagnostics.cls`, `src/classes/HttpDiagnosticEvent.cls`,
  `src/modules/HttpJson.bas`, `src/modules/HttpSecurity.bas`, and
  `src/classes/HttpClient.cls`.
- Tests: `src/modules/Tests/Unit/HttpDiagnosticsTests.bas` and the release
  consumer diagnostics smoke test.
- Specification: `docs/specs/diagnostics-policy.md`.

## Supersedes

- None

## Superseded by

- None
