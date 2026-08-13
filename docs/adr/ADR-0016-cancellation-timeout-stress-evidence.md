# ADR-0016: Cancellation and timeout stress evidence policy

## Status

`accepted`

## Background

The reliability and streaming integration suites already prove individual
cancellation and timeout paths, but a single successful cycle cannot establish
that handles, temporary files, or client state remain usable after repeated
abortions. Native buffered requests are also synchronous: the public total
deadline cannot interrupt a blocking `WinHttpReceiveResponse` call. The stress
evidence must therefore exercise only capabilities the transports actually
promise and must not claim stronger interruption semantics.

## Decision

- Add an excluded `WinHttpCancellationStressTests` module with four independent
  scenarios, each executed in its own temporary xlflow/Excel process:
  `com_active_cancellation` cancels four delayed asynchronous COM requests,
  `com_request_deadline` expires four COM request deadlines,
  `com_receive_timeout` repeats synchronous receive timeouts, and
  `native_download_cancellation` cancels a 64 KiB streaming download while
  preserving its sentinel destination and temporary-file count.
- Use 25 iterations by default, with a bounded 1..1000 override for diagnosis.
  Every iteration performs a loopback recovery `GET /status/204`; COM retry
  policy is explicitly limited to one attempt so retry behavior cannot hide a
  cancellation or timeout.
- The PowerShell runner starts the deterministic local server, records the
  existing Excel PID baseline for the complete gate, samples only newly created
  processes across all scenarios, waits for VBA start/done markers, captures
  before/peak/after/idle process snapshots,
  opens a release marker, and publishes JSON atomically. Existing user Excel
  processes are never terminated. Native idle handle growth must be <=8 and
  COM growth <=32. Memory peaks are recorded but remain host-dependent
  evidence; concurrent user-launched Excel is an inconclusive run condition.
- The result is loopback-only, contains no credentials or request headers, and
  is validated against the checked-in schema. Native total-deadline interruption
  during a blocking receive and repeated `Application.OnTime` active-cancel
  behavior beyond the bounded COM scenario remain outside this contract.

## Consequences

- Repeated cleanup and recovery behavior becomes a measurable Phase 9 gate
  instead of an informal claim based on one integration test.
- The gate is slower and requires an Excel host, so it is a separate Taskfile
  target and is not part of the normal pre-commit suite.
- The four scenarios deliberately reflect transport capability differences;
  native buffered operations do not promise mid-call cooperative cancellation.
  The receive-timeout scenario is the dedicated regression boundary for the
  synchronous COM failure cleanup described by ADR-0022.
- Cookie persistence and TLS certificate rejection are covered by their own
  policies and tests. Integrated authentication and process-wide memory
  thresholds remain separate hardening decisions.

## Rationale

- Code: `src/modules/Tests/Stress/WinHttpCancellationStressTests.bas`.
- Runner: `tools/Run-CancellationStressTests.ps1` and
  `tools/Validate-CancellationStressResult.ps1`.
- Evidence: `benchmarks/schema/cancellation-stress-result.schema.json` and
  `benchmarks/results/phase9-cancellation-stress.json`.
- Related contracts: `docs/specs/reliability-policy.md`,
  `docs/specs/streaming-download.md`, `docs/specs/streaming-upload.md`, and
  `ADR-0006-native-winhttp-callback-and-handle-ownership.md`.

## Supersedes

- None

## Superseded by

- None
