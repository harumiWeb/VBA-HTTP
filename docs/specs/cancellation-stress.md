# Cancellation and timeout stress specification

## Scope

`task test:cancellation-stress` runs four loopback-only scenarios against the
deterministic Go test server. The stress module lives under
`src/modules/Tests/Stress/` and is excluded from release workbooks.

| Scenario | Backend | Workload | Required invariant | Handle gate |
| --- | --- | --- | --- | ---: |
| `com_active_cancellation` | `WinHttpComTransport` | four `GET /delay/2000` requests, cancelled after one second | every item is cancelled; a recovery `GET /status/204` succeeds | <=32 |
| `com_request_deadline` | `WinHttpComTransport` | four `GET /delay/250` requests, 25 ms per-request deadline | every item is `HttpErrorTimeout`; recovery succeeds | <=32 |
| `com_receive_timeout` | `WinHttpComTransport` | repeated synchronous `GET /delay/10000` requests with a 1,000 ms receive timeout | every attempt maps to `HttpErrorTimeout`; recovery succeeds | <=32 |
| `native_download_cancellation` | `WinHttpNativeTransport` | 64 KiB `GET /stream/65536`, progress cancellation at 64 KiB | `HttpErrCancelled`, 8-byte sentinel and temp-file count unchanged; recovery succeeds | <=8 |

Each selected scenario runs exactly the configured number of iterations
(default 25, allowed range 1..1000), uses a fresh cancellation token/options for
each cycle, and sets `MaxAttempts=1` where retries could obscure the result.
The COM active-cancel scenario uses the existing cooperative `DoEvents` polling
boundary and does not introduce a native callback.

The receive-timeout scenario exercises the synchronous COM failure path. Its
bounded 250 ms abort drain is part of ADR-0022; the scenario is the evidence
gate for repeated timeout cleanup rather than a claim that WinHTTP has no
host-specific allocator variation.

## Process evidence

The runner records Excel PIDs present before the complete gate and samples only
new processes across every scenario. Each stress method performs a one-second
post-warmup settling pause after the start marker so the first measured
operation is separated from Excel startup/COM allocation. Existing Excel
processes are baseline-only and are never terminated;
the runner does not assume ownership of an unrelated user session. VBA writes
`start`, `done`, and waits for `release` markers. The runner captures process
handles, working set, and private bytes before the measured loop, at peak, one
second after completion, and at one-second intervals for up to 30 seconds while
COM teardown settles. A missing marker, lost process, failed test, or exceeded
idle handle limit fails closed. Memory values are evidence for same-host
comparison, not a universal pass threshold. Do not start another Excel process
during the gate: a concurrently launched process is intentionally treated as
unowned and can make the PID-scoped observation inconclusive.

## Result contract

The runner atomically publishes
`benchmarks/results/phase9-cancellation-stress.json`. It must have
`external_network=false`, a loopback base URL, one passing result per selected
scenario, exact iteration counts, scenario-specific transport and handle
limits, and PID-scoped snapshots. Validate it with
`benchmarks/schema/cancellation-stress-result.schema.json` and
`tools/Validate-CancellationStressResult.ps1`.

Native `TotalDeadlineMilliseconds` is not used to assert interruption of a
blocking buffered receive; that capability boundary remains documented by the
native transport ADR. Cookie state and credentials are not emitted in results.
