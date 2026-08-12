# HTTP benchmark methodology

## Scope

Benchmarks compare Raw `WinHttp.WinHttpRequest.5.1`, the pinned VBA-Web comparator, and VBA-HTTP against the same loopback test server. Results are engineering evidence, not universal performance claims.

## Controlled conditions

- External network access is prohibited.
- The server binds to an OS-selected loopback port and is reset before each implementation run.
- Each implementation uses the same HTTP method, path, payload size, timeout, warmup count, and measured iteration count.
- The Raw baseline creates a fresh request object per request. Comparator adapters must document any lifecycle they cannot match.
- Background workload should be minimized. A published comparison records Windows version, Excel version and bitness, xlflow version, CPU, and date outside the VBA-produced raw result.
- Compare only results collected in the same benchmark invocation or under demonstrably identical conditions.

## Phase 0 scenarios

| Scenario | Warmup | Measured iterations | Endpoint | Timeout | Metrics |
| --- | ---: | ---: | --- | --- | --- |
| Sequential latency | 5 | 50 | `GET /status/204` | resolve/connect 5s, send 30s, receive 300s | total, mean, min, max, requests/s |
| Buffered download | covered by prior warmup | 1 | `GET /bytes/104857600` | same | elapsed, bytes, MiB/s |

Time is measured with Windows `QueryPerformanceCounter`. Process working set, peak working set, private bytes, and handle count are captured immediately before and after each scenario. These snapshots detect persistent growth but are not a substitute for sampling peak transient allocation.

## Result storage

- Schema: `benchmarks/schema/benchmark-result.schema.json`
- Current machine-readable baseline: `benchmarks/results/raw-winhttp-baseline.json`
- Human summary: `docs/BENCHMARKS_BASELINE.md`

Raw result JSON uses schema version 1. Decimal numbers must use a period regardless of Windows locale. A run writes and validates a same-directory staging file before atomic replacement, so failure preserves the previous baseline. The xlflow structured error and test-server stderr are the diagnostic evidence for failed runs.
