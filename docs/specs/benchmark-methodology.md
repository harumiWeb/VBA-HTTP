# HTTP benchmark methodology

## Scope

Benchmarks compare Raw `WinHttp.WinHttpRequest.5.1`, the pinned VBA-Web comparator, and VBA-HTTP against the same loopback test server. Results are engineering evidence, not universal performance claims.

## Controlled conditions

- External network access is prohibited.
- The server binds to an OS-selected loopback port and is reset before each implementation run.
- Each implementation uses the same HTTP method, path, payload size, warmup count, and measured iteration count.
- The Raw baseline creates a fresh request object per request. The VBA-Web adapter likewise creates a fresh `WebClient`, `WebRequest`, and `WebResponse` lifecycle per request.
- Raw WinHttpRequest uses phase timeouts of 5s resolve, 5s connect, 30s send, and 300s receive. VBA-Web v4.1.6 exposes only one `TimeoutMs` value for all four phases, so its adapter uses 300s to preserve the shared receive deadline. This API limitation is recorded in its result JSON and must be considered when comparing failure behavior; it does not alter the loopback success-path workload.
- Background workload should be minimized. A published comparison records Windows version, Excel version and bitness, xlflow version, CPU, and date outside the VBA-produced raw result.
- Compare only results collected in the same benchmark invocation or under demonstrably identical conditions.

## Phase 0 scenarios

| Scenario | Warmup | Measured iterations | Endpoint | Timeout | Metrics |
| --- | ---: | ---: | --- | --- | --- |
| Sequential latency | 5 | 50 | `GET /status/204` | resolve/connect 5s, send 30s, receive 300s | total, mean, min, max, requests/s |
| Buffered download | covered by prior warmup | 1 | `GET /bytes/104857600` | same | elapsed, bytes, MiB/s |

Time is measured with Windows `QueryPerformanceCounter`. Process working set, peak working set, private bytes, and handle count are captured immediately before and after each scenario. These snapshots detect persistent growth but are not a substitute for sampling peak transient allocation.

## Phase 2 comparison

`task benchmark:phase2` runs Raw WinHttpRequest and VBA-HTTP from synchronized copies of the same development workbook, against fresh instances of the same loopback server, within one top-level task invocation. Both paths use the Phase 0 parameters above. The Raw path and VBA-HTTP path share the measurement, JSON serialization, and process-snapshot code; only the HTTP invocation differs.

`benchmarks/results/phase2-buffered-overhead.json` follows `benchmarks/schema/benchmark-comparison.schema.json`, compares sequential GET mean latency, and records the 15% engineering target. Exceeding the target does not make the evidence disappear: the comparison records `within_target: false`, and the human baseline must explain the absolute difference and the cost of the documented ownership/API guarantees. The comparison is a local engineering result, not a universal performance claim.

## VBA-Web comparator provenance and isolation

- Upstream: `https://github.com/VBA-tools/VBA-Web.git`
- Tag: `v4.1.6`
- Commit: `cefc320acc5372e0b86eed1d20eb3f31b331d598`
- License: MIT, retained in the pristine checkout
- Input workbook: upstream `VBA-Web - Blank.xlsm`, SHA-256 `4f0cd9c4338951708fb80018fbca63cf6c38f1ae202b9ebd595ce6f88e892891`

`task benchmark:vba-web:setup` is the only acquisition step and may access GitHub when the ignored reference checkout is absent. Normal benchmark execution uses `Setup-VBAWeb.ps1 -VerifyOnly`, fails if the exact pristine commit is unavailable, and performs no network access other than the loopback server. The tracked adapter is imported into an ignored temporary copy of the upstream workbook, compiled and run by xlflow, then discarded. Upstream specs are never imported or executed because several depend on public services such as httpbin and Facebook. The Raw runner likewise executes a temporary copy of the synchronized development workbook, so neither benchmark mutates its base workbook.

## Result storage

- Schema: `benchmarks/schema/benchmark-result.schema.json`
- Current machine-readable baselines: `benchmarks/results/raw-winhttp-baseline.json` and `benchmarks/results/vba-web-baseline.json`
- Phase 2 evidence: `benchmarks/results/phase2-raw-winhttp.json`, `vba-http-buffered.json`, and `phase2-buffered-overhead.json`
- Human summary: `docs/BENCHMARKS_BASELINE.md`

Result JSON uses schema version 1. Decimal numbers must use a period regardless of Windows locale. Each run writes and validates a same-directory staging file before atomic replacement, so failure preserves the previous baseline. The xlflow structured error and test-server stderr are the diagnostic evidence for failed runs. Run `task benchmark:phase0` to refresh both implementations in one top-level invocation.
