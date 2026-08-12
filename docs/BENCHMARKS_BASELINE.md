# Benchmark baseline

Recorded: 2026-08-12

This is the Phase 0 Raw WinHttpRequest and pinned VBA-Web comparison on the deterministic loopback harness. It is not yet a VBA-HTTP result and must not be presented as a general network-performance claim.

## Environment

- Windows 11 Home 10.0.22631 (build 22631)
- Microsoft Excel 16.0 x64
- 12th Gen Intel Core i7-12700, 12 cores / 20 logical processors
- xlflow development build (`commit: none`)
- Go local test server on an OS-selected `127.0.0.1` port

## Raw WinHttpRequest 5.1

| Scenario | Conditions | Result |
| --- | --- | ---: |
| Sequential GET | 5 warmups, 50 measured `GET /status/204` requests | 0.557 ms mean; 0.466 ms min; 0.705 ms max; 1,795.300 requests/s |
| Buffered download | One `GET /bytes/104857600` | 539.808 ms; 185.251 MiB/s; 100 MiB received |

The sequential scenario retained no additional handles (1,475 to 1,475), and the buffered download ended with two additional handles (1,475 to 1,477). Download peak working set rose from 160,399,360 to 475,009,024 bytes while the final working set returned to 160,423,936 bytes.

## VBA-Web 4.1.6

The comparator is the pristine upstream `VBA-Web - Blank.xlsm` at commit `cefc320acc5372e0b86eed1d20eb3f31b331d598`. xlflow compiled and ran only the tracked loopback adapter; upstream external-network specs were not executed.

| Scenario | Conditions | Result |
| --- | --- | ---: |
| Sequential GET | 5 warmups, 50 measured `GET /status/204` requests | 2.833 ms mean; 2.301 ms min; 5.391 ms max; 352.971 requests/s |
| Buffered download | One `GET /bytes/104857600` | 496.154 ms; 201.550 MiB/s; 100 MiB received |

VBA-Web's sequential latency was 5.09 times the Raw mean in this invocation. Its sequential scenario ended with eight additional handles (1,451 to 1,459); the download retained no additional handles. Download peak working set rose from 161,054,720 to 895,119,360 bytes while the final working set returned to 161,067,008 bytes. The apparent download throughput difference is within a single local run and is not a general performance claim.

These immediate snapshots are baseline signals, not proof of a persistent leak or precise peak allocation. VBA-Web exposes one timeout value for all request phases; its 300-second setting preserves the Raw receive deadline but differs from Raw's stricter resolve, connect, and send settings.

Authoritative machine-readable values are in `benchmarks/results/raw-winhttp-baseline.json` and `benchmarks/results/vba-web-baseline.json`. Methodology, provenance, setup, and limitations are in `docs/specs/benchmark-methodology.md`.

## Phase 2 buffered VBA-HTTP evidence

The Phase 2 matched run used the same 5 warmups, 50 measured `GET /status/204` requests, timeouts, and 100 MiB payload for Raw WinHttpRequest and VBA-HTTP.

| Implementation | Sequential mean | Requests/s | 100 MiB download | Throughput |
| --- | ---: | ---: | ---: | ---: |
| Raw WinHttpRequest 5.1 | 0.499 ms | 2,003.366 | 370.777 ms | 269.704 MiB/s |
| VBA-HTTP 0.2-dev | 0.833 ms | 1,199.849 | 797.747 ms | 125.353 MiB/s |

VBA-HTTP's sequential mean was 66.934% above Raw, so it does not meet the 15% relative target. The absolute mean difference was 0.334 ms on this loopback host. The measured path includes request snapshotting, URL/default-header processing, stable error mapping, response header/body ownership, and domain-object construction that Raw does not provide. Removing these guarantees to optimize a sub-millisecond loopback case would violate the public contract, so the variance is accepted and recorded for Phase 2.

Before the final implementation, a VBA per-byte defensive-copy loop limited the 100 MiB path to 12.845 MiB/s. Replacing only Byte-array copying with VBA's defensive SAFEARRAY assignment retained the ownership regression test and raised the measured result to 125.353 MiB/s. Buffered bodies still require multiple in-memory copies by contract; constant-memory transfer belongs to the native streaming phases.

The machine-readable comparison is `benchmarks/results/phase2-buffered-overhead.json`; its referenced raw results are the authoritative values for this run.

## Phase 3 bounded-concurrency evidence

The Phase 3 loopback run issued 100 requests to `GET /delay/100` through the same VBA-HTTP COM transport. A sequential limit of 1 took 11,041.019 ms; a concurrency limit of 16 took 858.266 ms, a 12.864x speedup. Server-side instrumentation observed a maximum of exactly 16 simultaneous requests, so the configured bound was not exceeded.

This measures cooperative async overlap on one local machine, not VBA multithreading or public-network throughput. The machine-readable result is `benchmarks/results/vba-http-concurrency.json` and follows `benchmarks/schema/concurrency-result.schema.json`.

## Phase 6 constant-memory download evidence

The Phase 6 stress run used the deterministic loopback server's chunked
`GET /stream/1073741824` endpoint. The harness performs a small native warm-up,
then gates the 1 GiB transfer so Excel startup is included in neither the
baseline nor the peak delta. This is an engineering memory signal, not a claim
about all Excel hosts.

| Measure | Result |
| --- | ---: |
| Payload | 1,073,741,824 bytes |
| Elapsed | 44,904.709 ms |
| Working-set peak delta | 6,836,224 bytes |
| Private-bytes peak delta | 19,017,728 bytes |
| SHA-256 | `3a33d58aa8ee1e9d21fd4f510cc5d1ce8d25ba5e24363d19c28e2bf866f4185c` |

The byte-for-byte destination hash matched the server's `/sha256/` response,
and integration tests verified cancellation, atomic replacement, and temporary
file cleanup. The machine-readable result is
`benchmarks/results/phase6-download-stress.json`.
