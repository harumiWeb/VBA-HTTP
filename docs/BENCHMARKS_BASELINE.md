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
