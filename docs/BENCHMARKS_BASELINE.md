# Benchmark baseline

Recorded: 2026-08-12

This is the initial Raw WinHttpRequest baseline for the Phase 0 loopback harness. It is not yet a VBA-Web or VBA-HTTP comparison and must not be presented as a general network-performance claim.

## Environment

- Windows 11 Home 10.0.22631 (build 22631)
- Microsoft Excel 16.0 x64
- 12th Gen Intel Core i7-12700, 12 cores / 20 logical processors
- xlflow development build (`commit: none`)
- Go local test server on an OS-selected `127.0.0.1` port

## Raw WinHttpRequest 5.1

| Scenario | Conditions | Result |
| --- | --- | ---: |
| Sequential GET | 5 warmups, 50 measured `GET /status/204` requests | 0.716 ms mean; 0.596 ms min; 1.048 ms max; 1,396.270 requests/s |
| Buffered download | One `GET /bytes/104857600` | 972.464 ms; 102.832 MiB/s; 100 MiB received |

The sequential scenario ended with two additional handles (1,467 to 1,469), and the buffered download ended with another two (1,469 to 1,471). Download peak working set rose from 159,825,920 to 474,439,680 bytes while the final working set returned to 159,854,592 bytes. These immediate snapshots are baseline signals, not proof of a persistent leak or precise peak allocation.

Authoritative machine-readable values are in `benchmarks/results/raw-winhttp-baseline.json`. Methodology and limitations are in `docs/specs/benchmark-methodology.md`.
