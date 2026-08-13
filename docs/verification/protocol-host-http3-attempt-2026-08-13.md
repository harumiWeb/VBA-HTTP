# HTTP/3 host validation attempt — 2026-08-13

## Scope

This record documents six operator-invoked attempts to investigate HTTP/3/QUIC
on the then-current host. It is historical diagnostic evidence, not a
compatibility claim. No `protocol-host-http3.json` passing artifact was
created, and ADR-0035 now classifies HTTP/3/QUIC as unsupported by policy.

## Runs

| Endpoint | Request | Result |
| --- | --- | --- |
| `nghttp2.org:4433` | HTTPS, `HTTP/3`, `required`, 90-second outer deadline | WinHTTP consumer terminated with RPC failure `0x800706BE`; no `ProtocolUsed` was observed |
| `quic.rocks:4433` | HTTPS, `HTTP/3`, `required`, 90-second outer deadline | WinHTTP consumer terminated with RPC failure `0x800706BE`; no `ProtocolUsed` was observed |
| `cloudflare-quic.com:443` | HTTPS, `HTTP/3`, `required`, 90-second outer deadline | WinHTTP consumer terminated with RPC failure `0x800706BE`; no `ProtocolUsed` was observed |
| `quic.aiortc.org:443` | HTTPS, `HTTP/3`, `required`, 90-second outer deadline | WinHTTP consumer terminated with RPC failure `0x800706BE`; no `ProtocolUsed` was observed |
| `cloudflare-quic.com:443` | HTTPS, `HTTP/3`, `required`, 60-second outer deadline after the native DWORD-option marshalling fix | WinHTTP consumer terminated with RPC failure `0x800706BE`; no `ProtocolUsed` was observed |
| `cloudflare-quic.com:443` | HTTPS, `HTTP/3`, `required`, 45-second outer deadline after explicit pointer-sized `WinHttpSetOption` marshalling | WinHTTP consumer terminated with RPC failure `0x800706BE`; Windows Application Error 1000 recorded an `EXCEL.EXE` access violation in `combase.dll` (`0xc0000005`); no `ProtocolUsed` was observed |

These endpoints are public HTTP/3 test-server candidates, but endpoint
availability alone is not evidence that this Windows/WinHTTP host negotiated
QUIC. The runner therefore remained fail-closed; the HTTP/3 matrix row is now
`unsupported-by-policy` under ADR-0035 rather than a pending promotion gate.

## Safety and cleanup

- The run started with no Excel process.
- `Run-ProtocolHostValidation.ps1` records the exact runner-owned Excel PID by
  mapping the new `Excel.Application.Hwnd` to its process; the watchdog and
  teardown can target only that proven PID.
- The run ended with no Excel process, no xlflow recovery state, and no
  `protocol-host-http3.json` output.
- No pre-existing or unrelated Excel process was stopped.
- The third run used the same ownership and cleanup checks; the owned Excel
  PID was `77108` and the watchdog PID was `57040`.
- The fourth run used the HWND-to-PID ownership guard; the watchdog targeted
  only its runner-owned Excel PID and the run ended without a passing artifact.
- The fifth run exercised the release artifact containing the explicit
  32-bit `WinHttpSetOption` buffer fix. The RPC failure remained; the
  runner-owned Excel process exited during teardown and no unrelated Excel
  process was stopped.
- The sixth run exercised the subsequent explicit `VarPtr(optionBuffer)` ABI
  fix. It reproduced the same RPC failure before a response was returned.
  Windows Error Reporting identified the faulting module as `combase.dll`
  (`EXCEL.EXE` exception `0xc0000005`). The ownership guard still targeted
  only the runner-owned Excel PID; no unrelated Excel process was stopped.

## Excel-free capability diagnosis

The direct x64 probe was run after the fifth consumer attempt to separate the
WinHTTP capability from Excel automation:

| Probe | Result |
| --- | --- |
| `cloudflare-quic.com:443`, mask `HTTP/3`, option 133 | `ERROR_NOT_SUPPORTED` (50) before send; no request was issued |
| `nghttp2.org:443`, mask `HTTP/2`, option 133 | option accepted and option 134 returned HTTP/2 |
| `nghttp2.org:443`, option 184 feature probe for options 133/145/118 | `ERROR_INVALID_PARAMETER` (87) for each query; no reliable preflight result |

This confirms that the current WinHTTP runtime exposes HTTP/2 but rejects an
HTTP/3-enabled mask. The native transport now treats error 50 as an explicit
capability miss: allow-fallback skips the option, while required mode raises
`HttpErrorProtocol`. The option-184 query did not provide a usable capability
signal on this host, so it is not used as a production preflight. The probe is
diagnostic evidence only and does not promote HTTP/3.

## Follow-up

The protocol-host runner now performs an Excel-free required-mask preflight
before creating Excel (ADR-0032). Re-running the same HTTP/3 candidate on this
host therefore stops at `WinHttpSetOption(133)` with `ERROR_NOT_SUPPORTED (50)`
and publishes no evidence; it does not repeat the earlier Excel automation
crash. The x64 HTTP/2 control run was refreshed with the passing preflight and
public-consumer result in `benchmarks/results/protocol-host-http2.json`.

If HTTP/3 support is reconsidered, a superseding ADR must first restore an
explicit promotion path on a Windows/WinHTTP host with confirmed UDP/QUIC
egress and a trusted HTTP/3 endpoint. Until then, HTTP/3 probes remain
diagnostic-only and no `ProtocolUsed=HTTP/3` result is a release claim.
