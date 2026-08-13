# HTTP/3 host validation attempt — 2026-08-13

## Scope

This record documents five operator-invoked attempts to obtain the pending
HTTP/3/QUIC negotiated-host evidence. It is a failure/inconclusive record,
not a compatibility claim. No `protocol-host-http3.json` passing artifact was
created.

## Runs

| Endpoint | Request | Result |
| --- | --- | --- |
| `nghttp2.org:4433` | HTTPS, `HTTP/3`, `required`, 90-second outer deadline | WinHTTP consumer terminated with RPC failure `0x800706BE`; no `ProtocolUsed` was observed |
| `quic.rocks:4433` | HTTPS, `HTTP/3`, `required`, 90-second outer deadline | WinHTTP consumer terminated with RPC failure `0x800706BE`; no `ProtocolUsed` was observed |
| `cloudflare-quic.com:443` | HTTPS, `HTTP/3`, `required`, 90-second outer deadline | WinHTTP consumer terminated with RPC failure `0x800706BE`; no `ProtocolUsed` was observed |
| `quic.aiortc.org:443` | HTTPS, `HTTP/3`, `required`, 90-second outer deadline | WinHTTP consumer terminated with RPC failure `0x800706BE`; no `ProtocolUsed` was observed |
| `cloudflare-quic.com:443` | HTTPS, `HTTP/3`, `required`, 60-second outer deadline after the native DWORD-option marshalling fix | WinHTTP consumer terminated with RPC failure `0x800706BE`; no `ProtocolUsed` was observed |

These endpoints are public HTTP/3 test-server candidates, but endpoint
availability alone is not evidence that this Windows/WinHTTP host negotiated
QUIC. The runner therefore remains fail-closed and leaves the HTTP/3 matrix
row pending.

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

## Excel-free capability diagnosis

The direct x64 probe was run after the fifth consumer attempt to separate the
WinHTTP capability from Excel automation:

| Probe | Result |
| --- | --- |
| `cloudflare-quic.com:443`, mask `HTTP/3`, option 133 | `ERROR_NOT_SUPPORTED` (50) before send; no request was issued |
| `nghttp2.org:443`, mask `HTTP/2`, option 133 | option accepted and option 134 returned HTTP/2 |

This confirms that the current WinHTTP runtime exposes HTTP/2 but rejects an
HTTP/3-enabled mask. The native transport now treats error 50 as an explicit
capability miss: allow-fallback skips the option, while required mode raises
`HttpErrorProtocol`. The probe is diagnostic evidence only and does not promote
HTTP/3.

## Follow-up

Repeat the same runner on a Windows/WinHTTP host with confirmed UDP/QUIC
egress and a trusted HTTP/3 endpoint. A result is promotable only when the
consumer returns `ProtocolUsed=HTTP/3`; a timeout, RPC failure, fallback, or
unsupported option must remain non-evidence.
