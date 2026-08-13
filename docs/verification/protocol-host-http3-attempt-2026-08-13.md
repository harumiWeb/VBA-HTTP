# HTTP/3 host validation attempt — 2026-08-13

## Scope

This record documents three operator-invoked attempts to obtain the pending
HTTP/3/QUIC negotiated-host evidence. It is a failure/inconclusive record,
not a compatibility claim. No `protocol-host-http3.json` passing artifact was
created.

## Runs

| Endpoint | Request | Result |
| --- | --- | --- |
| `nghttp2.org:4433` | HTTPS, `HTTP/3`, `required`, 90-second outer deadline | WinHTTP consumer terminated with RPC failure `0x800706BE`; no `ProtocolUsed` was observed |
| `quic.rocks:4433` | HTTPS, `HTTP/3`, `required`, 90-second outer deadline | WinHTTP consumer terminated with RPC failure `0x800706BE`; no `ProtocolUsed` was observed |
| `cloudflare-quic.com:443` | HTTPS, `HTTP/3`, `required`, 90-second outer deadline | WinHTTP consumer terminated with RPC failure `0x800706BE`; no `ProtocolUsed` was observed |

These endpoints are public HTTP/3 test-server candidates, but endpoint
availability alone is not evidence that this Windows/WinHTTP host negotiated
QUIC. The runner therefore remains fail-closed and leaves the HTTP/3 matrix
row pending.

## Safety and cleanup

- The run started with no Excel process.
- `Run-ProtocolHostValidation.ps1` recorded the exact runner-owned Excel PID
  in `Watch-ProtocolHostExcel.ps1`; only that PID was eligible for watchdog
  cleanup.
- The run ended with no Excel process, no xlflow recovery state, and no
  `protocol-host-http3.json` output.
- No pre-existing or unrelated Excel process was stopped.
- The third run used the same ownership and cleanup checks; the owned Excel
  PID was `77108` and the watchdog PID was `57040`.

## Follow-up

Repeat the same runner on a Windows/WinHTTP host with confirmed UDP/QUIC
egress and a trusted HTTP/3 endpoint. A result is promotable only when the
consumer returns `ProtocolUsed=HTTP/3`; a timeout, RPC failure, fallback, or
unsupported option must remain non-evidence.
