# ADR-0009: Native HTTP protocol negotiation policy

## Status

`accepted`

## Background

WinHTTP can enable HTTP/2 and HTTP/3 per request, report the negotiated
protocol, and optionally require that an enabled protocol be used. These
options are not exposed by the late-bound `WinHttp.WinHttpRequest.5.1` COM
object used by the default transport. The library therefore needs a
transport-specific policy that does not silently change the COM backend or
claim HTTP/2/3 support on hosts where WinHTTP lacks the capability. ADR-0035
later establishes that HTTP/3/QUIC is unsupported by policy for the current
distribution; the implementation below remains diagnostic/future-compatible.

## Decision

- Add `HttpProtocolOptions` to `HttpClient` and `HttpRequest`. The default has
  no override, so the operating system/WinHTTP default negotiation remains
  unchanged. `AllowHttp2` and `AllowHttp3` set the native WinHTTP protocol bit
  mask; `Mode = HttpProtocolAllowFallback` is the safe default when a mask is
  configured.
- `HttpProtocolRequired` is valid only with HTTP/2 or HTTP/3 enabled. Native
  requests set `WINHTTP_OPTION_ENABLE_HTTP_PROTOCOL` (133), and required
  requests additionally set `WINHTTP_OPTION_HTTP_PROTOCOL_REQUIRED` (145)
  before sending. A capability miss in allow-fallback mode falls back to the
  legacy protocol; a capability miss or negotiation mismatch in required mode
  raises the stable `HttpErrorProtocol` category. Both WinHTTP error 12009
  (`ERROR_WINHTTP_INVALID_OPTION`) and Win32 error 50 (`ERROR_NOT_SUPPORTED`)
  are capability misses, because current Windows builds can accept HTTP/2
  while rejecting an HTTP/3-enabled mask with error 50.
- HTTP/2 and HTTP/3 options require an HTTPS URL. Allow-fallback requests over
  plain HTTP skip the advanced option and use HTTP/1.1; required requests fail
  before sending.
- The COM transport rejects a configured advanced protocol mask before network
  I/O. It does not pretend that its legacy `WinHttpRequestOption` values are
  HTTP/2 or HTTP/3 controls.
- `HttpResponse.ProtocolUsed` remains the only negotiated-protocol evidence.
  The library does not infer HTTP/2 or HTTP/3 from requested flags. HTTP/2 host
  evidence remains part of the supported OS compatibility matrix; HTTP/3
  evidence is diagnostic-only under ADR-0035.
- Correct the Phase 7 upload length encoding: `dwTotalLength` is a DWORD
  bit-pattern for lengths below 4 GiB and uses
  `WINHTTP_IGNORE_REQUEST_TOTAL_LENGTH = 0` only for lengths at or above 4
  GiB when the explicit `Content-Length` header carries the full value.

## Consequences

- Consumers can opt into modern protocol negotiation without making the
  default client behavior host-dependent or breaking COM-only deployments.
- Required mode is intentionally strict and may fail on older Windows/WinHTTP
  versions or endpoints that cannot negotiate the requested protocol.
- Plain loopback HTTP remains a deterministic fallback test; a TLS-enabled
  HTTP/2 fixture is required before claiming negotiated HTTP/2 evidence for a
  release matrix. HTTP/3/QUIC is not a current release claim.
- Protocol options are not coupled to proxy credentials, authentication
  replay, or compression. Those policies remain separate Phase 8 decisions.

## Rationale

- Code: `src/classes/HttpProtocolOptions.cls`, `HttpRequest.cls`,
  `HttpClient.cls`, `WinHttpNativeTransport.cls`, `WinHttpNativeApi.bas`, and
  `WinHttpComTransport.cls`.
- Tests: `src/modules/Tests/Unit/HttpProtocolOptionsTests.bas`, native API
  mapping/length tests, request/client snapshot tests, and loopback native
  integration fallback/required tests.
- Current contract: `docs/specs/protocol-policy.md` and
  `docs/specs/native-winhttp-transport.md`.
- WinHTTP option behavior: [option flags](https://learn.microsoft.com/en-us/windows/win32/winhttp/option-flags)
  and [WinHttpSendRequest](https://learn.microsoft.com/en-us/windows/win32/api/winhttp/nf-winhttp-winhttpsendrequest).

## Supersedes

- None

## Superseded by

- None
