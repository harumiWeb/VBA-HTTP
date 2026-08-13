# Buffered COM transport

## Backend and selection

`WinHttpComTransport` implements `IHttpTransport` with a late-bound
`WinHttp.WinHttpRequest.5.1` object. `HttpClient` creates this transport by
default, while `HttpClient.Transport` remains replaceable for tests and custom
backends. The production source has no compile-time reference to a WinHTTP type
library.

## Request mapping

- The execution snapshot's uppercase method and absolute URL are passed to
  `Open` in synchronous mode after any URI fragment is removed; fragments are
  client-only state and never enter the HTTP request-target.
- Resolve, connect, send, and receive timeouts are passed to `SetTimeouts` in
  milliseconds. Zero retains WinHTTP's infinite-timeout meaning.
- Headers are emitted in insertion order. Repeated header names are emitted as
  repeated `SetRequestHeader` calls.
- Empty bodies call `Send` without an argument. Text bodies are encoded as
  UTF-8 bytes by `HttpBody`; binary bodies retain their exact bytes.
- `HttpRequest.FollowRedirects` defaults to `True` and
  `HttpRequest.MaxRedirects` defaults to 10. The transport applies both WinHTTP
  options before sending. `EnableHttpsToHttpRedirects` is explicitly set to
  `False`. Requests containing `Authorization`, `Proxy-Authorization`, or
  `Cookie` disable automatic redirects and return the 3xx response.
- The execution snapshot may contain a preemptive Basic or Bearer
  `IHttpAuthProvider`; the client has already materialized its `Authorization`
  header before `Open`/`Send`. Such snapshots always set `FollowRedirects` to
  `False`. The COM backend does not perform 401/407 challenge replay.

## Response mapping

Every completed exchange, including 4xx and 5xx, returns `HttpResponse`.
`Status`, `StatusText`, all response headers, authoritative `ResponseBody`
bytes, and elapsed milliseconds are copied before the COM object is released.
An empty response remains `HttpBodyEmpty`; it is not represented as a one-byte
or text body.

Loopback integration additionally proves that direct binary GET responses keep
their exact bytes, that UTF-8 text containing supplementary Unicode code points
is decoded deterministically, and that malformed UTF-8 is rejected when text is
requested. The raw `/malformed-headers` fixture must fail as
`HttpErrorProtocol`, whether rejection comes from WinHTTP or the transport's
header parser.

Header parsing splits each raw header line at its first colon, preserving
repeated fields and the field-value text after trimming optional whitespace.
Malformed response header lines fail as `HttpErrProtocol` rather than being
silently discarded.

## Transport failures

Raw COM/HRESULT values are not public error numbers. The low WinHTTP error code
is classified into the stable `HttpErrors` namespace:

| WinHTTP condition | Public category |
| --- | --- |
| invalid URL or scheme | `HttpErrorInvalidUrl` |
| name not resolved | `HttpErrorDns` |
| connect/connection failure | `HttpErrorConnection` |
| certificate or secure-channel failure | `HttpErrorTls` |
| timeout | `HttpErrorTimeout` |
| operation cancelled | `HttpErrorCancelled` |
| invalid response, header, or redirect | `HttpErrorProtocol` |
| other backend failure | `HttpErrorIo` |

Descriptions contain only a stable summary and numeric WinHTTP code. They do
not include the request URL, headers, credentials, response body, or the raw COM
description. Failures raised by VBA-HTTP before the COM call are propagated
unchanged.

Before releasing a backend after `Open`, `Send`, timeout, or response failure,
the transport makes a best-effort `Abort` call. The original mapped error is
preserved if that cleanup call fails. A 250 ms bounded drain follows a
synchronous failure and asynchronous cancellation so WinHTTP can finish
COM-handle teardown. The drain applies only to abort paths; repeated
receive-timeout stability is verified by the canonical Phase 9 stress result
on the current x64 host and must be rerun for each supported Office host.

## Scope

This transport is buffered and synchronous. Bounded asynchronous scheduling is
defined in Phase 3. Constant-memory download and upload use the native
transport defined by `streaming-download.md` and `streaming-upload.md`;
advanced protocol selection remains a later phase.

## Evidence

- Integration suite: `src/modules/Tests/Integration/WinHttpComTransportTests.bas`
- Repeated timeout evidence: `benchmarks/results/phase9-cancellation-stress.json`
- Unit suite: `src/modules/Tests/Unit/HttpRequestTests.bas`
- Loopback runner: `tools/Run-IntegrationTests.ps1`
