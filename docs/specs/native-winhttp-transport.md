# Native WinHTTP transport

## Scope

`WinHttpNativeTransport` is the native backend for buffered and streaming
operations. It is Windows-only, late-bound to `winhttp.dll`, and implements
`IHttpTransport`.
The default `HttpClient` transport remains `WinHttpComTransport`; consumers
select the native backend explicitly by assigning it to `HttpClient.Transport`.

## Call sequence

For every request the transport performs this sequence on the VBA thread:

1. parse the absolute HTTP or HTTPS URL into scheme, host, port, and origin
   path (fragments are not sent);
2. open a WinHTTP session using the default OS proxy configuration;
3. connect to the host and open a request handle;
4. apply timeout and redirect options and the optional native protocol policy;
5. add validated request headers, then send the buffered or streaming body with
   its exact byte count;
6. receive the response, query status and raw CRLF headers, and read all body
   chunks using `WinHttpQueryDataAvailable`/`WinHttpReadData`;
7. query `WINHTTP_OPTION_HTTP_PROTOCOL_USED` when supported and close all
   handles in reverse hierarchy order.

No callback is installed. A call that needs active cancellation or a total
deadline is rejected by the reliability engine because this synchronous
transport cannot observe those controls while WinHTTP is blocked.

## Handle ownership

`WinHttpSessionHandle`, `WinHttpConnectionHandle`, and `WinHttpRequestHandle`
each own one `HINTERNET`. A child retains its parent wrapper. `ReleaseHandle` is
idempotent; explicit transport cleanup is the normal path and `Class_Terminate`
is only a safety net. Handle values are pointer-sized (`LongPtr` under VBA7,
`Long` under the legacy conditional branch) and are never exposed through the
public HTTP API.

## 32/64-bit declarations

All declarations use `PtrSafe` and `LongPtr` under `VBA7`; the `#Else` branch
uses the 32-bit `Long` ABI. Pointer arguments are passed as `StrPtr`/`VarPtr`.
The project must record a real 32-bit Office compile on a dedicated host before
claiming the Phase 5 32-bit exit criterion. The current development evidence
is x64 only; static conditional compilation is not substituted for that
evidence.

## Request and response contract

- Supported methods and URL validation are inherited from `HttpRequest` and
  `HttpClient`.
- Request headers retain insertion order and repeated names.
- Response status, reason phrase, repeated headers, exact binary body, and
  elapsed time match `WinHttpComTransport`.
- HTTP 4xx/5xx responses are returned, not raised. Native failures are mapped
  by the shared WinHTTP error classifier to the ADR-0003 categories; raw URLs,
  headers, credentials, and response data never enter descriptions.
- TLS uses OS certificate validation. No certificate-ignore flag is exposed.
- Protocol flags are queried opportunistically. `HttpProtocolOptions` can
  enable HTTP/2/3 and can require the selected mask; allow-fallback treats an
  unsupported native option as a capability miss, while required mode maps
  unsupported options and protocol mismatch to `HttpErrorProtocol`. `HTTP/2`
  and `HTTP/3` are reported only when WinHTTP returns the corresponding flag;
  otherwise the response reports the negotiated legacy protocol as `HTTP/1.1`.
- Advanced protocol flags are applied only to HTTPS URLs. The default has no
  override, and plain HTTP fallback remains HTTP/1.1.

## Resource regression gate

The native integration suite warms up the backend, repeats complete requests,
and samples the Excel process handle count before and after. A small bounded
OS fluctuation is tolerated, but persistent growth beyond the documented
threshold fails the test. The test does not inspect private handle values.

## Evidence

- `src/classes/WinHttpNativeTransport.cls`
- `src/modules/WinHttpNativeApi.bas`
- `src/classes/WinHttp*Handle.cls`
- `src/modules/Tests/Integration/WinHttpNativeTransportTests.bas`
- `docs/verification/` records the host-specific bitness result.
