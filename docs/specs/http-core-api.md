# HTTP core API

## Compatibility contract

The public VBA component names, public procedure names, parameter order, enum values, and documented error numbers form the compatibility surface. Before v1.0, incompatible changes require a CHANGELOG entry and migration note. At and after v1.0, existing public calls remain source-compatible within the documented Windows and Office matrix; additions are preferred over semantic changes.

All public objects are late-bindable from consumers except VBA class interfaces used with `Implements`. The core has no dependency on `Scripting.Dictionary` or another external VBA reference.

Workbook-reference consumers cannot directly instantiate VBA classes marked `PublicNotCreatable`. `VBAHttp.CreateClient()` is the stable default-client factory; `VBAHttp.CreateNativeClient()` selects the synchronous native WinHTTP backend. `CreateRetryPolicy`, `CreateExecutionOptions`, `CreateBatchOptions`, and `CreateCancellationToken` expose reliability configuration across the same boundary. Source-vendored consumers may continue to use `New`.

## Synchronous API

```vb
Dim client As New HttpClient
Dim response As HttpResponse

Set response = client.GetResponse("https://example.com/api")
Debug.Print response.StatusCode
Debug.Print response.Text
```

`HttpClient.Execute(request)` is the primitive buffered operation. `GetResponse`, `PostResponse`, `PutResponse`, `PatchResponse`, and `DeleteResponse` are convenience methods that construct a request and delegate to `Execute`. The `Response` suffix is required because VBA reserves `Get`, `Put`, and related tokens; this keeps ordinary early-bound calls compilable. Batch execution is defined by `bounded-concurrency.md`; constant-memory downloads are defined by `streaming-download.md` and are available on transports that implement `IHttpDownloadTransport`. Constant-memory file and multipart uploads are defined by `streaming-upload.md` and are available on transports that implement `IHttpUploadTransport`.

`HttpResponse.ProtocolUsed` reports the native backend's negotiated protocol as
`HTTP/1.1`, `HTTP/2`, or `HTTP/3`. The COM backend leaves this property empty.
`HttpProtocolOptions` can be assigned to `HttpClient` or `HttpRequest` for
native-only HTTP/2/HTTP/3 opt-in; its default has no override. The complete
fallback, required-mode, and compatibility contract is in
`protocol-policy.md`.

`HttpDecompressionOptions` can be assigned to `HttpClient` or `HttpRequest` for
native-only gzip/deflate response decoding. Its default has no override;
`decompression-policy.md` defines header ownership, fallback/required behavior,
and the streaming download length contract.

`HttpProxyOptions` can be assigned to `HttpClient.ProxyOptions` or
`HttpRequest.ProxyOptions`. The request value overrides the client default;
`proxy-policy.md` defines OS/default, direct, and manual routing plus the
credential boundary.

`IHttpAuthProvider` can be assigned to `HttpClient.AuthProvider` or
`HttpRequest.AuthProvider`. The request value overrides the client default and
the provider is cloned into the execution snapshot after default headers are
merged. The first release provides preemptive Basic and Bearer header
providers through `VBAHttp.CreateBasicAuthProvider` and
`VBAHttp.CreateBearerAuthProvider`. Providers reject an existing
`Authorization` header, require HTTPS by default, and disable automatic
redirects on the execution snapshot so a credential-bearing header is never
forwarded implicitly. `auth-policy.md` defines the explicit loopback-only
insecure opt-in, secret redaction, and the no-challenge-replay boundary.

`HttpClient.BaseUrl` resolves a relative request URL. An absolute request URL is used as-is. Base and relative URL joining must not discard path segments accidentally and is covered by URL unit tests before network transports are enabled.

`HttpClient.DefaultHeaders` are copied into an execution request. Request headers override the same case-insensitive default name. Execution snapshots isolate in-flight behavior from later mutations of the client or original request.

## Transport contract

`IHttpTransport.Execute(request As HttpRequest) As HttpResponse` performs one request. The request supplied to a transport is an execution snapshot owned by the client for the duration of the call. A transport must not retain or mutate caller-owned objects after it returns.

A successful exchange returns a non-`Nothing` response. Input and transport failures raise through `HttpErrors`. HTTP 4xx and 5xx are successful exchanges at this boundary.

## Request and body ownership

`HttpRequest` contains method, URL, headers, query parameters, timeouts, and an optional body. Callers may configure a request until execution starts. `HttpClient` clones all mutable domain objects before invoking the transport.

Buffered requests also expose `FollowRedirects` and `MaxRedirects`. They default to `True` and 10 respectively, are copied into the execution snapshot, and are applied by transports that support automatic redirects.

Core buffered bodies use `HttpBody` with one of three kinds: empty, text, or bytes.

- Text is copied into the request snapshot as a VBA `String`.
- Bytes are copied into a private `Byte()` owned by `HttpBody`; getters return a copy.
- A response owns its buffered body independently of request and transport objects.
- File and stream bodies are not represented by a giant `String` or `Byte()`. `DownloadFile` owns its bounded read buffer and same-directory temporary file as defined by `streaming-download.md`. `UploadFile` and `UploadMultipart` keep the caller-owned source read-only and emit bounded `WinHttpWriteData` chunks as defined by `streaming-upload.md`.

## Response text decoding

`HttpResponse` stores the authoritative buffered bytes and materializes `Text` lazily. Decoding precedence is:

1. explicit charset from the `Content-Type` header;
2. UTF-8 when no charset is declared;
3. deterministic failure for an unsupported or malformed charset.

The initial core implements UTF-8 and US-ASCII. Additional charsets are additive. Decoding must not silently fall back to the host ANSI code page because results would differ across machines. An explicitly supplied text response used by a mock is encoded to UTF-8 when converted to bytes.

## Headers and query parameters

`HttpHeaders` is case-insensitive for lookup and preserves insertion order and original casing for emission. `Add` appends a value, `SetValue` replaces every value for that name, `GetValue` returns the first value or an empty string, and `GetValues` returns a zero-based `String()` copy. Header names reject empty strings, control characters, colon, and surrounding whitespace. Values reject CR and LF.

`HttpParams` preserves insertion order and repeated names. Query encoding uses UTF-8 percent encoding. RFC 3986 unreserved bytes (`A-Z`, `a-z`, `0-9`, `-`, `.`, `_`, `~`) remain literal; spaces encode as `%20`, never `+`. Both names and values are encoded.

## Timeouts

`HttpTimeouts` exposes `ResolveMilliseconds`, `ConnectMilliseconds`, `SendMilliseconds`, and `ReceiveMilliseconds`. Values must be non-negative `Long` values. Zero means the backend-defined infinite timeout, matching WinHTTP. New requests receive independent default values of 5,000; 5,000; 30,000; and 300,000 milliseconds.

## Error boundary

The authoritative decision is ADR-0003. `HttpErrors` reserves the public namespace from `vbObjectError + 21000`. Core validation and status conversion use stable named constants. A response status outside 200 through 299 is returned normally; `RaiseForStatus` raises the HTTP-status category while retaining the status code in its sanitized description.

## Evidence

- Domain tests: `src/modules/Tests/Unit/HttpHeadersTests.bas`, `HttpParamsTests.bas`, `HttpBodyTests.bas`, `HttpResponseTests.bas`
- Transport/client tests: `src/modules/Tests/Unit/HttpClientTests.bas`
- Native transport tests: `src/modules/Tests/Unit/WinHttpNativeTests.bas` and `src/modules/Tests/Integration/WinHttpNativeTransportTests.bas`
- Architectural boundary: `docs/adr/ADR-0002-dual-transport-boundary.md`
- Error decision: `docs/adr/ADR-0003-http-error-model.md`
- Buffered backend: `docs/specs/buffered-com-transport.md`
- Batch API and scheduler contract: `docs/specs/bounded-concurrency.md`
- Retry, total deadline, and cancellation: `docs/specs/reliability-policy.md`
- Constant-memory download, publication, progress, and cleanup: `docs/specs/streaming-download.md`
- Constant-memory file/multipart upload, source ownership, progress, cancellation, and challenge handling: `docs/specs/streaming-upload.md`
