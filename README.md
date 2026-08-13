# VBA-HTTP

VBA-HTTP is a Windows Excel/VBA HTTP client designed for deterministic testing, bounded concurrency, reliable retries, and constant-memory streaming. The project is built and verified with xlflow.

## Development status

Phase 6 adds constant-memory downloads and Phase 7 adds file/multipart uploads on the synchronous native WinHTTP backend. Phase 8 adds explicit native-only HTTP/2/HTTP/3 negotiation, gzip/deflate response decompression, and shared OS/direct/manual HTTP proxy routing. The default remains the late-bound `WinHttp.WinHttpRequest.5.1` transport; the native backend provides pointer-safe handles, negotiated protocol reporting, bounded file streaming, and OS-owned response decoding without a WinHTTP type-library reference.

```vb
Dim client As HttpClient
Dim request As New HttpRequest
Dim response As HttpResponse

Set client = VBAHttp.CreateClient()

request.Method = "GET"
request.Url = "https://example.com/api"
request.Query.Add "page", 1
request.Headers.SetValue "Accept", "application/json"

Set response = client.Execute(request)

If response.IsSuccess Then
    Debug.Print response.Text
Else
    Debug.Print response.StatusCode
End If
```

The convenience methods are named `GetResponse`, `PostResponse`, `PutResponse`, `PatchResponse`, and `DeleteResponse` because the corresponding bare HTTP verbs conflict with VBA language tokens.

Batch calls use cooperative WinHTTP async I/O while keeping VBA single-threaded. Results retain input order and isolate individual failures:

```vb
Dim urls As New Collection
Dim batch As HttpBatchResult
Dim options As New HttpBatchOptions

urls.Add "https://example.com/a"
urls.Add "https://example.com/b"
options.MaxConcurrency = 8

Set batch = client.GetMany(urls, options)
Debug.Print batch.SuccessCount, batch.FailureCount, batch.CancelledCount
```

The default maximum concurrency is 8. `HttpBatchOptions` also controls polling, host yielding, per-request scheduler deadlines, and an optional `HttpCancellationToken`.

For the native buffered backend, select it explicitly and inspect the negotiated protocol when needed:

```vb
Dim nativeClient As HttpClient
Dim nativeResponse As HttpResponse

Set nativeClient = VBAHttp.CreateNativeClient()
Set nativeResponse = nativeClient.GetResponse("https://example.com/api")
Debug.Print nativeResponse.ProtocolUsed
```

Modern protocol flags are opt-in on the native backend. The default leaves the
OS/WinHTTP negotiation policy unchanged; allow-fallback mode is suitable when
HTTP/1.1 remains acceptable:

```vb
Dim protocols As HttpProtocolOptions

Set protocols = VBAHttp.CreateProtocolOptions()
protocols.AllowHttp2 = True
protocols.AllowHttp3 = True
protocols.Mode = HttpProtocolAllowFallback
Set nativeClient.ProtocolOptions = protocols
Set nativeResponse = nativeClient.GetResponse("https://example.com/api")
Debug.Print nativeResponse.ProtocolUsed
```

Required mode is strict and requires HTTPS; unsupported capabilities raise the
stable `HttpErrorProtocol` category. The COM transport rejects advanced
protocol flags because its late-bound option surface has no HTTP/2/HTTP/3
control. See [`docs/specs/protocol-policy.md`](docs/specs/protocol-policy.md)
for the compatibility and evidence rules.

The offline loopback server intentionally speaks HTTP/1.1, so negotiated
HTTP/2/HTTP/3 promotion requires the separate fail-closed host runner. After a
release build, set `VBA_HTTP_PROTOCOL_HOST_URL` and
`VBA_HTTP_PROTOCOL_EXPECTED`, then run `task protocol:host`; it records only
path-free target metadata and the exact `ProtocolUsed` result.

Response decompression is also native-only. Select gzip/deflate explicitly; the
default leaves response bytes and `Accept-Encoding` unchanged:

```vb
Dim decompression As HttpDecompressionOptions

Set decompression = VBAHttp.CreateDecompressionOptions()
decompression.AllowGzip = True
decompression.Mode = HttpDecompressionAllowFallback
Set nativeClient.DecompressionOptions = decompression
Set nativeResponse = nativeClient.GetResponse("https://example.com/api")
Debug.Print nativeResponse.Text  ' decoded identity text
```

The COM backend rejects an active decompression option. See
[`docs/specs/decompression-policy.md`](docs/specs/decompression-policy.md) for
the `Accept-Encoding`, fallback, and streaming length contract.

Proxy routing is configured with the same snapshot-safe options on either
transport. Default mode uses the OS/WinHTTP configuration; `HttpProxyNoProxy`
connects directly; manual mode accepts an HTTP/HTTPS proxy URL and optional
semicolon-delimited bypass list:

```vb
Dim proxy As HttpProxyOptions

Set proxy = VBAHttp.CreateProxyOptions()
proxy.Mode = HttpProxyManual
proxy.ProxyUrl = "http://proxy.example.test:8080"
proxy.BypassList = "localhost;127.0.0.1"
Set client.ProxyOptions = proxy
```

Proxy credentials and HTTPS CONNECT authentication are deliberately handled by
the later authentication policy. See
[`docs/specs/proxy-policy.md`](docs/specs/proxy-policy.md).

Preemptive Basic and Bearer authentication is available through immutable
provider snapshots. Credentials require HTTPS unless an explicit insecure
opt-in is used for a controlled loopback test; providers reject conflicting
`Authorization` headers and disable automatic redirects:

```vb
Dim auth As IHttpAuthProvider

Set auth = VBAHttp.CreateBearerAuthProvider("opaque-token")
Set client.AuthProvider = auth
Set response = client.GetResponse("https://example.com/protected")
```

`VBAHttp.CreateBasicAuthProvider` uses UTF-8 RFC 7617 encoding. 401/407
responses are returned normally for preemptive providers. Buffered requests may
opt into bounded Windows/Digest/server-or-proxy challenge authentication with
`VBAHttp.CreateWindowsAuthProvider`; streaming uploads reject that provider
before opening a one-shot source. OAuth flows and interactive callbacks remain
outside the current provider contract; see
[`docs/specs/auth-policy.md`](docs/specs/auth-policy.md).

Requests carrying `Authorization`, `Proxy-Authorization`, or `Cookie` headers
also return 3xx responses instead of automatically forwarding those values to a
redirected origin. HTTPS-to-HTTP downgrade redirects are rejected by default;
see [`docs/specs/redirect-policy.md`](docs/specs/redirect-policy.md).

Cookie persistence is opt-in and caller-owned. Attach a jar to a client when a
session should store and replay matching `Set-Cookie` values:

```vb
Dim jar As HttpCookieJar

Set jar = VBAHttp.CreateCookieJar()
Set client.CookieJar = jar
Set response = client.GetResponse("https://example.com/session")
```

The jar applies host/path/secure/expiry rules, preserves an explicit caller
`Cookie` header, and suppresses automatic redirects when a cookie is present.
It is memory-only and is never shared with ambient WinHTTP state. See
[`docs/specs/cookie-policy.md`](docs/specs/cookie-policy.md).

Structured diagnostics are opt-in and caller-owned. The collector records
bounded top-level operation events with query-free targets and redacted
authorization, proxy-authorization, cookie, and authentication headers; it
never includes bodies or error descriptions:

```vb
Dim diagnostics As HttpDiagnostics

Set diagnostics = VBAHttp.CreateDiagnostics()
diagnostics.Enabled = True
Set client.Diagnostics = diagnostics
Set response = client.GetResponse("https://example.com/api")
Debug.Print diagnostics.ToJson
```

See [`docs/specs/diagnostics-policy.md`](docs/specs/diagnostics-policy.md) for
the schema and retention contract.

The native transport uses synchronous WinHTTP calls, OS certificate validation, and deterministic handle cleanup. Buffered native requests still reject active cancellation and total-deadline options because a blocking call cannot observe them; streaming downloads support cooperative cancellation and total-deadline checkpoints between bounded reads.

For a large file, use the native client and let the library publish only after the temporary file is complete:

```vb
Dim nativeClient As HttpClient
Dim download As HttpDownloadResult
Dim options As New HttpExecutionOptions

Set nativeClient = VBAHttp.CreateNativeClient()
nativeClient.BaseUrl = "http://127.0.0.1:8080"
Set download = nativeClient.DownloadFile("/large.bin", "C:\Temp\large.bin", options)
download.RaiseForStatus
Debug.Print download.BytesWritten, download.Published
```

File and multipart uploads use the same bounded native transport. The source file remains caller-owned and is never replayed automatically after cancellation, timeout, or a 401/407 challenge.

```vb
Dim upload As HttpUploadResult
Dim form As HttpMultipartForm

Set upload = nativeClient.UploadFile("/upload/hash", "C:\Temp\payload.bin")
upload.RaiseForStatus

Set form = VBAHttp.CreateMultipartForm()
form.AddField "title", "日本語"
form.AddFile "payload", "C:\Temp\payload.bin"
Set upload = nativeClient.UploadMultipart("/upload/multipart", form)
upload.RaiseForStatus
```

`DownloadFile` reads at most 64 KiB at a time, writes beside the destination, and atomically replaces the destination after a successful response. Existing destinations remain unchanged on HTTP failure, cancellation, timeout, or write failure. Implement `IHttpProgressSink` to receive integral byte counters; an unknown `Content-Length` is reported as `-1`.

GET, HEAD, PUT, DELETE, OPTIONS, and TRACE retry 408, 429, 500, 502, 503, 504 and transient DNS, connection, timeout, and I/O failures by default. POST, PATCH, and extension methods require explicit opt-in:

```vb
Dim policy As New HttpRetryPolicy
Dim options As New HttpExecutionOptions
Dim cancel As New HttpCancellationToken

policy.MaxAttempts = 4
policy.RetryNonIdempotentMethods = True
Set options.RetryPolicy = policy
options.TotalDeadlineMilliseconds = 10000
Set options.CancellationToken = cancel

Set response = client.PostResponse("https://example.com/jobs", Nothing, options)
```

`Retry-After` delta-seconds and HTTP-date are supported. Cancellation has priority over the total deadline, which includes every attempt and retry wait.

Use `VBAHttp.CreateClient()` when referencing the distributed workbook from another VBA project because VBA class modules are `PublicNotCreatable`. `VBAHttp.CreateRequest()` provides the same boundary for request-level timeout and transport-option configuration before `client.Execute`. `CreateRetryPolicy`, `CreateExecutionOptions`, `CreateBatchOptions`, `CreateCancellationToken`, and the protocol/decompression/proxy/authentication/cookie-jar factories expose the same configuration objects across that boundary. Source-vendored consumers may use `New`. The default transport can still be replaced through `HttpClient.Transport` for tests or custom backends.

## Contributor verification

```powershell
task verify
task release:build
```

Tests and benchmarks use only the deterministic loopback server. Release workbooks are generated with `xlflow build` and exclude tests, benchmarks, xlflow helpers, development modules, and test-only classes. The release gate verifies the manifest base/output, VBE/atomic-publication evidence, component source boundaries, and checksum before opening Excel; see [`docs/specs/release-security.md`](docs/specs/release-security.md).

The primary XLSM release is built with `task release:build`. An Excel add-in is
built independently from the tracked same-extension base with `task
release:xlam:build`; it never renames or mutates the XLSM artifact. See
[`docs/specs/distribution.md`](docs/specs/distribution.md) and ADR-0021.

Public API quick reference: [`docs/API.md`](docs/API.md). Contributor workflow:
[`CONTRIBUTING.md`](CONTRIBUTING.md). Distribution, install, and rollback rules:
[`docs/specs/distribution.md`](docs/specs/distribution.md). The handoff gate is
[`docs/RELEASE_CHECKLIST.md`](docs/RELEASE_CHECKLIST.md).

## Measured loopback evidence

The following values are recorded in
[`docs/BENCHMARKS_BASELINE.md`](docs/BENCHMARKS_BASELINE.md) and are not
universal Internet or hardware guarantees:

- bounded concurrency: 100 requests with 100 ms server delay, concurrency 16,
  measured 12.864x over sequential execution (maximum in-flight 16);
- constant-memory download: 1 GiB payload, working-set peak delta 6,836,224
  bytes and private-bytes peak delta 19,017,728 bytes;
- streaming upload: 1 GiB payload, server hash matched and 64 KiB write/read
  bounds remained active; the measured private-bytes peak delta was
  182,210,560 bytes;
- Phase 2 buffered comparison: Raw WinHttpRequest mean 0.499 ms versus
  VBA-HTTP 0.833 ms on the same loopback workload. The 15% target is therefore
  recorded as an engineering variance, not silently claimed as achieved.

Refresh these figures only with the documented benchmark methodology and keep
the machine-readable JSON evidence beside the narrative result.

Current contracts are documented in [`docs/specs/http-core-api.md`](docs/specs/http-core-api.md), [`docs/specs/native-winhttp-transport.md`](docs/specs/native-winhttp-transport.md), [`docs/specs/streaming-download.md`](docs/specs/streaming-download.md), [`docs/specs/streaming-upload.md`](docs/specs/streaming-upload.md), [`docs/specs/bounded-concurrency.md`](docs/specs/bounded-concurrency.md), and [`docs/specs/reliability-policy.md`](docs/specs/reliability-policy.md), with architectural decisions under [`docs/adr/`](docs/adr/).
