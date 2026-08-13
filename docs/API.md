# VBA-HTTP API quick reference

The distributed workbook exposes `PublicNotCreatable` classes through the
`VBAHttp` factory module. Consumers should use the factories instead of
instantiating public classes directly.

## Client and request

```vb
Dim client As HttpClient
Dim response As HttpResponse

Set client = VBAHttp.CreateClient()
Set response = client.GetResponse("https://example.test/status/204")
response.RaiseForStatus
```

Use `VBAHttp.CreateNativeClient()` when native WinHTTP capabilities are
required. `HttpRequest` can be used with `client.Execute` for method, headers,
query parameters, body, timeout, retry, cancellation, redirect, protocol,
decompression, proxy, authentication, cookie, and diagnostics options.

| Factory | Purpose |
| --- | --- |
| `CreateClient` / `CreateNativeClient` | COM-default or native WinHTTP client |
| `CreateRequest` / `CreateHeaders` / `CreateParams` | request composition |
| `CreateExecutionOptions` / `CreateBatchOptions` | per-call retry, deadline, and batch policy |
| `CreateRetryPolicy` / `CreateCancellationToken` | reliability controls |
| `CreateProtocolOptions` / `CreateDecompressionOptions` | native capability policy |
| `CreateProxyOptions` | default, direct, or manual proxy routing |
| `CreateBasicAuthProvider` / `CreateBearerAuthProvider` | preemptive HTTPS credentials |
| `CreateWindowsAuthProvider` | bounded buffered challenge credentials; COM uses `HttpAuthSchemeAuto`, while explicit Basic/Digest/NTLM/Negotiate selection is native-only |
| `CreateCookieJar` / `CreateDiagnostics` | explicit session state and redacted events |

## Response and errors

`HttpResponse.StatusCode`, `Headers`, `Body.Bytes`, `Text`, and
`ProtocolUsed` are available after a buffered request. `RaiseForStatus` raises
the stable library error model. Callers can inspect `HttpErrors.CategoryFromNumber`
or handle result objects for batch/download/upload operations.

The stable categories include validation, timeout, cancellation, connection,
TLS, protocol, HTTP status, capability, and I/O errors. Error descriptions and
diagnostics never contain request bodies, query values, credentials, or
sensitive header values.

## Streaming operations

```vb
Dim options As HttpExecutionOptions
Dim result As HttpDownloadResult

Set options = VBAHttp.CreateExecutionOptions()
Set result = client.DownloadFile("https://example.test/bytes/1048576", _
                                 "C:\Temp\payload.bin", options)
result.RaiseForStatus
```

`DownloadFile`, `UploadFile`, and `UploadMultipart` require the native client.
They use bounded buffers, preserve caller-owned source/destination files on
failure, and report cancellation/timeout without replaying one-shot bodies.

## Safety defaults

- credentials require HTTPS unless a caller explicitly opts into insecure
  loopback testing;
- buffered Windows/Digest challenge credentials are replayed only within their
  configured limit; streaming uploads reject them before opening the source;
- authorization, proxy-authorization, and cookie-bearing requests do not
  automatically follow redirects;
- POST/PATCH and streaming bodies are not retried by default;
- cookies are caller-owned in-memory state, never ambient WinHTTP state;
- gzip/deflate and HTTP/2/3 controls are explicit native options; COM rejects
  unsupported native-only options;
- release workbooks are filtered by `xlflow build` and never contain tests,
  benchmarks, or xlflow helper modules.

See the authoritative contracts in [`docs/specs/`](specs/README.md), especially
[`http-core-api.md`](specs/http-core-api.md),
[`reliability-policy.md`](specs/reliability-policy.md), and
[`distribution.md`](specs/distribution.md).
