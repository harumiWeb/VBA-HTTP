# VBA-HTTP

VBA-HTTP is a Windows Excel/VBA HTTP client designed for deterministic testing, bounded concurrency, reliable retries, and constant-memory streaming. The project is built and verified with xlflow.

## Development status

Phase 6 adds constant-memory downloads and Phase 7 adds file/multipart uploads on the synchronous native WinHTTP backend. The default remains the late-bound `WinHttp.WinHttpRequest.5.1` transport; the native backend provides pointer-safe handles, negotiated protocol reporting, and bounded file streaming without a WinHTTP type-library reference.

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

Use `VBAHttp.CreateClient()` when referencing the distributed workbook from another VBA project because VBA class modules are `PublicNotCreatable`. `CreateRetryPolicy`, `CreateExecutionOptions`, `CreateBatchOptions`, and `CreateCancellationToken` expose the same configuration objects across that boundary. Source-vendored consumers may use `New`. The default transport can still be replaced through `HttpClient.Transport` for tests or custom backends.

## Contributor verification

```powershell
task verify
task release:build
```

Tests and benchmarks use only the deterministic loopback server. Release workbooks are generated with `xlflow build` and exclude tests, benchmarks, xlflow helpers, development modules, and test-only classes.

Current contracts are documented in [`docs/specs/http-core-api.md`](docs/specs/http-core-api.md), [`docs/specs/native-winhttp-transport.md`](docs/specs/native-winhttp-transport.md), [`docs/specs/streaming-download.md`](docs/specs/streaming-download.md), [`docs/specs/streaming-upload.md`](docs/specs/streaming-upload.md), [`docs/specs/bounded-concurrency.md`](docs/specs/bounded-concurrency.md), and [`docs/specs/reliability-policy.md`](docs/specs/reliability-policy.md), with architectural decisions under [`docs/adr/`](docs/adr/).
