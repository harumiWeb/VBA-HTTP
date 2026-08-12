# VBA-HTTP

VBA-HTTP is a Windows Excel/VBA HTTP client designed for deterministic testing, bounded concurrency, reliable retries, and constant-memory streaming. The project is built and verified with xlflow.

## Development status

Phase 3 provides synchronous buffered HTTP plus bounded batch execution backed by late-bound `WinHttp.WinHttpRequest.5.1`. The public `HttpClient`, request/response, headers, query, timeout, body, encoding, stable error, cancellation, and ordered batch-result components require no WinHTTP type-library reference.

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

Use `VBAHttp.CreateClient()` when referencing the distributed workbook from another VBA project because VBA class modules are `PublicNotCreatable`. Source-vendored consumers may use `New HttpClient`. The default transport can still be replaced through `HttpClient.Transport` for tests or custom backends.

## Contributor verification

```powershell
task verify
task release:build
```

Tests and benchmarks use only the deterministic loopback server. Release workbooks are generated with `xlflow build` and exclude tests, benchmarks, xlflow helpers, development modules, and test-only classes.

Current contracts are documented in [`docs/specs/http-core-api.md`](docs/specs/http-core-api.md) and [`docs/specs/bounded-concurrency.md`](docs/specs/bounded-concurrency.md), with architectural decisions under [`docs/adr/`](docs/adr/).
