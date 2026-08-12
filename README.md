# VBA-HTTP

VBA-HTTP is a Windows Excel/VBA HTTP client designed for deterministic testing, bounded concurrency, reliable retries, and constant-memory streaming. The project is built and verified with xlflow.

## Development status

Phase 1 establishes the network-independent domain model and transport contract. The public `HttpClient`, request/response, headers, query, timeout, body, encoding, and error components are implemented. A real network backend arrives in Phase 2; until then, execution requires an injected `IHttpTransport` implementation.

```vb
Dim client As New HttpClient
Dim request As New HttpRequest
Dim response As HttpResponse
Dim transport As IHttpTransport

request.Method = "GET"
request.Url = "https://example.com/api"
request.Query.Add "page", 1
request.Headers.SetValue "Accept", "application/json"

' CreateMyTransport is the consumer's IHttpTransport implementation in Phase 1.
Set transport = CreateMyTransport()
Set client.Transport = transport
Set response = client.Execute(request)

If response.IsSuccess Then
    Debug.Print response.Text
Else
    Debug.Print response.StatusCode
End If
```

The convenience methods are named `GetResponse`, `PostResponse`, `PutResponse`, `PatchResponse`, and `DeleteResponse` because the corresponding bare HTTP verbs conflict with VBA language tokens.

## Contributor verification

```powershell
task verify
task release:build
```

Tests and benchmarks use only the deterministic loopback server. Release workbooks are generated with `xlflow build` and exclude tests, benchmarks, xlflow helpers, development modules, and test-only classes.

Current contracts are documented in [`docs/specs/http-core-api.md`](docs/specs/http-core-api.md), with architectural decisions under [`docs/adr/`](docs/adr/).
