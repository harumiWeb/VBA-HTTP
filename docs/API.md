# VBA-HTTP API quick reference

The complete, maintained API documentation now lives in the
[public API guide](guides/api-reference.md). The task-oriented guides are
indexed at [`guides/README.md`](guides/README.md).

## Minimal example

```vb
Dim client As HttpClient
Dim response As HttpResponse

Set client = VBAHttp.CreateClient()
Set response = client.GetResponse("https://example.test/status/204")
response.RaiseForStatus
```

Use `VBAHttp.CreateRequest` for explicit methods, headers, query parameters,
bodies, timeouts, and transport policies. `HttpErrorCategory` and
`HttpErrors.CategoryFromNumber` provide the stable error model. Streaming
download/upload APIs require `VBAHttp.CreateNativeClient`.

For installation and source distribution, see the
[getting started guide](guides/getting-started.md) and
[distribution guide](guides/distribution.md). Normative implementation
contracts remain in [`docs/specs/`](specs/README.md).
