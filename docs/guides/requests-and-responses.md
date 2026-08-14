# Requests and responses

## Build a request

Use `VBAHttp.CreateRequest()` when convenience methods are not enough:

```vb
Dim request As HttpRequest
Dim response As HttpResponse

Set request = VBAHttp.CreateRequest()
request.Method = "POST"
request.Url = "/search"
request.Query.Add "q", "日本語"
request.Headers.SetValue "Accept", "application/json"
request.Headers.SetValue "Content-Type", "application/json; charset=utf-8"
request.Body.Text = "{""limit"":10}"

Set response = client.Execute(request)
response.RaiseForStatus
```

The client clones the request before resolving its URL, merging defaults, and
applying providers. The original request, headers, body, and option objects are
not mutated by execution.

## URL resolution

`BaseUrl` must be an absolute HTTP or HTTPS URL. A request URL can be absolute
or relative to the base. Fragments are client-only and are removed from the
wire request target. User-info (`user:password@host`) is rejected. Query
parameters are encoded as UTF-8 and placed before a fragment.

```vb
client.BaseUrl = "https://api.example.test/v1/"
request.Url = "users"
request.Query.Add "page", 2
' Wire target: /v1/users?page=2
```

Do not concatenate untrusted query strings yourself. Use `HttpParams` so names
and values receive the same encoding and validation in every backend.

## Methods, headers, and body

Methods are normalized and validated before transport creation. `HttpHeaders`
is case-insensitive for lookup and preserves insertion order for repeated
fields. Use `Add` for repeated fields such as `Accept` and `Set-Cookie`; use
`SetValue` when a name must have one effective value.

Header names reject colon and control characters. Header values reject CR, LF,
NUL, DEL, and C0 controls while allowing HTAB and Unicode. This validation
prevents request-header injection in both COM and native serialization paths.

`HttpBody` has three states:

| State | Set with | Read with | Typical use |
| --- | --- | --- | --- |
| Empty | `Clear` | `IsEmpty=True` | GET/DELETE without a body. |
| Text | `Body.Text = ...` | `Text` or `Bytes` | UTF-8 JSON/XML/form text. |
| Bytes | `SetBytes(byteArray)` | `Bytes` | Binary payloads. |

The library owns an execution snapshot. `Body.Bytes` is a defensive copy, so
large buffered uploads incur an intentional ownership boundary. Use streaming
APIs for files or payloads that should not be materialized in a VBA array.

## Response body and headers

`HttpResponse.Headers` contains the response header collection. `Body.Bytes`
returns exact bytes; `Text` decodes according to the response charset contract.
For a binary response, use `Body.Bytes` and do not force text decoding. A
malformed UTF-8 text body maps to the stable validation/protocol boundary rather
than returning replacement characters silently.

```vb
If response.Headers.Contains("Content-Type") Then
    Debug.Print response.Headers.GetValue("Content-Type")
End If

If response.IsSuccess Then
    Debug.Print response.Text
End If
```

`ProtocolUsed` reports the backend's observed protocol. It is evidence of what
WinHTTP negotiated, not what a caller requested. `ElapsedMilliseconds` is a
diagnostic duration for the completed operation.

## Status handling

`IsSuccess` is true for status 200 through 299. HTTP status failures are not
raised automatically, which permits callers to inspect error response headers
and bodies:

```vb
Set response = client.GetResponse("https://api.example.test/resource")
If response.StatusCode = 404 Then
    Debug.Print "not found"
Else
    response.RaiseForStatus
End If
```

Calling `RaiseForStatus` on a non-2xx response raises `HttpErrStatus`. Network,
TLS, timeout, cancellation, protocol, and I/O failures use the stable
`HttpErrorCategory` values and do not create a response object.

## Redirect behavior

`FollowRedirects` and `MaxRedirects` apply to automatic redirects. A request
containing `Authorization`, `Proxy-Authorization`, or `Cookie`, or one with an
authentication provider, disables automatic redirects so sensitive values do
not cross origins. HTTPS-to-HTTP downgrade redirects are rejected by default.
Callers that receive a 3xx response must validate the new origin and create a
new request explicitly.

See [Security and state](security-and-state.md) and the normative
[`redirect-policy.md`](../specs/redirect-policy.md).

## Request-level overrides

Client defaults are cloned at the start of each operation. A request-level
`ProtocolOptions`, `DecompressionOptions`, `ProxyOptions`, `AuthProvider`, or
`CookieJar` overrides the corresponding client setting. This makes a shared
client safe to configure for a common policy while keeping individual requests
isolated.

The only intentionally shared control is `HttpCancellationToken`: the caller
can retain the token and call `Cancel()` while an asynchronous COM batch or a
cooperative native streaming operation is active.
