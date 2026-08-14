# Public API reference

This page is the complete consumer-facing API index. All names are VBA class,
module, property, or enum names and are case-insensitive in VBA. The source
package contains implementation helpers as well; helpers marked internal are
not part of the supported application-author API.

## Factory module: `VBAHttp`

Use factories when the project is referenced as a compiled workbook. The
library's public classes are `PublicNotCreatable` at the workbook boundary, so
factories are the portable construction mechanism for both referenced and
source-vendored consumers.

| Function | Returns | Purpose |
| --- | --- | --- |
| `CreateClient()` | `HttpClient` | Default late-bound WinHTTP COM client. |
| `CreateNativeClient()` | `HttpClient` | Native WinHTTP client for streaming and advanced options. |
| `CreateRequest()` | `HttpRequest` | Empty request with library defaults. |
| `CreateRetryPolicy()` | `HttpRetryPolicy` | Cloneable retry policy. |
| `CreateExecutionOptions()` | `HttpExecutionOptions` | Per-call retry/deadline/cancellation controls. |
| `CreateBatchOptions()` | `HttpBatchOptions` | Bounded batch scheduler controls. |
| `CreateCancellationToken()` | `HttpCancellationToken` | Cooperative cancellation signal. |
| `CreateProtocolOptions()` | `HttpProtocolOptions` | Native HTTP/2 policy; HTTP/3 is unsupported. |
| `CreateDecompressionOptions()` | `HttpDecompressionOptions` | Native gzip/deflate response decoding. |
| `CreateProxyOptions()` | `HttpProxyOptions` | Default, direct, or manual proxy routing. |
| `CreateCookieJar()` | `HttpCookieJar` | Explicit in-memory cookie state. |
| `CreateDiagnostics()` | `HttpDiagnostics` | Bounded, redacted operation events. |
| `CreateBasicAuthProvider(username, password, [allowInsecureHttp])` | `IHttpAuthProvider` | Preemptive Basic authentication. |
| `CreateBearerAuthProvider(token, [allowInsecureHttp])` | `IHttpAuthProvider` | Preemptive Bearer authentication. |
| `CreateWindowsAuthProvider(username, password, [scheme], [target], [allowInsecureHttp], [maxChallenges])` | `HttpWindowsAuthProvider` | Bounded buffered WinHTTP challenge authentication. |
| `CreateMultipartForm()` | `HttpMultipartForm` | Ordered multipart field/file form. |

Optional Boolean arguments default to `False` for insecure HTTP and the
provider's documented safe defaults for challenge settings.

## `HttpClient`

### Configuration

| Member | Type | Meaning |
| --- | --- | --- |
| `BaseUrl` | `String` | Absolute base URL used to resolve relative request URLs. |
| `DefaultHeaders` | `HttpHeaders` | Headers copied into every execution snapshot. |
| `Transport` | `IHttpTransport` | Replaceable buffered transport; default is COM. |
| `RetryPolicy` | `HttpRetryPolicy` | Client-level retry defaults. |
| `ProtocolOptions` | `HttpProtocolOptions` | Client-level native protocol override. |
| `DecompressionOptions` | `HttpDecompressionOptions` | Client-level native decoding override. |
| `ProxyOptions` | `HttpProxyOptions` | Client-level routing override. |
| `AuthProvider` | `IHttpAuthProvider` | Client-level authentication provider. |
| `CookieJar` | `HttpCookieJar` | Client-level caller-owned cookie state. |
| `Diagnostics` | `HttpDiagnostics` | Optional operation-event collector. |

All option objects are cloned when an operation starts. A request-level option
overrides its client-level counterpart. The cancellation token is intentionally
shared so the caller can signal an active operation.

### Methods

| Method | Result | Notes |
| --- | --- | --- |
| `Execute(request, [options])` | `HttpResponse` | Execute one buffered request. |
| `ExecuteMany(requests, [options])` | `HttpBatchResult` | Ordered bounded batch. |
| `GetMany(urls, [options])` | `HttpBatchResult` | Convenience batch of GET requests. |
| `GetResponse(url, [query], [options])` | `HttpResponse` | Convenience GET with optional `HttpParams`. |
| `PostResponse(url, [body], [options])` | `HttpResponse` | Convenience POST. |
| `PutResponse(url, [body], [options])` | `HttpResponse` | Convenience PUT. |
| `PatchResponse(url, [body], [options])` | `HttpResponse` | Convenience PATCH. |
| `DeleteResponse(url, [options])` | `HttpResponse` | Convenience DELETE. |
| `DownloadFile(url, destination, [options], [progress])` | `HttpDownloadResult` | Native constant-memory download. |
| `UploadFile(url, source, [contentType], [options], [progress])` | `HttpUploadResult` | Native file upload. |
| `UploadMultipart(url, form, [options], [progress])` | `HttpUploadResult` | Native streaming multipart. |

The same client is not reentrant. A callback must not call an operation on the
client currently executing.

## `HttpRequest`

| Member | Type | Default/contract |
| --- | --- | --- |
| `Method` | `String` | Uppercase HTTP method after validation. |
| `Url` | `String` | Absolute or client-relative URL; user-info is rejected. |
| `Headers` | `HttpHeaders` | Case-insensitive names, ordered values. |
| `Query` | `HttpParams` | Percent-encoded query parameters. |
| `Timeouts` | `HttpTimeouts` | Resolve/connect/send/receive phase limits. |
| `Body` | `HttpBody` | Empty, text, or byte-array body. |
| `FollowRedirects` | `Boolean` | `True` unless sensitive headers/provider force `False`. |
| `MaxRedirects` | `Long` | Bounded automatic redirect count. |
| `ProtocolOptions` | `HttpProtocolOptions` | Request override. |
| `DecompressionOptions` | `HttpDecompressionOptions` | Request override. |
| `ProxyOptions` | `HttpProxyOptions` | Request override. |
| `AuthProvider` | `IHttpAuthProvider` | Request provider overrides client provider. |
| `CookieJar` | `HttpCookieJar` | Request jar overrides client jar. |
| `Clone()` | `HttpRequest` | Snapshot-safe deep clone of mutable request state. |

## `HttpResponse` and `HttpBody`

`HttpResponse` exposes `StatusCode`, `ReasonPhrase`, `Headers`, `Body`,
`ElapsedMilliseconds`, `ProtocolUsed`, `IsSuccess`, and `Text`. `Text` decodes
the body according to the response charset contract; use `Body.Bytes` for
binary data. `RaiseForStatus()` raises `HttpErrStatus` for non-2xx responses.
`Clone()` returns a detached response copy.

`HttpBody` exposes `Kind`, `IsEmpty`, `Text` (get/let), `Bytes`, `SetBytes`,
`Clear`, and `Clone`. `Bytes` is a defensive copy. `SetBytes` accepts a VBA
byte array and takes a copy; the caller retains ownership of its array.

## Headers, parameters, and timeouts

### `HttpHeaders`

`Count`, `Add(name, value)`, `SetValue(name, value)`, `GetValue(name)`,
`GetValues(name)`, `Contains(name)`, `Remove(name)`, `Clear`, `NameAt(index)`,
`ValueAt(index)`, and `Clone()` are available. Names are case-insensitive but
the insertion order is retained. Names reject colon/control characters and
values reject CR, LF, NUL, DEL, and other C0 controls.

### `HttpParams`

`Count`, `Add(name, value)`, `SetValue(name, value)`, `Remove(name)`,
`ToQueryString()`, and `Clone()` are available. Values are UTF-8 percent
encoded. Query parameters are inserted before a URL fragment.

### `HttpTimeouts`

`ResolveMilliseconds`, `ConnectMilliseconds`, `SendMilliseconds`, and
`ReceiveMilliseconds` are non-negative phase limits. `Clone()` creates a
detached configuration object. A zero value means the backend default for that
phase.

## Reliability and batch types

### `HttpRetryPolicy`

`MaxAttempts`, `BaseDelayMilliseconds`, `MaxDelayMilliseconds`, `JitterRatio`,
`RespectRetryAfter`, `RetryNonIdempotentMethods`, and `Clone()` are available.
Defaults are 3 attempts, 200 ms base delay, 30 s maximum delay, 0.2 jitter,
`RespectRetryAfter=True`, and no retry for POST/PATCH.

### `HttpExecutionOptions`

`RetryPolicy`, `TotalDeadlineMilliseconds`, `CancellationToken`,
`PollIntervalMilliseconds`, `YieldToHost`, `YieldIntervalMilliseconds`, and
`Clone()` control one call. Cancellation wins over an expired deadline at each
checkpoint.

### `HttpCancellationToken`

`IsCancellationRequested` is read-only. `Cancel()` is idempotent and can be
called from a scheduled VBA callback or another cooperating host action.

### `HttpBatchOptions`, `HttpBatchResult`, and `HttpBatchItem`

`HttpBatchOptions` exposes `MaxConcurrency`, `PollIntervalMilliseconds`,
`YieldToHost`, `YieldIntervalMilliseconds`, `RequestDeadlineMilliseconds`,
`CancellationToken`, `RetryPolicy`, `TotalDeadlineMilliseconds`, and `Clone()`.
`HttpBatchResult` exposes `Count`, `ItemAt(index)`, `SuccessCount`,
`FailureCount`, and `CancelledCount`. Each `HttpBatchItem` exposes `Index`,
`Status`, `Response`, `ErrorNumber`, `ErrorCategory`, `ErrorSource`, and
`ErrorDescription`. Results retain input order and isolate individual failure.

`HttpAttemptOptions` is an advanced transport-attempt configuration type. It
is exposed for custom asynchronous transport implementations and is not needed
for normal `HttpClient` calls.

## Streaming and multipart types

`HttpDownloadResult` exposes `StatusCode`, `ReasonPhrase`, `Headers`,
`ProtocolUsed`, `DestinationPath`, `Published`, `BytesWritten`,
`ContentLength`, `ContentLengthKnown`, `ElapsedMilliseconds`, `IsSuccess`, and
`RaiseForStatus`.

`HttpUploadResult` exposes `StatusCode`, `ReasonPhrase`, `Headers`,
`ProtocolUsed`, `BytesWritten`, `ContentLength`, `ElapsedMilliseconds`,
`AuthenticationChallenged`, `IsSuccess`, and `RaiseForStatus`.

`HttpMultipartForm` exposes `Boundary`, `Count`, `AddField(name, value)`,
`AddFile(name, filePath, [fileName], [contentType])`, and `PartAt(index)`.
`HttpMultipartPart` exposes `IsFile`, `Name`, `Value`, `FilePath`, `FileName`,
and `ContentType`. Files are read incrementally and are never joined into one
payload-sized VBA array.

## Protocol, decompression, and proxy options

`HttpProtocolOptions` exposes `EnabledProtocols`, `Mode`, `HasOverride`,
`AllowHttp2`, `AllowHttp3`, `Clone`, and `Validate`. `AllowHttp3` is retained
for diagnostic compatibility only and is rejected as a supported production
capability; see [Compatibility](compatibility.md).

`HttpDecompressionOptions` exposes `EnabledEncodings`, `Mode`, `HasOverride`,
`AllowGzip`, `AllowDeflate`, `Clone`, and `Validate`. Native WinHTTP owns the
`Accept-Encoding` and decoded response-body contract. COM rejects an active
decompression override.

`HttpProxyOptions` exposes `Mode`, `ProxyUrl`, `BypassList`, `HasOverride`,
`Clone`, and `Validate`. The modes are `HttpProxyDefault`, `HttpProxyNoProxy`,
and `HttpProxyManual`; credentials must not be embedded in a proxy URL.

## Authentication, cookies, and diagnostics

`HttpBasicAuthProvider` and `HttpBearerAuthProvider` are created by factory and
expose only `AllowInsecureHttp`; secrets have no getter. `IHttpAuthProvider`
defines `Apply(request)` and `Clone()` for custom providers. Providers run on
the execution snapshot, reject conflicting `Authorization`, require HTTPS by
default, and disable automatic redirects.

`HttpWindowsAuthProvider` exposes `Scheme`, `Target`, `MaxChallenges`, and
`AllowInsecureHttp`. It is for buffered WinHTTP challenge exchange only;
streaming uploads reject it before opening a source.

`HttpCookieJar` exposes `Count`, `Clock`, `Clear`, `GetCookieHeader(url)`, and
`StoreResponseCookies(url, headers)`. `HttpCookieRecord` exposes `Name`,
`Value`, `Domain`, `Path`, `HostOnly`, `Secure`, `HasExpires`, `ExpiresAt`,
`SerializedLength`, `IsExpired(nowUtc)`, and `Matches(target, nowUtc)`.
`HttpCookieUrl` exposes `Scheme`, `Host`, `Path`, and `IsHttps`.

`HttpDiagnostics` exposes `Enabled`, `MaxEvents`, `Count`, `ItemAt(index)`,
`LastEvent`, `Clear`, and `ToJson()`. `HttpDiagnosticEvent` exposes
`Sequence`, `Operation`, `Method`, `Target`, `StatusCode`, `ProtocolUsed`,
`ElapsedMilliseconds`, `ErrorNumber`, `ErrorCategory`, `RequestHeaders`,
`ResponseHeaders`, and `ToJson()`. Events never contain bodies, query values,
or unredacted sensitive headers.

## Interfaces and extension points

| Interface | Required member(s) | Use |
| --- | --- | --- |
| `IHttpTransport` | `Execute(request)` | Custom buffered transport. |
| `IHttpBatchTransport` | `ExecuteMany(requests, options)` | Custom bounded batch backend. |
| `IHttpDownloadTransport` | `DownloadFile(request, destination, options, [progress])` | Custom streaming download backend. |
| `IHttpUploadTransport` | `UploadFile(...)`, `UploadMultipart(...)` | Custom streaming upload backend. |
| `IHttpProgressSink` | `OnProgress(bytesTransferred, totalBytes)` | Progress callback; `totalBytes=-1` means unknown. |
| `IHttpAuthProvider` | `Apply(request)`, `Clone()` | Snapshot-safe pre-send authentication. |
| `IHttpCookieClock` | clock implementation members | Deterministic cookie expiry tests. |

Custom transports must preserve the response/error and ownership contracts.
Supplying cancellation or total-deadline controls to a blocking transport that
does not implement the attempt interface is rejected before I/O.

`WinHttpComTransport` and `WinHttpNativeTransport` are exposed transport
implementations for advanced consumers. Prefer the `VBAHttp` factories so the
backend capability checks remain centralized.

## Enums and constants

| Enum | Values |
| --- | --- |
| `HttpBodyKind` | `HttpBodyEmpty`, `HttpBodyText`, `HttpBodyBytes` |
| `HttpErrorCategory` | `HttpErrorNone`, `Validation`, `InvalidUrl`, `Dns`, `Connection`, `Tls`, `Timeout`, `Cancelled`, `Protocol`, `Io`, `Status` |
| `HttpBatchItemStatus` | `Succeeded`, `Failed`, `Cancelled` |
| `HttpProtocolFlag` | HTTP/1.1 baseline, HTTP/2, retained HTTP/3 diagnostic flag |
| `HttpProtocolMode` | `HttpProtocolAllowFallback`, `HttpProtocolRequired` |
| `HttpDecompressionFlag` | Gzip, Deflate, All |
| `HttpDecompressionMode` | Allow fallback, Required |
| `HttpProxyMode` | Default, No proxy, Manual |
| `HttpAuthChallengeScheme` | Auto, Basic, NTLM, Digest, Negotiate |
| `HttpAuthChallengeTarget` | Server, Proxy |

Use symbolic constants from `HttpConstants` rather than numeric values.

## Errors

The public error numbers are exposed by `HttpErrors`: `HttpErrValidation`,
`HttpErrInvalidUrl`, `HttpErrDns`, `HttpErrConnection`, `HttpErrTls`,
`HttpErrTimeout`, `HttpErrCancelled`, `HttpErrProtocol`, `HttpErrIo`, and
`HttpErrStatus`. `HttpErrors.CategoryFromNumber`, `NumberFromCategory`, and
`CategoryName` map between the stable category and VBA error number.

HTTP status failures are ordinary responses until `RaiseForStatus` is called.
Transport failures use the stable category; descriptions are sanitized and do
not contain credentials, bodies, query values, or sensitive header values.

## Related contracts

- [Requests and responses](requests-and-responses.md)
- [Reliability and batches](reliability-and-batches.md)
- [Streaming](streaming.md)
- [Security and state](security-and-state.md)
- [Normative core API specification](../specs/http-core-api.md)
