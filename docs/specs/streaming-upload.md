# Streaming upload contract

## Public API

```vb
Dim result As HttpUploadResult
Dim options As HttpExecutionOptions

Set result = client.UploadFile( _
    "/upload/hash", _
    "C:\Temp\payload.bin", _
    "application/octet-stream", _
    options)
result.RaiseForStatus
Debug.Print result.BytesWritten, result.ContentLength
```

`HttpClient.UploadFile` and `HttpClient.UploadMultipart` clone the URL into a
`POST` request, resolve it with `BaseUrl`, apply default headers, and delegate
to `IHttpUploadTransport`. Both methods are synchronous and reject reentrant
use of the same client. A COM-only client reports a validation error before
opening network or file handles.

`HttpUploadResult` exposes response `StatusCode`, `ReasonPhrase`, `Headers`,
and `ProtocolUsed`; request `BytesWritten` and `ContentLength`; elapsed time;
`IsSuccess`; and `AuthenticationChallenged` for 401/407. Non-2xx responses are
returned without buffering a response body. `RaiseForStatus` is available for
exception-style handling.

## File upload

The source path is normalized, verified as a regular file, and measured before
the request starts. `WinHttpWriteData` receives no more than 64 KiB at a time.
The file reader reuses its 64 KiB array for full-size chunks and resizes it only
for the final short tail, so the source-size and read-only ownership contract
does not require one VBA array allocation per chunk.
The source is opened read-only and remains caller-owned. A source size change
causes an I/O failure rather than a silently truncated or overlong request.
The caller may set a content type; the default is `application/octet-stream`.

## Multipart upload

`VBAHttp.CreateMultipartForm()` returns an ordered `HttpMultipartForm`:

```vb
Dim form As HttpMultipartForm
Set form = VBAHttp.CreateMultipartForm()
form.Boundary = "----vba-http-test-boundary"
form.AddField "title", "日本語"
form.AddFile "payload", "C:\Temp\payload.bin", "payload.bin", "application/octet-stream"
Set result = client.UploadMultipart("/upload/multipart", form)
```

Each field and file is emitted as boundary, headers, content, and CRLF. The
encoder computes `Content-Length` from bounded UTF-8 chunks and file metadata,
then emits the same parts incrementally. It never joins the complete form into
one string or byte array. Names, filenames, content types, and boundaries are
validated against header injection.

## Controls and failure semantics

`HttpExecutionOptions.CancellationToken` and
`TotalDeadlineMilliseconds` are checked before request I/O, before each source
read/write, and after every native write. `IHttpProgressSink` receives an
initial `(0, TotalBytes)` callback and callbacks after each successful write.
Cancellation is cooperative while a native write is blocked; the request send
timeout remains the upper bound for that call.

Upload operations do not apply the buffered retry policy. A 401/407 response is
returned with `AuthenticationChallenged = True`; the library never replays the
body automatically. A transport, source, write, cancellation, deadline, or
source-stability failure raises through `HttpErrors`, closes native handles,
and leaves the source file untouched. Bytes already accepted by a remote server
cannot be rolled back.

An `IHttpAuthProvider` may add a preemptive Basic or Bearer header to the
execution snapshot, but a `HttpWindowsAuthProvider` challenge cannot replay a
one-shot source. Native file and multipart uploads reject that provider before
opening the source; without a provider, a 401/407 remains the normal
`AuthenticationChallenged` result. The provider therefore also disables
automatic redirects. Replayable streaming sources and host-specific proxy/
CONNECT evidence remain separate compatibility work.

## Verification

- Unit tests cover result/status/auth-challenge semantics, multipart boundary
  validation, UTF-8 length, and byte-count limits.
- Loopback integration tests cover file hash, Unicode multipart fields,
  multiple files, progress, cancellation, source stability, and handle cleanup.
- A stress runner uploads a 1 GiB deterministic file, verifies the server-side
  SHA-256, and records elapsed time, memory, and native handle observations.
- The release consumer smoke calls both file and multipart APIs against the
  filtered workbook without injecting test code into it.

The upload buffer-reuse path is an internal optimization. A new PID-scoped x64
stress run must compare throughput, CPU, memory, handle count, and source hash
against the existing 64 KiB baseline before the result is published as a
performance claim.

For files at or above 4 GiB, the native transport sends the full decimal
`Content-Length` header and passes WinHTTP's zero-valued
`WINHTTP_IGNORE_REQUEST_TOTAL_LENGTH` sentinel. Lengths below 4 GiB preserve
the DWORD bit pattern in the VBA `Long` declaration.
