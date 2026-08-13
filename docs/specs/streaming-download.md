# Streaming download contract

## Public API

```vb
Dim result As HttpDownloadResult
Dim options As HttpExecutionOptions
Dim progress As IHttpProgressSink

Set result = client.DownloadFile( _
    "/large.bin", _
    "C:\Temp\large.bin", _
    options, _
    progress)
```

`HttpClient.DownloadFile` clones the URL into a GET request, resolves it with
`BaseUrl`, applies default headers, and delegates to the configured
`IHttpDownloadTransport`. The method is synchronous and rejects reentrant use
of the same client. A COM-only client reports a validation error before opening
network or file handles.

`HttpDownloadResult` exposes:

- `StatusCode`, `ReasonPhrase`, `Headers`, and `ProtocolUsed` from the completed
  HTTP response;
- `IsSuccess` for status 200 through 299;
- `Published`, `DestinationPath`, `BytesWritten`, and `ElapsedMilliseconds`;
- `ContentLengthKnown` and `ContentLength`, where `ContentLength = -1` means
  unknown; and
- `RaiseForStatus`, which raises `HttpErrStatus` for a non-2xx result.

For non-2xx responses, `Published` is false and no destination replacement is
attempted. The response body is deliberately not buffered by this API.

## Progress and controls

Consumers may implement `IHttpProgressSink`:

```vb
Private Sub IHttpProgressSink_OnProgress( _
    ByVal BytesTransferred As Currency, _
    ByVal TotalBytes As Currency)
    ' TotalBytes = -1 when Content-Length is unknown.
End Sub
```

The callback runs after each chunk has been written to the temporary file. A
zero-byte response receives one progress callback with `BytesTransferred = 0`.
The callback must not call the same client reentrantly. Callback errors are
propagated after temporary-file cleanup.

`HttpExecutionOptions.CancellationToken` and
`TotalDeadlineMilliseconds` are checked before request I/O, before each native
read, and after each file write. Cancellation is cooperative while a native
read is blocked; `HttpTimeouts.ReceiveMilliseconds` remains the upper bound for
that individual blocking call.

The `RetryPolicy` property is not applied by streaming downloads in Phase 6.
Retry the complete `DownloadFile` call if desired; a future policy may add
per-attempt temporary-file semantics.

## File and memory guarantees

1. The destination path is normalized and its existing file is not opened for
   writing.
2. A unique temporary path is created in the destination's parent directory.
3. WinHTTP data is read into at most a 64 KiB byte array and written to the
   temporary file in bounded chunks.
4. The temporary file is closed before atomic replace with
   `MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH`.
5. On transport, callback, cancellation, deadline, or write failure, handles
   are closed and the temporary file is removed. The destination is unchanged.

The same-directory temporary path is required so publication does not become a
cross-volume copy. `Currency` byte counters are integral values and keep the
public declaration portable across the retained VBA ABI branches; they support
the Phase 6 2 GiB-plus tests without a platform-specific `LongLong` public
type. This source-level portability does not make 32-bit Office a supported
runtime; the current support boundary is x64 Office (ADR-0030).

## Verification

- Unit tests cover result/status semantics, unknown content length, and byte
  counter limits.
- Loopback integration tests cover known-length and chunked streaming, progress,
  atomic replacement, cancellation, and cleanup.
- A stress runner downloads a 1 GiB deterministic stream, verifies its SHA-256
  against the test server, and records file size, elapsed time, and resource
  observations.
- The latest x64 run used the fixed 64 KiB buffer and a small native warm-up
  before the memory gate. It transferred 1,073,741,824 bytes in 44,904.709 ms;
  the SHA-256 was
  `3a33d58aa8ee1e9d21fd4f510cc5d1ce8d25ba5e24363d19c28e2bf866f4185c`.
  Working-set peak delta was 6,836,224 bytes and private-bytes peak delta was
  19,017,728 bytes. The machine-readable result is
  `benchmarks/results/phase6-download-stress.json`.

## Evidence

- ADR: `docs/adr/ADR-0007-streaming-download-publication.md`
- Implementation: `src/classes/WinHttpNativeTransport.cls`
- Tests: `src/modules/Tests/Integration/WinHttpNativeTransportTests.bas` and
  `src/modules/Tests/Stress/WinHttpDownloadStressTests.bas`
