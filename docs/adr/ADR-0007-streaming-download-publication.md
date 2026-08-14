# ADR-0007: Streaming download ownership and publication

## Status

`accepted`

## Background

Large responses cannot be represented by the buffered `HttpResponse.Body`
without making peak VBA memory proportional to the payload. A download API also
must not expose a partially written destination when a request, callback, or
file operation fails. The Phase 5 native transport already owns the WinHTTP
handle hierarchy, but it intentionally has no public streaming or progress
boundary.

## Decision

- `HttpClient.DownloadFile` is an additive capability API. It requires the
  configured transport to implement `IHttpDownloadTransport`; the default COM
  transport fails validation before network I/O, while the native transport
  provides the capability.
- A download writes to a uniquely named temporary file in the destination's
  parent directory. The writer is closed before `MoveFileExW` publishes the
  temporary file with replace-existing and write-through flags. The destination
  is never opened for writing and remains unchanged when a download fails or is
  cancelled before publication.
- The transport reads bounded chunks directly with `WinHttpReadData`. It does
  not call `WinHttpQueryDataAvailable` in the hot path because the synchronous
  read already reports the number of bytes returned and zero is the end marker.
  It never constructs an `HttpBody` or a payload-sized `Byte()` array. A 64 KiB
  chunk is the default upper bound for both network and file writes.
- `IHttpProgressSink.OnProgress` is a synchronous, same-thread VBA interface.
  It is called after each successfully written chunk with byte counts expressed
  as integral `Currency` values; `TotalBytes = -1` means that the response did
  not provide a usable `Content-Length`. A sink exception aborts the download
  and removes the temporary file.
- Cancellation and total-deadline checks occur before opening the request,
  before each read, and after each file write. They are cooperative: a
  `WinHttpReadData` call already in progress is bounded by the request receive
  timeout and cannot be interrupted by a VBA callback.
- `HttpDownloadResult` returns status and metadata for every completed HTTP
  response. Only a 2xx result publishes the destination; non-2xx results have
  `Published = False` and leave the existing destination untouched. Transport,
  validation, cancellation, deadline, and file failures raise through the
  existing `HttpErrors` contract. `RaiseForStatus` is available for callers
  that want exception-style status handling.
- Streaming downloads do not apply the buffered retry policy in Phase 6. A
  caller may retry the entire operation, and a later ADR may add atomic
  per-attempt retry semantics after measuring the cost of discarding partial
  streams.

## Consequences

- Successful downloads have constant working memory with respect to payload
  size, at the cost of temporary disk space and one bounded chunk buffer.
- Atomic replacement avoids corrupting an existing destination, but a failed
  replacement can leave the destination unchanged and raises an I/O error.
- `Currency` is used for byte counts because it is available across the retained
  32-bit and 64-bit VBA declaration branches and exactly represents integral
  byte values through the supported file-size range without exposing a
  platform-specific `LongLong` type. This ABI portability does not expand the
  supported Office runtime beyond x64 (ADR-0030).
- Progress callbacks are easy to implement from VBA and deterministic in tests,
  but they must not perform reentrant calls on the same `HttpClient`.
- Direct reads reduce one WinHTTP availability query per chunk and preserve the
  same cooperative cancellation boundary: a read already in progress remains
  bounded by the receive timeout. The optimization is not a claim of improved
  throughput until a PID-scoped x64 before/after benchmark is recorded.
- HTTP error bodies are not buffered by `DownloadFile`; callers that need them
  can use `Execute` and inspect the returned `HttpResponse` instead.

## Rationale

- Code: `src/classes/IHttpDownloadTransport.cls`,
  `src/classes/IHttpProgressSink.cls`, `src/classes/HttpDownloadResult.cls`,
  `src/classes/WinHttpDownloadFileWriter.cls`, and
  `src/modules/WinHttpDownloadFileSystem.bas`.
- Transport: `src/classes/WinHttpNativeTransport.cls`.
- Tests: native download integration, cancellation cleanup, and stress hash
  verification under `src/modules/Tests/` and `tools/`.
- Related decision: `docs/adr/ADR-0006-native-winhttp-callback-and-handle-ownership.md`.

## Supersedes

- None

## Superseded by

- None
