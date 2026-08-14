# Streaming files and multipart forms

Streaming APIs require `VBAHttp.CreateNativeClient()`. The COM client rejects
these operations before opening network or file handles.

## Download a file

```vb
Dim result As HttpDownloadResult
Dim options As HttpExecutionOptions

Set options = VBAHttp.CreateExecutionOptions()
Set result = VBAHttp.CreateNativeClient().DownloadFile( _
    "https://files.example.test/archive.bin", _
    "C:\Temp\archive.bin", _
    options)
result.RaiseForStatus
Debug.Print result.BytesWritten, result.Published
```

`DownloadFile` reads at most 64 KiB at a time into a reusable buffer. It writes
to a unique temporary file beside the destination, closes the file, and then
publishes it with an atomic replacement. Existing destination bytes are not
opened for writing and remain unchanged after an HTTP, transport, callback,
cancellation, timeout, or write failure.

`HttpDownloadResult` reports `StatusCode`, `ReasonPhrase`, `Headers`,
`ProtocolUsed`, `DestinationPath`, `Published`, `BytesWritten`,
`ContentLength`, `ContentLengthKnown`, `ElapsedMilliseconds`, and `IsSuccess`.
`ContentLengthKnown=False` means `ContentLength=-1` and should not be used for a
progress denominator.

## Progress and cancellation

Implement `IHttpProgressSink` when a workbook needs progress updates:

```vb
Private Sub IHttpProgressSink_OnProgress( _
    ByVal BytesTransferred As Currency, _
    ByVal TotalBytes As Currency)
    ' TotalBytes is -1 when the server did not provide a usable length.
End Sub
```

Callbacks run after a chunk is written. A zero-byte response receives one
zero-progress callback. Callback errors abort and clean up the temporary file.
Cancellation and total-deadline checks occur before reads and after writes;
they are cooperative while a native WinHTTP call is blocked.

## Upload a file

```vb
Dim upload As HttpUploadResult
Set upload = VBAHttp.CreateNativeClient().UploadFile( _
    "https://files.example.test/upload", _
    "C:\Temp\payload.bin", _
    "application/octet-stream")
upload.RaiseForStatus
```

The source path is normalized, opened read-only, and sent in reusable chunks of
at most 64 KiB. A source-size change is reported as an I/O failure. The source
file is never modified, deleted, or implicitly replayed. A remote server may
already have accepted a prefix when a cancellation or write failure occurs;
that remote side effect cannot be rolled back by the client.

`HttpUploadResult` reports response metadata plus `BytesWritten`,
`ContentLength`, `ElapsedMilliseconds`, `AuthenticationChallenged`, and
`IsSuccess`. A 401/407 response is returned normally with
`AuthenticationChallenged=True`; the streaming body is not replayed.

## Multipart

```vb
Dim form As HttpMultipartForm
Dim upload As HttpUploadResult

Set form = VBAHttp.CreateMultipartForm()
form.AddField "title", "日本語"
form.AddFile "document", "C:\Temp\document.pdf", _
              "document.pdf", "application/pdf"

Set upload = client.UploadMultipart("/forms", form)
upload.RaiseForStatus
```

The encoder writes each boundary, header, field, and file incrementally. It
computes `Content-Length` from bounded UTF-8 chunks and file metadata without
joining the full form into one string or byte array. Names, filenames,
content-types, and boundaries are validated against header injection.

## Authentication and retries

Preemptive Basic/Bearer headers are allowed on streaming requests when the
provider's HTTPS policy permits them. `HttpWindowsAuthProvider` challenge
replay is rejected before opening a one-shot upload source. Streaming methods
do not apply the buffered retry policy; repeat the entire call only when the
application knows the operation is safe.

## Verification and contracts

The current chunk, atomic-publication, source-ownership, and failure-cleanup
invariants are specified in
[`../specs/streaming-download.md`](../specs/streaming-download.md) and
[`../specs/streaming-upload.md`](../specs/streaming-upload.md). They include
the 2 GiB-plus `Currency` counters and the 4 GiB WinHTTP content-length
boundary.
