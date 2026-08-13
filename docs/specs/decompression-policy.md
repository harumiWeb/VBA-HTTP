# Native response decompression policy

## Scope

`HttpDecompressionOptions` is an opt-in native WinHTTP capability for gzip and
deflate response bodies. It applies to buffered requests and native streaming
downloads. The default `WinHttpComTransport` rejects a non-empty option before
network I/O; it does not add a best-effort `Accept-Encoding` header.

## API and defaults

```vb
Dim decompression As HttpDecompressionOptions
Set decompression = VBAHttp.CreateDecompressionOptions()
decompression.AllowGzip = True
decompression.Mode = HttpDecompressionAllowFallback

Set nativeClient.DecompressionOptions = decompression
Set response = nativeClient.GetResponse("https://api.example.test/items")
```

`EnabledEncodings = 0` means no override. `HttpDecompressionGzip = 1`,
`HttpDecompressionDeflate = 2`, and `HttpDecompressionAll = 3` match the
WinHTTP mask. Required mode needs at least one selected encoding. Client and
request options are validated and cloned into the execution snapshot.

## Header and body ownership

When an override is active, the native transport sets
`WINHTTP_OPTION_DECOMPRESSION` (118) before adding request headers. WinHTTP
generates or overrides `Accept-Encoding` and returns decoded identity bytes.
Supplying `Accept-Encoding` on the same request is rejected with
`HttpErrValidation`; callers cannot create an ambiguous ownership order. A
request with no override keeps caller headers and receives the wire body as-is.

`HttpResponse.Body` and `HttpResponse.Text` therefore represent decoded bytes
when the option is active. Header behavior follows WinHTTP; callers must not
infer compressed length from a response that has been decoded.

## Fallback and errors

- Allow-fallback treats WinHTTP error 12009 (`ERROR_WINHTTP_INVALID_OPTION`) as
  a capability miss and omits the option. The request then follows normal
  server/WinHTTP behavior.
- Required mode maps that capability miss to `HttpErrorProtocol`.
- Other option failures and malformed compressed responses use the existing
  native WinHTTP error mapping. No raw body, header value, or credential enters
  an error description.
- The COM transport rejects active options with a stable validation error and
  does not make a network request.

## Streaming download contract

WinHTTP may expose the compressed wire `Content-Length` while returning decoded
bytes. For an active override, `HttpDownloadResult.ContentLengthKnown` is
always false and `ContentLength` is `-1`; `BytesWritten`, destination hash, and
progress transfer counters are authoritative. Temporary-file, atomic-replace,
failure cleanup, and cancellation rules remain those in
`streaming-download.md`.

## Compatibility and evidence

The option is documented by Microsoft as available from Windows 8.1. The
project's current real evidence and supported runtime target is x64 Office on
the configured host. 32-bit Office is unsupported by policy; conditional
declarations remain a static ABI guard only. The deterministic local
server exposes `/compress/gzip` and `/compress/deflate`; tests disable Go's
automatic decompression when verifying the wire fixture. The release artifact
consumer smoke uses the native factory and verifies decoded fixture text.

## Evidence

- Unit: `src/modules/Tests/Unit/HttpDecompressionOptionsTests.bas`,
  `HttpRequestTests.bas`, `HttpClientTests.bas`, and `VBAHttpTests.bas`.
- Integration: `WinHttpNativeTransportTests` gzip, deflate, header-conflict,
  and decompressed-download tests.
- Fixture: `tools/testserver/server.go` and `server_test.go`.
- Release: `tools/consumer/ReleaseBatchSmoke.bas` and
  `tools/Validate-ReleaseArtifact.ps1`.
- Decision: `docs/adr/ADR-0010-native-response-decompression-policy.md`.
