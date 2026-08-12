# ADR-0010: Native response decompression policy

## Status

`accepted`

## Background

WinHTTP can request gzip or deflate response decompression and return the
decoded bytes. The default `WinHttp.WinHttpRequest.5.1` COM option surface does
not expose this capability, and implementing a second zlib/gzip stack in VBA
would duplicate OS behavior and create a large-memory path. The library needs
one explicit ownership rule for `Accept-Encoding`, including streaming
downloads and runtimes where the WinHTTP option is unavailable.

## Decision

- Add `HttpDecompressionOptions` to `HttpClient` and `HttpRequest`. The default
  has no override, so existing response bytes and caller headers are unchanged.
  `AllowGzip` and `AllowDeflate` select the WinHTTP encoding mask; combined
  selection is `HttpDecompressionAll`. `HttpDecompressionAllowFallback` is the
  default configured mode, while required mode rejects an unavailable runtime.
- The native transport applies `WINHTTP_OPTION_DECOMPRESSION` (118) before
  request headers or body are sent. WinHTTP owns `Accept-Encoding` generation
  and response decoding. A configured decompression override together with a
  caller-supplied `Accept-Encoding` header is a validation error rather than an
  undocumented precedence rule.
- An invalid-option capability miss (WinHTTP error 12009) is skipped in
  allow-fallback mode. Required mode maps that miss to `HttpErrorProtocol`;
  other native failures use the existing WinHTTP mapping. No Brotli or manual
  decompression is claimed.
- The COM transport rejects a non-empty decompression override before network
  I/O because its late-bound option enum cannot configure native decoding.
- `HttpResponse` and successful `DownloadFile` results expose decoded identity
  bytes. When a download has an active decompression override,
  `ContentLengthKnown` is conservatively false because the wire
  `Content-Length` may describe compressed bytes; progress and hash assertions
  use bytes actually written.

## Consequences

- Native consumers get bounded WinHTTP decoding for buffered and streaming
  operations without a VBA compression dependency or a second body buffer.
- The default client and COM backend remain backward-compatible, but callers
  must select the native backend for decompression.
- Required mode can fail on Windows/WinHTTP versions before the option's
  compatibility floor (Windows 8.1) or on a host with the option disabled.
- A compressed response without an explicit option remains raw and is not
  silently decoded. Credential, proxy, and redirect policies remain separate
  decisions.

## Rationale

- Code: `src/classes/HttpDecompressionOptions.cls`, `HttpRequest.cls`,
  `HttpClient.cls`, `WinHttpNativeTransport.cls`, `WinHttpNativeApi.bas`, and
  `WinHttpComTransport.cls`.
- Tests: `src/modules/Tests/Unit/HttpDecompressionOptionsTests.bas`, request and
  client snapshot tests, deterministic Go compression fixtures, and native
  gzip/deflate/download integration tests.
- Current contract: `docs/specs/decompression-policy.md` and
  `docs/specs/native-winhttp-transport.md`.
- WinHTTP behavior: [option flags](https://learn.microsoft.com/en-us/windows/win32/winhttp/option-flags).

## Supersedes

- None

## Superseded by

- None
