# ADR-0014: Redirect security boundary

## Status

`accepted`

## Background

WinHTTP automatically follows redirects unless the caller disables them. A
user-defined `Authorization`, `Proxy-Authorization`, or `Cookie` header can be
carried to the redirected request, while the project deliberately does not
install native callbacks that would inspect and rewrite headers. The transports
also need one deterministic maximum so a redirect loop cannot stall a caller.

## Decision

- The prepared request is inspected after query and default-header merging. If
  it contains `Authorization`, `Proxy-Authorization`, or `Cookie`, automatic
  redirects are disabled for that execution snapshot. The caller receives the
  3xx response and must perform an origin-aware follow-up after deciding which
  headers are safe to copy.
- Both transports enforce this boundary as well as `HttpClient`, so a direct
  transport caller cannot accidentally bypass it. The original `HttpRequest`
  is never mutated.
- `MaxRedirects` remains an integer in the range 1..100 and is passed to both
  WinHTTP backends. Exceeding it is a protocol failure (`ERROR_WINHTTP_REDIRECT_FAILED`,
  category `HttpErrorProtocol`), not an unbounded retry.
- HTTPS-to-HTTP redirects remain disabled by default. COM sets
  `WinHttpRequestOption_EnableHttpsToHttpRedirects` to `False`; native WinHTTP
  uses `WINHTTP_OPTION_REDIRECT_POLICY_DISALLOW_HTTPS_TO_HTTP` (value 1).
- No ambient cookie persistence or implicit redirect state is introduced.
  Explicit caller-owned cookie state is defined separately by ADR-0018; a
  caller Cookie header or jar-applied Cookie header still forces the redirect
  boundary defined here.

## Consequences

- Credential-bearing requests do not silently leak headers across origins, but
  callers that need a redirect must handle the 3xx response themselves.
- A caller can still follow a redirect safely by constructing a new request and
  explicitly selecting headers for the new origin.
- The policy is compatible with synchronous COM, async COM batch, native
  buffered, and native streaming paths because it is enforced at the request
  snapshot and transport option boundaries.
- Cookie parsing/storage is intentionally separate from this redirect decision;
  cross-origin rewriting and interactive authentication remain out of scope.

## Rationale

- Code: `HttpSecurity.bas`, `HttpClient.cls`, `WinHttpComTransport.cls`,
  `WinHttpNativeTransport.cls`, and `WinHttpNativeApi.bas`.
- Tests: `HttpClientTests.bas`, COM/native sensitive-header and loopback
  redirect integration tests, and `tools/testserver/server_test.go`.
- Current contract: `docs/specs/redirect-policy.md`,
  `buffered-com-transport.md`, `native-winhttp-transport.md`, and
  `auth-policy.md`.
- Security references: [WinHTTP option flags](https://learn.microsoft.com/en-us/windows/win32/winhttp/option-flags),
  [WinHttpRequestOption](https://learn.microsoft.com/en-us/windows/win32/winhttp/winhttprequestoption),
  and [WinHTTP security considerations](https://learn.microsoft.com/en-us/windows/win32/winhttp/winhttp-security-considerations).

## Supersedes

- None

## Superseded by

- ADR-0018 (cookie-state portion only)
