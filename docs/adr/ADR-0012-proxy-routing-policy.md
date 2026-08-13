# ADR-0012: Explicit proxy routing policy

## Status

`accepted`

## Background

The two transports currently use different WinHTTP surfaces: the native
backend opens `winhttp.dll` sessions while the COM backend configures a
`WinHttpRequest` object. Corporate consumers need a deterministic way to
respect the OS WinHTTP configuration, bypass proxies completely for a local
call, or select a named HTTP proxy without depending on global machine state.
Proxy credentials are an authentication concern and must not be smuggled into
the routing value.

## Decision

- Add cloneable `HttpProxyOptions` to `HttpClient` and `HttpRequest`. The modes
  are `HttpProxyDefault`, `HttpProxyNoProxy`, and `HttpProxyManual`.
- Default mode delegates to the OS/WinHTTP proxy configuration. No-proxy mode
  uses direct connections. Manual mode requires an absolute HTTP or HTTPS
  proxy URL with no path, query, fragment, or user-info credentials and may
  include WinHTTP's semicolon-delimited `BypassList`.
- `HttpClient` owns the default options; a request-level value overrides it.
  Both are cloned into the execution snapshot so mutation after execution
  starts cannot change routing.
- The native backend maps the modes to `WINHTTP_ACCESS_TYPE_DEFAULT_PROXY`,
  `WINHTTP_ACCESS_TYPE_NO_PROXY`, and `WINHTTP_ACCESS_TYPE_NAMED_PROXY` when it
  calls `WinHttpOpen`. The COM backend calls `IWinHttpRequest::SetProxy` after
  `Open` and before request headers/body are sent.
- Proxy credentials, `Proxy-Authorization`, PAC/WPAD authoring, and SOCKS
  remain outside this routing slice. The deterministic HTTPS CONNECT boundary
  is defined separately by ADR-0031; trusted corporate CONNECT behavior still
  belongs to host-specific authentication/security evidence.

## Consequences

- Consumers can run deterministic local tests with `HttpProxyNoProxy` and can
  route both transports through an explicit loopback HTTP proxy.
- Manual proxy strings are intentionally less expressive than raw WinHTTP
  configuration; unsupported credential and PAC scenarios fail validation
  instead of creating hidden authentication behavior.
- The native session is recreated for each execution snapshot, so proxy mode
  is request-scoped and does not mutate global WinHTTP settings.
- HTTPS targets through a manual HTTP proxy are covered only by the loopback
  CONNECT boundary in ADR-0031; trusted corporate proxy behavior is not
  inferred from that fixture.

## Rationale

- Code: `src/classes/HttpProxyOptions.cls`, `HttpClient.cls`, `HttpRequest.cls`,
  `WinHttpComTransport.cls`, `WinHttpNativeTransport.cls`, and
  `WinHttpNativeApi.bas`.
- Tests: `src/modules/Tests/Unit/HttpProxyOptionsTests.bas`, request/client
  snapshot tests, COM/native loopback proxy integration tests, and the
  deterministic proxy in `tools/testserver/proxy.go`.
- Current contract: `docs/specs/proxy-policy.md` and
  `docs/specs/local-test-server.md`.
- WinHTTP references: [WinHttpRequest SetProxy](https://learn.microsoft.com/en-us/windows/win32/winhttp/iwinhttprequest-setproxy)
  and [WinHttpOpen access types](https://learn.microsoft.com/en-us/windows/win32/api/winhttp/nf-winhttp-winhttpopen).

## Supersedes

- None

## Superseded by

- None
