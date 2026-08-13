# Proxy policy

`HttpProxyOptions` is the shared routing contract for the COM and native
transports. It is available on both `HttpClient.ProxyOptions` and
`HttpRequest.ProxyOptions`; a request value overrides the client default and
is cloned into the execution snapshot.

```vb
Dim proxy As HttpProxyOptions
Set proxy = VBAHttp.CreateProxyOptions()
proxy.Mode = HttpProxyManual
proxy.ProxyUrl = "http://127.0.0.1:18080"
proxy.BypassList = "localhost;127.0.0.1"
Set client.ProxyOptions = proxy
```

## Modes

- `HttpProxyDefault` (the default) delegates to the current OS/WinHTTP
  configuration.
- `HttpProxyNoProxy` bypasses all proxy configuration and connects directly.
- `HttpProxyManual` requires `ProxyUrl` and passes it to the backend. The URL
  must be HTTP or HTTPS, contain a host, and contain no path, query, fragment,
  whitespace, or user-info credentials. `BypassList` is an optional
  semicolon-delimited WinHTTP bypass list and rejects control characters.

Default and no-proxy modes reject manual fields during `Validate`. Manual
credentials are intentionally rejected; future authentication providers own
credential material and `Proxy-Authorization`.

## Transport mapping

The COM backend calls `SetProxy` after `Open` and before headers/body. The
native backend maps the same mode to the WinHTTP session access type when it
calls `WinHttpOpen`. A proxy option never changes the public response/error
contract, and proxy URLs or credentials are not included in diagnostics.

This slice supports HTTP forwarding for deterministic loopback tests. A
separate authenticated loopback proxy verifies the `HttpAuthTargetProxy`
challenge path without changing the unauthenticated forwarding fixture. HTTPS
CONNECT, PAC/WPAD authoring, and SOCKS remain host-dependent compatibility work.

## Evidence

- Unit: `src/modules/Tests/Unit/HttpProxyOptionsTests.bas`,
  `HttpClientTests.bas`, and `HttpRequestTests.bas`.
- Integration: `WinHttpComTransportTests.Test_ComTransport_UsesManualProxy`
  and `WinHttpNativeTransportTests.Test_NativeTransport_UsesManualProxy`.
- Fixture: `tools/testserver/proxy.go`; `task testserver:test` covers its
  forwarding, target-isolation, and fixed Basic challenge behavior.
- Release/compatibility: `docs/specs/compatibility-matrix.md` and ADR-0012.
