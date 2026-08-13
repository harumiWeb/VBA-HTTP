# Buffered challenge authentication

## Public factory

```vb
Set client.AuthProvider = VBAHttp.CreateWindowsAuthProvider( _
    "user", "pass", HttpAuthSchemeAuto, HttpAuthTargetServer)
```

Use `VBAHttp.CreateNativeClient()` when the provider must request a specific
scheme, for example `HttpAuthSchemeBasic`.

`HttpWindowsAuthProvider` accepts `Auto`, `Basic`, `Digest`, `Ntlm`, or
`Negotiate`, and a server or proxy target. The native transport enforces the
configured challenge limit (1–3); its default is 3. COM accepts the default
value for compatibility but delegates the actual exchange count to WinHTTP.
Any custom limit on COM fails validation before backend creation.
The URL must be HTTPS unless `AllowInsecureHttp:=True` is explicitly supplied
for a loopback fixture. Username/password values are validated for header
control characters and are never returned by the public API.

The `Scheme` value is a backend capability boundary:

- Native WinHTTP queries the challenge bitmap and honors an explicit `Basic`,
  `Digest`, `Ntlm`, or `Negotiate` selection. `Auto` uses the documented native
  preference order and only selects a scheme advertised by the peer.
- COM `IWinHttpRequest::SetCredentials` accepts the credential target but has no
  scheme parameter. The COM transport therefore accepts only `Auto` and lets
  WinHTTP negotiate the scheme advertised by the server or proxy. An explicit
  scheme on the COM path fails validation before a backend is created; callers
  that require deterministic scheme selection must use
  `VBAHttp.CreateNativeClient()`.

## Replay boundary

- The provider is cloned into the execution snapshot and applied exactly once.
- Automatic redirects are disabled for every challenge provider.
- Buffered COM requests configure credentials before `Send`; WinHTTP performs
  its own challenge exchange using the server/proxy's advertised
  scheme. The COM API does not expose a challenge-count callback, so the
  library does not claim a strict per-provider limit there; custom limits are
  rejected instead of silently ignored.
- Native buffered requests query `WWW-Authenticate`/`Proxy-Authenticate`, set
  credentials on the existing handle, and resend the retained request body at
  most three times (or the provider's lower limit).
- Native GET/download requests are replayable because no request source is
  consumed. File and multipart uploads reject the provider before opening the
  source; callers receive the existing 401/407 result when no provider is set.
- A final 401/407 is a normal response. Unsupported schemes, target mismatch,
  and exhausted challenge attempts are not converted into a secret-bearing
  error description.

The deterministic proxy fixture starts on separate `proxy-auth-listen` and
`proxy-tls-auth-listen` listeners. It requires the fixed loopback-only
`proxy-user` / `proxy-pass` credential, verifies `HttpAuthTargetProxy` for both
transports, and leaves the existing unauthenticated forwarding proxy
unchanged. The TLS listener proves only that an authenticated CONNECT reaches
the intentionally untrusted origin before normal `HttpErrorTls` validation;
it does not claim trusted HTTPS CONNECT, PAC, or Windows-domain authentication
support. See `docs/specs/https-connect-proxy.md`.

## Security and compatibility

Credentials are sent only to the configured server/proxy target. Redirects,
diagnostics, benchmark JSON, and error descriptions must not contain raw
`Authorization`, `Proxy-Authorization`, `WWW-Authenticate`, or
`Proxy-Authenticate` values. Basic/Digest require TLS in production; NTLM and
Negotiate retain WinHTTP keep-alive behavior. Proxy CONNECT and real Windows
domain fixtures are host-dependent and remain separate compatibility evidence.

## Evidence

- Unit provider validation/snapshot tests: `HttpAuthTests.bas`.
- Deterministic loopback Basic challenge: `/auth/challenge/basic` in
  `tools/testserver/server.go` and COM/native integration tests.
- Deterministic proxy Basic challenge: `proxy_auth_url` from
  `tools/testserver/main.go`, COM/native proxy-target tests, and release smoke.
- Deterministic CONNECT boundary: `proxy_tls_url`, `proxy_tls_target_url`, and
  `proxy_tls_auth_url` from `tools/testserver/main.go`, with COM/native TLS
  rejection and `/__admin/proxy-stats` assertions.
- Release consumer smoke invokes the factory without importing test code into
  the release workbook.
