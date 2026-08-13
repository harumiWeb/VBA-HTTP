# Buffered challenge authentication

## Public factory

```vb
Set client.AuthProvider = VBAHttp.CreateWindowsAuthProvider( _
    "user", "pass", HttpAuthSchemeBasic, HttpAuthTargetServer)
```

`HttpWindowsAuthProvider` accepts `Auto`, `Basic`, `Digest`, `Ntlm`, or
`Negotiate`, and a server or proxy target. Its default challenge limit is 3.
The URL must be HTTPS unless `AllowInsecureHttp:=True` is explicitly supplied
for a loopback fixture. Username/password values are validated for header
control characters and are never returned by the public API.

## Replay boundary

- The provider is cloned into the execution snapshot and applied exactly once.
- Automatic redirects are disabled for every challenge provider.
- Buffered COM requests configure credentials before `Send`; WinHTTP performs
  its own bounded challenge exchange.
- Native buffered requests query `WWW-Authenticate`/`Proxy-Authenticate`, set
  credentials on the existing handle, and resend the retained request body at
  most three times (or the provider's lower limit).
- Native GET/download requests are replayable because no request source is
  consumed. File and multipart uploads reject the provider before opening the
  source; callers receive the existing 401/407 result when no provider is set.
- A final 401/407 is a normal response. Unsupported schemes, target mismatch,
  and exhausted challenge attempts are not converted into a secret-bearing
  error description.

The deterministic proxy fixture starts on a separate `proxy-auth-listen`
listener. It requires the fixed loopback-only `proxy-user` / `proxy-pass`
credential, verifies `HttpAuthTargetProxy` for both transports, and leaves the
existing unauthenticated forwarding proxy unchanged. It does not claim HTTPS
CONNECT, PAC, or Windows-domain authentication support.

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
- Release consumer smoke invokes the factory without importing test code into
  the release workbook.
