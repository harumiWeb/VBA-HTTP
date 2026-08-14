# Security, authentication, and state

## TLS and URL safety

Use HTTPS for credentials and sensitive data. The library relies on the
operating system's certificate validation and provides no certificate-ignore
switch. An untrusted certificate maps to `HttpErrorTls` in both supported
backends. User-info in URLs is rejected to prevent accidental credential
leakage.

Basic and Bearer providers require an absolute HTTPS execution URL by default.
`AllowInsecureHttp=True` is an explicit escape intended only for controlled
loopback tests; it does not make arbitrary HTTP safe.

## Basic and Bearer providers

```vb
Dim auth As IHttpAuthProvider

Set auth = VBAHttp.CreateBasicAuthProvider("user", "pass")
Set client.AuthProvider = auth
Set response = client.GetResponse("https://api.example.test/private")
```

Basic encodes UTF-8 `username:password` using Base64. The username cannot be
empty or contain a colon/control delimiter. Bearer tokens are opaque ASCII
values and reject whitespace/control characters. A provider rejects a
pre-existing `Authorization` header so precedence is never ambiguous.

Providers are cloned into the execution snapshot and have no public secret
getter. They run once after URL resolution and default-header merging. 401 and
407 are ordinary responses for preemptive providers; there is no hidden replay.

## Buffered challenge authentication

`CreateWindowsAuthProvider` enables bounded WinHTTP challenge authentication
for buffered requests. It can target the server or proxy and can select the
supported scheme through the native transport. COM uses `HttpAuthSchemeAuto`
and delegates the exchange to WinHTTP; explicit scheme selection and custom
challenge limits are rejected before COM backend creation. Streaming uploads
never replay a one-shot source and reject this provider before opening it.

Real Windows-domain credentials and trusted corporate proxy fixtures are
host-dependent. Do not use a challenge provider as a substitute for OAuth,
interactive login, or a secret store.

## Redirects and sensitive headers

WinHTTP can forward user-defined headers across redirects. To avoid sending
credentials or session state to a different origin, automatic redirects are
disabled when the execution snapshot contains `Authorization`,
`Proxy-Authorization`, or `Cookie`, including values injected by a provider or
cookie jar. HTTPS-to-HTTP downgrade redirects are rejected by default.

When a 3xx response is returned, validate the target origin, scheme, and
credential policy in application code before creating a new request.

## Cookie jar

Cookies are opt-in, in-memory, and caller-owned:

```vb
Dim jar As HttpCookieJar
Set jar = VBAHttp.CreateCookieJar()
Set client.CookieJar = jar
Set response = client.GetResponse("https://api.example.test/login")
Debug.Print jar.Count
```

The jar applies host-only/domain, path, secure, expiry, size, and creation-order
rules. An explicit caller `Cookie` header wins over jar-generated values. The
jar is never ambient WinHTTP state and is not persisted automatically. Cookie
values are secret material and are always redacted in diagnostics.

## Diagnostics and redaction

`HttpDiagnostics` is disabled by default and records bounded top-level events,
not a wire trace. Targets omit query values and user-info; bodies and reason
phrases are not recorded. Sensitive headers (`Authorization`,
`Proxy-Authorization`, `Cookie`, `Set-Cookie`, `WWW-Authenticate`, and
`Proxy-Authenticate`) are serialized as `[REDACTED]`.

```vb
Dim diagnostics As HttpDiagnostics
Set diagnostics = VBAHttp.CreateDiagnostics()
diagnostics.Enabled = True
Set client.Diagnostics = diagnostics
Set response = client.GetResponse("https://api.example.test/data")
Debug.Print diagnostics.ToJson
```

Do not concatenate raw URLs, credentials, or headers into application logs.
Use `HttpSecurity.RedactHeaderValue` when a custom diagnostic sink must render a
header.

## Replay and ownership rules

- POST/PATCH and streaming bodies are not retried by default.
- Buffered challenge replay is bounded by provider policy.
- A source file is opened read-only and never mutated by uploads.
- A download destination is published only after a complete successful write.
- Native handles and temporary files are cleaned up on cancellation, timeout,
  callback failure, and transport failure.
- The library does not use WinHTTP callbacks into VBA, executable memory,
  process injection, or global cookie/credential state.

Normative details are in [`../specs/auth-policy.md`](../specs/auth-policy.md),
[`../specs/challenge-authentication.md`](../specs/challenge-authentication.md),
[`../specs/redirect-policy.md`](../specs/redirect-policy.md),
[`../specs/cookie-policy.md`](../specs/cookie-policy.md), and
[`../specs/diagnostics-policy.md`](../specs/diagnostics-policy.md).
