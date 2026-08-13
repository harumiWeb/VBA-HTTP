# Authentication policy

## Provider boundary

`HttpClient.AuthProvider` and `HttpRequest.AuthProvider` accept an
`IHttpAuthProvider`. The request-level provider overrides the client default.
Both values are cloned into the execution snapshot. `HttpClient` invokes the
provider exactly once after URL/query resolution and default-header merging;
direct transport callers must supply their own headers and must not expect a
transport to retain or replay a provider.

The initial factories are:

```vb
Set client.AuthProvider = VBAHttp.CreateBasicAuthProvider("user", "pass")
Set client.AuthProvider = VBAHttp.CreateBearerAuthProvider(token)
```

The factories accept an optional `AllowInsecureHttp` Boolean. Its default is
`False`; only deterministic loopback tests should pass `True`.

## Basic and Bearer

- Basic rejects an empty username, a colon in the username, and control/header
  delimiter characters. The UTF-8 bytes of `username:password` are Base64
  encoded without line breaks and emitted as `Authorization: Basic ...`.
- Bearer rejects an empty token, whitespace, and control characters, then emits
  `Authorization: Bearer <opaque-token>` without transforming the token.
- A provider rejects any existing case-insensitive `Authorization` header. This
  includes a value inherited from `HttpClient.DefaultHeaders`.
- Credential-bearing requests require an absolute HTTPS URL after client base
  URL resolution unless the provider was explicitly created with insecure HTTP
  enabled.

## Redirects and challenge responses

Provider-generated credentials set the execution snapshot's
`FollowRedirects=False`. WinHTTP user-defined headers can otherwise cross an
automatic redirect unchanged. The same boundary applies to caller-supplied
`Authorization`, `Proxy-Authorization`, and `Cookie` headers. Callers receive
the 3xx response and must apply an origin-aware redirect policy before issuing a
follow-up request; see `redirect-policy.md`.

401 and 407 are returned as normal responses for preemptive providers.
Bounded buffered challenge authentication is an explicit
`HttpWindowsAuthProvider` capability documented in
`challenge-authentication.md`. COM uses `HttpAuthSchemeAuto`; explicit
challenge scheme selection is native-only and fails before COM backend
creation (ADR-0033). `MaxChallenges` is a strict native replay bound; COM
accepts only the default value and delegates the exchange count to WinHTTP,
rejecting custom limits before creation (ADR-0034). Streaming uploads remain
one-shot and must never be replayed implicitly.

## Redaction

`HttpSecurity.RedactHeaderValue` returns `[REDACTED]` for Authorization,
Proxy-Authorization, Cookie, Set-Cookie, WWW-Authenticate, and
Proxy-Authenticate. Diagnostics and benchmark writers must call this helper
before serializing headers. Public error descriptions never include provider
secrets or raw sensitive header values.

## Deferred capabilities

OAuth flows and interactive callbacks remain out of scope. Windows
Negotiate/NTLM, Digest, and proxy credentials are available only through the
bounded buffered provider contract in `challenge-authentication.md`; source
reset for streaming uploads and host-specific CONNECT/domain evidence remain
deferred.

## Evidence

- Unit: `src/modules/Tests/Unit/HttpAuthTests.bas`, `HttpClientTests.bas`,
  `HttpRequestTests.bas`.
- Integration: `WinHttpComTransportTests` and
  `WinHttpNativeTransportTests` auth cases against `/auth/basic` and
  `/auth/bearer`.
- Release: `tools/consumer/ReleaseBatchSmoke.bas` invokes Basic and Bearer
  providers without injecting test code into the release workbook.
