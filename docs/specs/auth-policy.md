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
automatic redirect unchanged; callers receive the 3xx response and must apply
an origin-aware redirect policy before issuing a follow-up request.

401 and 407 are returned as normal responses. There is no automatic challenge
replay, credential refresh, or proxy authentication in this slice. Streaming
uploads remain one-shot and must never be replayed implicitly.

## Redaction

`HttpSecurity.RedactHeaderValue` returns `[REDACTED]` for Authorization,
Proxy-Authorization, Cookie, Set-Cookie, WWW-Authenticate, and
Proxy-Authenticate. Diagnostics and benchmark writers must call this helper
before serializing headers. Public error descriptions never include provider
secrets or raw sensitive header values.

## Deferred capabilities

Windows integrated Negotiate/NTLM, Digest, proxy credentials, OAuth flows, and
interactive callbacks are not implemented here. A future native capability must
define bounded challenge replay, source reset semantics for uploads, keep-alive
requirements, and 401/407 loop limits in a separate ADR.

## Evidence

- Unit: `src/modules/Tests/Unit/HttpAuthTests.bas`, `HttpClientTests.bas`,
  `HttpRequestTests.bas`.
- Integration: `WinHttpComTransportTests` and
  `WinHttpNativeTransportTests` auth cases against `/auth/basic` and
  `/auth/bearer`.
- Release: `tools/consumer/ReleaseBatchSmoke.bas` invokes Basic and Bearer
  providers without injecting test code into the release workbook.
