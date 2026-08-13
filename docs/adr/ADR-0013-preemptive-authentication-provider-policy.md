# ADR-0013: Preemptive header authentication provider policy

## Status

`superseded`

## Background

The HTTP core currently has no credential ownership boundary. Callers can add
an `Authorization` header manually, but that makes credential-bearing requests
easy to copy, log, or accidentally redirect to another origin. WinHTTP's
challenge APIs (`WinHttpQueryAuthSchemes` and `WinHttpSetCredentials`) also
require bounded request replay, while streaming upload sources are one-shot and
the native callback boundary intentionally forbids callbacks into VBA logic.

The first authentication slice therefore needs a deterministic, backend-
independent contract for credentials that can be applied before a request is
sent, without pretending to implement Windows integrated authentication.

## Decision

- Add `IHttpAuthProvider` with `Apply(HttpRequest)` and `Clone()` methods.
  Providers are applied only to the execution snapshot after query and default
  header merging; caller-owned requests and providers are never mutated.
- Implement immutable, header-only `HttpBasicAuthProvider` and
  `HttpBearerAuthProvider` factories. Basic encodes the UTF-8 `username:password`
  octets with RFC 4648 Base64; Bearer preserves the opaque token after strict
  validation. Secret getters are not exposed.
- Basic and Bearer require an `https://` request URL by default. An explicit
  `AllowInsecureHttp` factory argument exists only for controlled loopback tests
  and must be visible in caller code.
- A provider rejects a pre-existing `Authorization` header rather than choosing
  an implicit precedence. Provider-generated credentials disable automatic
  redirects and return the 3xx response to the caller; redirect follow-up is a
  separate, origin-aware policy.
- 401 and 407 responses remain ordinary `HttpResponse` values. This slice does
  not replay buffered requests, configure proxy credentials, or challenge-retry
  one-shot uploads.
- Windows integrated authentication (Negotiate/NTLM), Digest, proxy
  authentication, and interactive credential callbacks remain a later native
  capability. That capability must use bounded replay and the same secret
  redaction contract before it is exposed.
- All diagnostics-facing header values use `HttpSecurity.RedactHeaderValue`.
  Sensitive names include Authorization, Proxy-Authorization, Cookie,
  Set-Cookie, WWW-Authenticate, and Proxy-Authenticate; values are never
  included in public errors, benchmark output, or test-server responses.

## Consequences

- Basic/Bearer work identically with the COM and native buffered transports and
  can also be attached to native streaming calls without enabling unsafe
  challenge replay.
- Callers must explicitly opt into insecure loopback testing and must handle a
  3xx response themselves when credentials are present.
- The API has no automatic credential refresh or integrated Windows login yet;
  those features require a new ADR rather than silently widening this provider.
- Provider objects are immutable after construction from the public factory,
  and execution snapshots isolate them from later caller mutation.

## Rationale

- Code: `IHttpAuthProvider.cls`, `HttpBasicAuthProvider.cls`,
  `HttpBearerAuthProvider.cls`, `HttpClient.cls`, `HttpRequest.cls`,
  `HttpEncoding.bas`, and `HttpSecurity.bas`.
- Tests: `HttpAuthTests.bas`, client/request snapshot tests, COM/native loopback
  auth integration tests, and the external release consumer smoke harness.
- Current contract: `docs/specs/auth-policy.md`, `http-core-api.md`,
  `buffered-com-transport.md`, and `streaming-upload.md`.
- Security references: [WinHTTP Security Considerations](https://learn.microsoft.com/en-us/windows/win32/winhttp/winhttp-security-considerations),
  [Authentication in WinHTTP](https://learn.microsoft.com/en-us/windows/win32/winhttp/authentication-in-winhttp),
  [RFC 7617](https://datatracker.ietf.org/doc/html/rfc7617), and
  [RFC 6750](https://datatracker.ietf.org/doc/html/rfc6750).

## Supersedes

- None

## Superseded by

- ADR-0023: Bounded buffered challenge authentication
