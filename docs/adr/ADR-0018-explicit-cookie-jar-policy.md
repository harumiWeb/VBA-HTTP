# ADR-0018: Explicit cookie jar policy

## Status

`accepted`

## Background

VBA-HTTP currently exposes `Set-Cookie` response headers but never persists or
replays them. Relying on WinHTTP ambient cookie state would make behavior
backend- and machine-dependent, and silently forwarding cookies across an
automatic redirect would create a cross-origin credential boundary that the
native callback policy deliberately does not support.

## Decision

- Add an explicit `HttpCookieJar` owned by the caller. A jar is opt-in through
  `HttpClient.CookieJar`; a client with no jar neither stores nor adds cookies.
- Snapshot the client/request jar reference at request preparation. A request
  jar overrides the client jar. The jar itself is mutable session state and is
  not cloned for every request; callers that need isolation create separate
  jars. The caller-owned request is never mutated while applying cookies.
- Store host-only cookies by default. Accept a valid `Domain` attribute only for
  an origin host suffix that is not a public-suffix-like bare suffix; never send
  a cookie to a different host. Implement RFC default-path matching, `Secure`,
  `Max-Age`, and `Expires`; a zero/negative lifetime deletes the matching
  cookie. Ignore unknown attributes and reject malformed names/values.
- Apply cookies only to HTTP(S) requests, only when domain/path/secure rules
  match, and only when the caller did not already supply a `Cookie` header.
  Caller-supplied `Cookie` remains authoritative. Cookie values are not exposed
  by diagnostics or benchmark serializers.
- Process `Set-Cookie` response headers after each completed response, including
  non-2xx responses. The jar uses the request URL and an injected/current UTC
  clock for expiry; an invalid `Expires` value does not create a persistent
  expiry. A bounded in-memory store rejects more than 300 cookies or a single
  cookie larger than 4096 bytes.
- Any request with a caller Cookie header or a jar-applied Cookie header forces
  `FollowRedirects=False`. Callers must inspect 3xx responses and construct an
  origin-aware follow-up. This preserves the existing redirect security policy
  without native callbacks.
- Cookie jars are not shared globally, serialized, or used for authentication
  challenge replay. Public-suffix lists, partitioned cookies, SameSite policy,
  and persistent disk storage are out of scope for this slice.

## Consequences

- COM and native transports receive the same prepared `Cookie` header and have
  deterministic behavior on the local server; no ambient WinHTTP cookie state
  is used.
- A jar is a mutable state boundary, so concurrent requests using one client
  are serialized by the existing client execution guard; separate clients or
  jars are required for independent sessions.
- The implementation is intentionally conservative around Domain and expiry;
  unsupported browser features are rejected or ignored rather than guessed.
- Cookie values remain secret material. Tests assert only server-side markers,
  not raw credential-bearing headers in output.

## Rationale

- Current response ownership: `src/classes/HttpResponse.cls` and
  `src/classes/HttpHeaders.cls`.
- Request preparation boundary: `src/classes/HttpClient.cls` and
  `src/classes/HttpRequest.cls`.
- Redirect security: `docs/adr/ADR-0014-redirect-security-boundary.md`.
- Implementation/spec/tests: `src/classes/HttpCookieJar.cls`,
  `src/modules/HttpCookieParsing.bas`, `docs/specs/cookie-policy.md`, and
  `src/modules/Tests/Unit/HttpCookieJarTests.bas`.

## Supersedes

- None

## Superseded by

- None
