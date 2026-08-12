# Cookie jar specification

## Public API

`VBAHttp.CreateCookieJar()` returns an empty `HttpCookieJar`. Set it on
`HttpClient.CookieJar` or `HttpRequest.CookieJar`; a request-level jar wins over
the client default. `HttpCookieJar.Count`, `Clear`, `GetCookieHeader(url)`, and
`StoreResponseCookies(url, headers)` are public for deterministic consumers and
tests. `GetCookieHeader` returns an empty string when no cookie matches.

The jar is mutable and caller-owned. The library uses it during request
preparation and response completion, but never clones or persists it. A caller
may use separate jars to isolate sessions.

## Parsing and matching

- Cookie name is a non-empty RFC token; values cannot contain semicolon, control,
  CR, or LF. Invalid `Set-Cookie` lines are ignored.
- `Domain` is normalized case-insensitively. A host-only cookie matches only the
  exact request host. A Domain cookie matches the domain itself or a dot-bound
  subdomain, never an IP's unrelated suffix; a Domain that does not match the
  origin is rejected.
- Without `Path`, default-path is `/` for a root or the request path up to the
  final slash. A cookie path matches the request path when equal, when it is a
  slash-prefix ending in `/`, or when the next request character is `/`.
- `Secure` cookies are sent only for `https`. `HttpOnly`, `SameSite`, and
  extension attributes are retained nowhere and do not change matching.
- `Max-Age` takes precedence over `Expires`; zero/negative deletes the matching
  name/domain/path. Positive `Max-Age` is evaluated against the jar clock.
  `Expires` accepts IMF-fixdate and the two legacy HTTP-date forms already used
  by retry parsing. Invalid dates leave the cookie as a session cookie.
- Matching order is longest path first, then creation order. The generated
  header is `name=value` pairs joined by `; `.

The store is capped at 300 cookies and 4096 bytes per serialized cookie. When a
new cookie exceeds the cap, validation fails before mutation; an expired cookie
is removed immediately.

## Client integration and redirect boundary

During `HttpClient.PrepareRequest`, the execution snapshot gets a cookie header
only if its own headers do not already contain `Cookie`. If a cookie is applied
or a caller supplied `Cookie`, `FollowRedirects` becomes false. After COM/native
execution returns, all `Set-Cookie` response values are offered to the selected
jar using the prepared request URL, including 3xx, 4xx, and 5xx responses.

The original `HttpRequest`, its headers, and its body remain unchanged. A direct
transport call bypasses jar integration by design; callers of `IHttpTransport`
must provide headers themselves.

## Security and diagnostics

Cookie values are secret material. No error, benchmark result, release manifest,
or diagnostic serializer may include them; `HttpSecurity.IsSensitiveHeader`
classifies `Cookie` and `Set-Cookie` for redaction. Cookies are never sent to an
HTTP URL when marked `Secure`, and no automatic redirect carrying cookies is
allowed.

## Verification

- Unit tests cover parsing, host/domain/path/secure matching, deletion, expiry,
  caps, clone/snapshot precedence, and caller-header preservation.
- The deterministic local server exposes `/cookie/set`, `/cookie/echo`, and
  `/cookie/redirect`; COM/native integration verifies opt-in persistence and
  redirect suppression without logging the header.
- Release consumer smoke sets and reads a fixed test cookie through the public
  factory; the release workbook contains `HttpCookieJar` but no test module.
