# Redirect policy

## Scope

This specification defines automatic redirect behavior for the COM and native
WinHTTP transports. It does not define a cookie jar or an authentication
challenge replay mechanism.

## Request snapshot boundary

`HttpClient` resolves the URL, merges query/default headers, and clones the
request before execution. If the resulting headers contain any of the
case-insensitive names below, the snapshot's `FollowRedirects` is forced to
`False`:

- `Authorization`
- `Proxy-Authorization`
- `Cookie`

The same check is repeated by each production transport for direct transport
callers. The source request and its headers remain unchanged. A 3xx response is
returned to the caller, who must create an origin-aware follow-up request.

## Bounds and failure

`HttpRequest.MaxRedirects` defaults to 10 and accepts 1 through 100. Both COM
and native transports pass this bound to WinHTTP. A finite redirect chain that
exceeds the bound, including `/redirect-loop` in the deterministic test server,
fails with the shared protocol category rather than looping indefinitely.

## Scheme downgrade

Automatic HTTPS-to-HTTP redirects are rejected by default. The COM backend sets
`WinHttpRequestOption_EnableHttpsToHttpRedirects` to `False`. The native backend
sets `WINHTTP_OPTION_REDIRECT_POLICY` to `DISALLOW_HTTPS_TO_HTTP` when redirects
are enabled and to `NEVER` when they are disabled. No certificate-ignore option
or insecure override is exposed.

## Cookies

There is no automatic cookie persistence. `Cookie` headers are caller-owned and
are redirect-sensitive; `Set-Cookie` values are exposed as response headers but
are not stored or replayed. A future cookie-jar feature requires an independent
API and deterministic domain/path/expiry/secure policy.

## Verification

- Unit tests assert that each redirect-sensitive request header disables
  redirects on the execution snapshot.
- COM and native loopback tests assert 302 is returned for those headers and
  that a two-hop bound rejects `/redirect-loop` as `HttpErrorProtocol`.
- Go testserver tests assert that `/redirect-loop` is deterministic.
