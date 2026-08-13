# ADR-0027: URL user-info boundary

## Status

`accepted`

## Background

HTTP URLs can contain an authority user-info component such as
`https://user:password@example.test/`.  It is ambiguous whether a backend
will treat that component as credentials, discard it, or include it in a
redirect/proxy/logging path.  The native parser already rejected `@`, while
the COM path and relative URL resolution did not share that boundary.  That
backend difference could leak credentials or make the same request behave
differently depending on transport selection.

## Decision

- Reject an `@` in the authority of every absolute HTTP(S) URL accepted by
  `HttpClient.BaseUrl`, absolute request resolution, and the COM transport
  boundary.  The native parser keeps the same rejection contract.
- Validate the resolved absolute URL after joining a relative request with
  `BaseUrl`, before query/default-header preparation or transport I/O.
- Reject user-info when `HttpCookieJar` parses its public URL input; the jar
  never strips or interprets embedded credentials.
- Keep `@` legal in path and query data when it is outside the authority.
- Callers that need credentials use `IHttpAuthProvider`; this decision does
  not add implicit Basic authentication or URL decoding.

## Consequences

- COM, native, mock-prepared client requests, and cookie matching fail with
  the same stable invalid-URL category before network I/O.
- Existing callers that embedded credentials in URLs must migrate to an auth
  provider; this is an intentional security-compatible breaking validation.
- User-info is removed from neither diagnostics nor a backend request because
  it is rejected before either can observe it.  Query/path `@` characters are
  unaffected.
- Direct transport callers are still responsible for supplying a valid
  execution request, but both concrete WinHTTP transports enforce the same
  rejection at their entry boundary.

## Rationale and evidence

- Shared check: `src/modules/HttpUrlValidation.bas`.
- Client and COM boundaries: `src/classes/HttpClient.cls` and
  `src/classes/WinHttpComTransport.cls`.
- Native boundary: `src/classes/WinHttpNativeUrl.cls`.
- Cookie boundary: `src/modules/HttpCookieParsing.bas`.
- Regression coverage: `HttpClientTests`, `HttpCookieJarTests`, and
  `WinHttpNativeTests` user-info cases.
- Current URL contract: `docs/specs/http-core-api.md`.

## Supersedes

- None

## Superseded by

- None
