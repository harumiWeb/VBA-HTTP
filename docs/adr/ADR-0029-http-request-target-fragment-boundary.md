# ADR-0029: HTTP request-target excludes URI fragments

## Status

`accepted`

## Background

URI fragments identify client-side representation state and are not part of an
HTTP request-target. The native parser already discarded fragments while
building its path, but the COM transport passed the original URL to
`WinHttpRequest.Open`, and a client execution snapshot could still expose a
fragment after query preparation. That made backend behavior dependent on
transport parsing and left custom transports with an ambiguous request URL.

## Decision

- Treat a raw `#` and the remainder of the URI as a local fragment, never as
  wire request data.
- After query parameters are merged, `HttpClient` removes the fragment from
  the execution snapshot before applying headers, cookies, or authentication.
- Direct COM transport calls remove the fragment from the URL passed to
  `WinHttpRequest.Open`; the native URL parser retains the same behavior.
- A percent-encoded `%23` remains ordinary query/path data and is not removed.
- The caller-owned `HttpRequest` is not mutated. A caller that needs a
  fragment for UI state must retain it separately from the request sent over
  HTTP.

## Consequences

- COM, native, and mock-prepared client requests use one deterministic
  request-target boundary.
- Query parameters are always inserted before a fragment and therefore remain
  part of the request URL after normalization.
- The public execution snapshot URL no longer contains client-only fragment
  state; this is a compatibility tightening for callers that inspected custom
  transport snapshots.
- Redirect, cookie, authentication, and diagnostics processing cannot
  accidentally treat fragment text as server-visible data.

## Rationale and evidence

- Shared normalization: `src/modules/HttpUrlValidation.bas`.
- Client and COM integration points: `src/classes/HttpClient.cls` and
  `src/classes/WinHttpComTransport.cls`.
- Native parsing boundary: `src/classes/WinHttpNativeUrl.cls`.
- Regression tests: `src/modules/Tests/Unit/HttpClientTests.bas` and
  `src/modules/Tests/Unit/HttpUrlValidationTests.bas`.
- Current API contract: `docs/specs/http-core-api.md`.

## Supersedes

- None

## Superseded by

- None
