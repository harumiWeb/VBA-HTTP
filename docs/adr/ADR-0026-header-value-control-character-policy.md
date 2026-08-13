# ADR-0026: Header value control-character policy

## Status

`accepted`

## Context

`HttpHeaders` is the shared input boundary for the COM transport, native
WinHTTP transport, mock transport, cookie jar, authentication providers, and
structured diagnostics. Its previous value validation rejected CR and LF but
accepted other C0 controls, NUL, and DEL. Those bytes are not valid visible
header content and can be interpreted differently by serializers, logs, or
intermediaries.

## Decision

- Reject U+0000 through U+001F except horizontal tab (U+0009), and reject
  U+007F, from every header value added through `HttpHeaders.Add` or
  `HttpHeaders.SetValue`.
- Preserve the existing dedicated CR/LF validation description for compatible
  diagnostics; other rejected controls use the stable validation description
  `Header value cannot contain control characters.`
- Allow HTAB, visible ASCII, and Unicode/obs-text values. This policy checks
  control characters only; it does not silently transcode or normalize values.
- Apply the rule before any transport, cookie, authentication, or diagnostics
  code serializes the headers. Direct transport callers remain responsible for
  using the same `HttpHeaders` contract.

## Consequences

- Malformed values that were previously accepted now fail deterministically as
  `HttpErrValidation`, closing a control-byte injection boundary shared by both
  backends.
- HTAB remains available for interoperable field values, while CR/LF and all
  other C0/DEL controls cannot reach wire or diagnostics serializers.
- This is a compatibility tightening: callers that intentionally used control
  bytes must remove them or encode them as data before adding a header.

## Evidence

- Implementation: `src/classes/HttpHeaders.cls`.
- Regression tests: `src/modules/Tests/Unit/HttpHeadersTests.bas`.
- Current contract: `docs/specs/http-core-api.md`.

## Supersedes

- None

## Superseded by

- None
