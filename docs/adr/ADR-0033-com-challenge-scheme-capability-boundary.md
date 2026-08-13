# ADR-0033: COM challenge-scheme capability boundary

## Status

`accepted`

## Context

`HttpWindowsAuthProvider` exposes `Auto`, `Basic`, `Digest`, `Ntlm`, and
`Negotiate` so a buffered request can choose a challenge scheme. The native
WinHTTP transport can query the peer's supported bitmap and pass an explicit
scheme to `WinHttpSetCredentials`. The late-bound COM transport uses
`IWinHttpRequest::SetCredentials`, whose public signature accepts only username,
password, and a server/proxy target flag; it has no scheme argument.

This follows the Windows API contracts for
[`IWinHttpRequest::SetCredentials`](https://learn.microsoft.com/en-us/windows/win32/winhttp/iwinhttprequest-setcredentials)
and native
[`WinHttpSetCredentials`](https://learn.microsoft.com/en-us/windows/win32/api/winhttp/nf-winhttp-winhttpsetcredentials).

Treating the same option as if both backends selected a scheme would silently
change authentication behavior when a caller switches transports.

## Decision

- The COM transport accepts `HttpAuthSchemeAuto` for
  `HttpWindowsAuthProvider`. WinHTTP chooses among the schemes advertised by
  the server or proxy during its bounded challenge exchange.
- An explicit `Basic`, `Digest`, `Ntlm`, or `Negotiate` selection is native-only.
  The COM path raises `HttpErrorValidation` before creating the COM backend, so
  no network request or credential exchange occurs under an ambiguous contract.
- The native transport retains explicit scheme selection and validates that the
  selected scheme is present in the peer's challenge bitmap before calling
  `WinHttpSetCredentials`.
- The public target (`server` or `proxy`), HTTPS-by-default rule, challenge
  limit, redirect suppression, redaction, and final 401/407 response semantics
  are shared by both transports.

## Consequences

- Default COM consumers remain portable by using `Auto`; existing native users
  can request deterministic scheme selection.
- A caller that requires Basic specifically must select the native client even
  when the target happens to advertise Basic. This is an intentional fail-fast
  capability boundary rather than a backend-dependent preference.
- COM integration and release smoke use `Auto`; native integration continues to
  exercise explicit Basic selection. Real Windows-domain and trusted corporate
  proxy evidence remain separate host-dependent gates.
- No credentials are logged or included in validation descriptions.

## Evidence

- COM boundary: `src/classes/WinHttpComTransport.cls` and
  `src/classes/HttpWindowsAuthProvider.cls`.
- Native selection: `src/classes/WinHttpNativeTransport.cls` and
  `src/modules/WinHttpNativeApi.bas`.
- Unit contract: `src/modules/Tests/Unit/HttpAuthTests.bas`.
- COM/native loopback and release smoke: the challenge tests in
  `WinHttpComTransportTests.bas`, `WinHttpNativeTransportTests.bas`, and
  `tools/consumer/ReleaseBatchSmoke.bas`.
- Current specification: `docs/specs/challenge-authentication.md`.

## Supersedes

- None

## Superseded by

- None
