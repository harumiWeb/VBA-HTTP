# ADR-0034: COM challenge-limit boundary

## Status

`accepted`

## Context

The native transport owns its challenge loop: it receives a 401/407 response,
queries the advertised schemes, sets credentials on the same request handle,
resends the retained buffered body, and stops at `HttpWindowsAuthProvider.MaxChallenges`.

The late-bound COM transport calls `IWinHttpRequest::SetCredentials` before
`Send`. The COM API exposes the server/proxy target flag but no callback or
per-request option that lets VBA observe and count each authentication
exchange. WinHTTP challenge-response schemes can require multiple transactions.

Claiming that the same `MaxChallenges` value is enforced by both backends would
make the public reliability contract backend-dependent without an observable
failure.

## Decision

- Native WinHTTP strictly enforces `MaxChallenges` in the buffered replay loop.
- COM accepts the historical default value of `3` for source compatibility, but
  delegates the actual challenge exchange to WinHTTP and does not claim a
  strict per-provider count.
- A COM request with a custom limit (1 or 2) fails with `HttpErrorValidation`
  before the COM backend is created. Callers requiring a strict bound must use
  `VBAHttp.CreateNativeClient()`.
- The existing COM scheme boundary remains in force: COM accepts
  `HttpAuthSchemeAuto`; explicit scheme selection is native-only under
  ADR-0033.
- Final 401/407 responses, HTTPS-by-default, redirect suppression, target
  isolation, streaming upload rejection, and redaction semantics remain
  unchanged.

## Consequences

- No caller can silently configure a challenge limit that the COM backend
  ignores.
- COM remains available for the common Auto/default loopback and corporate
  scenarios, while strict replay policy is explicit and testable on native.
- A future COM implementation that manually loops `Open`/`Send`/status checks
  may supersede this decision, but it must cover synchronous, async batch, and
  cancellation/deadline paths before removing the validation boundary.

## Evidence

- COM setup and validation: `src/classes/WinHttpComTransport.cls` and
  `src/classes/HttpWindowsAuthProvider.cls`.
- Native bounded loop: `src/classes/WinHttpNativeTransport.cls`.
- Unit contract: `src/modules/Tests/Unit/HttpAuthTests.bas`.
- COM/native integration and release smoke: the challenge tests and
  `tools/consumer/ReleaseBatchSmoke.bas`.
- Official API contracts: [IWinHttpRequest::SetCredentials](https://learn.microsoft.com/en-us/windows/win32/winhttp/iwinhttprequest-setcredentials)
  and [Authentication in WinHTTP](https://learn.microsoft.com/en-us/windows/win32/winhttp/authentication-in-winhttp).

## Supersedes

- None

## Superseded by

- None
