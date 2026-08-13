# ADR-0023: Bounded buffered challenge authentication

## Status

`accepted`

## Context

The preemptive Basic/Bearer provider deliberately leaves 401/407 responses
untouched. WinHTTP can negotiate Basic, Digest, NTLM, and Negotiate through its
challenge APIs, but a challenge may require resending the same request. An
arbitrary streaming upload source is one-shot and cannot be replayed safely;
native callbacks into VBA application logic are also prohibited by ADR-0006.

## Decision

- Add an immutable `HttpWindowsAuthProvider` factory with an explicit target
  (server or proxy), preferred WinHTTP scheme (Auto, Basic, Digest, NTLM, or
  Negotiate), bounded challenge count, and an explicit insecure-loopback opt-in.
- The provider is an `IHttpAuthProvider` execution-snapshot policy. It adds no
  header itself, rejects conflicting caller credentials, disables automatic
  redirects, and never exposes username/password getters to consumers,
  diagnostics, errors, or benchmark output.
- The COM transport configures credentials on each fresh buffered request
  before `Send`; WinHTTP owns the challenge exchange and keep-alive details.
  COM async batch operations use the same configuration path.
- The native buffered transport performs at most the provider's bounded number
  of challenge resends on the same request handle. It queries the server's
  supported schemes, selects the configured/first supported scheme, calls
  `WinHttpSetCredentials`, and resends the retained buffered body.
- Native GET/download requests may use the same bounded replay because their
  request body is empty. Native file and multipart uploads reject a challenge
  provider before opening a source; a 401/407 without a provider remains a
  normal `HttpUploadResult.AuthenticationChallenged` result.
- A final 401/407 is returned as an ordinary response/result. Challenge loops,
  unsupported schemes, or provider target mismatches do not silently fall back
  to a different credential target. No interactive callbacks, OAuth flow, or
  proxy URL credentials are added.

## Consequences

- Windows integrated and Digest authentication becomes available for buffered
  requests without weakening the streaming ownership contract.
- WinHTTP's native/COM implementations remain backend-specific internally but
  share the public provider snapshot, redirect, redaction, and final-response
  semantics.
- Credentials may remain in a process-owned provider until its snapshot is
  released; VBA offers no reliable memory zeroization. Callers must still use
  TLS in production and pass `AllowInsecureHttp:=True` only for loopback tests.
- Proxy challenge integration uses the separate deterministic listeners defined
  by ADR-0024 and ADR-0031. The CONNECT fixture proves only a loopback tunnel
  followed by normal untrusted-certificate rejection; trusted host-specific
  proxy/CONNECT and integrated credential evidence remain compatibility
  follow-up work.

## Evidence

- Code: `src/classes/HttpWindowsAuthProvider.cls`,
  `src/classes/WinHttpComTransport.cls`, `src/classes/WinHttpNativeTransport.cls`,
  and `src/modules/WinHttpNativeApi.bas`.
- Unit: `src/modules/Tests/Unit/HttpAuthTests.bas`.
- Loopback integration: `WinHttpComTransportTests` and
  `WinHttpNativeTransportTests` against `/auth/challenge/basic` and the
  authenticated proxy listener.
- Contract: `docs/specs/challenge-authentication.md` and the existing
  `docs/specs/auth-policy.md`.

## Supersedes

- None

## Superseded by

- None
