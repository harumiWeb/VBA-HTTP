# ADR-0024: Deterministic proxy challenge fixture boundary

## Status

`accepted`

## Context

`HttpWindowsAuthProvider` already supports a proxy target, but the original
loopback proxy intentionally forwarded unauthenticated HTTP requests. Reusing
that listener for 407 tests would make every existing proxy test depend on
credential setup and would not distinguish routing failures from authentication
failures. HTTPS CONNECT and real corporate proxy behavior are still
host-dependent.

## Decision

- Add a separate optional `-proxy-auth-listen` loopback listener to the Go test
  server. Its readiness JSON exposes `proxy_auth_url`; the existing
  `proxy_url` behavior is unchanged.
- The authenticated fixture forwards only to the same loopback alias as the
  ordinary proxy, rejects CONNECT and non-target hosts, and returns a fixed
  Basic 407 challenge until `proxy-user` / `proxy-pass` is supplied.
- Integration and release consumer smoke use `HttpAuthTargetProxy` with an
  explicit insecure loopback opt-in. No credential is embedded in a proxy URL,
  logged, echoed, or used for an external target.
- This fixture proves bounded HTTP forwarding challenge behavior only. It does
  not promote HTTPS CONNECT, PAC/WPAD, or Windows-domain Negotiate/NTLM to the
  compatibility matrix.

## Consequences

- COM and native proxy challenge paths can be tested deterministically on the
  current x64 host, including wrong-credential final 407 behavior.
- Existing routing tests remain focused on routing and do not inherit an auth
  dependency.
- A real CONNECT tunnel and domain-auth host remain required before claiming
  broad proxy/integrated-auth compatibility.

## Evidence

- Fixture: `tools/testserver/proxy.go`, `tools/testserver/main.go`, and
  `tools/testserver/server_test.go`.
- Integration: `WinHttpComTransportTests` and
  `WinHttpNativeTransportTests` proxy challenge cases.
- Release: `tools/consumer/ReleaseBatchSmoke.bas` and
  `tools/Validate-ReleaseArtifact.ps1`.
- Contract: `docs/specs/proxy-policy.md`,
  `docs/specs/challenge-authentication.md`, and
  `docs/specs/local-test-server.md`.

## Supersedes

- None

## Superseded by

- None
