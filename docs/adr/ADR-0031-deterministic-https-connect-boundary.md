# ADR-0031: Deterministic HTTPS CONNECT proxy boundary

## Status

`accepted`

## Background

The HTTP forwarding proxy and its fixed Basic challenge already prove named
proxy routing and 407 handling, but they previously rejected CONNECT. A
successful HTTPS request through a corporate proxy cannot be made a
deterministic offline test: certificate trust, proxy policy, and Windows
domain credentials belong to the host environment. The project nevertheless
needs to distinguish “the proxy never established a tunnel” from “the tunnel
was established and normal TLS validation rejected the intentionally untrusted
fixture.”

The test server already owns both a loopback HTTP listener and an intentionally
untrusted TLS listener. It can therefore provide a safe CONNECT target without
allowing arbitrary forwarding or changing the machine trust store.

## Decision

- Add opt-in `-proxy-tls-listen` and `-proxy-tls-auth-listen` listeners to the
  local test server. They accept CONNECT only for the server's own TLS alias
  and dial only the server-owned loopback TLS address.
- The authenticated listener returns a fixed 407 Basic challenge until the
  fixed loopback credential is supplied. CONNECT attempt and authorized counts
  are exposed only through the local `/__admin/proxy-stats` endpoint.
- COM and native integration tests use the manual proxy contract and require
  the request to reach the TLS boundary. They expect the normal `HttpErrorTls`
  result from the untrusted certificate and assert proxy stats increased; they
  do not install, ignore, or trust the certificate.
- The fixture rejects non-target CONNECT authorities and never connects to an
  external address. It remains a test-only component and is excluded from
  release workbooks.
- A passing CONNECT-boundary test does not promote trusted HTTPS proxy
  interoperability, PAC/WPAD, SOCKS, HTTP/3, or Windows-domain
  Negotiate/NTLM. Those require separate host-specific evidence.

## Consequences

- Proxy routing, proxy challenge, CONNECT establishment, and TLS validation
  are observable as separate boundaries on the supported x64 Office host.
- The integration suite remains deterministic and offline, while the TLS
  certificate rejection contract is preserved.
- A trusted CONNECT success test is intentionally deferred; adding a trust
  anchor or certificate-ignore switch would weaken the security contract.

## Rationale / Evidence

- Fixture: `tools/testserver/proxy.go`, `tools/testserver/main.go`,
  `tools/testserver/server.go`, and `tools/testserver/server_test.go`.
- Integration: `WinHttpComTransportTests` and
  `WinHttpNativeTransportTests` HTTPS CONNECT boundary cases.
- Contract: `docs/specs/https-connect-proxy.md`,
  `docs/specs/proxy-policy.md`, `docs/specs/challenge-authentication.md`,
  and `docs/specs/local-test-server.md`.

## Supersedes

- None

## Superseded by

- None
