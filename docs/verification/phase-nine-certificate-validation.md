# Phase 9 certificate validation evidence

## Fixture contract

`tools/testserver/` can be started with `-tls-listen 127.0.0.1:0`. It publishes
an `https_url` readiness field and serves the same deterministic routes over a
process-local self-signed certificate. The certificate has loopback SANs but is
not installed in the Windows trust store; the fixture therefore exercises the
normal OS trust decision without introducing a certificate-ignore switch.

## Proof

- `task testserver:unit`: passed, including certificate constraints and SAN
  checks.
- `task check`: lint, analyze, format, and class-source validation passed.
- `task test:integration`: 71/71 integration tests passed on the configured
  Windows x64 Office host.
- `WinHttpComTransportTests.Test_ComTransport_RejectsUntrustedCertificate`:
  self-signed HTTPS request mapped to `HttpErrorTls`.
- `WinHttpNativeTransportTests.Test_NativeTransport_RejectsUntrustedCertificate`:
  self-signed HTTPS request mapped to `HttpErrorTls`.
- No test imports the certificate into the OS store, sets a WinHTTP security
  bypass flag, or logs certificate/private-key material.

HTTP/2/HTTP/3 negotiated-protocol evidence and 32-bit Office execution remain
separate compatibility-matrix gates.
