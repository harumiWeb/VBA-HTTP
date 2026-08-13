# ADR-0035: HTTP/3 and QUIC support boundary

## Status

`accepted`

## Context

The native transport contains an opt-in WinHTTP protocol-mask path for HTTP/2
and HTTP/3, and `HttpResponse.ProtocolUsed` can report the flag returned by
WinHTTP. The implementation path is useful for capability diagnosis, but the
repository has no passing HTTP/3 negotiated-host record. The current x64 host
rejects the HTTP/3 option during the Excel-free preflight with
`ERROR_NOT_SUPPORTED (50)`, and the deterministic loopback TLS fixture is
intentionally untrusted. Neither requested flags nor a local certificate
rejection prove HTTP/3 interoperability.

Claiming HTTP/3 support would therefore expand the public compatibility promise
beyond the verified Windows/WinHTTP/TLS matrix and would make v1 promotion
depend on an environment that is not currently available.

## Decision

- HTTP/3 and QUIC are `unsupported-by-policy` for the current distribution.
  Consumers must not rely on `AllowHttp3`, `ProtocolUsed = "HTTP/3"`, or HTTP/3
  fallback behavior as a supported product capability.
- HTTP/1.1 remains the default supported protocol, and the x64 HTTP/2 host
  record remains a host-specific supported evidence row.
- The existing native HTTP/3 option and protocol-reporting code remain only as
  diagnostic/future-compatibility code. They are not a release, v1 promotion,
  compatibility, or security gate.
- The protocol-host promotion runner accepts HTTP/2 only. HTTP/3 capability
  checks may use the Excel-free probe for diagnosis, but they must not create a
  passing promotion artifact or start an Excel consumer proof.
- The offline loopback server remains HTTP/1.1. An optional HTTP/3 fixture such
  as a `quic-go` UDP listener may be added later for server-side and Go-client
  regression tests, but it cannot change the VBA/WinHTTP support boundary.
- Reconsidering this boundary requires a superseding ADR, a trusted QUIC
  endpoint, exact `ProtocolUsed = "HTTP/3"` evidence on the supported x64
  Office/WinHTTP host, focused integration, release build, and external
  consumer smoke evidence.

## Consequences

- The v1 release is no longer blocked by unavailable HTTP/3/QUIC host evidence.
- The public documentation has a clear distinction between diagnostic code and
  supported behavior; unsupported HTTP/3 requests remain caller responsibility.
- The HTTP/3 option path and historical failed attempts remain useful when
  diagnosing a future Windows/WinHTTP upgrade, without being mistaken for a
  compatibility claim.
- A future `quic-go` test fixture has infrastructure and maintenance cost but
  remains isolated under `tools/` and never becomes a workbook dependency.

## Evidence

- Current capability failure: `docs/verification/protocol-host-http3-attempt-2026-08-13.md`.
- Supported protocol record: `benchmarks/results/protocol-host-http2.json`.
- Native option implementation: `src/classes/WinHttpNativeTransport.cls`,
  `src/classes/HttpProtocolOptions.cls`, and `src/modules/WinHttpNativeApi.bas`.
- Promotion boundary: `tools/Run-ProtocolHostValidation.ps1`,
  `tools/Validate-ProtocolHostEvidence.ps1`, and
  `docs/specs/protocol-host-validation.md`.
- Existing protocol design: `docs/adr/ADR-0009-native-protocol-negotiation-policy.md`
  and `docs/specs/protocol-policy.md`.

## Supersedes

- None

## Superseded by

- None
