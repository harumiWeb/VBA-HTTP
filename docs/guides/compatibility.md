# Compatibility and support boundary

## Supported runtime

The supported runtime is Windows x64 Office with WinHTTP. The current evidence
path includes x64 compile, loopback integration, release build/inspection, and
consumer smoke. The default buffered COM backend and the native transport share
the public API/error contract on that boundary.

## Explicitly unsupported or host-dependent

| Area | Policy |
| --- | --- |
| 32-bit Office | Unsupported until a complete real-host evidence bundle and a superseding ADR exist. The legacy declaration branch is a static ABI guard only. |
| HTTP/3/QUIC | Unsupported. The native flag is retained for diagnostics/future work and is not a release or compatibility guarantee. |
| HTTP/2 | Opt-in native capability; required mode needs host-specific HTTPS evidence. Plain loopback HTTP/1.1 is not proof of HTTP/2. |
| XLAM | Independent optional target; not the primary source-package distribution. |
| PAC/WPAD/SOCKS | Host-dependent and outside the deterministic manual proxy contract. |
| OAuth/interactive authentication | Outside the current provider API. |
| Trusted corporate CONNECT/domain credentials | Host-dependent; loopback boundaries do not promote support. |

When a capability is unavailable, the library should fail with a stable
validation/protocol category or use the documented fallback. It must not claim
support based only on conditional VBA declarations or a requested option flag.

## Capability summary

- COM supports buffered requests, batches, default/direct/manual proxy modes,
  preemptive Basic/Bearer, and WinHTTP-managed automatic challenge exchange.
- Native supports the same buffered contract plus constant-memory streaming,
  explicit HTTP/2 policy, gzip/deflate decoding, and bounded native challenge
  controls.
- Both backends reject unsafe redirect/header combinations and preserve the
  same error categories.

## Evidence and reporting

Host evidence should record Windows/WinHTTP/Office versions, x64 architecture,
source/workbook revision, requested capability, observed error category, and
`ProtocolUsed`. Do not include credentials, query values, raw target paths, or
response bodies in evidence.

See the normative matrix in
[`../specs/compatibility-matrix.md`](../specs/compatibility-matrix.md), the x64
boundary ADR [`../adr/ADR-0030-32-bit-office-support-boundary.md`](../adr/ADR-0030-32-bit-office-support-boundary.md),
and the HTTP/3 policy ADR
[`../adr/ADR-0035-http3-support-boundary.md`](../adr/ADR-0035-http3-support-boundary.md).
