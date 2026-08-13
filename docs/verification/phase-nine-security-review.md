# Phase 9 security review record

## Review scope

The current release gate covers release-manifest identity, component allowlist
and exclusion boundaries, VBE/atomic-publication evidence, and SHA-256 pairing.
The authoritative implementation is
`tools/Validate-ReleaseSecurity.ps1`; its tamper matrix is run by
`task test:release-security` and the gate runs before
`tools/Validate-ReleaseArtifact.ps1` opens Excel.
The current blocker assertion is separately recorded in
`docs/security/risk-register.json` and validated by
`task test:security-risks` before release inspection.

## Threat and control matrix

| Asset / boundary | Control in this review | Result |
| --- | --- | --- |
| Release workbook and manifest bytes | Adjacent SHA-256 sidecar; stale/tampered pair fails closed | covered |
| Production component set | Exact policy allowlist/denylist and actual manifest count | covered |
| Source-to-artifact boundary | Canonical base/output and repository-relative source path checks; Tests/Benchmarks/Xlflow/Dev denied for included entries | covered |
| VBE compile/save/close and publication | Required manifest validation fields and `atomic_replace` | covered |
| Caller credentials and redirect headers | Auth provider, sensitive-header redaction, and redirect suppression specs/tests | covered by existing auth/redirect gates |
| TLS trust and certificate rejection | Self-signed loopback fixture is rejected by both COM/native transports and maps to `HttpErrorTls`; no ignore-certificate option exists | covered |
| Repeated COM receive-timeout cleanup | ADR-0022 `Abort`/bounded drain plus canonical 25-iteration loopback stress with idle handle budget | mitigated on current x64 host |
| HTTP/2/HTTP/3 TLS | x64 HTTP/2 is recorded at benchmarks/results/protocol-host-http2.json; HTTP/3/QUIC remains environment-dependent | partial / deferred |
| Buffered challenge auth | Bounded server Basic and proxy Basic challenges are covered by COM/native tests, release smoke, redaction, and streaming pre-network rejection | covered for loopback; host-specific domain/CONNECT evidence deferred |
| Current release blockers | Versioned risk register and fail-closed validator | covered (0) |

## Verification commands

```powershell
task check
task test:integration
task test:release-security
task test:security-risks
task build:plan
task release:security
```

`task release:security` writes a path-stable report under
`.xlflow/release-security/release-security.json`. A distributable evidence
bundle must retain that report together with the exact release workbook,
`*.build.json`, and `*.checksum.json`; the report records their SHA-256 values
without embedding host paths or secrets.

## Review outcome

Manifest/component integrity is accepted for the current x64 release path.
The risk register contains zero current release blockers; two deferred items
remain future-v1 evidence obligations and the repeated COM timeout item is
mitigated on the current x64 host. The x64 HTTP/2 host record is evidence for
that host only; it is not silently treated as HTTP/3 or cross-bitness
compatibility evidence.
This record does not declare the broader protocol/auth compatibility matrix
complete; HTTP/3 negotiation evidence and host-specific integrated/proxy
challenge authentication remain explicit follow-up gates. 32-bit Office is
outside the supported boundary under ADR-0030. The
certificate-negative test is covered by the deterministic HTTPS fixture and
both transport integration tests. The explicit cookie jar policy is covered by
ADR-0018 and its unit/integration/release smoke tests.

Detailed fixture and command evidence: `docs/verification/phase-nine-certificate-validation.md`.
The latest HTTP/3 host attempts and their fail-closed cleanup evidence are
recorded in
`docs/verification/protocol-host-http3-attempt-2026-08-13.md`.
