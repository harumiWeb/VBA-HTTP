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
| HTTP/2/HTTP/3 TLS and Office bitness | Compatibility evidence is still environment-dependent (current proof is x64) | deferred |
| Integrated/proxy challenge auth | Explicitly outside the preemptive Basic/Bearer contract and challenge replay | deferred |
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
The risk register contains zero current release blockers; its three deferred
items remain future-v1 evidence obligations and are not silently treated as
passed compatibility.
This record does not declare the broader protocol/auth compatibility matrix
complete; HTTP/2/HTTP/3 negotiation evidence, 32-bit Office evidence, and
integrated/proxy challenge authentication remain explicit follow-up gates. The
certificate-negative test is covered by the deterministic HTTPS fixture and
both transport integration tests. The explicit cookie jar policy is covered by
ADR-0018 and its unit/integration/release smoke tests.

Detailed fixture and command evidence: `docs/verification/phase-nine-certificate-validation.md`.
