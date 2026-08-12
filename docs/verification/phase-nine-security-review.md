# Phase 9 security review record

## Review scope

The current release gate covers release-manifest identity, component allowlist
and exclusion boundaries, VBE/atomic-publication evidence, and SHA-256 pairing.
The authoritative implementation is
`tools/Validate-ReleaseSecurity.ps1`; its tamper matrix is run by
`task test:release-security` and the gate runs before
`tools/Validate-ReleaseArtifact.ps1` opens Excel.

## Threat and control matrix

| Asset / boundary | Control in this review | Result |
| --- | --- | --- |
| Release workbook and manifest bytes | Adjacent SHA-256 sidecar; stale/tampered pair fails closed | covered |
| Production component set | Exact policy allowlist/denylist and actual manifest count | covered |
| Source-to-artifact boundary | Canonical base/output and repository-relative source path checks; Tests/Benchmarks/Xlflow/Dev denied for included entries | covered |
| VBE compile/save/close and publication | Required manifest validation fields and `atomic_replace` | covered |
| Caller credentials and redirect headers | Auth provider, sensitive-header redaction, and redirect suppression specs/tests | covered by existing auth/redirect gates |
| TLS trust and certificate rejection | OS validation is retained; no deterministic certificate-negative fixture exists yet | deferred, release-blocking for claiming certificate coverage |
| HTTP/2/HTTP/3 TLS and Office bitness | Compatibility evidence is still environment-dependent (current proof is x64) | deferred |
| Integrated/proxy challenge auth and cookie persistence | Explicitly outside the preemptive auth/no-cookie contracts | deferred |

## Verification commands

```powershell
task check
task test:release-security
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
This record does not declare Phase 9's certificate-negative test or the broader
protocol/auth/cookie compatibility matrix complete; those remain explicit
follow-up gates rather than being hidden by a green manifest check.
