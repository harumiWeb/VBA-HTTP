# ADR-0017: Release manifest security boundary

## Status

`accepted`

## Background

The xlflow build manifest records the component set and VBE validation result,
but the ordinary release smoke check historically learned most of its trust
from opening Excel. A malformed or stale manifest could therefore pass a
partial preflight even when its base/output paths, source paths, or exclusion
claims were not safe to audit. The checksum sidecar proves bytes, not that the
manifest describes a production-only workbook.

## Decision

- Add `tools/Validate-ReleaseSecurity.ps1` as a fail-closed, Excel-free gate
  before release component inspection and consumer smoke.
- Accept only the canonical `build/Release/VBA-HTTP.xlsm` output whose manifest
names `build/VBA-HTTP.xlsm` as its base and reports an atomic publication
(`atomic_create` for first publication or `atomic_replace` for replacement), source
  application, VBE compile, save, close, and clean Excel cleanup.
- Compare included and excluded component names with
  `tools/build-component-policy.json`. Included source paths must stay under
  `src/classes/`, `src/modules/`, or `src/workbook/` and must not contain
  `Tests`, `Benchmarks`, `Xlflow`, or `Dev` path segments. Excluded entries must
  be under one of those development-only segments. Related paths are subject to
  the same repository-relative boundary and cannot escape `src/`.
- Verify the adjacent SHA-256 sidecar before the security report is published.
  An optional deterministic JSON report records artifact, manifest, policy
  hashes, checks performed, and deferred security topics without host paths or
  secrets. Reports are written atomically under `.xlflow` by the release task.
- Keep the gate lexical and manifest-focused. It does not claim HTTP/2/HTTP/3
  TLS compatibility or integrated/proxy challenge authentication; those remain
  explicit residual risks until their dedicated fixtures and policies exist.
  Certificate rejection is covered by the loopback fixture in the integration
  suite, and cookie handling is covered by ADR-0018.

## Consequences

- A release cannot reach Excel inspection or external consumer smoke when a
  manifest points outside the approved source boundary, omits an excluded
  component, loses VBE/atomic-publication evidence, or has a stale checksum.
- The validator can run in a clean environment without Excel, while the later
  artifact validator still inspects the actual `VBProject` and executes the
  consumer smoke harness.
- The component policy remains a deliberate allowlist/denylist change point;
  adding a production component requires updating the policy and its tests.
- The gate is not a full threat model or protocol-compatibility test. Deferred
  topics are reported instead of being silently treated as passed security
  coverage.

## Rationale

- Code: `tools/Validate-ReleaseSecurity.ps1`,
  `tools/Test-ReleaseSecurity.ps1`, and `tools/Validate-ReleaseArtifact.ps1`.
- Specification: `docs/specs/release-security.md` and
  `docs/specs/release-checksum.md`.
- Policy: `tools/build-component-policy.json`.
- Verification: `task test:release-security` and the release `security-report`
  generated under `.xlflow/release-security/`.

## Supersedes

- None

## Superseded by

- None
