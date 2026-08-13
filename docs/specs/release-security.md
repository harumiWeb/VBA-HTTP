# Release manifest security specification

## Scope

`tools/Validate-ReleaseSecurity.ps1` validates the release artifact and its
adjacent xlflow manifest before Excel is opened. It is a manifest-integrity and
source-boundary gate; it does not replace `tools/Validate-ReleaseArtifact.ps1`,
which inspects the actual `VBProject` and runs the external consumer smoke.

## Required identity and publication evidence

- artifact: exactly `build/Release/VBA-HTTP.xlsm`;
- manifest `base`: exactly `build/VBA-HTTP.xlsm`;
- manifest `output`: the canonical artifact path;
- `schema_version=1`, `command=build`, and `backend=excel`;
- `validation.source_applied=true`, `vbe_compile=passed`,
  `workbook_saved=true`, `workbook_closed=true`, and `excel_cleanup=clean`;
- `validation.components_applied` equals the included component count; and
- `publication.method` is either `atomic_create` (first publication) or
  `atomic_replace` (replacement of an existing artifact); both are atomic.

Any missing or contradictory field fails closed.

## Component and path boundary

The included and excluded component name sets must exactly match
`tools/build-component-policy.json`. Every component has a safe VBA identifier,
an allowed type (`class`, `standard`, or `document`), and an extension matching
its type (`.cls` for classes, `.bas` otherwise).

Included and related paths are repository-relative and must remain under
`src/classes/`, `src/modules/`, or `src/workbook/`. An included path containing
the segments `Tests`, `Benchmarks`, `Xlflow`, or `Dev` is rejected. Excluded
paths must contain one of those development-only segments; this prevents a
production source path from being hidden in the exclusion list. Duplicate names
or source paths, absolute paths, traversal segments, and related paths outside
the approved source tree are rejected.

## Checksum and report

The validator invokes `tools/Verify-ReleaseChecksums.ps1` for the adjacent
`<artifact>.checksum.json` sidecar. When `-ReportPath` is supplied, it writes a
UTF-8, path-stable JSON report atomically. The report contains SHA-256 values for
the artifact, manifest, and component policy, counts, a fixed ordered list of
checks, and a fixed list of deferred security topics. It contains no host paths,
credentials, request headers, or timestamps.

The release task writes the report under `.xlflow/release-security/`; it is
evidence for the current artifact and is not a substitute for archiving the
artifact, manifest, and checksum together.

## Deferred security topics

The gate does not mark the following as implemented: HTTP/2/HTTP/3 TLS/Office
compatibility evidence or integrated/proxy challenge authentication. These
remain release-review items and are listed in the report and
`docs/verification/phase-nine-security-review.md`.

## Verification

- `task test:release-security` builds a source-derived fixture and checks one
  valid manifest plus seven tampering cases (base path, included test path,
  missing exclusion, failed compile evidence, non-atomic publication,
  related-path escape, and artifact checksum tamper).
- `task release:security` validates the current release artifact and writes the
  deterministic report before `task release:smoke` opens Excel.
