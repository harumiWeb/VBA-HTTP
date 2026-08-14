# ADR-0011: Deterministic release artifact checksum evidence

## Status

`accepted`

## Background

`xlflow build` already compiles the filtered workbook and publishes the release
artifact and build manifest atomically, but the manifest does not identify the
bytes that were actually published. A release consumer or archive process needs
to detect a stale workbook/manifest pair without opening Excel. A checksum file
must also be safe to regenerate: a failed evidence write must not truncate a
previously valid sidecar or touch the tracked development workbook.

## Decision

- Every release artifact has a companion `<artifact>.checksum.json` sidecar
  containing SHA-256 values for the workbook and its adjacent xlflow build
  manifest. The sidecar uses a versioned, path-free schema with no timestamps so
  unchanged inputs produce byte-identical evidence.
- `tools/Write-ReleaseChecksums.ps1` computes both hashes before writing and
  publishes the sidecar through a same-directory temporary file followed by
  `System.IO.File.Replace` (or an atomic move for a new sidecar).
- `tools/Verify-ReleaseChecksums.ps1` recomputes both hashes and rejects missing,
  malformed, stale, or cross-artifact sidecars. The external release smoke
  validator invokes it before inspecting VBA components or executing consumer
  macros.
- `release:build` runs checksum generation only after `xlflow build` succeeds;
  the tracked development workbook remains outside the checksum publication
  path. Generated release evidence is retained together by release archival
  tooling and remains ignored in ordinary Git worktrees.

## Consequences

- Release archives can be audited and verified without Excel, and a changed
  workbook or manifest fails closed before consumer smoke.
- The GitHub tag-release bundle uses the same SHA-256 discipline for the source
  ZIP, production-only pack XLSM, pack/release manifests, license, and notices.
  Its `SHA256SUMS` file also covers the release manifest and is validated before
  `gh release create`; this does not imply VBE validation for the pack.
- The checksum sidecar is supplementary evidence, not a replacement for
  xlflow's VBE compile, manifest validation, or atomic workbook publication.
  Because the two files are published by separate commands, a failed sidecar
  generation can leave a newly built artifact paired with an older sidecar;
  the release gate intentionally refuses to treat that pair as publishable.
- SHA-256 is stable and widely available in Windows PowerShell/.NET, but the
  sidecar format is versioned so a future algorithm can be introduced without
  silently changing validation semantics.

## Rationale

- Code: `tools/Write-ReleaseChecksums.ps1`,
  `tools/Verify-ReleaseChecksums.ps1`, `tools/Validate-ReleaseArtifact.ps1`,
  and `Taskfile.yml`.
- Tests: `tools/Test-ReleaseChecksums.ps1` and the `test:release-checksum`
  Taskfile task cover deterministic regeneration, tamper detection, missing
  input failure, and preservation of the previous sidecar after a failed write.
- Current contract: `docs/specs/release-checksum.md` and
  `docs/specs/development-and-release-workflow.md`.
- Existing publication boundary: ADR-0001 and the `publication.method` field in
  the xlflow build manifest.

## Supersedes

- None

## Superseded by

- None
