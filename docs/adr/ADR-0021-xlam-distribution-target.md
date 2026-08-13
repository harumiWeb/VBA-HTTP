# ADR-0021: Independent XLAM distribution target

## Status

`accepted`

## Context

The XLSM release is the primary consumer artifact, but an Excel add-in has a
different workbook identity and must not be produced by changing the extension
of an XLSM after build. The previous roadmap entry was blocked because no
same-extension base workbook or independent validation path existed.

## Decision

- Track `build/VBA-HTTP.xlam` as a verified development base for the add-in
  target, generated from the synchronized development workbook by Excel
  `xlOpenXMLAddIn` SaveAs.
- Build `build/Release/VBA-HTTP.xlam` with an explicit
  `xlflow build --base ... --out ...` invocation. The XLSM base and artifact are
  never modified by the XLAM build.
- Validate the XLAM dry-run against the same production component policy, then
  require VBE compile/save/close, atomic publication, checksum, `Workbook.IsAddin`
  identity, component inspection, risk-register validation, and an external
  component and identity validation. The shared external consumer harness is
  the primary runtime smoke path; an XLAM-specific consumer smoke can be added
  separately. The artifact itself has no scaffold entry macro.
- Keep XLAM generation as a separate Taskfile target; it is not an implicit
  second output of `task release:build`.

## Consequences

- Consumers can receive a real add-in artifact without manually renaming or
  removing modules from the XLSM release.
- The repository carries a second binary base and therefore must verify both
  artifact paths and checksums independently.
- Full network smoke remains shared with the XLSM target; the XLAM target adds
  identity and component-boundary evidence rather than duplicating every
  endpoint test.

## Evidence

- `build/VBA-HTTP.xlam`
- `tools/Validate-XlamBuildPlan.ps1`
- `tools/Validate-XlamArtifact.ps1`
- `tools/Test-XlamDistribution.ps1`
- `docs/specs/distribution.md`

## Supersedes

- None

## Superseded by

- None
