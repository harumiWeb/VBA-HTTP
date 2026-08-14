# ADR-0039: 32-bit Office unverified and community validation boundary

## Status

`accepted`

ADR-0030 remains as the historical record of the earlier conservative boundary
and is superseded by this decision for the classification of 32-bit Office.

## Background

`WinHttpNativeApi.bas` retains VBA7 pointer-sized declarations and legacy
32-bit `Long` branches. The Excel-free declaration contract confirms that both
source branches remain syntactically represented, but it cannot prove that a
real 32-bit Office host can compile, load, execute, publish, and clean up the
complete workbook. Current runtime evidence is x64 only.

The source therefore may work on some 32-bit Office hosts, but the repository
must not imply compatibility from conditional compilation alone. At the same
time, classifying an untested runtime as known-incompatible prevents
contributors from supplying useful evidence.

## Decision

- Windows x64 Office remains the only officially supported and release-verified
  runtime target.
- Windows 32-bit Office is classified as `unverified`: it is not a supported
  release target or compatibility guarantee, but it is not declared known
  incompatible by policy.
- The normal `task test:office-bitness` and release promotion path remain x64
  gates. They do not silently switch to an x86 bridge.
- `Run-OfficeBitnessValidation.ps1 -ExpectedArchitecture X86 -DiagnosticOnly`
  is the explicit contributor/evidence path. A complete x86 run records
  `architecture=X86`, `status=diagnostic`, and `support_status=unverified`.
- An x86 diagnostic result is never treated as a release pass. It may be
  attached to a contribution or compatibility discussion without changing the
  GitHub Release target, which remains `windows-x64-office`.
- A maintainer may promote the public label to `community-validated` only after
  reviewing a reproducible real-host bundle. Official support requires a later
  ADR decision and a complete compile, focused/full test, loopback integration,
  filtered release build, and external consumer-smoke record on representative
  32-bit Office hosts.
- Release manifests record `office_32_bit=unverified` so the absence of x86
  evidence is explicit and distinguishable from HTTP/3/QUIC, which remains
  `unsupported-by-policy` for a different reason.
- The process-ownership boundary is unchanged: validation snapshots existing
  Excel PIDs, proves a runner-owned PID, and never closes or terminates an
  unrelated user Excel process.

## Consequences

- Users running 32-bit Office receive an honest compatibility statement:
  the code may work, but the project offers no official support or release
  guarantee yet.
- Contributors have a documented, non-promotional way to run x86 validation
  and submit machine-readable evidence.
- x64 releases are not blocked by the absence of x86 hardware or Office, and
  the hosted GitHub workflow does not start Excel or claim x86 validation.
- Documentation, schemas, validators, risk records, and release manifests must
  use `unverified` consistently. Historical ADRs may retain their original
  wording when describing the decision they recorded, but current specs must
  point to this ADR.

## Rationale and evidence

- Conditional declaration and upload sentinel guard:
  `tools/Test-NativeDeclarationContract.ps1` and
  `src/modules/WinHttpNativeApi.bas`.
- Existing x64 runtime evidence:
  `benchmarks/results/office-bitness-x64.json`.
- Contributor runner and ownership guard:
  `tools/Run-OfficeBitnessValidation.ps1`,
  `tools/Validate-OfficeBitnessResult.ps1`, and ADR-0028.
- Current compatibility contract:
  `docs/specs/office-bitness-validation.md` and
  `docs/specs/compatibility-matrix.md`.

## Supersedes

- ADR-0030

## Superseded by

- None
