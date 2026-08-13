# ADR-0030: 32-bit Office support boundary

## Status

`accepted`

## Background

The repository contains conditional VBA7/legacy WinHTTP declarations, but the
available host evidence is from 64-bit Office only.  Conditional compilation
and an Excel-free declaration scan prove source shape; they do not prove that
32-bit Office can compile, load, execute, and publish the workbook safely.
Claiming 32-bit compatibility without that evidence would make the release
matrix stronger than the tested product.

## Decision

- The current supported runtime target is Windows x64 Office.  A release is
  not required to support or validate 32-bit Office; 32-bit Office is
  `unsupported-by-policy` until a later decision changes this boundary.
- `task test:office-bitness` and the protocol-host evidence runner are x64
  promotion gates.  They fail before opening a new Excel instance when the
  selected bridge is X86.
- The VBA7/legacy declaration contract remains an Excel-free regression gate.
  It protects pointer-size source branches but is not 32-bit runtime evidence
  and must not promote an x86 compatibility row.
- An operator may explicitly run the office runner in diagnostic-only mode on
  a future x86 host. Such output uses `status=diagnostic` and is labeled
  `support_status=unsupported-by-policy`; it is not a release or v1 promotion
  artifact and must not be used as support evidence.
- The existing process-ownership boundary remains mandatory: validation
  snapshots Excel PIDs, proves a newly created runner-owned PID, and only
  closes or terminates that owned PID.  A pre-existing user Excel process is
  never touched; ambiguous ownership fails closed.
- Reconsidering x86 support requires a superseding ADR, a real 32-bit Office
  host run covering compile, focused tests, loopback integration, release build,
  and consumer smoke, plus an updated compatibility matrix and release
  checklist.

## Consequences

- The compatibility matrix and release checklist can state a precise supported
  boundary instead of carrying an indefinitely pending x86 gate.
- Consumers running 32-bit Office should treat the workbook as unsupported;
  no compatibility or support expectation is implied by the legacy ABI branch.
- HTTP/3/QUIC evidence remains a separate x64 host capability question and is
  not resolved by this decision.
- Future x86 investigation remains possible without weakening the release
  gate, but its evidence must be clearly diagnostic and non-promotional.

## Rationale and evidence

- Current supported evidence: `benchmarks/results/office-bitness-x64.json`.
- Runner and ownership contract: `tools/Run-OfficeBitnessValidation.ps1`,
  `tools/Run-ProtocolHostValidation.ps1`, and ADR-0028.
- Current validation contract: `docs/specs/office-bitness-validation.md` and
  `docs/specs/compatibility-matrix.md`.
- Static declaration guard: `tools/Test-NativeDeclarationContract.ps1`.

## Supersedes

- ADR-0025. Its protocol evidence, trusted-endpoint, and ownership rules remain
  in force, while this ADR replaces its x86 promotion requirement with the
  x64-only support boundary.

## Superseded by

- None
