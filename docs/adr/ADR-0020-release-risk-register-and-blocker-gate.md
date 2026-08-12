# ADR-0020: Release risk register and blocker gate

## Status

`accepted`

## Background

The Phase 9 security review has completed the controls that can be proven on
the current x64 host, while protocol/bitness compatibility and integrated
challenge authentication require different Windows capabilities or a separate
replay contract. A prose-only list of residual risks can silently drift and
cannot prove that the current release has no known blocker.

## Decision

- Track residual security and compatibility issues in the versioned
  `docs/security/risk-register.json` file.
- Require an explicit `current_release_blockers` array. It must be empty for a
  release build; future-v1 blockers are allowed only when `blocks_current_release`
  is false and the required evidence is recorded.
- Validate schema, unique issue IDs, allowed status/severity values, and blocker
  consistency with `tools/Validate-SecurityRiskRegister.ps1`.
- Run the validator from `task verify` and before release artifact inspection.
  The gate is Excel-free and fails closed on malformed or contradictory data.
- Keep the register separate from implementation TODOs: it records risk and
  evidence obligations, while `tasks/todo.md` records work items.

## Consequences

- The current x64 release can prove zero known current blockers without falsely
  claiming HTTP/2/3, 32-bit, or integrated-auth compatibility.
- A future release cannot hide a newly discovered blocker in prose; the
  machine-readable register and release gate must be updated together.
- Deferred issues remain visible and block the v1.0 checklist until their
  dedicated evidence exists.

## Rationale and evidence

- Register: `docs/security/risk-register.json`.
- Validator and tamper tests: `tools/Validate-SecurityRiskRegister.ps1`,
  `tools/Test-SecurityRiskRegister.ps1`.
- Release integration: `tools/Validate-ReleaseSecurity.ps1` and `Taskfile.yml`.
- Review record: `docs/verification/phase-nine-security-review.md`.

## Supersedes

- None

## Superseded by

- None
