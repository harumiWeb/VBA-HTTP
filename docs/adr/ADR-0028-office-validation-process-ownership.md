# ADR-0028: Office validation process ownership

## Status

`accepted`

## Background

The bitness runner creates an Excel COM instance to read Office metadata and
to execute the validation loop. A machine can simultaneously contain a
user-owned Excel session for unrelated work. Calling `Quit` on an ambiguous
COM reference or using a name-only process cleanup could close that session,
which is unacceptable for a validation tool.

## Decision

- Snapshot all existing `EXCEL.EXE` PIDs before creating the metadata COM
  instance.
- Poll briefly until a new Excel PID appears and treat only the set difference
  as runner-owned. If no new PID can be proven, fail closed before calling
  `Quit`; do not infer ownership from the process name alone.
- Release artifact validation strengthens this proof by mapping the newly
  created `Excel.Application.Hwnd` to its owning PID with
  `GetWindowThreadProcessId`, then requiring that PID to be absent from the
  baseline and to remain an `EXCEL` process. A concurrent user launch cannot
  be selected merely because it happened to appear in a process-name
  difference.
- Release and quit only the owned COM instance. Other Excel PIDs remain
  untouched. The protocol-host runner and stress runners retain their own
  exact-PID ownership/watchdog rules.
- After `Quit` and COM release, cleanup waits five seconds for the owned PID to
  exit normally. Only then may the validator force-stop that exact proven PID;
  no name-based Excel termination is permitted.
- Keep the bitness evidence command host-specific: a concurrent user Excel
  session may make the run inconclusive, but it must never be a cleanup target.

## Consequences

- Office metadata validation is slightly more conservative and may require a
  retry when Excel startup is blocked or the COM server is reused by the host.
- A passing result proves that the metadata instance was newly created, while
  unrelated workbooks remain outside the runner's cleanup scope.
- This does not claim 32-bit evidence on an x64 host; the architecture gate
  remains separate and fail-closed.

## Rationale and evidence

- Implementation: `tools/Run-OfficeBitnessValidation.ps1` and
  `tools/Validate-ReleaseArtifact.ps1`.
- Current contract: `docs/specs/office-bitness-validation.md`.
- Related ownership controls: `ADR-0015-resource-stress-evidence-policy.md`,
  `ADR-0025-protocol-host-evidence-boundary.md`, and `tasks/lessons.md`.

## Supersedes

- None

## Superseded by

- None
