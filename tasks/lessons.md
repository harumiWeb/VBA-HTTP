## Proof-loop safety

- Do not chain `xlflow push` and `xlflow save` in one shell command. If compile
  fails after import, stop the session without saving, inspect the diagnostic,
  fix source, then rerun the proof loop.

- After recovery or an external source timestamp change, do not rely on
  `xlflow push --fast`'s changed-state shortcut as proof of workbook sync. If
  `xlflow status --json` still reports `src_newer_than_workbook`, run an
  explicit session `xlflow push --session --no-save --json`, then a separate
  `xlflow save --session --json`, and confirm the warning is gone before
  committing the development workbook.

- If a command wrapper times out while xlflow still reports workbook
  coordination as busy, do not retry or save. Inspect the recorded operation
  and PID, discard only a managed poisoned session through the recovery
  workflow, and reorient from a clean `xlflow status --json` before starting a
  new proof loop.

## Excel process ownership

- External protocol-host validation uses a hidden watchdog with the exact
  runner-owned Excel PID list. It must stop only those PIDs after the outer
  deadline and must publish no passing evidence on timeout.

- Stress runners must snapshot and exclude pre-existing Excel PIDs and must
  never terminate a process they did not create. A concurrent user-launched
  Excel process is intentionally treated as an inconclusive observation; do
  not broaden cleanup commands to `Get-Process Excel` or a name-only kill.

- Office bitness validation must prove that its metadata COM object created a
  new Excel PID before calling `Quit`; if ownership is ambiguous, fail closed
  and leave every pre-existing Excel process untouched.

## Support-boundary decisions

- Treat unverified protocol or Office-bitness combinations as explicit policy
  boundaries before adding fixtures or release gates. Diagnostic flags and
  capability probes must not become compatibility claims; promotion runners,
  manifests, and risk registers must accept only evidence for supported rows.

## Release workflow and pinned tools

- Validate release-gate scripts with the exact pinned xlflow version used by
  GitHub Actions; a local development build may accept different VBA lint or
  analyzer rules.
- When checking that a GitHub Release is absent, use a successful listing/query
  that returns an empty result. Do not intentionally invoke a 404-producing
  command under PowerShell's fail-fast error policy.
- Wrap every required Task invocation and check `$LASTEXITCODE`; otherwise a
  failed source gate can be masked while later packaging steps continue.

## Hosted CI boundaries

- GitHub-hosted Windows runners do not provide Excel COM. Keep Excel/VBE gates
  such as `task test:xlam` out of Excel-free workflows and use an explicit
  `task test:clean-checkout:excel-free` target for archive validation.
