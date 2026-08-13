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
