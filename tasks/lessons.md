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
