## Proof-loop safety

- Do not chain `xlflow push` and `xlflow save` in one shell command. If compile
  fails after import, stop the session without saving, inspect the diagnostic,
  fix source, then rerun the proof loop.
