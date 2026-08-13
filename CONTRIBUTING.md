# Contributing to VBA-HTTP

VBA-HTTP is maintained from exported VBA source. The tracked
`build/VBA-HTTP.xlsm` workbook is the development/verification target; it is
not the place to edit production code.

## Prerequisites

- Windows with Excel and VBIDE access (the current proof path is x64 Office).
- Go for the deterministic loopback test server.
- `xlflow`, Task, Lefthook, and PSScriptAnalyzer available on `PATH`.
- A repository checkout with the tracked development workbook present.

Run `task hooks:install` once to install the pre-commit hook. The hook stops on
any VBA lint/analyzer, PowerShell analyzer, formatter, or class-source failure.

## Source layout

- `src/classes/`: public and internal class modules.
- `src/modules/`: standard modules; tests are under `src/modules/Tests/`.
- `src/workbook/`: document modules.
- `tools/`: deterministic server, build validators, and external smoke harnesses.
- `docs/adr/`: architectural rationale; `docs/specs/`: current contracts.

Keep benchmark modules in `src/modules/Benchmarks/` and temporary development
helpers in `src/modules/Dev/`. Production components must not depend on those
excluded directories.

## Change and proof loop

1. Update the relevant ADR/spec before changing a public contract.
2. Implement the smallest focused change and add a regression or focused test.
3. Run `task check` (the same gates used by Lefthook, including
   `task powershell:lint`).
4. Start or attach an xlflow session, then run:

   ```powershell
   xlflow push --fast --session --no-save --json
   task test
   xlflow save --session --json
   xlflow session stop --json
   ```

5. For transport changes, run `task test:integration` and the applicable
   benchmark or stress task. Store machine-readable evidence under
   `benchmarks/results/`.
6. Confirm `xlflow status --json` reports a synchronized source/workbook pair.

Do not add external endpoints to deterministic tests. Host-specific required
HTTP/2 evidence is collected separately with `task protocol:host` only after a
release artifact is built and the operator has selected a trusted TLS endpoint;
HTTP/3/QUIC probing is diagnostic-only under ADR-0035. The runner records no
URL path, query, credentials, or body.

Never save a workbook after a failed proof. Follow xlflow recovery guidance if
the session reports a recovery-required state.

## Release artifact

The development workbook contains tests and tooling. A distributable workbook
is generated separately with `task release:build`, which runs `xlflow build`
with the configured exclusions, VBE compile/save/close, checksum generation,
manifest security validation, component inspection, and an external consumer
smoke harness. Do not remove modules manually and do not use `xlflow pack` for
the formal release path.

See [`docs/specs/distribution.md`](docs/specs/distribution.md) and
[`docs/RELEASE_CHECKLIST.md`](docs/RELEASE_CHECKLIST.md) for the artifact and
handoff rules.

## License and notices

Project-authored material is released under the MIT License in
[`LICENSE`](LICENSE). Keep [`THIRD_PARTY_NOTICES.md`](THIRD_PARTY_NOTICES.md)
with source or workbook handoffs. The ignored VBA-Web checkout is a pinned MIT
benchmark comparator only and is never a product dependency or release
component.

## Commit checklist

- focused and affected tests pass;
- `task check` is clean;
- source and the development workbook are synchronized;
- public API changes update README, API/spec docs, and CHANGELOG;
- performance claims include before/after evidence;
- release changes include the manifest, checksum, and security-risk checks.
