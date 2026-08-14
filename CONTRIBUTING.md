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
Pull requests targeting `main`, pushes to `main`, and the weekly scheduled
workflow run the hosted Excel-free CI contract in
[`docs/specs/github-ci.md`](docs/specs/github-ci.md). That workflow does not
start Excel or provide VBE compilation evidence; local `task precommit` remains
the Excel-backed pre-commit boundary.

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

### 32-bit Office validation contributions

32-bit Office is currently `unverified`, not known-incompatible. The official
release target remains Windows x64 Office, so the normal promotion gate does
not switch to an x86 bridge. Contributors with a real 32-bit Office host may
run the non-promotional diagnostic path:

```powershell
powershell -File tools/Run-OfficeBitnessValidation.ps1 `
  -ExpectedArchitecture X86 -DiagnosticOnly
```

Attach the resulting JSON and host/tool versions to a compatibility proposal;
do not label it as a release pass. The runner snapshots existing Excel PIDs,
proves ownership of the Excel instance it starts, and fails closed when that
ownership is ambiguous. Promotion to `community-validated` requires a
reproducible real-host bundle and maintainer review; official support requires
a later ADR and the full compile, test, integration, filtered-build, and
consumer-smoke evidence described in ADR-0039.

Never save a workbook after a failed proof. Follow xlflow recovery guidance if
the session reports a recovery-required state.

## Source package and release artifacts

The primary consumer handoff is the VBA-Web-style module package:

```powershell
task release:source
```

This creates `dist/VBA-HTTP-source.zip` with a manifest, production `.bas` and
`.cls` files, installer/uninstaller scripts, hashes, and license notices. The
package installer targets a closed workbook and owns only the Excel instance
it starts. Use `-Force` for component replacement and retain the generated
backup. See [`docs/specs/source-package.md`](docs/specs/source-package.md).

Compiled workbooks remain optional release artifacts. For a tag push, the
GitHub-hosted workflow creates the source ZIP and a production-only
`xlflow pack --experimental` XLSM without starting Excel; its release notes
explicitly say VBE validation was not performed. Run `task precommit` locally
for the ownership-safe x64 VBE compile gate.

The development workbook contains tests and tooling. A distributable workbook
is generated separately with `task release:build`, which runs `xlflow build`
with the configured exclusions, VBE compile/save/close, checksum generation,
manifest security validation, component inspection, and an external consumer
smoke harness. Do not remove modules manually. The GitHub tag path uses
`xlflow pack` only through production-only staging and has a separate
not-VBE-validated provenance contract.

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
