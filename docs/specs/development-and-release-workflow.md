# Development and release workflow

## Source layout and authority

- Edit VBA only under `src/`.
- Standard modules live in `src/modules/`, classes in `src/classes/`, forms in `src/forms/`, and workbook document modules in `src/workbook/`.
- Tests live in `src/modules/Tests/`. Unit, integration, and stress suites use `Unit/`, `Integration/`, and `Stress/` subdirectories respectively.
- Integration tests use the `integration` tag. Stress tests use the `stress` tag and are not part of the routine focused suite.
- Benchmark modules live in `src/modules/Benchmarks/`; temporary development helpers live in `src/modules/Dev/`.
- `build/VBA-HTTP.xlsm` is the tracked development workbook. It is synchronized from source and used to compile and execute VBA; it is not the editing authority.
- Git does not preserve empty directories. `task bootstrap:dirs` restores configured empty `src/classes/` and `src/forms/` roots before release planning or build; do not add non-VBA placeholder files to source roots.
- Exported `.cls` files use UTF-8 without BOM and CRLF records. The VBIDE can treat an otherwise valid LF-only class export as a standard module during a clean import, leaving the `VERSION 1.0 CLASS` header as executable text. `.gitattributes` enforces CRLF and `task class-source:check` rejects a non-importable worktree before build. Run `task class-source:normalize` after changing this policy in an existing checkout. Because the current xlflow formatter renders LF, `task format` checks content-equivalent temporary LF projections of every class while the original CRLF files are independently checked by `task class-source:check`; class formatting is not excluded.

If VBA is changed directly in the VBE during exceptional diagnosis or recovery, stop source editing, run `xlflow status --json`, pull the workbook source, and reconcile the resulting source diff before continuing.

## Development proof loop

Run the following from the repository root:

1. `xlflow status --json`
2. `xlflow session start --json` for a closed workbook, or attach to the configured user-open workbook.
3. `task check`
4. `xlflow push --fast --session --no-save --json`
5. Run the focused test in the live session. The routine full suite is `task test`, which uses module isolation so module-level VBA state and scheduled host callbacks cannot leak across suites.
6. `xlflow save --session --json`
7. `xlflow session stop --json`
8. Confirm source/workbook status is clean.

Never save after a failed proof. Follow xlflow recovery instructions when status reports that recovery is required.

### Temporary formatter exception

The current local `xlflow` development build rewrites `XlflowAssert.bas` into a form that its own lint command reports as `VB014 parser_recovery`, although symbol inspection reports a complete parse. `tools/Check-Format.ps1` therefore excludes only this bundled helper and checks every other `.bas` and `.cls` file. Do not broaden this exception. Remove it once the formatter output passes both `xlflow lint --json` and Excel compilation.

## Release component boundary

`xlflow.toml` excludes Tests, Benchmarks, Xlflow, and Dev modules only during `xlflow build`. Ordinary push operations keep them in the development workbook.

Production source must not call, instantiate, or otherwise depend on an excluded component. VBE compilation of the filtered workbook is the final dependency check. `tools/build-component-policy.json` is the explicit allowlist/denylist for the current component set; update it intentionally whenever components are added or renamed.

Before a release build:

1. Run `task verify`.
2. Inspect the dry-run included and excluded component sets.
3. Run `xlflow build --json --out build/Release/VBA-HTTP.xlsm`.
4. Run `task release:checksum` to generate the deterministic
   `VBA-HTTP.xlsm.checksum.json` sidecar for the artifact and build manifest.
5. Verify the build manifest reports successful source application, VBE compile, save, close, and atomic publication, and verify both SHA-256 values in the sidecar.
6. Run `task release:smoke` to inspect the actual VBA component collection and call `Main.Run` without injecting test code into the artifact.

Generated release artifacts and staging files are not committed. The tracked development workbook is never overwritten by a release build.

The checksum sidecar contract, schema, and failure behavior are defined in
`release-checksum.md`. The sidecar is generated and replaced atomically in the
same directory as the artifact; a stale or missing sidecar fails the external
release smoke gate.

## Documentation gates

- Public API changes update the relevant spec, README, examples, and CHANGELOG.
- Bug fixes add a regression test and update a spec when the invariant would otherwise be easy to lose.
- Performance changes preserve same-condition before/after benchmark results.
- Architectural trade-offs are recorded in `docs/adr/`; current behavior belongs in `docs/specs/`.
