# Distribution specification

## Two artifacts, two responsibilities

The repository maintains two distinct deliverables:

1. The tracked `build/VBA-HTTP.xlsm` is the development workbook. It mirrors
   source through xlflow and contains production code, unit/integration/stress
   tests, benchmarks, and xlflow support modules.
2. `build/Release/VBA-HTTP.xlsm` is a generated consumer artifact. It is built
   from the development workbook with `xlflow build` and contains only the
   production allowlist. Tests, benchmarks, Xlflow helpers, and Dev modules are
   excluded by `xlflow.toml`.
3. `build/VBA-HTTP.xlam` is the tracked same-extension add-in base. Its
   independent `build/Release/VBA-HTTP.xlam` target uses the same production
   allowlist but has a separate manifest, checksum, add-in identity check, and
   smoke command.

Source under `src/` is authoritative. The workbook is never edited manually as
the source of truth, and release modules are never removed by hand.

## Reproducible release flow

From a clean Windows checkout with Excel/VBIDE available:

```powershell
task verify
task release:build
task release:security
```

`task test:clean-checkout` creates a temporary `git archive` of `HEAD` and
re-runs source checks, documentation checks, both dry-run build plans, and the
tracked XLSM/XLAM base identity checks without relying on ignored local files.
Use `tools/Test-CleanCheckout.ps1 -FullRelease` when a clean-environment
release build and both consumer artifacts are required as evidence.

`task release:build` performs the following in order:

- validates the exact production include/exclude plan;
- runs `xlflow build --json --out build/Release/VBA-HTTP.xlsm` from the tracked
  development workbook;
- requires VBE compile, save, close, cleanup, and atomic publication;
- writes `VBA-HTTP.xlsm.build.json` and a SHA-256
  `VBA-HTTP.xlsm.checksum.json` sidecar;
- validates manifest paths, component policy, checksum, and deferred-risk
  register; and
- opens the generated workbook only through the external consumer smoke
  harness. The harness covers the HTTP proxy and the deterministic HTTPS
  CONNECT boundary for both COM and native public factories; it expects the
  intentionally untrusted TLS certificate to be rejected.

All release validation Excel instances are ownership-checked against the
pre-existing Excel PID set. Cleanup may terminate only a PID proven to have
been created by that validation run; unrelated user Excel processes are never
selected as cleanup targets.

The base workbook is not overwritten. If any build, validation, or smoke step
fails, an existing published release remains untouched. Generated release
files are ignored by Git and must be archived together with their manifest,
checksum, and security report when a release is handed off.
The handoff must also include the root `LICENSE` and
`THIRD_PARTY_NOTICES.md`; these package files are not embedded in the workbook
by `xlflow build`.

## XLAM build target

The add-in target is intentionally explicit:

```powershell
task test:xlam
task release:xlam:build
```

The target runs `xlflow build --base build/VBA-HTTP.xlam --out
build/Release/VBA-HTTP.xlam`, requires the filtered component policy and atomic
VBE build evidence, verifies the SHA-256 sidecar, checks `Workbook.IsAddin`,
and runs `Main.Run` through the external xlflow consumer path. It does not
overwrite either the XLSM base or the XLSM release artifact. The decision and
validation boundary are recorded in ADR-0021.

## Consumer installation and upgrade

For a consumer, copy the verified release workbook to a controlled add-in or
library location and reference it from the consuming VBA project. Call
`VBAHttp.CreateClient` (or another factory) across the workbook reference;
public classes are `PublicNotCreatable` by design.

To upgrade, stage the new workbook beside the current version, verify its
manifest/checksum and run the external smoke harness, then replace the old
file atomically while Excel is closed. Keep the previous verified artifact for
rollback. Do not copy test modules from the development workbook into the
consumer workbook.

Source-vendored consumers may instead import the `src/` components and run the
full xlflow proof loop. Such consumers own their workbook packaging and must
retain the same excluded-component and source-boundary rules. The vendored
`references/VBA-Web` tree is a benchmark comparator, not a product dependency;
its upstream specs are not run against external network services.

## Evidence retention

Release evidence consists of the exact workbook, `.build.json` manifest,
`.checksum.json` sidecar, `release-security.json` report, smoke output,
`LICENSE`, and `THIRD_PARTY_NOTICES.md`.
Evidence is path-stable and secret-free. Record the source revision and host
compatibility (Office bitness, Windows version, protocol fixture) beside the
bundle; x64 is the supported and currently verified path, and the x64 HTTP/2
host record is archived separately. HTTP/3/QUIC is unsupported by policy under
ADR-0035 and is not a release gate. 32-bit Office is outside the supported
distribution boundary under ADR-0030. Keep the XLSM and XLAM
manifest/checksum pairs separate so an artifact cannot be validated against
the wrong base or extension.
