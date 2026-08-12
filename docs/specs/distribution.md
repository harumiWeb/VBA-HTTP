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

Source under `src/` is authoritative. The workbook is never edited manually as
the source of truth, and release modules are never removed by hand.

## Reproducible release flow

From a clean Windows checkout with Excel/VBIDE available:

```powershell
task verify
task release:build
task release:security
```

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
  harness.

The base workbook is not overwritten. If any build, validation, or smoke step
fails, an existing published release remains untouched. Generated release
files are ignored by Git and must be archived together with their manifest,
checksum, and security report when a release is handed off.

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

## Future XLAM target

No tracked `.xlam` base artifact exists yet. Adding an XLAM distribution is a
separate build target: first add and verify a same-extension base workbook,
then add an explicit `xlflow build` output and its own manifest/checksum/smoke
contract. The XLSM release flow must not silently change extension or reuse an
unverified base.

## Evidence retention

Release evidence consists of the exact workbook, `.build.json` manifest,
`.checksum.json` sidecar, `release-security.json` report, and smoke output.
Evidence is path-stable and secret-free. Record the source revision and host
compatibility (Office bitness, Windows version, protocol fixture) beside the
bundle; x64 is the currently verified path, while 32-bit and HTTP/2/HTTP/3
negotiation remain explicit compatibility gates.
