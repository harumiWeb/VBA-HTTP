# Distribution specification

## Distribution artifacts and responsibilities

The repository maintains five distinct artifact responsibilities:

1. `dist/VBA-HTTP-source.zip` is the primary library distribution. It is a
   manifest-verified source package with an installer for existing consumer
   workbooks.
2. The tracked `build/VBA-HTTP.xlsm` is the development workbook. It mirrors
   source through xlflow and contains production code, unit/integration/stress
   tests, benchmarks, and xlflow support modules.
3. `build/Release/VBA-HTTP.xlsm` is an optional generated consumer artifact. It is built
   from the development workbook with `xlflow build` and contains only the
   production allowlist. Tests, benchmarks, Xlflow helpers, and Dev modules are
   excluded by `xlflow.toml`.
4. A GitHub tag release bundles the source package, a production-only
   `xlflow pack --experimental` XLSM, pack/release manifests, SHA-256 sums,
   `LICENSE`, and `THIRD_PARTY_NOTICES.md`. The pack manifest explicitly says
   `vbe_validation=not_performed`; GitHub Actions never starts Excel.
5. `build/VBA-HTTP.xlam` is the tracked same-extension add-in base. Its
   independent `build/Release/VBA-HTTP.xlam` target uses the same production
   allowlist but has a separate manifest, checksum, add-in identity check, and
   smoke command.

Source under `src/` is authoritative. The workbook is never edited manually as
the source of truth, and release modules are never removed by hand.

## Reproducible release flow

From a clean Windows checkout:

```powershell
task verify
task release:source
```

The source package does not start Excel. Workbook release commands additionally
require Excel/VBIDE:

```powershell
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
and validates identity and component policy. The XLSM external consumer harness
remains the primary runtime smoke path; an XLAM-specific consumer smoke may be
run separately after promotion. It does not overwrite either the XLSM base or
the XLSM release artifact. The decision and
validation boundary are recorded in ADR-0021.

## Source-package installation and upgrade

The module distribution follows the VBA-Web installation model. Unzip the
package and point the installer at a closed target workbook:

```powershell
Expand-Archive .\VBA-HTTP-source.zip -DestinationPath .\VBA-HTTP-source
.\VBA-HTTP-source\Install-VBAHttp.ps1 -Workbook .\MyApplication.xlsm
```

The installer validates the manifest and every SHA-256 record before opening
Excel, writes a complete `.vba-http.bak` backup, and imports the production
`.bas`/`.cls` components in one operation. Existing package component names
require `-Force` for replacement. `-WhatIf` validates the package and reports
the intended action without opening Excel. `Uninstall-VBAHttp.ps1` uses the
same manifest and backup boundary and also requires `-Force` before removing
existing components.

The installer starts one hidden Excel instance for the explicitly named target
workbook and closes only that instance. It does not enumerate or terminate
unrelated Excel processes. The target must be closed before installation and
must permit VBProject access. Consumers should keep the generated backup until
the new version passes their own smoke test.

## Workbook consumer installation and upgrade

For a consumer, copy the verified release workbook to a controlled add-in or
library location and reference it from the consuming VBA project. Call
`VBAHttp.CreateClient` (or another factory) across the workbook reference;
public classes are `PublicNotCreatable` by design.

To upgrade, stage the new workbook beside the current version, verify its
manifest/checksum and run the external smoke harness, then replace the old
file atomically while Excel is closed. Keep the previous verified artifact for
rollback. Do not copy test modules from the development workbook into the
consumer workbook.

Source-vendored consumers may instead use the package installer or import the
`components/` files from the package. They own their workbook packaging and
must retain the same excluded-component and source-boundary rules. The vendored
`references/VBA-Web` tree is a benchmark comparator, not a product dependency;
its upstream specs are not run against external network services.

## Evidence retention

Source-package evidence consists of the exact ZIP, extracted `manifest.json`,
source revision, and package hashes. Workbook release evidence consists of the
exact workbook, `.build.json` manifest,
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

## GitHub tag release

The authoritative tag-release contract is [`docs/specs/github-release.md`](github-release.md).
For a strict `vX.Y.Z` or prerelease tag, the GitHub-hosted Windows x64 workflow
validates the tag commit, runs Excel-free gates, stages only the production
allowlist for `xlflow pack`, and publishes the exact seven-asset set with
`gh release create --verify-tag`. The release notes state that VBE validation
was not performed. The local `task precommit` VBE gate and optional
`task release:build` artifact remain separate, Excel-backed evidence.
