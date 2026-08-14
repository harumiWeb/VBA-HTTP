# Release checklist

This checklist covers the source package and optional verified workbook
artifacts. It is intentionally separate
from `tasks/todo.md`: the roadmap tracks implementation progress, while this
document records the repeatable handoff gate.

For an automated GitHub tag release, also use the exact asset and provenance
contract in [`docs/specs/github-release.md`](specs/github-release.md). The
GitHub pack XLSM is production-only but explicitly **not VBE-validated**.

## Before building

- [ ] The source revision and intended version are recorded.
- [ ] `task check` is clean (lint, analyze, format, and class-source checks).
- [ ] `task test`, `task test:integration`, and the applicable benchmark/stress
      tasks are green against the deterministic loopback server.
- [ ] `task test:security-risks` reports zero current release blockers.
- [ ] `task test:license` passes; `LICENSE` and
      `THIRD_PARTY_NOTICES.md` identify the project and comparator boundary.
- [ ] `task test:source-package` passes and `task release:source` produces a
      manifest-verified source archive with installer, hashes, and notices.
- [ ] Any public API change has matching spec, README/API, example, and
      CHANGELOG updates.
- [ ] The compatibility record names the Windows version, Office bitness, and
      protocol evidence. The current supported proof path is x64 Office.
      For the supported negotiated HTTP/2 host row, archive the validated output
      from `task protocol:host` with the exact artifact and checksum. HTTP/3 is
      unsupported by policy under ADR-0035 and has no release evidence gate.
- [ ] `task test:clean-checkout` passes; for a clean-environment release
      evidence bundle, `tools/Test-CleanCheckout.ps1 -FullRelease` also passes.

## GitHub tag release (Excel-free)

- [ ] `task precommit` has passed locally on the supported x64 bridge.
- [ ] The tag matches the strict `vX.Y.Z`/prerelease expression and points to
      the intended commit.
- [ ] `task test:github-release` passes without starting Excel.
- [ ] The workflow uses `windows-2022`, `contents: write`, and the pinned
      `tools/release-toolchain.json`.
- [ ] Production-only pack staging and module counts match the allowlist.
- [ ] Pack provenance says `pure-go`, `experimental=true`, and
      `vbe_validation=not_performed`.
- [ ] The exact seven release assets are present and SHA-256 verified.
- [ ] Release notes explicitly state that VBE validation was not performed.
- [ ] `gh release view` confirms the tag, stable/prerelease state, and asset
      set; an existing Release was not overwritten.

## Source package handoff

Run:

```powershell
task release:source
```

Archive `dist/VBA-HTTP-source.zip` with its source revision. Validate it with
`Validate-SourcePackage.ps1` after extraction. The archive is the primary
library distribution and does not contain tests, workbook document modules, or
the old scaffold entry modules.

## Optional workbook build and inspect

Run from the repository root:

```powershell
task release:build
task release:security
```

The commands must produce, in `build/Release/`:

- `VBA-HTTP.xlsm`;
- `VBA-HTTP.xlsm.build.json`;
- `VBA-HTTP.xlsm.checksum.json`; and
- a corresponding security report under `.xlflow/release-security/`.

The build manifest must show the tracked `build/VBA-HTTP.xlsm` as its base,
successful source application, VBE compile/save/close/cleanup, and atomic
publication. The included component list must match
`tools/build-component-policy.json`; no `Tests`, `Benchmarks`, `Xlflow`, or
`Dev` component may be present in the artifact.

## Consumer verification

- [ ] `task release:smoke` opens the generated workbook through the external
      consumer harness without adding test code to it.
- [ ] Public factory calls, buffered GET, batch, retry/deadline, protocol,
      decompression, proxy, auth, cookie, diagnostics, download, and upload
      smoke scenarios pass where the host supports the capability.
- [ ] The artifact, manifest, checksum, security report, and smoke output are
      archived together with `LICENSE` and `THIRD_PARTY_NOTICES.md`. No
      secrets, request bodies, or host-local paths are included in the evidence
      bundle.
- [ ] The previous verified artifact remains available for rollback until the
      new workbook is accepted.

For an add-in release, run the independent target after the XLSM checklist:

```powershell
task test:xlam
task release:xlam:build
```

Archive `build/Release/VBA-HTTP.xlam`, its `.build.json` manifest, its checksum
sidecar, and the XLAM validation output separately from the XLSM evidence. The
validator must prove `Workbook.IsAddin=True` and must not reuse the XLSM base.

## Deferred promotion gates

The current XLSM/XLAM release is not evidence that every v1.0 compatibility target
is complete. Promotion of the v1.0 line additionally requires the remaining
risk-register item for integrated/proxy challenge authentication, plus any
other supported-target compatibility obligations, to be resolved with dedicated
fixtures and evidence. HTTP/3/QUIC is outside the supported distribution
boundary under ADR-0035. 32-bit Office is `unverified` under ADR-0039 and has
no official release guarantee; contributor diagnostic evidence is welcome but
is not a v1 promotion pass. The remaining deferred items are recorded in
`docs/security/risk-register.json` and must be re-evaluated before a v1
promotion.
