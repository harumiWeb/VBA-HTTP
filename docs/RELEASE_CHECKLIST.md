# Release checklist

This checklist is for a verified XLSM artifact. It is intentionally separate
from `tasks/todo.md`: the roadmap tracks implementation progress, while this
document records the repeatable handoff gate.

## Before building

- [ ] The source revision and intended version are recorded.
- [ ] `task check` is clean (lint, analyze, format, and class-source checks).
- [ ] `task test`, `task test:integration`, and the applicable benchmark/stress
      tasks are green against the deterministic loopback server.
- [ ] `task test:security-risks` reports zero current release blockers.
- [ ] Any public API change has matching spec, README/API, example, and
      CHANGELOG updates.
- [ ] The compatibility record names the Windows version, Office bitness, and
      protocol evidence. The current supported proof path is x64 Office.
- [ ] `task test:clean-checkout` passes; for a clean-environment release
      evidence bundle, `tools/Test-CleanCheckout.ps1 -FullRelease` also passes.

## Build and inspect

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
      archived together. No secrets, request bodies, or host-local paths are
      included in the evidence bundle.
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
is complete. Promotion of the v1.0 line additionally requires the risk-register
items for HTTP/2/HTTP/3 and 32-bit Office, plus integrated/proxy challenge
authentication, to be resolved with dedicated fixtures and evidence. The
