# Source package and installer specification

## Package purpose

The primary consumer distribution is a source package, following the
installation experience of VBA-Web. Development remains split by architectural
responsibility under `src/classes/` and `src/modules/`; consumers receive one
manifest-verified archive and do not select individual VBA files by hand.

The package targets Windows x64 Office for its official release guarantee.
32-bit Office is currently unverified under ADR-0039, and HTTP/3/QUIC is unsupported under
ADR-0035.

## Package contents

`task package:source` produces `dist/VBA-HTTP-source.zip`. The archive contains:

- `manifest.json`, schema version 1;
- `benchmarks/schema/source-package-manifest.schema.json`;
- production `.cls` and `.bas` components only;
- `Install-VBAHttp.ps1`, `Uninstall-VBAHttp.ps1`, and
  `Validate-SourcePackage.ps1`;
- `README.md`, `CONTRIBUTING.md`, `CHANGELOG.md`, the current `docs/specs/`
  contracts, `docs/API.md`, `LICENSE`, and `THIRD_PARTY_NOTICES.md`.

Workbook document modules (`ThisWorkbook`, sheets), xlflow scaffold entry
modules (`App`, `Main`, `Ui`), tests, benchmarks, and development helpers are
not source-package components. The package source paths retain provenance,
while the package paths use `components/classes/` and `components/modules/` so
the installer can import them deterministically.

Each component and package file has a lowercase SHA-256 record in the manifest.
The installer validates every record before opening Excel. A package with a
missing file, path traversal, duplicate name, development-only path, or hash
mismatch fails closed.

The manifest `source_revision` is resolved from the explicit
`-SourceRevision` argument when supplied. The clean-checkout contract passes
the archived commit through `VBA_HTTP_SOURCE_REVISION` because a `git archive`
does not contain a `.git` directory; that fallback must be a full hexadecimal
Git revision and is never used to bypass normal working-tree revision checks.

## Installation contract

The target workbook must be closed and writable (the installer probes an
exclusive file lock) and must grant VBProject access to the automation process.
The installer creates a backup before any import, opens
only the explicitly supplied workbook in its own hidden Excel instance, imports
all manifest components, saves, closes, and releases that instance. It never
enumerates or terminates unrelated Excel processes.

```powershell
Expand-Archive .\VBA-HTTP-source.zip -DestinationPath .\VBA-HTTP-source
.\VBA-HTTP-source\Install-VBAHttp.ps1 -Workbook .\MyApplication.xlsm
```

An existing component with a manifest name is not replaced implicitly. Pass
`-Force` to authorize replacement; the installer still creates a complete
workbook backup first. `-WhatIf` validates the package and reports the intended
target without opening Excel. Uninstall uses the same manifest and backup
boundary. Pass `-Force` to remove the named package components; without it, an
existing component causes a fail-closed refusal rather than an implicit
deletion.

## Upgrade and rollback

Install a new package beside the current one, verify its manifest and source
revision, and run the installer with `-Force` only after reviewing the backup
path. Roll back by closing the workbook and restoring the backup copy. The
installer does not merge or overwrite consumer-owned modules with names that
are not in the package manifest.

## Verification

- `task test:source-package` creates a temporary archive, validates all hashes,
  checks scaffold/document exclusion, and rejects a tampered component.
- `task package:source` creates the handoff archive after the source gates.
- `Validate-SourceArchive.ps1` expands the published ZIP in a temporary
  directory and reruns the manifest/hash gate on the actual archive.
- `task release:source` runs `task check` and then creates the package.
- Workbook compile and runtime smoke remain release-workbook gates; the source
  installer itself is Excel-free until a consumer explicitly runs it.
