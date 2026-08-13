# Licensing specification

## Canonical project license

Project-authored VBA source, documentation, examples, deterministic tooling,
source packages, and generated release workbooks are distributed under MIT. The complete grant
and warranty disclaimer are the root [`LICENSE`](../../LICENSE) file; its
copyright holder is currently `harumiWeb` for 2026.

No per-file header is required. A file with a more specific copyright or
third-party notice is governed by that notice for the indicated material.

## Packaging contract

Every source or workbook distribution must retain the root `LICENSE` and
`THIRD_PARTY_NOTICES.md`. A source package must additionally retain its
`manifest.json` beside the component payload.
- the release `.build.json` manifest, checksum sidecar, and security report

The filtered workbook is produced by `xlflow build`; license files are package
files and are not inserted into or removed from the workbook by the build. A
consumer must preserve the MIT copyright and permission notice in copies or
substantial portions of the project-authored material.

## Third-party and platform boundary

The ignored pinned `references/VBA-Web` tree is only a benchmark comparator.
It is not copied into development or release workbooks, and its upstream MIT
license remains authoritative for that checkout. Windows, Excel, WinHTTP, and
xlflow are external platform/build dependencies and are not relicensed or
redistributed by this project. The test server currently declares no external
Go modules beyond the standard library.

Any future third-party material in a release must be recorded in
`THIRD_PARTY_NOTICES.md`, retain its original notices, and pass provenance and
license verification before release promotion.

## Verification

`task test:license` runs `tools/Test-LicenseContract.ps1`. The test is part of
`task check` and therefore the Lefthook pre-commit gate. It checks the canonical
MIT markers, required public documentation, and the third-party inventory.
