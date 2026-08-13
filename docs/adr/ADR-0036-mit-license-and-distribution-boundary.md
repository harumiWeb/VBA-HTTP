# ADR-0036: MIT license and distribution boundary

## Status

`accepted`

## Background

The repository previously contained an empty root `LICENSE`, so consumers had
no authoritative license grant for the project-authored VBA source, tools, and
documentation. The generated release workbooks are distributed separately from
the development workbook, and the ignored `references/VBA-Web` checkout is
used only as a benchmark comparator. These boundaries must not accidentally
relicense third-party or platform material.

## Decision

- Project-authored source, documentation, deterministic test-server code,
  scripts, examples, and generated release workbooks are released under the
  MIT license in the root `LICENSE` file.
- `LICENSE` is the canonical license text. Per-file license headers are not
  required; new project-authored files inherit the root license unless they
  carry a more specific notice.
- A release handoff includes `LICENSE` and
  `THIRD_PARTY_NOTICES.md` beside the verified workbook, manifest, checksum,
  and security evidence. `xlflow build` filters workbook components but does
  not embed or replace the repository license files.
- The ignored pinned `references/VBA-Web` checkout is a benchmark input, not a
  product dependency or release component. Its upstream MIT notice remains in
  that checkout and is not relicensed by this repository.
- Windows, Excel, WinHTTP, xlflow, and other tools supplied by the operator are
  platform or build dependencies, not redistributed project components. Their
  own terms remain applicable.
- Adding any third-party source or binary to a release artifact requires a
  notice entry, provenance/license verification, and an ADR/spec update before
  the artifact can be promoted.

## Consequences

- Consumers may reuse, modify, and redistribute project-authored materials
  under the permissive MIT terms while preserving the copyright and permission
  notices.
- The license does not grant Microsoft, VBA-Web, or platform trademarks and
  does not change the support boundaries for x64 Office, HTTP/2, or the
  unsupported 32-bit Office and HTTP/3/QUIC combinations.
- Release automation must verify the license files as a packaging contract;
  a workbook-only handoff is incomplete.

## Rationale

- Canonical license text: `LICENSE`.
- Current license contract: `docs/specs/licensing.md`.
- Third-party inventory: `THIRD_PARTY_NOTICES.md`.
- Release boundary: `docs/specs/distribution.md` and
  `docs/RELEASE_CHECKLIST.md`.
- Comparator provenance and upstream-license verification:
  `tools/Setup-VBAWeb.ps1` and `docs/specs/benchmark-methodology.md`.
- Automated contract: `tools/Test-LicenseContract.ps1` and the `task check`
  pre-commit gate.

## Supersedes

- None

## Superseded by

- None
