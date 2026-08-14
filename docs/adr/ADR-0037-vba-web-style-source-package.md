# ADR-0037: VBA-Web-style source package installation

## Status

`accepted`

## Context

The implementation is intentionally divided into many VBA classes and modules
so that transports, domain objects, reliability, streaming, and security rules
can be tested independently. That layout is appropriate for maintenance but
forces a consumer who receives raw source files to perform a large manual copy
operation. The existing workbook release path is also a poor primary library
distribution: it is a compiled artifact rather than source that can be
vendored into an existing consumer workbook.

VBA-Web addresses the same adoption problem with a source-oriented package and
an installer experience. VBA itself cannot place multiple class definitions in
one `.cls` component, so mechanically concatenating the project into a handful
of files would either remove public types or require a fragile generated parser.

## Decision

- Keep responsibility-oriented source splitting in the repository and in the
  development workbook.
- Make `dist/VBA-HTTP-source.zip` the primary module distribution. It contains
  a manifest, production source components, package documentation, license
  notices, and an installer/uninstaller.
- Generate the package from the checked-in production component policy. Do not
  include workbook document modules, scaffold entry modules, tests, benchmarks,
  xlflow helpers, or development modules.
- Use `Install-VBAHttp.ps1` as the installer entrypoint. It validates every
  manifest hash, requires a closed target workbook, creates a backup, imports
  only manifest components, and owns only the Excel instance it starts.
- Require `-Force` for replacing an existing component name. Never silently
  remove a consumer-owned module, and never enumerate or terminate unrelated
  Excel processes.
- Keep XLSM/XLAM `xlflow build` artifacts as optional compiled release targets
  for demonstrations, smoke tests, and consumers who prefer a referenced
  workbook. They are not the primary module distribution.
- A GitHub tag release attaches this source ZIP as the primary asset and may
  attach a separately provenance-labeled production-only `xlflow pack` XLSM;
  the latter is not presented as VBE-validated.
- Remove the empty xlflow scaffold entry modules `App`, `Main`, and `Ui`; the
  external consumer harness calls the public `VBAHttp` factories directly.

## Consequences

- Consumers install one archive instead of selecting dozens of files.
- The package is inspectable and vendorable, while the repository retains
  small, reviewable source files.
- Installation requires Windows Excel VBProject access and a closed target
  workbook. A workbook installer can be added later as a UI wrapper around the
  same manifest without changing the package contract.
- Component replacement is intentionally explicit and backup-based; arbitrary
  same-name consumer modules are not automatically recoverable without the
  backup.
- Release automation now has a source-package gate in addition to workbook
  build gates.

## Evidence

- Current component policy: `tools/build-component-policy.json`.
- Package generator and validator: `tools/New-SourcePackage.ps1`,
  `tools/Validate-SourcePackage.ps1`.
- Installer boundary: `tools/Install-VBAHttp.ps1` and
  `tools/Uninstall-VBAHttp.ps1`.
- Package contract: `docs/specs/source-package.md` and
  `tools/Test-SourcePackage.ps1`.

## Supersedes

- None

## Superseded by

- None
