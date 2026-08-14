# Distribution and upgrades

## Primary distribution: source modules

VBA-HTTP is distributed primarily as a VBA-Web-style source package. A
versioned package contains production `.bas` and `.cls` modules, a manifest,
component hashes, installer/uninstaller scripts, `LICENSE`, and
`THIRD_PARTY_NOTICES.md`. Consumers copy the package into a closed workbook;
they do not need to copy tests, benchmarks, xlflow helpers, or development
modules.

The package is the preferred distribution when a consumer wants source
visibility, code signing, or a workbook-specific release process. Install with
`Install-VBAHttp.ps1`, keep the generated backup, and use `-Force` only for an
intentional upgrade. The uninstaller and backup provide rollback without
editing modules manually.

## Optional workbook artifacts

The tracked `build/VBA-HTTP.xlsm` is a development/verification workbook, not
the primary consumer handoff. `xlflow build` can generate a filtered XLSM with
VBE compile/save/close evidence. The same-extension XLAM target is independent
and must use its own base artifact.

The GitHub-hosted tag workflow has a different boundary: it stages only
production modules, runs `xlflow pack --experimental` without Excel, and marks
the pack manifest `vbe_validation=not_performed`. The XLSM attached to a GitHub
Release must not be described as VBE-validated. Local `task precommit` and
`task release:build` remain the Excel-backed validation paths.

## GitHub Release asset set

For tag `vX.Y.Z`, the release contains exactly:

```text
VBA-HTTP-vX.Y.Z-source.zip
VBA-HTTP-vX.Y.Z.xlsm
VBA-HTTP-vX.Y.Z.pack.json
VBA-HTTP-vX.Y.Z.release.json
VBA-HTTP-vX.Y.Z.SHA256SUMS.txt
LICENSE
THIRD_PARTY_NOTICES.md
```

The source ZIP is the primary distribution. Pack/release manifests record the
tag commit, production allowlist, module counts, tool version, support
boundary, and `vbe_validation=not_performed`; SHA-256 checksums are regenerated and
verified before publication. Existing Releases are never overwritten.

Download a release with GitHub CLI:

```powershell
gh release download v1.2.3 --repo <owner>/<repo> --dir .\vba-http-v1.2.3
Get-Content .\vba-http-v1.2.3\VBA-HTTP-v1.2.3.SHA256SUMS.txt
```

Verify the checksum before extracting or installing the source package. Keep
the license and notice files with redistributed modules.

## License and third-party material

Project-authored source, guides, specifications, tools, examples, and generated
artifacts are MIT-licensed. The pinned VBA-Web checkout is a benchmark-only
comparator and is not a product dependency. See [`../../LICENSE`](../../LICENSE),
[`../../THIRD_PARTY_NOTICES.md`](../../THIRD_PARTY_NOTICES.md), and
[`../specs/licensing.md`](../specs/licensing.md).

## Release boundaries

The package targets Windows x64 Office. 32-bit Office is currently unverified
and is not a release guarantee; HTTP/3/QUIC remains unsupported by policy, and
unverified host-specific protocol combinations are not promoted by the release
pipeline. Review [Compatibility](compatibility.md) and the normative
[`../specs/github-release.md`](../specs/github-release.md) before publishing a
consumer-facing claim.
