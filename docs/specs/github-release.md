# GitHub tag release specification

## Trigger and support boundary

Only a push of a tag matching the following expression is eligible:

```text
^v(0|[1-9][0-9]*)\.(0|[1-9][0-9]*)\.(0|[1-9][0-9]*)(-[0-9A-Za-z.-]+)?$
```

`+build` metadata is not supported. A suffix makes the GitHub Release a
prerelease; a tag without a suffix is stable. The tag object must resolve to
the workflow `GITHUB_SHA`. The runner is GitHub-hosted Windows x64
(`windows-2022`); the release target remains Windows x64 Office. 32-bit Office
is recorded as `unverified` and is not validated by this workflow; HTTP/3/QUIC
and XLAM remain outside this release contract.

## Excel-free release pipeline

The workflow runs source, documentation, security-risk, PowerShell, Go
test/vet, pack, and bundle gates without starting Excel. Tool versions and the
xlflow Windows archive SHA-256 are locked in `tools/release-toolchain.json`.

`New-PackArtifact.ps1` copies only the exact production component allowlist to
a temporary staging tree, runs `xlflow pack --experimental --template`, checks
module counts, and removes staging in a `finally` block. The generated pack
manifest must state:

```json
{
  "pack_backend": "pure-go",
  "experimental": true,
  "vbe_validation": "not_performed"
}
```

The pack path does not replace the local `xlflow build` VBE gate.

## Published assets

For tag `vX.Y.Z`, the bundle and GitHub Release contain exactly:

```text
VBA-HTTP-vX.Y.Z-source.zip
VBA-HTTP-vX.Y.Z.xlsm
VBA-HTTP-vX.Y.Z.pack.json
VBA-HTTP-vX.Y.Z.release.json
VBA-HTTP-vX.Y.Z.SHA256SUMS.txt
LICENSE
THIRD_PARTY_NOTICES.md
```

The source ZIP manifest records the tag and commit. The release manifest is
path-free and contains only tag, revision, tool/provenance, support boundary,
asset names, sizes, and SHA-256 values. `SHA256SUMS` covers all primary assets
and the release manifest. Validators reject missing, renamed, duplicated, or
tampered assets, development components in the pack, and any VBE provenance
other than `not_performed`.

The manifest's `toolchain` object records the locked xlflow, Task, Go, and
PSScriptAnalyzer versions; it never records a runner path, credential, URL
secret, or other host-specific detail.

## Publication and reruns

`gh release create <tag> --verify-tag` publishes only after validation. Stable
tags use a normal Release; suffix tags use `--prerelease`. The workflow first
checks that the Release does not already exist and never uses `--clobber`,
asset replacement, tag mutation, or automatic deletion. A failed publication
does not alter an existing Release; an administrator reviews any partial draft
before deletion or rerun.

## Local VBE boundary

Lefthook calls one `task precommit` command. It runs `task check`, confirms the
x64 bridge and non-recovery state, refuses a locked/open target workbook,
records existing Excel PIDs, and builds an atomic temporary workbook with
`xlflow build`. It requires source application, VBE compile, save, close,
cleanup, and atomic publication evidence. It never terminates an Excel PID it
did not create; if a new PID remains, it fails without cleanup. The temporary
artifact is removed after validation. This gate is intentionally not called by
the GitHub-hosted release workflow.

## Consumer verification

Download the source ZIP or XLSM and verify the adjacent `SHA256SUMS.txt` before
installation. The source package remains the primary VBA-Web-style module
distribution; the XLSM is an optional production-only convenience artifact and
must not be described as VBE-verified.
