# GitHub Excel-free CI specification

## Scope and triggers

`.github/workflows/ci.yml` is the read-only GitHub Actions validation path for
changes that do not require Excel. It runs for pull requests targeting `main`,
pushes to `main`, a weekly scheduled run, and explicit `workflow_dispatch`.
The workflow uses a GitHub-hosted `windows-2022` runner and cancels older
in-progress runs for the same ref.

The workflow has `contents: read` permission only. It never publishes a
release, mutates a tag, or uses a self-hosted runner.

## Toolchain and gates

The workflow installs the toolchain versions recorded in
`tools/release-toolchain.json`: xlflow from the locked archive, Task from the
locked Go module version, Go `1.24.0`, and PSScriptAnalyzer from the locked
PowerShell module version. The setup must fail if any required tool cannot be
installed or probed.

Every run executes these Excel-free gates:

- `task check`
- `task test:docs`
- `task test:security-risks`
- `task testserver:unit`
- `task test:release-checksum`
- `task test:release-security`
- `task build:plan`
- `task build:plan:xlam`
- `task test:xlam`
- `task test:pack-release`
- `task test:github-release`

The workflow fails if any gate changes tracked or untracked repository files.
The `task test:clean-checkout` contract runs only for scheduled and manually
dispatched executions because it repeats the source gates in a fresh archive.

## Excel boundary

The hosted workflow never starts Excel, imports source into VBIDE, runs the
VBA test suite, or claims VBE compilation evidence. Local `task precommit`,
`task release:build`, and the documented real-host validation paths remain the
authoritative Excel/VBE proof boundaries. Tag publication remains the separate
contract in `github-release.md` and `.github/workflows/release.yml`.

`tools/Test-CIWorkflow.ps1` is the executable contract for the workflow shape,
permissions, toolchain pins, gate list, and Excel/release exclusions.
