# PowerShell quality specification

## Scope and gate

All tracked PowerShell under `tools/` is analyzed by PSScriptAnalyzer using
the repository-root `PSScriptAnalyzerSettings.psd1`. `task powershell:lint`
runs `tools/Invoke-PSScriptAnalyzer.ps1`; the task is part of `task check` and
therefore the Lefthook pre-commit gate.

The gate fails on every PSScriptAnalyzer `Error` or `Warning`. No project-wide
rule suppression is permitted. A rule that is intentionally inapplicable must
be resolved by changing the code or by a narrow, documented source-level
exception with a linked rationale; `ExcludeRules` remains empty by default.

The settings intentionally use PSScriptAnalyzer default rules and do not claim
cross-platform PowerShell compatibility. These scripts run on the supported
Windows/x64 Office development and release host, and Excel ownership/cleanup
code must continue to protect pre-existing user Excel processes.

## Local setup

Install PSScriptAnalyzer for the current user or machine, then run:

```powershell
Install-Module PSScriptAnalyzer -Scope CurrentUser
task powershell:lint
```

The repository does not vendor the analyzer module. The module version and
PowerShell host are recorded by the operator when producing release evidence;
the source settings and clean result remain the repository contract.
