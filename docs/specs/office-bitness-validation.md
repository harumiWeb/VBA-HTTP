# Office bitness validation

The declarations in `WinHttpNativeApi.bas` retain both VBA7 pointer-sized
handles and legacy 32-bit `Long` branches for source-level ABI protection. The
supported runtime target is Windows x64 Office only. A conditional branch or
an x64 compile cannot prove that 32-bit Office loads every declaration and
workbook component, so 32-bit Office is `unsupported-by-policy` under
ADR-0030.

Run the supported runner on the x64 Office host:

```powershell
task test:office-bitness
```

`tools/Run-OfficeBitnessValidation.ps1` checks synchronized source/workbook
state, records the xlflow bridge architecture, runs the isolated VBA suite,
builds a temporary production workbook with VBE compile/save/close evidence,
executes the deterministic loopback integration suite. Artifact consumer smoke
is owned by the separate release harness, so this runner records
`consumer_smoke=deferred-to-release-harness`. It publishes a path-stable,
loopback-free JSON result under `benchmarks/results/` and removes the temporary
artifact. Before reading Office metadata, the runner snapshots existing Excel
PIDs and proves that its COM instance created a new PID. It calls `Quit` only
for that owned instance; if ownership cannot be proven, validation fails closed
without touching a pre-existing user Excel process. A normal promotion run
rejects an X86 bridge before opening a new Excel instance:

```powershell
powershell -File tools/Run-OfficeBitnessValidation.ps1 -ExpectedArchitecture X64
```

The current repository evidence is X64 only and is the supported release path.
An explicitly requested exploratory run on a future x86 host must use
`-DiagnosticOnly`; its result uses `status=diagnostic` and
`support_status=unsupported-by-policy`, and cannot promote a compatibility or
release row:

```powershell
powershell -File tools/Run-OfficeBitnessValidation.ps1 `
  -ExpectedArchitecture X86 -DiagnosticOnly
```

## Source-level ABI guard

The host-specific runner is complemented by the Excel-free
`task test:native-declarations` gate. It checks that every WinHTTP/Kernel32
declaration has a `VBA7` `PtrSafe`/`LongPtr` branch and a legacy `Long` branch,
and that the upload DWORD sentinel remains
`WINHTTP_IGNORE_REQUEST_TOTAL_LENGTH = 0`. The gate catches accidental
pointer-size regressions but does not promote or support the x86 Office row.
Changing this boundary requires a superseding ADR and a new real-host evidence
bundle; the diagnostic switch is not a release gate.
