# Office bitness validation

The declarations in `WinHttpNativeApi.bas` support both VBA7 pointer-sized
handles and legacy 32-bit `Long` handles, but an x64 compile cannot prove that
32-bit Office loads every declaration and workbook component. Bitness is
therefore an evidence property, not a source-only claim.

Run the same runner on each Office host:

```powershell
task test:office-bitness
```

`tools/Run-OfficeBitnessValidation.ps1` checks synchronized source/workbook
state, records the xlflow bridge architecture, runs the isolated VBA suite,
builds a temporary production workbook with VBE compile/save/close evidence,
executes the deterministic loopback integration suite, and invokes `Main.Run`
from that artifact. It publishes a path-stable,
loopback-free JSON result under `benchmarks/results/` and removes the temporary
artifact. On a 32-bit Office host use:

```powershell
powershell -File tools/Run-OfficeBitnessValidation.ps1 -ExpectedArchitecture X86
```

The current repository evidence is X64 only. The X86 result must be generated
on a real 32-bit Office host; changing the expected flag on an x64 host is
designed to fail rather than fabricate compatibility evidence.

## Source-level ABI guard

The host-specific runner is complemented by the Excel-free
`task test:native-declarations` gate. It checks that every WinHTTP/Kernel32
declaration has a `VBA7` `PtrSafe`/`LongPtr` branch and a legacy `Long` branch,
and that the upload DWORD sentinel remains
`WINHTTP_IGNORE_REQUEST_TOTAL_LENGTH = 0`. The gate catches accidental
pointer-size regressions but does not promote the real x86 Office row; only
`Run-OfficeBitnessValidation.ps1 -ExpectedArchitecture X86` on a 32-bit host
can provide that evidence.
