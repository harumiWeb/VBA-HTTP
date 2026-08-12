## Unreleased

### Added

- Defined the initial synchronous HTTP core API, dual-transport boundary, buffered body ownership, and stable error model.
- Added the default late-bound WinHTTP COM transport contract, including redirect controls and stable transport-error mapping.
- Added `VBAHttp.CreateClient` for consumers that reference the release workbook and an external release-artifact GET smoke test.
- Added matched Raw WinHttpRequest/VBA-HTTP Phase 2 benchmark evidence and optimized defensive Byte-array copies without weakening ownership.

### Fixed

- Made exported class sources clean-importable during `xlflow build` and added a preflight check for the required CRLF representation.
