## Unreleased

### Added

- Defined the initial synchronous HTTP core API, dual-transport boundary, buffered body ownership, and stable error model.
- Added the default late-bound WinHTTP COM transport contract, including redirect controls and stable transport-error mapping.

### Fixed

- Made exported class sources clean-importable during `xlflow build` and added a preflight check for the required CRLF representation.
