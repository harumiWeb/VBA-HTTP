# ADR-0025: Host-specific negotiated protocol evidence boundary

## Status

`accepted`

## Context

The deterministic loopback server intentionally speaks HTTP/1.1.  Its
fallback tests prove option validation and the `ProtocolUsed` contract, but
they cannot prove that a particular Windows/WinHTTP host negotiated HTTP/2 or
HTTP/3 over TLS.  Treating a requested flag as negotiated evidence would make
the compatibility matrix overstate support.  The remaining promotion gate
also needs to be repeatable on both x64 and x86 Office hosts without putting
external network access into the normal test suite.

## Decision

- The external consumer must use the public request factory and finite
  per-phase timeouts. A hidden watchdog monitors only the runner-owned Excel
  PID(s) and terminates them after the configured outer deadline; this is
  required because WinHTTP QUIC capability waits may not honor ordinary phase
  timeouts. A timed-out run publishes no evidence.

- Add an opt-in external consumer runner that targets one caller-supplied
  `https://` origin and requests exactly one protocol in `Required` mode.
- The external consumer harness creates its request through
  `VBAHttp.CreateRequest()` and applies finite per-phase timeouts before
  invoking `HttpClient.Execute`; the proof must fail in bounded time when a
  host or endpoint cannot complete the required negotiation.
- Record success only when the response's `HttpResponse.ProtocolUsed` exactly
  matches the requested protocol.  Unsupported capability, TLS failure,
  status failure, or protocol mismatch fails closed and publishes no passing
  evidence.
- Store only path-free target metadata (scheme, host, and port), artifact and
  manifest SHA-256 values, source revision, Office/bridge/Windows/WinHTTP
  metadata, requested protocol, observed protocol, and UTC time. Reject
  credentials, query/body values, headers, and local filesystem paths in the
  evidence schema.
- Keep this runner out of `task verify` and ordinary deterministic integration
  tests.  A host owner explicitly invokes it after selecting a trusted TLS
  endpoint and archives the result with the release artifact.
- Validate x64 and x86 evidence independently; an x64 result never promotes
  the x86 matrix row, and an HTTP/2 result never promotes HTTP/3.

## Consequences

- The compatibility matrix can distinguish fallback-tested behavior from real
  negotiated protocol evidence without relying on external URLs in CI.
- Host owners must provide a trusted TLS/QUIC endpoint and accept the external
  network boundary; the current loopback suite remains fully offline.
- Evidence generation opens a dedicated automation Excel instance only after
  proving ownership.  It never quits or force-stops a pre-existing Excel PID.
- HTTP/3 remains pending until a WinHTTP/QUIC-capable host produces a passing
  result; this ADR does not claim capability by itself.

## Evidence

- Consumer entrypoint: `tools/consumer/ReleaseBatchSmoke.bas`.
- Referenced-workbook request factory: `src/modules/VBAHttp.bas` and
  `src/modules/Tests/Unit/VBAHttpTests.bas`.
- Runner and schema validator: `tools/Run-ProtocolHostValidation.ps1` and
  `tools/Validate-ProtocolHostEvidence.ps1`.
- Offline validator tests: `tools/Test-ProtocolHostEvidence.ps1`.
- Contract: `docs/specs/protocol-host-validation.md` and
  `docs/specs/compatibility-matrix.md`.

## Supersedes

- None

## Superseded by

- None

The ownership watchdog is implemented by
tools/Watch-ProtocolHostExcel.ps1.
