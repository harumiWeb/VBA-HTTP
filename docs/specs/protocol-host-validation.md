# Protocol host validation specification

## Purpose and boundary

The normal test server is deliberately HTTP/1.1 and must not use external
network access.  `tools/Run-ProtocolHostValidation.ps1` is a separate,
operator-invoked evidence runner for a trusted TLS endpoint that is known to
support HTTP/2 or HTTP/3 on the selected Windows/WinHTTP host.

The runner requests one protocol with `HttpProtocolRequired`.  It succeeds
only when `HttpResponse.ProtocolUsed` is exactly the requested value.  A
fallback, unsupported option, TLS error, HTTP error status, or mismatch is a
failed host run, never a passing compatibility record.

## Invocation

Build and validate a release artifact first, then run one protocol per host:

```powershell
$env:VBA_HTTP_PROTOCOL_HOST_URL = 'https://trusted.example.test/probe'
$env:VBA_HTTP_PROTOCOL_EXPECTED = 'HTTP/2'
task protocol:host

# Or invoke the runner directly:
powershell -File tools/Run-ProtocolHostValidation.ps1 `
  -Url $env:VBA_HTTP_PROTOCOL_HOST_URL `
  -ExpectedProtocol $env:VBA_HTTP_PROTOCOL_EXPECTED
```

`ExpectedProtocol` is exactly `HTTP/2` or `HTTP/3`.  The URL must use HTTPS,
must not contain user-info, and is never written in full to the evidence
file.  The runner requires a compiled release artifact, its manifest, and its
checksum sidecar; it does not alter the development workbook.

The external consumer uses `VBAHttp.CreateRequest()` and sets all four
`HttpTimeouts` phases to 15 seconds before executing the required native
request. A host that cannot complete TLS or protocol negotiation therefore
fails in a bounded interval instead of inheriting the normal 300-second
receive default.

## Evidence schema

The runner starts a hidden ownership watchdog (90 seconds by default; the
MaxRuntimeSeconds parameter can be raised for a slower trusted endpoint). This
outer deadline is necessary because WinHTTP QUIC capability waits may not
honor ordinary phase timeouts. Before the watchdog starts, the runner maps the
new `Excel.Application.Hwnd` to its owning PID and requires that PID to be
absent from the pre-run snapshot. The watchdog can terminate only that exact
proven PID; a timed-out run publishes no evidence. Normal teardown gives the
same PID a five-second graceful-exit window before any force-stop is allowed.

The default output is
`benchmarks/results/protocol-host-http2.json` or
`benchmarks/results/protocol-host-http3.json`.

Required fields:

- `schema_version`: `1`;
- `benchmark`: `protocol-host-validation`;
- `status`: `passed`;
- `run_utc`: an ISO-8601 UTC timestamp;
- `source_revision`: a full hexadecimal Git revision;
- `requested_protocol` and `observed_protocol`: equal `HTTP/2` or `HTTP/3`;
- `mode`: `required`;
- `external_network`: `true`;
- `target`: `scheme=https`, non-empty `host`, and numeric `port`;
- `bridge`: xlflow bridge name/version/runtime and `X64` architecture;
- `office`: Excel version, build, and operating-system strings;
- `environment`: Windows version plus WinHTTP option IDs 133 (enable), 134
  (used), and 145 (required), with the observed protocol;
- `artifact`: basename plus 64-character SHA-256 values for artifact and
  manifest;
- `build`: VBE compile, source application, save, close, and cleanup evidence.

The validator rejects credentials, URL/path/query/body/header fields, secret
names, unknown protocol values, mismatched requested/observed values, missing
hashes, and non-passing build evidence.  It also rejects a top-level or nested
field whose name could contain a secret or request payload.

## Promotion rules

- Archive the JSON beside the exact release artifact, manifest, and checksum.
- Run only on the supported x64 Office target; the runner rejects an X86 bridge
  before opening a new Excel instance. The Office version/build and
  WinHTTP/Windows capability are still recorded by the runner.
- A passing HTTP/2 result promotes only the selected x64 host HTTP/2 row.
- A passing HTTP/3 result requires a QUIC-capable host and promotes only the
  selected x64 host HTTP/3 row.
- The current x64 HTTP/2 record is
  `benchmarks/results/protocol-host-http2.json` (trusted `nghttp2.org:443`,
  required mode, exact `ProtocolUsed=HTTP/2`). It does not promote HTTP/3.
- A 2026-08-13 attempt against `nghttp2.org:4433` and `quic.rocks:4433`
  ended in bounded RPC failures without an observed protocol; the attempt is
  recorded in
  `docs/verification/protocol-host-http3-attempt-2026-08-13.md` and does not
  create a passing HTTP/3 record.
- No host result changes the offline integration contract or permits a
  certificate-ignore option.

## Verification

- The ownership/non-Excel safety test is
  tools/Test-ProtocolHostWatchdog.ps1.

- Offline schema and fail-closed tamper cases:
  `powershell -File tools/Test-ProtocolHostEvidence.ps1`.
- Normal deterministic proof remains `task test:integration`; it does not call
  this runner or access external network.
- ADR rationale and process-ownership boundary: ADR-0025.
