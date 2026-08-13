# ADR-0032: Excel-free protocol capability preflight

## Status

`accepted`

## Background

The protocol-host runner is intentionally an external, host-specific HTTP/2
proof:
it must use the public release consumer and record the negotiated
`ProtocolUsed` value. The current Windows host can reject an HTTP/3-enabled
WinHTTP option before a request is sent. Repeated attempts also showed that a
WinHTTP/Excel automation boundary can terminate Excel before VBA receives a
recoverable error. Starting Excel before checking that capability wastes the
host-owned process and makes a concurrent user Excel session harder to reason
about, even though the ownership guard never targets a pre-existing PID.

The direct x64 `tools/Probe-WinHttpProtocol.ps1` already exercises the same
WinHTTP option and protocol query calls without Excel. Its result is not
promotion evidence, but it can safely distinguish an unavailable protocol
from an automation failure.

## Decision

- After release-artifact, checksum, security, doctor, and x64-bridge checks,
  `tools/Run-ProtocolHostValidation.ps1` runs the direct probe before creating
  `Excel.Application`.
- The preflight requests exactly the required HTTP/2 mask, uses a bounded
  timeout, and requires the probe's negotiated flag to equal that mask. Any
  unsupported option, timeout, malformed output, or mismatch fails closed
  before Excel starts and publishes no evidence. HTTP/3 probing remains
  diagnostic-only under ADR-0035.
- A passing preflight is retained only as redacted metadata in
  `environment.winhttp.preflight`. The release workbook's public consumer
  still runs afterward and must independently return the exact expected
  `HttpResponse.ProtocolUsed` value before a promotion record is written.
- The preflight is an opt-in host-runner concern. It is not used by the
  offline integration suite, release smoke, or production request execution.
  It never changes the fallback/required contract of the native transport.
- The runner remains x64-only under ADR-0030. The preflight's direct x64
  diagnostic path is not 32-bit Office evidence and cannot promote x86.

## Consequences

- Known WinHTTP capability misses no longer start Excel or risk an automation
  crash before the runner can report the reason.
- A host that passes the direct probe still needs the more expensive public
  consumer proof, so the preflight cannot create a false compatibility claim.
- A transient network or probe-specific failure may stop a host run earlier;
  the operator can retry on the same trusted endpoint without cleaning any
  user-owned Excel process.
- Passing evidence now carries a small, schema-validated preflight object;
  URL, headers, body, credentials, and raw error text are deliberately not
  stored.

## Rationale / Evidence

- Probe: `tools/Probe-WinHttpProtocol.ps1`.
- Runner: `tools/Run-ProtocolHostValidation.ps1`.
- Schema and tamper tests: `tools/Validate-ProtocolHostEvidence.ps1` and
  `tools/Test-ProtocolHostEvidence.ps1`.
- Ownership safety: `tools/OfficeProcessOwnership.ps1`,
  `tools/Watch-ProtocolHostExcel.ps1`, and `tasks/lessons.md`.
- Host failure evidence: `docs/verification/protocol-host-http3-attempt-2026-08-13.md`.

## Supersedes

- None

## Superseded by

- None
