# Security risk register contract

`docs/security/risk-register.json` is the current release risk register. It is
not a vulnerability database and contains no credentials, host paths, or
request data.

## Schema

- `schema_version` is `1`.
- `current_release_blockers` is an array of issue IDs and must be empty for the
  current release gate.
- `issues[].id` is unique; `status` is `open`, `known`, `deferred`, or
  `mitigated`; `severity` is `low`, `medium`, `high`, or `critical`.
- Each issue records `blocks_current_release`, `blocks_v1`, required evidence,
  and an evidence document.
- An issue listed in `current_release_blockers` must exist, be non-mitigated,
  and have `blocks_current_release=true`. The validator rejects contradictory
  entries and any critical/current blocker.

## Gate and evidence

`tools/Validate-SecurityRiskRegister.ps1` is run by `task verify` and by the
release security validator before Excel opens. It prints a deterministic count
summary and fails closed for missing, malformed, duplicate, or contradictory
records. The current register intentionally keeps HTTP/3/QUIC and 32-bit
evidence plus host-specific integrated/proxy challenge authentication as
future-v1 work without making them current release blockers. The bounded
loopback challenge contract is covered by ADR-0023; domain/CONNECT fixtures
remain deferred. The protocol-host runner has produced a passing x64 HTTP/2
record at benchmarks/results/protocol-host-http2.json; that record is
host-specific and does not satisfy the remaining HTTP/3 or x86 requirements.
Repeated COM receive-timeout handle growth is
mitigated on the current x64 host by the canonical cancellation stress result;
the same evidence must be rerun on each supported Office host.
