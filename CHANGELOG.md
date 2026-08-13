## Unreleased

- Added best-effort `Abort` cleanup plus a bounded 250 ms teardown drain for
  synchronous failures and asynchronous cancellation before releasing
  `WinHttpRequest` objects. Original timeout/error mapping is preserved;
  canonical x64 stress evidence now covers repeated receive-timeout cleanup.

- Extended the cancellation stress gate with repeated synchronous COM
  receive-timeout cycles and a dedicated idle-handle budget.

- Added a machine-validated Phase 9 security-risk register with an explicit
  zero-current-blocker gate, fail-closed tamper tests, and release-security
  integration. Deferred HTTP/2/HTTP/3 and 32-bit evidence plus host-specific
  integrated/proxy challenge fixtures remain visible as future-v1 obligations;
  repeated
  COM receive-timeout growth is mitigated by the canonical x64 stress result.

- Added the public API quick reference, contributor proof-loop guide, consumer
  examples, distribution/install/rollback specification, and repeatable XLSM
  release checklist. Documentation contracts now run as part of `task verify`.

- Added a real same-extension XLAM base and an independent `task
  release:xlam:build` target. Its dry-run, production component policy,
  checksum, add-in identity, VBE publication, and external consumer smoke are
  validated separately from the XLSM release.

- Added a clean-checkout contract that builds a temporary `git archive` and
  reruns the source, documentation, base-identity, and both release-plan gates;
  an opt-in full mode builds both XLSM and XLAM artifacts from that archive.

- Added a host-specific Office bitness validation runner and result schema. It
  records bridge architecture, isolated tests, loopback integration, VBE compile, and consumer smoke;
  the current evidence remains X64 until the same runner is executed on a real
  32-bit Office host.

- Added an opt-in fail-closed protocol-host evidence runner for required
  HTTPS HTTP/2/HTTP/3 negotiation. It records only path-free target metadata,
  artifact hashes, bridge/bitness, and exact `ProtocolUsed`; the normal
  loopback suite remains offline and the real HTTP/3 host gate remains pending.

- Added opt-in structured `HttpDiagnostics` operation events with bounded
  retention, query/user-info target sanitization, stable error categories, and
  automatic sensitive-header redaction. Release consumer smoke verifies that
  authorization, cookie, and query sentinels never appear in the JSON output.

- Added a fail-closed release manifest security gate and deterministic report:
  canonical base/output paths, VBE/atomic-publication evidence, exact
  included/excluded component policy, development-only source boundaries, and
  checksum verification are checked before release smoke. HTTP/2/HTTP/3 TLS
  negotiation evidence and host-specific integrated/proxy challenge fixtures
  remain explicitly deferred.

- Added a Phase 9 cancellation/timeout stress gate with repeated COM active
  cancellation and request-deadline scenarios, native streaming download
  cancellation, repeated COM receive-timeout cycles, recovery requests, and
  PID-scoped idle handle evidence. The gate is loopback-only and excluded from
  release workbooks. Existing Excel PIDs are baseline-only and are never
  terminated; the canonical current-host result is retained as release
  evidence.

- Added deterministic loopback response-boundary coverage for UTF-8 Unicode,
  exact binary GET bodies, and malformed UTF-8 rejection across the COM and
  native transports, plus raw malformed response-header protocol mapping.

- Added a Phase 9 resource-stress gate with 10,000-request warmup followed by
  10,000 measured sequential native and scheduled COM requests, PID-scoped
  Excel process sampling, idle handle-growth limits, atomic JSON evidence, and
  an excluded development-only stress module.

- Added a redirect security boundary: credential and cookie headers suppress
  automatic redirects, redirect loops honor the bounded maximum, and both COM
  and native transports explicitly reject HTTPS-to-HTTP downgrade redirects.
  Cookie persistence is now an explicit caller-owned `HttpCookieJar` policy.

- Added an explicit in-memory cookie jar with host/path/secure/expiry matching,
  caller-header precedence, response `Set-Cookie` storage, redirect suppression,
  deterministic loopback integration coverage, and release-consumer smoke.

- Added an optional deterministic HTTPS test-server listener with an untrusted
  self-signed certificate, plus COM/native `HttpErrorTls` rejection tests and
  compatibility/security evidence. The fixture is never installed or accepted
  as a trust anchor.

- Added snapshot-safe preemptive Basic/Bearer authentication providers for COM
  and native requests, deterministic loopback auth endpoints, redirect
  suppression, secret-header redaction helpers, and release consumer auth smoke.
  Bounded buffered server/proxy challenge authentication is now available via
  `HttpWindowsAuthProvider`; streaming upload replay, real Windows-domain
  credentials, and proxy CONNECT challenge fixtures remain deferred.

- Added shared default, direct, and manual HTTP proxy routing via
  `HttpProxyOptions` for COM and native transports, with deterministic loopback
  forwarding tests. Added a separate fixed Basic proxy-challenge listener and
  COM/native/release smoke for `HttpAuthTargetProxy`; HTTPS CONNECT and real
  Windows-domain authentication remain deferred.

### Added

- Added deterministic SHA-256 release checksum sidecars for the filtered
  workbook and xlflow build manifest, with atomic sidecar publication,
  tamper-detecting validation, and a release smoke gate.
- Added native `HttpDecompressionOptions` for bounded WinHTTP gzip/deflate
  response decoding, deterministic fallback/required behavior, COM validation,
  loopback integration coverage, and release-consumer decompression smoke.
- Added native `HttpProtocolOptions` with HTTP/2/HTTP/3 opt-in, explicit fallback or required mode, negotiated-protocol reporting, COM-backend validation, and release-consumer protocol smoke coverage.
- Added native constant-memory `HttpClient.UploadFile` and `UploadMultipart` APIs with deterministic content lengths, bounded `WinHttpWriteData` chunks, UTF-8 multipart fields, progress/cancellation checkpoints, source-preserving failure semantics, and authentication-challenge reporting.
- Added deterministic loopback upload endpoints, file/multipart integration coverage, a 1 GiB upload stress harness, and release-consumer upload smoke coverage.
- Added native constant-memory `HttpClient.DownloadFile` with bounded WinHTTP reads, same-directory temporary files, atomic publication, progress callbacks, cooperative cancellation/deadlines, and external release-artifact download smoke coverage.
- Defined the initial synchronous HTTP core API, dual-transport boundary, buffered body ownership, and stable error model.
- Added the default late-bound WinHTTP COM transport contract, including redirect controls and stable transport-error mapping.
- Added `VBAHttp.CreateClient` for consumers that reference the release workbook and an external release-artifact GET smoke test.
- Added matched Raw WinHttpRequest/VBA-HTTP Phase 2 benchmark evidence and optimized defensive Byte-array copies without weakening ownership.
- Added bounded `ExecuteMany` and `GetMany` execution with ordered per-item results, deadlines, cancellation, partial-failure isolation, and same-client reentrancy protection.
- Added deterministic concurrency integration tests, a server-observed in-flight bound, Phase 3 benchmark evidence, and an external release-artifact batch smoke test.
- Added client-owned retry policy and per-call execution options with idempotency-aware defaults, capped exponential backoff and jitter, `Retry-After`, total deadlines, and active WinHTTP cancellation.
- Extended bounded batches with per-item retry scheduling, retry waits that release concurrency slots, total batch deadlines, and cancellation of waiting retries.
- Added referenced-workbook factories for reliability configuration and external release retry/deadline smoke coverage.
- Added the synchronous native WinHTTP transport with pointer-sized session, connection, and request handle wrappers, protocol reporting, shared error classification, and repeated-request handle regression coverage.

### Fixed

- Corrected native upload `dwTotalLength` encoding across the unsigned DWORD range, including the `WINHTTP_IGNORE_REQUEST_TOTAL_LENGTH = 0` sentinel for payloads at or above 4 GiB.
- Made exported class sources clean-importable during `xlflow build` and added a preflight check for the required CRLF representation.
