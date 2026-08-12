## Unreleased

- Added a fail-closed release manifest security gate and deterministic report:
  canonical base/output paths, VBE/atomic-publication evidence, exact
  included/excluded component policy, development-only source boundaries, and
  checksum verification are checked before release smoke. Certificate-negative
  fixtures, HTTP/2/HTTP/3 TLS evidence, integrated auth, and cookie persistence
  remain explicitly deferred.

- Added a Phase 9 cancellation/timeout stress gate with repeated COM active
  cancellation and request-deadline scenarios, native streaming download
  cancellation, recovery requests, and
  PID-scoped idle handle evidence. The gate is loopback-only and excluded from
  release workbooks. The existing single COM receive-timeout integration test
  remains separate because repeated COM aborts currently expose handle-growth
  behavior requiring a follow-up fix.

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
  Cookie persistence remains caller-owned and deferred to a separate policy.

- Added snapshot-safe preemptive Basic/Bearer authentication providers for COM
  and native requests, deterministic loopback auth endpoints, redirect
  suppression, secret-header redaction helpers, and release consumer auth smoke.
  Windows integrated/Digest authentication, proxy credentials, and challenge
  replay remain explicitly deferred.

- Added shared default, direct, and manual HTTP proxy routing via
  `HttpProxyOptions` for COM and native transports, with deterministic loopback
  forwarding tests. Proxy credentials and CONNECT authentication remain
  deferred to the authentication policy.

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
