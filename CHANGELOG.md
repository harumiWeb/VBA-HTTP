## Unreleased

- Added a read-only Excel-free GitHub CI workflow for pull requests, `main`
  pushes, weekly maintenance runs, and manual validation. It installs the
  locked toolchain, runs source/release-plan contracts, and uses a dedicated
  Excel-free clean-checkout path; XLAM identity and other Excel/VBE checks
  remain on the local/Excel-host proof boundary (ADR-0041).

- Reclassified 32-bit Office from `unsupported-by-policy` to `unverified` under
  ADR-0039. The x64 runtime remains the official release target; contributors
  can submit real-host `-DiagnosticOnly` evidence for maintainer review, but
  diagnostic output alone does not create an official support guarantee.

- Consolidated the consumer-facing API reference, usage recipes, reliability,
  streaming, security, transport, compatibility, and distribution guidance
  under `docs/guides/`; the README and legacy `docs/API.md` now act as concise
  public entry points while `docs/specs/` remains the normative contract layer.

- Optimized the native buffered and streaming read paths to reuse a fixed
  64 KiB buffer, read directly with `WinHttpReadData`, pre-size known-length
  buffered bodies (with a bounded initial allocation), and avoid VBA per-byte
  copies for short download chunks.
- Reused full-size file-upload buffers and added a non-aligned 1,048,577-byte
  native upload regression case. Existing 1 GiB benchmark numbers remain the
  pre-optimization baseline until a PID-scoped x64 before/after run is saved.
- Hardened the Raw/VBA-HTTP benchmark runner to refuse pre-existing Excel
  processes and never terminate Excel during cleanup; current x64 measurements
  remain engineering evidence until repeated before/after runs are recorded.

- Added strict SemVer tag-triggered GitHub Release automation. A hosted
  Windows x64 workflow validates source/toolchain/pack manifests without
  starting Excel, stages only production components for `xlflow pack
  --experimental`, publishes the exact source/XLSM/manifest/checksum/license
  asset set with `gh release create --verify-tag`, and records
  `vbe_validation=not_performed`. Local Lefthook now runs `task precommit`,
  including the ownership-safe temporary x64 VBE compile gate.

- Changed the primary consumer distribution to a VBA-Web-style manifest-based
  source package with one-command installer/uninstaller, backup, and hash
  validation. The empty `App`, `Main`, and `Ui` scaffold entry modules were
  removed; compiled XLSM/XLAM artifacts remain optional release targets.

- Added a repository PSScriptAnalyzer settings contract and
  `task powershell:lint` gate for all `tools/*.ps1`; warnings and errors now
  stop the Lefthook pre-commit check.

- Established the MIT license for project-authored source, documentation,
  tooling, examples, and generated release workbooks. Added the canonical
  `LICENSE`, third-party notice inventory, license ADR/spec, and a pre-commit
  license contract; pinned VBA-Web remains a benchmark-only upstream MIT
  component.

- Established ADR-0035: HTTP/3/QUIC is unsupported by policy for the current
  distribution. The native HTTP/3 option and Excel-free probe remain
  diagnostic-only; protocol-host promotion now accepts HTTP/2 evidence only,
  and HTTP/3 no longer blocks v1 release promotion.

- Made the buffered challenge-authentication backend boundary explicit: COM
  now accepts only `HttpAuthSchemeAuto` because
  `IWinHttpRequest::SetCredentials` has no scheme argument; explicit
  Basic/Digest/NTLM/Negotiate selection fails before COM backend creation and
  remains available through the native transport (ADR-0033).

- Made challenge limits fail closed across backends: native WinHTTP enforces
  `MaxChallenges`, while COM accepts only the default and rejects custom limits
  before creating a backend because its API does not expose exchange-count
  callbacks (ADR-0034).

- Added an Excel-free WinHTTP capability preflight to the opt-in HTTP/2
  protocol-host runner. It checks the exact required mask before creating
  Excel, records only redacted preflight metadata, and fails closed without
  publishing evidence when the host lacks the capability (ADR-0032).

- Treated Win32 `ERROR_NOT_SUPPORTED` (50) as a native protocol/decompression
  capability miss, so HTTP/3 requests on WinHTTP runtimes without HTTP/3 now
  fail or fall back through the documented `HttpErrorProtocol` contract rather
  than leaking an unmapped native I/O error.

- Passed native `WinHttpSetOption` DWORD buffers as explicit pointer-sized
  addresses, matching the WinHTTP `LPVOID` ABI and avoiding VBA `Any`
  marshalling at unsupported protocol boundaries.

- Fixed native `WinHttpSetOption` DWORD marshalling so protocol, redirect,
  decompression, and upload options are passed through an explicit 32-bit
  buffer in both VBA7 and legacy declaration branches.

- Extended the filtered release consumer smoke to exercise ordinary and
  authenticated HTTPS CONNECT boundaries through both COM and native public
  factories. Release validation now proves ownership of every Excel PID it
  may clean up, leaving unrelated user Excel processes untouched.

- Centralized HWND-to-PID Excel ownership guards for release, Office bitness,
  and protocol-host validation runners. Concurrent user Excel sessions remain
  outside cleanup scope, with a five-second graceful-exit window before an
  exact owned-PID force cleanup.

- Added an offline HTTPS CONNECT proxy boundary fixture (including fixed Basic
  proxy challenge) with COM/native integration coverage. The fixture proves
  tunnel reachability before normal untrusted-certificate rejection; trusted
  corporate CONNECT remains an explicit host-dependent gate.

- Hardened `HttpHeaders` value validation to reject NUL, DEL, and C0 control
  characters (while preserving HTAB and Unicode), preventing control-byte
  injection before COM/native transport serialization.
- Corrected signed `AscW` handling so non-ASCII header names cannot bypass the
  ASCII token validation boundary.
- Corrected signed `AscW` handling in Bearer token validation so non-ASCII
  code points cannot bypass its ASCII-only contract.
- Hardened cookie-name parsing against signed `AscW` values so non-ASCII names
  are ignored instead of entering the jar.
- Hardened Office bitness validation to prove ownership of a new Excel PID
  before quitting its metadata COM instance.
- Fixed query-parameter merging so parameters are inserted before a URL
  fragment instead of being appended after it.
- Normalized execution request-targets to remove client-only URI fragments
  consistently across the client, COM transport, and native transport
  (ADR-0029).
- Rejected cookie `Domain` attributes that broaden a cookie to a single-label
  public-suffix-like scope, while preserving exact `localhost` fixtures.

- Superseded the earlier ADR-0030 32-bit Office exclusion with ADR-0039:
  Windows x64 Office remains the official runtime target, while 32-bit Office
  is `unverified` until a contributor evidence bundle is reviewed. Normal
  promotion runners continue to reject X86 before opening a new Excel instance;
  the explicit diagnostic-only path remains available.

- Added an Excel-free native declaration ABI gate that verifies the VBA7
  `PtrSafe`/`LongPtr` and legacy `Long` WinHTTP branches, including the
  upload DWORD sentinel. This guards source regressions without claiming
  unsupported 32-bit Office runtime compatibility.

- Added `VBAHttp.CreateRequest()` so workbook-reference consumers can
  configure request-level timeouts and transport options without constructing
  `PublicNotCreatable` classes directly. The protocol-host consumer probe now
  uses a bounded 15-second per-phase timeout.

- Added best-effort `Abort` cleanup plus a bounded 250 ms teardown drain for
  synchronous failures and asynchronous cancellation before releasing
  `WinHttpRequest` objects. Original timeout/error mapping is preserved;
  canonical x64 stress evidence now covers repeated receive-timeout cleanup.

- Extended the cancellation stress gate with repeated synchronous COM
  receive-timeout cycles and a dedicated idle-handle budget.

- Added a machine-validated Phase 9 security-risk register with an explicit
  zero-current-blocker gate, fail-closed tamper tests, and release-security
  integration. Host-specific integrated/proxy challenge fixtures remain a
  future-v1 obligation; HTTP/3/QUIC and 32-bit Office are explicitly outside
  the supported boundary;
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
  the supported evidence path is X64 and normal promotion rejects an X86
  bridge before opening a new Excel instance. Explicit diagnostic-only X86
  runs are labeled unsupported and cannot promote a release.

- Added an opt-in fail-closed protocol-host evidence runner for required
  HTTPS HTTP/2 negotiation. It records only path-free target metadata,
  artifact hashes, Office/Windows/WinHTTP metadata, bridge/bitness, and exact
  `ProtocolUsed`; the normal loopback suite remains offline. The current x64
  HTTP/2 host record is archived in
  `benchmarks/results/protocol-host-http2.json`; HTTP/3/QUIC is unsupported by
  policy and the runner is x64-only.

- Added opt-in structured `HttpDiagnostics` operation events with bounded
  retention, query/user-info target sanitization, stable error categories, and
  automatic sensitive-header redaction. Release consumer smoke verifies that
  authorization, cookie, and query sentinels never appear in the JSON output.

- Added a fail-closed release manifest security gate and deterministic report:
  canonical base/output paths, VBE/atomic-publication evidence, exact
  included/excluded component policy, development-only source boundaries, and
  checksum verification are checked before release smoke. Host-specific
  integrated/proxy challenge fixtures remain explicitly deferred; HTTP/3/QUIC
  and 32-bit Office are outside the supported boundary, and x64 HTTP/2
  evidence is archived separately.

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
  credentials, and trusted corporate proxy CONNECT remain deferred. The
  loopback CONNECT boundary is covered separately by ADR-0031.

- Added shared default, direct, and manual HTTP proxy routing via
  `HttpProxyOptions` for COM and native transports, with deterministic loopback
  forwarding tests. Added a separate fixed Basic proxy-challenge listener and
  COM/native/release smoke for `HttpAuthTargetProxy`; the loopback HTTPS
  CONNECT boundary is covered by ADR-0031, while trusted CONNECT and real
  Windows-domain authentication remain deferred.

### Added

- Added deterministic SHA-256 release checksum sidecars for the filtered
  workbook and xlflow build manifest, with atomic sidecar publication,
  tamper-detecting validation, and a release smoke gate.
- Added native `HttpDecompressionOptions` for bounded WinHTTP gzip/deflate
  response decoding, deterministic fallback/required behavior, COM validation,
  loopback integration coverage, and release-consumer decompression smoke.
- Added native `HttpProtocolOptions` with HTTP/2 opt-in, a retained
  diagnostic-only HTTP/3 flag, explicit fallback or required mode,
  negotiated-protocol reporting, COM-backend validation, and release-consumer
  protocol smoke coverage.
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
- Hardened URL validation so authority user-info is rejected consistently by
  `HttpClient`, COM/native transport boundaries, and the explicit cookie jar.
