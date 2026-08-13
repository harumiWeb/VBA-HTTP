# Specification documents

`docs/specs/` は現在有効なcontract、invariant、compatibility rule、validation requirementだけを保持する。判断理由と却下案はADRへ、作業予定は `tasks/todo.md` へ置く。

- 仕様変更と実装・testは同じchange setで更新する。
- 過去仕様を本文へ蓄積せず、user-visibleな変更履歴はCHANGELOGへ移す。
- public API、error behavior、security boundary、release validationにはauthoritativeなtestまたは検査scriptへの参照を含める。

## Current specifications

- `development-and-release-workflow.md`: source authority、proof loop、release component boundary
- `local-test-server.md`: deterministic integration server contract
- `benchmark-methodology.md`: benchmark conditions、provenance、result evidence
- `resource-stress.md`: 10,000-request scenarios, PID-scoped sampling, idle handle gates, and result evidence
- `cancellation-stress.md`: repeated COM/native cancellation and timeout workloads, cleanup invariants, and PID-scoped evidence
- `http-core-api.md`: Phase 1 public domain、transport、body、decoding、error contract
- `http-core-api.md` also defines the shared header name/value control-character contract enforced before either transport.
- `buffered-com-transport.md`: default WinHTTP COM request、response、redirect、failure contract
- `buffered-com-transport.md` also records best-effort COM `Abort` cleanup on failure (ADR-0022)
- `bounded-concurrency.md`: cooperative COM async scheduling、batch results、deadlines、cancellation、reentrancy contract
- `reliability-policy.md`: retry classification、backoff／jitter、Retry-After、total deadline、cancellation precedence
- `streaming-download.md`: native constant-memory download、atomic publication、progress、cancellation、cleanup contract
- `streaming-upload.md`: native constant-memory file／multipart upload、source ownership、progress、cancellation、challenge contract
- `protocol-policy.md`: native HTTP/2 opt-in、diagnostic-only HTTP/3 flag、fallback／required behavior、transport boundary、compatibility evidence
- `decompression-policy.md`: native gzip／deflate response decoding, header ownership, fallback, and streaming length contract
- `proxy-policy.md`: OS/default, direct, and manual proxy routing, validation, and transport mapping
- `https-connect-proxy.md`: deterministic loopback CONNECT tunnel, proxy challenge, TLS rejection boundary, and host-dependent limits
- `auth-policy.md`: preemptive Basic/Bearer providers, TLS/redirect boundary, and redaction
- `challenge-authentication.md`: bounded buffered Windows/Digest/proxy challenge replay and streaming no-replay boundary
- `redirect-policy.md`: bounded automatic redirects, credential-header suppression, downgrade protection, and caller-owned follow-up
- `cookie-policy.md`: explicit caller-owned cookie jar parsing, matching, expiry, redaction, and redirect boundary
- `diagnostics-policy.md`: opt-in structured operation events, bounded retention, target sanitization, and header redaction
- `release-checksum.md`: deterministic SHA-256 sidecar generation, validation, and atomic evidence publication
- `release-security.md`: fail-closed release manifest identity, component path boundary, checksum, and deferred-risk contract
- `security-risk-register.md`: machine-readable residual risks and zero-current-blocker release gate
- `compatibility-matrix.md`: observed Windows/Office/protocol evidence and promotion requirements
- `distribution.md`: development/release workbook separation, build, checksum, install, and upgrade contract
- `source-package.md`: VBA-Web-style manifest, module package, installer, backup, and upgrade contract
- `licensing.md`: MIT license scope, package files, third-party notices, and release verification
- `powershell-quality.md`: PSScriptAnalyzer settings, tools scope, warning gate, and setup
- `office-bitness-validation.md`: supported x64 compile, isolated test, consumer-smoke evidence, diagnostic X86 rejection, and source-level ABI guard
- `protocol-host-validation.md`: fail-closed HTTPS HTTP/2 negotiated-host evidence, diagnostic HTTP/3 probe boundary, Office/WinHTTP metadata, and redacted schema
- `distribution.md` also defines the independent same-extension XLAM base, build target, add-in identity, and smoke gate (ADR-0021)
