# OS and Office compatibility matrix

This matrix records observed evidence, not assumptions from conditional VBA
compilation. A row is only promoted to `supported` after the listed host runs
the focused compile, integration, and release-consumer smoke checks.

| Environment or capability | Current status | Evidence required before promotion |
| --- | --- | --- |
| Windows x64 Office, WinHTTP default HTTP/1.1 | `verified` | Native loopback integration and release consumer smoke on the current host |
| Windows x64 Office, native protocol option unavailable | `fallback-tested` | Plain HTTP fallback test; response must report `HTTP/1.1` |
| Windows x64 Office, required protocol on plain HTTP | `rejection-tested` | Required-mode integration must raise `HttpErrorProtocol` before a response |
| Windows x64 Office, untrusted HTTPS certificate | `verified-rejection` | Loopback self-signed fixture must map COM/native failures to `HttpErrorTls` without a certificate-ignore option |
| HTTPS HTTP/2 negotiation | `verified-host-specific` | x64 record at `benchmarks/results/protocol-host-http2.json` against `nghttp2.org:443`; required HTTP/2, Office 16.0 build 17932, Windows x64, and exact `ProtocolUsed=HTTP/2` |
| HTTPS HTTP/3 negotiation | `pending` | Same fail-closed host runner on a QUIC-capable Windows/WinHTTP endpoint; requested mask and exact `ProtocolUsed` |
| Native WinHTTP gzip/deflate decompression | `verified` | x64 loopback buffered/download tests and release consumer smoke; option 118 requires Windows 8.1+ |
| WinHTTP default/direct/manual HTTP proxy routing | `integration-tested` | x64 COM/native loopback forwarding and HTTPS CONNECT-boundary tests; WinHTTP access-type mapping and release smoke; trusted corporate CONNECT remains host-dependent |
| Preemptive Basic/Bearer auth providers | `integration-tested` | x64 COM/native loopback success, HTTPS-by-default validation, redirect suppression, and release consumer smoke |
| Buffered Windows/Digest/server-or-proxy challenge provider | `integration-tested` | x64 COM/native bounded server Basic, HTTP proxy Basic, and loopback HTTPS CONNECT Basic challenge; COM uses `Auto`, native honors explicit scheme selection; release smoke, wrong-credential 407, and upload pre-network rejection; real Windows-domain and trusted corporate CONNECT fixtures remain pending |
| Windows 32-bit Office, native declarations | `unsupported-by-policy` | No promotion evidence is required. A future support claim needs a superseding ADR and a complete real-host compile, loopback integration, release build, and consumer-smoke bundle |
| COM transport with advanced protocol options | `rejected-by-contract` | Public contract test must reject non-empty options without network I/O |

`HttpProtocolOptions.EnabledProtocols = 0` means no native override. Modern
protocol flags are attempted only for HTTPS URLs; a requested flag is never
treated as evidence that a peer negotiated that protocol. The authoritative
runtime value is `HttpResponse.ProtocolUsed`.

## Evidence record format

Each future host run should record:

- Windows edition/build and WinHTTP capability;
- Office version and 32/64-bit architecture;
- workbook/source revision and xlflow version;
- endpoint certificate and negotiated protocol support;
- requested flags/mode, observed error category, and `ProtocolUsed`;
- focused test, full suite, and release-consumer smoke results.

The current loopback server is intentionally plain HTTP/1.1, so it cannot
prove negotiated HTTP/2 or HTTP/3 support. The recorded HTTP/2 result is
host-specific and does not promote HTTP/3. The x64 row is the only supported
Office runtime row; the legacy declaration branch is a static ABI guard, not
runtime support for 32-bit Office.

The host-evidence runner and schema are defined in
`docs/specs/protocol-host-validation.md`. A passing record is required before
either pending protocol row is promoted.
