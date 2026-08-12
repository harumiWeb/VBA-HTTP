# OS and Office compatibility matrix

This matrix records observed evidence, not assumptions from conditional VBA
compilation. A row is only promoted to `supported` after the listed host runs
the focused compile, integration, and release-consumer smoke checks.

| Environment or capability | Current status | Evidence required before promotion |
| --- | --- | --- |
| Windows x64 Office, WinHTTP default HTTP/1.1 | `verified` | Native loopback integration and release consumer smoke on the current host |
| Windows x64 Office, native protocol option unavailable | `fallback-tested` | Plain HTTP fallback test; response must report `HTTP/1.1` |
| Windows x64 Office, required protocol on plain HTTP | `rejection-tested` | Required-mode integration must raise `HttpErrorProtocol` before a response |
| HTTPS HTTP/2 negotiation | `pending` | TLS fixture, Windows/WinHTTP version, Office bitness, requested mask, and `ProtocolUsed` |
| HTTPS HTTP/3 negotiation | `pending` | QUIC-capable Windows/WinHTTP host, fixture, requested mask, and `ProtocolUsed` |
| Native WinHTTP gzip/deflate decompression | `verified` | x64 loopback buffered/download tests and release consumer smoke; option 118 requires Windows 8.1+ |
| WinHTTP default/direct/manual HTTP proxy routing | `integration-tested` | x64 COM/native loopback proxy tests; WinHTTP access-type mapping and release smoke; HTTPS CONNECT/auth remains pending |
| Preemptive Basic/Bearer auth providers | `integration-tested` | x64 COM/native loopback success, HTTPS-by-default validation, redirect suppression, and release consumer smoke; Windows integrated/Digest auth remains pending |
| Windows 32-bit Office, native declarations | `pending` | Real 32-bit compile, loopback integration, and release build evidence |
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
prove negotiated HTTP/2 or HTTP/3 support.
