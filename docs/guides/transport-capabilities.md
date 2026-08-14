# Transport capabilities

## Backend matrix

| Capability | COM (`CreateClient`) | Native (`CreateNativeClient`) |
| --- | --- | --- |
| Buffered GET/POST/PUT/PATCH/DELETE | Supported | Supported |
| Bounded async batches | Supported | Client uses the native buffered boundary where applicable |
| Constant-memory download/upload | Rejected before I/O | Supported |
| HTTP/2 option | Rejected by contract | Opt-in, fallback or required |
| HTTP/3/QUIC | Rejected/unsupported | Diagnostic flag only; unsupported by policy |
| gzip/deflate response decoding | Rejected when active | Supported by WinHTTP option |
| Default/direct/manual proxy | Supported through COM proxy mapping | Supported through WinHTTP session mapping |
| Buffered challenge authentication | Auto scheme only | Explicit bounded scheme/target policy |
| Native handle protocol reporting | COM response contract | `HttpResponse.ProtocolUsed` from WinHTTP |

The public error and response contracts are shared. Capability-specific options
are validated before creating a backend when the selected transport cannot
honor them.

## Native protocol policy

`HttpProtocolOptions` is an explicit native policy. `AllowHttp2=True` with
`Mode=HttpProtocolAllowFallback` permits the OS to use HTTP/2 when available
and accepts HTTP/1.1 otherwise. `HttpProtocolRequired` fails with
`HttpErrorProtocol` when the capability or peer cannot satisfy the policy.

HTTP/2 is attempted only for HTTPS URLs. The loopback test server intentionally
speaks HTTP/1.1, so host-specific HTTP/2 evidence is collected separately and
must report the exact `ProtocolUsed` value. A requested flag is never proof of
negotiation.

HTTP/3/QUIC is unsupported by the current distribution. The `AllowHttp3`
property and native flag remain only so capability probing and future
compatibility work can be represented without changing the enum. Do not enable
it in production code or advertise HTTP/3 support.

## Native decompression

`HttpDecompressionOptions` lets WinHTTP decode gzip and/or deflate. WinHTTP
owns the `Accept-Encoding` header and returns decoded identity bytes. A caller
must not supply a contradictory `Accept-Encoding` header when an active option
is configured. In fallback mode an unavailable option is omitted; in required
mode it maps to `HttpErrorProtocol`. COM rejects active decompression options.

The same policy applies to buffered and streaming native responses. Callers
should verify `BytesWritten` and content hashes rather than trusting a wire
`Content-Length` after decoding.

## Proxy modes

`HttpProxyOptions.Mode` accepts:

- `HttpProxyDefault`: current OS/WinHTTP static configuration;
- `HttpProxyNoProxy`: direct connection; or
- `HttpProxyManual`: an HTTP/HTTPS proxy URL plus an optional WinHTTP bypass
  list.

Proxy URLs cannot contain user-info credentials. Proxy authentication is a
separate provider concern and must not be logged or embedded in a URL. PAC,
WPAD, SOCKS, and trusted corporate CONNECT behavior are host-dependent and not
part of the basic deterministic contract.

## Handle and memory model

Native operations create a session, connection, and request for one execution
and release them in reverse order. Streaming operations use bounded buffers and
same-directory temporary files. The library deliberately avoids VBA callbacks
from native WinHTTP, executable memory, injected code, and descriptor tricks.
These boundaries make the library easier to inspect under endpoint security
software and keep cleanup ownership explicit.

## References

- [Native transport specification](../specs/native-winhttp-transport.md)
- [Protocol policy](../specs/protocol-policy.md)
- [Decompression policy](../specs/decompression-policy.md)
- [Proxy policy](../specs/proxy-policy.md)
- [Compatibility](compatibility.md)
