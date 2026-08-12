# Native protocol policy

## Scope

`HttpProtocolOptions` is a native WinHTTP capability. It controls optional
HTTP/2 and HTTP/3 negotiation without changing the default operating-system
policy. The COM transport rejects a non-empty advanced protocol mask because
the COM option enum has no equivalent HTTP/2/HTTP/3 control.

## API

```vb
Dim protocols As HttpProtocolOptions
Set protocols = VBAHttp.CreateProtocolOptions()
protocols.AllowHttp2 = True
protocols.AllowHttp3 = True
protocols.Mode = HttpProtocolAllowFallback

Set client.ProtocolOptions = protocols
Set response = client.GetResponse("https://api.example.test/items")
Debug.Print response.ProtocolUsed
```

`EnabledProtocols = 0` means no native override. `HttpProtocolRequired` needs
at least one advanced protocol flag. Options are cloned into request snapshots;
later mutations by the caller cannot alter an in-flight request.

## Fallback and required behavior

- Native requests apply WinHTTP option 133 before headers/body are sent.
- In allow-fallback mode, an unsupported option is treated as a capability miss
  and the request may use HTTP/1.1. The response's `ProtocolUsed` is the
  authoritative result.
- In required mode, option 145 prevents fallback. Unsupported runtime options,
  protocol mismatch (WinHTTP error 12190), or a required advanced protocol on
  a plain HTTP URL raise `HttpErrProtocol` before a successful response is
  returned.
- HTTP/2/3 flags are never inferred from the request. A requested flag is not
  evidence that the peer negotiated that protocol.

## Compatibility

Modern protocol options are only attempted on HTTPS URLs. The loopback test
server intentionally remains HTTP/1.1; integration tests prove deterministic
fallback and required-mode rejection. A future TLS fixture must record Windows
version, WinHTTP capability, Office bitness, requested flags, and
`ProtocolUsed` before a matrix entry is marked supported.

## Evidence

- Unit: `src/modules/Tests/Unit/HttpProtocolOptionsTests.bas`,
  `HttpRequestTests.bas`, `HttpClientTests.bas`, and `WinHttpNativeTests.bas`.
- Integration: `WinHttpNativeTransportTests.Test_NativeTransport_ProtocolOptionsFallbackOnPlainHttp`
  and `Test_NativeTransport_RequiredProtocolRejectsPlainHttp`.
- ADR: `docs/adr/ADR-0009-native-protocol-negotiation-policy.md`.
