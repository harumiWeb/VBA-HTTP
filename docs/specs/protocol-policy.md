# Native protocol policy

## Scope

`HttpProtocolOptions` is a native WinHTTP capability. It controls optional
HTTP/2 negotiation without changing the default operating-system policy. The
retained HTTP/3 flag is diagnostic/future-compatibility code and is
`unsupported-by-policy`; it is not a supported product capability or release
gate. The COM transport rejects a non-empty advanced protocol mask because the
COM option enum has no equivalent HTTP/2/HTTP/3 control.

## API

```vb
Dim protocols As HttpProtocolOptions
Set protocols = VBAHttp.CreateProtocolOptions()
protocols.AllowHttp2 = True
protocols.Mode = HttpProtocolAllowFallback

Set client.ProtocolOptions = protocols
Set response = client.GetResponse("https://api.example.test/items")
Debug.Print response.ProtocolUsed
```

`EnabledProtocols = 0` means no native override. `HttpProtocolRequired` needs
at least one advanced protocol flag. Options are cloned into request snapshots;
later mutations by the caller cannot alter an in-flight request. Setting
`AllowHttp3` remains available for diagnostics only and must not be used as a
support or release claim.

## Fallback and required behavior

- Native requests apply WinHTTP option 133 before headers/body are sent.
- In allow-fallback mode, an unsupported option is treated as a capability miss
  and the request may use HTTP/1.1. The implementation treats both
  `ERROR_WINHTTP_INVALID_OPTION` (12009) and Win32 `ERROR_NOT_SUPPORTED` (50)
  as capability misses. The response's `ProtocolUsed` is the authoritative
  result.
- In required mode, option 145 prevents fallback. Unsupported runtime options
  (12009 or 50), protocol mismatch (WinHTTP error 12190), or a required
  advanced protocol on a plain HTTP URL raise `HttpErrProtocol` before a
  successful response is returned.
- HTTP/2/3 flags are never inferred from the request. A requested flag is not
  evidence that the peer negotiated that protocol. HTTP/3 is excluded from the
  supported compatibility matrix under ADR-0035.

## Compatibility

Modern protocol options are only attempted on HTTPS URLs. The loopback test
server intentionally remains HTTP/1.1; integration tests prove deterministic
fallback and required-mode rejection. HTTP/2 promotion uses the host runner;
HTTP/3 checks are diagnostic-only and cannot promote a matrix entry. Any future
support decision must record Windows version, WinHTTP capability, Office
bitness, requested flags, and `ProtocolUsed` under a superseding ADR.

## Evidence

- Unit: `src/modules/Tests/Unit/HttpProtocolOptionsTests.bas`,
  `HttpRequestTests.bas`, `HttpClientTests.bas`, and `WinHttpNativeTests.bas`.
- Integration: `WinHttpNativeTransportTests.Test_NativeTransport_ProtocolOptionsFallbackOnPlainHttp`
  and `Test_NativeTransport_RequiredProtocolRejectsPlainHttp`.
- ADR: `docs/adr/ADR-0009-native-protocol-negotiation-policy.md`.
