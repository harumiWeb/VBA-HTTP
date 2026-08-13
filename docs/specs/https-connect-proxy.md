# HTTPS CONNECT proxy fixture

## Scope

This specification defines the deterministic loopback fixture used to prove
that the COM and native transports can reach an HTTPS origin through a manual
HTTP proxy. It is a boundary test, not a claim that every corporate proxy or
trusted certificate environment is supported.

## Fixture contract

The integration runner starts the test server with:

```text
-tls-listen 127.0.0.1:0
-proxy-tls-listen 127.0.0.1:0
-proxy-tls-auth-listen 127.0.0.1:0
```

Readiness contains:

- `https_url`: the TLS origin with an intentionally untrusted certificate;
- `proxy_tls_url`: the ordinary CONNECT proxy URL;
- `proxy_tls_target_url`: the HTTPS origin alias used as the request base URL;
- `proxy_tls_auth_url`: the Basic-challenge CONNECT proxy URL.

The proxy accepts only a CONNECT authority equal to
`proxy_tls_target_url`'s host and port, then dials the server-owned TLS
listener address. It never resolves or forwards arbitrary external hosts.
The authenticated listener returns `407 Proxy Authentication Required` with a
fixed Basic challenge until `proxy-user` / `proxy-pass` is supplied.

`GET /__admin/proxy-stats` on the plain loopback origin returns the test-only
counter object:

```json
{"connect_attempts":1,"authorized_connects":1}
```

`POST /__admin/reset` resets these counters along with the other deterministic
fixture counters. The endpoint is not exposed by any release workbook.

## Transport assertions

Each COM/native test records the counters before the request, sends a GET
through `HttpProxyManual`, and requires:

1. the request fails with `HttpErrorTls` because the fixture certificate is
   not trusted by the operating system;
2. ordinary CONNECT increments `connect_attempts`; or
3. authenticated CONNECT increments `authorized_connects`.

The test must not accept a 200 response, a 407 final response after valid
credentials, a direct-origin response, or a certificate-ignore option. A
trusted CONNECT success remains host-dependent and is not inferred from this
fixture.

## Security boundary

The fixture is loopback-only, target-allowlisted, and uses no PAC/WPAD,
SOCKS, external network, trust-store mutation, or credentials in URLs. Real
Windows-domain Negotiate/NTLM and a production proxy's CONNECT policy require
separate opt-in host evidence. HTTP/3/QUIC evidence is also independent.

## Verification

- Go fixture unit tests cover successful TLS tunneling and the fixed Basic
  challenge (`task testserver:test`).
- COM/native loopback integration covers ordinary and authenticated CONNECT
  boundaries (`task test:integration`).
- The compatibility matrix records these as loopback boundary evidence only;
  it does not promote trusted corporate proxy interoperability.
