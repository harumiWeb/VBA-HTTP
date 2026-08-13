# Local HTTP test server

## Purpose

`tools/testserver/` provides deterministic HTTP behavior for VBA integration tests and benchmarks. It binds to loopback by default and requires no external network access.

## Lifecycle contract

- Start with `go run . -listen 127.0.0.1:0` from `tools/testserver/`. Non-loopback listeners are rejected.
- Passing `-proxy-listen 127.0.0.1:0` starts a deterministic HTTP forward proxy
  restricted to the target server; readiness then includes `proxy_url` and
  `proxy_target_url`. The target URL uses `vba-http.localhost` so WinHTTP does
  not silently bypass the explicitly configured proxy for a literal loopback
  address.
- Passing `-proxy-auth-listen 127.0.0.1:0` starts a separate deterministic
  Basic-authenticated HTTP forward proxy; readiness then includes
  `proxy_auth_url`. Its fixed test-only credential is `proxy-user` /
  `proxy-pass` and the listener remains loopback-only.
- Passing `-proxy-tls-listen 127.0.0.1:0` together with `-tls-listen` starts a
  deterministic HTTP CONNECT proxy restricted to the server's own TLS
  listener; readiness includes `proxy_tls_url` and `proxy_tls_target_url`.
- Passing `-proxy-tls-auth-listen 127.0.0.1:0` adds a fixed Basic-challenge
  CONNECT listener and readiness includes `proxy_tls_auth_url`. CONNECT
  attempts and authorized attempts are observable only through the plain
  origin's `GET /__admin/proxy-stats` endpoint.
- Passing `-tls-listen 127.0.0.1:0` starts a second HTTPS listener with a
  process-local, self-signed certificate whose trust chain is intentionally not
  installed in the OS. Readiness then includes `https_url`; the listener is
  used only for certificate-rejection tests and never as a trust bypass.
- Port `0` asks the OS for a free port. The first stdout line is JSON:
  `{"event":"ready","url":"http://127.0.0.1:<port>"}` or, when enabled,
  includes `https_url`, `proxy_url`, `proxy_target_url`, `proxy_auth_url`,
  `proxy_tls_url`, `proxy_tls_target_url`, and `proxy_tls_auth_url` for the
  requested listeners.
- The process accepts Ctrl+C or termination and allows five seconds for graceful shutdown.
- `GET /healthz` returns 200 when ready.
- `POST /__admin/reset` clears all flaky and rate-limit counters and returns 204.
- `POST /__admin/shutdown` returns 202 and requests the same graceful shutdown used for process signals.
- Stateful endpoints use the optional `id` query parameter as an isolation key. Tests that may overlap must use distinct IDs.

## Endpoint contract

| Endpoint | Behavior |
| --- | --- |
| `GET /status/{code}` | Returns the requested final HTTP status from 200 through 599. |
| `GET /delay/{milliseconds}` | Delays 0 through 300,000 ms, observes client cancellation, then returns JSON. |
| `GET /bytes/{size}` | Returns the deterministic byte pattern with `Content-Length`. |
| `GET /stream/{size}` | Returns the same pattern with periodic flushes and no declared length. |
| `GET /sha256/{size}` | Returns the expected SHA-256 digest and byte count for the generated pattern. |
| `GET /unicode` | Returns the fixed UTF-8 text fixture `VBA-HTTP unicode: 日本語🙂` with an explicit UTF-8 charset. |
| `GET /malformed-utf8` | Returns the fixed invalid UTF-8 byte sequence `C3 28` with an explicit UTF-8 charset; clients must reject text materialization. |
| `GET /malformed-headers` | Writes a raw HTTP response containing a header line without a colon; COM/native clients must classify the exchange as `HttpErrorProtocol`. |
| `GET /headers` | Returns method, path, query, and request headers as JSON. |
| `GET /auth/basic` | Verifies the fixed test-only Basic credential and returns 204 plus `X-Auth-Verified: 1`; missing or wrong credentials return 401 with a deterministic challenge and never reflect the value. |
| `GET /auth/bearer` | Verifies the fixed test-only Bearer token and returns 204 plus `X-Auth-Verified: 1`; missing or wrong credentials return 401 with a deterministic challenge and never reflect the value. |
| `POST /echo` | Streams the request body back with its content type. |
| `GET /redirect/{count}` | Returns 302 toward the next lower count and 200 at zero. |
| `GET /flaky/{failCount}` | Returns 503 for the first N attempts and 200 afterward. |
| `GET /rate-limit/{count}` | Returns 429 for the first N attempts and 200 afterward. `retry_after` controls the response header. |
| `GET／POST／PATCH /retry-status/{code}/{count}` | Returns the selected 4xx／5xx status for the first N attempts and 200 afterward. |
| `POST /upload/hash` | Streams the request body into a SHA-256 digest and returns `X-Upload-Digest`, `X-Upload-Bytes`, and the received `Content-Length`. |
| `POST /upload/slow/{milliseconds}` | Delays each request-body read by the selected amount, then streams the same hash response; used for cancellation and deadline tests. |
| `POST /upload/multipart` | Parses an ordered multipart form and returns file hashes, file byte counts, filenames, and UTF-8 field values in response headers and JSON. |
| `POST /upload/challenge` | Returns a deterministic 401 with `WWW-Authenticate` and never consumes the request body. |

The optional loopback HTTP proxy forwards only to the server's own target
address, adds `X-Test-Proxy-Forwarded: 1`, and returns 502 for any non-target
host. The authenticated proxy has the same target restriction and returns 407
with a fixed `Proxy-Authenticate: Basic` challenge until the test credential
is sent. The optional TLS proxy accepts only CONNECT to the server-owned TLS
alias, tunnels bytes to that listener, and never forwards arbitrary external
addresses. Its authenticated variant applies the same fixed 407 challenge.
The TLS certificate remains intentionally untrusted; the integration contract
expects normal certificate rejection after tunnel establishment.

The optional HTTPS listener uses the same route set as the HTTP listener. Its
certificate is deliberately untrusted, so clients must fail through normal OS
certificate validation; tests must never install or ignore that certificate.

Sizes are decimal bytes from zero through 2 GiB. The generated byte at zero-based offset `i` is `i mod 251`; both byte endpoints generate fixed-size chunks and never allocate the full payload. Tests and benchmarks verify payloads with SHA-256.

Invalid path parameters return 400 JSON. Endpoint behavior must remain backward compatible once referenced by a released integration test or benchmark schema.

`task testserver:test` runs unit tests, `go vet`, and a real-process lifecycle smoke. The smoke uses an OS-selected port, verifies readiness and reset, downloads the 100 MiB streaming fixture with constant client buffering, compares SHA-256, requests graceful shutdown, and confirms that no server process or generated artifact remains.
