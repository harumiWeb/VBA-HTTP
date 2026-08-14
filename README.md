<div align="center">
  <img src="docs/images/logo.png" width="160" alt="VBA-HTTP logo" />
</div>

<h1 align="center">VBA-HTTP</h1>

<p align="center">
  <strong>A high-performance HTTP client for Excel/VBA on Windows.</strong>
</p>

<p align="center">
  Concurrent requests · Native WinHTTP · Constant-memory streaming · Retries · HTTP/2 (host-specific)
</p>

<p align="center">
  <a href="README.md">English</a>
  |
  <a href="README.ja.md">日本語</a>
</p>

---

VBA-HTTP brings modern HTTP client capabilities to Excel/VBA without requiring
an external runtime or compiled application dependency.

It supports normal buffered requests through WinHTTP COM, while an optional
native WinHTTP backend provides bounded streaming, protocol control, and
advanced transport features. HTTP/2 remains an opt-in, host- and WinHTTP-
dependent capability rather than a universal network guarantee.

```vb
Dim client As HttpClient
Dim response As HttpResponse

Set client = VBAHttp.CreateClient()
Set response = client.GetResponse("https://example.com/api")

response.RaiseForStatus
Debug.Print response.Text
```

For workloads that go beyond a single request, VBA-HTTP includes bounded
concurrency, retries, deadlines, cancellation, streaming file transfer,
authentication, proxies, cookies, diagnostics, and opt-in, host-specific HTTP/2
protocol reporting.

## Why VBA-HTTP?

Most VBA HTTP code eventually becomes a thin wrapper around
`WinHttp.WinHttpRequest.5.1`.

VBA-HTTP goes further.

### Concurrent requests

Run multiple independent requests concurrently while keeping an explicit upper
bound on concurrency.

```vb
Dim urls As New Collection
Dim options As New HttpBatchOptions
Dim result As HttpBatchResult

urls.Add "https://example.com/a"
urls.Add "https://example.com/b"
urls.Add "https://example.com/c"

options.MaxConcurrency = 8

Set result = client.GetMany(urls, options)

Debug.Print result.SuccessCount
Debug.Print result.FailureCount
```

In the current deterministic loopback benchmark:

```text
100 requests × 100 ms server delay

Sequential       11.04 s
Concurrency 16    0.86 s

12.86× faster
```

Concurrency is bounded rather than fire-and-forget, and every item retains its
ordered success, failure, or cancellation result with a stable error category
for timeout and transport failures. Retries are internal attempts, not a
separate batch status.

### Stream large files without buffering them in VBA

The native transport can transfer files incrementally instead of materializing
the entire payload as a VBA `Byte()` array.

```vb
Dim client As HttpClient
Dim download As HttpDownloadResult

Set client = VBAHttp.CreateNativeClient()

Set download = client.DownloadFile( _
    "https://example.com/large.bin", _
    "C:\Temp\large.bin")

download.RaiseForStatus
Debug.Print download.BytesWritten
```

A recorded Phase 6 baseline 1 GiB download completed with approximately
**19 MB of peak private-memory growth** in the benchmarked Excel process. This
is a single x64 engineering measurement, not a memory budget or a post-
optimization guarantee; repeated before/after evidence is tracked separately.

Downloads are written to a temporary sibling file and atomically published only
after the transfer succeeds. Existing destination files remain intact after
HTTP failures, cancellation, timeout, or write errors.

Streaming uploads and multipart uploads use the same incremental model:

```vb
Set result = client.UploadFile( _
    "https://example.com/upload", _
    "C:\Temp\payload.bin")
```

```vb
Dim form As HttpMultipartForm

Set form = VBAHttp.CreateMultipartForm()
form.AddField "title", "example"
form.AddFile "payload", "C:\Temp\payload.bin"

Set result = client.UploadMultipart( _
    "https://example.com/upload", _
    form)
```

Large payloads do not need to be represented as one giant VBA string or byte
array.

### Native WinHTTP when you need it

The default client uses the late-bound Windows WinHTTP COM interface:

```vb
Set client = VBAHttp.CreateClient()
```

For streaming and advanced transport control:

```vb
Set client = VBAHttp.CreateNativeClient()
```

The native backend calls the documented Windows `winhttp.dll` API directly from
VBA and provides:

- streaming download and upload
- host-specific HTTP/2 protocol control and protocol reporting
- response decompression
- native proxy configuration
- bounded buffered Windows/server/proxy challenge authentication
- deterministic WinHTTP handle ownership and cleanup

No custom DLL is required.

### Reliability built in

Retries are not implemented as an unconditional loop around `Send`.

VBA-HTTP understands:

- idempotent vs. non-idempotent methods
- `Retry-After`
- exponential backoff
- jitter
- maximum attempt counts
- request deadlines
- total operation deadlines
- cooperative cancellation
- transient DNS, connection, timeout, and I/O failures

By default, retry-safe methods such as `GET`, `HEAD`, `PUT`, and `DELETE` can be
retried for transient failures and selected HTTP status codes.

`POST`, `PATCH`, and custom methods require explicit opt-in.

### Security-sensitive behavior is explicit

VBA-HTTP avoids silently carrying sensitive state across requests.

It includes:

- OS certificate validation
- Basic authentication
- Bearer authentication
- Windows challenge authentication
- bounded buffered proxy challenge authentication
- redirect downgrade protection
- suppression of sensitive headers across redirects
- optional caller-owned cookie jars
- diagnostics with credential and cookie redaction

Cookies and diagnostics are explicit objects rather than hidden global state.

## Performance

VBA-HTTP includes reproducible benchmarks backed by a deterministic local Go
HTTP server.

Recorded baseline highlights on Windows 11 / Excel x64 include:

| Workload | Result |
| --- | ---: |
| 100 requests × 100 ms, sequential | 11.04 s |
| 100 requests × 100 ms, concurrency 16 | 0.86 s |
| Concurrent speedup | **12.86×** |
| 1 GiB streaming download | **44.9 s** |
| Download peak private-memory growth | **~19 MB** |
| 1 GiB streaming upload | **16.0 s** |

The 1 GiB download/upload values are the Phase 6/7 baseline runs recorded
before the latest native hot-path optimization. The current repository does not
promote a new post-optimization throughput or memory claim until a repeated,
PID-scoped x64 before/after measurement is saved. The historical matched
buffered comparison also measured VBA-HTTP at 0.833 ms versus Raw
WinHttpRequest at 0.499 ms; the wrapper overhead is documented rather than
hidden.

These are measurements from one controlled machine and are not presented as
universal performance claims.

The benchmark suite records machine-readable results and validates the transfer
contents, concurrency bounds, resource cleanup, and payload hashes.

See [`benchmarks/`](benchmarks/) and the benchmark methodology documentation for
the complete evidence.

## Features

| Area | Capability |
| --- | --- |
| Requests | GET/POST/PUT/PATCH/DELETE, query parameters, headers, text and binary bodies |
| Responses | Status, headers, bytes, decoded text, negotiated protocol |
| Concurrency | Ordered bounded batches with configurable maximum concurrency |
| Reliability | Retries, `Retry-After`, backoff, jitter, deadlines, cancellation |
| Downloads | Constant-memory streaming with atomic publication |
| Uploads | Streaming file and multipart upload |
| Protocols | Native WinHTTP HTTP/2 opt-in and protocol reporting (host-specific) |
| Compression | Native gzip/deflate response decompression |
| Authentication | Basic, Bearer, bounded buffered server/proxy challenge auth |
| Proxy | System/default, direct, and manual proxy modes |
| Cookies | Optional caller-owned cookie jar |
| Diagnostics | Bounded structured events with sensitive-data redaction |
| Testing | Deterministic integration, stress, resource, and benchmark fixtures |

## Installation

The primary distribution is a source package that can be inspected before it is
imported into a workbook.

Download:

```text
VBA-HTTP-vX.Y.Z-source.zip
```

from a GitHub Release, extract it, and install it into a **closed** workbook:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass `
  -File .\Install-VBAHttp.ps1 `
  -Workbook C:\Work\Consumer.xlsm
```

The installer:

- validates the package manifest
- verifies component hashes
- creates a workbook backup
- imports production modules only
- supports deliberate upgrades with `-Force`

See the [distribution guide](docs/guides/distribution.md) for upgrade,
uninstall, rollback, and checksum procedures.

## Quick examples

### GET

```vb
Set response = client.GetResponse("https://example.com/api")
response.RaiseForStatus

Debug.Print response.StatusCode
Debug.Print response.Text
```

### Query parameters

```vb
Dim request As HttpRequest
Dim response As HttpResponse

Set request = VBAHttp.CreateRequest()
request.Method = "GET"
request.Url = "https://example.com/items"
request.Query.Add "page", 1
request.Query.Add "limit", 100

Set response = client.Execute(request)
response.RaiseForStatus
```

### Custom request

```vb
Dim request As HttpRequest

Set request = VBAHttp.CreateRequest()

request.Method = "GET"
request.Url = "https://example.com/api"
request.Headers.SetValue "Accept", "application/json"

Set response = client.Execute(request)
```

### Retry policy

```vb
Dim policy As HttpRetryPolicy

Set policy = VBAHttp.CreateRetryPolicy()
policy.MaxAttempts = 4

Set client.RetryPolicy = policy
```

### HTTP/2

```vb
Dim protocols As HttpProtocolOptions

Set client = VBAHttp.CreateNativeClient()
Set protocols = VBAHttp.CreateProtocolOptions()

protocols.AllowHttp2 = True
protocols.Mode = HttpProtocolAllowFallback

Set client.ProtocolOptions = protocols

Set response = client.GetResponse("https://example.com/")
Debug.Print response.ProtocolUsed
```

## Designed to be inspectable

VBA-HTTP is intentionally distributed as ordinary VBA source.

There is no hidden service process and no custom networking DLL behind the
public API.

The native implementation uses documented Windows APIs for networking,
resource management, and streaming while keeping explicit ownership around
WinHTTP handles.

The repository also contains:

- unit tests
- real Excel integration tests
- cancellation and resource stress tests
- deterministic HTTP and HTTPS fixtures
- proxy and authentication fixtures
- reproducible benchmarks
- architecture decision records
- machine-readable benchmark evidence

The goal is not only to make HTTP work in VBA, but to make its behavior
**testable, measurable, and explainable**.

## Compatibility

VBA-HTTP currently targets:

| Platform | Status |
| --- | --- |
| Windows x64 Office | **Supported** |
| Windows 32-bit Office | **Unverified; community validation welcome** |
| macOS Office | Unsupported |
| HTTP/1.1 | Supported |
| HTTP/2 | Native opt-in; host- and WinHTTP-dependent, with host-specific evidence |
| HTTP/3 / QUIC | **Unsupported by policy** |

The normal buffered client uses:

```text
WinHttp.WinHttpRequest.5.1
```

The native client uses Windows WinHTTP directly.

OAuth flows, interactive authentication, PAC/WPAD, SOCKS, and trusted
corporate credential environments are outside the current compatibility
guarantee.

32-bit Office may work on some hosts, but it is not an officially supported
release target yet. Contributors can run the non-promotional diagnostic path:

```powershell
powershell -File tools/Run-OfficeBitnessValidation.ps1 `
  -ExpectedArchitecture X86 -DiagnosticOnly
```

Successful evidence is reviewed separately and does not change the x64 release
boundary by itself.

See the [compatibility guide](docs/guides/compatibility.md) and
[compatibility matrix](docs/specs/compatibility-matrix.md).

## Documentation

Start here:

- [Guide index](docs/guides/README.md)
- [Getting started](docs/guides/getting-started.md)
- [API reference](docs/guides/api-reference.md)
- [Requests and responses](docs/guides/requests-and-responses.md)
- [Reliability and batches](docs/guides/reliability-and-batches.md)
- [Streaming](docs/guides/streaming.md)
- [Transport capabilities](docs/guides/transport-capabilities.md)
- [Security and state](docs/guides/security-and-state.md)
- [Distribution](docs/guides/distribution.md)
- [Examples](docs/guides/examples.md)
- [Compatibility](docs/guides/compatibility.md)

Normative contracts live under [`docs/specs/`](docs/specs/README.md).

Architecture decisions and implementation rationale live under
[`docs/adr/`](docs/adr/README.md).

## Built and verified with xlflow

VBA-HTTP is also a real-world stress test for
[xlflow](https://github.com/harumiWeb/xlflow), an environment for developing
Excel/VBA projects with modern source control, testing, diagnostics, and coding
agents.

The project is developed from source files rather than by manually editing the
VBE. Its verification loop includes real Excel compilation, unit tests,
integration tests, stress tests, static analysis, deterministic local services,
and reproducible benchmarks.

VBA-HTTP was used to push that workflow well beyond small macros and into a
native networking library with dozens of modules and system-level integration.

## Development

Production code lives in:

```text
src/classes
src/modules
src/workbook
```

Common verification commands:

```powershell
task check
task test:docs
task testserver:test
task precommit
```

Run the complete local verification workflow with:

```powershell
task verify
```

Do not edit the development workbook directly. The source tree is authoritative
and xlflow provides the source-to-workbook compile and test loop.

See [`CONTRIBUTING.md`](CONTRIBUTING.md).

## Release artifacts

The source ZIP is the primary consumer artifact.

GitHub Releases also publish:

- production source package
- production-only XLSM
- release and pack manifests
- SHA-256 checksums
- `LICENSE`
- `THIRD_PARTY_NOTICES.md`

GitHub-hosted packaging does not start Excel and therefore records the status
**VBE validation not performed**; the pack XLSM is not a VBE-validated
artifact.

The local release gate performs the real Excel/VBE compile validation before a
validated release artifact is accepted.

See the [legacy API quick reference](docs/API.md),
[`docs/specs/distribution.md`](docs/specs/distribution.md),
[`docs/specs/github-release.md`](docs/specs/github-release.md), and
[`docs/RELEASE_CHECKLIST.md`](docs/RELEASE_CHECKLIST.md).

## License

VBA-HTTP is released under the [MIT License](LICENSE).

Project-authored source code, documentation, tools, examples, source packages,
and generated release workbooks are covered by that license.

Redistributions must preserve the copyright and permission notice and include
[`THIRD_PARTY_NOTICES.md`](THIRD_PARTY_NOTICES.md).

The pinned VBA-Web checkout used by the benchmark suite is a comparator only
and is not a runtime dependency.

---

**VBA-HTTP — modern HTTP plumbing, pushed unusually far for VBA.**
