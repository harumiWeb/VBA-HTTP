<div align="center">
  <img src="docs/images/logo.png" width="160" alt="VBA-HTTP logo" />
</div>

<h1 align="center">VBA-HTTP</h1>

[MIT License](LICENSE) · Copyright (c) 2026 harumiWeb

VBA-HTTP is a Windows Excel/VBA HTTP client for applications that need a
small, inspectable module distribution, deterministic tests, bounded
concurrency, reliable retries, and constant-memory file transfer.

## Support boundary

- Supported runtime: Windows x64 Office with WinHTTP.
- Default backend: late-bound `WinHttp.WinHttpRequest.5.1` for buffered calls.
- Native backend: synchronous WinHTTP for streaming and advanced capabilities.
- 32-bit Office is unsupported by policy.
- HTTP/3/QUIC is unsupported by policy; HTTP/2 is an opt-in native capability
  that needs host-specific HTTPS evidence.
- OAuth, interactive authentication, PAC/WPAD, SOCKS, and trusted corporate
  credentials are outside the current compatibility guarantee.

See the complete [compatibility guide](docs/guides/compatibility.md) and the
normative [compatibility matrix](docs/specs/compatibility-matrix.md).

## Install the modules

The primary consumer distribution is the VBA-Web-style source package. Download
`VBA-HTTP-vX.Y.Z-source.zip` from a GitHub Release, extract it, and install into
a closed workbook:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass `
  -File .\Install-VBAHttp.ps1 `
  -Workbook C:\Work\Consumer.xlsm
```

The installer verifies the manifest and component hashes, creates a backup, and
imports only production modules. Use `-Force` only for an intentional upgrade;
retain the backup for rollback. The [distribution guide](docs/guides/distribution.md)
contains upgrade, uninstall, checksum, and GitHub Release procedures.

## First request

```vb
Option Explicit

Public Sub GetStatus()
    Dim client As HttpClient
    Dim response As HttpResponse

    Set client = VBAHttp.CreateClient()
    Set response = client.GetResponse("https://api.example.test/status/204")
    response.RaiseForStatus
    Debug.Print response.StatusCode, response.ProtocolUsed
End Sub
```

For request-level control, use `VBAHttp.CreateRequest()` and
`client.Execute(request)`. For native streaming or explicit WinHTTP options,
use `VBAHttp.CreateNativeClient()`.

## What is included

| Area | Public capability |
| --- | --- |
| Requests | Methods, query parameters, validated headers, text/binary bodies. |
| Responses | Status, headers, decoded text, defensive byte ownership, protocol. |
| Reliability | Idempotency-aware retries, `Retry-After`, deadlines, cancellation. |
| Batches | Ordered results with bounded concurrency and partial-failure isolation. |
| Streaming | Atomic constant-memory downloads, file uploads, multipart uploads. |
| Security | TLS validation, Basic/Bearer, bounded buffered challenge auth, redaction. |
| State | Explicit cookie jar and bounded diagnostics; no ambient session state. |
| Routing | Default, direct, and manual proxy modes on COM and native backends. |

## Documentation

The public documentation is organized under [`docs/guides/`](docs/guides/README.md):

- [Getting started](docs/guides/getting-started.md)
- [Complete API reference](docs/guides/api-reference.md)
- [Requests and responses](docs/guides/requests-and-responses.md)
- [Reliability and batches](docs/guides/reliability-and-batches.md)
- [Streaming](docs/guides/streaming.md)
- [Transport capabilities](docs/guides/transport-capabilities.md)
- [Security and state](docs/guides/security-and-state.md)
- [Distribution](docs/guides/distribution.md)
- [Examples](docs/guides/examples.md)
- [Compatibility](docs/guides/compatibility.md)

The legacy [API quick reference](docs/API.md) remains as a compatibility link.
Current contracts are maintained in [`docs/specs/`](docs/specs/README.md), and
architectural rationale is maintained in [`docs/adr/`](docs/adr/README.md).

## Release artifacts

The source ZIP is the primary distribution. A tag push matching `vX.Y.Z` or a
prerelease suffix runs the Excel-free GitHub Release workflow on a hosted
Windows x64 runner. It publishes the source ZIP, a production-only
`xlflow pack --experimental` XLSM, pack/release manifests, SHA-256 sums,
`LICENSE`, and `THIRD_PARTY_NOTICES.md`.

The GitHub pack path does not start Excel and always records **VBE validation not performed**.
It must not be described as VBE-validated. Local
`task precommit` performs the ownership-safe x64 VBE compile gate; optional
`xlflow build` artifacts are validated separately. Existing Releases are never
overwritten. See [`docs/specs/github-release.md`](docs/specs/github-release.md).

## Development

Production code is edited in `src/classes`, `src/modules`, and `src/workbook`;
tests and benchmarks remain excluded development components. The contributor
proof loop is documented in [`CONTRIBUTING.md`](CONTRIBUTING.md):

```powershell
task check
task test:docs
task testserver:test
task precommit
```

Run `task verify` for the complete local verification workflow. Do not edit the
development workbook directly; use xlflow's source-to-workbook proof loop.
Existing Excel processes are baseline-only and are never terminated by the
repository's validation scripts.

## License

Project-authored source, documentation, tools, examples, source packages, and
generated release workbooks are released under the [MIT License](LICENSE).
Redistributions must preserve the copyright and permission notice and include
[`THIRD_PARTY_NOTICES.md`](THIRD_PARTY_NOTICES.md). The pinned VBA-Web checkout
is a benchmark-only comparator and is not a product dependency.

For release handoff, use [`docs/RELEASE_CHECKLIST.md`](docs/RELEASE_CHECKLIST.md),
[`docs/specs/distribution.md`](docs/specs/distribution.md), and the GitHub
release contract above.
