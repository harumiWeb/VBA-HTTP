# ADR-0002: Dual transport boundary

## Status

`accepted`

## Background

VBA-HTTP needs an approachable buffered HTTP path now and constant-memory streaming later. `WinHttp.WinHttpRequest.5.1` provides a dependency-free COM API suitable for buffered requests, but it cannot provide the ownership and incremental I/O guarantees required by large downloads and uploads. Native WinHTTP provides those capabilities at the cost of explicit handles, bitness declarations, and a larger safety surface.

Exposing either backend directly from `HttpClient` would duplicate request semantics, leak implementation details into callers, and make backend selection a breaking public API decision.

## Decision

- `HttpClient` depends only on the class interface `IHttpTransport`.
- The transport boundary accepts an execution snapshot represented by `HttpRequest` and returns an `HttpResponse`; transport failures use the error contract in ADR-0003.
- `WinHttpComTransport` is the default buffered backend for ordinary synchronous and bounded asynchronous requests.
- `WinHttpNativeTransport` is added for streaming and protocol capabilities that require native WinHTTP.
- Public request, response, header, timeout, cancellation, and error semantics are backend-independent.
- `HttpClient.Transport` may be replaced before execution for tests or advanced consumers. A network-free mock transport implements the same interface.
- Automatic backend selection is capability-driven. It must not silently change observable semantics; an explicitly required unsupported capability fails before I/O.
- Production code must not depend on mock, test, benchmark, xlflow, or development-only components.

## Consequences

- Domain behavior and public examples can be tested without network access.
- COM and native implementations can evolve independently while remaining contract-compatible.
- The transport interface stays deliberately small, but future capability discovery must be added without weakening request semantics.
- Native streaming requires separate ownership and callback decisions before implementation; this ADR does not authorize callbacks into VBA application logic.
- Consumers can inject a custom transport, so compatibility tests must cover interface behavior rather than concrete class identity.

## Rationale

- Tests: `src/modules/Tests/Unit/HttpClientTests.bas`
- Code: `src/classes/IHttpTransport.cls`, `src/classes/HttpClient.cls`
- Related specs: `docs/specs/http-core-api.md`

## Supersedes

- None

## Superseded by

- None
