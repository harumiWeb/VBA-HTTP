# ADR-0003: HTTP error model

## Status

`accepted`

## Background

HTTP status failures and transport failures carry different information and support different recovery decisions. Treating every 4xx or 5xx as a VBA runtime error discards the response body and headers callers commonly need. Returning every DNS, TLS, timeout, validation, or cancellation failure as a partially initialized response makes failure handling ambiguous and complicates retry and batch execution.

VBA has no native exception type hierarchy, so the library also needs a stable error-number namespace and machine-readable categories without exposing raw COM or WinHTTP numbers as its public contract.

## Decision

- A completed HTTP exchange returns `HttpResponse` for every valid status code, including 4xx and 5xx.
- `HttpResponse.IsSuccess` is true only for status codes 200 through 299.
- `HttpResponse.RaiseForStatus` is the explicit conversion from a non-success response to a library VBA error.
- Invalid public input and transport failures raise within the reserved range beginning at `vbObjectError + 21000`.
- Public categories distinguish validation, invalid URL, DNS, connection, TLS, timeout, cancellation, protocol, I/O, and HTTP status failure.
- A raised library error uses `HttpErrors` for stable number, category, source, and description mapping. Backend-native codes may be retained as diagnostics but are not the primary public error number.
- Single-request APIs raise transport failures. Future batch APIs capture each failure as an item result so one failed request does not abort unrelated items.
- Secrets, credentials, and sensitive header values must not appear in public error descriptions.
- Structured operation evidence uses the opt-in `HttpDiagnostics` collector;
  its redaction and retention contract is defined by ADR-0019 rather than by
  localized `Err.Description` text.

## Consequences

- Callers can inspect error response headers and bodies without exception control flow.
- Retry policy can classify transport failure independently from HTTP status.
- Consumers must call `RaiseForStatus` when exception-style status handling is desired.
- Stable error numbers require additive evolution; existing meanings cannot be reassigned.
- VBA's global `Err` object remains the throw mechanism, so structured diagnostics beyond number/category/source require the explicit `HttpDiagnostics` helper defined by ADR-0019.

## Rationale

- Tests: `src/modules/Tests/Unit/HttpResponseTests.bas`, `src/modules/Tests/Unit/HttpClientTests.bas`
- Code: `src/modules/HttpErrors.bas`, `src/classes/HttpResponse.cls`
- Related specs: `docs/specs/http-core-api.md`

## Supersedes

- `ADR-0019-structured-diagnostics-boundary.md` (structured diagnostics and redaction)

## Superseded by

- None
