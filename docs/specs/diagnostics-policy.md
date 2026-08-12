# Structured diagnostics policy

`HttpDiagnostics` is an opt-in, caller-owned collector attached to a
`HttpClient`. It is disabled by default. The collector is safe for application
logging, but it is not a wire trace and must not be treated as one.

## Event contract

Each completed top-level operation appends one event per response/result:

- `operation`: `execute`, `batch`, `download`, or `upload`.
- `sequence`: monotonically increasing within the collector; `Clear` resets it.
- `method`: the prepared HTTP method, or `BATCH` for a batch item without a
  prepared request.
- `target`: scheme, authority without user-info, and path only. Query and
  fragment values are removed.
- `status_code`: HTTP status for a response/result, or `0` for a transport
  failure.
- `protocol`: negotiated protocol when available.
- `elapsed_ms`: non-negative measured duration when available.
- `error_number` and `error_category`: stable `HttpErrors` values; zero and
  `none` indicate no raised transport error.
- `request_headers` and `response_headers`: ordered arrays of `{name,value}`.

Events never contain request/response bodies, reason phrases, error
descriptions, query values, raw URLs with user-info, or backend-native text.
Sensitive header values are always `[REDACTED]`, including values added by a
cookie jar or authentication provider.

## Retention and serialization

`MaxEvents` must be between 1 and 10,000 and defaults to 100. The oldest event
is discarded when the bound is exceeded. `ToJson` returns:

```json
{"schema_version":1,"enabled":true,"events":[{"schema_version":1,...}]}
```

Field order is stable for snapshot tests. JSON escaping is performed by the
library; consumers must not concatenate event fields into JSON themselves.

## Integration boundary

`HttpClient` records events after the operation has a response/result and when
an operation raises a stable library error. Retry attempts are intentionally
not emitted. Direct calls to a replaceable `IHttpTransport` bypass the client
collector and must provide their own diagnostics policy.

Diagnostics recording is fail-open: an event construction/serialization error
is ignored and cannot change the HTTP result. `HttpSecurity.RedactHeaderValue`
is the single sensitive-header classification used by event construction.

## Verification

- `HttpDiagnosticsTests` covers bounds, ordering, target sanitization, JSON
  escaping, and redaction.
- `HttpClientTests` covers response and failure event capture with a mock
  transport.
- Release consumer smoke asserts that authorization/cookie sentinels do not
  occur in the serialized event stream.
