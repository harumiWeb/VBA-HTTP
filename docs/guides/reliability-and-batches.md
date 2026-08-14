# Reliability and batches

## Retry defaults

`HttpClient.RetryPolicy` is cloned into each buffered operation. The defaults
are:

| Setting | Default |
| --- | ---: |
| `MaxAttempts` | 3, including the first attempt |
| `BaseDelayMilliseconds` | 200 |
| `MaxDelayMilliseconds` | 30,000 |
| `JitterRatio` | 0.2 |
| `RespectRetryAfter` | `True` |
| `RetryNonIdempotentMethods` | `False` |

GET, HEAD, PUT, DELETE, OPTIONS, and TRACE may retry 408, 429, 500, 502, 503,
504 and transient DNS, connection, timeout, or I/O failures. POST, PATCH, and
unknown extension methods are not retried unless explicitly opted in. A final
retryable HTTP response remains an ordinary `HttpResponse`.

```vb
Dim policy As HttpRetryPolicy
Dim execution As HttpExecutionOptions

Set policy = VBAHttp.CreateRetryPolicy()
policy.MaxAttempts = 4
policy.RetryNonIdempotentMethods = True
Set execution = VBAHttp.CreateExecutionOptions()
Set execution.RetryPolicy = policy
execution.TotalDeadlineMilliseconds = 10000

Set response = client.PostResponse("/jobs", request.Body, execution)
```

The exponential delay is capped at `MaxDelayMilliseconds`. A valid
`Retry-After` delta-seconds or HTTP-date replaces the calculated delay when
`RespectRetryAfter=True`. Cancellation is never retried.

Streaming download and upload operations do not replay one-shot bodies. Repeat
the complete operation explicitly if the application can safely repeat it.

## Timeouts, deadlines, and cancellation

`HttpTimeouts` controls resolve, connect, send, and receive phases. A zero
phase value delegates to the backend default. `HttpExecutionOptions` adds a
total deadline and cancellation token:

```vb
Dim token As HttpCancellationToken
Dim execution As HttpExecutionOptions

Set token = VBAHttp.CreateCancellationToken()
Set execution = VBAHttp.CreateExecutionOptions()
Set execution.CancellationToken = token
execution.TotalDeadlineMilliseconds = 30000
' Later, from a cooperating host callback:
token.Cancel
```

Checkpoints occur before every attempt, while COM async work is polled, during
retry waits, and after an attempt. Cancellation has priority over an expired
total deadline. A blocking custom transport that cannot observe these controls
is rejected before I/O; native buffered calls cannot be interrupted mid-call,
while native streaming checks between bounded reads/writes.

## Bounded concurrency

`GetMany` accepts a `Collection` of URLs. `ExecuteMany` accepts a `Collection`
of `HttpRequest` objects. `HttpBatchOptions.MaxConcurrency` bounds in-flight
requests (default 8). Retry waits do not hold a slot, input order is preserved,
and one failed item does not fail the entire batch.

```vb
Dim urls As New Collection
Dim batchOptions As HttpBatchOptions
Dim results As HttpBatchResult

urls.Add "/users/1"
urls.Add "/users/2"
Set batchOptions = VBAHttp.CreateBatchOptions()
batchOptions.MaxConcurrency = 16
Set results = client.GetMany(urls, batchOptions)

Dim item As HttpBatchItem
Set item = results.ItemAt(1)
If item.Status = HttpBatchItemSucceeded Then
    Debug.Print item.Response.StatusCode
Else
    Debug.Print item.ErrorCategory, item.ErrorDescription
End If
```

`HttpBatchOptions.RequestDeadlineMilliseconds` applies to each scheduled
request attempt. `TotalDeadlineMilliseconds` covers the complete batch. A
cancelled or expired batch marks every unfinished item with the corresponding
stable category and returns the results already collected.

## Host yielding and polling

`PollIntervalMilliseconds` controls the COM async poll cadence. `YieldToHost`
and `YieldIntervalMilliseconds` allow controlled `DoEvents`/host yielding for
long-running Excel operations. Yielding is opt-in where a callback or workbook
reentrancy policy requires it; a progress/auth callback must not call the same
client recursively.

## Normative details

The complete retry classification, jitter calculation, deadline precedence,
and deterministic test seams are in
[`../specs/reliability-policy.md`](../specs/reliability-policy.md) and
[`../specs/bounded-concurrency.md`](../specs/bounded-concurrency.md).
