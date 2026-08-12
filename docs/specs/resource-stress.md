# Resource stress specification

## Workload

The Phase 9 resource gate runs against the deterministic local server's
`GET /status/204` endpoint with external network access disabled. Each scenario
uses the configured iteration count (default 10,000):

| Scenario | Backend | Scheduler | Default concurrency |
| --- | --- | --- | ---: |
| `sequential_native` | `WinHttpNativeTransport` | one synchronous request at a time | 1 |
| `scheduled_com` | `WinHttpComTransport` | `HttpClient.ExecuteMany` | 16 |

The runner can select one scenario for diagnosis, but the complete release
evidence requires both. The stress source is under `src/modules/Tests/Stress/`
and is excluded from the release workbook.

## Process sampling

`Run-ResourceStressTests.ps1` records the Excel process IDs present before the
scenario and samples only newly created Excel processes. The VBA test performs a
short warmup, writes a start marker immediately before the measured workload,
writes a done marker after all requests and native handle assertions pass, and
waits for a release marker. The runner uses that pause to capture:

- `process_before`: immediately after the start marker;
- `process_peak`: maximum observed handles, working set, and private bytes;
- `process_after`: one second after the done marker;
- `process_idle_after`: a second one-second idle sample before release.

If the target Excel process disappears or a marker is not observed, the runner
fails closed. Existing Excel processes are never included in the totals.

## Gates

- Sequential native idle handle delta must be `<= 8`.
- Scheduled COM idle handle delta must be `<= 32`.
- Every request must return status 204; scheduled results must have zero failed
  or cancelled items.
- Memory peaks are recorded but not given a universal pass threshold because
  Excel allocator reuse and Office versions differ. Review compares them with
  the same-host baseline.

## Result contract

The runner atomically publishes the JSON result only after both selected
scenarios pass. The result contains no request headers, credentials, file paths,
or external URLs beyond the loopback base URL. Validate it with
`benchmarks/schema/resource-stress-result.schema.json` and
`tools/Validate-ResourceStressResult.ps1`.
