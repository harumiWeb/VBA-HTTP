# ADR-0015: Resource stress evidence policy

## Status

`accepted`

## Background

The existing Phase 6/7 stress runners measure memory while a one-gigabyte
stream is active, and the integration suite checks only short native sequences.
That evidence does not prove that repeated request handles are released or that
the bounded scheduler remains stable after a long workload. Process-wide Excel
sampling also risks including an unrelated user workbook.

## Decision

- Add a dedicated excluded stress module with two independent scenarios:
  `sequential_native` uses `WinHttpNativeTransport` for a 10,000-request
  warmup followed by 10,000 measured buffered `GET /status/204` requests, and
  `scheduled_com` uses the COM bounded scheduler with the same warmup/measured
  shape and a configurable maximum concurrency (default 16).
- Run each scenario in its own temporary xlflow test process. The VBA test
  publishes start/done markers and waits at a release gate. The PowerShell
  runner samples only Excel processes created after the scenario began, captures
  before/peak/after/idle process handles and memory, then opens the release gate.
  This keeps the after snapshot alive long enough to make the handle check
  meaningful and excludes pre-existing user Excel processes.
- The warmup is part of the evidence contract: it absorbs one-time WinHTTP and
  Excel allocator pools. The strict leak gate is the measured idle handle
  delta after that warmup: at most 8 for sequential native requests and at most
  32 for scheduled COM requests. Working-set and private
  memory peaks are recorded as engineering evidence, not rejected by an
  allocator-dependent absolute threshold.
- Results are written atomically to
  `benchmarks/results/phase9-resource-stress.json` and validated against
  `benchmarks/schema/resource-stress-result.schema.json`. The default task runs
  both scenarios; a single scenario can be selected for diagnosis without
  replacing the complete gate artifact.
- The stress module and runner are development-only. They are excluded from
  release workbooks and never become a production dependency.

## Consequences

- A release-quality resource claim now has reproducible workload, process scope,
  idle sampling, and machine-readable evidence rather than only a short unit
  assertion.
- The complete gate is intentionally slower than ordinary integration tests and
  is a separate Taskfile target; it is not run on every pre-commit.
- Memory peaks can vary with Excel and Windows allocator behavior, so they are
  published for review while persistent handle growth remains a hard failure.
- Cancellation/timeout stress and cookie persistence remain separate decisions;
  this ADR does not add ambient state or long-lived session pooling.

## Rationale

- Code: `src/modules/Tests/Stress/WinHttpResourceStressTests.bas` and
  `tools/Run-ResourceStressTests.ps1`.
- Evidence: `benchmarks/schema/resource-stress-result.schema.json` and
  `benchmarks/results/phase9-resource-stress.json`.
- Existing comparison: `Run-DownloadStressTests.ps1`,
  `Run-UploadStressTests.ps1`, and `WinHttpNativeTransportTests` handle checks.
- Current contract: `docs/specs/resource-stress.md` and
  `docs/specs/benchmark-methodology.md`.

## Supersedes

- None

## Superseded by

- None
