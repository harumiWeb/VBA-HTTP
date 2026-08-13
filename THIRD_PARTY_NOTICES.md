# Third-party notices

The production VBA-HTTP source and filtered release workbooks do not bundle
the benchmark comparator or external platform binaries.

## VBA-Web benchmark comparator

- Project: [VBA-Web](https://github.com/VBA-tools/VBA-Web)
- Pinned revision: `cefc320acc5372e0b86eed1d20eb3f31b331d598` (v4.1.6)
- Usage: ignored, pristine benchmark input only; upstream external-network
  specs are not imported or executed.
- License: MIT, retained in the pinned checkout and verified by
  `tools/Setup-VBAWeb.ps1`.

The comparator is not a VBA-HTTP dependency and is never included in a release
workbook. This notice records provenance; it does not copy or relicense the
upstream source.

## Platform and build services

Windows, Microsoft Excel, WinHTTP, Go, xlflow, Task, and Lefthook are used as
platform or development services. They are not redistributed by the release
workbook and remain subject to their respective terms.

The deterministic Go test server currently uses only the Go standard library;
`tools/testserver/go.mod` declares no external module requirements.
