# Specification documents

`docs/specs/` は現在有効なcontract、invariant、compatibility rule、validation requirementだけを保持する。判断理由と却下案はADRへ、作業予定は `tasks/todo.md` へ置く。

- 仕様変更と実装・testは同じchange setで更新する。
- 過去仕様を本文へ蓄積せず、user-visibleな変更履歴はCHANGELOGへ移す。
- public API、error behavior、security boundary、release validationにはauthoritativeなtestまたは検査scriptへの参照を含める。

## Current specifications

- `development-and-release-workflow.md`: source authority、proof loop、release component boundary
- `local-test-server.md`: deterministic integration server contract
- `benchmark-methodology.md`: benchmark conditions、provenance、result evidence
- `http-core-api.md`: Phase 1 public domain、transport、body、decoding、error contract
- `buffered-com-transport.md`: default WinHTTP COM request、response、redirect、failure contract
- `bounded-concurrency.md`: cooperative COM async scheduling、batch results、deadlines、cancellation、reentrancy contract
- `reliability-policy.md`: retry classification、backoff／jitter、Retry-After、total deadline、cancellation precedence
- `streaming-download.md`: native constant-memory download、atomic publication、progress、cancellation、cleanup contract
- `streaming-upload.md`: native constant-memory file／multipart upload、source ownership、progress、cancellation、challenge contract
- `protocol-policy.md`: native HTTP/2／HTTP/3 opt-in、fallback／required behavior、transport boundary、compatibility evidence
- `decompression-policy.md`: native gzip／deflate response decoding, header ownership, fallback, and streaming length contract
- `compatibility-matrix.md`: observed Windows/Office/protocol evidence and promotion requirements
