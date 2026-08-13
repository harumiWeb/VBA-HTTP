# VBA-HTTP Roadmap

> VBAのHTTP runtimeを2026年基準で再設計し、xlflowで開発・検証・配布する。

この文書はVBA-HTTPの実行ロードマップであり、進捗管理の正本とする。
`docs/design.md` は構想を示す非規範のローカル資料として扱い、確定した判断理由は
`docs/adr/`、現行仕様は `docs/specs/`、実装状況はこの文書へ記録する。

## Status Legend

- `[ ]` 未着手
- `[~]` 進行中
- `[x]` 完了
- `[!]` blocked（理由と解除条件を同じ項目に記載する）

## Roadmap Rules

- 1 Todoは原則として、1つのarchitectural concern、1つのtest family、1つの測定可能な完了条件に限定する。
- クラスは `src/classes/`、標準モジュールは `src/modules/`、テストは `src/modules/Tests/` に配置する。
- benchmark用VBAは `src/modules/Benchmarks/`、一時的な開発支援コードは `src/modules/Dev/` に隔離する。
- `docs/design.md` にあるフラットな `src/` やルート `tests/` の構成案は採用しない。
- 外部networkをunit test・integration test・通常benchmarkの依存先にしない。
- test数は品質指標として追跡するが、件数だけをphaseのexit gateにしない。
- bug fixにはregression testを追加する。
- performance最適化には同一条件のbefore／after benchmarkを添付する。
- public APIを変更する場合は、実装前に該当するADRまたはspecを更新する。
- native handleは作成、所有権移譲、closeの状態を追跡できる設計にする。
- WinHTTP native callbackからVBA application logicを直接呼び出さない。
- Authorization、Cookie、proxy passwordなどのsecretをlogへ出力しない。

## Definition of Done

すべての実装Todoは、該当する項目を満たすまで完了にしない。

- [ ] 必要なADRまたはspecを実装前に作成・更新した。
- [ ] focused testを追加または更新し、成功を確認した。
- [ ] 影響範囲のtest suiteを実行し、成功を確認した。
- [ ] `xlflow lint --json` が成功した。
- [ ] `xlflow analyze --json` が成功した。
- [ ] format checkが成功した。
- [ ] sourceをxlflowで開発用workbookへpushし、VBE compileが成功した。
- [ ] 実Excel上のtest、macro、またはworkbook inspectionで要求動作を証明した。
- [ ] verified workbookを保存し、sourceとの差分がないことを確認した。
- [ ] 性能変更では再現可能なbefore／after benchmarkを保存した。
- [ ] public API変更ではREADME、spec、CHANGELOGを更新した。
- [ ] release対象milestoneではRelease Build Gateを通過した。

## xlflow Development Loop

通常のVBA変更では、sourceを正本として次のproof loopを使用する。

```text
xlflow status --json
  -> xlflow session start/attach --json
  -> edit source
  -> xlflow lint --json
  -> xlflow analyze --json
  -> xlflow push --fast --session --no-save --json
  -> focused test / diagnostic macro
  -> affected test suite
  -> xlflow save --session --json
  -> xlflow session stop --json
```

recovery-required状態では通常操作やsaveを続行せず、xlflowのrecovery workflowへ移る。

## Release Build Policy

開発用workbookにはproduction source、test、benchmark、xlflow支援コードを含める。
配布用workbookは開発用workbookを手作業で加工せず、必ず `xlflow build` で別artifactとして生成する。

Phase -1で `xlflow.toml` に次のfilterを追加する。

```toml
[build]
exclude = [
  "src/modules/Tests/**",
  "src/modules/Benchmarks/**",
  "src/modules/Xlflow/**",
  "src/modules/Dev/**",
]
```

Release buildの原則：

- 製品コードからexcluded componentへの依存を禁止する。
- `xlflow build --dry-run --json` でinclude／exclude planを検査してから実buildする。
- baseは追跡された開発用 `build/VBA-HTTP.xlsm` とする。
- default出力は `build/Release/VBA-HTTP.xlsm` とする。
- companion build manifestをrelease構成の機械可読な証跡として保存する。
- manifestのincluded／excluded componentをallowlistと照合する。
- release workbookにtest、benchmark、xlflow helper、dev moduleが存在しないことをinspectする。
- VBE compile、save、close、atomic publicationの成功を必須とする。
- runtimeはrelease workbook内へtest codeを戻さず、外部consumer smoke harnessからpublic APIを呼び出して検証する。
- `pack` はexperimentalでVBE compileを検証しないため、正式なrelease pipelineでは使用しない。
- 将来の `.xlam` は同じ拡張子のbaseを用意し、独立したbuild targetとして生成する。

### Release Build Gate

- [x] 開発用workbookに全sourceをpushし、full test suiteが成功した。
- [x] `xlflow build --dry-run --json` にunmatched exclude warningがない。
- [x] dry-runのincluded componentがproduction allowlistと一致した。
- [x] dry-runでTests、Benchmarks、Xlflow、Dev配下がすべてexcludedになった。
- [x] `xlflow build --json` がVBE compile、save、closeに成功した。
- [x] build manifestがartifactと共に生成された。
- [x] release workbookのcomponent inspectionでdevelopment-only codeが存在しないことを確認した。
- [x] 外部consumer smoke harnessがrelease artifactのpublic APIを実行できた。
- [x] artifactとmanifestのchecksumを保存するsidecar生成・検証を自動化した。

## Milestones

| Milestone | 内容 | 公開方針 |
| --- | --- | --- |
| v0.1 | Core domain model＋basic buffered HTTP | 内部checkpoint |
| v0.2 | Bounded concurrency | 内部checkpoint |
| v0.3 | Retry／timeout／cancellation | 最初の利用可能release |
| v0.4 | Native WinHTTP foundation | Preview |
| v0.5 | Constant-memory download | Public milestone |
| v0.6 | Streaming upload／multipart | Public milestone |
| v0.7 | HTTP/2／advanced protocol | Preview |
| v0.8 | Proxy／auth／cookies | Preview |
| v0.9 | Hardening／benchmarks／docs | Release candidate |
| v1.0 | Stable API | Stable release |

---

## Phase -1 — Repository & xlflow Bootstrap

### Goal

fresh cloneから同じ開発・test・release buildを再現できるclean baselineを確立する。

### Repository and source authority

- [x] `.gitignore` を更新し、`build/VBA-HTTP.xlsm` を開発用workbookとしてGit追跡対象にする。
- [x] `build/Release/**`、一時staging、generated manifestは通常の開発差分へ混入しないignore規則を定義する。
- [x] VBA sourceを編集の正本、開発用xlsmを同期・実行対象とする規約をspec化する。
- [x] VBEでのproduction VBA直接編集を禁止し、例外時のpull／reconcile手順を文書化する。
- [x] fresh cloneで `xlflow status --json` が設定済みworkbookを解決できることを確認する。
- [x] sourceと追跡workbookの初期差分を解消する。

### Clean xlflow baseline

- [x] `src/modules/Tests/SampleTests.bas` のsample／TODO testを製品向けbootstrap smoke testへ置き換える。
- [x] 現在format checkで報告されるscaffold 7ファイルを正規化する。
- [x] `XlflowDebug` の既存analyzer finding 5件を、修正または根拠付き局所抑制で解消する。
- [x] `XlflowUI` のunused inline suppression warningを解消する。
- [x] lint、analyze、format check、testのclean baseline結果を記録する。
- [x] analyzer／lintのproject-wide suppressionは追加せず、意図した境界だけを根拠付き局所抑制にする。

### Test and benchmark conventions

- [x] unit testを `src/modules/Tests/Unit/` に配置する規約を定義する。
- [x] integration testを `src/modules/Tests/Integration/` に配置し、`integration` tagを付ける規約を定義する。
- [x] stress testを `src/modules/Tests/Stress/` に配置し、通常suiteから分離する規約を定義する。
- [x] VBA benchmark moduleを `src/modules/Benchmarks/` に配置する規約を定義する。
- [x] deterministic testで使用するclock、random、transportの差し替え方針をspec化する。
- [x] x64 Officeのsmoke test証跡を取得する。
- [x] 32-bit Officeのcompile／smoke testを行う環境と証跡保存方法を文書化する。

### Release build foundation

- [x] `xlflow.toml` に `[build].exclude` を追加する。
- [x] `src/modules/Benchmarks/` と `src/modules/Dev/` のdirectory規約を作成する。
- [x] production sourceからTests、Benchmarks、Xlflow、Dev componentへの依存を禁止するspecを作成する。
- [x] `xlflow build --dry-run --json` のmanifestを検査するscriptまたは自動checkを作成する。
- [x] exclude globが1件もmatchしない場合をwarningとして失敗させるrelease checkを定義する。
- [x] production component allowlistとdry-run結果を照合する。
- [x] `xlflow build --json` で `build/Release/VBA-HTTP.xlsm` を生成する。
- [x] build manifestからVBE compile、save、close、publication結果を検証する。
- [x] release artifactのVBA component一覧をinspectする。
- [x] release artifactにexcluded componentが存在しないことを自動検証する。
- [x] release artifactを呼び出す外部consumer smoke harnessの構成を決定・作成する。
- [x] consumer smoke harnessがrelease artifactの最小public APIを実行できることを確認する。

### Governance and contributor workflow

- [x] `docs/adr/` の命名、status、template、supersede規約を定義する。
- [x] `docs/specs/` に現行仕様だけを置く規約を定義する。
- [x] benchmark結果のJSON schemaと保存先を定義する。
- [x] contributor向けに標準xlflow proof loopを文書化する。
- [x] contributor向けにrelease build loopを文書化する。
- [x] public API、bug fix、performance変更ごとのdocumentation gateをチェックリスト化する。

### Exit Criteria

- [x] fresh cloneで追跡workbookを使って `xlflow test --json` を実行できる。
- [x] lint、analyze、format checkがcleanである。
- [x] sample testとTODO testが残っていない。
- [x] sourceと開発用workbookに意図しない差分がない。
- [x] build dry-runが意図したproduction componentだけをincludeする。
- [x] release buildがVBE compileに成功する。
- [x] release artifactにexcluded componentが存在しない。
- [x] 外部consumer smoke harnessがrelease artifactに対して成功する。
- [x] x64の実行証跡と32-bitの検証経路が存在する。

---

## Phase 0 — Deterministic Test & Benchmark Infrastructure

### Goal

実装前に公平で再現可能な比較基準を固定し、以降のHTTP機能を外部networkなしで証明できるようにする。

### Todo

- [x] `tools/testserver/` にGo製local HTTP serverを作成する。
- [x] test serverのport選択、startup readiness、shutdown、state resetを自動化する。
- [x] `GET /status/{code}` を実装する。
- [x] `GET /delay/{milliseconds}` を実装する。
- [x] `GET /bytes/{size}` と `GET /stream/{size}` を実装する。
- [x] `GET /headers` と `POST /echo` を実装する。
- [x] `GET /redirect/{count}` を実装する。
- [x] `GET /flaky/{failCount}` と `GET /rate-limit/{count}` を実装する。
- [x] fixture、hash検証、大容量payload生成を整備する。
- [x] Excel process memoryとnative handleを測定するutilityを整備する。
- [x] Raw WinHttpRequestのbenchmark harnessを作成する。
- [x] VBA-Webのupstream commitを固定したsetup手順を作成する。
- [x] VBA-Webのbenchmarkをlocal serverへ向け、upstreamの外部network specsを実行対象から除外する。
- [x] warmup、試行回数、payload、timeout、集計方法を `docs/specs/` に記録する。
- [x] benchmark結果の機械可読JSON schemaを実装する。
- [x] baselineを `docs/BENCHMARKS_BASELINE.md` とJSONへ保存する。

### Exit Criteria

- [x] test／benchmarkが外部networkなしでdeterministicに再実行できる。
- [x] Raw WinHttpRequestとVBA-Webのsequential／latency比較を取得できる。
- [x] 100MB以上のdownload benchmarkを実行できる。
- [x] benchmark条件、環境、失敗時診断が記録される。
- [x] release buildからbenchmark moduleが除外される。

---

## Phase 1 — Core Domain Model

### Goal

network実装から独立したpublic APIとdomain modelを確立する。

### Todo

- [x] dual transportとtransport abstractionのADRを作成する。
- [x] public API shapeとcompatibility policyのspecを作成する。
- [x] error modelのADRを作成する。
- [x] `IHttpTransport` とmock transportを実装する。
- [x] `HttpClient`、`HttpRequest`、`HttpResponse` を実装する。
- [x] `HttpHeaders`、`HttpParams`、`HttpTimeouts` を実装する。
- [x] HTTP method、URL encoding、query constructionを実装する。
- [x] text／binary body表現とownershipを定義する。
- [x] response text decodingとcharset failure policyを定義する。
- [x] error category、VBA error number、response failureの境界を実装する。
- [x] network-free unit test matrixを実装する。

### Exit Criteria

- [x] networkアクセス0で全domain testが成功する。
- [x] requestからtransport、response／errorまでmockで検証できる。
- [x] public API例とcompatibility policyがspec化されている。
- [x] production componentだけのrelease buildが成功する。

---

## Phase 2 — Buffered COM Transport

### Goal

dependency-freeで扱いやすい基本HTTP機能をWinHTTP COM transportで提供する。

### Todo

- [x] late-bound `WinHttp.WinHttpRequest.5.1` transportを実装する。
- [x] GET、POST、PUT、PATCH、DELETEを実装する。
- [x] request headers、query、text body、binary bodyを実装する。
- [x] resolve、connect、send、receive timeoutを適用する。
- [x] 基本redirect policyを実装する。
- [x] status、headers、text／binary response bodyを取得する。
- [x] COM／WinHTTP failureを統一error modelへ変換する。
- [x] local server integration testを追加する。
- [x] Raw WinHttpRequestに対するlibrary overheadを測定する。

### Exit Criteria

- [x] 基本HTTP操作がlocal serverで成功する。
- [x] transport failureが統一error modelへ変換される。
- [x] 小規模requestのoverheadがbaseline比15%以内、または差異の根拠が記録されている。
- [x] built artifactを使うconsumer GET smoke testが成功する。

---

## Phase 3 — Bounded Concurrency

### Goal

VBAをmulti-thread化せず、COM async I/Oをbounded schedulerで並行処理する。

### Todo

- [x] concurrency semanticsとreentrancy policyをspec化する。
- [x] COM async request lifecycleを実装する。
- [x] bounded schedulerとcompletion pollingを実装する。
- [x] `ExecuteMany` と `GetMany` を実装する。
- [x] `HttpBatchResult` と `HttpBatchItem` を実装する。
- [x] per-request deadlineとbatch cancellationを実装する。
- [x] controlled message pumpとyield intervalを実装する。
- [x] 同一clientからのreentrant executionを拒否する。
- [x] 個別failureを保持してbatch継続できる結果モデルを実装する。
- [x] fairness、result order、timeout、cancel、reentrancy regression testsを追加する。

### Exit Criteria

- [x] 100 request × 100ms、concurrency 16でsequential比6倍以上を達成する。
- [x] 設定した上限を超えるin-flight requestが存在しない。
- [x] partial failureとcancellation結果が決定的である。
- [x] built artifactからconcurrent consumer smoke testが成功する。

---

## Phase 4 — Retry, Timeout & Cancellation

### Milestone

最初の利用可能release `v0.3`。

### Goal

retry、deadline、cancellationを統一した信頼性policyとして提供する。

### Todo

- [x] retry semanticsのADRを作成する。
- [x] exponential backoffと上限delayを実装する。
- [x] deterministicに差し替え可能なjitter sourceを実装する。
- [x] `Retry-After` delta-secondsとHTTP-dateを実装する。
- [x] 408、429、500、502、503、504のdefault retry判定を実装する。
- [x] retry可能なnetwork error分類を実装する。
- [x] idempotent methodだけをdefault retryする。
- [x] non-idempotent methodの明示opt-inを実装する。
- [x] per-attempt timeoutとtotal deadlineの優先順位を実装する。
- [x] polling、retry wait、送受信checkpointのcancellationを実装する。
- [x] flaky／rate-limit endpointを使うdeterministic policy testsを追加する。
- [x] retry／timeout／cancelの利用例を追加する。

### Exit Criteria

- [x] 408、429、500、502、503、504、timeout、接続失敗を決定的に検証できる。
- [x] POST／PATCHは明示opt-inなしでretryされない。
- [x] total deadlineとcancelがretry loopを確実に停止する。
- [x] clean release buildとconsumer reliability smoke testが成功する。
- [x] Core＋COM＋concurrency＋reliabilityを備えたv0.3を配布できる。

---

## Phase 5 — Native WinHTTP Foundation

### Goal

streamingとadvanced protocolの土台となる安全なNative WinHTTP transportを構築する。

### Todo

- [x] native callbackからVBA application logicを呼ばないADRを作成する。
- [x] native handle ownershipとcleanup policyをspec化する。
- [x] 32-bit／64-bit対応WinHTTP declarationを実装する。
- [x] session、connection、request handle wrapperを実装する。
- [x] buffered GET、request headers、response headersを実装する。
- [x] protocol query、TLS validation、error mappingを実装する。
- [x] repeated requestのhandle計測とleak regression testsを追加する。
- [x] COM transportとのcontract suiteを追加する。

### Exit Criteria

- [x] x64 compile evidenceがある（`benchmarks/results/office-bitness-x64.json`で135 unit／71 integration、VBE compile、consumer smokeを実証）。
- [~] 32-bit Office compile evidenceを取得する（`Run-OfficeBitnessValidation.ps1 -ExpectedArchitecture X86`の実機実行待ち）。
- [x] repeated request後にpersistent handle growthがない。
- [x] COM transportと同じpublic response／error contractを満たす。
- [x] production-only release buildが成功する（`xlflow build` のVBE compileと外部consumer smokeを確認済み）。

---

## Phase 6 — Constant-Memory Download

### Goal

巨大responseをmemoryへ全展開せず、安全にfileへstreamingする。

### Todo

- [x] streaming ownership／failure semanticsのADRまたはspecを作成する。
- [x] `WinHttpQueryDataAvailable`／`WinHttpReadData` loopを実装する。
- [x] temporary fileへのchunked writeとflushを実装する。
- [x] 成功時のatomic replacementを実装する。
- [x] failure／cancel時のpartial file cleanupを実装する。
- [x] `IHttpProgressSink` とprogress reportingを実装する。
- [x] streaming checkpointでcancellationを検査する。
- [x] 2GB超のfile size／byte count表現を実装する。
- [x] hash、memory、cleanupのintegration／stress testsを追加する。

### Exit Criteria

- [x] 1GB downloadのhashが一致する（loopback stress: SHA-256 `3a33d58aa8ee1e9d21fd4f510cc5d1ce8d25ba5e24363d19c28e2bf866f4185c`）。
- [x] additional working memory 32MB未満をengineering targetとして実測する（working-set peak delta 6,836,224 bytes、private-bytes peak delta 19,017,728 bytes）。
- [x] failure／cancel時に既存destinationを破壊しない。
- [x] partial fileとnative resourceが残らない。
- [x] built artifactからdownload consumer smoke testが成功する（native 1 MiB loopback download、published/file-size検証）。

---

## Phase 7 — Streaming Upload & Multipart

### Goal

fileとmultipart bodyをmemoryへ全展開せず送信する。

### Todo

- [x] `IHttpUploadTransport` と `HttpUploadResult` の公開境界を実装し、COM-only clientのcapability failureをnetwork前に検証する。
- [x] `WinHttpWriteData` を使うfile upload loopを実装する。
- [x] field／fileを逐次送信するmultipart encoderを実装する。
- [x] content lengthとchunk handling policyをADR-0008／specへ記録する。
- [x] upload progressとcancellationを実装する。
- [x] authentication challenge中のupload failure policyをADR-0008／specへ記録する。
- [x] deterministic serverのhash、memory、multipart parsing、cleanup testsを追加する。
- [x] release artifactをtest codeへ戻さずfile／multipart APIを呼ぶexternal consumer smoke harnessを追加する。

### Exit Criteria

- [x] 1GB uploadのserver-side hashが一致する（loopback stress結果を `docs/BENCHMARKS_BASELINE.md` に保存する）。
- [x] multipart uploadの各field／fileが正しく復元される。
- [x] payload全体をByte配列や巨大Stringへ展開しない（64 KiB write/read upper boundをunit testと実装で固定する）。
- [x] cancellation後にsource file／native resourceが残らない。
- [x] built artifactからfile／multipart upload consumer smoke testが成功する。

---

## Phase 8 — Protocol, Proxy & Authentication

### Goal

OSが提供するmodern WinHTTP capabilityとcorporate environment対応を安全に公開する。

### Todo

- [x] protocol enable／fallback／required modeをADR-0009と`docs/specs/protocol-policy.md`で確定する。
- [x] native HTTP/2 opt-inと`HttpResponse.ProtocolUsed`によるnegotiated protocol取得を実装する。
- [x] HTTP/3 opt-inのWinHTTP flagとvalidationを実装する（requested-mask、unsupported fallback、required errorをnative transportで検証済み）。
- [x] HTTP/2／HTTP/3 negotiated-host evidenceのfail-closed収集runner、redacted schema、validatorを追加する（ADR-0025／`tools/Run-ProtocolHostValidation.ps1`）。
- [x] HTTP/2のTLS negotiated-host evidenceをx64実機で取得する（`benchmarks/results/protocol-host-http2.json`、trusted `nghttp2.org:443`、required mode、`ProtocolUsed=HTTP/2`）。
- [~] HTTP/3のTLS/QUIC negotiated-host evidenceを実機で取得する（対応Windows/WinHTTP hostとfixture待ち）。
- [x] unsupported option、plain HTTP、required mismatchのfallback／`HttpErrorProtocol`を実装する。
- [x] gzip／deflate decompressionをADR-0010と`docs/specs/decompression-policy.md`で確定し、native WinHTTP option 118、COM拒否、loopback／release smokeを実装する。
- [x] OS default、no-proxy、manual HTTP proxyをADR-0012と`docs/specs/proxy-policy.md`で確定し、COM/native transportとloopback fixtureで検証する。
- [x] Basic／Bearerのpreemptive authenticationを実装する（COM/native、HTTPS default、loopback opt-in、redirect suppression、release smokeまで検証済み）。
- [x] Windows integrated／Digest／proxy challenge authenticationのbounded buffered契約をADR-0023／`docs/specs/challenge-authentication.md`で確定し、COM/nativeのchallenge replay、redaction、release smokeを実装する。streaming sourceはresetなしでpre-network拒否し、実Windows domain／proxy CONNECT fixtureはcompatibility gateとして残す。
- [x] `IHttpAuthProvider` を実装する（execution snapshotでcloneし、401／407の自動replayは行わない）。
- [x] credential／secret redaction helperを実装し、sensitive headerがdiagnosticsへ出ないunit gateを追加する。
- [x] `HttpDiagnostics` のbounded structured event schemaをADR-0019／`docs/specs/diagnostics-policy.md`で確定し、HttpClientとrelease consumer smokeへ統合する。
- [x] OS version／Office bitness compatibility matrixを作成する（`docs/specs/compatibility-matrix.md`）。
- [~] matrixのprotocol host evidenceと32-bit実測を完了する（x64 HTTP/2は実測済み、HTTP/3とx86 Officeは実機証跡待ち）。

### Exit Criteria

- [x] HTTP/1.1 fallbackとrequested-mask contractを識別できる（plain HTTP、unsupported、required mismatchを検証済み）。
- [~] TLS上のHTTP/2と対応環境のHTTP/3 negotiated protocolを識別する（x64 HTTP/2は`benchmarks/results/protocol-host-http2.json`で実測済み、HTTP/3はTLS/QUIC host evidence待ち）。
- [x] unsupported環境のfallback／errorがspec通りである。
- [x] HTTP forwarding proxyとfixed Basic proxy-challengeのlocal integration testsがCOM/nativeで成功する（HTTPS CONNECTと実Windows domain authはcompatibility gateとして継続）。
- [x] secretがdiagnosticsやbenchmark outputへ出ない（diagnostics serializerはquery／user-info／body／descriptionを除外し、sensitive headerを常時redactする。benchmark serializerは既存契約を維持する）。
- [x] release artifactでprotocol／auth consumer smoke testが成功する（protocol fallbackとBasic／Bearer auth smoke）。

---

## Phase 9 — Production Hardening

### Goal

security、malformed input、長時間実行、resource stabilityをrelease品質まで引き上げる。

### Todo

- [x] cookie jarと最小cookie policyを実装する（caller-owned jar、host/path/secure/expiry、Set-Cookie、redirect boundary、COM/native/release smoke）。
- [x] redirect loopとmaximum redirectを実装する（WinHTTP上限とloopback検証、release consumer smoke）。
- [x] cross-originでAuthorization／Proxy-Authorizationを除去する（native callbackを使わず、credential-bearing automatic redirectを停止して3xxをcallerへ返す）。
- [x] HTTPSからHTTPへのdowngrade redirectをdefault拒否する（COM option 12=false、native policy=DISALLOW_HTTPS_TO_HTTP）。
- [x] Unicode、direct binary、malformed UTF-8 response testsを追加する（COM/native loopbackでresponse boundaryを検証）。
- [x] wire-level malformed response-header fixtureを追加し、COM/nativeの`HttpErrorProtocol` mappingを検証する。
- [x] 10,000 sequential／scheduled request stress testsを追加する（10,000 warmup後の10,000測定区間）。
- [x] repeated cancellation／timeout stress testsを追加する（COM active cancellation／request deadline／receive-timeout、native download cancellationの4 scenario。`task test:cancellation-stress`で25 iteration、recovery 204、PID-scoped handle gate、destination/temp cleanupの証跡を生成する）。
- [x] process memory／native handle leak checksを自動化する（PID scoped、idle handle delta gate、memory peak evidence）。
- [x] security policyとthreat reviewを文書化する（manifest境界、checksum、残余リスクを`docs/adr/ADR-0017-*`、`docs/specs/release-security.md`、`docs/verification/phase-nine-security-review.md`へ記録）。
- [x] release manifestのincluded component一覧をsecurity reviewする（`task release:security`のfail-closed検査と証跡レポート）。

### Exit Criteria

- [x] known critical bugとcurrent release blocker bugが0件である（`docs/security/risk-register.json`を`task test:security-risks`とrelease security gateで検証。将来v1.0 gateは別途明示）。
- [x] 10,000 request測定区間後にpersistent resource growthがない（x64実測、native delta <=8／COM delta <=32）。
- [x] redirect、credential、certificate security testsが成功する（redirect／credential／untrusted TLS certificateのCOM/native loopback検証済み。HTTP/2/3 negotiated evidenceと32-bit証跡はcompatibility gateとして継続）。
- [x] release artifactのcomponent構成がreview済みである（`task release:security`と`task release:smoke`でproduction allowlist／development denylistのmanifest・実Workbook構成を検証）。

---

## Phase 10 — Developer Experience & Distribution

### Goal

利用者とcontributorが導入、利用、検証、配布を再現できる状態にする。

### Todo

- [x] README、API reference、examplesを完成させる（README、`docs/API.md`、`examples/`、`test:docs`契約ゲート）。
- [x] xlflow-based contributor guideを完成させる（`CONTRIBUTING.md`、Lefthookの`task check`ゲート）。
- [x] `xlflow build` によるrelease workbook生成を自動化する（`task release:build`、manifest／VBE／checksum／smoke）。
- [x] build manifest検証とartifact checksum生成を自動化する。
- [x] external consumer smoke harnessをrelease pipelineへ統合する。
- [x] source distributionとworkbook distributionの責務を分離する（`docs/specs/distribution.md`）。
- [x] source vendoring、install、upgrade手順を文書化する。
- [x] `.xlam` 用base workbookと独立build targetを追加する（追跡`build/VBA-HTTP.xlam`、`task test:xlam`、`task release:xlam:build`、ADR-0021、manifest/checksum/add-in identity/smoke gate）。
- [x] CHANGELOGとcompatibility documentationを完成させる（x64 HTTP/2実測と未完了の32-bit／HTTP/3 gateを明記）。

### Exit Criteria

- [x] clean environmentでtest、build、smoke、packageを再現できる（`tools/Test-CleanCheckout.ps1 -FullRelease`でgit archiveからXLSM/XLAM build・checksum・smokeを実証。32-bitはcompatibility gateとして継続）。
- [x] 全public APIに例とerror behaviorの説明がある（`docs/API.md`、README、spec、examples）。
- [x] development-only codeを含まない配布workbookを生成できる。
- [x] build失敗時に既存release artifactが破壊されない（atomic publication／checksum／tamper tests）。
- [x] release manifestとchecksum sidecarが保存される。

---

## Phase 11 — Release Benchmark & v1.0

### Goal

同一条件の公開benchmarkと全quality evidenceを揃え、stable APIをreleaseする。

### Todo

- [x] Phase 0 baselineを同一条件で再測定し、`docs/BENCHMARKS_BASELINE.md`とJSONへ保存する。
- [x] sequential／concurrency benchmarkを公開する。
- [x] 100MB／1GB download benchmarkを公開する。
- [x] multipart upload benchmarkを公開する。
- [x] memory／handle stability結果を公開する。
- [x] x64 Office validationを完了する（`benchmarks/results/office-bitness-x64.json`）。
- [~] 32-bit Office validationを完了する（`Run-OfficeBitnessValidation.ps1 -ExpectedArchitecture X86`の実機実測待ち）。
- [x] release checklistとartifact checksumsを作成する（`docs/RELEASE_CHECKLIST.md`、checksum sidecar gate）。
- [x] READMEには実測値だけを掲載する（loopback条件と未達15% targetを明記）。
- [x] 最終 `xlflow build` とexternal consumer smoke testを実行する（現行x64 release evidence）。

### Exit Criteria

- [x] concurrency、streaming、resource stabilityの実測証跡がある。
- [x] 現行x64のquality gateが成功する（unit、integration、lint、analyze、format、VBE、release smokeを検証済み）。
- [~] v1.0 promotion quality gateを完了する（x64 HTTP/2は実測済み、32-bit／HTTP/3 negotiated／host-specific challenge-auth evidenceは未完了）。
- [x] API、security、compatibility documentationが完成している（現行contract、matrix、security gate、deferred riskを文書化済み）。
- [x] test、benchmark、xlflow支援コードを含まないproduction-only artifactを生成できる（XLSM/XLAM build・manifest・checksum・smoke済み）。
- [~] v1.0 promotion evidenceを完了する（x64 HTTP/2は`benchmarks/results/protocol-host-http2.json`、32-bit／HTTP/3 negotiated／host-specific challenge-auth riskは未解消）。
- [x] build manifestからrelease構成を再現・監査できる。

---

## Architecture Decision Gates

実装前に、少なくとも次の判断をADRまたはspecとして確定する。

- [x] dual transportとtransport selection policy
- [x] public `HttpClient` の同期、batch、download、upload API（Phase 1〜7で実装済み）
- [x] `IHttpTransport` と `IHttpProgressSink` の契約（auth providerとは分離）
- [x] `IHttpAuthProvider` の契約（ADR-0013／`docs/specs/auth-policy.md`で確定）。
- [x] buffered／streaming request・response bodyのownership（buffered bodyはPhase 1、downloadはADR-0007、uploadはADR-0008/specで確定）
- [x] error分類とVBA error／result objectの境界
- [x] retry、deadline、redirect、cancellationの優先順位（ADR-0005／`docs/specs/reliability-policy.md`、cancel > total deadline > attempt outcome > retry/delay。unit/integrationで検証済み）
- [x] native handle lifecycleとnative callback禁止
- [x] protocol fallback policy（ADR-0009、`docs/specs/protocol-policy.md`）
- [x] response decompression ownership、fallback／required、streaming length contract（ADR-0010、`docs/specs/decompression-policy.md`）
- [~] OS compatibility evidence（x64 HTTP/2 host evidenceを追加、HTTP/3と32-bit実測は未完了）
- [x] diagnostics schemaとsecret redaction（ADR-0019／`docs/specs/diagnostics-policy.md`、unit、client、release smokeで検証済み）。
- [x] development-only componentへのproduction依存禁止
- [x] development workbook、release workbook、source distributionの責務

ADRは判断理由とtrade-offを保持し、現行のinterface、validation、compatibility contractは
`docs/specs/` に保持する。既存ADRと重複する場合は新規作成せず、更新またはsupersedeする。

## v1.0 Scope Boundaries

次の機能はv1.0の対象外とする。

- macOS
- WebSocket
- Server-Sent Events
- HTML scraping
- JSON parser内蔵
- OAuth1／OAuth2 flowそのもの
- GraphQL専用client
- API SDK generator
- VBA自体のmulti-threading
- 独自TLS implementation

## Assumptions

- `docs/design.md` はGit管理せず、このTodo単体でroadmapを理解できる内容を維持する。
- `build/VBA-HTTP.xlsm` は追跡するが、VBEでproduction VBAを直接編集しない。
- 開発用workbookには全sourceとtest infrastructureを含める。
- 配布用workbookは常に `xlflow build` から生成し、手作業でmoduleを削除しない。
- `pack` は正式release生成に使用しない。
- VBA-Webはversion固定setupで取得し、製品dependencyにはしない。
- v0.1／v0.2は内部checkpoint、最初の利用可能releaseはv0.3とする。
