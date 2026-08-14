<div align="center">
  <img src="docs/images/logo.png" width="160" alt="VBA-HTTP ロゴ" />
</div>

<h1 align="center">VBA-HTTP</h1>

<p align="center">
  <strong>Windows環境のExcel/VBA向け高性能HTTPクライアント</strong>
</p>

<p align="center">
  並行リクエスト対応 · WinHTTPネイティブ実装 · メモリ固定ストリーミング · リトライ機能 · HTTP/2（ホスト依存）
</p>

<p align="center">
  <a href="README.md">English</a>
  |
  <a href="README.ja.md">日本語</a>
</p>

---

VBA-HTTPは、外部ランタイムやコンパイル済みアプリケーションへの依存なしに、
現代的なHTTPクライアント機能をExcel/VBA環境に提供します。
WinHTTP COM経由の通常のバッファリングリクエストに対応するほか、オプションで
ネイティブなWinHTTPバックエンドを選択することで、帯域制限付きストリーミング、プロトコル制御、
高度なトランスポート機能が利用可能になります。HTTP/2は普遍的なネットワーク保証ではなく、
ホストおよびWinHTTP依存のオプショナル機能として実装されています。

```vb
Dim client As HttpClient
Dim response As HttpResponse

Set client = VBAHttp.CreateClient()
Set response = client.GetResponse("https://example.com/api")

response.RaiseForStatus
Debug.Print response.Text
```

単一リクエストを超えるワークロードに対応するため、VBA-HTTPでは以下の特徴を提供します：
・帯域制限付き並行処理
・リトライ機能
・タイムアウト設定
・キャンセル機能
・ファイル転送ストリーミング
・認証機構
・プロキシサポート
・クッキー管理
・診断情報取得機能
・ホスト依存のオプショナルなHTTP/2プロトコルレポート機能

## なぜVBA-HTTPを選ぶのか？

多くのVBA HTTPコードは最終的に、`WinHttp.WinHttpRequest.5.1`を薄いラッパーで包んだだけのものになりがちです。
しかしVBA-HTTPはさらに進化しています。

### 並行リクエスト処理

複数の独立したリクエストを同時に実行可能で、同時に実行する数に明確な上限を設けられます。

```vb
Dim urls As New Collection
Dim options As New HttpBatchOptions
Dim result As HttpBatchResult

urls.Add "https://example.com/a"
urls.Add "https://example.com/b"
urls.Add "https://example.com/c"

options.MaxConcurrency = 8

Set result = client.GetMany(urls, options)

Debug.Print result.SuccessCount
Debug.Print result.FailureCount
```

現在実施されている決定論的ループバックベンチマークでは：

```text
100 requests × 100 ms server delay

Sequential       11.04 s
Concurrency 16    0.86 s

12.86× faster
```

並列度は放任型ではなく、明確に制限されており、各アイテムはタイムアウトやトランスポートエラー時も安定したエラーカテゴリを保持したまま、成功/失敗/キャンセル結果を維持します。リトライ処理は個別のバッチステータスではなく、内部的な試行として扱われます。

### VBAでバッファリングせずに大容量ファイルをストリーミング転送

ネイティブなトランスポート機構により、ファイル全体をVBAの`Byte()`配列として実体化することなく、段階的にデータを転送できます。

```vb
Dim client As HttpClient
Dim download As HttpDownloadResult

Set client = VBAHttp.CreateNativeClient()

Set download = client.DownloadFile( _
    "https://example.com/large.bin", _
    "C:\Temp\large.bin")

download.RaiseForStatus
Debug.Print download.BytesWritten
```

記録されたフェーズ6ベースラインテストでは、1GiBサイズのダウンロードが約**19MBのピークプライベートメモリ使用量増加**で完了しました。この値はベンチマーク対象のExcelプロセスにおける単一のx64環境計測値であり、単なるメモリ予算や最適化後の保証値ではありません。事前/事後比較データは別途個別に追跡されています。

ダウンロード処理は一時的ファイルに一時的に書き込まれ、転送が正常に完了した後でのみ原子的に公開される。既存の宛先ファイルは、HTTP通信エラーやキャンセル、タイムアウト、または書き込みエラーが発生した場合でもそのまま保持される。
ストリーミングアップロードとマルチパートアップロードでは、同じ増分モデルが採用されている：

```vb
Set result = client.UploadFile( _
    "https://example.com/upload", _
    "C:\Temp\payload.bin")
```

```vb
Dim form As HttpMultipartForm

Set form = VBAHttp.CreateMultipartForm()
form.AddField "title", "example"
form.AddFile "payload", "C:\Temp\payload.bin"

Set result = client.UploadMultipart( _
    "https://example.com/upload", _
    form)
```

大容量ペイロードを単一の巨大なVBA文字列やバイト配列として表現する必要はない。

### 必要な場合にはネイティブWinHTTPを使用

デフォルトクライアントは、遅延バインド方式でWindows WinHTTP COMインターフェースを利用している：

```vb
Set client = VBAHttp.CreateClient()
```

ストリーミング転送や高度な通信制御が必要な場合には：

```vb
Set client = VBAHttp.CreateNativeClient()
```

ネイティブバックエンドは直接、ドキュメント化されたWindowsの`winhttp.dll` APIをVBAから呼び出し、以下の機能を提供する：

- ストリーミングダウンロードおよびアップロード
- ホスト固有のHTTP/2プロトコル制御とプロトコルレポート機能
- 応答データのデ圧縮処理
- ネイティブプロキシ設定管理
- 境界付きバッファリングによるWindows/サーバー/プロキシ認証チャレンジへの対応
- 決定論的なWinHTTPハンドル所有権管理とクリーンアップ機構
  追加のカスタムDLLは不要である。

### 信頼性が組み込まれている

リトライ機能は、`Send`関数を無条件に囲むループとして実装されているわけではない。
VBA-HTTPは以下の点を適切に理解している：

- 冪等性を持つメソッドとそうでないメソッドの区別
- `Retry-After`ヘッダーの扱い方
- 指数バックオフアルゴリズム
- ジッター処理
- 最大試行回数設定
- リクエスト期限管理
- 操作全体のタイムアウト設定
- 協調型キャンセル機構
- 一時的DNSエラー、接続エラー、タイムアウト、I/Oエラーなどのトランジェント障害への対応
  デフォルトでは、`GET`、`HEAD`、`PUT`、`DELETE`といったリトライセーフなメソッドについては、一時的な障害が発生した場合に再試行が可能であり、特定のHTTPステータスコードを選択的に対象とすることができる。
  `POST`、`PATCH`、およびカスタムメソッドの場合は、明示的なオプトイン設定が必要となる。

### セキュリティ重視の動作は明確に定義されている

VBA-HTTPは、リクエスト間で機密状態を暗黙的に引き継ぐことを回避している。具体的には以下の機能を備えている：

- OSレベルの証明書検証機構
- Basic認証方式
- Bearer認証方式
- Windows認証チャレンジへの対応
- 境界付きバッファリングによるプロキシ認証チャレンジ管理
- リダイレクト時のプロトコルダウングレード防止機能
- リダイレクト処理中における機密ヘッダー情報の抑制機能
- オプションとしてコールヤーが管理するクッキージャー機能
- 資格情報やクッキー情報をマスキングした診断ログ出力機能
  クッキーや診断データは、隠蔽されたグローバル状態ではなく、明示的なオブジェクトとして管理される。

## パフォーマンスについて

VBA-HTTPには、決定論的なローカルGo HTTPサーバーを基盤とした再現性のあるベンチマーク結果が付属している。
Windows 11/Excel x64環境で記録されたベースラインのハイライトは以下の通りである：

| 負荷条件                                         |        結果 |
| ------------------------------------------------ | ----------: |
| 100リクエスト×100ms（順次処理）                  |     11.04秒 |
| 100リクエスト×100ms（同時実行数16）              |      0.86秒 |
| 並行処理による速度向上率                         | **12.86倍** |
| 1GiBのストリーミングダウンロード                 |  **44.9秒** |
| ダウンロード時のプライベートメモリ使用量ピーク値 |  **約19MB** |
| 1GiBのストリーミングアップロード                 |  **16.0秒** |

1GiB規模のダウンロード/アップロード値は、第6フェーズおよび第7フェーズにおけるベースライン測定結果であり、最新のネイティブホットパス最適化が適用される前に記録されたものである。現在のリポジトリでは、複数回にわたるPIDスコープ付きx64環境下での事前・事後計測データを保存するまで、最適化後のスループットやメモリ使用量に関する新たな数値は公表していない。過去のデータとの比較においては、VBA-HTTPが0.833msであったのに対し、Raw WinHttpRequestでは0.499msという結果が得られており、ラッパー層によるオーバーヘッドは明らかにされている。
これらの測定値は特定の制御環境下における単一マシンでの計測結果であり、普遍的な性能指標として提示するものではない。
ベンチマークスイートでは、機械可読形式の計測結果を記録するとともに、転送内容の検証、同時実行数の制約条件、リソースクリーンアップ処理、およびペイロードハッシュ値の確認を行っている。
完全なエビデンスについては、[`benchmarks/`](benchmarks/)ディレクトリおよびベンチマーク実施方法に関するドキュメントを参照されたい。

## 機能

| 領域             | 対応機能                                                                                                                        |
| ---------------- | ------------------------------------------------------------------------------------------------------------------------------- |
| リクエスト       | GET/POST/PUT/PATCH/DELETEメソッド、クエリパラメータ、ヘッダ情報、テキスト形式・バイナリ形式のボディデータ                       |
| レスポンス       | ステータスコード、ヘッダ情報、バイト単位データ、デコード済みテキスト、交渉されたプロトコルバージョン                            |
| 同時実行制御     | 順序付けされた制限付きバッチ処理と、設定可能な最大並行実行数                                                                    |
| 信頼性機能       | リトライ機構、`Retry-After`ヘッダーによる再試行間隔指定、バックオフアルゴリズム、ジッター調整、デッドライン管理、キャンセル処理 |
| ダウンロード機能 | 定数メモリ使用量を維持するストリーミング方式、および原子的な公開メカニズム                                                      |
| アップロード機能 | ファイル形式およびマルチパート形式でのストリーミングアップロード                                                                |
| プロトコル対応   | ネイティブWinHTTP HTTP/2オプション実装とプロトコル情報報告（ホスト環境依存）                                                    |
| 圧縮機能         | ネイティブgzip/deflateフォーマットによる応答データの自動解凍処理                                                                |
| 認証機構         | 基本認証、Bearerトークン方式、制限付きバッファリングを伴うサーバー/プロキシ向けチャレンジ認証                                   |
| プロキシ対応     | システム標準設定・デフォルトモード、ダイレクト接続モード、および手動プロキシ設定モード                                          |
| クッキー管理     | オプションとしてコール側が独自に管理するクッキージャー機能                                                                      |
| 診断機能         | 機密データのマスキング処理が施された制限付き構造化イベントログ                                                                  |
| テスト環境       | 決定論的な統合テスト、ストレステスト、リソース計測、およびベンチマーク用フィクスチャ                                            |

## インストール方法

主要な配布形態はソースコードパッケージであり、ワークブックにインポートする前に内容を精査することが可能である。
ダウンロード手順:

```text
VBA-HTTP-vX.Y.Z-source.zip
```

GitHubリリースページから入手し、解凍後、**閉じられた状態の**ワークブック環境へインストールすること。

```powershell
powershell -NoProfile -ExecutionPolicy Bypass `
  -File .\Install-VBAHttp.ps1 `
  -Workbook C:\Work\Consumer.xlsm
```

インストーラーの機能：

- パッケージマニフェストの検証
- コンポーネントハッシュの照合
- ワークブックバックアップの作成
- 本番モジュールのみのインポート処理
- `-Force`オプションによる意図的なアップグレード対応
  アップグレード、アンインストール、ロールバック、およびチェックサム手順については、[配布ガイド](docs/guides/distribution.md)を参照してください。

## クイック使用例

### GETリクエスト

```vb
Set response = client.GetResponse("https://example.com/api")
response.RaiseForStatus

Debug.Print response.StatusCode
Debug.Print response.Text
```

### クエリパラメータの指定方法

```vb
Dim request As HttpRequest
Dim response As HttpResponse

Set request = VBAHttp.CreateRequest()
request.Method = "GET"
request.Url = "https://example.com/items"
request.Query.Add "page", 1
request.Query.Add "limit", 100

Set response = client.Execute(request)
response.RaiseForStatus
```

### カスタムリクエストの送信例

```vb
Dim request As HttpRequest

Set request = VBAHttp.CreateRequest()

request.Method = "GET"
request.Url = "https://example.com/api"
request.Headers.SetValue "Accept", "application/json"

Set response = client.Execute(request)
```

### リトライポリシー設定

```vb
Dim policy As HttpRetryPolicy

Set policy = VBAHttp.CreateRetryPolicy()
policy.MaxAttempts = 4

Set client.RetryPolicy = policy
```

### HTTP/2プロトコル対応

```vb
Dim protocols As HttpProtocolOptions

Set client = VBAHttp.CreateNativeClient()
Set protocols = VBAHttp.CreateProtocolOptions()

protocols.AllowHttp2 = True
protocols.Mode = HttpProtocolAllowFallback

Set client.ProtocolOptions = protocols

Set response = client.GetResponse("https://example.com/")
Debug.Print response.ProtocolUsed
```

## 検査可能な設計思想

VBA-HTTPは意図的に通常のVBAソースコードとして配布されています。
隠れたサービスプロセスや、公開APIの背後にカスタムネットワークDLLが存在することはありません。
ネイティブ実装では、Windows APIドキュメントに準拠したネットワーク処理、リソース管理、ストリーミング機能を採用しつつ、WinHTTPハンドルに対する明確な所有権管理を維持しています。
リポジトリには以下も含まれています：

- 単体テストコード
- Excel統合環境における実運用テストケース
- キャンセル処理およびリソース負荷試験用のテストシナリオ
- HTTP/HTTPS通信のための決定論的フィクスチャ実装
- プロキシ接続と認証機構に関するテスト用フィクスチャ
- 再現性のあるベンチマークデータ
- アーキテクチャ設計判断記録
- 機械可読形式のベンチマーク検証証拠資料
  本プロジェクトの目的は、VBA環境でHTTP機能を実現するだけでなく、その動作を**試験可能・計測可能・説明可能な状態にすること**にあります。

## 互換性情報

現在VBA-HTTPが対象としているプラットフォームは以下の通りです：
| プラットフォーム | 対応状況 |
| --- | --- |
| Windows x64版Office | **正式サポート済み** |
| Windows 32bit版Office | **未検証；コミュニティによる動作確認を歓迎します** |
| macOS版Office | 非対応 |
| HTTP/1.1プロトコル | 対応可 |
| HTTP/2プロトコル | ネイティブ対応（オプション機能） - ホスト環境およびWinHTTP依存、プラットフォーム固有の検証データあり |
| HTTP/3 / QUICプロトコル | **ポリシーにより未サポート** |
標準バッファリングクライアントでは以下を使用します：

```text
WinHttp.WinHttpRequest.5.1
```

一方、ネイティブクライアントはWindowsのWinHTTP機能を直接利用します。
OAuthフロー処理、インタラクティブ認証、PAC/WPAD設定、SOCKS接続、および企業向け信頼済み資格情報環境については、現在の互換性保証対象外となっています。
32bit版Officeについては一部環境で動作する可能性がありますが、公式にサポートされたリリース対象ではありません。貢献者は以下の非プロモーション用診断パスを実行できます：

```powershell
powershell -File tools/Run-OfficeBitnessValidation.ps1 `
  -ExpectedArchitecture X86 -DiagnosticOnly
```

動作確認結果は別途個別に審査され、x64版リリース境界に影響を与えるものではありません。
詳細は[互換性ガイド](docs/guides/compatibility.md)および[互換性マトリックス](docs/specs/compatibility-matrix.md)を参照してください。

## ドキュメント案内

まずは以下のページからご覧ください：

- [ガイド目次](docs/guides/README.md)
- [入門ガイド](docs/guides/getting-started.md)
- [APIリファレンス](docs/guides/api-reference.md)
- [リクエストとレスポンスの処理](docs/guides/requests-and-responses.md)
- [信頼性とバッチ処理](docs/guides/reliability-and-batches.md)
- [ストリーミング機能](docs/guides/streaming.md)
- [トランスポート機能](docs/guides/transport-capabilities.md)
- [セキュリティと状態管理](docs/guides/security-and-state.md)
- [分散処理アーキテクチャ](docs/guides/distribution.md)
- [使用例集](docs/guides/examples.md)
- [互換性情報](docs/guides/compatibility.md)
  規範的な契約仕様は[`docs/specs/`](docs/specs/README.md)以下に配置されています。
  アーキテクチャに関する決定事項や実装の根拠となる文書は
  [`docs/adr/`](docs/adr/README.md)以下に収録されています。

## xlflowによる構築と検証

VBA-HTTPは、現代的なソース管理、テスト、診断機能、コーディング支援エージェントを備えたExcel/VBAプロジェクト開発環境である
[xlflow](https://github.com/harumiWeb/xlflow)の実運用環境におけるストレステストとしても機能しています。
本プロジェクトはVBEを手動で編集するのではなく、ソースコードファイルから直接開発されています。その検証プロセスには、実際のExcelコンパイル、単体テスト、統合テスト、負荷試験、静的解析、決定論的なローカルサービス、および再現性のあるベンチマーク測定が含まれています。
VBA-HTTPは、このワークフローを単なる小さなマクロの範囲を超え、数十ものモジュールとシステムレベルの統合を備えた本格的なネットワークライブラリへと発展させる上で重要な役割を果たしました。

## 開発プロセスについて

本番コードは以下のディレクトリに配置されています：

```text
src/classes
src/modules
src/workbook
```

一般的な検証コマンドは以下の通りです：

```powershell
task check
task test:docs
task testserver:test
task precommit
```

完全なローカル検証ワークフローを実行するには以下を実行してください：

```powershell
task verify
```

直接的に開発環境ワークブックを編集しないでください。ソースツリーが正式なものであり、xlflowが提供するのはソースコードからワークブックへのコンパイルとテスト実行の一連の流れです。
詳細は[`CONTRIBUTING.md`](CONTRIBUTING.md)を参照してください。

## リリース成果物について

主要な消費者向け成果物としてソースZIPファイルがあります。
GitHub Releasesではさらに以下のものが公開されます：

- 本番環境用ソースコードパッケージ
- 本番環境専用のXLSMファイル
- リリースマニフェストおよびパックマニフェスト
- SHA-256チェックサム値
- `LICENSE`ファイル
- `THIRD_PARTY_NOTICES.md`ファイル
  GitHubでホストされるパッケージングではExcelを起動しないため、ステータスとして**VBE検証未実施**と記録されます。このパック用XLSMはVBEによる検証を受けた成果物ではありません。
  ローカルリリースゲートでは、本番環境でのリリースが承認される前に、実際のExcel/VBEコンパイル検証が行われます。

[レガシーAPI簡易リファレンス](docs/API.md)、
[`docs/specs/distribution.md`](docs/specs/distribution.md)、
[`docs/specs/github-release.md`](docs/specs/github-release.md)、および
[`docs/RELEASE_CHECKLIST.md`](docs/RELEASE_CHECKLIST.md) を参照してください。

## ライセンスについて

VBA-HTTPは[MITライセンス](LICENSE)の下で公開されています。
プロジェクトが作成したソースコード、ドキュメント、ツール、サンプルコード、ソースパッケージ、
および生成されたリリース用ワークブックはすべてこのライセンスの対象となります。
再配布を行う場合、著作権表示と使用許諾通知を保持し、
[`THIRD_PARTY_NOTICES.md`](THIRD_PARTY_NOTICES.md)を含める必要があります。
ベンチマークスイートで使用されている固定バージョンのVBA-Webチェックアウトは比較用のみに提供されており、

## 実行時依存関係ではありません。

**VBA-HTTP - 現代的なHTTP処理機能を、VBAとしては異例のレベルで実装したライブラリです。**
