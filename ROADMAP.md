# elixcee Roadmap

## Product direction

elixcee は、Excelを起動せずに、データ処理用のVBAとワークブックを安全かつ
再現可能に処理する Rust エンジンです。ClosedXML、openpyxl、SheetJS などの
「ワークブックを操作するライブラリ」と同じ機能数を短期に競うのではなく、次の
組み合わせで選ばれることを目標にします。

1. Excel不要のVBA実行（移行対象のマクロをコード変更なし、または小さな変更で処理）
2. Linux/macOS/Windows、CLI/Python/JavaScript-WASMで同じ結果になること
3. 大きな表を扱うバッチ処理で、メモリ使用量と実行時間を予測できること
4. 未対応OOXMLを黙って壊さず、保持・警告・エラーを明確に選べること
5. Excel/LibreOffice/既存ライブラリとの比較結果を、再実行可能なfixtureとして公開すること

ClosedXMLが提供する広いワークブックAPI（ワークシート、テーブル、スタイル、数式など）と、
openpyxlが提供するPythonからの扱いやすい入出力・編集APIは重要な比較対象です。ただし、
競合の機能表をそのまま受け入れるのではなく、各機能について
「読み書きできる」「意味を評価できる」「実Excelとの往復で壊さない」を分離して判定します。

## 識別子

- `[BUILD]`: 通常の実装・小さなfixture・ユニット/property testで完了判定できる。
- `[MEASURE]`: 実Excelまたは独立oracle、大規模fixture、時間/RSS/成果物サイズなどの
  測定が必要。測定データと環境を保存するまで完了にしない。
- `[GATE]`: 先行phaseの合格が必要。未検証の機能を次の互換性主張に含めない。

重い測定は、機能実装のたびに行わず、各 `[MEASURE]` phase の最後に集約します。
小さなfixtureでの正しさは `[BUILD]` で先に確認し、ベンチマークで正しさを推測しません。

## 現状（0.23系開発ライン）

実装済みの主な範囲は、VBAの基本制御構文・変数/配列・Range/Cells・複数シート・
エラー処理、主要な数式カテゴリ、Python/CLI、XLSX/XLSM/ODS入出力、テーブル・検証・
AutoFilter、JavaScript/WASMの実験的APIです。既存の `compat/` には、SheetJS、
openpyxl、LibreOffice、独立したVBA意味論fixtureとの比較基盤があります。

大きなXLSX/XLSM向けに、前方向のstreaming reader/writerも提供しています。readerは
行番号付き返却、空行、relationship targetの正規化、行バッファ上限を扱います。

ただし、Microsoft Excel本体をoracleにしたVBA実行・保存後再実行、大規模ワークブックの
時間/RSS比較、OOXML part単位の保持保証、1.0の公開契約は未完了です。

## Phase 0 — 競争軸と測定基盤を固定する

### `[BUILD]` 比較対象・サポート階層

- ClosedXML（.NET）、openpyxl（Python）、SheetJS（JavaScript）、LibreOffice、可能な範囲で
  Microsoft Excelを比較対象として登録する。
- API互換ではなく、`read/write`、`formula/VBA evaluation`、`round-trip preservation`、
  `resource safety` の4軸で機能マトリクスを作る。
- 機能ごとに `supported` / `preserved` / `evaluated` / `warned` / `rejected` を記録する。
- 既存の `compat/` の判定語彙を共通化し、`MATCH` と「比較不能」を混同しない。

### `[MEASURE]` 再現可能な性能・互換性ハーネス

- fixtureを小・中・大（例: 1万、10万、100万行相当）に分け、生成条件を固定する。
- wall time、peak RSS、出力サイズ、失敗理由、CPU/OS/ランタイム版をJSONで保存する。
- cold start、warm run、parse-only、execute-only、read/write、full pipelineを分離する。
- 結果は「勝った/負けた」ではなく、中央値・p95・RSS・正確性の組で報告する。
- ここでベースラインを作り、以後の最適化は同じfixtureと同じ測定手順で比較する。

## Phase 1 — 安全なワークブック基盤

### `[BUILD]` 入出力契約と保存ポリシー

- 1-based座標、日付/時刻、エラー値、空セル、共有文字列、数式キャッシュを型として固定する。
- `.xlsx` / `.xlsm` / `.ods` の読み書き境界をAPIドキュメントとエラー型に反映する。
- 未対応OOXMLは、可能なら保持し、無理なら警告または明示エラーにする。
- マクロbytes、relationships、defined names、テーブル、検証、コメント、ハイパーリンク、
  drawing、freeze pane、hidden stateをpart単位のfixtureで回帰化する。
- zip bomb、過大XML、過大セル範囲、循環数式などの資源上限を明示する。

### `[MEASURE]` 大規模I/Oとメモリ上限

- streaming readerと通常readerを同一fixtureで比較し、RSSが入力サイズにどう増えるか測る。
- 破損/悪意あるZIP/XMLを含め、上限到達までの時間、エラーの決定性、部分出力の有無を測る。
- 競合との比較は速度だけでなく、保存後の論理同値と未対応要素の保持率を同時に合格条件にする。

## Phase 2 — Excel数式エンジンの信頼性

### `[BUILD]` 意味論の土台

- A1/RC参照、範囲、名前、相対/絶対参照、依存グラフ、循環検出を完成させる。
- `Variant` の型変換、空値、エラー伝播、比較、日付シリアル、文字列連結を仕様化する。
- SUM/IF/LOOKUP等の既存関数を、関数レジストリ・独立期待値・失敗ケースで管理する。
- 通常・配列・動的配列・共有数式を「保存」と「評価」に分けて実装する。
- Automatic/Manual計算と、再計算範囲の無駄な全走査を検出する。

### `[MEASURE]` 数式の正確性・計算性能

- Excel/LibreOffice/ClosedXML等で同じfixtureを評価し、値・エラー・スピル範囲を比較する。
- 関数カテゴリ別に正答率、未対応率、比較不能率を出す。比較不能は成功に数えない。
- 依存グラフの疎/密、再計算回数、範囲サイズを変え、full recalcとincremental recalcを測る。
- 速度最適化の合格条件は、値の一致を満たしたケースでのみ適用する。

## Phase 3 — VBA実行互換性

### `[BUILD]` データ処理マクロの拡張

- Sub/Function、ByRef、Variant、配列、For/For Each/Do、If/Select Case/Withを安定化する。
- On Error、Resume、Err、型不一致、範囲外アクセスのエラー契約を固定する。
- Range/Cells、複数シート、名前付き範囲、AutoFilter、Sort、テーブル操作を実装する。
- MsgBox/UserForm/COM/ActiveX等は skip / dummy / error をオプションで選択可能にする。
- 静的検査で「実行時に黙って無視される命令」を事前に報告する。

### `[MEASURE]` Excel意味論と実運用マクロ

- Microsoft Excelを利用できるWindows runnerで、入力fixture→VBA実行→保存→再読込を測る。
- 既存の581シナリオを拡張し、実Excel結果を独立oracleとして登録する。
- 値だけでなく、エラー、変更セル集合、数式、書式、マクロbytes、警告を比較する。
- macro実行時間、parse cacheの効果、複数VM並列時のRSSと再現性を測る。
- Excel oracleが使えない環境では、`ORACLE_UNAVAILABLE` として扱い、成功扱いにしない。

## Phase 4 — 大規模バッチと並列実行

### `[BUILD]` スループット設計

- read-only workbook snapshot、parse cache、columnar/range bulk API、streaming writeを設計する。
- VM間の状態分離、キャンセル、タイムアウト、決定的なseed、構造化診断を提供する。
- Python GIL、WASMのメモリ制約、Rustネイティブの並列実行をそれぞれ文書化する。

### `[MEASURE]` 競合を越える性能の立証

- ClosedXML/openpyxl/SheetJS/LibreOfficeと、同じ操作・同じ出力検証で比較する。
- 行数・列数・数式密度・style密度・シート数を振ったマトリクスで測る。
- throughput、p50/p95、peak RSS、CPU時間、出力サイズ、失敗率を保存する。
- 「高速」を名乗る条件を、対象fixture・環境・比較対象・測定日付きで限定する。
- 目標値は測定後に決める。根拠のない倍率目標や、1ケースだけのベンチマークは採用しない。

## Phase 5 — API・配布・開発者体験

### `[BUILD]` 使い始めるまでの摩擦を下げる

- Python 3.9+、CLI、JavaScript/WASMの最短例と、VBA→Python移行例を整備する。
- 型stub、安定したエラーコード、diagnose JSON、ログ抑制、入力検証を公開契約にする。
- Linux/macOS/Windowsのwheel、CLI、npm packageを同じリリース番号で検証する。
- feature flagと実験的APIを分離し、semver上の破壊的変更方針を決める。

### `[MEASURE]` 配布物と実利用の検証

- clean environmentでwheel/npm/CLIをインストールし、最小例・大規模例・保存再読込を実行する。
- cold start、バイナリ/WASMサイズ、インストール時間、対応Python/Node/Rust環境を測る。
- 代表的な利用者シナリオを匿名化fixtureとして収集し、成功率と診断品質を評価する。

## Phase 6 — 1.0サポート契約

### `[GATE]` リリース判定

- サポート対象を「評価できる機能」「保持できる機能」「未対応/警告/拒否」に固定する。
- 主要fixtureで、未説明の差分、未分類のクラッシュ、データ消失がゼロであることを確認する。
- Excel oracle測定、性能測定、安全性測定、配布物検証の結果をリリース成果物に添付する。
- `README`、`FUNCTIONS.md`、`docs/limits.md`、互換性文書、CHANGELOGの主張を一致させる。
- 1.0以後は、機能追加、性能改善、互換性修正、破壊的変更を別々の受け入れ基準で管理する。

## 測定が必要なタスクの運用ルール

1. 先に小さなfixtureと期待値を追加し、機能の正しさを `[BUILD]` で確認する。
2. 測定対象、入力分布、oracle、環境、収集項目、合格条件を先に文書化する。
3. 最低3回以上の反復と、中央値・p95・peak RSSを記録する。外れ値を都合よく捨てない。
4. 測定不能・比較不能・oracle不在は、成功ではなく明示的な未検証として記録する。
5. ベンチマークの数字だけで公開主張を更新せず、正確性と保存安全性の結果を併記する。

## 直近の実行順

1. Phase 0 `[BUILD]`: 競合比較マトリクスとfixture分類を確定する。
2. Phase 1 `[BUILD]`: OOXML保存ポリシーと未対応要素の警告契約を固める。
3. Phase 2 `[BUILD]`: 数式の依存グラフ・エラー意味論・再計算を固める。
4. Phase 0/1/2 `[MEASURE]`: 最初の互換性・I/O・数式ベースラインを取得する。
5. Phase 3 `[BUILD]` → `[MEASURE]`: VBA実行の実Excel検証へ進む。
6. Phase 4以降は、測定結果からボトルネックを選び、根拠のある最適化だけを行う。
