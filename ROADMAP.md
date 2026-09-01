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

## 明示する未達項目

現時点では、次の3点を互換性・安全性の未達として扱う。これらを解消するまで、
「OOXML完全互換」「全Excel関数互換」「全VBAオブジェクト互換」「任意サイズ入力を
安全に処理できる」とは主張しない。

1. **OOXML全機能の完全な往復互換ではない** — 未知part、拡張属性、relationship、
   drawing/chart/pivot等が、編集・保存後に失われたり切断されたりする可能性がある。
2. **全Excel関数・全VBAオブジェクトの互換性はない** — 未対応関数・オブジェクトを
   空値や成功に変換せず、unsupported / warned / rejectedとして診断できる状態が必要。
3. **reader全体の処理時間予算・キャンセル機構が未実装** — 個別のサイズ上限だけでは、
   多数の合法なpartや高コストな組み合わせによるCPU消費を防げない。

### 解除条件

- OOXML: feature/part/relationshipごとの保持率、編集可否、Excel再オープン結果をfixture付きで公開する。
- Excel関数/VBA: 対応表、独立期待値、Excel oracle結果、未対応診断をカテゴリ別に公開する。
- reader: deadline・cancellation token・総work budgetをAPI/CLI/Pythonで提供し、タイムアウト時に
  部分的なWorkbookを返さず、確実に後始末して終了する。

## XLSX editing compatibility track

現状、既存XLSXを総合的に読み込み、編集し、保存する用途では ClosedXML / Aspose.Cells が
上位です。この差を埋める対象は単なるAPI数ではなく、次の4段階です。

1. **読み込める** — OOXMLを解釈し、値・型・数式・表示情報を取得できる
2. **編集できる** — セルだけでなく、シート構造・スタイル・テーブル・描画等を変更できる
3. **壊さず保存できる** — 未編集のpart、relationship、拡張属性を保持して往復できる
4. **Excelで再利用できる** — Excelで開け、再計算・印刷・表示・マクロ実行に支障がない

ClosedXML / Aspose.Cellsを機能の優先順位付けとAPI比較の対象にするが、互換性の合格判定は
実Excelまたは独立したOOXML検証で行う。特に「読み込めた」「保存できた」だけでは完了とせず、
保存前後の論理差分と、未対応要素の保持・警告・拒否を記録する。

## XLSX Phase X0 — 編集互換性の棚卸し

### `[XLSX-BUILD]`

- OOXML part、relationship、content type、拡張属性を一覧化し、機能マトリクスを作る。
- セル/数式/書式だけでなく、行列寸法、結合、名前、テーブル、AutoFilter、検証、コメント、
  ハイパーリンク、画像、図形、チャート、ピボット、条件付き書式、印刷設定、保護、
  非表示状態、外部リンク、VBA projectを分類する。
- 各機能を `read` / `edit` / `write` / `round-trip` / `excel-verified` に分けて表示する。
- 競合の挙動をそのまま仕様にせず、elixceeの保持・警告・エラー方針を機能ごとに決める。

### `[XLSX-MEASURE]`

- ClosedXML、Aspose.Cells、openpyxl、Excelで同一fixtureを開き、機能単位の差分を収集する。
- real-world fixtureと最小再現fixtureを分け、比較不能・未対応・破損を別の判定にする。
- fixtureごとに、保持できたpart数、失われた属性、警告、Excelでの再オープン結果を保存する。

## XLSX Phase X1 — 基本モデルと編集API

### `[XLSX-BUILD]`

- Workbook / Worksheet / Row / Column / Cell / Range / Style / Relationshipの内部モデルを整理する。
- 1-based座標、Excelの型、数式とcached value、日付シリアル、エラー値を一貫した型で扱う。
- bulk range、行列操作、insert/delete、move/copy、sort、merge、名前付き範囲を実装する。
- styleを値コピーではなく共有・差分適用できるモデルにし、意図しない全体書式変更を防ぐ。
- mutation後のdirty partを追跡し、編集していないpartを再生成しない保存経路を用意する。

### `[XLSX-MEASURE]`

- 1セル、1行、矩形範囲、疎な範囲、大量スタイルの編集について、API結果と保存後結果を測る。
- 編集前後のXMLをバイト比較せず、正規化した論理モデル・relationship・part一覧で比較する。

## XLSX Phase X2 — 表示・構造化オブジェクト

### `[XLSX-BUILD]`

- テーブル、構造化参照、AutoFilter、データ検証、条件付き書式、コメント、ハイパーリンクを
  読み書き・編集できるようにする。
- 行高、列幅、hidden、freeze pane、print area、page setup、sheet protectionを保持・編集する。
- 画像、drawing、chart、pivot、external linkは、編集対応・無変更保持・明示拒否を分けて実装する。
- unsupported objectを検出した場合、保存前に診断へ出し、静かな消失を許さない。
- unknown part/attribute/relationshipを、保持・再配置・明示拒否のいずれかに分類し、分類漏れをエラーにする。

### `[XLSX-MEASURE]`

- Excelで作成したfixtureを保存→elixcee編集→Excel再オープンし、表示・印刷・フィルタ・参照を確認する。
- 競合ライブラリで同じ編集を行い、機能の可否ではなく、出力の論理同値とExcelでの再利用性を比較する。

## XLSX Phase X3 — 数式・再計算・マクロ共存

### `[XLSX-BUILD]`

- 数式の保存と評価を分離し、未対応関数を勝手に空値へ変換しない。
- Excel関数をカテゴリ別の対応表で管理し、未対応関数・引数型・戻り値・エラー値を診断可能にする。
- 依存関係、shared/array/dynamic formula、calculation chain、manual/automatic calculationを扱う。
- `.xlsm` のVBA project bytes、関連relationship、署名や未知partは、編集時の保持方針を明示する。
- マクロを実行しない編集と、VBA VMで実行する処理をAPI上で明確に分離する。
- VBAオブジェクトをApplication/Workbook/Worksheet/Range等の対応表で管理し、未対応メンバーを明示的に拒否する。

### `[XLSX-MEASURE]`

- Excelで再計算した値、エラー、spill範囲、calculation chain、保存後のマクロ実行結果を比較する。
- Excel関数・VBAオブジェクトを網羅表のカテゴリ単位で測定し、成功率と未対応率を分けて報告する。
- 数式密度、依存グラフ、未対応関数、マクロ有無を変え、再計算時間と結果の安定性を測る。

## XLSX Phase X4 — lossless寄りの保存と障害復旧

### `[XLSX-BUILD]`

- 保存をtransactionalにし、書き込み途中の失敗で元ファイルを破壊しない。
- unknown part、unknown attribute、relationship、名前空間、順序依存を可能な範囲で保持する。
- preserve mode、edit mode、strict modeを用意し、互換性と安全性のトレードオフを明示する。
- 出力の自己検査、ZIP/XML整合性検査、再読込検査を保存処理に組み込む。

### `[XLSX-MEASURE]`

- 全機能を含む複合fixtureを複数回往復し、データ消失・Excel警告・破損・part欠落を確認する。
- OOXML partとrelationshipを正規化して比較し、未編集partの消失、属性欠落、namespace変更を検出する。
- 異常終了、ディスク不足、権限エラー、キャンセル後に元ファイルと一時ファイルが安全か測る。

## XLSX Phase X5 — 性能・メモリ・API互換性

### `[XLSX-BUILD]`

- read/write API、bulk編集、streaming API、Python/CLI/WASMの役割と制約を整理する。
- 互換APIは必要な範囲で提供するが、実装できない機能を同名APIで誤認させない。
- parse cache、style deduplication、shared strings、dirty-part保存をボトルネックに応じて実装する。

### `[XLSX-MEASURE]`

- ClosedXML、Aspose.Cells、openpyxlと同じ操作で、time、p50/p95、peak RSS、出力サイズを測る。
- 小規模・大規模、dense/sparse、style-heavy、formula-heavy、object-heavyの入力群を使う。
- 性能比較は、Excelで開けることと論理差分ゼロまたは明示済み差分を満たした結果だけ採用する。
- 「ClosedXML/Aspose.Cellsより速い」という主張は、fixture、環境、バージョン、測定日を併記する。

## XLSX Phase X6 — 互換性リリースゲート

### `[XLSX-GATE]`

- priority機能について、`read`・`edit`・`write`・`round-trip`・`excel-verified` の状態が公開されている。
- 未対応partの静かな消失、未分類のExcel警告、再現不能な出力差分をゼロにする。
- ClosedXML/Aspose.Cells/openpyxlとの差分は、バグ・意図的差分・未対応・比較不能に分類する。
- 互換性の改善を、機能追加、保存安全性、数式/VBA、性能の別々のリリースノートで説明する。

## Security-first track

セキュリティ上の「穴を無くす」は、未発見の脆弱性がゼロだと宣言することではなく、
不信なファイル・マクロ・数式・出力を境界ごとに制御し、失敗を安全で再現可能なエラーに
することを意味します。以下のPhaseは通常の機能Phaseと並行して進め、下位Phaseの合格前に
新しい入力経路や実行権限を増やしません。

### 判定ラベル

- `[SECURITY-BUILD]`: 型、境界チェック、拒否規則、エラー契約、単体/property testで確認する。
- `[SECURITY-MEASURE]`: fuzz、悪意あるfixture、大規模入力、CPU/RSS/時間、実環境での
  隔離検証が必要。測定結果を保存するまで完了にしない。
- `[SECURITY-GATE]`: 合格するまでリリースや新しいデフォルト許可に進まない。

## Security Phase S0 — 脅威モデルと資産境界

### `[SECURITY-BUILD]`

- 入力を「XLSX/XLSM/ODS/ZIP/XML」「VBA」「数式」「外部リンク」「生成HTML/JSON」に分類する。
- 攻撃者が制御できる値、Rust/Python/JS/WASM/CLIをまたぐ境界、守る資産を一覧化する。
- prototype pollution、ZIP bomb、XML/文字列爆発、パストラバーサル、無限ループ、
  unsafe HTML/URL、秘密情報のログ漏えいを脅威カテゴリとして固定する。
- `supported`、`warned`、`rejected`、`disabled-by-default` を機能ごとに定義する。
- `docs/xlsx-security-model.md`、`docs/limits.md`、診断JSON、READMEの主張を同期する。

### `[SECURITY-MEASURE]`

- 実在fixture、過去の不具合、手作業で生成した悪意あるfixtureを脅威カテゴリ別に収集する。
- 各ケースで、クラッシュ・ハング・過剰RSS・意図しない外部アクセス・危険な出力がないかを測る。

## Security Phase S1 — ZIP/XML入力の防御

### `[SECURITY-BUILD]`

- ZIP entry数、entryごとの展開後サイズ、総展開サイズ、圧縮率を上限付きで検査する。
- XMLの要素数、属性数、属性値/テキスト長、共有文字列数、シート/行/列/セル数を制限する。
- 外部エンティティ、外部DTD、外部参照、想定外のencoding、異常なrelationship targetを拒否する。
- パスを正規化し、出力先・relationship・ZIP entryから作業ディレクトリ外へ出られないようにする。
- 上限超過は一貫したエラーコードで返し、部分的に信頼できるWorkbookを返さない。

### `[SECURITY-MEASURE]`

- 各上限を、正常な大規模ファイルと悪意ある最小ファイルの両方で測定して決める。
- 展開時間、peak RSS、CPU時間、entry数、圧縮率を記録し、連続実行で資源が回収されることを確認する。
- 既存の `ZIP_ENTRY_MAX_BYTES` とJS側の範囲上限を、全体予算との組み合わせで再検証する。

## Security Phase S2 — parser / formula / VM の耐性

### `[SECURITY-BUILD]`

- parserに深さ、トークン数、識別子長、式長、配列サイズ、ASTサイズの上限を設ける。
- 数式の循環、巨大範囲、再計算爆発、NaN/Infinity、ゼロ除算、型変換を安全なエラーとして扱う。
- VBAに命令数、wall-clock、メモリ、再帰/呼び出し深度、配列/文字列サイズのbudgetを適用する。
- `On Error` や診断機能が、制限超過・拒否エラーを握りつぶして成功扱いにできないようにする。
- 外部ファイル、Shell、COM、ActiveX、UserForm、ネットワーク相当の機能は明示的に無効化または
  opt-inとし、デフォルトのVMから到達できないことを型/APIで確認する。

### `[SECURITY-MEASURE]`

- `fuzz/fuzz_targets/` のparser、formula、VBA、readerを継続実行し、crash/timeout/RSSを保存する。
- 巨大な入れ子式、循環参照、長大文字列、分岐の多いマクロで、制限が決めた時間内に停止するか測る。
- 同じseed・入力を複数回実行し、結果・エラー・消費資源が決定的であることを確認する。

## Security Phase S3 — 出力・互換APIの安全化

### `[SECURITY-BUILD]`

- JSON/object keyは `__proto__`、`constructor`、`prototype` を含めてもprototypeを変更しない構造にする。
- HTML属性、テキスト、URLを別々にescapeし、許可しないURL schemeはリンクとして出力しない。
- raw HTML/rich textは入力経路と信頼境界を明示し、未信頼ファイル由来の値をそのまま描画しない。
- 上書き、symlink、temporary file、permission、保存失敗時のcleanupを安全側に固定する。
- エラー・ログ・diagnose出力からセル値、パス、マクロ本文などの秘密情報が漏れないようにする。

### `[SECURITY-MEASURE]`

- 生成HTMLをブラウザで実際に読み込み、script実行、unsafe URL、属性脱出が起きないことを確認する。
- JSON、XLSX、XLSMの往復で、危険キー・リンク・rich text・macro bytesの保持と安全性を同時に測る。
- 競合oracleとの比較で安全側に意図的に差分を出す場合は、`INTENTIONAL_SECURITY_DIVERGENCE`
  として理由とfixtureを登録する。

## Security Phase S4 — 隔離・実行環境・依存関係

### `[SECURITY-BUILD]`

- CLI、Python、WASMで権限モデルを明示し、入力ディレクトリ・出力ディレクトリ・一時領域を分離する。
- マクロ実行をプロセス/コンテナ等で隔離できる実行モードと、利用者が選べる安全なデフォルトを用意する。
- 依存crate/npm/Python packageの固定、SBOM、license、脆弱性スキャン、更新手順をCIに組み込む。
- release binary、wheel、npm、WASMに不要な権限・デバッグ情報・秘密を含めない。

### `[SECURITY-MEASURE]`

- Linux/macOS/Windowsのclean環境で、ファイル・ネットワーク・プロセス・symlinkへの到達可能性を検証する。
- 隔離あり/なし、タイムアウト、kill、キャンセル、並列実行時のRSSと残存プロセスを測る。
- 依存更新後に、既知脆弱性の再現fixture、fuzz smoke、パッケージ内容監査を実行する。

## Security Phase S5 — 継続的攻撃検証とリリースゲート

### `[SECURITY-BUILD]`

- security regression test、fuzz corpus、限界値fixture、公開脆弱性対応の手順をCIに登録する。
- 脆弱性のseverity、影響範囲、修正commit、回帰テスト、公開判断を記録する。
- 重要なセキュリティ変更には、入力境界・エラー契約・互換性差分のレビューを必須にする。

### `[SECURITY-MEASURE]`

- 定期的な長時間fuzz、大規模悪意入力、clean-install、主要OS matrixを実行する。
- false positiveでなく、crash、hang、resource exhaustion、意図しない外部作用を合格判定の対象にする。

### `[SECURITY-GATE]` 1.0および各リリース

- 未分類のcrash、無期限hang、制限なしの資源消費、未承認の外部作用、prototype/HTML injectionをゼロにする。
- すべての上限値に根拠となる測定または保守的なセキュリティ判断を添付する。
- 安全側の意図的な非互換、既知の制限、利用者が有効化すべき危険機能をCHANGELOGとdocsに記載する。
- 「完全に安全」「任意のVBAを安全に実行」とは表現せず、検証済みの境界と前提条件だけを公開する。

## Security work order

1. S0 `[SECURITY-BUILD]`: 脅威モデル、資産、境界、エラー語彙を固定する。
2. S1 `[SECURITY-BUILD]`: ZIP/XMLの複合上限と安全な失敗を実装する。
3. S2 `[SECURITY-BUILD]`: parser/formula/VMのbudgetと外部作用の遮断を実装する。
4. S1/S2 `[SECURITY-MEASURE]`: fuzz・悪意入力・大規模入力で上限を校正する。
5. S3 `[SECURITY-BUILD]`: HTML、URL、JSON key、ファイル出力を安全化する。
6. S4 `[SECURITY-BUILD]` → `[SECURITY-MEASURE]`: OS・配布物・隔離を検証する。
7. S5 `[SECURITY-GATE]`: 未分類の危険な挙動がないことを確認してからリリースする。

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
- reader全体にdeadline、cancellation token、総work budgetを通し、CLI/Python/APIから中断できるようにする。
- キャンセル時は入力・一時ファイル・ZIP/XML parserの状態を確実に解放し、部分Workbookを返さない。

### `[MEASURE]` 大規模I/Oとメモリ上限

- streaming readerと通常readerを同一fixtureで比較し、RSSが入力サイズにどう増えるか測る。
- 破損/悪意あるZIP/XMLを含め、上限到達までの時間、エラーの決定性、部分出力の有無を測る。
- reader開始から終了までの処理時間予算、途中キャンセル、deadline直前、キャンセル競合を測る。
- 入力サイズだけでなく、entry数・XML要素数・relationship数の組み合わせで総work budgetを検証する。
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

1. XLSX Phase X0 `[XLSX-BUILD]`: 競合との差分と、read/edit/write/round-tripの状態を確定する。
2. Security Phase S0/S1 `[SECURITY-BUILD]`: 脅威モデルと入力境界を先に固定する。
3. XLSX Phase X1/X2 `[XLSX-BUILD]`: 基本編集モデル、構造化オブジェクト、未対応要素の方針を固める。
4. XLSX Phase X0/X1/S1 `[MEASURE]`: 小規模fixture、実Excel、悪意ある入力のベースラインを取得する。
5. XLSX Phase X3/X4 `[XLSX-BUILD]`: 数式・マクロ共存、transactional保存、自己検査を実装する。
6. XLSX Phase X2/X3/X4 `[XLSX-MEASURE]`: Excel再オープン、論理差分、データ消失を検証する。
7. XLSX Phase X5 `[XLSX-MEASURE]`: ClosedXML/Aspose.Cells/openpyxlとの性能・メモリ比較を行う。
8. X6 `[XLSX-GATE]` と S5 `[SECURITY-GATE]` を満たしてから互換性・安全性の主張を更新する。
