# Changelog

重要な変更を [Keep a Changelog](https://keepachangelog.com/en/1.0.0/) 形式で記録します。

## [Unreleased]

次の変更はここに記録します。

- G5aの内部経路として、defined namesの読込・報告で`xl/workbook.xml`だけを検証付き取得し、worksheet/table等の兄弟payloadを再読込・保持しないようにしました。保存時の全part遅延化やconstant-memory達成を意味しません。
- G5aの内部経路として、未編集table XMLを保存時のraw mapへ保持せず、編集対象だけを元ZIPから再取得してpatchするようにしました。全passthrough遅延化やconstant-memory達成を意味しません。
- G5bの内部経路として、worksheet relationship XMLの解析用索引を統合し、同一`.rels` payloadの解析用二重保持を削減しました。passthrough bytesの遅延化やconstant-memory達成を意味しません。
- G5bの内部経路として、relationship XMLのraw bytesを保存時mapへ保持せず、元ZIPから遅延copyするようにしました。新規tableのworksheet `.rels`だけはpatchして出力します。全payload遅延化やconstant-memory達成を意味しません。
- G5bの内部経路として、relationship接続・pruning解析後に`raw_entries`側の`.rels` bytesを解放し、解析用索引とentry名だけを残すようにしました。全payload遅延化やconstant-memory達成を意味しません。
- G5の測定記録を更新し、現行1.0.4 release wheelのappend 100,000／250,000行を別processで再測定しました。normal-VMの長時間化ケースは未測定として明示しています。
- G5の通常VM測定ハーネスを4,096行単位の`set_range`バッチへ変更し、100,000／250,000行のnormal-fresh保存・RSS結果を記録しました。250,000行RSSは2回間で大きく変動したため、constant-memoryや3 OS対応の根拠にはしていません。
- 大規模な範囲編集でtransaction内の各操作が全VM undo snapshotを追加しないようにし、commit時にtransaction全体を1つのundo単位として記録するようにしました。通常VMの測定ハーネスもこのatomic bulk-edit経路を使用します。
- 同じ条件のnormal-fresh再測定で、100,000行のRSS p95を623.22 MiBから118.30 MiB、250,000行を1,920.73 MiBから203.61 MiBへ削減しました。これはundo履歴の削減効果であり、constant-memoryの証拠ではありません。
- transaction経路を1,000,000行まで測定し、appendはRSS p95 19.44 MiB、normal-freshは753.20 MiBでした。ZIP・worksheet形状・最終行・出力検証は成功しましたが、Linux/Windowsやnormal VMのconstant-memoryを示すものではありません。
- Python type stubとREADMEに、transaction・undo/redo APIと「transaction全体を1つのundo単位として扱う」動作を追記しました。
- transaction開始時に既存redo履歴を複製しないようにし、abort時は既存redo履歴を保持する回帰を追加しました。
- 現行1.0.4のappend-cache Criterionベースラインを更新し、5,000行でreference 106.88 ms、cached 94.98 msを記録しました。これは局所microbenchmarkの観測であり、1.2倍やend-to-end速度の主張ではありません。
- formula dirty propagationのcontrolled matrixを再測定し、single-input chain 1.2825 ms、warm noop 1.2363 ms、structure rebuild 1.3090 ms、独立1,000入力 1.1570 msを記録しました。dirty経路の一般的な速度優位は主張せず、negative resultとして扱います。
- dirty closure queueがformula plan indexを直接運ぶ局所最適化を追加し、single-input chainを1.2538 ms（直前比約2.2%改善）で測定しました。rebuildは1.3223 msで変化なしのため、一般的な速度優位は主張していません。
- queue index化後のwarm noop 1.2012 ms、独立1,000入力1.1502 msを再測定しましたが、いずれも有意差はありませんでした。効果はsingle-input chainに限定して扱います。
- 現行1.0.4 candidateのworkspace全target回帰を実行し、Rust unit 1,549件、blackbox、CLI、property、XLSX round-trip 52件、bench smoke、WASM crateの成功を確認しました。これはlocal regression evidenceであり、3 OSやExcel oracleの証拠ではありません。
- 追跡済みformula planでは毎回の全formula整合性scanを省略し、`cells_mut()`等でtrackingが無効化された場合だけfallback検査するようにしました。dirty formula、cycle、manual→automaticの関連回帰を確認しています。

## [1.0.4] - 2026-09-07

- READMEとv1 support contractの製品説明を、VBA実行専用ではなく、数式計算・workbook編集・保存・VBA実行を含むヘッドレスworkbook自動化ランタイムとして整理しました。
- G3の初段として、sheet-local named rangeの登録APIをRust VM／Pythonへ追加し、formula評価でhost sheetのlocal nameをglobal nameより優先するようにしました。ロード時は単純A1／範囲definedNameをscope付きで取り込み、単純なqualified targetも評価します。dynamic／structured nameは未対応です。
- ロード済みの単純definedNameについて、行／列の挿入・削除およびsheet rename時に参照先を追従させるようにしました。sheet移動・追加に伴うlocalSheetIdの完全更新は未対応です。
- `OFFSET`で表現できるdynamic definedNameを数式ASTへ展開し、`SUM(DynamicRange)`などの集計入力として評価できるようにしました。structured／external referenceは未対応です。
- 単純な`TableName[ColumnName]`をtableのデータ行範囲へ評価時に正規化し、`SUM(TableName[ColumnName])`をworkbook依存グラフで評価できるようにしました。複雑なspecifierは未対応です。
- `[#Headers]`、`[#Data]`、`[#Totals]`、`[#All]`と列指定の静的specifierをtable範囲へ正規化できるようにしました。`[@Column]`などhost行依存のspecifierは未対応です。
- 同一tableのデータ行で`Table[@Column]`および`Table[[#This Row],[Column]]`をhost行のセル参照へ正規化できるようにしました。table外・別sheetの暗黙的な行コンテキストは未対応です。
- `Table[[Column1]:[Column2]]`と各静的specifierの連続複数列指定をtable範囲へ正規化できるようにしました。
- table名・列名に数字が含まれる場合も内部正規化識別子がセル参照と衝突しないようにし、列削除後のstructured formula追従を回帰テストで固定しました。
- structured referenceの正規化で、文字列リテラル内と長い識別子の一部を変換対象から除外し、循環するformula dependency graphをWASM診断の`hasFormulaCycle`で報告できるようにしました。循環式の再計算は従来どおりcached valueを使うbest-effortです。
- `IF`の条件式がExcel Error値を返した場合に、真偽分岐でErrorをfalse扱いせず、そのError値を伝播するようにしました。真の分岐・偽の分岐は従来どおり必要な側だけを遅延評価します。
- `NOT`の引数がExcel Error値の場合も、Errorをfalse扱いせず伝播するようにしました。
- `AND`/`OR`は全引数を評価し、真偽値で途中決定できる場合でもExcel Error値を結果へ伝播するようにしました。`IFERROR`での回復も回帰テストで固定しています。
- `XOR`も全引数を評価し、引数中のExcel Error値を結果へ伝播するようにしました。
- workbook formula dependency graphのready nodeと循環残りを安定ソートし、同じ入力での再計算・診断順を決定的にしました。
- 循環参照診断がqualifiedなcross-sheet cell referenceも依存辺として扱う回帰を追加しました。

- Pivot cache ownerの再出力と、cache definition／recordsへ至る内部relationship到達性の保存時検査を追加しました。Pivotの再集計・編集は未対応です。
- sheet rename／row・column構造編集／sheet move・delete後は、参照更新未実装のDrawing／Pivot ownerを保存時に復元しない安全境界を追加しました。値セル編集では従来どおり保持します。

- LogiSheets対抗トラックを追加し、workbook数式依存関係・共有Rust/WASM runtime・undo/redo・安全なplugin/AI境界をL0–L6へ分解しました。
- シート修飾参照を含むformulaについて、workbook全体の依存順再計算を追加しました。WASM/Node/browser API、増分dependency graph、undo/redoは引き続き未完了です。
- 大きなrangeを全セル展開せずにformula node間の依存辺を構築するinterval-style経路と、単純なruntime named rangeのformula展開を追加しました。
- workbook dirty trackingを追加し、値セルの変更から直接参照・range参照・formula chainの依存先だけを差分再計算します。完全なinterval tree、動的/structured named range、WASM/Node/browser parityは未完了です。
- 共有Rust formula runtimeのfull workbook計算をWASM `calculateWorkbook(bytes)`として追加し、Node/browser同梱bridgeを更新しました。JS側の増分API、diagnostics、worker/async loadingは未完了です。
- WASMに`diagnoseWorkbook(bytes)`を追加し、シート数・数式数・qualified formula数・parse error数を構造化JSONで取得できるようにしました。
- Rust VM・Python APIに、セル/範囲値とセル数式の最大128操作undo/redoを追加しました。シート操作とJS/WASM history APIは未完了です。
- Rust VM・Python APIにedit transactionのbegin/commit/abortを追加し、abort時に編集内容とprior undo/redo履歴を復元します。ネストしたtransactionは拒否します。
- WASMにstateful `WorkbookEditor`を追加し、`setNumber`・`recalculate`・`undo`/`redo`・transactionをNode/browserの共有Rust runtimeから利用できるようにしました。OOXML書き込みは従来どおりJavaScript側です。
- 互換root exportを変更せず、`@elixcee/xlsx/runtime`サブパスから`WorkbookEditor`・`calculateWorkbook`・`diagnoseWorkbook`を利用できるようにしました。
- WASM smokeにruntimeサブパスのNode／browser条件検証と、意図的に更新したpayload baselineに対する10%サイズ成長ゲートを追加しました。現行のbaselineはWorkbookEditor追加後の実測値です。
- externalReferences ownerを再生成workbookへ戻し、sourceのType/Targetに対応する新しいrelationship IDへ書き換えるようにしました。未解決の外部参照はdangling ownerを出力せず省略します。Pivot cacheや外部取得は未対応です。
- worksheetのdrawing／legacyDrawing ownerについて、r:idからrels、相対target、出力partまでの接続を保存時に検査し、欠落・重複時はdangling ownerを復元しないようにしました。Chart／imageの推移的検証は未完です。
- Drawing等の内部partに付随する`.rels`も推移的に検査し、Chart／imageへの欠落targetを検出するようにしました。Pivot cacheの接続検証は未完です。

## [1.0.3] - 2026-09-06

- G5の保存メモリ削減を段階実装しました。画像・VBAなどの大きなpassthrough payload、未編集table、条件付きworksheet rels、未編集stylesをsource ZIPから遅延copyし、worksheet/workbook XMLとpassthrough XMLを出力単位で解放します。編集対象partは従来どおり保持・patchします。
- Python追記Writerに行単位・総work量・列数の上限を実装し、abort時のtemp cleanup、close-before-rename、sync経路を共通化しました。`create_stream_bounded`、RSS/temp disk測定、XML escape／巨大文字列校正、3 OS手動CI matrix、測定JSON validatorを追加しました。
- G5のローカル検証としてRust全workspace／全target 1,709件、XLSX round-trip 52件、strict clippy、rustfmt、actionlintを確認しました。Linux/Windows実測、Excel再open、constant-memory保証、workbook rels等の完全遅延化は未完了です。

## [1.0.2] - 2026-09-06

性能値は各リンク先の測定時点（1.0.1表記の開発ツリー）の記録です。公開済み1.0.1との比較や、1.0.2タグそのものの再測定と混同しないでください。

- JSの`sheet_to_html`が既定で`cell.h`を捨てずにエスケープ表示するよう修正しました。生HTMLの出力は引き続き明示的な`rawHtml: true`に限定します。JSパッケージはprivateのままです。
- CLI契約・ロードマップ・各言語README・互換性／資源／配布境界の文書を整理し、0.x変更履歴を別文書へ移動しました。過去の測定条件と未達結果は保持しています。
- 数式校正バイナリもCargo配布対象から除外し、通常のvタグでCLI配布物を作成するよう公開workflowを統一しました。
- Linuxでの測定補助ツールとWASMでの警告チェックに対応しました。atomic出力は引き続きファイルを閉じてからrenameします。
- 大規模XLSX向けに、保存時の行別tree／配列を単一の座標sortへ変更し、セル番地をスタック上で生成するようにしました。Readerでは属性配列とセル型の作業領域を再利用し、タグ境界・重複属性・禁止宣言の検査コストを削減しました。直前の実装に対する20組の交互測定で、10万／40万／100万セル（25万×4シート）の全体中央値は1.180／1.318／1.196倍でした。厳密な1.2倍目標の達成は40万セルのみです。全ZIP部品の展開後バイト一致、全ターゲットtest／全feature clippyを確認し、保存耐久性・入力制限・versionを維持しました。生データと小規模／文字列混在の回帰測定は `docs/benchmarks/workbook-large-speedup-2026-09-06.md` を参照してください。
- opaque XML要素検索で、不在の要素名を先に判定し、タグ名の一時`String`生成をborrowへ置換しました。共有文字列収集では文字列セルだけを座標順に並べ替え、数値セルの不要なsort／lookupを削減しました。直前の実装を基準とした40組の交互測定で、数値1万／10万セル・混在1万セルは約2倍に高速化し、全ZIP部品の展開後バイト一致を確認しました。小規模は約1.09倍で1.2倍目標未達です。保存耐久性とversionは維持し、詳細は `docs/benchmarks/workbook-speedup120-2026-09-06.md` に記録しました。
- ClosedXML 0.105.1／.NET SDK 10.0.400を固定した比較用workerを追加しました。macOSの同一入力・編集・`F_FULLFSYNC`・atomic rename・再読込条件で、elixcee／ClosedXML／openpyxlを3ケース各30反復×2回測定しました。両runの全ケースでelixceeの中央値がClosedXMLを下回りましたが、高負荷環境の大きなp95変動とラウンド単位の逆転も含め、暫定値として `docs/benchmarks/workbook-closedxml-2026-09-06.md` に記録しました。C#依存は製品の依存関係に追加していません。
- 公開されている低レベルcell map経由で数式本文が変更された場合も、warmな依存計画を無効化して再評価するよう修正しました。
- worksheet XMLを一時的な巨大`String`へ構築せず、保存先の`ZipWriter`へ直接ストリーミングするsink経路へ変更しました。passthrough entry全量保持の削減は引き続き未完了です。
- passthrough entry はメタデータ解析後に所有権移動するようにし、保存用payloadの不要な二重cloneを除去しました。raw ZIP全体の遅延展開は引き続き未完了です。
- 通常Writerのshared-string本文をindexと別の`Vec<String>`へ二重保持せず、所有indexから直接出力するようにしました。style／全補助索引とraw ZIP全体の遅延展開は引き続き未完了です。
- 保存時の画像・VBAなど非XML passthrough payloadをraw mapへ常駐させず、元ZIPから一件ずつ再読込して転送するようにしました。XML／relsの遅延展開と全体のconstant-memory化は引き続き未完了です。
- 遅延passthrough partは一時`Vec<u8>`を作らず、元ZIPのentryから出力ZIPへ直接copyするようにしました。必要なXML／relsとVM全セル保持は引き続きメモリ上に残ります。
- 保存時のwriter-owned styles/workbook/[Content_Types] partはraw mapから所有権移動し、解析用値との不要なcloneを削減しました。cell/row/column style編集のeffective mapは対象sheetだけをoverlay cloneします。
- worksheet source XMLも削除判定後に所有権移動し、raw mapとの重複cloneを削減しました。
- font/fill/borderを変更しないstyle編集では、巨大な共有style tableを展開せずcellXfsだけを再構成するようにしました。
- 遅延passthroughの直接copy前に元ZIPを再検証し、entryの実コピー量が期待サイズと一致することを確認するようにしました。
- dirty/full-rescan の代表値一致を検査しながら p50/p95 を出力する、formula dirty propagation のローカル校正バイナリを追加しました。
- formula dirty propagation の校正を100式／1,000式のcontrolled matrixへ拡張し、各ケースのp50/p95とfull-rescan一致を記録しました。
- formula dirty propagation の測定バイナリに、Unixのpeak RSS・user/system CPUカウンタ出力を追加しました。
- manual→automatic遷移と循環参照のp50/p95校正も追加し、dirty propagationの測定行列を拡張しました。
- release-profile 10反復の dirty/full matrix と resource counter 結果を測定成果物へ保存しました。
- Writerのworksheet sink／passthrough clone削減後に代表XLSXのrelease保存soakを3回実行し、round-tripとallocator結果を保存しました。
- formula dirty propagation の release matrix を30反復へ拡張し、p95の再現性を測定成果物へ保存しました。
- elixcee／openpyxl／LibreOfficeの初回XLSX速度比較と、比較不能条件を含む公開用ベンチマーク成果物を追加しました。
- openpyxl比較に保存後fsyncを追加した初期5反復を記録しました。ただしmacOSのRust `sync_all`（`F_FULLFSYNC`）とは不一致だったため、この初期測定を公平な比較とする説明を撤回しました。
- 標準の耐久性保証を変えず、最終fsyncを省略する明示的な `save_workbook_fast` APIを追加しました。高速経路はクラッシュ耐性が弱いため用途を限定します。
- XLSX writerを圧縮前後の64 KiBバッファとシーク不要のZIP出力へ変更し、標準の `sync_all`／atomic renameを維持したまま細切れ書き込みを削減しました。数値1万／10万セルの保存中央値を変更前より約49%／55%短縮しました。
- macOS `F_FULLFSYNC`まで一致させたopenpyxl 3.1.2比較を、3ケース・30反復＋独立40反復で実測しました。未公開Rust APIのload/edit/save/reload中央値は再測定で1.90〜2.98倍高速でした。条件・p95・全生データは `docs/benchmarks/workbook-equal-durable-2026-09-06.md` を参照してください。
- 純粋な VM 内 `Scripting.Dictionary` adapter を追加しました。`New`、キーの追加・検索・削除、`CompareMode`、挿入順のキー列挙をサポートし、`CreateObject` の外部効果遮断は維持します。
- VBA built-in `Collection`を追加しました。`New Collection`、1-based index、case-insensitive key、`Add`/`Remove`/`Item`/`Count`、値/object `For Each`、object-valued itemの`Set`取得、`With`、参照alias、到達可能性回収、要素budgetに対応します。
- Collectionの`Call c.Add(...)`/`Call c.Remove(...)`構文と、export済みclass moduleの基本object連携（生成、instance field、初期化、Sub/Function、Collection/`For Each`/`With`）を追加しました。
- class moduleの非indexed Property Get/Let/Set、object-valued field/引数、`Class_Terminate`、`Implements`署名検証、Private/Friend境界、class object arrayを追加しました。
- indexed `Property Get/Let/Set`、型付き引数を受けるobject-returning class `Function`、interface宣言型を保持する厳密なSub/Function/Property dispatch（`With`を含む）を追加しました。
- Range objectの相対`.Cells`/`.Range`、1-based default `Item`/`Value`、`With`内の相対参照を追加しました。
- Worksheetの`Name`/`Index`/`Visible`/`UsedRange`/`Activate`/`Select`と、Workbookの`Name`/`ActiveSheet`/`Worksheets.Count`/`Sheets.Count`をdirect/object参照へ追加しました。シート名変更後も保持中のWorksheet/Range参照は同じシートを指し続けます。
- `Range.SpecialCells(Type[, Value])`を全`XlCellType`定数へ拡張し、定数/数式の型mask、空白、最終セル、入力規則、条件付き書式、legacy comment、該当なし1004、走査budgetに対応しました。
- Range、Collection、イベント、外部オブジェクトの残作業を安全境界ごとに分割してロードマップへ追加しました。
- 小容量入力が監視threadの初回pollより先に完了しても、事前作成済みCLI cancel-fileを確実に反映するよう競合を修正しました。

## [1.0.1] - 2026-09-04

- readerの総work budget、deadline、協調キャンセルをRust/Python/CLIへ通し、CLIではSIGINTとcancel-fileから安全に中断できるようにしました。
- ZIP/XMLのchunk境界でキャンセルを確認し、キャンセル・大容量入力・反復資源回収・read/mutate/write経路をローカルfixtureで校正しました。
- XMLイベントをborrow/`Cow`化し、worksheetの構文・資源・shared-string検証とセル構築を単一走査へ統合しました。
- `WorkbookSheet`からVMへの変換を所有権移動に変更し、セル値・数式・シートメタデータの不要なcloneを除去しました。
- macOS arm64のrelease build、dense 400,000セル、3回のローカル比較で、reader→VMロード中央値を約1,386.5 msから650.5 msへ短縮しました（約2.13倍、条件と制約は`docs/measurements/reader-optimization-2026-09-04.md`参照）。
- reader測定成果物のsemantic validator、拒否系self-test、Cargo package境界チェックをCIへ追加しました。

## [1.0.0] - 2026-09-01

- 文書化されたデータ処理サブセットと、安全な失敗動作をv1サポート契約として固定しました。
- 完全なExcel/VBA互換性を主張せず、`supported`、`preserved`、`warned`、`rejected`、`unverified`の境界を公開しました。
- READMEとセキュリティモデルのリリース表記を1.0.0へ更新しました。

## Older releases

[0.xの変更履歴](docs/history/CHANGELOG-0.x.md)へ移動しました。過去の記録は保持しています。
