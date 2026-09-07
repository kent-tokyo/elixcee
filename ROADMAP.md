# elixcee Roadmap

更新日: 2026-09-07。対象versionは **1.0.4** です。
完了項目は記載した実装・測定の範囲に限ります。公開先の状態はリリースごとに別途確認します。
版ごとの変更は [CHANGELOG](CHANGELOG.md)、実装範囲は
[FUNCTIONS](FUNCTIONS.md)、保証範囲は [v1契約](docs/v1-support-contract.md) を参照してください。

## 方針と完了の定義

Excel不要の、Rust/Pythonによる安全なファイル入出力・数式計算・VBAデータ処理を中心に開発します。
GUI、任意のCOM、完全なExcel/VBA互換、無条件のlossless保存は保証しません。
JavaScript互換APIは別トラックで、`packages/xlsx` はprivate・未公開です。

- `BUILD`: 実装と回帰テスト。`MEASURE`: 固定条件による測定。`GATE`: 配布・互換性・安全性の判定。
- `[x]` は記載した範囲の実装／ローカル検証が完了した意味です。公開・3 OS検証・Excel完全一致を含意しません。
- 互換性は `supported / preserved / warned / rejected / unverified` を区別します。
- 外部効果は既定で遮断します。保存耐久性、入力制限、意味論を速度のために緩めません。
- 測定前に対象・baseline・反復数・成功条件を固定し、未達・p95悪化・除外条件も残します。

## 次の実行順

1. **G0–G1 土台**: 下記3領域の現状を固定し、Writerの入力受け入れを先に有界化する。
2. **G2 OOXML保持**: owner XML・relationship・partを一組として往復検証する。保持と編集・再計算は分離する。
3. **G3–G4 数式**: シート横断の評価基盤を整え、既存関数の意味論校正と不足機能追加を小さな組に分ける。
4. **G5 メモリ**: 通常保存のpassthrough遅延処理と、追記専用Writerの行数非依存メモリを別々に実装・測定する。
5. **G6 判定**: quiet-host再測定、Excel oracle、3 OS、配布・安全性ゲート。以下の性能バックログも継続する。
6. **LogiSheets対抗 L0–L6**: workbook数式・共有runtime・操作履歴を、既存の安全性と互換性ゲートを維持したまま段階導入する。

## 互換性・数式・省メモリ強化（G0–G6）

1.0.4公開後の開発計画です。次の変更はUnreleasedに記録します。
各PhaseはBUILDを小さく実装し、MEASUREが未完なら未検証として残します。
EPPlus／Aspose.Cellsとの一般的な同等性や、関数名の個数だけでの優劣は達成条件にしません。

### G0 — 現状・合格条件の固定（X0 / S0）

- [x] ソース棚卸し: drawing/legacyDrawingの選択的保持はあるが、workbookの`pivotCaches`／`externalReferences`は再出力していない。partの存在だけでは接続を保証できない。
- [x] 数式はA1/RC・シート修飾参照のparse/rewriteと評価を分離しており、シート修飾参照の評価は現在拒否する。
- [x] 通常Writerは元ZIP全展開・セルmap・文字列索引を保持。Python追記WriterはZIPへ逐次出力するが、行の収集とXML生成があり、`pending_bytes`は保持RSSではなく累積受け入れ量。
- [ ] 機械可読matrix: Charts / Pivot / Drawings / External Linksごとにread・preserve・edit・recalculate・Excel再openを別状態にし、fixtureと結ぶ。

### G1 — Writerの受け入れ上限（S1 / X5、最初のBUILD）

- [x] 行数はiterator開始前、列数と累積推定byte数はセル取り込み中に検査。空文字列でもセル本体の費用を計上する。
- [x] 不正型・iterator例外・予算超過で、拒否行のXML／行数／byteカウンタを更新しない。正常行を追加して再開できる回帰を追加する。
- [x] インストールしたwheelのPythonテストをCI設定へ追加。上限直前／一致／超過と、保存前の既存出力保護をローカル検証（この変更のGitHub CI実行は未確認）。
- [ ] MEASURE: 巨大な単一文字列、XML escape拡大、allocator／Python入力側のメモリも含めRSSを校正する。ここだけではconstant-memory達成にしない。

2026-09-06 macOSローカル: [Python回帰10件](tests/python/test_stream_writer_limits.py)は
公開1.0.3で5件の問題を検出し、修正版wheelでは10/10成功。Rust全workspace／全target
1,709 tests、全feature strict clippyも成功。これはG1のBUILD検証で、RSS／Excel互換性の測定ではありません。

### G2 — OOXMLの接続を保つ（X2 / X4）

- [x] G2a 部分 BUILD: `externalReferences` owner要素を再出力し、source relationshipのType/Targetを基準に再採番後の`r:id`へ書き換える経路と、未解決relationship時にownerを省略する安全策を追加した。Pivot cache、cacheId／外部参照順序、content types／namespaceのfixture検査は未完。
- [x] G2b 部分 BUILD: worksheetのdrawing／legacyDrawing ownerについて、r:id・worksheet .rels・相対target・出力に残るpartを一組で照合し、missing target／重複relationship ID時はownerを再出力しない経路を追加した。さらにDrawing等の内部partに付随する.relsを推移的に辿り、Chart／image targetの欠落も検出する。PivotTableからcache definition／recordsまでの検証とfixture化は未完。
- [x] G2a/G2b 追加 BUILD: pivotCaches ownerを再生成workbookへ戻し、workbook relationshipの再採番とcache definition／recordsへの内部relationship到達性検査を適用した。Pivotの再集計・編集、cacheIdの意味更新、fixtureによるExcel再open検証は未完。
- [ ] G2c: sheet rename／行列挿入削除に伴うchart参照・anchor・pivot sourceを更新する。更新できない編集は明示診断／拒否し、古い参照を黙って保存しない。
- [ ] G2d: Charts / Drawingsの作成・編集API、Pivotのsource/cache更新を一機能ずつ追加。描画再現・Pivot再集計は保持とは別の未完項目として扱う。
- [x] G2c safety BUILD: sheet rename、row/column insert/delete、sheet move/deleteを構造編集として追跡し、未更新のDrawing／Pivot ownerを保存時に復元しない安全境界を追加した。参照の実更新と明示的な編集APIは未完。
- [ ] External Linksは既定で非取得・非実行。保持するURLを辿らない。削除／拒否policyと外部参照数式の非評価を明示する。
- [ ] MEASURE: 自作最小packageとExcel由来fixtureで、part／rels／owner／cacheを比較し、実Excelの修復警告と編集後の再利用を確認する。

### G3 — Workbook単位の数式評価（X3）

- [x] G3 部分 BUILD: workbook数式評価へsheet-local named rangeを追加し、host sheetのlocal nameをglobal nameより優先する解決規則を固定した。ロード時は単純A1／範囲definedNameと、既存の`OFFSET`で表現できるdynamic definedNameをglobal／local scope付きで取り込む。tableの単純列参照、連続複数列、静的specifier、同一tableデータ行内の`[@Column]`／`[#This Row]`をqualified rangeへ評価時正規化する。qualified target、行／列insert・delete、sheet renameにも追従する。table外・別sheetの行コンテキスト、複雑なstructured／external reference、sheet位置変更や追加を含む完全な参照更新は未完。
- [x] G3 診断 BUILD: parsed formula graphの循環有無を検出し、WASM `diagnoseWorkbook`へ`hasFormulaCycle`を追加した。再計算のbest-effort挙動は維持し、循環の解消・Excel oracle照合・structured／named rangeを含む完全診断は未完。
- [x] G3 安全境界 BUILD: structured referenceの評価時正規化で文字列リテラルと識別子境界を考慮し、unsupported式の部分一致による誤変換を防止した。
- [x] G3 型意味論 BUILD: `IF`の条件式に含まれるExcel Error値を伝播する回帰を追加し、未選択分岐の遅延評価を維持した。Empty／Error全関数の独立oracle校正、日付系・丸めは未完。
- [x] G3 型意味論 BUILD: `NOT`のError引数もErrorとして伝播させる回帰を追加した。AND／ORを含む全関数のError優先順位とExcel oracle照合は未完。
- [x] G3 型意味論 BUILD: `AND`／`OR`で全引数のErrorを保持・伝播し、真偽値による途中短絡でErrorを隠さない回帰を追加した。全関数のError優先順位とExcel oracle照合は未完。
- [x] G3 型意味論 BUILD: `XOR`でも全引数のErrorを保持・伝播する回帰を追加した。全関数のError優先順位とExcel oracle照合は未完。
- [x] G3 決定性 BUILD: workbook formula graphのready nodeと循環残りを安定ソートし、HashMap列挙順に依存しない再計算・診断順を固定した。
- [x] G3 循環診断 BUILD: qualifiedなcross-sheet cell referenceを含む循環を診断グラフで検出する回帰を追加した。named／structured rangeを含む完全診断とExcel oracle照合は未完。

- [ ] 安定sheet ID付き参照解決、workbook/sheet-local name、構造化参照を段階導入。現在の単一sheet lookupと誤って混在させない。
- [ ] シート横断dirty graph、循環検出、manual→automatic、削除／rename、cached valueの扱いを統合。既存の総work・深さ・参照budgetを維持する。
- [ ] 型変換、Empty／Error、1900/1904日付、丸め、IF/IFERRORの遅延評価を独立期待値で校正する。
- [ ] MEASURE: 複数sheetの鎖／fan-out／循環でfull再走査との一致、p50/p95・CPU・RSSを測定する。

### G4 — 関数・配列互換性の拡張（X3）

- [ ] dispatcherからcanonical関数名とaliasを棚卸しし、FUNCTIONS・引数形・未対応mode・oracle fixtureとの対応を検査する。aliasを水増し計上しない。
- [x] G4 dispatcher棚卸し基盤: `scripts/check-formula-dispatch.py`で実際の`eval_func`からcanonical名・alias・重複名を抽出する検査を追加した。FUNCTIONS・引数形・未対応mode・oracle fixtureとの対応検査は未完。
- [x] G4 dispatcher文書対応: `--check-docs`でworksheet関数表と実dispatchの218名（canonical 207 / alias 11）を照合し、未記載・stale記載をエラーにするlocal gateを追加した。引数形・未対応mode・oracle fixtureの対応検査は未完。
- [x] G4 lookup引数形: `XLOOKUP` / `XMATCH`のmatch/search modeを整数値だけ受け付け、非整数・非有限・i32範囲外を`#VALUE!`にした。全関数の引数形検査は未完。
- [x] G4 INDEX配列形式: `INDEX`で行番号0・列番号0をそれぞれ列配列・行配列として扱い、両方0では全範囲配列を返す経路を追加した。worksheetへのspill配置・shape metadata・Excel oracle校正は未完。
- [x] G4 動的配列shape基盤: `elixcee-types`に`ArrayShape`と`Variant::array_shape()`を追加し、既存のflat `Variant::Array`を維持したまま、非空配列は1行、空配列は占有領域なしとするspill footprint契約を固定した。二次元shape生成・spill配置・衝突／依存更新は未完。
- [x] G4 spill矩形基盤: 1-based worksheet座標の`SpillRect`、境界検証、セル位置計算、空配列の無占有、矩形衝突判定を共有型へ追加した。実際のworksheet書込み・既存セルとの`#SPILL!`判定・依存更新は未完。
- [x] G4 spill計画API: VMに`plan_spill_for_value()`を追加し、配列結果の占有矩形を算出して既存の非Emptyセル衝突を`#SPILL!`相当のエラーとして事前検出できるようにした。実際のspill値配置、anchor以外のCellContent生成、依存更新は未完。
- [x] G4 spill適用API: VMに`apply_spill_for_value()`を追加し、計画成功後にのみanchorとspillセルを一括反映し、anchorのformulaを保持するようにした。自動再計算への接続、旧spill領域の追跡／解放、二次元shapeは未完。
- [x] G4 spill所有権: anchorごとのspill矩形をVMとedit historyで追跡し、同一anchorの再適用時だけ旧spillセルを再利用、範囲外の旧セルを解放、undo/redoで所有情報を復元するようにした。自動再計算接続、複数行shape、spill表示のExcel oracle校正は未完。
- [x] G4 spill stale lifecycle: 旧spill範囲内の手動編集とanchor formulaの差し替えで、該当anchorの所有矩形とspillセルを先に解放するようにした。別anchorとの依存判定、自動再計算接続、複数行shapeは未完。
- [x] G4 spill依存順序: formula graphのcell／range依存でspill先を同じanchor formulaへの依存へ解決し、spill参照formulaがanchorより後に評価されるtopological orderを追加した。実際の自動spill再計算、複数行shape、別sheet依存は未完。
- [x] G4 spill再計算API: `recalculate_all_with_spills()`で各sheetのformula再計算、配列anchorの事前衝突検査／一括spill適用、依存先の再計算を接続し、呼出し元のactive sheetを復元するようにした。既存`recalculate_all()`は互換性のため変更せず、複数行shape・Excel oracle校正は未完。
- [x] G4 2D spill適用: `apply_spill_matrix()`で矩形matrixの行幅検証と`ArrayShape(rows, cols)`を反映し、2行1列などflat表現では失われる方向情報を保持できるようにした。formula evaluatorの2D metadata生成、既定再計算接続、Excel oracle校正は未完。
- [x] G4 formula shape接続: `recalculate_all_with_spills()`で`SEQUENCE` / `RANDARRAY` / `TRANSPOSE` / `WRAPCOLS` / `WRAPROWS`のflat配列結果からshapeを復元し、既定spill経路でも2D矩形として依存セルへ反映するようにした。関数全体のshape伝播、Excel oracle校正は未完。
- [x] G4 stack shape伝播: `VSTACK` / `HSTACK`の引数formula shapeを復元し、同幅・同高の2D結果を既定spill再計算へ接続した。異幅・異高入力は不足領域を`#N/A`でpaddingする。全配列関数のshape伝播は未完。
- [x] G4 FILTER shape伝播: 元の2D range幅と出力要素数から`FILTER`の行列shapeを復元し、2D抽出結果を既定spill再計算へ接続した。include形状の全組合せ、異幅入力、Excel oracle校正は未完。
- [x] G4 TAKE/DROP shape接続: 元配列shapeに基づく行・列単位のTAKE/DROPと任意列数引数を追加し、2D結果を既定spill再計算へ接続した。異常引数、全shape伝播、Excel oracle校正は未完。
- [x] G4 axis shape伝播: 1列sourceの`UNIQUE` / `SORT`と、`TOCOL` / `TOROW`の出力軸を既定spill再計算へ接続した。2D sourceの全意味論、全shape伝播、Excel oracle校正は未完。
- [x] G4 INDEX array shape接続: `INDEX(range,0,0)`の全範囲、行配列、列配列について元rangeと引数からshapeを復元し、既定spill再計算へ接続した。2D切出し全体、Excel oracle校正は未完。
- [x] G4 choose axis接続: 2D入力の`CHOOSECOLS` / `CHOOSEROWS`を行列単位で選択し、選択後shapeを既定spill再計算へ接続した。異常引数、1D legacy経路、Excel oracle校正は未完。
- [x] G4 2D unique/sort接続: 2D入力の`UNIQUE`を行／列単位の重複排除、`SORT`を行列のsort_index・sort_order・by_colに接続し、結果shapeを既定spill再計算へ接続した。exactly_once、型変換、Excel oracle校正は未完。
- [x] G4 2D SORTBY接続: 2Dデータを同じ行のsort-by列で並べ替え、行数・1列sort-byの整合性を検証して結果shapeを既定spill再計算へ接続した。複数sort-by配列、型変換、Excel oracle校正は未完。
- [x] G4 2D FILTER接続: Range既存経路を維持しつつ、生成された2D配列を行includeで抽出し、元配列幅から結果shapeを既定spill再計算へ接続した。列include、異幅入力、Excel oracle校正は未完。
- [ ] 第1組は参照／条件集計／検索、次に日付／統計／金融。既存SUMIFS・XLOOKUP等を再実装せず、未対応modeと意味論差分から埋める。
- [x] G4 第1組の意味論補完: `IFNA`を追加し、`#N/A`だけをfallback対象として、それ以外のErrorは伝播する遅延評価を回帰テストで固定した。Excel oracleとの一致確認と、他関数の未対応mode棚卸しは未完。
- [x] G4 検索modeの安全境界: `XLOOKUP`のwildcard `match_mode=2`と、ソート済み数値範囲向けbinary `search_mode=2/-2`を追加した。wildcardとbinaryの組み合わせや未知modeは明示エラーにし、順序を満たさない入力を推測処理しない。Excel oracleとbinary modeの網羅的校正は未完。
- [x] G4 検索関数の整合: `XMATCH`にもwildcard `match_mode=2`と、ソート済み数値範囲向けbinary `search_mode=2/-2`を追加し、未知modeの黙った通常検索を廃止した。Excel oracleと検索関数全体のmode網羅性は未完。
- [x] G4 lookup境界: `HLOOKUP`の行番号0を`#VALUE!`、範囲外を`#REF!`として`VLOOKUP`と同じ明示的な境界に揃えた。Excel oracleによる追加の型変換校正は未完。
- [x] G4 lookup境界: `INDEX`の行列番号0を`#VALUE!`、範囲外を`#REF!`として明示的に検出し、整数アンダーフローによる誤セル参照を防いだ。配列形式・Excel oracle校正は未完。
- [x] G4 MATCH exact互換性: `match_type=0`で文字列wildcard `*` / `?`を大文字小文字非依存で評価し、未定義match typeは`#VALUE!`として返すようにした。Excel oracleと近似matchの並び順校正は未完。
- [x] G4 MATCH wildcard補完: `~*` / `~?` / `~~`をリテラルwildcardとして扱うbounded DP経路を追加した。criteria系を含む全wildcard仕様のExcel oracle校正は未完。
- [ ] 動的配列のspill衝突・shape・依存更新を整えてから、既存FILTER／LET／LAMBDA等の配列経路を拡張する。
- [ ] MEASURE: Excel／EPPlus／Aspose.Cellsはversion・計算設定・license利用条件を固定して比較。実行していない公式対応表と、実測一致率を分ける。

### G5 — 保存メモリの段階削減（X4 / X5）

進行中。G5c/G5dと共通のabort実装は完了。G5a/G5bとRSS測定が残っているため、
G5全体は未完了であり、constant-memoryを主張しない。

- [x] G5a 部分 BUILD: 画像・VBAと、関係解析対象外のXML payloadをraw mapへ保持せず、元ZIPから一件ずつdestinationへ直接copyする経路へ変更した。未編集table XMLと、新規table追加が無い場合のworksheet `.rels` XMLも遅延copyへ移し、編集対象tableや新規table用relsだけを保持してpatchする。worksheet source XMLは保存ループで1シートずつ所有権移動して処理済みpayloadを解放し、workbook XMLもowned fragment抽出後に元全文を解放する。copy前のZIP再検証とentry期待サイズ一致を確認し、reader回帰テストで1 MiBの画像payloadが索引上空のまま保持されることも固定した。workbook rels等の完全遅延化、元file全体の同一性測定、CRC／3 OS検証は未完了のため、G5a全体は未完了。
- [x] G5a 部分 BUILD: defined namesの読込・報告は`xl/workbook.xml`だけを検証付きで直接取得する経路へ分離し、worksheet/table等の兄弟XMLを再読込・保持しない回帰を追加した。これは保存時の全part遅延化や3 OS RSS測定を完了扱いにしない。
- [x] G5a 部分 BUILD: 未編集のtable XMLを保存時のraw mapへ保持せず、table編集がある場合だけ対象partを元ZIPから再取得してpatchする経路へ変更した。全passthrough partの遅延化、保存RSS測定、3 OS検証は未完了。
- [x] G5b 部分 BUILD: worksheet relationship XMLの解析用文字列索引を専用mapから統合relationship mapへまとめ、同一`.rels` payloadの解析用二重保持を削減した。passthrough bytes自体の遅延化とRSS測定は未完了。
- [x] G5b 部分 BUILD: relationship XMLのraw bytesを保存時のpassthrough mapへ保持せず、接続解析後は元ZIPから遅延copyするようにした。新規tableのworksheet `.rels`だけは解析済みXMLをpatchして出力し、重複entryを防止する。全payload遅延化とRSS測定は未完了。
- [x] G5b 部分 BUILD: relationship接続・pruning解析後に`raw_entries`側の`.rels` bytesを解放し、解析用文字列索引とentry名だけを保持するようにした。最終出力は元ZIPから再取得するため、全payload遅延化とRSS測定は未完了。
- [x] G5b 部分 BUILD: shared-string本文を`Vec<String>`へ二重保持せず、所有するindexから参照を座標順に並べて直接出力する経路へ変更した。未編集styles XMLはstyle解決後に解放し、元ZIPから直接copyする。passthrough XMLも出力時に一件ずつdrainして処理済みpayloadを解放する。styles/workbook/[Content_Types]/worksheetのwriter-owned cloneも所有権移動し、cell/row/column style編集時は対象sheetだけをoverlay cloneするようにした。font/fill/borderを使わないstyle編集ではcellXfsだけを展開する。その他の補助索引とdisk spoolは未完了のため、G5b全体は未完了。
- [x] G5c BUILD: 追記専用APIに`create_stream_bounded`を追加し、「1行上限」と「総work量」を分けた。既存`max_pending_bytes`の累積制限は維持し、stub・README・limitsへ移行例を追加した。
- [x] G5d BUILD: bounded追記Writerは固定worksheet構造・列上限・row buffer・64 KiB codec bufferを使い、inline stringsでunique stringsを蓄積しない。通常VMの全セル保持は対象外。
- [x] 共通 BUILD: I/O失敗・context例外時のabort、temp削除、close失敗後の状態、Windowsのclose-before-rename、標準syncの経路を実装した。Windows実機検証は未完。
- [x] 部分 MEASURE: macOS arm64／CPython 3.13の別processで、appendとtransaction付きnormal-freshを100,000・250,000・1,000,000行まで測定し、RSS・wall・temp/output・ZIP／worksheet／最終行検証を記録した。Linux/Windows、Excel oracle、full matrixは未完了。
- [ ] MEASURE: 通常保存と追記を別processで測り、10万→100万行のRSS増分・p95・temp disk・出力同値を記録。通常VMの全セル保持までconstant-memoryと呼ばない。

追加測定（2026-09-07、macOS arm64、CPython 3.13、現行1.0.4 release wheel）では、appendの100,000／250,000行を各2回完了し、RSS p50/p95は19.39/19.83 MiB、19.38/19.46 MiB、出力検証は全件成功した。normal-VMは1行ごとのundo snapshotによる入力オーバーヘッドと保存メモリを分離するため4,096行単位の`set_range`へ変更し、さらに全batchを1 transactionへまとめて再測定した。100,000行はRSS 116.92/118.30 MiB・wall 173/202 ms、250,000行はRSS 203.56/203.61 MiB・wall 418/420 ms（各p50/p95）だった。全件のZIP・worksheet形状・最終行・出力検証は成功した。これはundo履歴の全体clone回数を抑えた効果であり、normal VMのconstant-memoryの根拠にはしない。append APIの結果と直接比較せず、3 OS・100万行・Excel oracle・通常保存の完全matrixは未完了とする。詳細は[測定記録](docs/measurements/writer-streaming-2026-09-07.md)。
追加測定（同条件、1,000,000行、各2回）も完了し、appendはRSS 19.38/19.44 MiB・wall 2,750/2,789 ms、normal-fresh（一つのtransaction）はRSS 751.56/753.20 MiB・wall 2,185/3,829 ms（各p50/p95）だった。temp/outputはappend 14.24/14.24 MiB、normal-fresh 12.43/12.43 MiBで、ZIP・worksheet形状・最終行・出力検証は全件成功した。1M行でtransaction経路を確認できたが、normal VMのconstant-memoryや3 OS対応の根拠にはしない。append APIの結果と直接比較せず、Linux/Windows・Excel oracle・通常保存の完全matrixは未完了とする。詳細は[測定記録](docs/measurements/writer-streaming-2026-09-07.md)。

測定入口: `python3 scripts/measure-stream-writer-memory.py --mode append --rows 100000 250000 1000000`。
入力形状は`--value-profile plain|escape|giant`で切り替え、XML escape膨張と1 MiB単一文字列を
行数スケール測定と分離して校正する。
RSS取得はmacOS/Linuxでは`ru_maxrss`、WindowsではWin32のpeak working setを使うため、
同じchild-process harnessを3 OSで実行できる。
`.github/workflows/ci.yml`には、通常CIを重くしない`workflow_dispatch`限定の
Ubuntu/macOS/Windows測定matrixを追加した。Actions実行と結果の固定は未完了である。
測定スクリプトの`--output`でJSONを保存し、workflowはOS別artifactとして保管する。
`check-stream-writer-measurements.py`でschema、p50/p95、全sampleの出力検証を自動検査する。
validator自身の正常系・schema破損・p95逆転・未検証sampleのself-testもCIで実行する。
遠隔の直近成功CI（release commit `43cf3d0`）にはこの手動G5 jobがまだ含まれないため、
既存CI成功をG5の3 OS測定根拠には昇格しない。
通常保存は`--mode normal`で同じ別process harnessを使う。
`--repetitions N`を指定すると各caseのp50/p95を集計し、caseディレクトリの
peak temp disk使用量も記録する（10ms pollingの観測値）。
出力は別processのpeak RSS、wall time、出力byte、行数を保存し、同一入力を再読込して
値一致を確認してから記録する。未実行・wheel未確認の結果はG5完了根拠にしない。

部分測定（2026-09-06、macOS arm64、CPython 3.13、修正版wheel）:

| 追記行数 | peak RSS | wall time | 出力 | 検証 |
|---:|---:|---:|---:|---|
| 100,000 | 18.33 MiB | 325 ms | 1.43 MiB | ZIP／最終行 OK |
| 250,000 | 18.45 MiB | 752 ms | 3.68 MiB | ZIP／最終行 OK |
| 1,000,000 | 18.42 MiB | 2,901 ms | 14.93 MiB | ZIP／最終行 OK |

追記経路のこの3点ではRSS増分は約0.08 MiBだったが、Python入力・allocator・ZIP
metadataの影響を含む一環境の結果であり、通常Writerのconstant-memoryや3 OS対応を示さない。

通常保存の部分測定（2026-09-06、macOS arm64、CPython 3.13、1列）:

| 行数 | peak RSS | wall time | 出力 | 検証 |
|---:|---:|---:|---:|---|
| 10,000 | 22.01 MiB | 37 ms | 61.2 KiB | ZIP／最終行／再読込 OK |
| 25,000 | 26.06 MiB | 75 ms | 149.4 KiB | ZIP／最終行／再読込 OK |
| 50,000 | 33.63 MiB | 133 ms | 295.0 KiB | ZIP／最終行／再読込 OK |

100,000行以上の通常保存は、現行ReaderのXML要素上限（1,000,000）に先に達するため、
この条件では未測定。上限を緩めることは安全性仕様の変更になるため、G5完了条件とは
別に扱う。

新規VMの通常Writer測定（同日、3列、`--mode normal-fresh`）:

| 行数 | peak RSS | wall time | 出力 | 検証 |
|---:|---:|---:|---:|---|
| 100,000 | 109.25 MiB | 182 ms | 1.24 MiB | ZIP／最終行 OK |
| 250,000 | 200.42 MiB | 436 ms | 3.11 MiB | ZIP／最終行 OK |
| 1,000,000 | 748.54 MiB | 1,802 ms | 12.43 MiB | ZIP／最終行 OK |

この行列はVM全セル保持を含むためRSSは行数依存であり、通常VMをconstant-memoryとは呼ばない。

入力形状校正（2026-09-06、macOS arm64、CPython 3.13、追記、各3回）:

| profile | 行数×列数 | peak RSS p50/p95 | peak temp disk | 出力検証 |
|---|---:|---:|---:|---|
| escape | 1,000×2 | 19.39/19.44 MiB | 561.5 KiB | ZIP／最終行 OK |
| giant | 10×2 | 31.62/32.52 MiB | 12.5 KiB | ZIP／最終行 OK |

`giant`の1 MiB文字列は同一内容のためDeflate後の出力が小さくなる。これは入力側の
文字列・escape・allocatorの校正であり、constant-memoryや3 OS対応の証拠ではない。

1M行の反復測定（同日、append 3回／normal-fresh 2回）では、appendのwall p50/p95が
2,718/2,935 ms、peak RSS p50/p95が18.55/18.59 MiB、normal-freshのwall p50/p95が
1,783/2,672 ms、peak RSS p50/p95が748.55/750.44 MiBだった。いずれも出力検証に成功した。
temp directoryの1M行観測はappend 3回で14.24 MiB、normal-fresh単発で12.43 MiB
（いずれも最終出力サイズと一致）だった。監視を含むwall timeはI/O負荷で大きく変動するため、
前段のp50/p95と混ぜず別観測として扱う。

### G6 — 互換性・資源・公開判定

- [ ] G2の接続graphとExcel再open、G3–G4の独立oracle、G5の行数別RSSを根拠として、対応matrixと既知の損失を更新する。
- [ ] Linux/macOS/Windows、失敗時の元出力保護、fuzz・依存監査・Python/Rust API回帰を実施する。
- [ ] 大規模速度は同じ入力・編集・耐久性・反復条件で再測定。互換性やRSSの悪化を速度向上で相殺しない。

### LogiSheets 対抗トラック（L0–L6）

LogiSheetsの公開metadataは実装・測定・運用実績の証拠ではないため、
「公開API」「実行可能なfixture」「3 runtime」「速度・依存更新・履歴」の軸を分けて比較する。
既存のX/Gゲートを置き換えず、数式とWASMの共有を優先する。

- [x] L0 現状固定: Rust/WASM/Node/ブラウザのread経路は共有済み。数式はsingle-sheet fast pathとworkbook qualified-reference slow path、undo/redoとJS write/recalculateは未完として固定した。
- [x] L1 workbook formula BUILD: sheet-qualified cell/range referenceを明示的なsheet bandへremapし、全sheetのformula nodeを依存順に評価するslow pathを追加。case-insensitive sheet名、mixed host/qualified参照、formula chain、cycleのbest-effortを単体テストで固定した。dirty graphの増分化とstructured referenceは未完。
- [x] L2a named range BUILD: runtime named rangeと、ロード時に取り込んだ単純A1／qualified／`OFFSET` definedNameをformula ASTへ展開し、qualified referenceと混在する`SUM(MyRange)`をworkbook再計算で評価する。table column、連続複数列、静的specifier、同一tableデータ行のthis-row structured referenceを同じ依存グラフへ接続し、行／列構造変更とrenameの単純参照更新にも対応する。table外の行コンテキスト依存・複雑なstructured referenceは未完。
- [x] L2b interval-style dependency BUILD: 大きなrangeを全セルの依存キーへ展開せず、sheet・行・列の区間として保持し、sheet-local formula-node indexとの包含判定で依存辺を構築する。完全な更新用interval treeとincremental workbook dirty propagationは未完。
- [x] L2 dependency graph BUILD（部分完了）: range依存の過剰展開を避けるsheet-local interval-style index、値セルを起点にした直接・range・formula-chainのdirty closure、manual/automatic再計算をworkbook単位で統合した。full interval tree、sheet rename/delete、sheet-scoped/dynamic/structured named range、循環診断、完全な構造変更追跡は未完。
- [x] L3 shared runtime BUILD（部分完了）: 同一Rust coreのworkbook計算を`calculateWorkbook(bytes)`としてWASMへ公開し、`diagnoseWorkbook(bytes)`でsheet/formula/qualified formula/parse errorのJSON summaryを提供、Node/browser同梱runtimeを再生成した。cross-sheet formulaのNode実行、`@elixcee/xlsx/runtime`のNode/browser条件、CJS/ESM bundle、意図的に固定したWASM payload baselineに対する10%成長gateを確認した。incremental JS API、sync/async loading、browser worker境界は未完。
- [x] L4 editing history BUILD（部分完了）: 明示的なセル/範囲値書き込みとセル数式設定を対象に、最大128操作のbounded undo/redoとtransaction abortをRust VM・Python APIへ追加し、transaction commitは全編集を1つのundo単位として記録することで大規模範囲編集の全体snapshot増殖を抑えた。transaction開始時も既存redo履歴を複製せず、abortでは履歴を保持する。formula cache・dirty状態・tile cacheの復元、prior history保持、nested transaction拒否をテストした。WASMにはstateful `WorkbookEditor`（`setNumber`、`recalculate`、`undo`/`redo`、transaction）を追加し、互換rootを汚さない`@elixcee/xlsx/runtime`サブパスからNode/browserへ公開した。sheet操作、OOXML dirty partsと外部効果の一体復元は未完。
- [ ] L5 structured data/plugin/AI boundary: table/validationをtyped data APIへ写像し、pluginはcapability allowlistとresource budget内で実行する。AI操作は提案・dry-run・差分承認を既定にし、任意コード実行や外部取得を許可しない。
- [ ] L6 MEASURE/GATE: LogiSheetsを固定commit/versionで比較し、formula correctness、dependency update、WASM/Node/browser parity、undo/redo、startup/throughput/RSS/bundle sizeを同じfixtureで測る。WASM payload baselineと10%成長gate、runtime subpathのNode/browser条件、CJS/ESM bundle smokeは先行実装した。LogiSheets固定版比較、同一fixtureの速度/RSS/値parity、Excel oracleは未完。stars/commitsや機能数は補助情報に留め、未測定の優位性は主張しない。

直近の実装成果（2026-09-06）: L1のworkbook再計算、L2aのnamed range展開、L2bの区間依存辺構築、L2のdirty closure（値セルからの逆引き、formula chain、range）と`set_cell_formula`のqualified-reference初期値処理、L3の`calculateWorkbook(bytes)`／`diagnoseWorkbook(bytes)`共有WASM入口、L4のbounded undo/redo／transaction abort／WASM `WorkbookEditor`と`@elixcee/xlsx/runtime`公開サブパスを実装。
`cargo test --workspace --all-targets --offline`、`cargo clippy -p elixcee --all-targets --offline -- -D warnings`、`packages/xlsx`の`npm run wasm:smoke`を確認した。
これはBUILDとconsumer smokeの証拠であり、LogiSheetsとの速度比較、WASM/Node/browserの完全な値parity、Excel oracle一致を示すものではない。

比較仕様の参照先（実測証拠ではありません）:
[EPPlus公式対応表](https://github.com/EPPlusSoftware/EPPlus/wiki/Supported-Functions)、
[Aspose.Cells計算仕様](https://docs.aspose.com/cells/net/calculate-formulas/)、
[OOXML PivotCaches](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.pivotcaches?view=openxml-3.0.1)。

## 性能バックログ（優先順）

| 優先度 | 未完了の作業 | 完了条件 |
|---|---|---|
| P0 | 元ZIPのpassthrough entryを全量展開・保持しないWriter | 大規模partのclone削減、peak RSS・保存時間、全part／relationship同値 |
| P0 | 小規模ファイルの1.2倍目標 | 17セルの固定保存費用を改善し、標準sync・atomic renameを維持。quiet-hostでp95も確認 |
| P0 | 大規模全条件で追加1.2倍 | 10万・40万・100万セルの中央値比で判定。style／数式密度・RSS・単一sheetの上限も校正 |
| P0 | formula dirty propagationの完全校正 | 大規模controlled matrix、full再走査との値一致、p50/p95、CPU/RSS、循環・manual/automatic |

P0速度の現行ベースライン（2026-09-07、Criterion 10 samples／2 seconds、現行1.0.4）では、5,000行appendのreference rescanが106.88 ms、cached pathが94.98 msだった。これは当該microbenchmarkで約11.1%短縮した観測であり、1.2倍目標やend-to-end優位性の根拠にはしない。詳細は[VMホットパス測定記録](docs/measurements/vm-hotpath-optimization-2026-09-05.md)。
formula dirty propagationの同日controlled matrixでは、single-input chain 1.2825 ms、warm noop 1.2363 ms、structure rebuild 1.3090 ms、独立1,000入力 1.1570 ms（各median）だった。single-inputがrebuildを上回る短縮は確認できなかったため、closure bookkeepingが支配的になる条件のnegative resultとして記録し、P0最適化候補を維持する。
その後、依存先を座標ではなくformula plan indexでqueueへ渡す局所変更を実装し、single-input chain 1.2538 ms、structure rebuild 1.3223 msを再測定した。直前chain比で約2.2%の改善だが、rebuildは不変であり、独立入力・end-to-end・1.2倍目標の根拠にはしない。range/cycle回帰は成功し、独立入力は未再測定。
残るwarm noopは1.2012 ms、独立1,000入力は1.1502 msで、いずれもCriterion上の有意差なしだった。queue index化はsingle-input chainに限る局所改善として確定し、一般的なdirty propagation高速化とは扱わない。

2026-09-07のcandidate健全性確認では、`cargo test --workspace --all-targets --offline`を実行し、Rust unit 1,549件、blackbox、CLI、property、XLSX round-trip 52件、bench smoke、WASM crateを含む全targetが成功した。これはlocal regression evidenceであり、3 OS、Excel oracle、公開artifactの証拠ではない。
同日の`cargo clippy --workspace --all-targets --offline -- -D warnings`も成功し、workspace全targetで警告は検出されなかった。これは静的検査の証拠であり、3 OS・Excel oracle・公開artifactの証拠ではない。
主要workflow（CI、publish、release、crates-publish）は`actionlint`でエラーなしだった。workflow定義の静的検査であり、GitHub Actionsの実行結果や3 OS測定完了を意味しない。
測定境界検査（`scripts/check-measurement-boundary.sh`）とreader／stream writer validatorのself-testも成功した。これは未検証測定を公開artifactへ混入させないためのlocal gateであり、3 OS実測やExcel oracleを代替しない。
追跡済みformula planでは整合性scanを省略し、`cells_mut()`等でtrackingが無効化された場合だけlive formula再検査へfallbackする変更を実装した。dirty formula、cycle、manual→automaticを含む関連テストとworkspace全target回帰は成功した。
同一Criterion条件の再測定ではsingle-input chain 1.0217 ms、warm noop 0.96091 ms、独立1,000入力 0.89674 msとなり、直前baseline比で約18.5%・20.0%・22.0%改善した。structure rebuildは1.3737 msで約3.9%増のため、tracked-edit fast pathの局所改善として扱い、一般的な再計算高速化とは主張しない。
| P1 | VBA bytecode／symbol intern | AST経路との一致、演算主体・セルI/O主体の両方を測定 |
| P1 | canonical packed/tiled cell storage | 公開HashMap APIを保つoverlay。Range・formula・style・save・fork・GCの回帰 |
| P1 | Writerのstyle／shared-string overlay | 全map／文字列cloneを削減し、密度別RSS・CPU・出力同値を確認 |
| P2 | Range中間grid削減 | Sort・転置・広域copyでrow permutation／tile／iteratorを検討。1-based・Empty・formulaを保持 |
| P2 | Python GIL解放・複数VM並列 | Python objectを触らない区間を対象にthroughput・RSS・例外伝播を確認 |
| P2 | byte指向lexer | char配列・一時文字列を削減。Unicode識別子と診断位置を保持 |
| P2 | object GCの世代別／dirty graph化 | Class_Terminateの一度性、Collection cycle、budget、長時間soak |

### 直近の性能実績

- [x] Readerのborrow化・単一走査・VMへの所有権移動、read budget／deadline／協調キャンセル。
- [x] worksheet XMLのZipWriterへの直接出力、passthrough payloadの二重clone除去。
  **raw ZIP全体の遅延展開は未完了**。
- [x] formula dirty/full比較、100／1,000式matrix、manual→automatic・循環参照、CPU／RSS、30反復の校正。
  完全校正ではありません。[測定一覧](docs/measurements/README.md)
- [x] durable saveのbuffering・seek不要ZIP出力、opaque XML検索とshared-string収集の最適化。
- [x] 行別treeから座標sortへの変更、セル番地のstack生成、Readerの属性・型buffer再利用。
  [大規模測定](docs/benchmarks/workbook-large-speedup-2026-09-06.md)
- [x] 固定版openpyxl・ClosedXMLとの同条件比較と、撤回した初期比較を含む生データの保存。
  [比較一覧](docs/benchmarks/README.md)

2026-09-06の直前実装に対する大規模20組の交互測定:

| セル数 | before / afterの全体中央値比 | 厳密な1.2倍目標 |
|---|---:|---|
| 100,000（1 sheet） | 1.180 | 未達 |
| 400,000（1 sheet） | 1.318 | 達成 |
| 1,000,000（250,000 × 4 sheets） | 1.196 | 未達 |

これはローカルの特定入力・特定baselineに限る結果です。1.196を丸めて達成扱いしません。
中央値比とpaired speedupは別の指標です。既存XML要素数制限は緩めていません。
過去の競合比を掛け合わせて、最新実装の競合比とすることもしません。

## XLSX編集互換性トラック

各行は残る評価・拡張範囲です。既存APIがあることと、全ケースの互換性を証明したことを分けます。

| ID | BUILD | MEASURE / GATE |
|---|---|---|
| X0 棚卸し | read／edit／write／round-trip／Excel確認の機能matrixを維持 | 未実装・保持のみ・拒否・未検証をfixtureと結び付ける |
| X1 基本モデル | セル型・数式・style・merge・行列・sheet操作とdirty part管理を整合 | 挿入／削除／移動／copy・大規模疎密入力・参照追従 |
| X2 構造化オブジェクト | table・filter・validation・conditional formatting・defined name等を段階拡張 | owner XMLとrelationshipを含む往復、Excel実機で修復警告・意味を確認 |
| X3 数式・マクロ共存 | A1/RC・sheet間参照・named range・cached value・再計算方針 | エラー／日付／型変換・manual/automatic・VBA binary保持。保持と実行を区別 |
| X4 保存・復旧 | 未知part保持、dirty part再生成、transactional save | 中断・失敗時に元ファイルを保護。content type／relationship／署名等の限界を明示 |
| X5 性能・API | streaming、Python/Rust/JS境界、互換API差分を整理 | 固定版競合、疎密／数式／style／多sheet、CPU・RSS・p50/p95・出力サイズ |
| X6 リリース | サポートmatrix・既知の損失・migration例を固定 | clean-install、3 OS、回帰・ライセンス・配布物・公開情報を別々に確認 |

入力には大きさだけでなくXML・ZIP・work budgetの上限があります。
[limits](docs/limits.md) と [architecture](docs/xlsx-architecture.md) を参照してください。

## VBA・数式の残作業

- [x] Range相対参照、default Item/Value、Worksheet/Workbookの基本member、SpecialCells拡張。
- [x] VM-local Collectionとexport済みclass moduleのobject連携、Property・interface dispatch。
- [x] VM-local Dictionary adapter。キー正規化などの制限は [FUNCTIONS](FUNCTIONS.md#in-memory-dictionary) に明記。
- [ ] Save／Close／外部リンクなど作用を持つmemberは、安全境界とoracle fixtureを先に定義して段階実装。
- [ ] Workbook_Open／Worksheet_Change等のイベントは、EnableEvents・再入抑止・決定的dispatch順・budget付きで設計。
- [ ] runtime errorの文字列分類依存を減らし、型付き診断へ移行。
- [ ] 複数module間のUDT名衝突と診断位置など、解決規則の残差を明確化。
- [ ] Excelとの型変換・丸め・日付・Empty／Error・配列境界・再計算の独立oracle比較を拡張。
- [ ] 実運用macroの回帰と、数式の依存グラフ・循環・volatile／dynamic array等の対応境界を検証。

ローカルsynthetic fixtureの通過数を、実Excelの意味論一致件数として扱いません。
[互換性ハーネス](compat/README.md) と [CLI契約](docs/agent-contract.md) が検証の入口です。

## セキュリティトラック

| ID | 維持・追加する作業 | 残る検証 |
|---|---|---|
| S0 脅威モデル | 入力、出力、保存先、formula／VBA、互換API、秘密情報の境界 | 攻撃面と失敗時の非破壊性をfixtureへ対応付ける |
| S1 ZIP/XML | path／relationship検証、暗号化・特殊entry・不正XML拒否、資源上限 | limit近傍・圧縮爆弾・中断・大規模の3 OS校正 |
| S2 parser/formula/VM | 深さ・式・配列・実行budget、型変換とエラーを検査 | adversarial corpus、fuzz、長時間CPU/RSS、panic／hang／leak |
| S3 出力・互換API | prototype安全性、HTML既定escape、明示的rawHtml、CSV／外部参照境界 | 悪意ある入力の安全な差分、保存失敗・relationship保持 |
| S4 隔離・依存 | 最小権限、外部効果遮断、必要に応じprocess／container隔離 | worker強制終了と協調cancelの区別、依存監査・供給網・clean install |
| S5 継続ゲート | 回帰corpus、fuzz、配布物とsecurity policy | freshなadvisory、license、3 OS、長時間soak。未検証は明記 |

- [x] ZIP/XML構文・entry／part／全体・work budget、協調中断のローカルfixture。
- [x] reader資源回収・read/mutate/write、semantic validatorと拒否self-test、Cargo測定境界の検査。
- [ ] Linux／Windowsを含む資源校正と、長時間fuzz／CPU／RSS・隔離環境検証を完了する。

既存のmacOS測定と安全策は [測定記録](docs/measurements/README.md)、
[脅威モデル](docs/xlsx-security-model.md) に集約します。限界値の緩和を達成条件にしません。

## 配布・サポートゲート

1.0.3のローカル検証（2026-09-06、macOS）:

- [x] Rust全workspace／全targetの1,709 tests、strict clippy／Rustdoc、依存監査、測定記録検証。
- [x] crate検証、wheel／sdist生成、独立Python環境でのimport・VBA・通常／fast保存往復。
- [x] JS型／differential／WASM bundle／packed consumerと、VBA corpus 581件・意味論386件の既存ゲート。

以下は継続的な配布・サポート要件であり、上記のローカル検証だけでは完了しません。

- [ ] Python wheel／sdist、Rust crate、CLIのclean-installと最小利用例を各対象OSで確認。
- [ ] JS packageはutils／read／write／browser／Node／型定義を別々に検証。公開する場合は独立した判断を行う。
- [ ] license／notice／package内容／依存監査／回帰・互換性matrixを確認。
- [ ] サポート外・既知の損失・セキュリティ境界・移行例を公開文書へ反映。
- [ ] tag、registry、workflow、正式Release、worktree状態を別々に確認。

このゲートは将来のリリース判定です。文書の更新だけで完了にしません。
実装・測定は自律的なローカル作業として進め、version変更・push・公開は別の操作として扱います。
