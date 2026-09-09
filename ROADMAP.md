# elixcee Roadmap

更新日: 2026-09-09。対象versionは **1.0.5** です。
完了項目は記載した実装・測定の範囲に限ります。公開先の状態はリリースごとに別途確認します。
版ごとの変更は [CHANGELOG](CHANGELOG.md)、実装範囲は
[FUNCTIONS](FUNCTIONS.md)、保証範囲は [v1契約](docs/v1-support-contract.md) を参照してください。

## 方針と完了の定義

Excel不要の、Rust/Pythonによる安全なワークブック編集・数式計算・VBAデータ処理を一体化した自動化ランタイムを中心に開発します。
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

1.0.5公開後の開発計画です。次の変更はUnreleasedに記録します。
各PhaseはBUILDを小さく実装し、MEASUREが未完なら未検証として残します。
EPPlus／Aspose.Cellsとの一般的な同等性や、関数名の個数だけでの優劣は達成条件にしません。

### G0 — 現状・合格条件の固定（X0 / S0）

- [x] ソース棚卸し: drawing/legacyDrawingの選択的保持はあるが、workbookの`pivotCaches`／`externalReferences`は再出力していない。partの存在だけでは接続を保証できない。
- [x] 数式はA1/RC・シート修飾参照のparse/rewriteと評価を分離しており、シート修飾参照の評価は現在拒否する。
- [x] 通常Writerは元ZIP全展開・セルmap・文字列索引を保持。Python追記WriterはZIPへ逐次出力するが、行の収集とXML生成があり、`pending_bytes`は保持RSSではなく累積受け入れ量。
- [x] 機械可読matrix: [OOXML feature matrix](compat/ooxml-feature-matrix.json)でCharts / Pivot / Drawings / External Linksごとにread・preserve・edit・recalculate・Excel再openを別状態にし、既存fixtureまたは未作成状態と結び付けた。`preserved`は編集可能を意味せず、`unverified`は互換性成功に数えない。

### G1 — Writerの受け入れ上限（S1 / X5、最初のBUILD）

- [x] 行数はiterator開始前、列数と累積推定byte数はセル取り込み中に検査。空文字列でもセル本体の費用を計上する。
- [x] 不正型・iterator例外・予算超過で、拒否行のXML／行数／byteカウンタを更新しない。正常行を追加して再開できる回帰を追加する。
- [x] インストールしたwheelのPythonテストをCI設定へ追加。上限直前／一致／超過と、保存前の既存出力保護をローカル検証（この変更のGitHub CI実行は未確認）。
- [x] 部分 MEASURE: 現行1.0.4 release wheelを別processで実行し、XML escape-heavy 1,000行と1 MiB単一文字列10行を各3回測定した。peak RSSはPython入力・Rust allocator・XML／ZIP生成を含むprocess値として記録し、ZIP／worksheet形状／最終行を全件検証した。allocator内訳、Linux／Windows、通常VMのconstant-memoryは未完であり、この測定だけでconstant-memoryを主張しない。詳細は[入力校正](docs/measurements/writer-input-calibration-2026-09-09.md)。

2026-09-06 macOSローカル: [Python回帰10件](tests/python/test_stream_writer_limits.py)は
公開1.0.3で5件の問題を検出し、修正版wheelでは10/10成功。Rust全workspace／全target
1,679 tests、全feature strict clippyも成功。これはG1のBUILD検証で、RSS／Excel互換性の測定ではありません。

### G2 — OOXMLの接続を保つ（X2 / X4）

- [x] G2a 部分 BUILD: `externalReferences` owner要素を再出力し、source relationshipのType/Targetを基準に再採番後の`r:id`へ書き換える経路と、未解決relationship時にownerを省略する安全策を追加した。Pivot cache、cacheId／外部参照順序、content types／namespaceのfixture検査は未完。
- [x] G2b 部分 BUILD: worksheetのdrawing／legacyDrawing ownerについて、r:id・worksheet .rels・相対target・出力に残るpartを一組で照合し、missing target／重複relationship ID時はownerを再出力しない経路を追加した。さらにDrawing等の内部partに付随する.relsを推移的に辿り、Chart／image targetの欠落も検出する。PivotTableからcache definition／recordsまでの検証とfixture化は未完。
- [x] G2a/G2b 追加 BUILD: pivotCaches ownerを再生成workbookへ戻し、workbook relationshipの再採番とcache definition／recordsへの内部relationship到達性検査を適用した。Pivotの再集計・編集、cacheIdの意味更新、fixtureによるExcel再open検証は未完。
- [x] G2c safety BUILD: Chart/DrawingまたはPivotを含む入力に対するsheet rename／行列挿入削除などの構造編集を、参照更新経路がない場合は保存前に明示拒否するようにした。古い参照を黙って保存しない。実際のchart参照・anchor・pivot source更新はG2d以降の未完項目。
- [ ] G2d: Charts / Drawingsの作成・一般編集API、Pivotの一般source/cache更新を一機能ずつ追加。既存Chart系列のformula／cache、worksheet-backed Pivot source、two-cell Drawing anchorの限定編集は部分BUILD済み。描画再現・Pivot再集計は保持とは別の未完項目として扱う。
- [x] G2d 部分 BUILD: sheet renameに限定したChart `<c:f>`とPivot `worksheetSource@sheet`の安全な参照更新、およびDrawing／relationship chain保持を実装・回帰検証した。Chart/Drawing作成・一般編集、row/column編集に伴うanchor更新、Pivot source/cache編集・再集計は未完。
- [x] G2d Chart rename BUILD: 単純なsheet renameに限り、Chart XMLの`<c:f>`に含まれるqualified sheet referenceを既存formula parserで安全に書き換え、Drawing owner／relationshipを保持して保存する経路を実fixtureで検証した。chart creation、一般のChart/Drawing編集、Pivot更新、Excel再openは未完。
- [x] G2d Pivot rename BUILD: 単純なsheet renameに限り、Pivot cache definitionの`worksheetSource@sheet`をXML escape付きで更新し、cache本体・cacheId・table source・再集計には触れずに保存する経路と回帰を追加した。Pivot source/cacheの一般編集、再集計、Excel再openは未完。
- [x] G2d Chart series BUILD: 既存Chart XMLに対する`Vm.set_chart_series_formulas`（Python binding／型stub含む）を追加し、0-based series indexでcategory／valueの`<c:f>`だけを明示更新できるようにした。対象Chart・系列・参照の欠損は保存前に拒否し、Drawing／relationship chainは保持する。Chart作成、cache再計算、Drawingの一般shape編集、Excel再openは未完。
- [x] G2d Chart series-name BUILD: 既存Chart XMLに対する`Vm.set_chart_series_name_formula`（Python binding／型stub含む）を追加し、0-based系列の`<c:tx>`内`<c:f>`だけを更新できるようにした。系列・name referenceの欠損は拒否し、category/value/cache/Drawing relationshipは変更しない。Chart作成、Excel再openは未完。
- [x] G2d Chart marker BUILD: 既存Chart XMLに対する`Vm.set_chart_series_marker_symbol`（Python binding／型stub含む）を追加し、0-based系列の既存`<c:marker>/<c:symbol@val>`だけを限定更新できるようにした。DrawingML標準値をallowlistし、marker／系列欠損は拒否する。Chart作成、cache再計算、Excel再openは未完。
- [x] G2d Chart marker-size BUILD: 既存Chart XMLに対する`Vm.set_chart_series_marker_size`（Python binding／型stub含む）を追加し、0-based系列の既存`<c:marker>/<c:size@val>`だけを2..72範囲で限定更新できるようにした。marker／系列欠損は拒否する。Chart作成、cache再計算、Excel再openは未完。
- [x] G2d Chart smooth BUILD: 既存Chart XMLに対する`Vm.set_chart_series_smooth`（Python binding／型stub含む）を追加し、0-based系列の既存`<c:smooth@val>`だけをboolean OOXML値へ限定更新できるようにした。欠損要素は推測生成せず拒否し、系列の他要素・cache・Drawing relationshipは保持する。Chart作成、cache再計算、Excel再openは未完。[G2d測定](docs/measurements/g2d-object-editing-2026-09-09.md)
- [x] G2d Chart title BUILD: Chart XMLに対する`Vm.set_chart_title`（Python binding／型stub含む）を追加し、既存`<c:title>`では最初の`<a:t>`だけを、title欠損時は`plotArea`前に最小rich-text titleをXML escape付きで追加できるようにした。既存titleのtext欠損、制御文字、16KiB超は拒否し、他のChart XML・Drawing／relationship chainは保持する。複数text runの完全編集、Chart作成、Excel再openは未完。
- [x] G2d Chart legend BUILD: Chart XMLの最初の`<c:legendPos@val>`を`Vm.set_chart_legend_position`（Python binding／型stub含む）から`b`／`tr`／`r`／`l`／`t`の範囲で限定更新し、凡例欠損時は`plotArea`前に最小凡例を追加できるようにした。legend以外のChart XML・Drawing／relationship chainは保持し、Chart作成・Excel再openは未完。
- [x] G2d Chart style BUILD: 既存Chartのchart-space `<c:style val>`を`Vm.set_chart_style`（Python binding／型stub含む）から1..48の範囲で更新し、style欠損時は`<c:chart>`直前へ追加できるようにした。系列・タイトル・凡例・Drawing／relationship chainは保持し、Chart作成・Excel再openは未完。
- [x] G2d Chart axis-title BUILD: cat／val／date／ser軸のdocument-order指定について、`Vm.set_chart_axis_title`（Python binding／型stub含む）から軸内の最初のtitle text runを更新し、欠損時は最小titleを追加できるようにした。Chart作成、Excel再openは未完。
- [x] G2d Chart legend-overlay BUILD: 既存凡例の`c:overlay@val`を`Vm.set_chart_legend_overlay`（Python binding／型stub含む）からboolean OOXML値で更新し、属性欠損時は凡例末尾へ追加できるようにした。凡例欠損、Chart作成、Excel再openは未完。
- [x] G2d Chart data-label BUILD: 既存Chartの最初の`c:dLbls@showVal`を`Vm.set_chart_data_labels_show_value`（Python binding／型stub含む）からboolean OOXML値で追加／更新できるようにした。系列・軸・relationshipは保持し、data-label要素の新規作成、他のラベル属性、Chart作成、Excel再openは未完。
- [x] G2d Chart data-label category BUILD: 既存Chartの最初の`c:dLbls@showCat`を`Vm.set_chart_data_labels_show_category`（Python binding／型stub含む）からboolean OOXML値で追加／更新できるようにした。showValとの併用を保持し、data-label要素の新規作成、他のラベル属性、Chart作成、Excel再openは未完。
- [x] G2d Chart data-label series-name BUILD: 既存Chartの最初の`c:dLbls@showSerName`を`Vm.set_chart_data_labels_show_series_name`（Python binding／型stub含む）からboolean OOXML値で追加／更新できるようにした。showVal／showCatとの併用と既存子要素を保持し、data-label要素の新規作成、Chart作成、Excel再openは未完。
- [x] G2d Chart data-label percent BUILD: 既存Chartの最初の`c:dLbls@showPercent`を`Vm.set_chart_data_labels_show_percent`（Python binding／型stub含む）からboolean OOXML値で追加／更新できるようにした。既存data-label属性・子要素を保持し、data-label要素の新規作成、Chart作成、Excel再openは未完。
- [x] G2d Chart data-label leader-lines BUILD: 既存Chartの最初の`c:dLbls@showLeaderLines`を`Vm.set_chart_data_labels_show_leader_lines`（Python binding／型stub含む）からboolean OOXML値で追加／更新できるようにした。既存data-label属性・子要素を保持し、data-label要素の新規作成、Chart作成、Excel再openは未完。
- [x] G2d Chart data-label bubble-size BUILD: 既存Chartの最初の`c:dLbls@showBubbleSize`を`Vm.set_chart_data_labels_show_bubble_size`（Python binding／型stub含む）からboolean OOXML値で追加／更新できるようにした。既存data-label属性・子要素を保持し、data-label要素の新規作成、Chart作成、Excel再openは未完。
- [x] G2d Chart data-label legend-key BUILD: 既存Chartの最初の`c:dLbls@showLegendKey`を`Vm.set_chart_data_labels_show_legend_key`（Python binding／型stub含む）からboolean OOXML値で追加／更新できるようにした。既存data-label属性・子要素を保持し、data-label要素の新規作成、Chart作成、Excel再openは未完。
- [x] G2d Chart data-label position BUILD: 既存Chartの最初の`c:dLbls/c:dLblPos@val`を`Vm.set_chart_data_labels_position`（Python binding／型stub含む）からOOXML定義値に限定して追加／更新できるようにした。既存data-label属性・子要素を保持し、Chart作成、描画再現、Excel再openは未完。
- [x] G2d Chart data-label number-format BUILD: 既存Chartの最初の`c:dLbls/c:numFmt@formatCode`を`Vm.set_chart_data_labels_number_format`（Python binding／型stub含む）から検証付きで追加／更新できるようにした。`sourceLinked`と他のdata-label内容を保持し、Chart作成、描画再現、Excel再openは未完。
- [x] G2d Chart data-label separator BUILD: 既存Chartの最初の`c:dLbls/c:separator@val`を`Vm.set_chart_data_labels_separator`（Python binding／型stub含む）から検証付きで追加／更新できるようにした。既存data-label属性・子要素を保持し、Chart作成、描画再現、Excel再openは未完。
- [x] G2d Chart data-label number-format BUILD: 既存Chartの最初の`c:dLbls/c:numFmt@formatCode`を`Vm.set_chart_data_labels_number_format`（Python binding／型stub含む）から検証付きで追加／更新できるようにした。`sourceLinked`と他のdata-label内容を保持し、Chart作成、描画再現、Excel再openは未完。
- [x] VBA event trigger BUILD: Python `Vm.set_cell(..., trigger_events=True)`で、直前にparse済みVBAの`Worksheet_Change(Target)`を変更対象A1へ自動dispatchできるようにした。既存のEnableEvents・再入抑止・timeoutを適用し、既定の`trigger_events=False`と明示dispatch APIは維持する。同一Program内の複数handler順序とExcel oracleは未完。
- [x] VBA cell-write consistency BUILD: Python `Vm.set_cell`をVM共通の値書込み経路へ接続し、1-based座標検証、variant budget、undo、spill解放、formula AST／dirty依存無効化を一括適用するようにした。イベントtriggerは明示opt-inのまま維持し、Excel oracleは未完。
- [x] G2d Drawing anchor BUILD: 既存Drawing XMLの指定したtwo-cell anchorについて、1-basedのfrom/toセルを`Vm.set_drawing_anchor`（Python binding／型stub含む）から限定更新できるようにした。anchor index・one-cell anchor・欠損marker・逆順座標は拒否し、shape content／offset／relationshipは保持する。Drawing作成・一般shape編集・row/column変更への自動追従・Excel再openは未完。
- [x] G2d Chart cache BUILD: Chart系列のcategory/value cacheについて、`Vm.set_chart_series_cache`（Python binding／型stub含む）から既存`strCache`／`numCache`の点数・値を限定更新し、cache欠損時は既存`strRef`／`numRef`から対応cacheを生成できるようにした。formula・既存formatCode・周辺XMLを保持し、制御文字・有限数値・サイズ上限・cacheable reference欠損を検証する。cache再計算、Chart作成、Excel再openは未完。
- [x] G2d Drawing shape metadata BUILD: 既存two-cell／one-cell／absolute anchor内の`<xdr:cNvPr name>`／`descr`／`title`を`Vm.set_drawing_shape_name`／`Vm.set_drawing_shape_description`／`Vm.set_drawing_shape_title`（Python binding／型stub含む）からXML escape付きで追加／更新できるようにした。geometry・shape content・relationshipは保持し、Chart/Drawing作成、shape属性全般、Excel再openは未完。
- [x] G2d Drawing rotation BUILD: 既存shapeの`<a:xfrm rot>`を`Vm.set_drawing_shape_rotation`（Python binding／型stub含む）から整数度で限定更新できるようにした。geometry・shape content・relationshipは保持し、transform欠損時は推測生成せず拒否する。Drawing作成、一般shape編集、Excel再openは未完。
- [x] G2d Drawing flip BUILD: 既存shapeの`<a:xfrm flipH/flipV>`を`Vm.set_drawing_shape_flip`（Python binding／型stub含む）から限定更新できるようにした。geometry・shape content・relationshipは保持し、transform欠損時は推測生成せず拒否する。Drawing作成、一般shape編集、Excel再openは未完。
- [x] G2d Drawing fill BUILD: 既存shapeの`<a:solidFill>/<a:srgbClr>`を`Vm.set_drawing_shape_fill`（Python binding／型stub含む）から6桁RGB／8桁ARGBで限定更新できるようにした。既存fillは色だけを更新し、欠損時は既存`spPr`へ最小fillを追加する。theme／gradient fill、Drawing作成、Excel再openは未完。
- [x] G2d Drawing line-color BUILD: 既存shapeの`<a:ln>/<a:solidFill>/<a:srgbClr>`を`Vm.set_drawing_shape_line_color`（Python binding／型stub含む）から6桁RGB／8桁ARGBで限定更新できるようにした。line要素・solid fill・色要素が欠損する場合は推測生成せず拒否する。line作成、dash／width等の一般編集、Excel再openは未完。
- [x] G2d Drawing line-width BUILD: 既存shapeの`<a:ln@w>`を`Vm.set_drawing_shape_line_width`（Python binding／型stub含む）からポイント指定でEMUへ変換して限定更新できるようにした。既存lineの幅だけを変更し、色・dash・shape構造は保持する。line作成、一般style編集、Excel再openは未完。
- [x] G2d Drawing line-dash BUILD: 既存shapeの`<a:ln>/<a:prstDash@val>`を`Vm.set_drawing_shape_line_dash`（Python binding／型stub含む）からDrawingML標準preset値で限定更新できるようにした。既存lineの他要素は保持し、custom dash、line作成、Excel再openは未完。
- [x] G2d Drawing text BUILD: 既存anchor内の最初の`<a:t>` text runを`Vm.set_drawing_shape_text`、任意の既存runを0-based indexで`Vm.set_drawing_shape_text_run`（Python binding／型stub含む）からXML escape付きで限定更新できるようにした。欠損text runは推測生成せず拒否し、他のrun・shape content・geometry・relationshipは保持する。Drawing作成、Excel再openは未完。[G2d測定](docs/measurements/g2d-object-editing-2026-09-09.md)
- [x] G2d Drawing text regression: 最小Drawing-backed XLSXの保存後ZIP再読込で、選択anchorの先頭および0-based指定text runのXML escape付き更新を確認した。indexed-run追加後の低ディスクRust workspace全targetは1,681 library testsと全integration／benchmark targetが成功した。[G2d測定](docs/measurements/g2d-object-editing-2026-09-09.md)
- [x] G2d Drawing hidden BUILD: 既存two-cell／one-cell／absolute anchor内の`<xdr:cNvPr hidden>`を`Vm.set_drawing_shape_hidden`（Python binding／型stub含む）からboolean OOXML値で追加／更新できるようにした。geometry・shape content・relationshipは保持し、Drawing作成、shape属性全般、Excel再openは未完。
- [x] G2d Pivot refresh BUILD: `Vm.set_pivot_cache_refresh_on_load`（Python binding／型stub含む）で既存Pivot cache definitionの`refreshOnLoad`だけを更新できるようにした。外部source取得・cache records変更・Pivot再集計は行わず、属性の追加／置換と既存source編集の併用を回帰検証した。Excel側のrefresh実行と再open確認は未完。
- [x] G2d Pivot field-caption BUILD: 既存Pivot cacheの`cacheField@name`を`Vm.set_pivot_cache_field_caption`（Python binding／型stub含む）から限定更新できるようにした。sharedItems・cache records・field順序・PivotTable layoutは変更せず、属性欠損・範囲外field・制御文字を拒否する。cache values変更、再集計、Excel再openは未完。
- [x] G2d Pivot source BUILD: 既存のworksheet-backed Pivot cacheに対する`Vm.set_pivot_worksheet_source`（Python binding／型stub含む）を追加し、`worksheetSource`のsheet／A1 `ref`を限定更新できるようにした。cache records・PivotTable layout・再集計は変更せず、対象cache／source／属性欠損は拒否する。table-backed source、一般cache編集、Excel再openは未完。
- [x] G2c safety BUILD: sheet rename、row/column insert/delete、sheet move/deleteを構造編集として追跡し、未更新のDrawing／Pivot ownerを保存時に復元しない安全境界を追加した。参照の実更新と明示的な編集APIは未完。
- [x] G2 External Links安全policy: 既定の`preserve`は外部URLを取得・実行せずrelationshipをラウンドトリップ用に保持し、`reject`ではモデル構築前に拒否、`drop`ではowner／relationship／`xl/externalLinks/` partsを保存時に除去する。外部参照数式のExcel oracle校正は未完。
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

- [x] G3 stable identity lookup BUILD: loaded XLSX/XLSMのraw `sheetId`をtab位置・lowercase lookup keyから分離し、Python `Vm.sheet_id()`／`sheet_name_for_id()`でrename／reorder後も解決できるようにした。formula node keyは従来のlowercase sheet keyを維持し、ODS／新規sheetにはIDを推測付与しない。sheet delete後の履歴解決とraw IDをformula node keyにする設計は未完。
- [x] G3 stable identity boundary BUILD: formula nodeはrename時に一括re-keyするlowercase sheet keyで安定して扱い、loaded XLSXのraw `sheetId`は保存／snapshot metadataとして保持する。raw `sheetId`自体をformula node keyに使う設計、sheet delete／位置変更を含む完全なidentity解決は未完。
- [x] G3 local BUILD: workbook／sheet-local nameと単純structured referenceをqualified formula評価へ接続し、sheet rename・row/column shiftの単純参照更新を回帰固定した。stable sheet IDそのものをformula node keyに使う設計は未完。
- [x] 部分 BUILD: シート横断dirty graph、循環検出、manual→automatic、rename、cached valueの扱いを統合し、既存の総work・深さ・参照budgetを維持する。sheet delete／完全な構造変更と全budget組合せは未完。
- [x] G3 local BUILD: cross-sheet dirty closure、循環診断、Manual→Automatic、rename、cached value rollbackを実装・単体検証した。sheet delete／構造変更全体と全budget組合せは未完。
- [x] G3 VBA write consistency BUILD: `Cells(...).Value =` のVBA scalar writeを、Python writeと同じundo／spill解放／formula dirty invalidation経路へ接続し、後続formulaの再計算を回帰固定した。worksheetイベント自動発火・複数module連鎖・bounded連鎖は後続BUILDで追加済みで、Excel oracleは未完。
- [x] G3 VBA range write consistency BUILD: `Range(...).Value` と qualified range writeも対象シート指定可能な共通経路へ接続し、矩形全体を1 undo単位で更新しながらspill解放・formula dirty invalidationを維持した。worksheetイベント自動発火・複数module連鎖・bounded連鎖は後続BUILDで追加済みで、Excel oracleは未完。
- [x] G3 VBA offset write consistency BUILD: `Range(...).Offset(...).Value` も共通書込み経路へ接続し、1-based座標境界を検証して負のwraparoundを拒否する回帰を追加した。worksheetイベント自動発火・複数module連鎖・bounded連鎖は後続BUILDで追加済みで、Excel oracleは未完。
- [x] G3 VBA range clear consistency BUILD: `Range(...).Clear` を範囲共通経路へ接続し、formula dirty invalidation・spill解放・矩形単位undoを適用した。書式・コメント等のClear細分類とworksheetイベント自動発火は後続BUILDで追加済みで、Excel oracleは未完。
- [x] G3 VBA object-range clear BUILD: `Set r = ...Range(...); r.Clear/ClearContents` も束縛Rangeのsheetと矩形を共通クリア経路へ接続し、アクティブシート変更後も誤シートへ書かない回帰を追加した。Clear細分類とworksheetイベント自動発火は後続BUILDで追加済みで、Excel oracleは未完。
- [x] G3 Worksheet_Change auto-dispatch BUILD: `run_sub_with_events` 実行中のVBAセル値書き込みで、同一Programに一意な `Worksheet_Change(Target As Range)` がある場合だけactive sheetの変更対象を自動dispatchする経路を追加した。EnableEvents・再入抑止・明示実行との分離を維持し、同一Program内の複数handler順序とExcel oracleは未完。
- [x] G3 Worksheet_Change range target BUILD: 矩形のRange値書き込み・Clearでも変更後に一度だけ全矩形をTargetへ渡す自動dispatchへ拡張し、`Target.Columns.Count` の回帰を追加した。数式書き込み・複数module連鎖・同一Program内の複数handler順序・Excel oracleは未完。
- [x] G3 Worksheet_Change formula-write BUILD: 単一セルのVBA `Formula` 書き込みも値書き込みと同じWorksheet_Change通知へ接続した。矩形Formula書き込みの一括Target化・複数module連鎖・同一Program内の複数handler順序・Excel oracleは未完。
- [x] G3 Worksheet_Change formula-range BUILD: 矩形のVBA `Formula` 書き込みをセルごとに発火せず、書き込み完了後に一つの全矩形Targetへ通知するようにした。書式細分類・複数module連鎖・同一Program内の複数handler順序・Excel oracleは未完。
- [x] G3 multi-module Worksheet_Change BUILD: `run_sub_multi_with_events`でactive worksheet名とmodule名が一致する`Worksheet_Change`を複数moduleから決定的に選択し、entrypointの値／数式書き込み後にTarget通知する経路を追加した。一致するhandlerがない、または複数ある曖昧な場合は拒否し、同一Program内の複数handler順序とExcel oracleは未完。
- [x] G3 bounded Worksheet_Change chain BUILD: handler内の変更を累積64 dispatch以内のqueueで順次処理し、キュー長が増えない無限連鎖も固定上限で拒否するbounded連鎖を追加した。同一Program内の複数handler順序とExcel oracleは未完。
- [x] G3 Worksheet_Change target metadata BUILD: `Worksheet_Change(Target As Range)`で`Target.Address`／`Target.Row`／`Target.Column`を読み取り可能にし、矩形書き込みの先頭セルと絶対A1アドレスを回帰固定した。Areas全体のExcel意味論・同一Program内の複数handler順序・Excel oracleは未完。
- [x] G3 Worksheet_Change target cell-count BUILD: 単一矩形の`Target.Cells.Count`を`Rows.Count × Columns.Count`として評価し、既存のTarget寸法メタデータとの整合を回帰固定した。複数Areaの総数意味論とExcel oracleは未完。
- [x] G3 Worksheet_Change target parent BUILD: `Target.Parent.Name`で束縛Rangeのworksheet名を参照できるようにし、active sheetの変更対象が別シートへ誤って解決されないことを回帰固定した。イベント所有者のcodeName読込は後続BUILDで追加済みだが、完全なVBAプロジェクト意味論とExcel oracleは未完。
- [x] G3 Clear semantics BUILD: `ClearContents` は追跡済みのセル数値書式・コメントを保持し、`Clear` は対象セルの同メタデータを除去する差分を追加した。OOXML全書式・validation・conditional formatting・Excel oracleは未完。
- [x] G3 Range dimension member BUILD: Rangeオブジェクトの `Rows.Count`／`Columns.Count` を単一矩形で評価できるようにし、イベントTargetの矩形サイズ検証を実行可能にした。複数Areaの総数意味論とExcel oracleは未完。
- [x] G3 Worksheet-to-Range object BUILD: `Set ws = Sheets(...); Set r = ws.Range(...)` と `ws.Cells(row, col)` のRangeオブジェクト生成を追加し、Worksheetから取得したRangeでも束縛シートの編集経路を利用する。複雑なWorksheet memberとExcel oracleは未完。
- [x] 部分 BUILD: Empty／Error伝播、IF/IFERRORの遅延評価、1900系DATE／日付関数、ROUND系の境界をローカル回帰で固定した。VMはロード元の1904 date-system metadataを保持・公開し、Rust／Pythonのロード入口と保存ラウンドトリップで同じ値を維持する。実wheelのPython回帰11件でもこの契約を確認した。serial変換は数式キャッシュと保存の契約が未確定のため未接続。全型変換・丸め規則、独立期待値／Excel oracle校正は未完。
- [x] 部分 MEASURE: 100／1,000 formulaのchainに加え、Data!A1からCalcシート1,000式へのcross-sheet fan-outを構成し、dirty／forced full再走査の代表セル一致、Manual→Automatic、cycleをrelease profile・macOS arm64で100反復測定した。p50/p95・CPU・RSSを記録し、dirtyが常に高速とは主張しない。大規模topology、独立oracle、Linux／Windowsは未完。詳細は[formula dirty calibration](docs/measurements/formula-dirty-calibration-2026-09-05.md)。

### G4 — 関数・配列互換性の拡張（X3）

- [x] G4 dispatcher契約監査: `compat/formula-contracts.json`と`check-formula-dispatch.py --check-contracts`で、mode-rich関数の引数形・対応／未対応mode・回帰テスト参照・oracle未検証状態を機械検査する。Excel oracle fixtureの一致校正は未完。
- [x] G4 VBA WorksheetFunction BUILD: 数式側に既存実装のある`TEXTJOIN`／`XLOOKUP`／`XMATCH`をVBA `WorksheetFunction` dispatchへ接続し、範囲flatten、空値除外、exact／wildcard match、forward/reverse／binary searchの回帰を追加した。完全な型変換とExcel oracle校正は未完。
- [x] G4 dispatcher棚卸し基盤: `scripts/check-formula-dispatch.py`で実際の`eval_func`からcanonical名・alias・重複名を抽出する検査を追加した。FUNCTIONS・引数形・未対応mode・oracle fixtureとの対応検査は未完。
- [x] G4 dispatcher文書対応: `--check-docs`でworksheet関数表と実dispatchの218名（canonical 207 / alias 11）を照合し、未記載・stale記載をエラーにするlocal gateを追加した。引数形・未対応mode・oracle fixtureの対応検査は未完。
- [x] G4 独立formula oracle拡張: LibreOfficeと同一生成fixtureで丸め（FLOOR／CEILING／MROUND）・日付境界（YEAR／MONTH／DAY／EDATE／EOMONTH）・Empty/Error（COUNTBLANK／ISBLANK／ERROR.TYPE／未検出検索）・文字列（SUBSTITUTE／SEARCH／EXACT／PROPER）・配列入力集計（SUMPRODUCT／AVERAGEIF）・統計／型変換（MEDIAN／PRODUCT／RANK／ISNUMBER／ISTEXT／VALUE）を追加校正し、65/65 comparableでelixceeと一致した。DAYS／IFNA／MAXIFS／MINIFS／XMATCH／TEXTJOINは当該LibreOffice buildの未評価として6件を明示skip。これはExcel oracle、全型変換、1904 serial補正、配列境界の完了を意味しない。[測定記録](docs/measurements/formula-independent-oracle-2026-09-10.md)
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
- [x] G4 動的配列サイズ境界: `SEQUENCE` / `RANDARRAY`の行・列サイズを正の整数として検証し、浮動小数の暗黙切り捨て、zero／負数の空配列化、乗算overflowを拒否するようにした。Excel oracle校正は未完。
- [x] G4 stack shape伝播: `VSTACK` / `HSTACK`の引数formula shapeを復元し、同幅・同高の2D結果を既定spill再計算へ接続した。異幅・異高入力は不足領域を`#N/A`でpaddingする。全配列関数のshape伝播は未完。
- [x] G4 FILTER shape伝播: 元の2D range幅と出力要素数から`FILTER`の行列shapeを復元し、2D抽出結果を既定spill再計算へ接続した。include形状の全組合せ、異幅入力、Excel oracle校正は未完。
- [x] G4 TAKE/DROP shape接続: 元配列shapeに基づく行・列単位のTAKE/DROPと任意列数引数を追加し、2D結果を既定spill再計算へ接続した。行／列件数は整数かつ非zeroであることを検証する。全shape伝播、Excel oracle校正は未完。
- [x] G4 axis shape伝播: 1列sourceの`UNIQUE` / `SORT`と、`TOCOL` / `TOROW`の出力軸を既定spill再計算へ接続した。2D sourceの全意味論、全shape伝播、Excel oracle校正は未完。
- [x] G4 INDEX array shape接続: `INDEX(range,0,0)`の全範囲、行配列、列配列について元rangeと引数からshapeを復元し、既定spill再計算へ接続した。行／列番号は整数値として検証し、非整数値を`#VALUE!`として拒否する。2D切出し全体、Excel oracle校正は未完。
- [x] G4 choose axis接続: 2D入力の`CHOOSECOLS` / `CHOOSEROWS`を行列単位で選択し、選択後shapeを既定spill再計算へ接続した。行／列indexは整数値として検証し、非整数値を`#VALUE!`として拒否する。1D legacy経路、Excel oracle校正は未完。
- [x] G4 2D unique/sort接続: 2D入力の`UNIQUE`を行／列単位の重複排除、`SORT`を行列のsort_index・sort_order・by_colに接続し、結果shapeを既定spill再計算へ接続した。`UNIQUE`のexactly_once／by_colと`SORT`のby_colはtruthy型変換に揃え、1D／2D `SORT`のsort_index／sort_orderは不正値を`#VALUE!`として拒否する。型意味論全体とExcel oracle校正は未完。
- [x] G4 2D SORTBY接続: 2Dデータを同じ行のsort-by列で並べ替え、複数sort-by配列の行数・1列制約と各昇降順を検証して結果shapeを既定spill再計算へ接続した。1D／2Dの各sort orderは`1/-1`以外を`#VALUE!`として拒否する。型変換、Excel oracle校正は未完。
- [x] G4 2D FILTER接続: Range既存経路を維持しつつ、生成された2D配列を行／列includeで抽出し、元配列幅と選択列数から結果shapeを既定spill再計算へ接続した。生成配列の比較条件もoperand shapeを継承し、`*`／`+`による複合includeを要素単位のAND／ORとして評価し、異常shapeは明示エラーにする。Excel oracle校正は未完。
- [x] 部分 BUILD: 第1組の参照／条件集計／検索について、IFNA、XLOOKUP／XMATCHのwildcard・binary mode、VLOOKUP／HLOOKUP／INDEX／MATCHの境界とwildcardを追加した。日付／統計／金融、未対応mode全体、Excel oracle校正は未完。
- [x] G4 第1組の意味論補完: `IFNA`を追加し、`#N/A`だけをfallback対象として、それ以外のErrorは伝播する遅延評価を回帰テストで固定した。Excel oracleとの一致確認と、他関数の未対応mode棚卸しは未完。
- [x] G4 検索modeの安全境界: `XLOOKUP`のwildcard `match_mode=2`と、ソート済み数値範囲向けbinary `search_mode=2/-2`を追加した。wildcardとbinaryの組み合わせや未知modeは明示エラーにし、順序を満たさない入力を推測処理しない。Excel oracleとbinary modeの網羅的校正は未完。
- [x] G4 検索関数の整合: `XMATCH`にもwildcard `match_mode=2`と、ソート済み数値範囲向けbinary `search_mode=2/-2`を追加し、未知modeの黙った通常検索を廃止した。Excel oracleと検索関数全体のmode網羅性は未完。
- [x] G4 lookup境界: `HLOOKUP`の行番号0を`#VALUE!`、範囲外を`#REF!`として`VLOOKUP`と同じ明示的な境界に揃え、行／列番号の非整数値も`#VALUE!`として拒否する。Excel oracleによる追加の型変換校正は未完。
- [x] G4 lookup境界: `INDEX`の行列番号0を`#VALUE!`、範囲外を`#REF!`として明示的に検出し、整数アンダーフローによる誤セル参照を防いだ。配列形式・Excel oracle校正は未完。
- [x] G4 MATCH exact互換性: `match_type=0`で文字列wildcard `*` / `?`を大文字小文字非依存で評価し、`-1/0/1`以外と非整数のmatch typeは`#VALUE!`として返すようにした。Excel oracleと近似matchの並び順校正は未完。
- [x] G4 MATCH wildcard補完: `~*` / `~?` / `~~`をリテラルwildcardとして扱うbounded DP経路を追加した。criteria系を含む全wildcard仕様のExcel oracle校正は未完。
- [x] 部分 BUILD: 動的配列のspill衝突・shape・所有権・stale解放・依存更新を主要2D配列関数へ接続し、FILTER等の既存経路も拡張した。LET／LAMBDAを含む全配列関数のshape伝播とExcel oracle校正は未完。
- [x] G4 local BUILD: spill矩形の衝突検出、anchor所有権、stale解放、undo/redo復元、主要2D配列関数のshape接続とatomic rollbackを実装・回帰検証した。全配列関数のshape伝播とExcel oracleは未完。
- [x] 部分 MEASURE: 現行環境のLibreOffice oracleを算術8ケースで再実行したが、8件すべてrunner timeout、比較可能な出力は0件だった。これはExcel oracleや数式一致率の証拠には数えず、Range/Cellsを含む既知のoracle実行境界として記録した。[local oracle記録](docs/measurements/formula-oracle-local-2026-09-10.md)
- [x] 部分 MEASURE: Basic object modelを使わず数式入りXLSXを直接LibreOfficeで再計算する独立経路を追加し、基本集計・エラー値／回復・文字列・検索・criteria集計、数値、文字列正規化と混在型範囲、1904 workbook epochの日付を含む42ケースを比較した。比較可能39ケースは39/39一致し、同一fixtureをelixcee wheelでも再計算して日付serial／Error表示をfixture-level正規化後39/39一致した。LibreOfficeが処理できないIFNA／XMATCH／TEXTJOIN probe 3件はskippedとして記録した。LibreOfficeはExcel oracleではなく、1904シリアル変換・日付／型変換・動的配列全体の互換性は未完。[formula-only oracle記録](docs/measurements/formula-independent-oracle-2026-09-10.md)
- [ ] MEASURE: Excel／EPPlus／Aspose.Cellsはversion・計算設定・license利用条件を固定して比較。実行していない公式対応表と、実測一致率を分ける。

### G5 — 保存メモリの段階削減（X4 / X5）

進行中。G5c/G5dと共通のabort実装、およびmacOSでの部分RSS測定は完了。
G5a/G5bの全passthrough遅延化、3 OS・出力同値の完全測定が残っているため、
G5全体は未完了であり、constant-memoryを主張しない。

- [x] G5a 部分 BUILD: 画像・VBAと、関係解析対象外のXML payloadをraw mapへ保持せず、元ZIPから一件ずつdestinationへ直接copyする経路へ変更した。未編集table XMLと、新規table追加が無い場合のworksheet `.rels` XMLも遅延copyへ移し、編集対象tableや新規table用relsだけを保持してpatchする。worksheet source XMLは保存ループで1シートずつ所有権移動して処理済みpayloadを解放し、workbook XMLもowned fragment抽出後に元全文を解放する。copy前のZIP再検証とentry期待サイズ一致を確認し、reader回帰テストで1 MiBの画像payloadが索引上空のまま保持されることも固定した。workbook rels等の完全遅延化、元file全体の同一性測定、CRC／3 OS検証は未完了のため、G5a全体は未完了。
- [x] G5a 部分 BUILD: defined namesの読込・報告は`xl/workbook.xml`だけを検証付きで直接取得する経路へ分離し、worksheet/table等の兄弟XMLを再読込・保持しない回帰を追加した。これは保存時の全part遅延化や3 OS RSS測定を完了扱いにしない。
- [x] G5a 部分 BUILD: 未編集のtable XMLを保存時のraw mapへ保持せず、table編集がある場合だけ対象partを元ZIPから再取得してpatchする経路へ変更した。全passthrough partの遅延化、保存RSS測定、3 OS検証は未完了。
- [x] G5b 部分 BUILD: worksheet relationship XMLの解析用文字列索引を専用mapから統合relationship mapへまとめ、同一`.rels` payloadの解析用二重保持を削減した。passthrough bytes自体の遅延化とRSS測定は未完了。
- [x] G5b 部分 BUILD: relationship XMLのraw bytesを保存時のpassthrough mapへ保持せず、接続解析後は元ZIPから遅延copyするようにした。新規tableのworksheet `.rels`だけは解析済みXMLをpatchして出力し、重複entryを防止する。全payload遅延化とRSS測定は未完了。
- [x] G5b 部分 BUILD: passthrough relationshipの接続解析とcarry-over判定を借用UTF-8 viewで行い、解析用の一時String cloneを除去した。relationship内容は保存時に必要な範囲だけ所有し、全payload遅延化・RSS測定・速度効果の実測は未完了。
- [x] G5b 部分 BUILD: writer-owned `xl/_rels/workbook.xml.rels`もcarry-over判定後にraw mapから所有権移動し、workbook relationship XMLの一時cloneを除去した。XLSX round-trip 54件とstrict clippyを再実行したが、速度／RSS効果は未測定。
- [x] G5b 部分 BUILD: relationship接続・pruning解析後に`raw_entries`側の`.rels` bytesを解放し、解析用文字列索引とentry名だけを保持するようにした。最終出力は元ZIPから再取得するため、全payload遅延化とRSS測定は未完了。所有権移動の前後release binaryはSHAが一致し、速度／RSSの改善値は得られなかった（[G5測定](docs/measurements/g5-large-hotpath-2026-09-10.md)）。
- [x] G5b 部分 BUILD: shared-string本文を`Vec<String>`へ二重保持せず、所有するindexから参照を座標順に並べて直接出力する経路へ変更した。未編集styles XMLはstyle解決後に解放し、元ZIPから直接copyする。passthrough XMLも出力時に一件ずつdrainして処理済みpayloadを解放する。styles/workbook/[Content_Types]/worksheetのwriter-owned cloneも所有権移動し、cell/row/column style編集時は対象sheetだけをoverlay cloneするようにした。font/fill/borderを使わないstyle編集ではcellXfsだけを展開する。その他の補助索引とdisk spoolは未完了のため、G5b全体は未完了。
- [x] G5c BUILD: 追記専用APIに`create_stream_bounded`を追加し、「1行上限」と「総work量」を分けた。既存`max_pending_bytes`の累積制限は維持し、stub・README・limitsへ移行例を追加した。
- [x] G5d BUILD: bounded追記Writerは固定worksheet構造・列上限・row buffer・64 KiB codec bufferを使い、inline stringsでunique stringsを蓄積しない。通常VMの全セル保持は対象外。
- [x] 共通 BUILD: I/O失敗・context例外時のabort、temp削除、close失敗後の状態、Windowsのclose-before-rename、標準syncの経路を実装した。Windows実機検証は未完。
- [x] 部分 MEASURE: macOS arm64／CPython 3.13の別processで、appendとtransaction付きnormal-freshを100,000・250,000・1,000,000行まで測定し、RSS・wall・temp/output・ZIP／worksheet／最終行検証を記録した。Linux/Windows、Excel oracle、full matrixは未完了。
- [x] 部分 MEASURE: 通常保存と追記を別processで測り、10万→100万行のRSS・p95・temp diskを記録し、ZIP／worksheet形状／最終行に加えて生成入力とのstreaming semantic digest一致を検証するハーネスを追加した。macOSの1M行でも両経路の`semantic_equal: true`を確認済み。Linux／Windows、Excel oracle、通常VMの全セル保持までconstant-memoryと呼ばない。
- [x] 部分 MEASURE: 現行候補の独立Python環境でappend Writerを100,000／250,000／1,000,000行×3列、各2回再測定し、RSS p50/p95 21.62/21.72、21.58/21.66、21.58/21.83 MiBと、ZIP／worksheet／semantic digest全件成功を確認した。これはmacOS append経路の証跡であり、通常VM、Linux／Windows、Excel oracleは未完。[RSS測定](docs/measurements/writer-rss-rerun-2026-09-10.md)
- [x] 部分 MEASURE: 現行候補の独立Python環境でnormal-fresh Writerを100,000／250,000／1,000,000行×3列、各1回測定し、RSS 120.31／207.17／754.95 MiB、ZIP／worksheet／semantic digest全件成功を確認した。通常VMはセルモデルを保持するためRSSが行数依存であり、constant-memoryとは呼ばない。Linux／Windows、Excel oracleは未完。[RSS測定](docs/measurements/writer-rss-normal-fresh-rerun-2026-09-10.md)

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
- [x] 部分 G6 local BUILD/GATE: macOSで失敗時の元出力保護、4 targetの短時間fuzz、`cargo audit --no-fetch --stale`、Python/Rust API回帰を実施した。Linux／Windows、長時間fuzz／soak、3 OS資源校正は未完。
- [x] G5 paired MEASURE: `v1.0.5`タグと現行release binaryを同一fixture・編集・耐久保存・再読込・ZIP全part同値・streaming全セル検証で100k／400k／1Mセル各20ペア比較し、p50総時間で1.320x／1.304x／1.304x（全て1.2x目標達成）、p95でも全ケース短縮を確認した。RSS、他OS、互換性評価は未完。[G5測定](docs/measurements/g5-paired-v1.0.5-2026-09-10.md)

### LogiSheets 対抗トラック（L0–L6）

LogiSheetsの公開metadataは実装・測定・運用実績の証拠ではないため、
「公開API」「実行可能なfixture」「3 runtime」「速度・依存更新・履歴」の軸を分けて比較する。
既存のX/Gゲートを置き換えず、数式とWASMの共有を優先する。

- [x] L0 現状固定: Rust/WASM/Node/ブラウザのread経路は共有済み。数式はsingle-sheet fast pathとworkbook qualified-reference slow path、undo/redoとJS write/recalculateは未完として固定した。
- [x] L1 workbook formula BUILD: sheet-qualified cell/range referenceを明示的なsheet bandへremapし、全sheetのformula nodeを依存順に評価するslow pathを追加。case-insensitive sheet名、mixed host/qualified参照、formula chain、cycleのbest-effortを単体テストで固定した。dirty graphの増分化とstructured referenceは未完。
- [x] L2a named range BUILD: runtime named rangeと、ロード時に取り込んだ単純A1／qualified／`OFFSET` definedNameをformula ASTへ展開し、qualified referenceと混在する`SUM(MyRange)`をworkbook再計算で評価する。table column、連続複数列、静的specifier、同一tableデータ行のthis-row structured referenceを同じ依存グラフへ接続し、行／列構造変更とrenameの単純参照更新にも対応する。table外の行コンテキスト依存・複雑なstructured referenceは未完。
- [x] L2b interval-style dependency BUILD: 大きなrangeを全セルの依存キーへ展開せず、sheet・行・列の区間として保持し、sheet-local formula-node indexとの包含判定で依存辺を構築する。完全な更新用interval treeとincremental workbook dirty propagationは未完。
- [x] L2 dependency graph BUILD（部分完了）: range依存の過剰展開を避けるsheet-local interval-style index、値セルを起点にした直接・range・formula-chainのdirty closure、manual/automatic再計算をworkbook単位で統合した。full interval tree、sheet rename/delete、sheet-scoped/dynamic/structured named range、循環診断、完全な構造変更追跡は未完。
- [x] L3 shared runtime BUILD（部分完了）: 同一Rust coreのworkbook計算を`calculateWorkbook(bytes)`としてWASMへ公開し、`diagnoseWorkbook(bytes)`でsheet/formula/qualified formula/parse errorのJSON summaryを提供、Node/browser同梱runtimeを再生成した。cross-sheet formulaのNode実行、stateful `WorkbookEditor`、transaction付きdata-only operation planのdry-run／apply／undo、`@elixcee/xlsx/runtime`のNode/browser条件、CJS/ESM bundle、意図的に固定したWASM payload baselineに対する10%成長gateを確認した。sync/async loading、browser worker境界、複雑なJS incremental APIは未完。
- [x] L4 editing history BUILD（部分完了）: 明示的なセル/範囲値書き込みとセル数式設定を対象に、最大128操作のbounded undo/redoとtransaction abortをRust VM・Python APIへ追加し、transaction commitは全編集を1つのundo単位として記録することで大規模範囲編集の全体snapshot増殖を抑えた。transaction開始時も既存redo履歴を複製せず、abortでは履歴を保持する。formula cache・dirty状態・tile cacheの復元、prior history保持、nested transaction拒否をテストした。WASMにはstateful `WorkbookEditor`（`setNumber`／`setString`／`setBoolean`、recalculate、undo/redo、transaction）を追加し、座標境界・有限数値を検証した上で、互換rootを汚さない`@elixcee/xlsx/runtime`サブパスからNode/browserへ公開した。sheet操作、OOXML dirty partsと外部効果の一体復元は未完。
- [x] 部分 BUILD: table／validationをtyped data APIへ写像し、data-only操作計画をcapability allowlist・resource budget内でdry-run既定／明示applyに制限した。一般plugin登録・実行、AI差分承認UI、任意コード実行・外部取得は未完または未実装。
- [x] L5 local BUILD: Pythonのtable／validation TypedDict、WASMのread-only validation projection、data-only operation plan、capability allowlist、operation/JSON budget、dry-run既定、明示apply、WASM `WorkbookEditor`へのtransaction一括適用を実装・検証した。
- [x] L5 declarative plugin BUILD: private JS runtimeに、任意callback／module／外部I/Oを許さない名前付きimmutable operation-plan registryを追加し、plugin単位のcapability grant、maxPlugins／operation／JSON budget、dry-run既定、Workbook／WASM editorへのatomic applyと回帰を固定した。任意コードplugin、外部取得、AI承認UIは対象外。
- [ ] L5 general plugin execution: plugin登録・実行の追加は、任意コード実行や外部I/Oを許可しない capability/resource sandbox仕様が固まるまで未実装とする。
- [x] L5 safety boundary 部分 BUILD: private runtimeにdata-onlyの`setNumber`／`setString`／`setBoolean`操作計画を追加し、1-based座標・型付き値・capability allowlist・操作数／JSON bytes budgetを検証する。dry-runを既定とし、`apply: true`の明示時だけ差分を反映する。table／validationのtyped projection、一般plugin実行、外部取得は未完。
- [x] L5 typed projection 部分 BUILD: 既存Python APIの`tables()`／`data_validations()`を`elixcee.pyi`の`TypedDict`へ反映し、table column・範囲・validation ruleを型付き構造metadataとして利用できるようにした。構造情報の参照に限定し、calculated formulaの評価、cell値検証、JS/WASM projection、一般plugin実行は未完。
- [x] L5 WASM projection 部分 BUILD: `readWorkbook`の各worksheetへ`!dataValidations`（validation typeと1-based `sqref`）をread-only構造metadataとして投影し、TypeScript型宣言・WASM回帰・Node/browser payload gateを固定した。table詳細、formula評価、cell値検証、一般plugin実行は未完。
- [ ] L6 MEASURE/GATE: LogiSheetsを固定commit/versionで比較し、formula correctness、dependency update、WASM/Node/browser parity、undo/redo、startup/throughput/RSS/bundle sizeを同じfixtureで測る。WASM payload baselineと10%成長gate、runtime subpathのNode/browser条件、CJS/ESM bundle smokeは先行実装した。LogiSheets固定版比較、同一fixtureの速度/RSS/値parity、Excel oracleは未完。stars/commitsや機能数は補助情報に留め、未測定の優位性は主張しない。

直近の実装成果（2026-09-09）: G3のcross-sheet dirty fan-out測定、G5のstreaming semantic digest検証、L1のworkbook再計算、L2aのnamed range展開、L2bの区間依存辺構築、L2のdirty closure（値セルからの逆引き、formula chain、range）と`set_cell_formula`のqualified-reference初期値処理、L3の`calculateWorkbook(bytes)`／`diagnoseWorkbook(bytes)`共有WASM入口、L4のbounded undo/redo／transaction abort／WASM `WorkbookEditor`、L5のtyped validation projectionとdata-only operation planを実装。
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
| P0 | 現行比1.1倍の速度目標 | 17セル・1,000×10・10,000×10の現行中央値を基準に、各ケースの処理時間を90.9%以下へ削減。標準sync・atomic rename・出力検証を維持 |
| P0 | 大規模全条件で追加1.1倍 | 10万・40万・100万セルの中央値比で判定。style／数式密度・RSS・単一sheetの上限も校正し、1.2倍はストレッチ目標として別管理 |
| P0 | formula dirty propagationの完全校正 | 大規模controlled matrix、full再走査との値一致、p50/p95、CPU/RSS、循環・manual/automatic |

P0速度の新しい目標は、現行同一条件の17セル・1,000×10・10,000×10ケースで処理時間を90.9%以下（速度1.1倍以上）にすることとする。ユーザー提示の暫定値はelixcee 5.02／35.6／317.6 ms、openpyxl 9.54／84.5／947.3 msであり、これは現行実装の比較基準候補である。固定fixture・耐久性・出力検証を再実行するまで、達成済みとは扱わない。従来の1.2倍はストレッチ目標として残す。
2026-09-09の現行1.0.5同一条件再測定では、WriterのDeflate level 1候補により、現行ベースライン比のelixcee中央値は17セル4.9566→4.8646 ms（1.02倍、未達）、1,000×10 16.9883→14.5888 ms（1.16倍、達成）、10,000×10 160.5500→107.7993 ms（1.49倍、達成）だった。全セル値・数式のround-trip検証、F_FULLFSYNC相当の同期、atomic renameは維持した。小規模ケースは同期固定費が支配的であり、目標は未完了。level 2は出力サイズとのバランスは改善したが、大規模中央値153.3230 msでlevel 1を下回ったため採用しない。出力サイズは10,000×10で306,676→575,254 bytesとなるため、速度とサイズのトレードオフを次の評価項目とする。
同期APIを`sync_data`へ弱める実験は17セル5.7711 ms、1,000×10 13.8165 ms、10,000×10 128.4340 msで、level 1の確定計測を改善しなかった。さらにメタデータ同期保証を弱めるため不採用とし、`sync_all`を維持する。
未変更passthrough entryと未変更`styles.xml`を`raw_copy_file`で圧縮済みのまま移送する候補を追加した。変更前実装と同一実行内で交互測定した結果、17セルは4.9381→4.9618 ms（0.995倍、未達）、1,000×10は20.7003→15.9346 ms（1.30倍）、10,000×10は163.8505→126.4021 ms（1.30倍）だった。全セル・数式のround-trip、F_FULLFSYNC相当、atomic renameは成功した。workbook.xml/relsのraw copyは関係ID不整合を起こすため撤回し、Writer生成を維持する。
さらにsource ZIPを解析時とraw copy時に二度open/validateしていた経路を、検証済みarchiveの保持・再利用へ統合した。同一実行内の再測定では、17セル4.8867→4.7842 ms（1.02倍、未達）、1,000×10 17.7606→13.8260 ms（1.28倍）、10,000×10 132.3231→104.5360 ms（1.27倍）だった。検証済みハンドルを保持するためTOCTOU条件を弱めず、全セル・数式round-tripも成功した。小規模はなお同期固定費が目標を阻んでいる。
`[Content_Types].xml`とroot `_rels/.rels`のraw copyは試行したが、source側の追加Defaultやrelationship ID順をそのまま持ち込み、paired ZIP-part比較と一致しないため候補から撤回した。これらはwriter生成を維持し、静的partの差分を安全性ゲートで見逃さない。sharedStrings／stylesと一般passthroughのraw copyだけを候補として残す。
削減後実装の安定測定（before/after各120サンプル、8ラウンド×15反復）では、17セル4.8847→4.6196 ms（1.06倍、未達）、1,000×10 17.6805→14.0657 ms（1.26倍）、10,000×10 133.3955→97.3674 ms（1.37倍）だった。17セルのp50でも同期固定費が支配的であり、単発runの1.1倍達成値は採用判定に使わない。
未変更の`sharedStrings.xml`についても、現在の共有文字列テーブルとsource entryの内容・順序が一致し、構造編集がない場合に限り、生成・再圧縮せず検証済みsource ZIPからraw copyする経路を追加した。before/after各40サンプル（4ラウンド×10反復）の再測定では、17セル4.9971→4.8726 ms（1.03倍、未達）、1,000×10 16.6352→12.8169 ms（1.30倍）、10,000×10 123.4998→91.4622 ms（1.35倍）だった。全ケースでセル値・数式のround-trip検証、F_FULLFSYNC相当、atomic renameを維持した。文字列テーブルが変わる保存では従来の生成経路を使うため、共有文字列の安全性境界は変えていない。
共有文字列の一致判定は、source tableと所有indexを位置ごとに直接照合する方式へ整理し、判定専用の一時`Vec<String>`を作らないようにした。順序・件数・内容の不一致を回帰テストで固定し、保存時の小規模な一時割り当てを削減した。
formula dirty propagationの同日controlled matrixでは、single-input chain 1.2825 ms、warm noop 1.2363 ms、structure rebuild 1.3090 ms、独立1,000入力 1.1570 ms（各median）だった。single-inputがrebuildを上回る短縮は確認できなかったため、closure bookkeepingが支配的になる条件のnegative resultとして記録し、P0最適化候補を維持する。
その後、依存先を座標ではなくformula plan indexでqueueへ渡す局所変更を実装し、single-input chain 1.2538 ms、structure rebuild 1.3223 msを再測定した。直前chain比で約2.2%の改善だが、rebuildは不変であり、独立入力・end-to-end・1.2倍目標の根拠にはしない。range/cycle回帰は成功し、独立入力は未再測定。
残るwarm noopは1.2012 ms、独立1,000入力は1.1502 msで、いずれもCriterion上の有意差なしだった。queue index化はsingle-input chainに限る局所改善として確定し、一般的なdirty propagation高速化とは扱わない。
2026-09-09の候補版大規模確認測定（macOS arm64、同一実行ファイルのbefore/after、20ペア、標準sync・atomic rename・独立openpyxl全セル検証）では、100k cellsがbefore 152.822→after 118.540 ms（p50比1.289倍）、400kが630.931→503.945 ms（1.252倍）、1m・4 sheetsが1,393.716→1,165.116 ms（1.196倍）だった。全ケースでZIP member比較と出力検証は成功した。100k／400kは追加1.2倍目標を満たすが、1mは1.196倍であり丸めて達成扱いにしない。paired median speedupは順に1.284／1.290／1.250倍だが、判定は固定したp50比を優先する。測定JSONは`/private/tmp/elixcee-large-speedup-candidate.json`に保存した（release artifactではない）。

2026-09-07のcandidate健全性確認では、`cargo test --workspace --all-targets --offline`を実行し、Rust unit 1,549件、blackbox、CLI、property、XLSX round-trip 52件、bench smoke、WASM crateを含む全targetが成功した。これはlocal regression evidenceであり、3 OS、Excel oracle、公開artifactの証拠ではない。
同日の`cargo clippy --workspace --all-targets --offline -- -D warnings`も成功し、workspace全targetで警告は検出されなかった。これは静的検査の証拠であり、3 OS・Excel oracle・公開artifactの証拠ではない。
2026-09-09の自己完結ローカルゲートでは、version／measurement boundary／formula dispatch／OOXML matrix／stream measurement self-test、Rust 1,583 tests、strict clippy、Rustdoc、offline audit、fuzz smoke 4種、JS typecheck／operation plan／pack audit／WASM smoke／実tarball CJS・ESM consumer／実Chrome browser smokeを完了した。fuzzはformula parser 225,513、formula eval 162,516、VBA parser 203,945、XLSX reader 17,773 runsで、各終了コード0・RSS上限内だった。これはmacOS上のローカル証跡であり、Linux／Windows、Excel oracle、外部レビュー、registry公開の完了を意味しない。
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
- [x] 部分 G5 BUILD/MEASURE: 現行候補の同一bench binaryでcached appendとreference rescan、dirty closureとstructure rebuildを各10サンプル・2秒条件で再測定した。appendは約1.128倍、dirty closureは約1.323倍だったが、XLSX end-to-end、100k/400k/1Mセル、RSS・耐久保存を含む大規模総合判定は未完。[G5測定](docs/measurements/g5-large-hotpath-2026-09-10.md)
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

| セル数 | before / afterの全体中央値比 | 厳密な1.1倍目標 |
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
- [x] 部分 BUILD: 未実装のVBA `ThisWorkbook.Save`／`ThisWorkbook.Close`（および同じmember判定に入るSave系呼出し）を既定のheadless実行で外部効果として拒否し、`SECURITY`／`E1011`へ構造化分類する回帰を追加した。実際の保存先指定、Close後state、イベント連携、外部リンクの実行、Excel oracleは未完。
- [x] Workbook_Open／Worksheet_Change等のイベントは、EnableEvents・再入抑止・決定的dispatch順・budget付きで設計。複数moduleの`Worksheet_Change`はactive worksheetの表示名またはOOXML `sheetPr@codeName`とmodule名が一致するhandlerだけを選択し、曖昧な場合は拒否する。Excel oracle、codeNameの完全なVBAプロジェクト／Excel実機意味論は未完。
- [x] 部分 BUILD: `Vm.run_event`／Python `Vm.run_event`で明示指定したzero-argumentのWorkbook／Worksheetイベントをdispatchし、`Application.EnableEvents`による無効化、再入抑止、既定execution budgetを適用した。さらに`run_worksheet_change`／Python bindingで、active sheet上の明示A1 targetを`As Range`引数へ一時束縛し、`Value`／`Address`／`Row`／`Column`と単一矩形の行列数を参照できるようにした。自動発火とbounded連鎖は後続BUILDで追加済み、複数handlerのworksheet単位意味論とExcel oracleは未完。
- [x] 部分 BUILD: `Vm.run_sub_multi_with_events`で複数moduleから`Workbook_Open`を一意に解決してentrypointより先にdispatchし、標準module間および同一Program内の重複handlerを拒否する決定性境界を追加した。Worksheetイベントの自動発火とbounded連鎖、active worksheet名に基づく`Worksheet_Change`選択は後続BUILDで追加済み、同一Program内の複数handler順序とExcel oracleは未完。
- [x] 部分 BUILD: VMが実行中に確定したruntime failure categoryをside channelで保持し、CLI JSON診断が文字列再分類なしに利用する経路を追加した。blocked external effectとMsgBox拒否は発生箇所で直接分類し、未移行経路だけがメッセージ再分類へfallbackする。blocked external effectには新しい`E1011`を割り当て、既存の`E1006`（duplicate module name）を壊さない。entrypoint／compile／事前エラーと全エラー生成箇所の完全な型付き移行、独立Excel oracleは未完。
- [x] 部分 BUILD: 同一moduleで同名のUDT（`Type ... End Type`）を検出し、実行前と`check --json`で拒否するようにした。大文字小文字を統一し、module-qualified UDTの単一module解決と診断位置も追加した。複数module間の同名UDTは次項のmodule-local scopeで解決する。
- [x] 複数module間のUDT解決規則を実装し、同名UDTでもSub／Function／Property本体のmodule-localなbare参照は実行中moduleを優先し、qualified参照は明示moduleへ解決するようにした。nested UDTも同じscopeで初期化する。重複UDTの同一module内拒否、手製ASTのflat fallback、外部module間のSub／Function衝突拒否は維持する。
- [ ] Excelとの型変換・丸め・日付・Empty／Error・配列境界・再計算の独立oracle比較を拡張。
- [x] 部分 G4 local regression: 生成VBA corpus 581件を現行CLIで再実行し、572 PASS、8 EXPECTED_RUNTIME_ERROR、1 NONDETERMINISTIC、MISMATCH／UNEXPLAINED 0を確認した。実Excel由来の実運用macro、依存グラフ・循環・volatile／dynamic arrayのExcel意味論は未完。[測定記録](docs/measurements/vba-corpus-local-2026-09-10.md)

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
- [x] reader cancellation BUILD: worksheetのCPU-bound XML event validation中もdeadline／cancelを確認し、SIGINT時は構造XML budget超過より`READER_CANCELED`を優先する回帰を固定した。
- [x] 部分 FUZZ SMOKE: nightlyの4 target（`fuzz_formula_parser`／`fuzz_formula_eval`／`fuzz_vba_parser`／`fuzz_xlsx_reader`）を各5秒・RSS上限1 GiB・既存corpusで実行し、最新の低ディスク再測定で順に213,931／145,288／209,211／21,408 runs、panicなしを確認した。全targetの長時間fuzz、Linux／Windows、soak、隔離検証は未完。[測定記録](docs/measurements/local-gates-2026-09-10.md)
- [x] local gate BUILD: `scripts/check-local-gates.sh`に、version／measurement／formula／OOXML契約、Rust test／clippy／doc、offline audit、4 fuzz smoke、JS type／WASM／pack consumer／browser smokeを統合した。外部oracle、他OS、比較測定、公開操作は意図的に含めない。
- [x] local gate evidence: 2026-09-10のmacOS arm64実行で、低ディスク設定のRust全target 1,679件、4 fuzz smoke、WASM／npm tarball／実Chrome smoke、offline auditを成功させ、標準設定のリンク容量不足と外部未検証境界を[測定記録](docs/measurements/local-gates-2026-09-10.md)へ固定した。Excel oracle、他OS、外部レビュー、公開操作は未完。
- [ ] Linux／Windowsを含む資源校正と、長時間fuzz／CPU／RSS・隔離環境検証を完了する。

既存のmacOS測定と安全策は [測定記録](docs/measurements/README.md)、
[脅威モデル](docs/xlsx-security-model.md) に集約します。限界値の緩和を達成条件にしません。

## 配布・サポートゲート

1.0.3のローカル検証（2026-09-06、macOS）:

- [x] Rust全workspace／全targetの1,679 tests、strict clippy／Rustdoc、依存監査、測定記録検証。
- [x] crate検証、wheel／sdist生成、独立Python環境でのimport・VBA・通常／fast保存往復。
- [x] JS型／differential／WASM bundle／packed consumerと、VBA corpus 581件・意味論386件の既存ゲート。

以下は継続的な配布・サポート要件であり、上記のローカル検証だけでは完了しません。

- [ ] Python wheel／sdist、Rust crate、CLIのclean-installと最小利用例を各対象OSで確認。
- [x] 部分 G6 local BUILD/GATE: JS packageのutils／read／write／browser／Node／型定義を、WASM smoke・実Chrome・real tarball consumer・CJS/ESM parityで検証した。公開可否と他OSのclean-installは未完。
- [x] 部分 G6 local BUILD/GATE: license／notice、real npm package内容、ローカルCargo依存監査、回帰・OOXML compatibility matrixを確認した。fresh advisory DB、他OS配布検証、外部oracleは未完。
- [x] G6 local BUILD: サポート範囲・既知の損失・セキュリティ境界・移行例を [migration guide](docs/migration.md)、v1 support contract、OOXML feature matrixへ集約した。Excel再open、他OS、第三者比較の未検証状態は明記し、公開Release判定とは分離する。
- [ ] tag、registry、workflow、正式Release、worktree状態を別々に確認。

このゲートは将来のリリース判定です。文書の更新だけで完了にしません。
実装・測定は自律的なローカル作業として進め、version変更・push・公開は別の操作として扱います。

## 残課題の依存分類

未チェック項目は、未実装のローカル作業と、外部環境・外部成果物が必要な判定作業を混同しない。

| 分類 | 対象 | 完了条件 |
|---|---|---|
| ローカル実装が残る | G2dのChart/Drawing作成・一般編集、Pivotのcache一般編集（worksheet-backed sourceの限定編集は実装済み）、VBAの実保存／Close後state・イベント・型付きruntime error | API設計、実装、fixture回帰、Rust/Python/WASM境界の検証 |
| ローカル測定が残る | 大規模1mの1.2倍、formula dirty propagation完全校正、実運用macro corpus、長時間fuzz／CPU／RSS | 固定入力・反復・資源上限・失敗条件を記録した再現可能な測定 |
| 外部環境に依存 | Excel再open／修復警告、Excel oracle、Linux／Windows clean-install・資源校正、3 OS検証 | 対象環境の実行結果とversionを取得。macOSローカル結果では代替しない |
| 外部サービス・将来公開に依存 | LogiSheets固定版の取得を伴う競合比較、外部レビュー、registry／GitHub Release／tag公開 | 取得元・固定version・公開状態を別途記録。未実施の推測は完了扱いにしない |

現在の候補版では、自己完結ローカルゲートとmacOS測定を完了した項目だけを `[x]` とし、上表の外部依存項目は未完のまま維持する。
