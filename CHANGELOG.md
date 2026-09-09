# Changelog

重要な変更を [Keep a Changelog](https://keepachangelog.com/en/1.0.0/) 形式で記録します。

## [Unreleased]

次の変更はここに記録します。

- G2dの限定Chart系列編集を拡張し、既存要素のsmooth、負値反転表示、可視性（`delete`）をRust/Python APIと型stubから更新できるようにしました。系列・cache・Drawing relationshipを保持し、Chart作成、cache再計算、Excel再openは未完です。実Excel由来fixtureの一時copyによるload/save/ZIP再読込回帰を追加しました。

- G4の独立formula oracleを72ケースへ拡張しました。LibreOfficeとelixceeの比較可能65ケースは65/65一致し、LibreOffice buildが評価しない6ケースは明示的にskipとして保存しています。結果JSONを`compat/corpus/results/formula-independent-20260910.json`へ固定しました。

- G5bの保存前relationship carry-over判定で、保持対象partの存在確認を線形リスト探索からHashSet参照へ変更しました。出力形式・接続判定は維持し、`cargo check --lib --offline`（debuginfo無効の一時target）で型検証しました。速度／RSS効果は未測定です。

- G2dの限定Drawing編集として、既存shapeの`<a:prstDash>`を`Vm.set_drawing_shape_line_dash`（Python binding／型stub含む）からDrawingML標準値で更新できるようにしました。line作成、custom dash、Excel再openは未完です。

- 複数moduleの`Worksheet_Change`自動dispatchで、表示名に加えて既存worksheet XMLの`sheetPr@codeName`をmodule名解決に利用するようにしました。codeNameはOOXMLから読み取る範囲に限定し、VBAプロジェクトの完全なExcelイベント意味論とExcel実機oracleは未完です。

- G2dの限定Chart編集として、凡例の`c:overlay@val`を`Vm.set_chart_legend_overlay`（Python binding／型stub含む）から追加／更新できるようにしました。凡例の他の子要素・系列・title・relationshipは保持し、凡例欠損、Chart作成、Excel再openは未完です。
- G2dの限定Chart編集として、既存Chartの最初の`c:dLbls@showVal`を`Vm.set_chart_data_labels_show_value`（Python binding／型stub含む）から追加／更新できるようにしました。既存のラベル子要素・系列・軸・relationshipは保持し、data-label要素の新規作成とExcel再openは未完です。
- G2dの限定Chart編集として、既存Chartの最初の`c:dLbls@showCat`を`Vm.set_chart_data_labels_show_category`（Python binding／型stub含む）から追加／更新できるようにしました。showValとの併用、既存ラベル子要素・系列・軸・relationshipを保持し、data-label要素の新規作成とExcel再openは未完です。
- G2dの限定Chart編集として、既存Chartの最初の`c:dLbls@showSerName`を`Vm.set_chart_data_labels_show_series_name`（Python binding／型stub含む）から追加／更新できるようにしました。showVal／showCatとの併用と既存ラベル子要素・系列・軸・relationshipを保持し、data-label要素の新規作成とExcel再openは未完です。
- G2dの限定Chart編集として、既存Chartの最初の`c:dLbls@showPercent`を`Vm.set_chart_data_labels_show_percent`（Python binding／型stub含む）から追加／更新できるようにしました。既存data-label属性・子要素・系列・軸・relationshipを保持し、data-label要素の新規作成とExcel再openは未完です。
- G2dの限定Chart編集として、既存Chartの最初の`c:dLbls@showLeaderLines`を`Vm.set_chart_data_labels_show_leader_lines`（Python binding／型stub含む）から追加／更新できるようにしました。既存data-label属性・子要素・系列・軸・relationshipを保持し、data-label要素の新規作成とExcel再openは未完です。
- G2dの限定Chart編集として、既存Chartの最初の`c:dLbls@showBubbleSize`を`Vm.set_chart_data_labels_show_bubble_size`（Python binding／型stub含む）から追加／更新できるようにしました。既存data-label属性・子要素・系列・軸・relationshipを保持し、data-label要素の新規作成とExcel再openは未完です。
- G2dの限定Chart編集として、既存Chartの最初の`c:dLbls@showLegendKey`を`Vm.set_chart_data_labels_show_legend_key`（Python binding／型stub含む）から追加／更新できるようにしました。既存data-label属性・子要素・系列・軸・relationshipを保持し、data-label要素の新規作成とExcel再openは未完です。
- G2dの限定Chart編集として、既存Chartの最初の`c:dLbls/c:dLblPos@val`を`Vm.set_chart_data_labels_position`（Python binding／型stub含む）からOOXML定義値に限定して追加／更新できるようにしました。既存data-label属性・子要素・系列・軸・relationshipを保持し、Chart作成とExcel再openは未完です。
- G2dの限定Chart編集として、既存Chartの最初の`c:dLbls/c:numFmt@formatCode`を`Vm.set_chart_data_labels_number_format`（Python binding／型stub含む）から検証付きで追加／更新できるようにしました。`sourceLinked`と他のdata-label内容を保持し、Chart作成とExcel再openは未完です。
- G2dの限定Chart編集として、既存Chartの最初の`c:dLbls/c:separator@val`を`Vm.set_chart_data_labels_separator`（Python binding／型stub含む）から検証付きで追加／更新できるようにしました。既存data-label属性・子要素・系列・軸・relationshipを保持し、Chart作成とExcel再openは未完です。
- VBAの`Cells(...).Value`書き込みを統一書き込み経路へ接続し、formula dirty invalidation・spill解放・undo履歴との整合を改善しました。
- VBAの`Range(...).Value`およびqualified range書き込みも共通経路へ接続し、矩形編集を1 undo単位で扱いながら依存無効化とspill解放を適用しました。
- VBAの`Range(...).Offset(...).Value`も共通書き込み経路へ接続し、1-based座標外へのwraparoundを拒否するようにしました。
- VBAの`Range(...).Clear`を範囲共通経路へ接続し、formula依存無効化・spill解放・矩形単位undoを適用しました。
- `Set` で取得した Range オブジェクトの `.Clear` / `.ClearContents` も、保護・undo・spill解放・数式dirty無効化を共有するクリア経路へ接続しました。束縛先とアクティブシートが異なる場合も、束縛先のシートを更新します。
- Worksheetオブジェクトの `.Range(...)` / `.Cells(row, col)` をRangeオブジェクトとして取得できるようにし、取得後の値書き込み・クリアを束縛シートへ接続しました。
- `run_sub_with_events` のVBA実行中に、同一Programの `Worksheet_Change(Target As Range)` をセル値書き込み後に自動dispatchできるようにしました。既存のEnableEvents・再入抑止・通常の `run_sub` のイベントなし契約は維持します。
- 矩形Rangeの値書き込み・クリアでは、セルごとの発火ではなく変更された全矩形を `Target` として一度だけ通知するようにしました。
- 単一セルのVBA数式書き込みでも `Worksheet_Change` を通知するようにしました。
- 矩形のVBA数式書き込みも、全矩形を一つの `Target` として書き込み後に一度だけ通知するようにしました。
- 複数module実行の `run_sub_multi_with_events` でも、一意な `Worksheet_Change` handlerを値／数式書き込み後に自動通知できるようにしました。重複handlerは拒否します。
- `Worksheet_Change` handler内の追加変更をbounded queueで順次処理し、累積64件を超える無限イベント連鎖はエラーにするようにしました。
- `Worksheet_Change(Target As Range)`から `Target.Address`、`Target.Row`、`Target.Column`を参照できるようにしました。
- `Worksheet_Change(Target As Range)`から単一矩形の `Target.Cells.Count` を参照できるようにしました。
- `Worksheet_Change(Target As Range)`から `Target.Parent.Name` で束縛先worksheet名を参照できるようにしました。
- G2dの限定Drawing編集として、既存shapeの回転を `Vm.set_drawing_shape_rotation`（Python binding／型stub含む）から整数度で更新できるようにしました。既存transformがないshapeは安全のため拒否します。
- G2dの限定Drawing編集として、既存shapeの水平／垂直反転を `Vm.set_drawing_shape_flip`（Python binding／型stub含む）から更新できるようにしました。既存transformがないshapeは安全のため拒否します。
- G2dの限定Drawing編集として、既存shapeのsolid fill色を `Vm.set_drawing_shape_fill`（Python binding／型stub含む）からRGB／ARGB hexで更新できるようにしました。theme／gradient fillは対象外です。
- G2dの限定Drawing編集として、既存shapeのline色を `Vm.set_drawing_shape_line_color`（Python binding／型stub含む）からRGB／ARGB hexで更新できるようにしました。line構造が欠ける場合は安全のため拒否します。
- G2dの限定Drawing編集として、既存shapeのline幅を `Vm.set_drawing_shape_line_width`（Python binding／型stub含む）からポイント指定で更新できるようにしました。
- VBAイベントの複数module対応として、active worksheet名に一致する`Worksheet_Change` moduleを決定的に選択し、それ以外の曖昧な重複handlerは従来どおり拒否するようにしました。
- 2026-09-10のmacOS arm64で自己完結ローカルゲートを再実行し、Rust全target 1,663件、Clippy、Rustdoc、offline audit、4 fuzz smoke、WASM、npm tarball、実Chrome smokeを成功させました。外部Excel oracle・他OS・外部レビュー・公開操作は未完です。
- `ClearContents` と `Clear` を区別し、追跡済みのセル数値書式・コメントを前者では保持、後者では対象セルから除去するようにしました。
- Rangeオブジェクトの単一矩形について `Rows.Count` と `Columns.Count` を評価できるようにしました。
- G2dの限定Chart編集として、既存Chartの最初の`c:dLbls/c:numFmt@formatCode`を`Vm.set_chart_data_labels_number_format`（Python binding／型stub含む）から検証付きで追加／更新できるようにしました。`sourceLinked`と他のdata-label内容を保持し、Chart作成とExcel再openは未完です。
- VBAイベントの部分BUILDとして、Python `Vm.set_cell(..., trigger_events=True)`から直前にparse済みVBAの`Worksheet_Change(Target)`を変更対象A1へ自動dispatchできるようにしました。既定動作は従来どおりイベントなしで、EnableEvents・再入抑止・timeoutを適用します。イベント連鎖・複数handler順序・Excel oracleは未完です。
- VBAセル書込みの整合性を補強し、Python `Vm.set_cell`をVM共通の値書込み経路へ接続しました。1-based座標、variant budget、undo、spill解放、formula AST／dirty依存無効化を適用します。Excel oracleは未完です。
- G2dの限定Chart編集として、軸タイトルを`Vm.set_chart_axis_title`（Python binding／型stub含む）からdocument orderの0-based indexで更新／追加できるようにしました。軸内の最初のtitle text runだけを変更し、系列・Chart title・relationshipは保持します。Chart作成、Excel再openは未完です。
- G2dの限定Drawing編集として、既存anchorの`cNvPr@hidden`を`Vm.set_drawing_shape_hidden`（Python binding／型stub含む）から追加／更新できるようにしました。geometry・shape content・relationshipは保持し、Drawing作成とExcel再openは未完です。
- G4/VBA互換性の部分BUILDとして、`WorksheetFunction.TextJoin`、`WorksheetFunction.XLookup`、`WorksheetFunction.XMatch`をVMへ接続しました。TEXTJOINは範囲flattenと空文字除外、XLOOKUP／XMATCHはexact／wildcard matchおよびsearch mode 1/-1/2/-2を実装し、旧compat corpusの「未実装」期待値を削除しました。完全な型変換とExcel oracle照合は未完です。
- G2dの限定Chart編集として、`Vm.set_chart_style`（Python binding／型stub含む）で既存Chartの`c:style@val`を1..48の範囲で更新し、欠損時はchart-spaceへ追加できるようにしました。系列・タイトル・凡例・Drawing relationshipは保持し、Chart作成とExcel再openは未完です。
- イベントhandlerの決定性を強化し、同一Program内の重複`Workbook_Open`／`Worksheet_Change`を先頭選択せず拒否するようにしました。
- `Worksheet_Change(Target As Range)`の部分BUILDとして、`run_worksheet_change`（Rust／Python）に明示A1 targetを渡せるようにしました。Target bindingは一時的なRange objectとして行い、不正範囲・型不一致・再入を拒否または抑止します。セル編集からの自動発火とイベント連鎖は未完です。
- 複数module実行でも`Workbook_Open`をopt-in先行dispatchできる`Vm.run_sub_multi_with_events`を追加しました。標準module間の重複handlerはsource traversal順に依存せず拒否します。
- G2dの限定Chart編集として、既存系列のname formulaを`Vm.set_chart_series_name_formula`（Python binding含む）から更新できるようにしました。系列index・`<c:tx>`・`<c:f>`の欠損、制御文字、16KiB超を拒否し、他の系列・cache・Drawing relationshipは保持します。Excel再openは未完です。
- G2dの限定Chart編集として、`Vm.set_chart_title`（Python binding含む）で既存Chartの最初のtitle text runを更新し、title欠損時は最小rich-text titleを追加できるようにしました。XML escape、既存titleのtext欠損、制御文字、16KiB上限を検証し、周辺のChart XML・Drawing relationshipは保持します。Chart作成、複数runの完全編集、Excel再openは未完です。
- G2dの限定Chart編集として、`Vm.set_chart_legend_position`（Python binding含む）から既存Chartの凡例位置を`b`／`tr`／`r`／`l`／`t`の範囲で更新し、凡例欠損時は最小凡例を追加できるようにしました。凡例以外のChart XML・Drawing／relationship chainは保持し、Chart作成・Excel再openは未完です。
- G2dの限定Drawing編集として、既存two-cell anchorのfrom/toセルを`Vm.set_drawing_anchor`（Python binding含む）から1-based座標で更新できるようにしました。shape content、offset、relationshipを保持し、one-cell anchorや不正座標は拒否します。Drawing作成・一般shape編集・Excel再openは未完です。
- G2dの限定Chart編集として、既存系列のcategory/value cacheを`Vm.set_chart_series_cache`（Python binding含む）から更新できるようにしました。既存`strCache`／`numCache`の点数・値だけを置換し、formula・formatCode・cache種別・周辺XMLを保持します。cache作成／再計算、Chart作成、Excel再openは未完です。
- Chart系列にcacheが存在しない場合も、既存の`strRef`／`numRef`を変えずに`strCache`／`numCache`を生成してから点数・値を保存できるようにしました。
- G2dの限定Drawing編集として、既存drawing anchor内の`cNvPr name`を`Vm.set_drawing_shape_name`（Python binding含む）から更新できるようにしました。two-cell／one-cell／absolute anchorに対応し、geometry・shape content・relationshipを保持します。Drawing作成・shape属性全般編集・Excel再openは未完です。
- G2dの限定Drawing編集として、既存drawing anchor内のoptionalな`cNvPr descr`を`Vm.set_drawing_shape_description`（Python binding含む）から追加／更新できるようにしました。XML escapeと既存shape構造を保持し、Drawing作成・shape属性全般編集・Excel再openは未完です。
- G2dの限定Drawing編集として、既存drawing anchor内のoptionalな`cNvPr title`を`Vm.set_drawing_shape_title`（Python binding含む）から追加／更新できるようにしました。geometry・shape content・relationshipは保持し、Drawing作成・shape属性全般編集・Excel再openは未完です。
- 複数moduleで同名のUDTを定義した場合でも、Sub／Function／Propertyの所属moduleを実行時scopeとしてbare型名を解決し、`Module.Type`形式のqualified参照とnested UDTのmodule-local解決を維持するようにしました。同一module内の重複Type、Sub／Functionのflat namespace衝突は引き続き拒否します。
- G2dの限定Pivot編集として、既存cache definitionの`refreshOnLoad`を`Vm.set_pivot_cache_refresh_on_load`（Python binding含む）から追加／更新できるようにしました。外部取得・cache records変更・Pivot再集計は行わず、Excel側のrefreshと再openは未検証です。
- G2dの限定Pivot編集として、既存cache fieldの`name`を`Vm.set_pivot_cache_field_caption`（Python binding含む）から更新できるようにしました。sharedItems・cache records・PivotTable layoutは保持し、cache values変更・再集計・Excel再openは未完です。
- イベント処理を拡張し、`run_with_events`（Rust／Python）で`Workbook_Open`を明示的に先行dispatchできるようにしました。通常の`run`／`run_sub`は従来どおりイベントを自動実行せず、open handlerの失敗時は本体Macroを実行しません。
- VBAイベントの部分BUILDとして、明示指定したzero-argumentの`Workbook_Open`／`Workbook_BeforeClose`／`Worksheet_Change`／`Worksheet_Calculate`／`Worksheet_SelectionChange`を`Vm.run_event`（Python binding含む）から実行できるようにしました。`Application.EnableEvents`の無効化と再入抑止、既存execution budgetを適用しています。自動発火、`Worksheet_Change(Target)` binding、複数handler順序、イベント連鎖は未完です。
- VBAの安全境界を補強し、`ThisWorkbook.Save`／`ThisWorkbook.Close`を既定のheadless実行で暗黙の成功扱いにせず、外部効果として拒否するようにしました。`E1011`の構造化runtime failure分類と回帰テストで検証しています。実際の保存・Close後state・イベント連携・Excel oracleは未完です。
- G2dのローカル実装として、既存Chartの系列ごとのcategory／value formulaを`Vm.set_chart_series_formulas`から更新できるようにしました。0-based系列指定、欠損Chart／系列／参照の拒否、Drawing／relationship chain保持を実fixtureで検証しています。Chart作成、cache再計算、Drawingの一般shape編集、Excel再openは未完です。
- G2dのローカル実装として、既存のworksheet-backed Pivot cacheの`worksheetSource`を`Vm.set_pivot_worksheet_source`から限定更新できるようにしました。sheetまたはA1 `ref`を変更できますが、cache records／PivotTable layout／再集計は変更しません。合成fixtureでChart/Pivotの参照 chain保持と回帰を検証し、table sourceとExcel再openは未完です。
- G2dのChart/Pivot限定編集の再現手順と検証結果を`docs/measurements/g2d-object-editing-2026-09-09.md`へ記録しました。ローカルBUILD証拠と、Excel再open・再集計・他OS・外部比較の未検証境界を分離しています。
- VBA runtime error診断の部分BUILDとして、VMが実行中の失敗カテゴリを構造化side channelで保持し、CLI JSONがそれを優先利用するようにしました。blocked external effectには`E1011`を割り当て、既存の`E1006`（duplicate module name）を維持しています。事前／compile errorと全エラー生成箇所の完全移行は未完です。
- runtime failure分類を発生箇所へ寄せ、blocked external effectと`error_on_msgbox`によるMsgBox拒否をVMが直接`RuntimeFailureKind`へ記録するようにしました。既存利用者向けの文字列エラーは維持し、未移行経路だけをfallback分類します。全エラー生成箇所の型付き移行とExcel oracleは未完です。
- 同一／複数moduleの同名UDTを大文字小文字非依存で検出し、実行前と`check --json`で`E1012`として明示拒否するようにしました。重複宣言・後勝ちの型定義上書きを防ぎますが、module-qualified UDT解決と診断位置は未完です。
- module-qualifiedなUDT型名（例: `Types.Point`）をパーサーで保持し、module名付きの型定義へ解決できるようにしました。同名UDTの複数module混在は従来どおり曖昧性として拒否します。
- 次期高速化目標を、同一条件の17セル・1,000×10・10,000×10で現行比1.1倍（処理時間90.9%以下）に設定しました。openpyxl比の提示値は再現計測前の暫定基準として扱い、耐久性・出力検証を含む再測定後に達成判定します。
- WriterのDeflate level 1候補を同一条件で測定し、1,000×10と10,000×10では現行比1.1倍を達成しました。17セルは1.02倍に留まり、同期固定費を含む全ケース目標は未完了です。出力サイズ増加も確認したため、公開採用前の評価項目として残します。
- 未変更のpassthrough entryとstyles.xmlを圧縮済みのまま移送するWriter最適化を追加しました。大規模ケースでは変更前との同一実行内比較で約1.30倍を確認しましたが、17セルは未達です。workbook.xml/relsのraw copyは関係ID不整合のため採用していません。
- source ZIPの解析・raw copyで同じ検証済みarchiveを再利用する経路へ整理しました。追加の同一条件測定で1,000×10は1.28倍、10,000×10は1.27倍でした。17セルは1.02倍で、同期固定費を含む目標は継続中です。
- `[Content_Types].xml` と root `_rels/.rels` のraw copy候補は、source固有のDefaultやrelationship ID順がpaired ZIP-part契約と一致しないため撤回しました。writer生成を維持し、sharedStrings／stylesと一般passthroughの安全なraw copyだけを残しています。
- 120サンプルの安定測定でも17セルは1.06倍に留まり、単発runの揺れを除外して未達と判定しました。大規模ケースは1.26〜1.37倍を維持しています。
- 未変更で順序も一致する`sharedStrings.xml`を検証済みsource ZIPからraw copyする候補を追加しました。40サンプルの同一条件測定では17セル1.03倍、1,000×10 1.30倍、10,000×10 1.35倍でした。文字列テーブルが変わる場合は従来どおり生成し、全ケースでround-trip・同期・atomic renameを検証しています。17セルを含む1.1倍目標は未完了です。
- `sharedStrings.xml` の一致判定をindexへの直接照合へ変更し、判定用の一時文字列テーブルを生成しないようにしました。順序・件数・内容の不一致を回帰テストで検証しています。
- 自己完結ローカルゲートを再実行し、Rust 1,583 tests、strict clippy／Rustdoc／offline audit、4種のfuzz smoke、JS typecheck、実tarball consumer、WASM、実Chrome browser smokeを成功させました。これはmacOSローカル証跡であり、Linux／Windows、Excel oracle、外部レビュー、公開を完了扱いにしません。
- 大規模20ペア測定を追加し、100k／400k cellsではp50 1.289／1.252倍、1m・4 sheetsでは1.196倍でした。全ケースでZIP member比較と独立openpyxl全セル検証に成功し、1mは1.2倍未達として記録しています。
- ROADMAPに残課題の依存分類を追加し、ローカル実装・ローカル測定・Excel/他OS環境・外部サービス／将来公開を分離しました。未実施の外部成果をローカル証跡で代替しない方針を明記しています。
- crates.ioの公開検証で検出したworkspace依存crateのsource driftを解消するため、`elixcee-types`の公開patch版をroot crateの検証前に明示pinするようにしました。

## [1.0.5] - 2026-09-09

- G3として、loaded XLSX/XLSMのraw `sheetId`をlowercase lookup key・タブ位置から分離し、rename／reorder後もPython `Vm.sheet_id()`／`Vm.sheet_name_for_id()`で解決できるようにしました。新規／ODS sheetへの推測ID付与、削除後履歴、raw IDをformula node keyにする設計は未完です。
- L5の宣言型plugin schemaをallowlist化し、`module`／`code`など未定義の実行主体フィールドも登録時に拒否するようにしました。
- LogiSheets対抗L5として、private JavaScript runtimeに宣言型plugin registryを追加しました。pluginはimmutableなdata-only operation planに限定し、登録callback／module、任意コード、外部I/Oを拒否します。pluginごとのcapability grantとmaxPlugins／operation／JSON budget、dry-run既定、Workbook／WASM editorへのatomic applyを回帰検証しています。
- G6の移行・境界ガイドを追加し、native Python／CLI／private JavaScriptの使い分け、外部リンクpolicy、OOXML／VBA／constant-memoryの既知の境界、ローカル検証とExcel／他OS／競合比較の証拠分離を短く整理しました。
- LogiSheets対抗L5の安全境界として、private runtimeにdata-onlyの操作計画APIを追加しました。有限数値・文字列・真偽値のcell writeを対応するcapability・件数・JSON bytes budget付きで検証し、dry-runを既定、`apply: true`を明示した場合だけ適用します。任意コード実行、外部I/O、table／validationのtyped projectionは未完です。
- L4 shared runtimeのWASM `WorkbookEditor`に`setString`／`setBoolean`を追加し、`setNumber`を含むtyped cell writesの座標境界・有限数値検証をNode/browser配布アーティファクトへ反映しました。
- WASM buildで`wasm-opt -Oz`を明示し、typed editor追加後のpayloadを961,710 bytes（baseline比+9.72%）へ抑え、既存の10%成長ゲートを基準値変更なしで通過させました。
- Pythonの既存`tables()`／`data_validations()` metadata APIを`elixcee.pyi`のTypedDictへ追加し、table/validationの構造projectionを静的利用側にも公開しました。数式評価やcell値検証は行いません。
- G2dとして、単純なsheet renameに限りChart XMLの`<c:f>`参照を既存formula parserで検証しながら書き換え、Drawing owner／relationshipを保持して保存する経路を実fixtureで検証しました。chart creation、一般のChart/Drawing編集、Pivot更新、Excel再open検証は未完です。
- G2dとして、単純なsheet renameに限りPivot cache definitionの`worksheetSource@sheet`を更新する経路を追加しました。cache本体・cacheId・table source・Pivot再集計は変更せず、一般のPivot source/cache編集とExcel再open検証は未完です。
- G2のExternal Links安全policyを追加しました。既定の`preserve`は外部URLを取得・実行せず、`reject`では読込前に拒否し、`drop`では保存時にowner／relationship／partsを除去します。外部参照数式のExcel oracle校正は未完です。
- G4のdispatcher契約監査を追加し、mode-rich関数の引数形、対応／未対応mode、回帰テスト参照、Excel oracle未検証状態を`compat/formula-contracts.json`で宣言してlocal gateで検査できるようにしました。これはExcel一致率を示すものではありません。
- G2c safety BUILDとして、Chart/DrawingまたはPivotを含むworkbookに対するsheet rename／行列構造変更を、参照更新できない場合は保存前に明示拒否するようにしました。古いchart anchorやpivot sourceを成功扱いで保存しません。実際の参照更新APIは未完です。
- G0のOOXML互換性matrixを追加し、Charts／Pivot／Drawings／External Linksについてread・preserve・edit・recalculate・Excel再openを独立した状態で記録・検査できるようにしました。`preserved`と`unverified`を互換性成功として混同しない境界を固定しています。
- README、v1 support contract、docs索引、内部方針の製品説明を、VBA専用ツールではなく、ワークブック編集・対応数式の再計算・データ処理VBAの実行／診断を同じモデルで扱うヘッドレスExcelワークブック自動化ランタイムとして統一しました。対応範囲とExcel完全互換でない境界は維持しています。
- G4の2D `SORT`引数境界を補完し、`sort_index`の非整数値と`sort_order`の`1/-1`以外を`#VALUE!`として拒否するようにしました。`by_col`は既存のtruthy変換に揃え、暗黙の整数丸めやBoolean限定による誤動作を防ぎます。Excel oracle照合は未完です。
- G4の2D `SORTBY`でも各sort orderを`1/-1`に限定し、複数キーの途中にある不正な順序指定を`#VALUE!`として明示的に拒否するようにしました。Excel oracle照合は未完です。
- G4の`TAKE`／`DROP`で行・列件数の非整数値、非有限値、zeroを黙って処理せず明示エラーにしました。正負方向と利用可能範囲を超えた件数の既存挙動は維持しています。Excel oracle照合は未完です。
- G4の`CHOOSECOLS`／`CHOOSEROWS`で行・列indexの非整数値と非有限値を`#VALUE!`として拒否し、浮動小数の暗黙切り捨てを防ぎました。Excel oracle照合は未完です。
- G4の2D `FILTER`で`SEQUENCE`など生成配列の比較条件を要素単位に評価し、比較元の行列shapeをincludeへ伝播するようにしました。行／列条件の両方を既定spill経路へ接続しています。複合includeとExcel oracle照合は未完です。
- G4の`FILTER`で`(条件1)*(条件2)`をAND、`(条件1)+(条件2)`をORとして配列要素単位に合成できるようにしました。条件長のbroadcastと不一致拒否を含む回帰テストを追加しています。Excel oracle照合は未完です。
- G4の`SEQUENCE`／`RANDARRAY`で行・列サイズを正の整数として検証し、浮動小数の暗黙切り捨てとサイズ計算のoverflowを防止しました。Excel oracle照合は未完です。
- G4の1D `SORT`／`SORTBY`でもsort orderを`1/-1`に限定し、2D経路と異なる暗黙丸めを廃止しました。Excel oracle照合は未完です。
- G4の`INDEX`で行・列番号の非整数値と非有限値を`#VALUE!`として拒否し、配列形式の0指定・範囲外エラーとの境界を統一しました。Excel oracle照合は未完です。
- G4の`MATCH`で`match_type`を`-1/0/1`の整数に限定し、`0.5`などの暗黙切り捨てを`#VALUE!`として拒否するようにしました。Excel oracle照合は未完です。
- G4の`VLOOKUP`／`HLOOKUP`で戻り列・行番号の非整数値と負数を`#VALUE!`として拒否し、範囲外は従来どおり`#REF!`とする境界を維持しました。Excel oracle照合は未完です。
- G4の動的配列基盤として、既存のflat `Variant::Array`を変更せずに`ArrayShape`と`Variant::array_shape()`を追加しました。非空配列は1行、空配列は占有領域なしとして扱います。二次元shape生成、worksheetへのspill配置、衝突処理は未実装です。
- G4のspill準備として、1-based worksheet座標の`SpillRect`を追加しました。座標境界、セル位置、空配列、矩形衝突を共通判定できますが、実際のworksheet書込みと`#SPILL!`生成は未実装です。
- G4のspill計画APIとして、VMの`plan_spill_for_value()`で配列結果の占有範囲と既存の非Emptyセル衝突を事前検査できるようにしました。worksheetへの実際のspill配置と依存更新は未実装です。
- G4のspill適用APIとして、`apply_spill_for_value()`を追加しました。衝突検査後にanchorとspillセルを一括反映し、anchorのformulaを保持します。自動再計算への接続と旧spill領域の追跡は未実装です。
- G4のspill所有権管理として、anchorごとの矩形をVMとedit historyで追跡するようにしました。同じanchorの再適用では範囲外の旧spillセルを解放し、undo/redoで所有情報も復元します。自動再計算接続と複数行shapeは未実装です。
- G4のspill stale対策として、旧spillセルへの手動編集やanchor formulaの差し替え時に、該当anchorの旧spill所有権と派生セルを解放するようにしました。
- G4のspill依存順序として、spill先セルを参照するformulaをanchor formulaへの依存としてtopological sortへ反映しました。自動spill再計算と別sheet依存は未実装です。
- G4のspill再計算APIとして、`recalculate_all_with_spills()`を拡張しました。各sheetの配列anchorを処理し、呼出し元のactive sheetを復元してから依存formulaを再計算します。既存`recalculate_all()`の既定動作は維持しています。
- G4の2D spill準備として、`apply_spill_matrix()`を追加しました。矩形matrixの行幅を検証し、`ArrayShape(rows, cols)`を保持したままanchor formulaを維持して適用できます。
- G4の2D spill接続として、`SEQUENCE` / `RANDARRAY` / `TRANSPOSE` / `WRAPCOLS` / `WRAPROWS`のformula shapeを既定の`recalculate_all_with_spills()`で復元し、2D矩形のspillと依存再計算へ接続しました。全配列関数のshape伝播とExcel oracle照合は未完です。
- G4のshape伝播を拡張し、`VSTACK` / `HSTACK`の2D結果を既定spill再計算で正しい矩形へ配置できるようにしました。異幅・異高入力は不足領域を`#N/A`でpaddingします。Excel oracle照合は未完です。
- G4のshape伝播を拡張し、2D rangeを元にした`FILTER`の出力幅を保持して既定spill再計算へ接続しました。include形状の全組合せ、異幅入力とExcel oracle照合は未完です。
- G4のshape伝播を拡張し、元配列shapeに基づく行・列単位の`TAKE` / `DROP`と任意列数引数を既定spill再計算へ接続しました。異常引数とExcel oracle照合は未完です。
- G4のshape伝播を拡張し、1列sourceの`UNIQUE` / `SORT`と`TOCOL` / `TOROW`の出力軸を既定spill再計算へ接続しました。2D source全体の意味論とExcel oracle照合は未完です。
- G4のshape伝播を拡張し、`INDEX(range,0,0)`の全範囲・行配列・列配列を元rangeのshapeで既定spill再計算へ接続しました。2D切出し全体とExcel oracle照合は未完です。
- G4の2D配列対応を拡張し、`CHOOSECOLS` / `CHOOSEROWS`を行列単位で選択して、選択後のshapeを既定spill再計算へ接続しました。異常引数とExcel oracle照合は未完です。
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
- 現行1.0.4 candidateでworkspace全target strict clippy（`-D warnings`）を実行し、警告なしで成功しました。これは静的検査の証拠であり、3 OSやExcel oracleの証拠ではありません。
- CI・publish・release・crates-publish workflowを`actionlint`で検査し、エラーなしを確認しました。GitHub Actionsの実行結果や3 OS測定完了とは分けて扱います。
- 測定境界検査とreader／stream writer measurement validatorのself-testを実行し、すべて成功しました。未検証測定の公開artifact混入を防ぐlocal gateであり、3 OS実測やExcel oracleの証拠ではありません。
- 追跡済みformula planでは毎回の全formula整合性scanを省略し、`cells_mut()`等でtrackingが無効化された場合だけfallback検査するようにしました。dirty formula、cycle、manual→automaticの関連回帰を確認しています。
- 同じCriterion条件で再測定し、single-input chain 1.0217 ms、warm noop 0.96091 ms、独立1,000入力0.89674 msを記録しました。改善はtracked-edit fast pathに限定し、structure rebuild 1.3737 msの増加を含めて記録しています。
- G4の数式互換性補完として`IFNA`を追加しました。`#N/A`だけをfallbackし、`#DIV/0!`など他のErrorは伝播する遅延評価を回帰テストで固定しています。Excel oracle照合や関数全体の未対応mode棚卸しは未完です。
- G4の検索互換性補完として`XLOOKUP`のwildcard `match_mode=2`と、ソート済み数値範囲向けbinary `search_mode=2/-2`を追加しました。wildcardとbinaryの組み合わせや未知modeは明示的にエラーとし、Excel oracle照合と網羅的なbinary mode校正は未完です。
- G4の検索関数整合性として`XMATCH`にもwildcard `match_mode=2`と、ソート済み数値範囲向けbinary `search_mode=2/-2`を追加しました。未知modeは通常検索へ置換せず明示的にエラーとします。Excel oracle照合は未完です。
- G4のlookup境界を補完し、`HLOOKUP`の行番号0を`#VALUE!`、範囲外を`#REF!`として返すようにしました。`VLOOKUP`と同じ明示的な境界を回帰テストで固定しています。
- G4のlookup境界を補完し、`INDEX`の行列番号0を`#VALUE!`、範囲外を`#REF!`として返すようにしました。整数アンダーフローによる誤セル参照を防止しています。
- G4の`MATCH` exact検索で文字列wildcard `*` / `?`を大文字小文字非依存で扱うようにし、未定義match typeを`#VALUE!`として返すようにしました。
- `MATCH`のwildcard検索でExcel式のエスケープ`~*`、`~?`、`~~`をリテラルとして扱うようにしました。エスケープを含む場合もbounded DPで評価します。
- G4のdispatcher棚卸し基盤として、実際のformula dispatch tableからcanonical名・alias・重複名を抽出する`check-formula-dispatch.py`を追加しました。関数数の未測定な水増しを防ぐためのlocal gateであり、Excel oracle一致率は未測定です。
- dispatcher棚卸しに`--check-docs`を追加し、`FUNCTIONS.md`のworksheet関数表と実dispatchの218名（canonical 207 / alias 11）を照合できるようにしました。引数形・未対応mode・Excel oracle一致率は未測定です。
- `XLOOKUP` / `XMATCH`のmatch/search modeで非整数・非有限・範囲外の値を切り捨てず、`#VALUE!`として返すようにしました。
- `INDEX`で行番号0・列番号0を指定した配列形式を追加し、列・行・全範囲を`Variant::Array`として返すようにしました。worksheetへのspill配置は未実装です。

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
- G4の2D配列対応を拡張し、`UNIQUE`の行／列単位重複排除と`SORT`の行列ソートを追加して、結果shapeを既定spill再計算へ接続しました。exactly_once、型変換、Excel oracle照合は未完です。
- G4の2D配列対応を拡張し、`SORTBY`で2Dデータを複数の1列sort-by配列と各昇降順に従って行単位に並べ替え、結果shapeを既定spill再計算へ接続しました。型変換とExcel oracle照合は未完です。
- G4の2D配列対応を拡張し、生成された2D配列に対する`FILTER`の行／列includeと結果shape接続を追加しました。行／列ベクトル以外のincludeは明示エラーにし、既存Range経路を維持しています。複合includeとExcel oracle照合は未完です。
- G4の2D `UNIQUE`で`exactly_once`／`by_col`を既存のtruthy型変換に揃え、Boolean以外の数値フラグも扱えるようにしました。型意味論全体とExcel oracle照合は未完です。
