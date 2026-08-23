# 0.10.0 Lossless Worksheet Preservation — 設計調査

## Status

Draft — 実装未着手。この文書自体もまだcommitしていない（ユーザー指示: 「コード変更、commit、
pushは行わず、設計上の選択肢と推奨案を報告して停止する」）。`docs/xlsx-architecture.md`の
「Root-crate writer: regenerate vs. preserve-and-merge」節（0.8.0/0.9.0の既存決定）を前提とし、
その続きとして書く。単独ファイルにしたのは、既存ファイルが737行と既に大きく、この調査単体でも
それに匹敵する分量になるため——0.10.0が実装フェーズに入った時点で、要点を
`docs/xlsx-architecture.md`本体へ合流させることを推奨する。

## 調査方法

推測ではなく、実物で確認した。

- **ソース**: `src/lib.rs`の`is_writer_owned_part`/`save_xlsx_impl`/`build_xlsx_sheet`/
  `build_xlsx_workbook`/`xlsx_cell_xml`、`src/reader.rs`の`xlsx_sheet_cells`/`WorkbookSheet`/
  `XlsxSheetData`/`BufferSheet`を実際に読んだ（行番号は本文中に記載）。
- **実データ**: `compat/oracle-excel-com/fixtures/pristine/`の5つの実Excel authored fixtureを
  `unzip`/`python3 -m zipfile`で直接展開し、`fixture3`（table/validation/conditional）・
  `fixture4`（hyperlink/comment/definedName）・`fixture5`（chart/image/print、ただし後述の
  通り実際にはfreeze paneは含まれていなかった）のworksheet XML・`.rels`・
  `[Content_Types].xml`を生で確認した。これは0.9.0-Aで既に実Excel検証済みの資産であり、新たな
  実Excelでの動作確認は今回行っていない（設計調査段階のため）。
- **実測**: 4節の中心的な主張（worksheet-levelのrelationshipが「生き残るが誰にも参照され
  ない」状態になる）は導出だけで終わらせず、実際に`fixture3`をelixceeでロード→1セル編集
  →`--output`保存→`mechanical_check.py`実行、まで走らせて確認した（4節・9節に詳細と結果）。

## 1. worksheet XMLでelixceeが現在生成・所有している要素

`build_xlsx_sheet`（`src/lib.rs:915-1014`）が生成する`<worksheet>`要素の中身は、実質これだけ：

```
<worksheet xmlns="...spreadsheetml/2006/main">
  <cols>                      -- hidden列区間のみ（min/max/hidden="1"）。width等はなし
  <sheetData>
    <row r=".." [hidden="1"]>
      <c r=".." [s=".."] [t=".."]>[<f>..</f>]<v>..</v></c>
  <mergeCells>                -- mergeCell ref="..." の列挙
</worksheet>
```

ルート`<worksheet>`タグ自体もdefault namespace 1つのみで、`xmlns:r`（relationship参照に必須）・
`mc:Ignorable`・`xr`/`xr2`/`xr3`系のnamespaceは一切宣言されない（`src/lib.rs:920-923`のリテラル
文字列で固定）。

`xl/workbook.xml`（`build_xlsx_workbook`、`src/lib.rs:820-838`）も同様に極小で、
`<sheets><sheet name=".." sheetId=".." r:id="rIdN"/></sheets>`のみ。`<definedNames>`・
`<bookViews>`・`<calcPr>`・`<workbookPr>`・`<fileVersion>`・`<extLst>`はゼロ。

**この文書内での注記**: 「defined names」はworksheet埋め込みではなくworkbook.xml埋め込みの機能。
ユーザーの元の分類（低結合リストに含まれる）は範囲としては正しいが、実装箇所は
worksheet-XMLの再構築ロジックではなく workbook.xml側の再構築ロジック
（`build_xlsx_workbook`）になる——後述10節のmilestone分割で明記する。

## 2. 読み取っているが書き戻していない要素

このカテゴリには2つある。

- **`<dimension ref="...">`** — `reader.rs`の`XlsxSheetData::dimension`
  （`src/reader.rs:759-761`）→`BufferSheet::dimension`（`src/reader.rs:132-137`）として
  読み取られ保持されているが、`build_xlsx_sheet`は一度も`<dimension>`を出力しない。
- **`<sheet sheetId="...">`（workbook.xml側、worksheet-embeddedではない）** —
  `WorkbookSheet::sheet_id`（`src/reader.rs:20-24`、doc commentに明記: 「real, and not
  necessarily contiguous」）としてsourceの元のsheetIdが読み取られているが、
  `build_xlsx_workbook`（`src/lib.rs:829-834`）は`sheetId="{n}"`を**現在の列挙位置から
  regenerate**し、読み取った値を使わずに捨てている。この事実は既に別の場所
  （`src/snapshot.rs:10-15`のdoc comment）で明示的に文書化されている:
  「for any elixcee-written `.xlsx`, since this repo's own writer regenerates `sheetId`
  sequentially from current sheet order on every save」——つまりelixcee開発陣は既に
  この事実を把握した上で`snapshot.rs`側だけ対処済み（後述`stable_id`、6節）。

それ以外（`raw_style_indices`→`s="N"`、`merged_ranges`→`<mergeCells>`、
`hidden_rows`/`hidden_columns`→`hidden="1"`、`formulas`→`<f>`）は0.8.0/0.9.0で読み取り・
書き戻し双方とも実装済み——このカテゴリには入らない。

3節で挙げる要素は基本的に**そもそも読んでいない**（＝reader.rs側にパース処理が存在しない）
——0.10.0は「書き戻す」だけでなく「読む」ことから作る必要がある機能がほとんど、という
点は変わらない。

## 3. 完全に未知のまま失われる要素・属性・namespace（実fixtureで確認）

`fixture3_table_validation_conditional.xlsm`の`xl/worksheets/sheet1.xml`（実Excel出力、整形
済み抜粋）:

```xml
<worksheet xmlns="..." xmlns:r="..." xmlns:mc="..." xmlns:x14ac="..."
           xmlns:xr="..." xmlns:xr2="..." xmlns:xr3="..."
           mc:Ignorable="x14ac xr xr2 xr3" xr:uid="{92FAC149-...}">
  <sheetPr codeName="Sheet1"/>
  <dimension ref="A1:F5"/>
  <sheetViews><sheetView tabSelected="1" workbookViewId="0">
    <selection activeCell="D8" sqref="D8"/>
  </sheetView></sheetViews>
  <sheetFormatPr baseColWidth="10" defaultRowHeight="20"/>
  <sheetData>...</sheetData>
  <phoneticPr fontId="1"/>
  <conditionalFormatting sqref="F1:F5">
    <cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>700</formula></cfRule>
  </conditionalFormatting>
  <dataValidations count="1">
    <dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1"
                    sqref="E1" xr:uid="{BF4C2CDE-...}">
      <formula1>"Yes,No,Maybe"</formula1>
    </dataValidation>
  </dataValidations>
  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
  <tableParts count="1"><tablePart r:id="rId1"/></tableParts>
</worksheet>
```

`fixture4_hyperlink_comment_name.xlsm`は`<hyperlinks><hyperlink ref="D6" r:id="rId1" .../>`
（External URL、`.rels`に`Target="https://yahoo.co.jp/" TargetMode="External"`）と
`<legacyDrawing r:id="rId2"/>`（コメントの吹き出しマーカーを描くVML）を追加で持つ。

**訂正（レビュー指摘）**: `<hyperlink>`要素には`r:id`（`.rels`経由でexternal URL/fileや外部
workbookへ接続）と`location`（同一workbook内のシート/セル/defined nameを直接テキスト参照、
relationship不要）の2形態があり、SpreadsheetMLの仕様上は排他ではない。`fixture4`が
たまたま`r:id`のみの外部URL例だったため、当初「`<hyperlinks>`全体がrelationship依存」と
断定したのは言い過ぎだった。**このリポジトリには`location`属性を使う同一workbook内
ハイパーリンクの実例が現状ない**——3節の他の「実データ未確認」項目（freeze pane）と同様、
正直に「未実証」として扱う。10節のmilestone分割でこの区別を反映する。

**更新（0.10.0-A中に解消）**: `fixture6_internal_hyperlink.xlsm`をユーザー提供の実Excel
ファイルから追加、`<hyperlink ref="A1" location="Sheet2!B2" .../>`（`r:id`なし）の実例を
確認済み。詳細は`fixtures/pristine/INVENTORY.md`。

同fixtureの`xl/workbook.xml`には
`<definedNames><definedName name="test" comment="test desu!!!">Sheet1!$F$5</definedName></definedNames>`
（workbook-level、2節参照）。

`fixture5_chart_image_freeze_print.xlsm`は`<pageSetup .../>`と`<drawing r:id="rId1"/>`
（chart/image）を追加で持つ——ただし実際に`grep -o '<pane'`した結果、**このfixtureに
`<pane>`（freeze pane）要素は存在しなかった**。ファイル名は"freeze"を含むが、authoring時に
実際には設定されなかったか、別の保存経路で失われた可能性がある。**freeze paneについては
このリポジトリに実Excel authoredの正例が現状ない**——0.10.0-Bでこの機能に着手する前に、
新規fixtureで実データを確保する必要がある（正直に記録：机上のOOXMLスキーマ知識のみでの
実装は、このプロジェクトが再三経験してきた「synthetic fixtureでは検出できないreal-Excel
特有のバグ」を再現するリスクがある）。

**更新（0.10.0-A中に解消）**: `fixture7_freeze_pane.xlsm`をユーザー提供の実Excelファイルから
追加、`<pane xSplit="1" ySplit="1" topLeftCell="B2" activePane="bottomRight" state="frozen"/>`
の実例を確認済み。詳細は`fixtures/pristine/INVENTORY.md`。

まとめると、`build_xlsx_sheet`が一度も出力しない、かつreader.rsも一度もパースしない
worksheet-XML要素:

`<sheetPr>`・ルート`<worksheet>`のnamespace束/`mc:Ignorable`/`xr:uid`・`<sheetViews>`
（`<selection>`・`<pane>`含む）・`<sheetFormatPr>`・`<sheetProtection>`・`<autoFilter>`・
`<phoneticPr>`・`<conditionalFormatting>`・`<dataValidations>`・`<hyperlinks>`・
`<pageMargins>`・`<pageSetup>`・`<headerFooter>`・`<drawing>`・`<legacyDrawing>`・
`<tableParts>`・`<extLst>`。

workbook.xmlレベル: `<definedNames>`・`<bookViews>`・`<calcPr>`・`<workbookPr codeName="...">`
・`<fileVersion>`・`<extLst>`。

**明示的に0.10.0のスコープ外として宣言すべきもの**: `fixture5`で確認された`richData`機能
（`xl/richData/*.xml`、セルの`vm="1"`属性、Excelの「リンクされたデータ型」）。これは
チャート/画像/テーブルよりさらに複雑な依存グラフを持ち、実利用頻度も低い。opaque-fragment
機構（7節）の対象からも明示的に除外し、「未知のpart/属性として素通しはするが、worksheet XML
内の`vm`属性やrich-value参照そのものは復元しない」という診断（9節）付きの既知の非対応として
扱うことを推奨する。

## 4. worksheet relationshipと対象partの一覧 — 今回の調査で見つかった具体的な穴

`is_writer_owned_part`（`src/lib.rs:469-481`）は`xl/worksheets/*.xml`を**パターンマッチ**で
writer-owned扱いにするが、`xl/worksheets/_rels/sheetN.xml.rels`は`"xl/worksheets/"`の後に
`/`を含むためこのパターンに一致せず、**汎用passthroughループ（`src/lib.rs:596-641`）で
バイト単位でそのまま生き残る**。その参照先（`xl/tables/table1.xml`・`xl/drawings/drawing1.xml`
・`xl/charts/chart1.xml`・`xl/comments1.xml`・`xl/drawings/vmlDrawing1.vml`等）もwriter-owned
ではないため同様に生き残り、content-type宣言も0.9.0で実装済みの`carried_overrides`機構
（source自身の`[Content_Types].xml`から解決）で正しく引き継がれる——ここまでは機能している。

**しかし、regenerateされる`sheetN.xml`側は`<hyperlinks r:id>`・`<drawing r:id>`・
`<legacyDrawing r:id>`・`<tableParts><tablePart r:id>`のどれも一切出力しない。** つまり現状:

- `.rels`ファイルとその参照先partは構造的に正しく生き残る（`mechanical_check.py`の
  dangling/orphan検査は通る——partからの参照は生きているので）。
- だが regenerate された`sheetN.xml`のどの要素からも、その`.rels`内のrIdへの参照が存在しない。
  結果、`.rels`ファイル自体が**パッケージ内で誰からも参照されない孤立したpart**になる。

これは0.9.0のMilestone 4で見つかった「orphaned relationship」バグ（workbook-level
`.rels`が固定テンプレートで再生成されるため、themeやdocPropsへの関係が丸ごと消えていた）とは
**別の、まだ未修正の穴**——あちらは関係が消える話、こちらは関係もpartも生き残るのに
機能として無効化される話。

**実際に確認した（derivedのままにしなかった）**: `fixture3_table_validation_conditional.xlsm`
を実際にelixceeでロードし、1セルを編集するVBAマクロを実行して`--output`で別ファイルへ保存、
その結果を検証した。

- 出力ZIPには`xl/worksheets/_rels/sheet1.xml.rels`（`rId1`→`../tables/table1.xml`）と
  `xl/tables/table1.xml`本体が共にbyte-for-byteで残っている。
- しかし出力の`xl/worksheets/sheet1.xml`を`grep`すると、`tableParts`・`rId1`ともに
  **一致ゼロ**——導出通り、参照は本当に失われていた。
- `compat/oracle-excel-com/mechanical_check.py`をこの保存前後のペアに対して実行した結果は
  **`STRUCTURALLY_CLEAN`（violations: []）**。既存のorphan検査（Milestone 4で「orphaned
  part」を実際に検出した実績のある検査）は、この失敗モードを検出できない——理由は9節で
  補足するが、既存検査は「`.rels`が宣言する関係の対象partが存在するか」と「partが何らかの
  関係から参照されているか」は見ているが、「`sheetN.xml`の**中身**が、その`.rels`のrIdを
  実際に使っているか」は見ていないため。つまり4節のこの穴は、0.9.0時点の検証資産だけでは
  **静かに見逃され続ける**——0.10.0の作業対象であると同時に、検証ツール自体の拡張が
  同時に必要という発見。

ネストしたrelationship連鎖も確認: `xl/drawings/_rels/drawing1.xml.rels`
（→`chart1.xml`）・`xl/charts/_rels/chart1.xml.rels`（→`colors1.xml`/`style1.xml`）・
`xl/richData/_rels/richValueRel.xml.rels`（`fixture5`）。これらはchart1.xml自身が
byte-identicalでpassthroughされる限り内部的には自己無矛盾——壊れるのは常に
「writer-owned partから最初にrelationship graphへ入る入り口」（＝worksheet→drawing/table/
hyperlinkの一歩目）のみ。

## 5. content typesとの依存関係

**これは0.10.0の新規課題ではない** — 0.9.0の`carried_overrides`機構
（`src/lib.rs:607-640`、source自身の`[Content_Types].xml`をOverride優先→拡張子Default→
vbaProject限定の防御的フォールバックの順で解決）が、`fixture5`の豊富なOverride群
（`richData/*`・`charts/*`・`drawings/*`・`metadata.xml`まで全10種）をすべて正しく解決できる
ことを`unzip`での実データ確認で裏取り済み。4節の穴はcontent-type宣言の欠落ではなく、
純粋にrelationship graphの接続性（`r:id`参照）の欠落。

## 6. sheet rename/add/delete時に更新が必要な参照

現状、`sheet_names()`の**現在の並び順**から`build_xlsx_workbook`/`build_xlsx_workbook_rels`が
`sheet{i+1}.xml`という**位置ベース**の命名で毎回regenerateする（workbook.xml⇄workbook.xml.rels
⇄実partの3点は常にwriterが単独で完全に握っているため、この輪の中だけなら自己無矛盾で安全）。

問題は4節の穴を塞いだ後に生じる: worksheet-levelの`.rels`はsourceの**元のpart名**
（例:「削除前3番目のシートだった`sheet3.xml`」）に紐づいている。VBAでシートを削除・追加・
リネームした場合、現在の位置ベース命名では「削除前3番目のシート」が保存後に
`sheet2.xml`（新2番目）に化ける可能性があり、そこへ元`sheet3.xml`用の`.rels`
（ハイパーリンク/テーブル/チャート）を単純に位置で結びつけると**別シートに誤って
関係を付け替えてしまう**——4節を直しただけでは終わらない、6節を同時に設計しないと
新しい正しさのバグを作る。

推奨方針（8節でも詳述）: **worksheet-level rels/relationshipの持ち越しは、位置やindexでは
絶対にキーにしない。**

**訂正（レビュー指摘）**: 初稿では`src/snapshot.rs:10-33`の`stable_id`をそのまま
relationship持ち越しキーとして「流用」する案を書いたが、これは正確ではない。
`stable_id`は表示・識別用の文字列（`sheetId`があれば使い、なければ1-based位置に
フォールバックする）であり、`save_xlsx_impl`が必要とする「relationshipを安全に
再接続するための識別」とは要件が異なる——`stable_id`はあくまで
**「表示名（sheet name、リネームで変わる）と永続identity（sheetId、リネームで
変わらない）は別物である」という設計上の先例**として参照すべきものであり、その実装を
そのまま転用できるわけではない。

**worksheet originのidentityとして、save時に以下をセットで保持することを推奨する**:

```rust
struct WorksheetOrigin {
    original_sheet_id: Option<String>,       // <sheet sheetId="..">（workbook.xml側）
    original_workbook_rel_id: Option<String>, // workbook.xml.rels側のrId（workbook→sheet part）
    original_part_name: Option<String>,       // 例: "xl/worksheets/sheet3.xml"
}
```

**実装時の修正（0.10.0-A）**: 初稿にあった`stable_key: SheetKey`（rename後も変わらない
VM内部識別子）フィールドは実装しなかった。`src/vm/mod.rs`・`src/parser/`を`grep`した
結果、**このVMにはシートをリネームするVBA機能が現状一切存在しない**ことを確認——
今のところ、全ての per-sheet `Vm`マップ（`merged_ranges`・`sheet_visibility`・
`cell_style_indices`、そして`worksheet_origins`自身）は既にシート名（小文字化）を
キーにしており、rename機能がない以上これは既に安定したキーとして機能している。
検証されない抽象化を先回りして作るのは過剰設計と判断し、rename-safeな識別子は
sheet-rename VBA機能が実装される時まで見送ることにした。

**追記（0.10.0-C着手時に発覚・修正済み）**: この節冒頭で「位置ベースの命名」と表現していた
`sheet_names()`は、実際には**アルファベット順ソート**だった——add/delete/renameが一切
起きない、ロードしたファイルをそのまま保存するだけの最も単純なケースでも、シート名が
アルファベット順でなければ（例: "Zebra"→"Alpha"の2シート）保存のたびにタブ順が入れ替わる
という、この節が想定していたより一段階手前のバグだった。全ての既存fixtureが
たまたまアルファベット順（Sheet1/2/3）だったため発見が遅れた。`Vm::sheet_order`
（挿入順、`ensure_sheet`/`Sheets(...).Delete`で`sheets`と同期）を新設し、
`save_xlsx_impl`の並び順ソースをこちらに切り替えて修正済み——ただし`sheet_names()`
自体は`Sheets(i)`/`Worksheets(i)`のランタイム解決にも使われており、そちらの
アルファベット順は別の既知ギャップ（`docs/agent-contract.md`）として意図的に
未変更のまま残した。この節が本来警告していた「add/delete/rename後にworksheet-level
relsを安全に付け替える」設計（origin基準のcarry-over、0.10.0-D）はこの修正の対象外で、
引き続き未実装。

役割を明確に分ける: シート名（ユーザーに見える可変値）／`sheetId`（workbook内の識別子、
`.ods`など`sheetId`を持たないソースではNone）／workbook.xmlの`r:id`（workbook.xmlから
worksheet partへの関係）／worksheet part path（`xl/worksheets/sheetN.xml`という文字列、
現状は位置から再生成されているため最も不安定）——**この4つは別々の軸であり、どれか1つ
だけで代用しようとしない**（VM内部identityは上記の通り、現時点ではシート名自身がその
役割を兼ねる）。

特にrelationship carry-overでは、`sheetId`単独よりも**「元worksheet part path＋元workbook
relationship＋VM内部identity」の組み合わせで追跡する方が安全**——`sheetId`が欠落・重複
しうる実ファイル（`.ods`経由や手編集されたXML）に対しても、part pathとworkbook.xmlの
`r:id`という2つの独立した手がかりが残るため、1点の欠落だけで持ち越し判定全体が
壊れない。すべての手がかりが失われた場合にのみ（＝対応するsourceのsheetが
どの基準でも追跡できない場合にのみ）、持ち越しをスキップし9節の診断へ回す。

2節で触れた`build_xlsx_workbook`の`sheetId`位置ベース再生成（毎save時に列挙位置から
振り直している）はこの`WorksheetOrigin`設計と地続きの問題——既存シートの`sheetId`は
可能な限り保持し、新規シートにのみ既存と衝突しない新規idを振る、という修正を
0.10.0-Aの一部として合わせて設計する。

もう一点: `<definedNames>`（`fixture4`実例: `Sheet1!$F$5`）はシート名を**テキストとして
埋め込む**。VBAでシートをリネームした場合、opaque-fragment passthrough
（7節）のまま単純に持ち越すと定義名の参照が古いシート名を指したまま残り、silent
dangling referenceになる——これも9節の診断対象にすべき既知のギャップとして明示する
（自動書き換えは0.10.0のスコープ外を推奨。理由は次節）。

## 7. preserve-and-merge方式と全面再生成方式の比較

**ZIP part単位での方針は0.8.0で既に決定済み**（`docs/xlsx-architecture.md`該当節）。
0.10.0で新たに決めるべきは、**1つのwriter-owned XML part（worksheet/workbook）の"中"**で
個々の未知要素をどう扱うかという、一段細かい粒度の話。選択肢は2つ：

- **(a) 構造化モデル**: 各要素（`<hyperlinks>`等）をパースしてVm上の型に変換し、書き込み時に
  Rust側で再構築する。VBAから編集可能になる、が実装コストが高く、属性順序やホワイトスペース
  が原文と一致しなくなるリスクを新たに生む。
- **(b) opaque fragment passthrough**: パース時に要素を「どのCT_Worksheetスキーマ位置に
  属するか」というタグ付きの生XML文字列として捕捉し、書き込み時に該当スキーマ位置へ
  そのまま差し込む。編集はできないが、実装コストが低く、原文のバイトをほぼそのまま
  保持できる。

**確認した事実**: `src/vm/mod.rs`・`src/parser/*.rs`を`grep`した限り、Hyperlink・
FreezePanes・Comment・FormatCondition・Validation・definedNameを**VBAオブジェクトとして
操作する機能は現在ゼロ**（`Range.Interior.Color`/`Range.NumberFormat`と同様、そもそも
この方向のVBA表面が実装されていない）。つまりこれらの要素は今のVMにとって
「読んでも誰も使わない、書いても誰も作らない」データであり、(a)の構造化コストを正当化する
利用者側の需要が現状存在しない。

**推奨: 6節の`<hyperlinks>`/`<definedNames>`のような「参照先が変わりうる」要素を除き、
基本は(b) opaque fragment passthroughを採用する。** 将来VBAに`Range.AddComment`や
`ActiveWindow.FreezePanes =`等の対応を追加する時が来れば、その機能に限定して(a)へ
昇格させればよい——今それを先回りして全要素に適用するのは過剰設計であり、このセッション
全体を通じて一貫している「タスクが要求する以上の抽象化・投機的な作り込みをしない」方針
（本セッションの一連のスコープ判断・ROADMAP.mdの各所の「later slice」区切りに一貫して
表れている姿勢）に反する。

## 8. XML順序・namespace・relationship IDを壊さない設計

**更新（0.10.0-A、実XSDで確認済み）**: 初稿時点では下記の並びは記憶からの再構成で
未検証だった。0.10.0-A着手時に実際のECMA-376第5版XSD
（`OfficeOpenXML-XMLSchema-Transitional/sml.xsd`、`QtExcel/ecma-376-5th`のGitHubミラー経由
で取得——real-world Excelファイルはtransitional schemaに従う。strict schemaは実質使われ
ないため対象外）の`CT_Worksheet`定義を直接取得して突き合わせた。**結果、記憶ベースの並びは
`drawingHF`（`legacyDrawingHF`と`picture`の間、`CT_DrawingHF`型）が1つ丸ごと抜け落ちていた
以外は完全に一致**——「記憶を信用しない」というhard gate自体が実際に1つの間違いを
検出した、という結果になった。以下は実XSDで確認済みの並び（`fixture3`の実データで
直接位置関係を確認できたのは`sheetPr, dimension, sheetViews, sheetFormatPr, [cols,]
sheetData, [mergeCells,] phoneticPr, conditionalFormatting, dataValidations, pageMargins,
tableParts`の10〜12要素のみで、残りはXSD一次情報での確認）。

```
sheetPr, dimension, sheetViews, sheetFormatPr, cols, sheetData, sheetCalcPr,
sheetProtection, protectedRanges, scenarios, autoFilter, sortState,
dataConsolidate, customSheetViews, mergeCells, phoneticPr,
conditionalFormatting, dataValidations, hyperlinks, printOptions,
pageMargins, pageSetup, headerFooter, rowBreaks, colBreaks,
customProperties, cellWatches, ignoredErrors, smartTags, drawing,
legacyDrawing, legacyDrawingHF, drawingHF, picture, oleObjects, controls,
webPublishItems, tableParts, extLst
```

参考までに`CT_Workbook`（`xl/workbook.xml`、0.10.0-C対象）の並びも同じXSDで確認済み:

```
fileVersion, fileSharing, workbookPr, workbookProtection, bookViews, sheets,
functionGroups, externalReferences, definedNames, calcPr, oleSize,
customWorkbookViews, pivotCaches, smartTagPr, smartTagTypes, webPublishing,
fileRecoveryPr, webPublishObjects, extLst
```

7節(b)を採用する場合、各opaque fragmentにはこの列内での**スキーマ位置スロット名**を
タグ付けし、書き込み時はパース順ではなくこの固定順で再構成する必要がある——
`<cols>`/`<sheetData>`/`<mergeCells>`は既に正しい位置に書けている（現行コードのハードコード
順序がこの並びと矛盾しないことを確認済み）ので、新スロットはこの並びに挿入する形になる。

**namespace/ルート属性**: 現行`build_xlsx_sheet`はルート`<worksheet>`タグをリテラル文字列で
固定生成しており、default namespaceしか宣言しない。`r:id`属性を使う要素（hyperlinks/
drawing/legacyDrawing/tableParts）を一つでも復元するなら`xmlns:r`宣言が最低限必須になる。
推奨: sourceの元のルートタグの属性文字列（namespace宣言・`mc:Ignorable`・`xr:uid`）を
そのままキャプチャして再利用する——書き込み側で必要なnamespaceだけ選んで再構築するより、
「元のルート属性を丸ごと引き継ぐ」方が実装が単純かつ安全（未知の追加namespaceを個別に
把握する必要がなくなる）。

**relationship ID**: workbook-levelの`carry_over_rels`（`src/lib.rs:494-515`）は新しいrIdを
連番で振り直す——これはworkbook-level rels内にwriter自身が発行するworksheet/sharedStrings/
styles/vbaProject用のrIdと衝突しうるため必要な処置。**worksheet-level relsには
このような衝突が存在しない**（今のwriterはworksheet-level relationshipを一つも自分で
発行していない）ため、より単純な方針で足りる: **sourceの`.rels`ファイルをバイト単位で
無変更のままpassthroughし、opaque fragment内の`r:id="rIdN"`もパース時の文字列をそのまま
再利用する**（renumberしない）。6節の`WorksheetOrigin`による持ち越し判定さえ正しければ、
ID自体は触らないのが最小リスク。

## 9. unsupported featureをsilent lossさせない診断方法

既存の実装パターンを流用すべき、新規発明は不要——`src/check.rs`に既に
`Diagnostic { severity, code, kind, message, location }`
（`src/check.rs:78-95`、`kind: "unsupported_construct"`、code `I1002`）というVBAコンパイル時の
構造化診断の型がある。これをそのまま再利用し、**save時専用の新しいcode系列**
（例: `I3xxx`/`W3xxx`、既存のVBAコンパイル時`E1xxx`/`I1xxx`系列と衝突しない別namespace）を
割り当てることを推奨する。

現状`save_workbook`/`save_xlsx_impl`は`Result<(), String>`で成否のみを返し、「成功したが
一部要素を保持できなかった」を伝える経路がない。0.10.0-Aで`Vec<Diagnostic>`を返すよう
シグネチャを拡張し（PyO3側・CLI側どちらも新しいオプショナルフィールドとして追加可能、
既存の呼び出しコードを壊さない設計にできる）、次のケースで診断を出す:

- 6節の「`WorksheetOrigin`のどの手がかりでも一致せずrelationship持ち越しをスキップした」場合
- 3節でスコープ外と宣言したrichData等の未対応要素をsourceで検出した場合
- freeze pane等、reader側にまだパース実装がない要素をsourceのXMLに検出した場合
  （＝「存在は検知できるが復元はできない」ことを明示——検知すらできないよりずっと良い）

**`mechanical_check.py`自体の拡張も0.10.0-Aのスコープに含めることを推奨する。** 4節で
実際に確認した通り、既存のorphan検査は「`.rels`の関係とその対象partが存在するか」しか
見ておらず、「`sheetN.xml`の中身が実際にそのrIdを参照しているか」は見ていないため、
今回発見した失敗モードを`STRUCTURALLY_CLEAN`として素通ししてしまう。

**訂正（レビュー指摘）**: 初稿では「`.rels`の全rIdがworksheet XML内に文字列として
存在するか」という単純検査を提案したが、これは採用しない——relationship typeによって
worksheet側での参照のされ方が異なるため、文字列一致だけでは誤検知/見逃しの両方を生む
（例えば`r:id`という文字列自体は`<hyperlinks>`にも`<drawing>`にも`<tableParts>`にも
現れるが、`type="table"`の関係が`<drawing r:id="rId1">`から「たまたま」参照されていても
それは無関係な関係を誤って正当と判定してしまう）。**relationship type ごとに、
worksheet側のどの要素のどの属性が参照元になり得るかを明示的に対応させる**
（type-aware mapping）:

**実装済み（0.10.0-A）**: 以下の表は`compat/oracle-excel-com/mechanical_check.py`の
`_WORKSHEET_RID_ELEMENT_XPATHS`として既に実装され、`check_source_references()`
（新関数）がこの表を使って`SOURCE_REFERENCE_LOSS`を検出する。実fixture（`fixture3`/
`4`/`5`）とECMA-376実XSD（8節参照）の両方で裏取りした4種のみを実装し、fixtureが
存在しない行は実装しなかった——「分かっている範囲を正しく書く」の実践。

| relationship type | worksheet側の参照元 | 根拠 |
|---|---|---|
| table | `<tableParts><tablePart r:id="..">`（`r:id` required） | fixture3実データ＋`CT_TablePart`実XSD |
| drawing | `<drawing r:id="..">`（`r:id` required） | fixture5実データ＋`CT_Drawing`実XSD |
| hyperlink（`r:id`形式） | `<hyperlinks><hyperlink r:id="..">`（`r:id` optional、`location`と排他ではない） | fixture4実データ＋`CT_Hyperlink`実XSD |
| vmlDrawing | `<legacyDrawing r:id="..">`（`r:id` required。`legacyDrawingHF`も同型だが未実装fixtureのため対象外） | fixture4実データ＋`CT_LegacyDrawing`実XSD |

**実装しなかった行（fixtureなし、XSDのみ確認済み）**: printerSettings →
`<pageSetup r:id="..">`（`CT_PageSetup`の`r:id`は`optional`属性）／oleObject →
`<oleObjects><oleObject r:id="..">`（**訂正**: `r:id`は`<oleObject>`要素自身に直接
付く`optional`属性——初稿では入れ子の`<objectPr>`側についていると誤って想定していたが、
実XSD確認で訂正）／control → `<controls><control r:id="..">`（`r:id`は`<control>`
自身に`required`）。いずれも実fixtureが存在するまで実装しない（10節のhard gate通り）。

**comments・threadedComments（実測で確定、初稿の「個別確認が必要」を解消）**:
`fixture4`の`xl/worksheets/_rels/sheet1.xml.rels`が持つ4つの関係
（table…ではなくhyperlink=rId1／vmlDrawing=rId2／comments=rId3／threadedComment=rId4）
のうち、`sheet1.xml`本文が実際に`r:id`として参照しているのは`rId1`と`rId2`だけ——
**`rId3`（comments）と`rId4`（threadedComment）はsheet1.xml本文のどこにも一切現れない**
ことを`grep`で直接確認した。つまりcomments/threadedCommentsはworksheet content側の
`r:id`参照を一切必要とせず、`.rels`ファイル自身のType属性だけで存在が決まる
——この2種類は`SOURCE_REFERENCE_LOSS`検査の対象外とし（対応表に含めない）、
既存の`ORPHANED_PART`検査（`.rels`グラフレベルの参照有無のみを見る）がそのまま
正しくカバーする、という結論で確定した。

**新しい違反分類を追加する**（既存の`ORPHANED_PART`/`DANGLING_RELATIONSHIP`とは別カテゴリ
として区別する——今回見つかったのはどちらでもなく、3つ目の新しい種類の壊れ方）:

- `SOURCE_REFERENCE_LOSS` — `.rels`と対象partはどちらも存在するが、regenerateされた
  writer-owned part（`sheetN.xml`等）のどこからもそのrIdが参照されていない（今回の発見）。
- `DANGLING_RELATIONSHIP` — 既存（`.rels`が指すtargetが存在しない）。
- `ORPHANED_PART` — 既存（partが存在するがどの`.rels`からも参照されていない、Milestone 4で
  発見済み）。

**実装済み（0.10.0-A）**: `mechanical_check.py --self-test`のCase Hとして、実装した4種
（table/drawing/hyperlink/vmlDrawing）それぞれについて、元の`.rels`と対象partを
そのまま残した状態でworksheet側の参照要素だけを個別に取り除いた破壊fixtureを用意し、
全て`SOURCE_REFERENCE_LOSS`として検出されることを確認済み。同じCase Hで、
comments relationship（`rId5`、対応表に含まれない未マップtype）を持つがsheet1.xml
から一切参照されないケースが正しく`CLEAN`と判定される（誤検知しない）ことも
併せて確認——「comments/threadedCommentsをこの検査の対象外とする」という設計判断
そのものを固定するregression guardになっている。printerSettings/oleObject/control
は表自体を実装していないため、対応する破壊ケースも実装していない（fixture確保が
先——10節のhard gate通り）。

**実fixtureへの適用（0.10.0-A、実測で確定）**: self-testが通った後、実際に
`fixture1`〜`5`それぞれをelixceeでロード→1セル編集→保存→`check_source_references()`
を実行した。`fixture1`/`fixture2`（worksheet-level relationshipを一切持たない）は
`CLEAN`。**worksheet-level relationshipを持つ`fixture3`/`4`/`5`は全て
`SOURCE_REFERENCE_LOSS`を検出**——`fixture3`（table）・`fixture4`
（hyperlink・vmlDrawingの両方が同一保存で同時に消失）・`fixture5`（drawing）。
これは当初4節で確認した「1つの実例」ではなく、**worksheet-level relationshipを持つ
実fixtureの100%で再現する、体系的な欠落**であることが確定した。

これは0.9.0-Aの「macro実行結果はNOT_EVALUATEDと書く、擁護的な言い回しをしない」という
標準方針、および`compat/differential/classify.mjs`のORACLE_AMBIGUITY/NONDETERMINISTIC
判定規律と同じ精神をsaveパスにも広げるもの——「黙って消える」を「消えたことが分かる」に
変える。

## 10. 0.10.0milestoneへの分割

**訂正（レビュー指摘）**: 初稿の3分割（A: architecture / B: 低結合 / C: relationship-heavy）
は、`<hyperlinks>`を一括で低結合とみなした誤りを含んでいたほか、B/Cの境界がやや粗かった。
以下の4分割へ改める——分け方の原則は「relationship graphに触れるかどうか」を最優先の
軸にし、その中で「新規fixture/検証基盤が要るか」「workbook-level か worksheet-level か」を
副軸にする。各milestoneは独立にmerge/revert可能。

**着手前のhard gate（全milestone共通）**: このセッション・0.9.0で繰り返し確認された
教訓（synthetic fixtureや記憶ベースの知識だけでは real Excelのrepair対象になるバグを
見逃す）を踏まえ、要素単位で以下を満たすまでwriterコードを書かないことを明文化する。

> No writer code may be added for an element until:
> 1. an Excel-authored fixture exists,
> 2. its actual XML and relationships are recorded,
> 3. the applicable XSD sequence is confirmed,
> 4. mechanical_check has a negative test for its loss.

- **0.10.0-A — Foundation**（新機能ゼロ、writer機能を増やさない）
  - [x] fixture inventoryの棚卸し（`fixtures/pristine/INVENTORY.md`、新規）——既存5
    fixtureが実際に何を含み何を含まないかをスクリプトで一覧化、3節で発覚した
    「fixture5に実はfreeze paneがない」を含め、filenameを信用しない棚卸しを確定した。
    副産物として`fixture5`に`_xlnm.Print_Area`（builtin defined name）があることも
    新規発見（0.10.0-C対象）。
  - [x] relationship type → worksheet側source element 対応表の確定（9節）——実装した
    4種（table/drawing/hyperlink/vmlDrawing）はfixture実データ＋実XSD両方で裏取り、
    printerSettings/oleObject/controlはXSDのみ確認しfixture不在のため未実装、
    comments/threadedCommentsは「worksheet content側に`r:id`参照が一切ない」ことを
    実測で確定し対応表から除外。
  - [x] CT_Worksheet/CT_WorkbookのXML/XSD順序の確定（8節）——実際の ECMA-376第5版XSD
    （`sml.xsd`、`QtExcel/ecma-376-5th`経由）を取得し突き合わせ。記憶ベースの並びは
    `drawingHF`が1要素欠落していた以外は正しかった——hard gate自体が実際に1件の
    誤りを検出した。
  - [x] `mechanical_check.py`への`check_source_references()`実装と
    `SOURCE_REFERENCE_LOSS`違反分類の追加（9節）。
  - [x] negative self-testの追加（`self_test()`のCase H、4種の破壊ケース＋comments
    誤検知なしの確認）——`--self-test`で全green。
  - [x] 実fixtureへの適用——`fixture1`〜`5`全てで実際に確認した結果、`fixture3`/`4`/
    `5`（worksheet-level relationshipを持つ全fixture）で`SOURCE_REFERENCE_LOSS`を
    確認、`fixture1`/`2`は`CLEAN`。当初の「1実例」から「体系的な欠落」へ確度が上がった。
  - [x] `WorksheetOrigin`のidentity設計（6節）の実装——`reader::WorkbookSheet`に
    `workbook_rel_id`/`source_part_name`を追加（`rid`/`zip_path`は元々`read_workbook_from_archive`
    内で計算されていたが破棄されていた値）、`vm::WorksheetOrigin`
    （`original_sheet_id`/`original_workbook_rel_id`/`original_part_name`、6節の初稿から
    `stable_key`フィールドは削除——VBAにsheet rename機能が現状ゼロと確認済みのため、
    rename-safeな別identityは時期尚早と判断）、`Vm.worksheet_origins`
    （`merged_ranges`等と同じsheet名キーのHashMap）を実装。`build_xlsx_workbook`
    （`src/lib.rs`）を修正し、既存シートは元の`sheetId`をそのまま再利用、新規シートのみ
    既存の最大idを超える新しいidを割り当てるよう変更——`snapshot.rs`の`stable_id`が
    前提としていた「`sheetId`はcross-save-stable」という約束を、writer側が実際に
    満たしていなかった問題を修正。`r:id="rIdN"`の位置ベース採番はwriter内部で
    自己完結しているため変更していない。
  - [x] 実データでの確認——2シート・非連番`sheetId`（5・9）を持つ合成`.xlsx`を実際に
    elixceeでロード→1セル編集→保存し、出力`workbook.xml`で両シートとも元の`sheetId`が
    保持されることを確認。この検証中、ロード元に存在しない3つ目のシート
    （`Vm::new()`のデフォルト`"sheet1"`がロード時に上書きされず残ったもの）が出力に
    混入する事象を発見——当初は「本タスクとは無関係の既存動作」として報告のみに
    留めたが、根本原因を追ったところ**再現性のある独立したバグ**と判明したため、
    別コミットとして修正した（次項）。
  - [x] **発見事項の修正——`populate_from_sheets`が`Vm::new()`のデフォルト`"sheet1"`を
    クリアしていなかったバグ**。`Vm::new()`はワークブック未ロード時にもマクロが
    セルへ書き込めるよう、常に空の`"sheet1"`を`self.sheets`へ事前挿入する。
    `populate_from_sheets`は実際に読み込んだシートを`ensure_sheet`で*追加*するだけで
    このデフォルトを削除しないため、**ロードしたワークブックのどのシートも
    "Sheet1"という名前でない限り、保存のたびに空シートが1枚混入していた**——
    5つの実Excelフィクスチャで顕在化しなかったのは、たまたま全て`Sheet1`という
    シートを含んでいたため。`populate_from_sheets`（`src/vm/mod.rs`）の先頭に
    `self.sheets.clear()`を追加して修正（呼び出しは`Vm::new()`直後・マクロ実行前の
    一度きりであることを呼び出し元3箇所で確認済みのため、データ消失リスクなし）。
    回帰テスト3件を追加、および元と同じ合成fixture（"First"/"Second"、`sheetId` 5・9）
    で実CLIを再実行し、出力シート数が2枚（3枚ではなく）であることを確認。
  - [x] internal hyperlink（`location`属性）のfixture新規作成——ユーザー提供の実Excel
    ファイル（`fixture6_internal_hyperlink.xlsm`）を`compat/oracle-excel-com/fixtures/pristine/`
    へ追加。`<hyperlink ref="A1" location="Sheet2!B2" .../>`（`r:id`なし、relationship非依存）
    を実際に含むことを確認済み。`docProps/core.xml`の作成者名と`xl/workbook.xml`の
    `x15ac:absPath`（ローカルパス含む）は既存fixtureと同じ慣行でコミット前にscrub済み。
  - [x] freeze paneのfixture新規作成——同様にユーザー提供の実Excelファイル
    （`fixture7_freeze_pane.xlsm`）を追加。`<pane xSplit="1" ySplit="1" topLeftCell="B2"
    activePane="bottomRight" state="frozen"/>`を実際に含むことを確認済み。同じscrubを適用。
    `INVENTORY.md`を両fixture追加分で更新（「全5fixtureで確認absent」の記述も、この2件が
    もはや当てはまらないため「全7fixture」へ修正）。
  - 検証: `cargo test --workspace`（833件）・`cargo fmt --check`・
    `cargo clippy --all-targets`（python feature有無両方）・`cargo doc
    --document-private-items`、いずれもクリーン。`compat/corpus`（581件、0
    UNEXPLAINED/0 MISMATCH）・`compat/vba-semantics`（386件、0 BUG/0 UNCLASSIFIED）とも
    無変化を確認。
- **0.10.0-B — Inline worksheet preservation**（relationship非依存、7節(b)の
  opaque-fragment機構をworksheet-XML側へ適用）。1要素ずつ独立にmerge可能なslice単位で
  進める（0.10.0-Aの`fixture inventory→checker→writer`という順序をslice単位で反復）。
  - **B1（実装済み）— `<sheetViews>`（`<pane>`＝freeze pane・`<selection>`）＋ルート
    `<worksheet>`タグの属性文字列引き継ぎ**
    - [x] checker先行実装: `mechanical_check.py`に`check_inline_worksheet_elements()`
      （新しい違反分類`INLINE_ELEMENT_LOSS`、`SOURCE_REFERENCE_LOSS`ともcheck_roundtrip
      とも別軸）を追加、`_INLINE_WORKSHEET_ELEMENTS = ["sheetViews"]`から開始。
      self-testのCase Iで検出を確認してからwriterへ着手（0.10.0-Aと同じ順序）。
    - [x] writer未着手の時点で実fixture7種全てに`check_inline_worksheet_elements()`を
      適用し、全シートで`<sheetViews>`が体系的に失われることを実測確定（pre-fix
      baseline）——0.10.0-Aの「先にcheckerで実バグを実測してから直す」を踏襲。
    - [x] `reader.rs`に`extract_root_attrs`/`extract_raw_element`（共有の
      `find_next_open_tag`スキャン primitive経由）を実装——生バイトをそのまま
      切り出すopaque-fragment抽出、フルXMLパーサーではない（7節(b)通り）。単体テスト
      8件（自己閉じタグ・属性なしルート・名前不一致・namespace prefix等の境界値）。
    - [x] `save_xlsx_impl`に`sheet_source_xml`（sheet名キー、`WorksheetOrigin.
      original_part_name`経由で`raw_entries`から解決）を追加、`build_xlsx_sheet`へ
      `root_attrs`/`sheet_views`の2つの`Option<&str>`として橋渡し（advisor助言通り、
      3つ目のスロットが増えるまで構造体化は見送り）。XSD順序（8節）通り、ルートタグ
      直後・`<cols>`より前に`<sheetViews>`を挿入。origin不明のシート（新規シート・
      `.ods`ソース）は従来通りhardcodedのminimalなルートタグにフォールバック。
    - [x] 実fixture7種全てに対する事後確認——全シートで`check_inline_worksheet_elements`
      が`CLEAN`に、`check_roundtrip`（`STRUCTURALLY_CLEAN`）・`check_formula_preservation`
      （`CLEAN`）は不変。`check_source_references`の`SOURCE_REFERENCE_LOSS`は
      fixture3/4/5で意図通り変化なし（0.10.0-Dの担当、このsliceの回帰ではない）。
    - [x] `tests/xlsx_roundtrip.rs`に実fixture7（freeze pane）を使った専用テスト
      `real_excel_freeze_pane_sheetviews_survive_a_save`を追加——`<pane .../>`が
      byte-for-byteで生き残ること、ルートタグがsource由来のnamespace宣言を持つことを
      直接assert。既存の合成fixture（`<sheetViews>`を持たない）ベースのテストは
      無変化で通過を確認（advisor指摘の懸念は実害なしと確認済み）。
    - [x] **実Excel検証済み（ユーザー確認、2026-08-23）**: `<sheetView
      workbookViewId="0">`は`xl/workbook.xml`側の`<bookViews>`（0.10.0-C対象、
      現状writerは一切出力しない）への暗黙の参照だが、`<bookViews>`なしでも実Excelは
      repair警告なしに開けることを確認——`elixcee_freeze_pane_check.xlsm`
      （fixture7をload→A1編集→save）を実際にExcelで開き、(1)修復警告なし、
      (2) freeze pane（1行目・1列目固定）が復元されている、(3) A1が12345、の3点を
      確認済み。0.9.0-Aと同じ完了条件（実Excel再オープンでrepair警告0件）をB1は
      満たした——**B1は実装・実Excel検証ともに完了**。
    - 検証: `cargo test --workspace`（841件）・`cargo fmt --check`・
      `cargo clippy --all-targets`（python feature有無両方）・`cargo doc
      --document-private-items`、いずれもクリーン。`compat/corpus`（581件）・
      `compat/vba-semantics`（386件）とも無変化。
  - **B2（実装済み）— `<sheetPr>`・`<sheetFormatPr>`・`<phoneticPr>`・
    `<dataValidations>`**
    - [x] checker先行実装: `_INLINE_WORKSHEET_ELEMENTS`に4要素追加、self-testの
      Case J（1要素ずつ独立に破壊し、他の要素を誤検知しないことも確認）を追加。
      writer未着手の時点で実fixture7種に適用し、体系的な欠落を実測確定
      （sheetFormatPr/phoneticPrは全7fixture、sheetPrはfixture2/3/4、
      dataValidationsはfixture3のみ——sheetViews自体はB1修正のままCLEANを再確認）。
    - [x] `build_xlsx_sheet`の引数を2つの`Option<&str>`から
      `OpaqueWorksheetFragments`構造体（advisor助言通り、3つ目以降のスロットが
      増えた時点で構造体化）へ変更。XSD順序（8節）通り、`sheetPr`→`sheetViews`→
      `sheetFormatPr`をルートタグ直後・`<cols>`より前に、`phoneticPr`→
      `dataValidations`を`<mergeCells>`より後・`</worksheet>`直前に挿入
      （`conditionalFormatting`は意図的にスキップ——スロット間の順序はXSD上、
      存在する要素同士の相対順序のみが問題になるため省略可）。
    - [x] 実fixture7種全てに対する事後確認——全シートで`inline_elements`が
      `CLEAN`に。`structural`・`formulas`・`source_references`（fixture3/4/5の
      `SOURCE_REFERENCE_LOSS`含め）は無変化。
    - [x] `tests/xlsx_roundtrip.rs`に`real_excel_sheetpr_and_data_validations_survive_a_save`
      を追加（fixture3——sheetPr/sheetFormatPr/phoneticPr/dataValidationsの4つ全てを
      持つ唯一の実fixture）。`dataValidations`内の`<dataValidation>`が`xr:uid`属性を
      持つため、B1のroot_attrs引き継ぎ（`xmlns:xr`宣言）が無いと不正なXMLになる
      ——この依存関係も込みで丸ごと生き残ることをassert。
    - 検証: `cargo test --workspace`（841件）・`cargo fmt --check`・
      `cargo clippy --all-targets`（python feature有無両方）・`cargo doc
      --document-private-items`、いずれもクリーン。`compat/corpus`（581件）・
      `compat/vba-semantics`（386件）とも無変化。
    - [x] **実Excel検証済み（ユーザー確認、2026-08-23）**: `elixcee_datavalidation_check.xlsm`
      （fixture3をload→B2編集→save）を実際にExcelで開き、(1)修復警告なし、
      (2) E1セルのデータ入力規則ドロップダウン（Yes/No/Maybe）が復元されている、
      (3) B2が999、の3点を確認済み。0.9.0-Aと同じ完了条件をB2も満たした
      ——**B2は実装・実Excel検証ともに完了**。
  - **B3（実装済み）— `<pageMargins>`**
    - [x] checker先行実装: `_INLINE_WORKSHEET_ELEMENTS`に追加、self-testの
      Case Jの合成fixture・mutation dictへ組み込み。writer未着手の時点で
      実fixture7種に適用し、体系的な欠落を実測確定——全7fixtureで`pageMargins`
      のみが検出され、slice 1/2で直した要素は引き続きCLEANであることも確認。
    - [x] `OpaqueWorksheetFragments`へ`page_margins`フィールドを追加。XSD順序
      （8節：`dataValidations, hyperlinks, printOptions, pageMargins`）通り、
      `phoneticPr`→`dataValidations`の直後・`</worksheet>`直前に挿入
      （`hyperlinks`/`printOptions`は未実装のためスキップ）。
    - [x] 実fixture7種全てに対する事後確認——全シートで`inline_elements`が
      `CLEAN`。`structural`・`formulas`・`source_references`は無変化。
    - [x] `tests/xlsx_roundtrip.rs`の`real_excel_freeze_pane_sheetviews_survive_a_save`
      （fixture7）へ`pageMargins`のassertionを追加。
    - 検証: `cargo test --workspace`（841件）・`cargo fmt --check`・
      `cargo clippy --all-targets`（python feature有無両方）・`cargo doc
      --document-private-items`、いずれもクリーン。`compat/corpus`（581件）・
      `compat/vba-semantics`（386件）とも無変化。
  - **B4（実装済み）— internal hyperlinkの`location`属性**
    - [x] 実XSD確認: `CT_Hyperlinks`の`<hyperlink>`は`minOccurs="1"`——全部r:id
      形式だった場合は`<hyperlinks>`ごと省略必須、空`<hyperlinks/>`は不正なXML。
    - [x] checker先行実装: `_INLINE_WORKSHEET_ELEMENTS`への単純追加ではなく、
      専用の`check_internal_hyperlinks()`（子要素単位で`ref`照合、r:id保持
      children は対象外、出力側の空`<hyperlinks/>`自体も別途違反として検出）を
      新規実装。理由: 単純なpresent/absent比較だとfixture4（全r:id-only、
      `<hyperlinks>`がsourceにはあるが正しい出力には無いのが正しい）を誤検知
      するため。self-testのCase Kで合成の混在fixture（r:id 1つ＋location-only
      1つ）を用意し、(a) location-only側の欠落検出、(b) r:id側は誤検知しない
      こと、(c) 空`<hyperlinks/>`自体の不正検出、の3点を確認。writer未着手の
      時点で実fixtureに適用——fixture4（全r:id）は正しくCLEAN、fixture6
      （location-only）はcurrent writerで検出（真の欠落）を確認。
    - [x] `reader.rs`に`extract_relationship_free_hyperlinks`を実装
      （`find_next_open_tag`・既存の`parse_attrs`/`attr_get`を再利用、`r:id`の
      有無で子要素をフィルタし、該当する生の`<hyperlink .../>`スパンのみを
      `Vec<String>`で返す——コンテナ全体のバイトコピーではない、B1〜B3とは
      異なるパターン）。単体テスト5件（不在・全location・全r:id・混在
      ——synthetic、位置に依存しないこと・複数件の順序保持）。
    - [x] `OpaqueWorksheetFragments`へ`internal_hyperlinks: &'a [String]`を追加
      （他フィールドと違い`Option<&str>`ではなく、`build_xlsx_sheet`側で
      `<hyperlinks>`/`</hyperlinks>`ラッパーを空判定込みで合成）。XSD順序通り
      `dataValidations`の後・`pageMargins`の前に挿入。
    - [x] 実fixture7種全てに対する事後確認——全シートで`internal_hyperlinks`が
      `CLEAN`（fixture6の実出力を直接確認: `<hyperlinks>`コンテナが正しく
      再構成され、`xr:uid`属性込みでbyte一致。fixture4の実出力を直接確認:
      `<hyperlinks>`が文字列としても一切出現しない——空タグを出さず完全省略）。
      `structural`・`formulas`・`source_references`・`inline_elements`は無変化。
    - [x] `tests/xlsx_roundtrip.rs`に`real_excel_internal_hyperlink_survives_a_save`
      （fixture6、location-onlyの生存確認）と
      `real_excel_external_only_hyperlink_omits_the_hyperlinks_container_entirely`
      （fixture4、`<hyperlinks`という文字列すら出現しないことの否定的確認）を追加。
    - 検証: `cargo test --workspace`（846件）・`cargo fmt --check`・
      `cargo clippy --all-targets`（python feature有無両方）・`cargo doc
      --document-private-items`、いずれもクリーン。`compat/corpus`（581件）・
      `compat/vba-semantics`（386件）とも無変化。
    - **注記（正直な限界の記録）**: 混在`<hyperlinks>`コンテナ（r:id形式と
      location形式が同一シートに同居するケース）は実fixture上の実例がまだ
      ない。実装・検証は「2つの確認済み端点（fixture6=全location-only、
      fixture4=全r:id-only）＋実XSDのminOccurs制約」からの一般化であり、
      混在ケース自体はCase Kのsynthetic self-testのみで検証されている
      （fixture6/4個別の実データ確認と組み合わせても、真に混在した実
      Excelファイルでの動作は未確認）。
  - **B5以降（未着手）**: `<autoFilter>`（実fixtureにstandalone要素としての
    実例が現状なし、fixture新規待ち）・行/列プロパティ（`<cols>`/`<row>`の
    幅・スタイル等、hidden以外の属性）。
    - **調査メモ（B4完了後の`/greenlane`ラウンドで発見、未実装）**: `width`属性の
      実証拠は実際に存在する——`fixture1_values_styles_merge_hidden.xlsm`の隠し列が
      `<col min="4" max="4" width="0" hidden="1" customWidth="1"/>`を持つ。ただし
      これはB1〜B4のopaque-fragment機構（未知の要素をまるごと切り出して差し込むだけ）
      とは性質が異なる——`<col>`自体は既にB1以前から`hidden_columns`（Vm state）
      駆動でwriterが生成している既存要素であり、`width`/`customWidth`はその
      **既存の生成ロジックに追加attributeを持ち込む**形になる。source側の`<col
      min max>`範囲と、Vmが持つ`hidden_columns`の範囲が常に1:1で対応する保証は
      ない（macroが隠し列を追加/変更した場合の扱いが未設計）ため、B1〜B4と同じ
      「引き継ぐだけ」では済まず、方針決定が必要——実装せず保留する。
  - **`<conditionalFormatting>`は要注意——`dxfId`（`xl/styles.xml`の`<dxfs>`）や
    `<extLst>`拡張を参照する場合があり、完全な参照非依存とは言い切れない。** 最初は
    「参照先の妥当性検証はせず、生のsubtreeをそのまま保存する」raw subtree preservation
    として扱い、`dxf`参照の検証・復元は将来の別スライスへ回すことを推奨する。
  - **`<dimension>`は0.10.0-Bのスコープに含めない（advisor指摘で確定）**——2節で
    「読んでいるが書き戻していない」として挙げていたが、これはセル値から導出される
    データであり、view状態のようなopaque-fragment passthrough向きのデータではない。
    `BufferSheet`止まりで`WorkbookSheet`/`Vm`まで到達しておらず、マクロが元の範囲外へ
    書き込んだ場合はsource値をそのまま出すと逆に古い情報を確定的に出してしまう。
    やるならcurrent cellsから再計算する別種の作業になるため、0.10.0-Bには含めない。
- **0.10.0-C — Workbook-level preservation**（worksheet XMLではなくworkbook.xml側、
  1節の指摘通りスコープを混ぜない）。sliceに分割し、位置非依存の要素から先に着手
  （C1）、位置依存の要素は個別に設計してから着手（C2/C3）——0.10.0-B同様、hard gate
  （実fixture確認→XSD確認→checker→writer）を1要素ずつ通す。
  - **C1（完了）**: `<workbookPr>`・`<calcPr>`・`<extLst>`、およびroot`<workbook>`タグ
    自身のnamespace宣言（opaque root_attrs passthrough、0.10.0-B同様の仕組み）。
    3要素とも位置非依存（シート順・シート数と無関係）で、fixture1〜7全てに
    fixture実データがある。`mechanical_check.py`に`check_workbook_elements()`と
    `WORKBOOK_ELEMENT_LOSS`カテゴリを追加（`INLINE_ELEMENT_LOSS`とは別カテゴリ——
    こちらはworkbook.xmlという単一の固定パスpartが対象で、
    `_sheet_name_to_part`によるシート名マッチングが不要）。fixture1〜7全てで
    実装前`WORKBOOK_ELEMENT_LOSS`確認→実装後`CLEAN`確認済み。
    **実Excel検証済み（ユーザー確認、2026-08-23）**: 詳細はC3の項目末尾にまとめて記載
    （C1〜C3は同じ2fixture・3出力ファイルで一括検証したため）。
  - **C2（完了）**: `<bookViews>`。**訂正（当初案からの変更）**: 当初はadd/delete
    発生時に備えてシート順・シート数の一致を条件にverbatim持ち越しをgateする設計を
    検討したが、fixture1〜7全てを確認した結果、どのfixtureの`<workbookView>`も
    `activeTab`/`firstSheet`属性を一切持たない（両方ともXSD規定のデフォルト値0）——
    つまり「位置がずれて壊れる」という具体的なfixture実例が現状ゼロであり、
    実証されていないハザードに備えたgating機構を先回りして作るのは0.10.0が
    一貫して掲げる「hard gate（実fixture確認前にwriterコードを書かない）」
    そのものに反する過剰設計と判断し、この案は採用しなかった。C1と同じopaque
    fragment passthroughとして実装し、`_WORKBOOK_ELEMENTS`/
    `check_workbook_elements()`に`bookViews`を追加（C1と同じ`WORKBOOK_ELEMENT_LOSS`
    カテゴリを共有——`<hyperlinks>`のような別抽出機構が必要な要素ではなく、
    workbookPr/calcPr/extLstと同じ「まるごとコピー」で十分なため）。fixture1〜7
    全てで実装前`WORKBOOK_ELEMENT_LOSS`確認→実装後`CLEAN`確認済み、
    `xr2:uid`属性（root属性のxmlns:xr2宣言に依存）も含め検証済み。
    将来activeTab/firstSheetが実際に非デフォルト値を持つfixtureが追加された場合
    への備えとして、gating機構ではなく`checker`のdocstringに検知方針を明記する
    形にとどめた（`check_workbook_elements()`のdocstring参照）——それ自体が
    silent lossではなく「まだ正しさが未検証」という別種のリスクなので、
    実装するとしても9節の診断の枠組みで扱うべき将来課題として記録。
  - **C3（完了、簡略版）**: `<definedNames>`（print area・print titles含む——
    `_xlnm.Print_Area`等は`<definedNames>`内の特殊named rangeとして表現される、
    fixture5に実例あり）。C2とは異なり`localSheetId`（sheet位置への0-based
    インデックス）は実fixtureに実例が存在する（fixture5の
    `_xlnm.Print_Area localSheetId="0"`）ため、「実証されていないハザードだから
    先送り」というC2の理屈は使えない——ただしfixture5は単一シートのため、
    「シート削除でlocalSheetIdがずれる」という具体的な壊れ方自体はfixtureで
    実証されていない（合成fixtureでのみ再現）。**採用した方針（簡略版、
    per-name remapping ではない）**: `Vm::worksheet_origins`の全キーが
    現在の`Vm::sheet_order`にまだ存在するか（＝ロード後にシート削除が一度も
    起きていないか）だけを見るゲート——1つでもシートが削除されていれば
    `<definedNames>`全体を丸ごと省略し、削除が一度も起きていなければverbatim
    でそのまま持ち越す。個別のdefinedName単位でlocalSheetIdを再計算して
    生存させる（部分的remapping）という、より精密だが実装コストの高い代替案は
    見送った——シート追加のみ（削除なし）ならlocalSheetIdは不変のままなので、
    この簡略版でも「よくある操作（読み込み→編集→保存、シート追加のみ）」では
    100%持ち越され、「シート削除」という比較的まれな操作でのみdefinedNames
    全体を失う、という設計判断。`mechanical_check.py`に専用の
    `check_defined_names()`を追加（`_WORKBOOK_ELEMENTS`の単純な有無チェックとは
    別関数——「削除が起きた場合は省略が正しい」という非対称な判定が必要なため、
    `check_internal_hyperlinks()`が独自関数になったのと同じ理由）。
    自己テストで両方向（削除なし→verbatim必須、削除あり→完全省略必須、
    どちらの逆方向の違反も検知）を確認済み。実fixture（fixture4/5、単一シート
    のためこのテストでは削除分岐を再現できない）と、シート削除分岐専用の
    合成fixture（2シート、`Sheets(...).Delete`マクロ）の両方で実CLI経由の
    動作を確認済み。

  **C1〜C3、実Excel検証済み（ユーザー確認、2026-08-23、Mac Excel）**: fixture4/fixture5を
  コード`571df1e`でsave-as・in-place両経路で保存し（fixture4は両方、fixture5はsave-asのみ
  ——ユーザー判断で十分と確認）、実Excelで再オープン。3出力ファイル全てで修復警告0件。
  fixture4: 「名前の管理」で`test`（参照先`=Sheet1!$F$5`・範囲=ブック・コメント
  `test desu!!!`）を実際に開いて全フィールド一致を確認、編集セルA1=12345も確認。
  fixture5: 「名前の管理」で`_xlnm.Print_Area`（参照先`=Sheet1!$E$3`・範囲=Sheet1）を確認、
  さらに印刷プレビュー（⌘+P）で実際の印刷範囲が空セルE3のみになっている（データ表
  A〜C列が印刷対象に含まれていない）ことを確認——print areaが実効的に機能している
  positive controlとして機能した。**0.10.0-Cはmechanical_check・実Excel検証の両方が
  完了し、正式に完了とする**（ユーザー判定、2026-08-23）。

  **この過程で発見した別件のバグ（0.10.0-Cとは無関係、既存の挙動）**:
  fixture5のD8セル（`t="e"`、`#VALUE!`エラー値）が、elixceeの保存後は`t="s"`
  （共有文字列）の平文テキスト`"#VALUE!"`に変わっていた。`git blame`で
  `src/reader.rs:1292`が`72b5cc38`（2026-06-21）由来と確認——`SheetCell` enumが
  そもそもError variantを持たず（`Integer`/`Float`/`Str`/`Bool`のみ）、`t="e"`を
  `t="str"`と同様に文字列として読んでいるため。今回のC1〜C3のどのコミットも
  触っていない箇所であり、C判定には影響しないと判断（ユーザー承認）。
  `ROADMAP.md`のKnown gaps item 14として別途記録済み、対応は未着手。

  - ~~元の`sheetId`の保持（2節/6節、位置ベース再生成の是正）~~ ——**訂正（stale
    cross-reference）**: 初稿ではここに挙げていたが、実際には0.10.0-Aで
    `WorksheetOrigin`実装の一部として既に完了済み（10節の0.10.0-Aチェックリスト
    参照、commit `ae030b7`）。0.10.0-Cが対象とする残りのworkbook-level要素とは
    別軸（sheetIdはworksheet identityの一部として6節で扱われ、A段階で先に
    片付いた）——ここから削除。
- **0.10.0-D — Relationship-backed features**（relationship graphの再接続そのものが
  主目的、4節の穴を実際に塞ぐmilestone）
  - 前提: 0.10.0-Aの`WorksheetOrigin`と source-reference graph checker
  - external hyperlink（`r:id`形式）・`<tableParts>`・comments/notes・threaded
    comments/persons・`<drawing>`（chart/image）・`<legacyDrawing>`（VML）・
    printerSettings・OLE/controls
  - content-type解決自体は5節の通り既に0.9.0で解決済みなので新規作業ではない——
    このmilestoneの主目的は純粋にrelationship graphの接続性（4節）

  **方針決定（ユーザー承認、着手前の設計。実装はまだ開始していない）**:
  worksheet part命名は**origin-based**（`WorksheetOrigin.original_part_name`を
  維持）を採用する。現行writerが行っている「毎回`sheet{i+1}.xml`へ位置ベースで
  付け直し、worksheet-level `.rels`側をremapする」方式は採用しない。

  理由: Open XML packageはpart発見をrelationshipの連鎖に依存しており、
  worksheet part名が`sheet1.xml`/`sheet2.xml`という連番である必要はない
  （Microsoft Open XML SDKの最小SpreadsheetML例でも`/xl/worksheets/sheet.xml`
  という非連番target）。既存part URIを変更しない方が、未知のpartや将来の
  Excel拡張まで含めた場合に、全参照を書き換えるより安全——lossless
  preservationという0.10.0全体の目的と直接一致する。

  **出力計画（`WorksheetOutputPlan`、設計スケッチ——実コードではない）**:
  保存開始時に一度だけ計画を作り、writerの複数箇所が独立に`i + 1`を
  計算している現状の構造をやめる。

  ```
  struct WorksheetOutputPlan {
      sheet_key: String,           // Vm内部のlowercaseキー
      display_name: String,        // WorksheetOrigin.original_display_name優先
      sheet_id: String,            // WorksheetOrigin.original_sheet_id優先
      workbook_rel_id: String,     // workbook.xml.rels側のrId（position-based、既存通り）
      output_part_name: String,    // 既存シート: original_part_name。新規: 衝突しない新規名
      output_rels_name: String,    // output_part_nameに対応する_rels/*.xml.rels名
      is_existing: bool,           // originがあるか（true）／純粋にVBAで作られたか（false）
  }
  ```

  - **既存シート**: `output_part_name = WorksheetOrigin.original_part_name`をそのまま使う。
    シートの表示順序（`Vm::sheet_order`）を変えても、part名自体は変えない——変更するのは
    `workbook.xml`内`<sheet>`要素の順序と、workbookから各worksheetへのrelationshipのみ。
    worksheet自身の`.rels`（`xl/worksheets/_rels/sheetN.xml.rels`）はpart名が不変な限り
    そのまま維持できる。
  - **新規シート**（`Sheets.Add`）: 既存`sheetN.xml`の最大N + 1を割り当てる。予約済み集合
    として、生存worksheet part名・削除されたsheetが使っていたpart名・passthrough entry名・
    同一save中に新規割当済みのpart名、全てを確認してから決定する——削除された番号を
    即座に再利用すると、古い未知参照（例えば別partに残っているrelationship）が誤って
    新しいシートに結び付く事故になりうるため、単調増加のみを許可する。
  - **削除シート**: `xl/worksheets/<original-part>.xml`と対応する`.rels`は出力しない。
    さらに、削除されたworksheetの`.rels`からのみ到達可能なtarget part（table/drawing/
    chart/image/comments/VML/printerSettings等）を洗い出し、他のsurviving partからも
    参照されているものは残し、削除sheetからしか到達できないものだけ削除する
    （reference counting／package reachability判定。無条件削除は共有image等を
    誤って消す危険があり、無条件保持はorphan partを増やす）。

  **実装分割（D1〜D4、1コミットずつ）**:
  - **D1（完了）**: `WorksheetOutputPlan`の導入と、`workbook.xml`／`workbook.xml.rels`／
    `[Content_Types].xml`をこの計画に基づいて出力するように切り替える。この段階では
    worksheet-level `.rels`の機能復元（relationship-backed要素の実復元）はまだ行わない。
  - **D2（不要と判明——D1で既に満たされていた）**: 生存sheetの元`.rels`をoriginal part名の
    ままrelationship ID不変で引き継ぐ、という目標自体は、D1完了時点のbinaryで
    `check_source_references()`をfixture3に対して実行し確認した——`.rels`ファイル自体が
    差分として検出された違反は0件で、全て「`sheet1.xml`がrIdを参照していない」という
    D3側の欠落だった。generic passthroughは0.9.0からworksheet `.rels`をbyte-identicalで
    運んでおり、D1がpart名の共存置を正しくしたことで、この目標は既に満たされていた。
    別コミットとしての実装は発生しなかった。
  - **D3（tableParts分だけ完了）**: `<tableParts>`・external `<hyperlinks>`・`<drawing>`・
    `<legacyDrawing>`・`<pageSetup r:id>`をtype-awareに復元する（4節の対応表通り）。
    advisorレビューの指摘（B同様、1要素1コミットで進める）に従い、fixtureの実証がある
    `<tableParts>`（fixture3のみ）から着手し、完了。実装は7節(b)のopaque-fragment
    passthrough機構をそのまま再利用——`pageMargins`の直後（8節の実XSD順序で
    `pageMargins`と`tableParts`の間に他要素が未実装のため、現状はこの位置が正しい）へ
    生の`tableParts` XML片を挿入するだけで済んだ。`rels_survived`ゲート（`is_existing`
    かつ該当`.rels`が実際にこのsaveのpassthrough集合に存在するか）を新設し、`.rels`が
    生き残らなかった場合にdangling r:idを絶対に出力しないことをnegative testで確認
    （ゲートを一時的に外してテストが実際に落ちることも確認済み）。`<drawing>`・
    `<legacyDrawing>`・`<hyperlinks>`（B4の既存filter実装の書き換えが必要——最後に回す）・
    `<pageSetup r:id>`（実証fixtureが無ければ着手しない）は未着手。
  - **D4**: sheet rename／reorder／deletion／新規追加／非連番part名／shared・exclusive
    targetのreachability——実fixtureとnegative testで固める。

  **必須テストケース**（D4完了条件の一部）:

  | ケース | 期待結果 |
  |---|---|
  | `sheet5.xml`を持つ1シートを保存 | `sheet5.xml`のまま |
  | `sheet2.xml`／`sheet7.xml`を持つ2シートを並べ替え | part名は不変、表示順だけ変更 |
  | シートrename（このVMには未実装だが将来に備え） | part名と`.rels`名は不変 |
  | 新規シート追加 | 未使用の新part名 |
  | relationship付きシート削除 | worksheet・`.rels`・exclusive targetが消える |
  | shared targetを持つシート削除 | shared targetは残る |
  | table／drawing／external hyperlink | checkerがCLEAN |
  | source参照だけ削除 | `SOURCE_REFERENCE_LOSS` |
  | `.rels`だけ削除 | `DANGLING_RELATIONSHIP`または対応分類 |
  | target partだけ削除 | `DANGLING_RELATIONSHIP` |

  **着手条件（満たされた、2026-08-23）**: 0.10.0-Cの実Excel確認が完了するまでD1を
  含め一切のコード実装に着手しない、という条件を設けていた——B/C/Dを混ぜると実Excelで
  問題が起きた際に原因の切り分けができなくなるため（ユーザー指示）。C1〜C3の実Excel
  検証がユーザー確認済み（上記C3末尾参照）となったため、D1着手可能。

  **D1実装済み（2026-08-23、ユーザーへの実Excel確認依頼はまだ行っていない）**:
  `WorksheetOutputPlan`と`plan_worksheet_output`を実装し、`build_xlsx_content_types`／
  `build_xlsx_workbook`／`build_xlsx_workbook_rels`／per-sheet書き込みループの全てを
  この計画一本に統一した。D1で許可されたスコープ（出力計画とテストのみ）を厳密に守り、
  relationship復元（D2/D3の担当）には一切着手していない。

  スタブや仮説ではなく実バグ修正であることを確認済み: 修正前は生存sheetの内容が
  位置ベースのpart名に書かれる一方、そのsheet自身のpassthrough `.rels`は元のpart名の
  ままだったため、より前のsheetが削除されて位置がずれると、内容と`.rels`が食い違い、
  `.rels`が孤立していた。`git stash`／`git stash pop`で修正前後のコードを同一シナリオ
  （3シート合成fixture、Sheet3が実hyperlink `.rels`を持つ、Sheet2を削除）に対して
  実行して直接比較し、修正前は本当にこの不整合が起きることを確認した。

  さらに2つの独立した検証を追加で実施した（advisor指摘を受けて）:
  1. **relationship持ちシート自体を削除するCLIシナリオ**（先の`git stash`比較は
     relationship無しのSheet2を削除していたため、これとは異なる形）——Sheet3
     （`_rels/sheet3.xml.rels`を持つ）を削除して`mechanical_check.py`のフル
     チェックを通したところ、`structural_verdict: STRUCTURALLY_CLEAN`が返った。
     しかしこれは「本当に問題がない」ことを意味しない——`_rels/sheet3.xml.rels`は
     対応する`sheet3.xml`が無いまま出力に生き残る（orphan）が、`check_roundtrip`の
     どのチェックもこの形の孤立を検出できないことをソース読解で確認した（2b節は
     「partが参照されているか」しか見ず、「`.rels`ファイル自身がOPC命名規則上
     結び付くはずのpartが実在するか」は見ない）。実Excelには無害（何も参照しない
     ためExcelが開く理由がない）で、D1で新規発生したものではない（D1以前も別の
     stale part名で同じ形の孤立が起きていた）が、checker側の既知の穴として
     ROADMAP.md Known gaps項目15に記録した。D4のreachability清掃で実体は消える
     見込みだが、checker専用のnegative testはまだ書いていない。
  2. **`sheet1.xml`／`sheet3.xml`（欠番2）を持つ合成fixtureに対する`Sheets.Add`の
     end-to-endシナリオ**——新規part割当分岐（`next_fresh_part_n`相当）が
     unit testでしか検証されていなかったため、実際にCLIを通して確認した。
     この過程で`Stmt::SheetsAdd`自体の独立したバグ（`self.sheets.len() + 1`のみで
     衝突チェックが無く、番号に欠番がある状態で`Add`すると既存sheetと衝突して
     `ensure_sheet`が黙って何もしないバグ——`72b5cc38`由来、0.10.0とは無関係）を
     発見し、別コミットで先に修正した。修正後に同じCLIシナリオを再実行し、
     新規sheetが`sheet4.xml`（`max(1,3)+1`、欠番2の再利用なし）に正しく割り当てられ、
     `[Content_Types].xml`のOverrideと`workbook.xml.rels`のTargetが両方とも同じ
     part名を指していることを確認した。

  **未解決事項（D4への申し送り）**: 上記「必須テストケース」表の「`sheet2.xml`／
  `sheet7.xml`を持つ2シートを並べ替え」は、このVMに現状シートの並べ替え・rename
  primitiveが存在しない（`sheet_order`はsource順のまま、`Sheets.Add`/`Delete`しか
  無い）ため、現時点では到達不能なテストケースである。D4着手時に並べ替えprimitiveを
  実装するか、このテストケースを別の到達可能な形に置き換えるか判断が必要。

各milestoneの完了条件は0.9.0-Aと同じ形式（実Excel再オープンでrepair警告0件、
mechanical_check clean、既存回帰テスト無変化）を踏襲することを推奨する。

## 未解決のまま残す論点（この文書では決めない）

- ~~freeze pane・internal hyperlink（`location`属性）の実データ確保（3節）~~ ——
  0.10.0-A中に解消。ユーザー提供の実Excelファイルから`fixture6_internal_hyperlink.xlsm`・
  `fixture7_freeze_pane.xlsm`を追加、`INVENTORY.md`更新済み（§10）。0.10.0-Bのwriter実装
  着手時は、これらのfixtureに対する具体的なXML構造・XSD確認・negative self-test追加が
  依然として必要な作業として残る。
- `<conditionalFormatting>`の`dxfId`/`extLst`参照の扱い（10節、0.10.0-B）——raw subtree
  preservationで一旦済ませるか、参照検証まで踏み込むかは実装時に判断。
- comments/threadedComments/richData等、単純なr:id対応表に収まらない機能の個別確認
  （9節）——実fixtureとXSDを見ながら実装時に確定する。
- richDataを"検知はするが触れない"扱いにする際の、9節診断の具体的な文言——実装時に決定。
