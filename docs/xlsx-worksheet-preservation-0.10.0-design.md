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
  - **B3以降（未着手）**: internal hyperlinkの`location`属性（advisor指摘:
    `<hyperlinks>`はB1/B2と同じまるごとpassthroughができない——`r:id`を持つ
    子要素とlocation-onlyの子要素が混在しうるため、子要素単位でのfilteringが
    必要。かつ`CT_Hyperlinks`の`<hyperlink>`の`minOccurs`を実XSDで確認し、
    「全部r:id形式だった場合は`<hyperlinks>`ごと省略する」のか「空の
    `<hyperlinks/>`を出力してよいのか」を実装前に確定する必要がある）・
    `<autoFilter>`（実fixtureにstandalone要素としての実例が現状なし、
    fixture新規待ち）・`<pageMargins>`・行/列プロパティ（`<cols>`/`<row>`の
    幅・スタイル等、hidden以外の属性）。
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
  1節の指摘通りスコープを混ぜない）
  - `<definedNames>`（6節のシートリネーム時dangling注意、9節の診断で対応）
  - print area・print titles（`<definedNames>`内の特殊named range、`_xlnm.Print_Area`等）
  - workbook metadata（`<bookViews>`・`<calcPr>`・`<workbookPr>`）
  - 元の`sheetId`の保持（2節/6節、位置ベース再生成の是正）
  - workbook-levelの`<extLst>`拡張ノード
- **0.10.0-D — Relationship-backed features**（relationship graphの再接続そのものが
  主目的、4節の穴を実際に塞ぐmilestone）
  - 前提: 0.10.0-Aの`WorksheetOrigin`と source-reference graph checker
  - external hyperlink（`r:id`形式）・`<tableParts>`・comments/notes・threaded
    comments/persons・`<drawing>`（chart/image）・`<legacyDrawing>`（VML）・
    printerSettings・OLE/controls
  - content-type解決自体は5節の通り既に0.9.0で解決済みなので新規作業ではない——
    このmilestoneの主目的は純粋にrelationship graphの接続性（4節）

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
