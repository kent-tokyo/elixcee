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
    // rename後も変わらない、VM内部でのみ使う識別子。elixcee側で新規発行してもよいし、
    // 上記3つのどれかが安定して使えるならそれをそのまま採用してもよい——ロード時に
    // 一度だけ決定し、以後そのシートの寿命が尽きるまで変えない、という制約が本質。
    stable_key: SheetKey,
}
```

役割を明確に分ける: シート名（ユーザーに見える可変値）／`sheetId`（workbook内の識別子、
`.ods`など`sheetId`を持たないソースではNone）／workbook.xmlの`r:id`（workbook.xmlから
worksheet partへの関係）／worksheet part path（`xl/worksheets/sheetN.xml`という文字列、
現状は位置から再生成されているため最も不安定）／VM内部identity（renameを跨いで安定させる
必要がある唯一の値）——**5つは別々の軸であり、どれか1つだけで代用しようとしない。**

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

**重要な注記（provenance）**: 以下のうち`fixture3`の実データで**実際に位置関係を確認した**の
は`sheetPr, dimension, sheetViews, sheetFormatPr, [cols,] sheetData, [mergeCells,] phoneticPr,
conditionalFormatting, dataValidations, pageMargins, tableParts`の10〜12要素分の並びのみ
（`cols`/`mergeCells`は現行writerが既に正しい位置に出力できていることを既存コードから
確認、fixture3自体には含まれていなかった）。残りの位置は**ECMA-376/ISO 29500の
`CT_Worksheet`スキーマ定義をこちらの記憶から再構成したもので、実XSDと突き合わせて
検証していない**。0.10.0-Aの実装着手時には、この表を鵜呑みにせず、
`sml.xsd`（`CT_Worksheet`の`xsd:sequence`定義）を一次情報として直接参照し直すこと
——万一この表の並びを1箇所でも間違えたまま実装すると、real Excelがファイルを
repair対象にする、というこのプロジェクトが実際に0.9.0で繰り返し踏んだ失敗パターンを
再現する。

```
sheetPr, dimension, sheetViews, sheetFormatPr, cols, sheetData, sheetCalcPr,
sheetProtection, protectedRanges, scenarios, autoFilter, sortState,
dataConsolidate, customSheetViews, mergeCells, phoneticPr,
conditionalFormatting, dataValidations, hyperlinks, printOptions,
pageMargins, pageSetup, headerFooter, rowBreaks, colBreaks,
customProperties, cellWatches, ignoredErrors, smartTags, drawing,
legacyDrawing, legacyDrawingHF, picture, oleObjects, controls,
webPublishItems, tableParts, extLst
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

| relationship type | worksheet側の参照元 |
|---|---|
| table | `<tableParts><tablePart r:id="..">` |
| drawing | `<drawing r:id="..">` |
| hyperlink（external） | `<hyperlinks><hyperlink r:id="..">` |
| vmlDrawing | `<legacyDrawing r:id="..">` |
| printerSettings | `<pageSetup r:id="..">`（`xl/printerSettings/printerSettingsN.bin`を指す） |
| oleObject | `<oleObjects><oleObject r:id="..">` |

comments・threadedComments・拡張機能系（richData等）は関係の張り方がworksheet直下の
単純な`r:id`属性1つに収まらない場合がある（`fixture4`では`legacyDrawing`経由の間接参照と
`comments1.xml`への直接relationshipが併存している）ため、この対応表に機械的に含めず、
**実fixtureとXSDを個別に確認してから**表に追加することを推奨する——分かっている範囲を
正しく書く方が、分かっていない範囲まで表に含めて誤った安心感を与えるより安全。

**新しい違反分類を追加する**（既存の`ORPHANED_PART`/`DANGLING_RELATIONSHIP`とは別カテゴリ
として区別する——今回見つかったのはどちらでもなく、3つ目の新しい種類の壊れ方）:

- `SOURCE_REFERENCE_LOSS` — `.rels`と対象partはどちらも存在するが、regenerateされた
  writer-owned part（`sheetN.xml`等）のどこからもそのrIdが参照されていない（今回の発見）。
- `DANGLING_RELATIONSHIP` — 既存（`.rels`が指すtargetが存在しない）。
- `ORPHANED_PART` — 既存（partが存在するがどの`.rels`からも参照されていない、Milestone 4で
  発見済み）。

**self-testに追加すべき破壊ケース**（`mechanical_check.py --self-test`が確実にこれらを
検出できることを、新検査を信用する前に必ず確認する——0.9.0-Aの教訓「検出できない
checkerはゼロ違反を返すだけで正しさの証拠にならない」を踏襲）: 元の`.rels`と対象partを
そのまま残した状態で、worksheet側の参照要素だけを個別に取り除いたケースを7種
（`<tableParts>`削除・`<drawing>`削除・external hyperlinkの`r:id`属性削除・
`<legacyDrawing>`削除・`<pageSetup r:id>`のid属性のみ削除、を最低ラインとし、上記表が
確定した時点で追加する）用意し、いずれも`SOURCE_REFERENCE_LOSS`として検出されることを
確認してから、実fixtureへの適用を信用する。

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
  - `WorksheetOrigin`のidentity設計（6節）の実装
  - CT_Worksheet/CT_WorkbookのXML/XSD順序の確定（8節、記憶ベースの表を実XSDで裏取り）
  - source-reference graph checker（type-aware mapping、9節）
  - `mechanical_check.py`のnegative self-test追加（9節、7種の破壊ケース）
  - fixture inventory の棚卸し（既存5 fixtureが実際に何を含み何を含まないかの一覧化、
    3節で発覚した「fixture5に実はfreeze paneがない」のような食い違いを潰す）
  - internal hyperlink（`location`属性）のfixture新規作成——現状ゼロ（3節）
  - freeze paneのfixture新規作成——現状ゼロ（3節）
  - relationship type → worksheet側source element 対応表の確定（9節）
  - 検証: 既存5 fixtureの構造/mechanical_check再確認、新規機能なし（回帰ゲート）
- **0.10.0-B — Inline worksheet preservation**（relationship非依存、7節(b)の
  opaque-fragment機構をworksheet-XML側へ適用）
  - `<sheetViews>`（`<selection>`・`<pane>`＝freeze pane含む）・`<sheetPr>`・
    `<sheetFormatPr>`・`<dataValidations>`・internal hyperlinkの`location`属性・
    `<autoFilter>`・`<pageMargins>`・行/列プロパティ（`<cols>`/`<row>`の幅・スタイル等、
    hidden以外の属性）
  - **`<conditionalFormatting>`は要注意——`dxfId`（`xl/styles.xml`の`<dxfs>`）や
    `<extLst>`拡張を参照する場合があり、完全な参照非依存とは言い切れない。** 最初は
    「参照先の妥当性検証はせず、生のsubtreeをそのまま保存する」raw subtree preservation
    として扱い、`dxf`参照の検証・復元は将来の別スライスへ回すことを推奨する。
  - 0.10.0-Aの機構へスロットを追加するのみ
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

- freeze pane・internal hyperlink（`location`属性）の実データ確保（3節）——どちらも
  現状このリポジトリに実例がなく、新規fixtureが0.10.0-A着手の前提作業になる。
- `<conditionalFormatting>`の`dxfId`/`extLst`参照の扱い（10節、0.10.0-B）——raw subtree
  preservationで一旦済ませるか、参照検証まで踏み込むかは実装時に判断。
- comments/threadedComments/richData等、単純なr:id対応表に収まらない機能の個別確認
  （9節）——実fixtureとXSDを見ながら実装時に確定する。
- richDataを"検知はするが触れない"扱いにする際の、9節診断の具体的な文言——実装時に決定。
