# elixcee

[English](README.md) | **日本語** | [中文](README_zh.md)

Microsoft Excelをインストールせずに、Excel VBAマクロを実行・テスト・原因診断できるRust製ヘッドレスランタイムです。存在しないシート、配列範囲外、コピー／貼り付け範囲の不一致、保護シートへの書き込みなど、Excelが操作を拒否する理由を具体的な範囲と根拠付きで報告します。

コアエンジンは **Rust**、Python バインディングは **pyo3 + maturin** で提供します。

## 名前の由来

**elixcee** = **Excel** + **elixir**（エリクサー） + **C**

Excel 依存という「呪い」を解く*エリクサー*（万能薬）— Rust による C レベルの速度で動作します。

---

## 類似ツールとの比較

| 機能 | **elixcee** | xlwings | LibreOffice UNO | openpyxl | xlcalculator |
|------|:-----------:|:-------:|:---------------:|:--------:|:------------:|
| VBA マクロの実行 | あり | あり | あり（一部） | なし | なし |
| Excel が必要 | なし | あり | なし | なし | なし |
| LibreOffice が必要 | なし | なし | あり | なし | なし |
| 数式の評価 | あり | あり | あり | なし | あり |
| macOS / Linux / Windows | あり | 一部 | あり | あり | あり |
| シンプルな Python API | あり | あり | なし | あり | あり |
| .xlsx の読み込み | あり | あり | あり | あり | あり |
| .ods の読み込み | あり | あり | あり | なし | なし |
| .xlsx の書き込み | あり | あり | あり | あり | なし |
| .ods の書き込み | あり | あり | あり | なし | なし |
| 実行速度 | Rust（ネイティブ） | COM/IPC（低速） | IPC（低速） | — | Python |

**補足:**
- **xlwings** は macOS では AppleScript 経由で Excel for Mac が、Windows では COM 経由で Excel が必要です。Linux サポートには Excel インスタンスまたはクラウドブリッジが必要です。
- **LibreOffice UNO** は起動に 1 秒以上かかる場合があり、API も複雑です。VBA は LibreOffice 独自のインタプリタで実行されるため、Excel の動作と完全には一致しない場合があります。
- **openpyxl** は .xlsx ファイルからキャッシュ済みの数式値を読み込みますが、実行時に数式を再評価する機能はありません。
- **xlcalculator** は Excel の数式を Python で再評価できますが、VBA はサポートしていません。
- elixcee の VBA インタプリタは、一般的なデータ処理マクロで使われる VBA のサブセット（ループ、条件分岐、セルの読み書き、文字列・数学関数、複数シートへのアクセス）をカバーしています。グラフ作成や書式設定など、Excel の UI 操作の大半は未対応または no-op です。`MsgBox` だけは特別扱いで、モードに応じて標準出力への表示・JSON 出力への収集・エラーとしての送出のいずれかになります。

---

## インストール

```bash
pip install elixcee
```

開発版（ソースからビルド）:

```bash
python3 -m venv .venv && source .venv/bin/activate
maturin develop
```

---

## CLI（Windows / Linux / macOS）

Python 不要のスタンドアロンバイナリを [Releases](https://github.com/kent-tokyo/elixcee/releases) ページで配布しています。

| ダウンロード | 対象プラットフォーム |
|---|---|
| [elixcee-x86_64-windows.exe](https://github.com/kent-tokyo/elixcee/releases/latest/download/elixcee-x86_64-windows.exe) | Windows x64 |
| [elixcee-x86_64-linux](https://github.com/kent-tokyo/elixcee/releases/latest/download/elixcee-x86_64-linux) | Linux x64 |
| [elixcee-aarch64-macos](https://github.com/kent-tokyo/elixcee/releases/latest/download/elixcee-aarch64-macos) | macOS Apple Silicon |

### 使い方

```
elixcee <vba_file>... <MacroName> [OPTIONS]

引数:
  <vba_file>...  VBA ソースファイルのパス（.vbs / .bas / .txt）を1つ以上。
                 複数ファイルの場合、モジュールをまたいで同名の Sub/Function
                 があれば Module.Sub で区別する。
  <MacroName>    実行する Sub の名前（最後の引数）

オプション:
  --file <path>    スプレッドシートからセルデータを読み込む（.xlsx / .xlsm / .ods）
  --sheet <name>   アクティブシート名（デフォルト: --file の先頭シート）
  --output <path>  結果セルをスプレッドシートに保存（.xlsx / .ods）
  --json           プレーンテキストの代わりに単一の JSON オブジェクト（結果またはエラー）を出力
```

### 実行例

VBA ファイルを実行して結果を標準出力に表示:

```bat
elixcee macro.vbs ProcessData
```

Excel ファイルからデータを読み込み、マクロを実行し、結果を保存:

```bat
elixcee macro.vbs ProcessData --file input.xlsx --output result.xlsx
```

出力形式 — 非空セルを1行ずつ、アドレスと値をタブ区切りで表示:

```
A1    Hello
B1    42
A2    3.14
```

`MsgBox` の内容は標準出力に表示されます。

### 複数ファイル（マルチモジュールプロジェクト）

複数の VBA ファイルを渡すと、複数モジュールにまたがるプロジェクトを実行できます。Sub/Function 名はプロジェクト全体で共有されるため、同じ名前が複数モジュールに存在する場合は `Module.Sub` で特定のものを指定します（モジュール名は `Attribute VB_Name` があればその値、なければファイル名から決まります）:

```bat
elixcee Helpers.bas Main.bas Main.ProcessData
```

プロジェクトマニフェストはまだありません（対応範囲の詳細、モジュール間で名前が衝突した場合の扱いなどは [docs/agent-contract.md](docs/agent-contract.md) を参照）。

### JSON 出力（スクリプト・AI エージェント向け）

`--json` を付けると、プレーンテキストの代わりに単一の機械可読な JSON オブジェクトを出力します:

```bat
elixcee macro.vbs ProcessData --json
```

```json
{"schema_version":1,"ok":true,"entrypoint":"ProcessData","duration_ms":0.42,"cells":[{"sheet":"sheet1","address":"A1","value":42}],"messages":[]}
```

契約の全容（エラーコード・終了コード・`messages` の仕様）: [docs/agent-contract.md](docs/agent-contract.md)

### マクロを実行せずに静的解析する

`elixcee check` は1つ以上の `.bas` ファイルを**実行せずに**検査します: parse エラー、指定した macro の存在確認、本文中の未定義 Sub/Function 呼び出し、`MsgBox` などの対話操作の検出。位置引数はすべてファイルとして扱われ、エントリポイント（指定する場合）は常に `--entry` で渡します（位置引数では渡しません）— そのため `elixcee check *.bas` は特定のエントリポイントを前提とせずプロジェクト内の全モジュールを検査できます。

```bat
elixcee check macro.vbs --entry ProcessData --json
```

```json
{"schema_version":1,"ok":true,"diagnostics":[]}
```

### ワークブックのスナップショット

`elixcee snapshot` は `.xlsx`/`.xlsm`/`.ods` ファイルを VBA を実行せずに直接読み込み、全シートの非空セルを Markdown（デフォルト）または `--json` で JSON として出力します。

```bat
elixcee snapshot Book1.xlsx --json
```

```json
{"schema_version":1,"ok":true,"file":"Book1.xlsx","sheets":[{"name":"Sheet1","sheet_id":"1","stable_id":"sheet1","cells":[{"address":"A1","value":42}]}]}
```

`stable_id` はファイル自身が持つ `sheetId`（無ければ位置ベースのフォールバック）から導出したものであり、VBA の `CodeName` プロパティ**ではありません**。詳細な設計理由は [docs/agent-contract.md](docs/agent-contract.md) を参照してください。

### プロパティベースのワークブックテスト

`elixcee test-workbook` は、生成された境界値入力（空欄・`0`・`1`・`-1`・オーバーフロー付近の数値・空/短/長い文字列）を使ってマクロを何度も実行し、panic・ランタイムエラー・タイムアウト・Excel エラー値の混入を検査します。各ケースは必ずまっさらなワークブック状態から開始します。

```toml
# fixture.toml
name = "order calculation"
workbook = "orders.xlsx"
vba_files = ["Main.bas"]
macro = "Main.Process"
cases = 100
seed = 42

[[inputs]]
range = "Input!B2:B10"
strategy = "boundary_numeric"

[[assertions]]
range = "Result!A1:F100"
rule = "no_excel_errors"
```

```bat
elixcee test-workbook fixture.toml --json
```

失敗したケースは seed と case index を報告するため、`elixcee test-workbook fixture.toml --seed 42 --case 17` で正確に再現できます。スキーマ・strategy・assertion ルールの詳細は [docs/agent-contract.md](docs/agent-contract.md) を参照してください。

### Excel操作の診断

`elixcee diagnose` はマクロを一度だけ実行し、存在しないシート・存在しないワークブック・配列の範囲外アクセス・Copy/Paste の形状不一致・保護されたシートへの書き込み・結合セルのレイアウトと衝突する Copy/Paste など、Excelがその操作を拒否する具体的な理由を根拠付きで説明します（単なるエラー文字列ではありません）：

```bat
elixcee diagnose Main.bas --file report.xlsx --json Main.Run
```

```json
{
  "schema_version": 1,
  "ok": false,
  "message": "Sheet 'Sales2025' not found",
  "location": {"file": "Main.bas", "line": 2, "column": 5},
  "root_causes": [
    {
      "code": "WORKSHEET_NOT_FOUND",
      "certainty": "definite",
      "expression": "Worksheets(\"Sales2025\")",
      "requested": "Sales2025",
      "available": ["input", "sales2026", "summary"],
      "suggested": "sales2026",
      "suggestions": ["did you mean 'sales2026'?"]
    }
  ],
  "messages": []
}
```

`Range("A1:C10").Copy` の後に `Range("E1:F10").PasteSpecial` を実行すると、形状の不一致と両方の文の位置を報告します:

```json
{
  "code": "PASTE_SHAPE_MISMATCH",
  "source_addr": "A1:C10", "source_rows": 10, "source_cols": 3,
  "dest_addr": "E1:F10", "dest_rows": 10, "dest_cols": 2,
  "copy_location": {"file": "Main.bas", "line": 2, "column": 5},
  "suggestions": [
    "resize the destination to E1:G10",
    "or specify only the top-left cell E1"
  ]
}
```

`.Protect` されたシートへ書き込むと、どのシートかと修正方法を報告します:

```json
{
  "code": "SHEET_PROTECTED",
  "sheet": "sheet1",
  "suggestions": ["unprotect the sheet first: Worksheets(\"sheet1\").Unprotect"]
}
```

`A1:C10` を `E1:G10` へ貼り付ける際、貼り付け先の1行目だけが結合されている（`E1:G1`）が貼り付け元は結合されていない場合、レイアウトの衝突と両方の文の位置を報告します:

```json
{
  "code": "PASTE_MERGE_LAYOUT_MISMATCH",
  "source_addr": "A1:C10", "dest_addr": "E1:G10",
  "conflicts": ["E1:G1"],
  "copy_location": {"file": "Main.bas", "line": 2, "column": 5},
  "suggestions": [
    "unmerge E1:G1 before pasting",
    "or make the source and destination merge layouts identical"
  ]
}
```

分類ルールと JSON スキーマの詳細は [docs/agent-contract.md](docs/agent-contract.md) を参照してください。

### 生成された入力群にわたる診断

`elixcee diagnose-workbook` は上記2つの機能を組み合わせたものです。`test-workbook` が生成するケース群に対してマクロを繰り返し実行し、失敗した場合は単なるエラー文字列ではなく分類結果を報告します。配列範囲外エラーのように、一部の入力値でしか再現しない入力依存の失敗にこそ真価を発揮します——形状不一致・結合セルの衝突・シート保護といった構造的な問題は、そもそも入力値に依存しないため `diagnose` を1回実行するだけで十分見つかります:

```bat
elixcee diagnose-workbook fixture.toml --json
```

```json
{
  "schema_version": 1,
  "ok": false,
  "seed": 42,
  "case_index": 3,
  "inputs": [{"address": "sheet1!B2", "value": 999999999}],
  "failure": {
    "rule": "no_runtime_error",
    "message": "Array 'arr': index 999999999 out of bounds (len=6)"
  },
  "root_causes": [
    {
      "code": "ARRAY_INDEX_OUT_OF_BOUNDS",
      "name": "arr", "index": 999999999, "lower": 0, "upper": 5,
      "suggestions": ["check that 'arr' is large enough for index 999999999 (valid range is 0 To 5)"]
    }
  ]
}
```

フィクスチャ形式と `--seed`/`--case` による再現は `test-workbook` と同一、加えて今回の実行だけフィクスチャのケース数を上書きする `--cases N` を追加しました。完全なスキーマは [docs/agent-contract.md](docs/agent-contract.md) を参照してください。

### 複数領域Range

`Range("A1:A10,C1:C10")` のような非連続な複数領域Rangeを `.Copy` が認識するようになりました。ただし貼り付けは診断専用です——`diagnose`/`diagnose-workbook` が、黙って何もしない代わりに理由を分類して報告します:

```json
{
  "code": "MULTI_AREA_TO_SINGLE_AREA_PASTE",
  "source_areas": [
    {"address": "A1:A10", "rows": 10, "columns": 1},
    {"address": "C1:C10", "rows": 10, "columns": 1}
  ],
  "destination_areas": [
    {"address": "E1:F10", "rows": 10, "columns": 2}
  ],
  "suggestions": [
    "paste each source area separately",
    "copy a contiguous rectangular range",
    "use destination areas with matching count and shapes"
  ]
}
```

`Union()`、`Areas`、`Dim rng As Range`/`Set` によるオブジェクト変数、および領域数・形状が完全一致する場合の複数領域貼り付けは、現在対応済みです——詳細は下記「VBAオブジェクトモデル」を参照してください。上記の4つの分類コードは、領域数や形状が一致しない、または片側のみ複数領域といった、完全一致しないすべてのケースに引き続き適用されます。全体像は [docs/agent-contract.md](docs/agent-contract.md) を参照してください。

### 非表示行・列の証拠情報

`diagnose`/`diagnose-workbook` は、`.Copy` したRangeが非表示の行・列（実際のXLSXの `hidden="1"` メタデータから読み取り）と重なっている場合に報告するようになりました——これはエラーではなく、`root_causes` とは別の新しい `observations` フィールドとして（併記、または単独で）出力されます:

```json
{
  "code": "RANGE_CONTAINS_HIDDEN_CELLS",
  "certainty": "observed",
  "range": {"sheet": "sheet1", "address": "A1:C100", "rows": 100, "columns": 3},
  "visibility": {
    "hidden_rows": ["11:14", "30:39"],
    "hidden_columns": ["B:B"],
    "total_cells": 300,
    "visible_cells": 172
  },
  "message": "The range contains hidden rows or columns. Excel operations using visible cells only may produce a multi-area range."
}
```

これは下記の `SpecialCells(xlCellTypeVisible)` が土台にしている情報です——通常のCopy/Paste自体の挙動は変わりません（非表示セルもこれまで通りコピー・貼り付けされます）。対応はXLSXのみで、ODSは後続課題です。全体像は [docs/agent-contract.md](docs/agent-contract.md) を参照してください。

### VBAオブジェクトモデル

```vb
Dim rng As Range
Set rng = Range("A1:B2")
rng.Value = 5                        ' 実際のSet参照セマンティクス——コピーではなくエイリアス

Dim u As Range
Set u = Union(Range("A1"), Range("D1"))
Range("C1").Value = u.Areas.Count    ' 2

Dim ws As Worksheet
Set ws = ActiveSheet
ws.Range("A1").Value = 1

Range("F1").Value = 7 Mod 3          ' 1
Range("F2").Value = 2 ^ 3            ' 8
Range("F3").Value = 7 \ 3            ' 2（整数除算）
If Not (a And b) Then MsgBox "ok"

With Cells(r, c)                     ' 任意の対象式を、一度だけ評価
  .Value = 5
  If .Value > 0 Then .Value = .Value + 1   ' .memberはどの深さのネストでも解決される
End With

Set rng = Range("A1"): Set rng2 = rng: Set rng = Nothing
rng2.Value = 1                       ' rngがNothingになってもエイリアスは生き続ける
rng.Value = 2                        ' "Object variable or With block variable not set"

Dim n
n = Null
If IsNull(n + 5) Then MsgBox "Nullは+を伝播する"   ' True

Function DoubleIt(x As Integer) As Integer
  DoubleIt = x * 2
End Function
```

`Set`で代入された`Range`/`Worksheet`/`Workbook`オブジェクト変数——実際の参照セマンティクスに加え、真のunset/`Nothing`状態も持つ（一度も`Set`されていない、または明示的に`Nothing`にされた変数へのメンバーアクセスは実際のVBAの「オブジェクト変数、または With ブロック変数が設定されていません」エラーを送出し、`Set x = Nothing`は`x`自身のみを解除し、以前に作成したエイリアスには影響しない）——`Union`/`Areas`、`SpecialCells(xlCellTypeVisible)`（上記の非表示行・列情報を利用）、領域数・形状が一致する複数領域Copy/Paste、`Mod`/`\`/`^`、式中の`And`/`Or`/`Xor`/`Not`（非Boolean値に対する実際のビット演算）、ランタイムの`With`スタック（`With Cells(r, c)`のような計算対象式も含め一度だけ評価され、`.member`は`If`/`For`/`Do`/`Select Case`内のどの深さのネストでも正しく解決される）、`Variant`の`Null`（`Empty`とは異なる、`+`/`&`/比較演算子を通じた文書化されたVBAの伝播規則）、`:`による複数文区切り、型付き`Function`の引数・戻り値、1つの`Dim`文で複数の変数を宣言する構文（`Dim a As Integer, b As Range`）、単一行`If cond Then stmt [Else stmt]`——すべて対応済みです。

**既知の制限**: 複数領域の貼り付けは、両側が複数領域で`Areas.Count`と各領域の形状が一致する場合のみ実行され、それ以外の組み合わせは引き続き診断のみです（上記参照）。

### XLSX.read()/write() — `@elixcee/xlsx`（npm、公開準備済み・未公開）

同期的でWebAssembly版の`XLSX.read(bytes)`——`await init()`不要——がnpmパッケージ`@elixcee/xlsx`（互換性の取り組みと同期ブリッジの設計は [docs/xlsx-architecture.md](docs/xlsx-architecture.md) を参照）に実装されており、`readFile()`/`readFileSync()`（Node専用。ブラウザ向けエントリポイントは偽のファイルシステムを装う代わりに例外を送出）も加わりました。シート名、`!ref`、`!merges`、`!rows`/`!cols`（非表示行・列）、セルごとの`{t, v, f, w, z}`（値・数式テキスト・書式済み表示文字列・日付型セル、実際の`styles.xml`/数値書式解析による）を返します。実物の`xlsx@0.18.5`パッケージに対する差分テストで、33/33 MATCH・開示ゼロ（以前の版で記載していた`src/reader.rs`の`xml:space="preserve"`トリム欠陥は修正済み。詳細はCHANGELOG.md参照）。Node（CJS/ESM）とブラウザの両方で動作します（`"browser"` export conditionが、インライン化されたバイト列と`initSync`によるWASM artifactへ配線済み)——この配線は、Nodeが当該export conditionをシミュレートするだけでなく、実際のヘッドレスChromeプロセスが実物のbundleを読み込み、ページ自身のDOMから`XLSX.read()`の結果を取得することでも検証済みです（Safariは非対応・未検証）。ブラウザ向けエントリポイントは依然としてバンドル利用を前提としています（共有コードにCJSの`require('ssf')`が含まれるため）——ビルド不要の`<script type="module">`でそのまま使える形ではありません——が、実物のnpm tarballインストール（このリポジトリへの相対importではなく）とCJS/ESMバンドルはいずれも、手動でのアセットコピー不要でそのまま動作するようになりました。

`XLSX.write(wb, opts)`/`writeFile()`/`writeFileSync()`——純粋なJS/XML/ZIP生成のみでRust側の書き込み実装は不要——も実装済みです（`bookType: "xlsx"`のみ）。実物のoracleに対して双方向で差分テスト済み：36 MATCH + 開示済み1ケース（`bookType: "ods"`は未実装）。`package.json`のdescriptionは更新済みですが、`version`（`0.0.0-development`のまま）・`private`（`true`のまま）・`publishConfig`（未設定のまま）は意図的に無変更です——**`npm publish`はまだ一度も実行していません**。この環境からは`@elixcee` npm scopeの所有権も確認できません（詳細はROADMAP.mdの「Known gaps」参照）。

### ソースからビルド

```bash
cargo build --release --bin elixcee
# 生成物: target/release/elixcee（Windows では elixcee.exe）
```

---

## クイックスタート

```python
import elixcee

# VBA マクロを実行し、結果セルをすべて取得
cells = elixcee.run_macro("""
Sub FillSquares()
    For i = 1 To 5
        Cells(i, 1).Value = i * i
    Next i
End Sub
""", "FillSquares")
# cells == {(1,1): 1, (2,1): 4, (3,1): 9, (4,1): 16, (5,1): 25}

# Python からセルを事前設定してマクロを実行
vm = elixcee.Vm()
vm.set_cell(1, 1, 100)
vm.set_cell(2, 1, 200)
vm.run("""
Sub CalcTotal()
    total = Cells(1,1).Value + Cells(2,1).Value
    Cells(3,1).Value = total
End Sub
""", "CalcTotal")
print(vm.get_cell(3, 1))   # 300
print(vm.variables())       # {"total": 300}

# Excel ファイルのセルデータを読み込んでマクロを実行
vm = elixcee.load_workbook("data.xlsx")
vm.run(vba_code, "ProcessData")
result_cells = vm.cells()   # {(row, col): value, ...}

# セルにワークシート数式を設定して評価
vm.set_cell_formula(4, 1, "=SUM(A1:A3)")
print(vm.get_cell(4, 1))   # A列1〜3行の合計

# MsgBox の動作を制御
vm = elixcee.Vm(on_msgbox="skip")   # MsgBox を無視（デフォルト）
vm = elixcee.Vm(on_msgbox="error")  # MsgBox 時に RuntimeError を発生
```

---

## Python API

| メソッド | 説明 |
|---|---|
| `Vm(on_msgbox="skip")` | VM を作成。`on_msgbox="error"` で MsgBox 時に RuntimeError を発生。 |
| `vm.run(vba_code, macro_name)` | 指定した Sub を解析・実行。 |
| `vm.set_cell(row, col, value)` | セルに値を書き込む（1始まり）。 |
| `vm.get_cell(row, col)` | セルの値を読み取る。空セルは `None`。 |
| `vm.cells()` | アクティブシートの全非空セルを `{(row, col): value}` で返す。 |
| `vm.variables()` | VBA 変数を `{name: value}` で返す。 |
| `vm.set_cell_formula(row, col, formula)` | 数式（例: `"=SUM(A1:A3)"`）をセルに設定して評価。 |
| `vm.set_cell_formula_batch(formulas)` | 複数の数式を一括設定: `{(row, col): 数式文字列}`。 |
| `vm.recalculate()` | すべての数式セルを再評価。 |
| `vm.set_sheet(name)` | アクティブシートを切り替え（存在しない場合は作成）。 |
| `vm.active_sheet()` | 現在のアクティブシート名を返す。 |
| `vm.sheet_names()` | すべてのシート名のリストを返す。 |
| `vm.get_sheet(name)` | 指定シートの全非空セルを `{(row, col): value}` で返す。 |
| `vm.save_workbook(path)` | 全シートを `.xlsx` または `.ods` に保存。 |
| `vm.cells_df()` | アクティブシートを **pandas DataFrame** として返す（pandas 要インストール）。 |
| `elixcee.run_macro(vba, name)` | 一発実行: マクロを実行して `{(row, col): value}` を返す。 |
| `elixcee.load_workbook(path)` | `.xlsx` / `.ods` を読み込んで `Vm` を返す。 |

---

## 対応状況

詳細は **[FUNCTIONS.md](FUNCTIONS.md)** を参照してください（全関数・VBA 構文対応表、Excel バージョン列付き）。

**主な対応状況:**
- **Classic (Excel 2003-)**: SUM、VLOOKUP、IF、PMT、FV、PV、NPER、RATE、IPMT、PPMT、NPV、IRR、MIRR、XNPV、XIRR、DGET、DSUM、DAVERAGE、DCOUNT、DCOUNTA、DMAX、DMIN ほか 100+ の基本関数
- **2007〜2019**: IFERROR、COUNTIFS/SUMIFS、XOR、IFS、SWITCH、TEXTJOIN、MAXIFS/MINIFS
- **365/2021**: XLOOKUP、XMATCH、FILTER、SORT、UNIQUE、SEQUENCE、LET、LAMBDA、MAP、REDUCE
- **2024/365**: TEXTSPLIT、TEXTBEFORE、TEXTAFTER、VSTACK、HSTACK、TAKE、DROP、CHOOSECOLS、CHOOSEROWS
- **VBA**: For/If/While/With/On Error/Function/`Type...End Type`/名前付き範囲/UDT配列

### 名前付き範囲

VBA で `Range("A1:B5").Name = "MyData"` として名前付き範囲を登録すると、範囲アドレスを受け取るあらゆる箇所でその名前を使えます:

```vba
Range("MyData").Value = 0          ' 範囲内の全セルに書き込む
For Each cell In Range("MyData")   ' セルを走査する
    total = total + cell
Next cell
```

名前付き範囲は `vm.named_ranges`（`dict[str, str]`、小文字化した名前 → アドレス）に保存されます。

### 条件構文（COUNTIF / SUMIF / SUMIFS など）

| 条件 | 例 | 意味 |
|---|---|---|
| 数値 | `10` | 完全一致する数値 |
| 文字列 | `"apple"` | 大文字小文字を区別しない文字列一致 |
| 比較 | `">5"`、`"<=10"`、`"<>"` | 数値比較 |
| ワイルドカード | `"a*"`、`"?bc"` | `*` = 任意の文字列、`?` = 任意の1文字 |

### Application オブジェクト

| プロパティ/メソッド | 説明 | 挙動 |
|---|---|---|
| `Application.Calculation = xlCalculationManual` | 自動再計算を無効化 | **有効** |
| `Application.Calculation = xlCalculationAutomatic` | 自動再計算を有効化し、全数式セルを再評価 | **有効** |
| `Application.ScreenUpdating = False/True` | 画面更新を抑制 | **No-op**（画面が存在しない） |
| `Application.EnableEvents = False/True` | イベント発火を無効化/有効化 | **No-op**（イベントが存在しない） |
| `Application.DisplayAlerts = False/True` | ダイアログ表示を抑制 | **No-op**（ダイアログが存在しない） |
| `Application.StatusBar = "..."` / `False` | ステータスバーの文字列を設定/クリア | **No-op**（UI が存在しない） |
| `Application.Cursor = xlWait` / `xlDefault` | カーソル形状を変更 | **No-op**（UI が存在しない） |
| `Application.CutCopyMode = False` | クリップボードモードを解除 | **有効**（内部で保持しているクリップボード状態をクリア） |

> **No-op** のプロパティはパースと受理はされますが、効果はありません。これにより、マクロ冒頭の `Application.ScreenUpdating = False` のような VBA のパフォーマンス最適化の書き方を、変更せずそのまま実行できます。

## Microsoft Excel での round-trip 検証

elixcee のワークブック保存経路は、Microsoft Excel for Mac 上で実際に作成した、
サニタイズ済みの Microsoft Excel 製 `.xlsm` フィクスチャ 5 件を用いて検証されています。

**検証済みの範囲：**

- Excel で作成したワークブックを開く
- elixcee でセルを編集する
- 別名保存・上書き保存（in-place save）の両方
- 修復警告なしで Microsoft Excel から再度開ける
- 数式・既存のセル書式・結合セル・非表示の行/列・VBA プロジェクトのバイト列・
  未知の ZIP パーツ・relationship が保持される

**未検証の範囲：**

- 保存後の VBA マクロの実行
- 再生成されるワークシート XML に埋め込まれる、テーブル・データの入力規則・
  条件付き書式・ハイパーリンク・コメント・名前の定義・グラフ・画像・印刷設定

詳細な結果は
[`compat/oracle-excel-com/results/0.9.0-A_summary.md`](compat/oracle-excel-com/results/0.9.0-A_summary.md)
を参照してください。

## 未対応関数

詳細リストは **[FUNCTIONS.md — Not Yet Supported](FUNCTIONS.md#not-yet-supported)** を参照してください。

主な未対応カテゴリ:
- **統計**: NORM.S.DIST、T.INV、F.DIST、CHISQ.DIST ほか
- **テキスト**: REPT、NUMBERVALUE、PHONETIC
- **対象外**: IMAGE（URL参照）、GROUPBY（ピボット集計）、TRIMRANGE（使用頻度低）

---

## ステータス凡例

| マーク | 意味 |
|---|---|
| 完了 | 実装・テスト済み |
| 未定 | スケジュール未決定 |

---

## 開発フェーズ

| フェーズ | 内容 | 状況 |
|---|---|---|
| Phase 1 | Rust プロジェクト初期化 + pyo3 バインディング | 完了 |
| Phase 2 | VBA パーサー MVP（Sub/End Sub, 代入, Cells） | 完了 |
| Phase 3 | 仮想 Excel VM（変数, セルストレージ, インタプリタ） | 完了 |
| Phase 3.5 | Excel フォーミュラエンジン（SUM, IF, VLOOKUP, Application.Calculation など） | 完了 |
| Phase 4 | 制御構文（For ループ, If 分岐, 算術式） | 完了 |
| Phase 5 | Python インターフェース（Vm クラス, run_macro, load_workbook, MsgBox） | 完了 |
| Phase 6 | ワークシート関数の大幅拡充（100+ 関数, 118 テスト） | 完了 |
| Phase 7 | 高度な VBA 構文（ElseIf, Exit, For Each, On Error, Function, 配列, While-Wend） | 完了 |
| Phase 8 | Range API（ClearContents, Offset, Sheets.Cells, WorksheetFunction, マルチシート） | 完了 |
| Phase 9 | マルチシート対応（Sheets HashMap, With Sheets, Python API, load_workbook 全シート） | 完了 |
| Phase 10 | ワークシート関数拡充（数学・三角・統計・配列スピル・Lambda 関数） | 完了 |
| Phase 11 | ユーザー定義型（Type...End Type）、名前付き範囲、RANDARRAY、pandas 連携、型スタブ | 完了 |
| Phase D1 | rust_xlsxwriter 削除、手書き XLSX（zip）出力（依存: 5→4） | 完了 |
| Phase D2 | pest/pest_derive 削除、手書き再帰下降 VBA パーサー（依存: 4→3） | 完了 |
| Phase D3 | calamine をランタイム依存から除去、手書き XLSX/ODS リーダー（依存: 3→2） | 完了 |
| Perf R4 | SUM/AVERAGE/MIN/MAX fast path（`Vec<Variant>` 省略）、RangeWrite dirty フラグ集約 | 完了 |
| CLI | スタンドアロン `elixcee` バイナリ、pyo3 オプション化、GitHub Actions リリースワークフロー | 完了 |
| Milestone A | JSON Agent Contract（`--json`）、エラー分類、MsgBox メッセージログ | 完了 |
| Milestone A.1 | JSON contract hardening（`serde_json` による構造検証テスト、メッセージログのライフサイクル、エラーコード表のドキュメント化） | 完了 |
| Milestone A.5 | ソースロケーション追跡 — parse/runtime エラーへの line/column 付与 | 完了 |
| Milestone B1 | `check` サブコマンド — parse 診断、エントリポイント存在確認、`MsgBox` 等の対話操作検出 | 完了 |
| Milestone B1.1 | `check`: 未定義 Sub/Function 呼び出し検出、未対応構文（no-op）検出 | 完了 |
| Milestone B2 | マルチモジュールプロジェクト — 複数 `.bas` ファイル、`Module.Sub` 修飾エントリポイント、モジュール間名前衝突検出 | 完了 |
| Milestone B3 | 決定的なブラックボックステスト（`tests/blackbox.rs`、宣言的な `.toml` フィクスチャ） | 完了 |
| Milestone B4 | `snapshot` サブコマンド — VBA を実行せずにワークブックのセルを読む | 完了 |
| Milestone B5a | `test-workbook` サブコマンド — 生成された境界値入力によるプロパティベーステスト | 完了 |
| Milestone B6a | `diagnose` サブコマンド — 存在しないシート/ワークブック、配列範囲外の根本原因診断 | 完了 |
| Milestone B6b | `diagnose`: Copy/Paste 形状不一致 + クリップボード状態 | 完了 |
| Milestone B6c | `diagnose`: シート保護（`Protect`/`Unprotect`） | 完了 |
| Milestone B6c2 | `diagnose`: 結合セルを考慮したCopy/Paste診断 | 完了 |
| Milestone B6d | `diagnose-workbook` — 生成ケース群にわたる根本原因診断 | 完了 |
| Milestone B7a | Copy/Paste診断のための複数領域`Range`/`Union`/`Areas`基盤 | 完了 |
| Milestone B7b | Copy/Paste診断のための非表示行・列メタデータ基盤 | 完了 |
| Phase 3A-1 | `compat/vba-semantics/` 値正しさスイート: 208→301ケース（新カテゴリ6件）。単一行`If`文dispatch・`Boolean`算術（`True`=-1）・`WorksheetFunction`のBoolean係数・`Empty`等価比較の各バグを修正 | 完了 |
| Phase 3A-2 | CI `wasm` job: `wasm-pack`実ビルド（nodejs/webターゲット）+ Node/browser条件スモークテストをGitHub Actionsへ配線 | 完了 |
| 0.5.0 | VBA構造的意味論（`:`複数文区切り、文書化された伝播規則を持つ`Variant::Null`、alias安全な実`Nothing`状態、runtime `With`ターゲットスタック）と`@elixcee/xlsx`の実consumer/実ブラウザ検証（packed tarballインストール、実headless Chromeスモーク、bundle-safe WASMローディング、`readFile()`）を統合。`compat/vba-semantics/`は301→386ケースへ。新規publicな`Variant::Null` enum variantのため`elixcee-types`を0.2.0へbump。crates.io・PyPI・GitHub Releasesへ公開済み | 完了 |
