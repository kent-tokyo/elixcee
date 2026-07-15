# elixcee

[English](README.md) | **日本語** | [中文](README_zh.md)

Microsoft Excel をインストールすることなく、Linux / macOS / Windows 上で Excel マクロ（VBA）のデータ処理ロジックを高速にエミュレートして実行するライブラリです。

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
- elixcee の VBA インタプリタは、一般的なデータ処理マクロで使われる VBA のサブセット（ループ、条件分岐、セルの読み書き、文字列・数学関数、複数シートへのアクセス）をカバーしています。Excel の UI 操作（グラフ作成、書式設定、ダイアログ）は no-op です。

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

`elixcee diagnose` はマクロを一度だけ実行し、存在しないシート・存在しないワークブック・配列の範囲外アクセス・Copy/Paste の形状不一致など、Excelがその操作を拒否する具体的な理由を根拠付きで説明します（単なるエラー文字列ではありません）：

```bat
elixcee diagnose Main.bas --file report.xlsx --entrypoint Main.Run --json
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

分類ルールと JSON スキーマの詳細は [docs/agent-contract.md](docs/agent-contract.md) を参照してください。

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
