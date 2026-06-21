# elixcee

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

詳細は **[FUNCTIONS_ja.md](FUNCTIONS_ja.md)** を参照してください（全関数・VBA 構文対応表、Excel バージョン列付き）。

**主な対応状況:**
- **Classic (Excel 2003-)**: SUM、VLOOKUP、IF、PMT、DGET ほか 100+ の基本関数
- **2007〜2019**: IFERROR、COUNTIFS/SUMIFS、XOR、IFS、SWITCH、TEXTJOIN、MAXIFS/MINIFS
- **365/2021**: XLOOKUP、XMATCH、FILTER、SORT、UNIQUE、SEQUENCE、LET、LAMBDA、MAP、REDUCE
- **2024/365**: TEXTSPLIT、TEXTBEFORE、TEXTAFTER、VSTACK、HSTACK、TAKE、DROP、CHOOSECOLS、CHOOSEROWS
- **VBA**: For/If/While/With/On Error/Function/`Type...End Type`/名前付き範囲/UDT配列


## 未対応関数

詳細リストは **[FUNCTIONS_ja.md — 未対応関数](FUNCTIONS_ja.md#未対応関数)** を参照してください。

主な未対応カテゴリ:
- **財務系**: FV、PV、RATE、NPER、NPV、IRR、XNPV、XIRR ほか
- **数学**: FACT、PERMUT、GCD、LCM、SIGN ほか
- **統計**: NORM.DIST、CORREL、COVARIANCE.S ほか
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
