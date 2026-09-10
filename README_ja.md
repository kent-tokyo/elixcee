# elixcee

[English](README.md) | **日本語** | [中文](README_zh.md)

Microsoft Excelなしで、Excelワークブックを直接編集し、対応する数式を再計算し、
データ処理向けVBAの実行・テスト・診断も行えるRust/Python製のヘッドレス
Excelワークブック自動化ランタイムです。
PyO3によるPython API、単体CLI、実験的な`@elixcee/xlsx` JavaScript/WASMパッケージを提供します。

バージョンは **1.0.8** です。変更点は[CHANGELOG](CHANGELOG.md)を参照してください。
JavaScriptパッケージはprivate・未公開です。
現在のブランチは公開済み1.0.8の契約に対応しています。次の変更は`[Unreleased]`に記録し、次回リリースまで公開版の機能とは区別します。

elixceeはVBA専用の実行ツールではありません。同じworkbookモデル上で、直接の
データ編集、対応する数式の再計算、VBAの実行・診断・テストを行えます。
ExcelをインストールできないCIやサーバーでの
`.xlsx`/`.xlsm`の読み書きにも利用できます。
VBAを使わず、PythonやCLIから読み込み、編集、保存、数式再計算だけを行うこともできます。

Excelデスクトップアプリの完全な代替ではありません。画面更新、グラフ、
ダイアログなどのUI機能は、スキップ・簡易モデル化・エラー化されます。
既存Chartの系列formula・cache・marker・smooth・負値表示・可視性・title、two-cell
Drawing anchor、worksheet-backed Pivot sourceには限定編集APIがあります。既存Drawingへの
line/bar/area/pie Chart作成APIもBUILD段階で追加していますが、macOS Excelでの完全な
再open検証は未完です。一般のDrawing shape編集とPivot再集計は対象外です。
完全なExcelオブジェクトモデルやOOXML互換性が必要な場合は、対応範囲を確認してください。

### 用途別の選択

| 必要なこと | 選択の目安 |
|---|---|
| セルの直接編集だけ | 一般的なExcel編集ライブラリで十分な場合があります |
| 編集と対応数式のヘッドレス再計算 | elixcee |
| Linux/macOS/CIでデータ処理VBAを実行・診断 | elixcee |
| Excelの完全なオブジェクトモデル、UI、完全なOOXML互換性 | Excelまたは専用ライブラリとelixceeの対応範囲を比較してください |

## インストール

```bash
pip install elixcee
```

CLIのバイナリは[GitHub Releases](https://github.com/kent-tokyo/elixcee/releases)から取得できます。ソースからは次のようにビルドします。

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install maturin
maturin develop
```

## CLI

```text
elixcee <file.bas>... <MacroName> [--file input.xlsx] [--sheet Sheet1]
                         [--output result.xlsx] [--json]
elixcee check <file.bas>... [--entry MacroName] [--json]
elixcee snapshot <workbook.xlsx|ods> [--json]
elixcee test-workbook fixture.toml [--json] [--seed N] [--case N]
elixcee diagnose <file.bas>... <MacroName> --file input.xlsx [--json]
elixcee diagnose-workbook fixture.toml [--json] [--seed N] [--case N] [--cases N]
```

複数モジュールでは`Module.Sub`でエントリポイントを指定できます。機械処理
には`--json`を使ってください。仕様は[docs/agent-contract.md](docs/agent-contract.md)にあります。

## Pythonの最小例

```python
import elixcee

vm = elixcee.Vm()
vm.set_cell(1, 1, 10)          # 行・列はExcelと同じ1ベース
vm.run("""
Sub DoubleIt()
    Cells(1, 2).Value = Cells(1, 1).Value * 2
End Sub
""", "DoubleIt")
print(vm.get_cell(1, 2))       # 20

vm = elixcee.load_workbook("input.xlsx")
vm.set_cell(1, 1, 10)                    # セルを編集
vm.set_cell_formula(1, 2, "=A1*2")      # 数式を設定
vm.recalculate()                        # 数式を再計算
vm.save_workbook("output.xlsx")

vm.set_cell(1, 1, 20)
vm.undo()                                # 編集を取り消し
```

数式評価、範囲、シート、VBA `Collection`／class moduleサブセット、スタイル、テーブル、データ検証、AutoFilter、
名前定義、pandas連携、`.xlsx`/`.xlsm`/`.ods`入出力にも対応しています。
`INDIRECT`はboundedなA1範囲と絶対R1C1参照、`OFFSET`は元範囲の寸法とboundedな`height`／`width`を扱います。
相対R1C1、sheet-qualified／外部参照、Excel oracle校正は現在の契約対象外です。
明示的なzero-argumentのWorkbook／Worksheetイベントは`vm.run_event(...)`で呼び出せます。
`EnableEvents`、VBAセル／数式書き込み後の一意な`Worksheet_Change`自動発火、
boundedなイベント連鎖が反映されます。明示的なA1 targetを
`Worksheet_Change(Target As Range)`へ渡す場合は`vm.run_worksheet_change(...)`、
`Workbook_Open`を先に実行する場合はopt-inの`vm.run_with_events(...)`を使用します。
APIの詳細は[elixcee.pyi](elixcee.pyi)を参照してください。
`Vm.tables()` と `Vm.data_validations()` は、テーブル列・範囲・検証規則を構造化された型付きmetadataとして返します。計算列数式や検証数式の評価は行いません。
読込済みXLSX/XLSMでは、`Vm.sheet_id(name)` と `Vm.sheet_name_for_id(sheet_id)` により、タブ順やrenameから独立した元ファイルのsheet IDを参照できます。新規sheetやODS sheetには推測したIDを付けません。
読込済みworkbookでは、既存Drawingへの限定的なline/bar/area/pie Chart作成（系列追加を含む）と、既存Chart系列の限定編集をBUILD段階で提供しています。Chart作成には既存Drawingが必要です。lineChart単独はmacOS Excelで再openできましたが、複数Chart／barChartを含むケースは修復警告が残っています。一般的なobject編集とPivot cache再集計は現在の契約対象外です。

大きなXLSX/XLSMには、全体を展開しない`open_stream(path, sheet=None)`を使えます。
`include_row_numbers=True`では`(行番号, 値)`を返し、`max_rows=N`で読み取り行数を
制限できます。`max_row_bytes=N`では1行のXMLバッファ上限も指定できます。
`max_columns=N`では1行の列数上限も指定できます。
`timeout_ms=N`では次の行を待つ時間（ミリ秒）を制限できます。
`create_stream(path)`はXLSX用の追記型writerです。`max_rows=N`や
`max_columns=N`や`max_pending_bytes=N`で受け入れる出力を制限できます。
byte予算はcloseまでの累積値で、保持RSSやconstant-memoryの保証ではありません。
[制限とUnreleased G1の変更](docs/limits.md)を参照してください。
行ごとの上限と総作業量を分ける場合は`create_stream_bounded(...)`を使います。
`Vm(timeout_ms=N)`または`run_macro(..., timeout_ms=N)`でVBA実行時間を制限できます。
同じ`Vm`で同じソースを再実行する場合は、解析済みASTを再利用します。
`vm.fork()`でバッチ処理用の独立したVMコピーを作成できます。
`vm.snapshot()`は全シートの値、タブ順、名前定義、計算モード、表示状態、結合範囲、非表示区間を返します。
`include_formulas=True`で数式本文も追加できます。CLIのsnapshotとは異なる、より詳細な形式です。
`diagnose_macro(vba_code, macro_name, workbook_path)`でCLIの`diagnose --json`と同じ構造化診断JSONを取得できます。

通常readerでは`load_workbook(..., max_work_units=N, timeout_ms=N,
cancellation=token)`で総work量、期限、協調キャンセルを指定できます。CLIの
`snapshot`は`--max-work-units N`、`--timeout-ms N`、`--cancel-file PATH`を受け付けます。
runの`--file`読込は既定budgetとSIGINTによる中断に対応し、これら3つのオプションは受け付けません。キャンセルは協調方式のため、
OSのブロッキング読込中は次のZIP chunk境界で検出されます。

対応するVBA構文・ワークシート関数は[FUNCTIONS.md](FUNCTIONS.md)にまとめています。
既知の制約と診断形式は[docs/](docs/)を参照してください。

初心者向けには[日本語クイックスタート](docs/quickstart-ja.md)、[日本語チュートリアル](docs/tutorial-beginners-ja.md)、
[ブラウザーplayground](playground/README-ja.md)を用意しています。playground内で英語・日本語・簡体中文を切り替えられます。

## 開発

```bash
cargo test --workspace
cargo fmt --all -- --check
cargo clippy --all-targets --all-features -- -D warnings
```

開発フェーズ、現在の状態、リリースゲートは[ROADMAP.md](ROADMAP.md)に集約しています。
公開済み1.0.8は、文書化したG0/G1の基盤と、段階的なG2/G3/G4の機能を含みます。
互換性の制約とセキュリティ方針は[docs/](docs/)を参照してください。
`Vm.tables()` と `Vm.data_validations()` は、テーブル列・範囲・検証規則を構造化された型付きmetadataとして返します。計算列数式や検証数式の評価は行いません。
