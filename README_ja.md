# elixcee

[English](README.md) | **日本語** | [中文](README_zh.md)

Microsoft Excelなしで、ワークブックの編集、対応する数式の再計算、
データ処理向けVBAの実行・テスト・診断を行えるRust/Python製のヘッドレス
Excelワークブック自動化ランタイムです。
PyO3によるPython API、単体CLI、実験的な`@elixcee/xlsx` JavaScript/WASMパッケージを提供します。

バージョンは **1.0.5** です。変更点は[CHANGELOG](CHANGELOG.md)を参照してください。
JavaScriptパッケージはprivate・未公開です。

elixceeはVBA専用の実行ツールではありません。同じworkbookモデル上で、直接の
データ編集、対応する数式の再計算、VBAの実行・診断・テストを行えます。
ExcelをインストールできないCIやサーバーでの
`.xlsx`/`.xlsm`の読み書きにも利用できます。

Excelデスクトップアプリの完全な代替ではありません。画面更新、グラフ、
ダイアログなどのUI機能は、スキップ・簡易モデル化・エラー化されます。
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
APIの詳細は[elixcee.pyi](elixcee.pyi)を参照してください。
`Vm.tables()` と `Vm.data_validations()` は、テーブル列・範囲・検証規則を構造化された型付きmetadataとして返します。計算列数式や検証数式の評価は行いません。
読込済みXLSX/XLSMでは、`Vm.sheet_id(name)` と `Vm.sheet_name_for_id(sheet_id)` により、タブ順やrenameから独立した元ファイルのsheet IDを参照できます。新規sheetやODS sheetには推測したIDを付けません。

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

## 開発

```bash
cargo test --workspace
cargo fmt --check
cargo clippy --all-targets --all-features -- -D warnings
```

計画は[ROADMAP.md](ROADMAP.md)、文書一覧は[docs/README.md](docs/README.md)、
ライセンスは[MITの説明](docs/licensing.md)と[第三者表記](THIRD_PARTY_NOTICES.md)を参照してください。
`Vm.tables()` と `Vm.data_validations()` は、テーブル列・範囲・検証規則を構造化された型付きmetadataとして返します。計算列数式や検証数式の評価は行いません。
