# elixcee

[English](README.md) | **日本語** | [中文](README_zh.md)

Microsoft Excelなしで、データ処理向けのExcel VBAのサブセットを実行・
テスト・診断するRust製ランタイムです。PyO3によるPython API、単体CLI、
実験的な`@elixcee/xlsx` JavaScript/WASMパッケージを提供します。

現在のリリースは **0.28.0** です。

Excelデスクトップアプリの完全な代替ではありません。画面更新、グラフ、
ダイアログなどのUI機能は、スキップ・簡易モデル化・エラー化されます。

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
```

数式評価、範囲、シート、スタイル、テーブル、データ検証、AutoFilter、
名前定義、pandas連携、`.xlsx`/`.xlsm`/`.ods`入出力にも対応しています。
APIの詳細は[elixcee.pyi](elixcee.pyi)を参照してください。

大きなXLSX/XLSMには、全体を展開しない`open_stream(path, sheet=None)`を使えます。
`include_row_numbers=True`では`(行番号, 値)`を返し、`max_rows=N`で読み取り行数を
制限できます。`max_row_bytes=N`では1行のXMLバッファ上限も指定できます。
`max_columns=N`では1行の列数上限も指定できます。
`timeout_ms=N`では次の行を待つ時間（ミリ秒）を制限できます。
`create_stream(path)`はXLSX用の追記型writerです。`max_rows=N`や
`max_columns=N`や`max_pending_bytes=N`も指定して、保留中の出力を制限できます。
`Vm(timeout_ms=N)`または`run_macro(..., timeout_ms=N)`でVBA実行時間を制限できます。
同じ`Vm`で同じソースを再実行する場合は、解析済みASTを再利用します。
`vm.fork()`でバッチ処理用の独立したVMコピーを作成できます。
`vm.snapshot()`で全シートの独立した読み取り専用スナップショットを取得できます。
`include_formulas=True`を指定すると、計算結果とは別に保存数式も取得できます。
スナップショットにはワークシートのタブ順も含まれます。
スナップショットには実行時の名前定義も含まれます。
スナップショットには現在の`calculation_mode`（`automatic`または`manual`）も含まれます。
`diagnose_macro(vba_code, macro_name, workbook_path)`でCLIの`diagnose --json`と同じ構造化診断JSONを取得できます。

対応するVBA構文・ワークシート関数は[FUNCTIONS.md](FUNCTIONS.md)にまとめています。
既知の制約と診断形式は[docs/](docs/)を参照してください。

## 開発

```bash
cargo test --workspace
cargo fmt --check
cargo clippy --all-targets --all-features -- -D warnings
```

計画は[ROADMAP.md](ROADMAP.md)、ライセンスは[THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md)を参照してください。
