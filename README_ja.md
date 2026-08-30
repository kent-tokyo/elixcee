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
`include_row_numbers=True`では`(行番号, 値)`を返し、`create_stream(path)`はXLSX用の
追記型writerです。

対応するVBA構文・ワークシート関数は[FUNCTIONS.md](FUNCTIONS.md)にまとめています。
既知の制約と診断形式は[docs/](docs/)を参照してください。

## 開発

```bash
cargo test --workspace
cargo fmt --check
cargo clippy --all-targets --all-features -- -D warnings
```

計画は[ROADMAP.md](ROADMAP.md)、ライセンスは[THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md)を参照してください。
