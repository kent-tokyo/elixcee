# elixcee

[![CI](https://github.com/kent-tokyo/elixcee/actions/workflows/ci.yml/badge.svg)](https://github.com/kent-tokyo/elixcee/actions/workflows/ci.yml)
[![Docs](https://docs.rs/elixcee/badge.svg)](https://docs.rs/elixcee/1.0.11/elixcee/)
[![Version](https://img.shields.io/badge/version-1.0.11-blue.svg)](https://github.com/kent-tokyo/elixcee/releases/tag/v1.0.11)

elixceeは、Microsoft Excelなしで`.xlsx`／`.xlsm`を編集し、対応する数式を
再計算し、データ処理向けVBAを実行・診断できるRust/Python製ランタイムです。
CLIと実験的なJavaScript/WASM版も提供します。

[クイックスタート](docs/quickstart-ja.md) · [初心者向けチュートリアル](docs/tutorial-beginners-ja.md) ·
[ブラウザーplayground](https://kent-tokyo.github.io/elixcee/playground/)

[English](README.md) | **日本語** | [中文](README_zh.md)

## インストール

```bash
pip install elixcee
```

## Python

```python
import elixcee

vm = elixcee.load_workbook("input.xlsx")
vm.set_cell(1, 1, 10)              # (行, 列)の1ベース
vm.set_cell_formula(1, 2, "=A1*2")
vm.recalculate()
vm.run("""Sub Macro()
    Cells(1, 3).Value = Cells(1, 2).Value + 1
End Sub""", "Macro")
vm.save_workbook("output.xlsx")
```

## CLI

```text
elixcee <file.bas>... <MacroName> --file input.xlsx --output result.xlsx
elixcee check <file.bas>... [--entry MacroName] [--json]
elixcee diagnose <file.bas>... <MacroName> --file input.xlsx [--json]
elixcee snapshot <workbook.xlsx|ods> [--json]
```

## 対応範囲

ワークブック編集、数式再計算、データ処理VBA（`If`、`For`、`Do`、配列、
複数シート、`Range`／`Cells`）に対応します。UI操作と未対応OOXMLは操作に応じて
簡易モデル化、保持、拒否、または報告されます。[サポート契約](docs/v1-support-contract.md)、
[関数一覧](FUNCTIONS.md)、[制限](docs/limits.md)を確認してください。

[ロードマップ](ROADMAP.md) · [CHANGELOG](CHANGELOG.md) · [MIT license](docs/licensing.md)
