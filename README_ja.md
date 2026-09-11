# elixcee

[![CI](https://github.com/kent-tokyo/elixcee/actions/workflows/ci.yml/badge.svg)](https://github.com/kent-tokyo/elixcee/actions/workflows/ci.yml)
[![Docs](https://docs.rs/elixcee/badge.svg)](https://docs.rs/elixcee/1.0.10/elixcee/)
[![Version](https://img.shields.io/badge/version-1.0.10-blue.svg)](https://github.com/kent-tokyo/elixcee/releases/tag/v1.0.10)

Microsoft Excelなしで、`.xlsx`／`.xlsm`を編集し、対応する数式を再計算し、
データ処理向けVBAを実行・診断できるRust/Python製のヘッドレスExcelランタイムです。
CLIと実験的なJavaScript/WASMパッケージも提供します。

[English](README.md) | **日本語** | [中文](README_zh.md)

## まずここから

1. [クイックスタート](docs/quickstart-ja.md) — 最初の操作を実行します。
2. [初心者向けチュートリアル](docs/tutorial-beginners-ja.md) — 編集、数式、VBAを学びます。
3. [ブラウザーplayground](https://kent-tokyo.github.io/elixcee/playground/) — インストールなしで試せます。

playgroundは英語・日本語・簡体中文を切り替えられます。
ローカル開発は[playgroundの説明](playground/README-ja.md)を参照してください。

## インストール

```bash
pip install elixcee
```

CLIバイナリは[GitHub Releases](https://github.com/kent-tokyo/elixcee/releases)から取得できます。

## Python

```python
import elixcee

vm = elixcee.load_workbook("input.xlsx")
vm.set_cell(1, 1, 10)              # Excelと同じ1ベース
vm.set_cell_formula(1, 2, "=A1*2")
vm.recalculate()
vm.run("""Sub Macro()
    Cells(1, 3).Value = Cells(1, 2).Value + 1
End Sub""", "Macro")
vm.save_workbook("output.xlsx")
```

範囲、スタイル、テーブル、検証、Undo/Redo、イベント、ストリーミング入出力にも対応します。
APIの詳細は[elixcee.pyi](elixcee.pyi)を参照してください。

## CLI

```text
elixcee <file.bas>... <MacroName> --file input.xlsx --output result.xlsx
elixcee check <file.bas>... [--entry MacroName] [--json]
elixcee snapshot <workbook.xlsx|ods> [--json]
elixcee diagnose <file.bas>... <MacroName> --file input.xlsx [--json]
elixcee test-workbook fixture.toml [--json]
```

自動処理には`--json`を使います。仕様は[docs/agent-contract.md](docs/agent-contract.md)にあります。

## 対応範囲

一般的なデータ処理VBA（`If`、`For`、`Do`、配列、複数シート、`Range`／`Cells`など）と、
算術・検索・文字列・論理・日付・統計・財務・動的配列などの数式に対応します。

Excelデスクトップの完全な代替ではありません。UI操作はスキップ、簡易モデル化、または報告されます。
Chart／Drawing／Pivotと数式互換性は限定対応で、未対応OOXMLは操作により保持・拒否・欠落する場合があります。

[v1サポート契約](docs/v1-support-contract.md)、[関数一覧](FUNCTIONS.md)、[制限](docs/limits.md)を確認してください。

## 開発

現在の優先事項とリリース条件は[ロードマップ](ROADMAP.md)を参照してください。
未リリースの変更は[CHANGELOG.md](CHANGELOG.md)にあります。

License: [MIT](docs/licensing.md)。[third-party notices](THIRD_PARTY_NOTICES.md)も参照してください。
