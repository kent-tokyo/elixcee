# elixcee

[![CI](https://github.com/kent-tokyo/elixcee/actions/workflows/ci.yml/badge.svg)](https://github.com/kent-tokyo/elixcee/actions/workflows/ci.yml)
[![Docs](https://docs.rs/elixcee/badge.svg)](https://docs.rs/elixcee/1.0.11/elixcee/)
[![Version](https://img.shields.io/badge/version-1.0.11-blue.svg)](https://github.com/kent-tokyo/elixcee/releases/tag/v1.0.11)

elixcee是用Rust/Python编写的无头Excel运行时。无需Microsoft Excel即可编辑
`.xlsx`／`.xlsm`、重新计算受支持的公式，并运行或诊断数据处理VBA。
项目也提供CLI和实验性的JavaScript/WASM版本。

[快速开始](docs/quickstart-zh.md) · [初学者教程](docs/tutorial-beginners-zh.md) ·
[浏览器playground](https://kent-tokyo.github.io/elixcee/playground/)

[English](README.md) | [日本語](README_ja.md) | **中文**

## 安装

```bash
pip install elixcee
```

## Python

```python
import elixcee

vm = elixcee.load_workbook("input.xlsx")
vm.set_cell(1, 1, 10)              # 1-based (行, 列)
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

## 支持范围

支持工作簿编辑、公式计算和数据处理VBA（`If`、`For`、`Do`、数组、多工作表、
`Range`／`Cells`）。UI操作和未建模OOXML会按操作被简化、保留、拒绝或报告。
请先阅读[支持契约](docs/v1-support-contract.md)、[函数列表](FUNCTIONS.md)和[限制](docs/limits.md)。

[路线图](ROADMAP.md) · [CHANGELOG](CHANGELOG.md) · [MIT license](docs/licensing.md)
