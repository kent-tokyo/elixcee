# elixcee

[![CI](https://github.com/kent-tokyo/elixcee/actions/workflows/ci.yml/badge.svg)](https://github.com/kent-tokyo/elixcee/actions/workflows/ci.yml)
[![Docs](https://docs.rs/elixcee/badge.svg)](https://docs.rs/elixcee/1.0.12/elixcee/)
[![Version](https://img.shields.io/badge/version-1.0.12-blue.svg)](https://github.com/kent-tokyo/elixcee/releases/tag/v1.0.12)

elixcee 是使用 Rust/Python 编写的无头 Excel 运行时。无需安装 Microsoft Excel，
即可编辑 `.xlsx`／`.xlsm`、重新计算受支持的公式，并运行或诊断数据处理 VBA。
项目也提供 CLI 和实验性的 JavaScript/WASM 包。

当前版本为 `1.0.12`。浏览器 playground 可体验区域选择、单元格编辑、格式、筛选、
表格、Chart 预览、透视汇总、工作表操作、公式重算以及受限的 VBA 执行与诊断。浏览器中的 VBA 仅支持受限的
数据处理子集，不保证任意 VBA 或原生 PivotTable 的编辑。

[English](README.md) | [日本語](README_ja.md) | **中文**

## 从这里开始

1. [快速开始](docs/quickstart-zh.md) — 完成第一次工作簿操作。
2. [初学者教程](docs/tutorial-beginners-zh.md) — 学习编辑、公式和 VBA。
3. [浏览器 playground](https://kent-tokyo.github.io/elixcee/playground/) — 无需安装即可试用。

playground支持英语、日本語和简体中文切换。
本地开发说明见[playground文档](playground/README-zh.md)。

## 安装

```bash
pip install elixcee
```

CLI二进制文件可从[GitHub Releases](https://github.com/kent-tokyo/elixcee/releases)获取。

可复用的 Rust crate 包括 `elixcee`（原生/Python runtime）和
`elixcee-wasm`（浏览器/Node.js runtime bridge）。playground是独立的静态
Web应用：它把 WASM runtime、JavaScript XLSX writer 和 Excel风格界面组合在一起。
[elixcee-wasm API文档](https://docs.rs/elixcee-wasm)可单独查看。
详见[crate/API边界](docs/crate-api-boundary.md)。

## Python

```python
import elixcee

vm = elixcee.load_workbook("input.xlsx")
vm.set_cell(1, 1, 10)              # 与 Excel 一样从 1 开始
vm.set_cell_formula(1, 2, "=A1*2")
vm.recalculate()
vm.run("""Sub Macro()
    Cells(1, 3).Value = Cells(1, 2).Value + 1
End Sub""", "Macro")
vm.save_workbook("output.xlsx")
```

Python API还支持范围、样式、表格、验证、撤销/重做、事件和流式读写。
接口详情见[elixcee.pyi](elixcee.pyi)。

## CLI

```text
elixcee <file.bas>... <MacroName> --file input.xlsx --output result.xlsx
elixcee check <file.bas>... [--entry MacroName] [--json]
elixcee snapshot <workbook.xlsx|ods> [--json]
elixcee diagnose <file.bas>... <MacroName> --file input.xlsx [--json]
elixcee test-workbook fixture.toml [--json]
```

自动化场景请使用 `--json`。输出契约见 [docs/agent-contract.md](docs/agent-contract.md)。

## 支持范围

支持常见数据处理 VBA（`If`、`For`、`Do`、数组、多工作表、`Range`／`Cells`等），
以及算术、查找、文本、逻辑、日期、统计、财务和动态数组公式。

这不是 Excel 桌面应用的完整替代品。UI 操作会被跳过、简化建模或报告错误。
Chart／Drawing／Pivot 和公式兼容性属于有限支持；未建模的 OOXML 可能被保留、拒绝或丢失。
浏览器 playground 会在无法保证无损保留时拒绝外部链接、PivotTable／PivotCache、嵌入媒体、threaded comments、slicer 和 custom XML。

请先阅读 [v1支持契约](docs/v1-support-contract.md)、[函数列表](FUNCTIONS.md)和[限制](docs/limits.md)。

## 开发

当前优先事项和发布条件见[路线图](ROADMAP.md)。未发布变更见[CHANGELOG.md](CHANGELOG.md)。

License: [MIT](docs/licensing.md)。另请参阅 [third-party notices](THIRD_PARTY_NOTICES.md)。
