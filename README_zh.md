# elixcee

[English](README.md) | [日本語](README_ja.md) | **中文**

elixcee 是一个使用 Rust 编写的无头运行时，可在不安装 Microsoft Excel 的情况下运行、测试和诊断面向数据处理的 Excel VBA 子集。项目提供 PyO3 Python API、独立 CLI，以及实验性的 `@elixcee/xlsx` JavaScript/WASM 包。

它不是 Excel 桌面应用的完整替代品。屏幕更新、图表和对话框等 UI 功能会被跳过、简化建模或报告错误。

当前版本：**0.28.0**。

## 安装

```bash
pip install elixcee
```

CLI 二进制文件可从 [GitHub Releases](https://github.com/kent-tokyo/elixcee/releases) 获取。源码构建：

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

多模块项目可使用 `Module.Sub` 指定入口。脚本和 CI 应使用 `--json`；完整契约见 [docs/agent-contract.md](docs/agent-contract.md)。

## Python 示例

```python
import elixcee

vm = elixcee.Vm()
vm.set_cell(1, 1, 10)          # 行列索引从 1 开始
vm.run("""
Sub DoubleIt()
    Cells(1, 2).Value = Cells(1, 1).Value * 2
End Sub
""", "DoubleIt")
print(vm.get_cell(1, 2))       # 20
```

Python API 还支持公式、范围、工作表、样式、表格、数据验证、AutoFilter、
名称定义、pandas，以及 `.xlsx`/`.xlsm`/`.ods` 文件。接口签名见 [elixcee.pyi](elixcee.pyi)，
VBA 和工作表函数列表见 [FUNCTIONS.md](FUNCTIONS.md)。

对于大型 XLSX/XLSM 文件，可使用 `open_stream(path, sheet=None)` 逐行读取；设置
`include_row_numbers=True` 后返回 `(行号, 值)` 元组，也可用 `max_rows=N` 限制读取行数，
或用 `max_row_bytes=N` 限制单行 XML 缓冲区大小，
或用 `max_columns=N` 限制每行列数，
或用 `timeout_ms=N` 限制等待下一行的时间（毫秒），
`create_stream(path)` 提供 XLSX 追加式写入器；可用 `max_rows=N` 或
`max_columns=N` 或 `max_pending_bytes=N` 限制待处理输出。
可用 `Vm(timeout_ms=N)` 或 `run_macro(..., timeout_ms=N)` 限制 VBA 执行时间。
同一个 `Vm` 重复执行相同源码时会复用已解析的 AST。
可使用 `vm.fork()` 创建用于批处理的独立 VM 副本。
可使用 `vm.snapshot()` 获取所有工作表的独立只读快照。
指定 `include_formulas=True` 可单独获取保存的公式，而不与计算结果混合。
快照还包含工作表的标签顺序。
快照还包含运行时的名称定义。
快照还包含当前的 `calculation_mode`（`automatic` 或 `manual`）。
快照还包含每个工作表的显示状态（`visible`、`hidden` 或 `veryHidden`）。
可使用 `diagnose_macro(vba_code, macro_name, workbook_path)` 获取与 CLI `diagnose --json` 相同的结构化诊断 JSON。

## 开发

```bash
cargo test --workspace
cargo fmt --check
cargo clippy --all-targets --all-features -- -D warnings
```

计划见 [ROADMAP.md](ROADMAP.md)，其他公开政策和限制见 [docs/](docs/)。
