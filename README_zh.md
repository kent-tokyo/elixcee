# elixcee

[English](README.md) | [日本語](README_ja.md) | **中文**

elixcee 是一个使用 Rust/Python 编写的无头 Excel 工作簿自动化运行时，可在不安装 Microsoft Excel 的情况下编辑工作簿、重新计算受支持的公式，并运行、测试和诊断面向数据处理的 VBA 子集。项目提供 PyO3 Python API、独立 CLI，以及实验性的 `@elixcee/xlsx` JavaScript/WASM 包。

elixcee 不是 VBA 专用执行器，而是工作簿自动化运行时：可以在同一个工作簿模型上直接编辑数据、重新计算受支持的公式，以及执行、诊断和测试受支持的 VBA。它适合没有安装 Excel 的 CI 和服务器环境中的 `.xlsx`/`.xlsm` 读写。

它不是 Excel 桌面应用的完整替代品。屏幕更新、图表和对话框等 UI 功能会被跳过、简化建模或报告错误。需要完整 Excel 对象模型或完整 OOXML 兼容性时，请先检查支持边界。

版本：**1.0.5**。变更记录见 [CHANGELOG](CHANGELOG.md)。
JavaScript 包仍为 private，尚未发布。

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

vm = elixcee.load_workbook("input.xlsx")
vm.set_cell(1, 1, 10)                    # 编辑单元格
vm.set_cell_formula(1, 2, "=A1*2")      # 设置公式
vm.recalculate()                        # 重新计算公式
vm.save_workbook("output.xlsx")

vm.set_cell(1, 1, 20)
vm.undo()                                # 撤销编辑
```

Python API 还支持公式、范围、工作表、VBA `Collection` 与类模块子集、样式、表格、数据验证、AutoFilter、
名称定义、pandas，以及 `.xlsx`/`.xlsm`/`.ods` 文件。接口签名见 [elixcee.pyi](elixcee.pyi)，
VBA 和工作表函数列表见 [FUNCTIONS.md](FUNCTIONS.md)。
`Vm.tables()` 和 `Vm.data_validations()` 返回表格列、范围及验证规则的结构化类型元数据；不会执行计算列公式或验证公式。
对于已加载的 XLSX/XLSM 工作表，`Vm.sheet_id(name)` 和 `Vm.sheet_name_for_id(sheet_id)` 可在标签顺序或重命名变化后解析原始 sheet ID；新建或 ODS 工作表不会推测 ID。

对于大型 XLSX/XLSM 文件，可使用 `open_stream(path, sheet=None)` 逐行读取；设置
`include_row_numbers=True` 后返回 `(行号, 值)` 元组，也可用 `max_rows=N` 限制读取行数，
或用 `max_row_bytes=N` 限制单行 XML 缓冲区大小，
或用 `max_columns=N` 限制每行列数，
或用 `timeout_ms=N` 限制等待下一行的时间（毫秒），
`create_stream(path)` 提供 XLSX 追加式写入器；可用 `max_rows=N` 或
`max_columns=N` 或 `max_pending_bytes=N` 限制接受的输出量。
字节预算累计到close，并非实际保留RSS或恒定内存保证。
参见[限制与Unreleased G1变更](docs/limits.md)。
如需分别限制单行和总工作量，请使用`create_stream_bounded(...)`。
可用 `Vm(timeout_ms=N)` 或 `run_macro(..., timeout_ms=N)` 限制 VBA 执行时间。
同一个 `Vm` 重复执行相同源码时会复用已解析的 AST。
可使用 `vm.fork()` 创建用于批处理的独立 VM 副本。
`vm.snapshot()` 返回所有工作表的值、标签顺序、名称定义、计算模式、可见性、合并范围和隐藏区间。
设置 `include_formulas=True` 可添加公式文本；此格式比 CLI snapshot 更详细。
可使用 `diagnose_macro(vba_code, macro_name, workbook_path)` 获取与 CLI `diagnose --json` 相同的结构化诊断 JSON。

普通读取器支持通过 `load_workbook(..., max_work_units=N, timeout_ms=N,
cancellation=token)` 设置总工作量、截止时间和协作式取消。CLI 的 `snapshot`
支持 `--max-work-units N`、`--timeout-ms N` 和 `--cancel-file PATH`。
run 的 `--file` 读取使用默认预算并支持 SIGINT，不接受这三个选项。取消请求会在下一个 ZIP 数据块边界被检测；
正在进行的操作系统阻塞读取不会被强制中断。

## 开发

```bash
cargo test --workspace
cargo fmt --check
cargo clippy --all-targets --all-features -- -D warnings
```

计划见 [ROADMAP.md](ROADMAP.md)，其他公开政策和限制见 [docs/](docs/)。
