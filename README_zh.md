# elixcee

[English](README.md) | [日本語](README_ja.md) | **中文**

无需安装 Microsoft Excel，即可在 Linux、macOS 和 Windows 上高速模拟并执行 Excel 宏（VBA）数据处理逻辑的库。

核心引擎使用 **Rust** 编写，通过 **pyo3 + maturin** 提供 Python 绑定。

## 名称由来

**elixcee** = **Excel** + **elixir**（万灵药） + **C**

治愈 Excel 依赖的"万灵药"——借助 Rust 以 C 语言级别的速度运行。

---

## 与同类工具的对比

| 功能 | **elixcee** | xlwings | LibreOffice UNO | openpyxl | xlcalculator |
|------|:-----------:|:-------:|:---------------:|:--------:|:------------:|
| 运行 VBA 宏 | 是 | 是 | 是（部分） | 否 | 否 |
| 需要 Excel | 否 | 是 | 否 | 否 | 否 |
| 需要 LibreOffice | 否 | 否 | 是 | 否 | 否 |
| 公式求值 | 是 | 是 | 是 | 否 | 是 |
| macOS/Linux/Windows | 是 | 部分 | 是 | 是 | 是 |
| 简洁的 Python API | 是 | 是 | 否 | 是 | 是 |
| 读取 .xlsx | 是 | 是 | 是 | 是 | 是 |
| 读取 .ods | 是 | 是 | 是 | 否 | 否 |
| 写入 .xlsx | 是 | 是 | 是 | 是 | 否 |
| 写入 .ods | 是 | 是 | 是 | 否 | 否 |
| 执行速度 | Rust（原生） | COM/IPC（慢） | IPC（慢） | — | Python |

**说明：**
- **xlwings** 在 macOS 上需要通过 AppleScript 调用 Excel for Mac，在 Windows 上需要通过 COM 调用 Excel。Linux 支持需要运行中的 Excel 实例或云端桥接。
- **LibreOffice UNO** 启动耗时超过 1 秒，且 API 复杂。VBA 通过 LibreOffice 自有解释器运行，行为可能与 Excel 不完全一致。
- **openpyxl** 从 .xlsx 文件中读取缓存的公式值，但不支持运行时重新求值。
- **xlcalculator** 可在 Python 中重新求值 Excel 公式，但不支持 VBA。
- elixcee 的 VBA 解释器覆盖了典型数据处理宏所用的 VBA 子集（循环、条件分支、单元格读写、字符串/数学函数、多工作表访问）。Excel UI 操作（图表、格式设置、对话框）均为 no-op。

---

## 安装

```bash
pip install elixcee
```

开发版（从源码构建）：

```bash
python3 -m venv .venv && source .venv/bin/activate
maturin develop
```

---

## CLI（Windows / Linux / macOS）

无需 Python 的独立可执行文件，可从 [Releases](https://github.com/kent-tokyo/elixcee/releases) 页面下载。

| 下载 | 适用平台 |
|---|---|
| [elixcee-x86_64-windows.exe](https://github.com/kent-tokyo/elixcee/releases/latest/download/elixcee-x86_64-windows.exe) | Windows x64 |
| [elixcee-x86_64-linux](https://github.com/kent-tokyo/elixcee/releases/latest/download/elixcee-x86_64-linux) | Linux x64 |
| [elixcee-aarch64-macos](https://github.com/kent-tokyo/elixcee/releases/latest/download/elixcee-aarch64-macos) | macOS Apple Silicon |

### 用法

```
elixcee <vba_file>... <MacroName> [OPTIONS]

参数：
  <vba_file>...  一个或多个 VBA 源文件（.vbs / .bas / .txt）。传入多个文件时，
                 Sub/Function 名称在整个项目内共享，遇到同名时用 Module.Sub 指定。
  <MacroName>    要执行的 Sub 名称（最后一个参数）

选项：
  --file <path>    从电子表格加载单元格数据（.xlsx / .xlsm / .ods）
  --sheet <name>   活动工作表名称（默认：--file 的第一个工作表）
  --output <path>  将结果单元格保存到电子表格（.xlsx / .ods）
  --json           输出单个机器可读的 JSON 对象（结果或错误），而非纯文本
```

### 示例

执行 VBA 文件并将结果打印到标准输出：

```bat
elixcee macro.vbs ProcessData
```

从 Excel 文件加载数据，执行宏，并保存结果：

```bat
elixcee macro.vbs ProcessData --file input.xlsx --output result.xlsx
```

输出格式 — 每行一个非空单元格，地址与值用制表符分隔：

```
A1    Hello
B1    42
A2    3.14
```

`MsgBox` 的内容将输出到标准输出。

### 多文件（多模块项目）

传入多个源文件即可运行跨多个模块的项目。Sub/Function 名称在整个项目内共享——如果同名的 Sub/Function 存在于多个模块中，使用 `Module.Sub` 指定具体的一个（模块名优先取 `Attribute VB_Name`，否则取文件名）：

```bat
elixcee Helpers.bas Main.bas Main.ProcessData
```

目前还没有项目清单文件（具体支持范围、跨模块名称冲突的处理方式等详见 [docs/agent-contract.md](docs/agent-contract.md)）。

### JSON 输出（面向脚本 / AI Agent）

加上 `--json` 可输出单个机器可读的 JSON 对象，而非纯文本：

```bat
elixcee macro.vbs ProcessData --json
```

```json
{"schema_version":1,"ok":true,"entrypoint":"ProcessData","duration_ms":0.42,"cells":[{"sheet":"sheet1","address":"A1","value":42}],"messages":[]}
```

完整契约（错误码、退出码、`messages` 语义）：[docs/agent-contract.md](docs/agent-contract.md)。

### 不执行宏的静态分析

`elixcee check` 在不执行的情况下检查一个或多个 `.bas` 文件：parse 错误、指定入口宏是否存在、代码中任何位置的未定义 Sub/Function 调用，以及 `MsgBox` 等交互操作。所有位置参数都视为文件，入口点（如果指定）始终通过 `--entry` 传入而非位置参数——因此 `elixcee check *.bas` 可以在不假设任何特定入口点的前提下检查项目中的每个模块。

```bat
elixcee check macro.vbs --entry ProcessData --json
```

```json
{"schema_version":1,"ok":true,"diagnostics":[]}
```

### 工作簿快照

`elixcee snapshot` 直接读取 `.xlsx`/`.xlsm`/`.ods` 文件——不执行 VBA——并以 Markdown（默认）或 `--json` 时以 JSON 打印每个工作表的非空单元格：

```bat
elixcee snapshot Book1.xlsx --json
```

```json
{"schema_version":1,"ok":true,"file":"Book1.xlsx","sheets":[{"name":"Sheet1","sheet_id":"1","stable_id":"sheet1","cells":[{"address":"A1","value":42}]}]}
```

`stable_id` 来自文件自身的 `sheetId`（若不存在则按位置回退生成)，它**不是** VBA 的 `CodeName` 属性。完整设计理由见 [docs/agent-contract.md](docs/agent-contract.md)。

### 基于属性的工作簿测试

`elixcee test-workbook` 会用生成的边界值输入（空白、`0`、`1`、`-1`、接近溢出的数值、空/短/长字符串）反复对同一个起始工作簿运行宏，并在每次运行中检查 panic、运行时错误、超时以及 Excel 错误值——每个用例都从完全独立的工作簿状态开始：

```toml
# fixture.toml
name = "order calculation"
workbook = "orders.xlsx"
vba_files = ["Main.bas"]
macro = "Main.Process"
cases = 100
seed = 42

[[inputs]]
range = "Input!B2:B10"
strategy = "boundary_numeric"

[[assertions]]
range = "Result!A1:F100"
rule = "no_excel_errors"
```

```bat
elixcee test-workbook fixture.toml --json
```

失败的用例会报告其 seed 和 case index，以便精确复现：`elixcee test-workbook fixture.toml --seed 42 --case 17`。完整 schema、strategy 与 assertion 规则见 [docs/agent-contract.md](docs/agent-contract.md)。

### Excel 操作诊断

`elixcee diagnose` 执行一次宏，并给出证据说明 Excel *为什么* 会拒绝该操作——缺失的工作表、缺失的工作簿、数组越界、Copy/Paste 形状不匹配，或写入受保护的工作表——而不是只给出一句裸的错误字符串：

```bat
elixcee diagnose Main.bas --file report.xlsx --json Main.Run
```

```json
{
  "schema_version": 1,
  "ok": false,
  "message": "Sheet 'Sales2025' not found",
  "location": {"file": "Main.bas", "line": 2, "column": 5},
  "root_causes": [
    {
      "code": "WORKSHEET_NOT_FOUND",
      "certainty": "definite",
      "expression": "Worksheets(\"Sales2025\")",
      "requested": "Sales2025",
      "available": ["input", "sales2026", "summary"],
      "suggested": "sales2026",
      "suggestions": ["did you mean 'sales2026'?"]
    }
  ],
  "messages": []
}
```

`Range("A1:C10").Copy` 之后执行 `Range("E1:F10").PasteSpecial`，会同时报告形状不匹配以及两条语句各自的位置：

```json
{
  "code": "PASTE_SHAPE_MISMATCH",
  "source_addr": "A1:C10", "source_rows": 10, "source_cols": 3,
  "dest_addr": "E1:F10", "dest_rows": 10, "dest_cols": 2,
  "copy_location": {"file": "Main.bas", "line": 2, "column": 5},
  "suggestions": [
    "resize the destination to E1:G10",
    "or specify only the top-left cell E1"
  ]
}
```

写入已 `.Protect` 的工作表会报告是哪个工作表以及修复方法：

```json
{
  "code": "SHEET_PROTECTED",
  "sheet": "sheet1",
  "suggestions": ["unprotect the sheet first: Worksheets(\"sheet1\").Unprotect"]
}
```

完整分类规则与 JSON schema 见 [docs/agent-contract.md](docs/agent-contract.md)。

### 从源码构建

```bash
cargo build --release --bin elixcee
# 生成文件：target/release/elixcee（Windows 为 elixcee.exe）
```

---

## 快速开始

```python
import elixcee

# 运行 VBA 宏并获取所有结果单元格
cells = elixcee.run_macro("""
Sub FillSquares()
    For i = 1 To 5
        Cells(i, 1).Value = i * i
    Next i
End Sub
""", "FillSquares")
# cells == {(1,1): 1, (2,1): 4, (3,1): 9, (4,1): 16, (5,1): 25}

# 从 Python 预设单元格数据，再执行宏
vm = elixcee.Vm()
vm.set_cell(1, 1, 100)
vm.set_cell(2, 1, 200)
vm.run("""
Sub CalcTotal()
    total = Cells(1,1).Value + Cells(2,1).Value
    Cells(3,1).Value = total
End Sub
""", "CalcTotal")
print(vm.get_cell(3, 1))   # 300
print(vm.variables())       # {"total": 300}

# 从现有 Excel 文件加载数据，再执行宏
vm = elixcee.load_workbook("data.xlsx")
vm.run(vba_code, "ProcessData")
result_cells = vm.cells()   # {(row, col): value, ...}

# 在单元格上设置工作表公式并求值
vm.set_cell_formula(4, 1, "=SUM(A1:A3)")
print(vm.get_cell(4, 1))   # 第1~3行A列的合计

# 控制 MsgBox 行为
vm = elixcee.Vm(on_msgbox="skip")   # 静默忽略 MsgBox（默认）
vm = elixcee.Vm(on_msgbox="error")  # MsgBox 时抛出 RuntimeError
```

---

## Python API

| 方法 | 说明 |
|---|---|
| `Vm(on_msgbox="skip")` | 创建 VM。`on_msgbox="error"` 时 MsgBox 抛出 RuntimeError。 |
| `vm.run(vba_code, macro_name)` | 解析并执行指定的 Sub。 |
| `vm.set_cell(row, col, value)` | 向单元格写入值（1-based）。 |
| `vm.get_cell(row, col)` | 读取单元格值。空单元格返回 `None`。 |
| `vm.cells()` | 以 `{(row, col): value}` 返回活动工作表的所有非空单元格。 |
| `vm.variables()` | 以 `{name: value}` 返回所有 VBA 变量。 |
| `vm.set_cell_formula(row, col, formula)` | 设置公式（如 `"=SUM(A1:A3)"`）并求值。 |
| `vm.set_cell_formula_batch(formulas)` | 批量设置公式：`{(row, col): 公式字符串}`。 |
| `vm.recalculate()` | 重新求值所有公式单元格。 |
| `vm.set_sheet(name)` | 切换活动工作表（不存在则创建）。 |
| `vm.active_sheet()` | 返回当前活动工作表名称。 |
| `vm.sheet_names()` | 返回所有工作表名称列表。 |
| `vm.get_sheet(name)` | 以 `{(row, col): value}` 返回指定工作表的所有非空单元格。 |
| `vm.save_workbook(path)` | 将所有工作表保存为 `.xlsx` 或 `.ods`。 |
| `vm.cells_df()` | 将活动工作表作为 **pandas DataFrame** 返回（需安装 pandas）。 |
| `elixcee.run_macro(vba, name)` | 一次性执行：运行宏并返回 `{(row, col): value}`。 |
| `elixcee.load_workbook(path)` | 将 `.xlsx` / `.ods` 文件加载到 `Vm` 中。 |

---

## 函数覆盖范围

详见 **[FUNCTIONS.md](FUNCTIONS.md)**（完整函数与 VBA 语法参考，含 Excel 版本列）。

**主要覆盖：**
- **Classic（Excel 2003-）**：SUM、VLOOKUP、IF、PMT、FV、PV、NPER、RATE、IPMT、PPMT、NPV、IRR、MIRR、XNPV、XIRR、DGET、DSUM、DAVERAGE、DCOUNT、DCOUNTA、DMAX、DMIN 等 100+ 核心函数
- **2007–2019**：IFERROR、COUNTIFS/SUMIFS、XOR、IFS、SWITCH、TEXTJOIN、MAXIFS/MINIFS
- **365/2021**：XLOOKUP、XMATCH、FILTER、SORT、UNIQUE、SEQUENCE、LET、LAMBDA、MAP、REDUCE
- **2024/365**：TEXTSPLIT、TEXTBEFORE、TEXTAFTER、VSTACK、HSTACK、TAKE、DROP、CHOOSECOLS、CHOOSEROWS
- **VBA**：For/If/While/With/On Error/Function/`Type...End Type`/命名范围/UDT 数组

### 命名范围

在 VBA 中使用 `Range("A1:B5").Name = "MyData"` 注册命名范围，之后可在任何接受范围地址的地方使用该名称：

```vba
Range("MyData").Value = 0          ' 向范围内所有单元格写入
For Each cell In Range("MyData")   ' 遍历单元格
    total = total + cell
Next cell
```

命名范围存储在 `vm.named_ranges` 中（`dict[str, str]`，键为小写名称，值为地址）。

### 条件语法（COUNTIF / SUMIF / SUMIFS 等）

| 条件 | 示例 | 含义 |
|---|---|---|
| 数字 | `10` | 精确数值匹配 |
| 字符串 | `"apple"` | 不区分大小写的字符串匹配 |
| 比较 | `">5"`、`"<=10"`、`"<>"` | 数值比较 |
| 通配符 | `"a*"`、`"?bc"` | `*` = 任意字符，`?` = 单个字符 |

### Application 对象

| 属性 / 方法 | 说明 | 行为 |
|---|---|---|
| `Application.Calculation = xlCalculationManual` | 禁用自动重算 | **有效** |
| `Application.Calculation = xlCalculationAutomatic` | 启用自动重算并重新求值所有公式 | **有效** |
| `Application.ScreenUpdating = False/True` | 抑制屏幕刷新 | **No-op**（无界面） |
| `Application.EnableEvents = False/True` | 禁用/启用事件触发 | **No-op**（无事件） |
| `Application.DisplayAlerts = False/True` | 抑制对话框 | **No-op**（无对话框） |
| `Application.StatusBar = "..."` / `False` | 设置/清除状态栏文本 | **No-op**（无界面） |
| `Application.Cursor = xlWait` / `xlDefault` | 更改光标形状 | **No-op**（无界面） |
| `Application.CutCopyMode = False` | 取消剪贴板模式 | **No-op**（无剪贴板） |

> **No-op** 属性会被解析并接受，但不产生任何效果。这使得 VBA 宏中的性能优化写法（如 `Application.ScreenUpdating = False`）能够不经修改直接运行。

---

## 暂不支持

详见 **[FUNCTIONS.md — Not Yet Supported](FUNCTIONS.md#not-yet-supported)**。

主要缺口：
- **统计函数**：NORM.S.DIST、T.INV、F.DIST、CHISQ.DIST 等
- **文本函数**：REPT、NUMBERVALUE、PHONETIC
- **超出范围**：IMAGE（URL 图片获取）、GROUPBY（透视聚合）、TRIMRANGE

---

## 状态说明

| 标记 | 含义 |
|---|---|
| 完成 | 已实现并测试 |
| 待定 | 尚未排期 |

---

## 开发阶段

| 阶段 | 内容 | 状态 |
|---|---|---|
| Phase 1 | Rust 项目初始化 + pyo3 Python 绑定 | 完成 |
| Phase 2 | VBA 解析器 MVP（Sub/End Sub、赋值、Cells） | 完成 |
| Phase 3 | 虚拟 Excel VM（变量、单元格存储、解释器） | 完成 |
| Phase 3.5 | Excel 公式引擎（SUM、IF、VLOOKUP、Application.Calculation 等） | 完成 |
| Phase 4 | 控制流（For 循环、If 分支、算术表达式） | 完成 |
| Phase 5 | Python 接口（Vm 类、run_macro、load_workbook、MsgBox） | 完成 |
| Phase 6 | 工作表函数大幅扩充（100+ 函数，118 个测试） | 完成 |
| Phase 7 | 高级 VBA 语法（ElseIf、Exit、For Each、On Error、Function、数组、While-Wend） | 完成 |
| Phase 8 | Range API（ClearContents、Offset、Sheets.Cells、WorksheetFunction、多工作表） | 完成 |
| Phase 9 | 多工作表支持（Sheets HashMap、With Sheets、Python API、load_workbook 全表） | 完成 |
| Phase 10 | 工作表函数扩充（数学、三角、统计、数组溢出、Lambda 函数） | 完成 |
| Phase 11 | 用户自定义类型（Type...End Type）、命名范围、RANDARRAY、pandas 集成、类型存根 | 完成 |
| Phase D1 | 移除 rust_xlsxwriter，手写 XLSX（zip）输出（依赖：5→4） | 完成 |
| Phase D2 | 移除 pest/pest_derive，手写递归下降 VBA 解析器（依赖：4→3） | 完成 |
| Phase D3 | 从运行时依赖中移除 calamine，手写 XLSX/ODS 读取器（依赖：3→2） | 完成 |
| Perf R4 | SUM/AVERAGE/MIN/MAX 快速路径（跳过 `Vec<Variant>`），RangeWrite dirty 标志批量更新 | 完成 |
| CLI | 独立 `elixcee` 可执行文件；pyo3 可选化；GitHub Actions 发布工作流 | 完成 |
| Milestone A | JSON Agent Contract（`--json`）、错误分类、MsgBox 消息日志 | 完成 |
| Milestone A.1 | JSON contract 加固（`serde_json` 结构化校验测试、消息日志生命周期、错误码文档化） | 完成 |
| Milestone A.5 | 源码位置追踪 — 为 parse/runtime 错误附加 line/column | 完成 |
| Milestone B1 | `check` 子命令 — parse 诊断、入口点存在性检查、`MsgBox` 等交互操作检测 | 完成 |
| Milestone B1.1 | `check`：未定义 Sub/Function 调用检测、不支持语法（no-op）检测 | 完成 |
| Milestone B2 | 多模块项目 — 支持多个 `.bas` 文件、`Module.Sub` 限定入口点、跨模块名称冲突检测 | 完成 |
| Milestone B3 | 确定性黑盒测试（`tests/blackbox.rs`，声明式 `.toml` fixture） | 完成 |
| Milestone B4 | `snapshot` 子命令 — 不执行 VBA，直接读取工作簿单元格 | 完成 |
| Milestone B5a | `test-workbook` 子命令 — 基于生成的边界值输入的属性测试 | 完成 |
| Milestone B6a | `diagnose` 子命令 — 缺失工作表/工作簿、数组越界等根因诊断 | 完成 |
| Milestone B6b | `diagnose`：Copy/Paste 形状不匹配 + 剪贴板状态 | 完成 |
| Milestone B6c | `diagnose`：工作表保护（`Protect`/`Unprotect`） | 完成 |
