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
- **Classic（Excel 2003-）**：SUM、VLOOKUP、IF、PMT、DGET、DSUM、DAVERAGE、DCOUNT、DCOUNTA、DMAX、DMIN 等 100+ 核心函数
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
- **财务函数**：FV、PV、RATE、NPER、NPV、IRR、XNPV、XIRR 等
- **数学函数**：FACT、PERMUT、GCD、LCM、SIGN 等
- **统计函数**：NORM.DIST、CORREL、COVARIANCE.S 等
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
