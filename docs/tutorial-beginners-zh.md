# elixcee 初学者教程

本教程介绍完整的入门流程：在内存中创建工作簿、修改单元格、重新计算公式、运行数据处理用 VBA，并在不安装 Microsoft Excel 的情况下保存 `.xlsx` 或 `.xlsm` 文件。

## 安装

```bash
python3 -m venv .venv
source .venv/bin/activate       # Windows: .venv\\Scripts\\activate
python -m pip install elixcee
```

## 创建并计算工作簿

```python
import elixcee

vm = elixcee.Vm()
vm.set_cell(1, 1, 10)
vm.set_cell(2, 1, 20)
vm.set_cell_formula(3, 1, "=SUM(A1:A2)")
print(vm.get_cell(3, 1))  # 30

vm.set_cell(1, 1, 100)
vm.recalculate()
print(vm.get_cell(3, 1))  # 120
```

坐标从1开始，顺序为`（行，列）`，与VBA的`Cells(row, column)`一致。

## 运行数据处理用VBA

```python
vba = """
Sub DoubleTotal()
    Cells(1, 2).Value = Cells(1, 1).Value * 2
End Sub
"""
vm.run(vba, "DoubleTotal")
print(vm.get_cell(1, 2))  # 200
```

运行时不需要图形界面。`MsgBox`默认跳过；如果希望对话框导致运行失败，请使用`Vm(on_msgbox="error")`。

## 读取、修改并保存文件

```python
vm = elixcee.load_workbook("input.xlsx")
vm.set_cell(1, 1, "已处理")
vm.set_cell_formula(1, 2, "=COUNTA(A1:A10)")
vm.recalculate()
vm.save_workbook("output.xlsx")
```

同一API也支持`.xlsm`。支持的往返操作会保留已有宏项目；图表、数据透视表、绘图、外部链接及其他OOXML部件仍受支持边界约束。

## 撤销修改

```python
vm.set_cell(1, 1, "临时值")
vm.undo()
```

使用`begin_transaction()`和`commit_transaction()`可以把多次修改合并为一次可撤销操作。

## 使用CLI

```bash
elixcee process.bas ProcessData --file input.xlsm --output output.xlsm --json
elixcee check process.bas --json
elixcee diagnose process.bas ProcessData --file input.xlsm --json
```

接下来可以阅读[函数覆盖范围](../FUNCTIONS.md)、[v1支持契约](v1-support-contract.md)或运行[playground](../playground/README.md)。

其他语言：[English](tutorial-beginners.md) · [日本語](tutorial-beginners-ja.md)
