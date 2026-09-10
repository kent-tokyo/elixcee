# elixcee 快速开始

五分钟内，你将学会在不安装 Microsoft Excel 的情况下修改单元格、计算公式，并运行一个简单的 VBA 数据处理宏。

## 1. 安装

```bash
python3 -m venv .venv
source .venv/bin/activate       # Windows: .venv\\Scripts\\activate
python -m pip install elixcee
```

## 2. 创建 `quickstart.py`

```python
import elixcee

vm = elixcee.Vm()
vm.set_cell(1, 1, 10)                 # A1；坐标顺序是（行，列）
vm.set_cell(2, 1, 20)                 # A2
vm.set_cell_formula(3, 1, "=SUM(A1:A2)")

vba = """
Sub DoubleTotal()
    Cells(3, 2).Value = Cells(3, 1).Value * 2
End Sub
"""
vm.run(vba, "DoubleTotal")

print(vm.get_cell(3, 1))               # 30
print(vm.get_cell(3, 2))               # 60
```

运行：

```bash
python quickstart.py
```

如果看到`30`和`60`，说明运行时已经正常工作。elixcee与Excel/VBA一样使用从1开始的`（行，列）`坐标。

## 3. 编辑Excel工作簿

```python
vm = elixcee.load_workbook("input.xlsx")
vm.set_cell(1, 1, "已处理")
vm.set_cell_formula(1, 2, "=COUNTA(A1:A10)")
vm.recalculate()
vm.save_workbook("output.xlsx")
```

完整的入门流程请阅读[初学者教程](tutorial-beginners-zh.md)。可运行的示例位于[playground](../playground/README-zh.md)。

函数覆盖范围和兼容性限制见[函数列表](../FUNCTIONS.md)以及[v1支持契约](v1-support-contract.md)。

其他语言：[English](quickstart.md) · [日本語](quickstart-ja.md)
