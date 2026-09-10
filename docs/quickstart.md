# elixcee Quick Start

In five minutes, you will edit cells, calculate a formula, and run a small VBA data-processing macro without installing Microsoft Excel.

## 1. Install

```bash
python3 -m venv .venv
source .venv/bin/activate       # Windows: .venv\\Scripts\\activate
python -m pip install elixcee
```

## 2. Create `quickstart.py`

```python
import elixcee

vm = elixcee.Vm()
vm.set_cell(1, 1, 10)                 # A1; coordinates are (row, column)
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

Run it:

```bash
python quickstart.py
```

If you see `30` and `60`, the runtime is working. elixcee uses the same one-based `(row, column)` coordinates as Excel/VBA.

## 3. Edit an Excel workbook

```python
vm = elixcee.load_workbook("input.xlsx")
vm.set_cell(1, 1, "processed")
vm.set_cell_formula(1, 2, "=COUNTA(A1:A10)")
vm.recalculate()
vm.save_workbook("output.xlsx")
```

For the full beginner path, see [Beginner Tutorial](tutorial-beginners.md). The runnable sample is in the [playground](../playground/README.md).

Function coverage and compatibility limits are documented in [FUNCTIONS.md](../FUNCTIONS.md) and the [v1 support contract](v1-support-contract.md).

Translations: [日本語](quickstart-ja.md) · [简体中文](quickstart-zh.md)
