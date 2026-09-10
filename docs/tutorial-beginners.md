# elixcee Beginner Tutorial

This tutorial introduces the complete first workflow: create a workbook in memory, edit cells, recalculate formulas, run data-processing VBA, and save an `.xlsx` or `.xlsm` file without Microsoft Excel.

## Install

```bash
python3 -m venv .venv
source .venv/bin/activate       # Windows: .venv\\Scripts\\activate
python -m pip install elixcee
```

## Create and calculate a workbook

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

Coordinates are one-based and ordered `(row, column)`, matching `Cells(row, column)` in VBA.

## Run data-processing VBA

```python
vba = """
Sub DoubleTotal()
    Cells(1, 2).Value = Cells(1, 1).Value * 2
End Sub
"""
vm.run(vba, "DoubleTotal")
print(vm.get_cell(1, 2))  # 200
```

The runtime is headless. `MsgBox` is skipped by default; use `Vm(on_msgbox="error")` when a dialog should fail the run.

## Load, edit, and save a file

```python
vm = elixcee.load_workbook("input.xlsx")
vm.set_cell(1, 1, "processed")
vm.set_cell_formula(1, 2, "=COUNTA(A1:A10)")
vm.recalculate()
vm.save_workbook("output.xlsx")
```

The same API accepts `.xlsm`. Existing macro projects are preserved on supported round trips, while charts, pivots, drawings, external links, and other OOXML parts remain subject to the documented support boundaries.

## Undo a change

```python
vm.set_cell(1, 1, "temporary")
vm.undo()
```

Use `begin_transaction()` and `commit_transaction()` to group several edits into one undoable operation.

## Use the CLI

```bash
elixcee process.bas ProcessData --file input.xlsm --output output.xlsm --json
elixcee check process.bas --json
elixcee diagnose process.bas ProcessData --file input.xlsm --json
```

## Where to go next

- [Function coverage](../FUNCTIONS.md)
- [v1 support contract](v1-support-contract.md)
- [Runnable playground](../playground/README.md)
- [Python API signatures](../elixcee.pyi)

Translations: [日本語](tutorial-beginners-ja.md) · [简体中文](tutorial-beginners-zh.md)
