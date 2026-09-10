# elixcee クイックスタート

5分で、Excelなしのセル編集・数式計算・VBA実行を体験します。

## 1. インストール

```bash
python3 -m venv .venv
source .venv/bin/activate       # Windows: .venv\\Scripts\\activate
python -m pip install elixcee
```

## 2. ファイルを作る

`quickstart.py`を作成します。

```python
import elixcee

vm = elixcee.Vm()
vm.set_cell(1, 1, 10)                 # A1に10
vm.set_cell(2, 1, 20)                 # A2に20
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

## 3. 実行する

```bash
python quickstart.py
```

`30`と`60`が表示されれば成功です。座標はExcel/VBAと同じ1ベースで、`(行, 列)`の順です。

## 4. Excelファイルを編集する

```python
vm = elixcee.load_workbook("input.xlsx")
vm.set_cell(1, 1, "処理済み")
vm.set_cell_formula(1, 2, "=COUNTA(A1:A10)")
vm.recalculate()
vm.save_workbook("output.xlsx")
```

詳しい使い方は[初心者向けチュートリアル](tutorial-beginners-ja.md)、すぐ試せるサンプルは[playground](../playground/README.md)を参照してください。

対応関数と制限は[関数一覧](../FUNCTIONS.md)と[サポート契約](v1-support-contract.md)に記載しています。
