# elixcee はじめてのチュートリアル

このチュートリアルでは、Microsoft Excelをインストールせずに、Pythonからワークブックを編集し、数式を計算し、データ処理用のVBAを実行します。

## 1. インストール

Python 3.9以上の仮想環境を作り、elixceeをインストールします。

```bash
python3 -m venv .venv
source .venv/bin/activate       # Windows: .venv\\Scripts\\activate
python -m pip install elixcee
```

インストール確認:

```bash
python -c "import elixcee; print('elixcee is ready')"
```

## 2. 最初のワークブックを作る

`playground/hello.py`という名前で次のコードを保存します。

```python
import elixcee

vm = elixcee.Vm()
vm.set_cell(1, 1, 10)             # A1: 行・列は1から始まる
vm.set_cell(2, 1, 20)             # A2
vm.set_cell_formula(3, 1, "=SUM(A1:A2)")

print(vm.get_cell(3, 1))          # 30
```

実行:

```bash
python playground/hello.py
```

`set_cell`は値を書き込み、`set_cell_formula`は数式を保存して評価します。座標はExcel/VBAと同じ1ベースです。`A1`は`(1, 1)`、`B3`は`(3, 2)`です。

## 3. 値を変更して再計算する

```python
vm.set_cell(1, 1, 100)
vm.recalculate()
print(vm.get_cell(3, 1))          # 120
```

自動再計算が有効なVMでは、依存する数式が更新されます。明示的に全数式を更新したい場合は`recalculate()`を呼びます。

## 4. VBAでデータ処理する

elixceeはVBAをExcel画面なしで実行できます。画面操作ではなく、セルの読み書きなどデータ処理に使うコードが対象です。

```python
vba = """
Sub DoubleTotal()
    Cells(1, 2).Value = Cells(1, 1).Value * 2
End Sub
"""

vm.run(vba, "DoubleTotal")
print(vm.get_cell(1, 2))          # 200
```

`MsgBox`など画面に依存する命令は、既定ではスキップされます。失敗として扱いたい場合は`elixcee.Vm(on_msgbox="error")`を使います。

## 5. `.xlsx`を読み、編集し、保存する

```python
import elixcee

vm = elixcee.load_workbook("input.xlsx")
vm.set_cell(1, 1, "処理済み")
vm.set_cell_formula(1, 2, "=COUNTA(A1:A10)")
vm.recalculate()
vm.save_workbook("output.xlsx")
```

`.xlsm`も読み書きできます。対応範囲内ではVBAプロジェクトなどの既存パーツを保持しますが、Chart、Pivot、Drawing、外部リンクなどは[互換性資料](v1-support-contract.md)の状態を確認してください。

## 6. 1回の操作を取り消す

```python
vm.set_cell(1, 1, "変更")
vm.undo()
vm.redo()
```

複数の編集を1つの操作として扱うにはトランザクションを使います。

```python
vm.begin_transaction()
vm.set_cell(1, 1, 1)
vm.set_cell(1, 2, 2)
vm.commit_transaction()
vm.undo()                         # 2セル分をまとめて取り消す
```

## 7. CLIでVBAを実行する

標準モジュールを`process.bas`に保存し、次のように実行します。

```bash
elixcee process.bas ProcessData \
  --file input.xlsm \
  --output output.xlsm \
  --json
```

構文だけを確認する場合:

```bash
elixcee check process.bas --json
```

失敗箇所を機械的に診断する場合:

```bash
elixcee diagnose process.bas ProcessData --file input.xlsm --json
```

## 8. つまずいたとき

- `get_cell`の座標は`(row, column)`です。`get_cell(1, 2)`はB1です。
- 数式文字列には通常`=`を付けます。
- Excelの全関数・全VBAオブジェクトを再現するものではありません。対応関数は[FUNCTIONS.md](../FUNCTIONS.md)、制限は[サポート契約](v1-support-contract.md)を参照してください。
- 外部I/Oや画面操作は安全のため既定で実行しません。

## 次に読む資料

- [Python APIの型・シグネチャ](../elixcee.pyi)
- [対応関数一覧](../FUNCTIONS.md)
- [互換性と既知の制限](compatibility-known-defects.md)
- [実行可能なplayground](../playground/README.md)
