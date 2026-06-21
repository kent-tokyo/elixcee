# elixcee

Microsoft Excel をインストールすることなく、Linux / macOS / Windows 上で Excel マクロ（VBA）のデータ処理ロジックを高速にエミュレートして実行するライブラリです。

コアエンジンは **Rust**、Python バインディングは **pyo3 + maturin** で提供します。

---

## インストール

```bash
pip install elixcee
```

開発版（ソースからビルド）:

```bash
python3 -m venv .venv && source .venv/bin/activate
maturin develop
```

---

## クイックスタート

```python
import elixcee

# VBA マクロを実行し、結果セルをすべて取得
cells = elixcee.run_macro("""
Sub FillSquares()
    For i = 1 To 5
        Cells(i, 1).Value = i * i
    Next i
End Sub
""", "FillSquares")
# cells == {(1,1): 1, (2,1): 4, (3,1): 9, (4,1): 16, (5,1): 25}

# Python からセルを事前設定してマクロを実行
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

# Excel ファイルのセルデータを読み込んでマクロを実行
vm = elixcee.load_workbook("data.xlsx")
vm.run(vba_code, "ProcessData")
result_cells = vm.cells()   # {(row, col): value, ...}

# セルにワークシート数式を設定して評価
vm.set_cell_formula(4, 1, "=SUM(A1:A3)")
print(vm.get_cell(4, 1))   # A列1〜3行の合計

# MsgBox の動作を制御
vm = elixcee.Vm(on_msgbox="skip")   # MsgBox を無視（デフォルト）
vm = elixcee.Vm(on_msgbox="error")  # MsgBox 時に RuntimeError を発生
```

---

## 対応状況

### VBA 構文

| 構文 | 例 | 状況 |
|---|---|---|
| Sub / End Sub | `Sub MySub() ... End Sub` | 完了 |
| 変数代入 | `a = 10` | 完了 |
| セル書き込み | `Cells(1, 1).Value = a` | 完了 |
| セル読み取り | `x = Cells(1, 1).Value` | 完了 |
| コメント | `' コメント` | 完了 |
| Application.Calculation | `Application.Calculation = xlCalculationAutomatic` | 完了 |
| 式中変数参照 | `Cells(1, 1).Value = a + 1` | 完了 |
| For ループ | `For i = 1 To N ... Next i` | 完了 |
| For ループ（ステップ指定） | `For i = 10 To 1 Step -2` | 完了 |
| If 条件分岐 | `If x > 0 Then ... Else ... End If` | 完了 |
| Do While ループ | `Do While x > 0 ... Loop` | 未定 |
| Select Case | `Select Case x ... End Select` | 未定 |

### Excel ワークシート関数（セル数式）

#### 算術・統計系

| 関数 | 説明 |
|---|---|
| `SUM` | 合計 |
| `AVERAGE` | 平均 |
| `AVERAGEIF` | 条件付き平均 |
| `AVERAGEIFS` | 複数条件の平均 |
| `MIN` / `MAX` | 最小値 / 最大値 |
| `MINIFS` / `MAXIFS` | 条件付き最小値 / 最大値 |
| `COUNT` / `COUNTA` | 数値 / 空白以外のセル数 |
| `COUNTIF` / `COUNTIFS` | 条件付きカウント |
| `SUMIF` / `SUMIFS` | 条件付き合計 |
| `SUMPRODUCT` | 要素ごとの積の和 |
| `PRODUCT` | 積 |
| `MEDIAN` | 中央値 |
| `MODE.MULT` | 最頻値 |
| `LARGE` / `SMALL` | K 番目に大きい / 小さい値 |
| `RANK` | 順位 |
| `PERCENTILE` / `PERCENTILE.INC` | パーセンタイル（inclusive） |
| `PERCENTRANK` / `PERCENTRANK.INC` | パーセントランク |
| `ROUND` / `ROUNDUP` / `ROUNDDOWN` | 四捨五入 / 切り上げ / 切り捨て |
| `INT` | 床関数（負の無限大方向に切り捨て） |
| `TRUNC` | 零方向への切り捨て |
| `MOD` | 余り |
| `RAND` | 0〜1 の擬似乱数 |
| `RANDBETWEEN` | 範囲指定の乱数整数 |
| `SUBTOTAL` | 集計方法を指定して集計（1〜6, 9, 101〜106, 109） |
| `AGGREGATE` | 拡張 SUBTOTAL（1〜6, 9, 12〜16） |

#### 論理系

| 関数 | 説明 |
|---|---|
| `IF` | 条件分岐 |
| `IFS` | 複数条件の分岐 |
| `SWITCH` | 値に応じた多分岐 |
| `AND` / `OR` / `NOT` | 論理演算 |
| `XOR` | 排他的論理和 |
| `IFERROR` | エラー時の代替値 |

#### 文字列系

| 関数 | 説明 |
|---|---|
| `LEFT` / `RIGHT` / `MID` | 文字数ベースの文字列抽出 |
| `LEFTB` / `RIGHTB` / `MIDB` | DBCS バイト数ベースの文字列抽出 |
| `LEN` / `LENB` | 文字数 / バイト数 |
| `UPPER` / `LOWER` / `PROPER` | 大文字化 / 小文字化 / 単語先頭大文字化 |
| `TRIM` | 余分な空白の削除 |
| `FIND` | 大文字小文字区別の位置検索 |
| `SEARCH` | 大文字小文字無視・ワイルドカード検索 |
| `SUBSTITUTE` | 値による置換 |
| `REPLACE` | 位置による置換 |
| `CONCATENATE` / `CONCAT` | 文字列結合 |
| `TEXTJOIN` | 区切り文字付き結合 |
| `TEXT` | 書式付き文字列変換 |
| `VALUE` | 文字列から数値への変換 |
| `EXACT` | 大文字小文字区別の完全一致 |
| `CHAR` / `UNICHAR` | コードから文字へ変換 |
| `CODE` / `UNICODE` | 先頭文字のコードポイント |
| `ASC` | 全角 → 半角変換（DBCS） |
| `JIS` | 半角 → 全角変換（DBCS） |

#### 日付・時刻系

| 関数 | 説明 |
|---|---|
| `DATE` | 日付シリアル値の生成（Excel エポック） |
| `TODAY` / `NOW` | 今日の日付 / 現在日時 |
| `YEAR` / `MONTH` / `DAY` | 年 / 月 / 日の抽出 |
| `WEEKDAY` | 曜日番号（戻り値タイプ 1〜3 対応） |
| `DAYS` | 2つの日付の日数差 |
| `EDATE` | N ヶ月後の同日 |
| `EOMONTH` | N ヶ月後の月末日 |
| `DATEDIF` | 日付差（Y / M / D / MD / YM / YD 単位） |
| `DATEVALUE` | "YYYY/MM/DD" / "YYYY-MM-DD" のパース |
| `TIME` | 時刻シリアル値の生成 |
| `TIMEVALUE` | "HH:MM:SS" のパース |
| `HOUR` / `MINUTE` / `SECOND` | 時 / 分 / 秒の抽出 |
| `NETWORKDAYS` | 稼働日数（土日除外） |
| `NETWORKDAYS.INTL` | カスタム週末指定の稼働日数 |
| `WORKDAY.INTL` | カスタム週末指定の N 稼働日後 |

#### 検索・参照系

| 関数 | 説明 |
|---|---|
| `VLOOKUP` / `HLOOKUP` | 縦方向 / 横方向検索 |
| `XLOOKUP` | 柔軟な検索（完全一致・以上・以下モード） |
| `LOOKUP` | ソート済みベクターの検索 |
| `INDEX` | 行列オフセットで値を取得 |
| `MATCH` | 値の位置を返す |
| `XMATCH` | モード・検索方向指定付きの MATCH |
| `CHOOSE` | インデックスで選択 |
| `ROW` / `COLUMN` | セル参照の行番号 / 列番号 |

#### 情報系

| 関数 | 説明 |
|---|---|
| `ISBLANK` | 空白かどうか |
| `ISERROR` / `ISERR` | エラーかどうか |
| `ISNA` | #N/A かどうか（常に FALSE — N/A 型未実装） |
| `ISNUMBER` | 数値かどうか |
| `ISTEXT` | 文字列かどうか |
| `ISLOGICAL` | 論理値かどうか |
| `ISNONTEXT` | 文字列でないかどうか |

### 条件式の書式（COUNTIF / SUMIF / SUMIFS 等）

| 条件式 | 例 | 意味 |
|---|---|---|
| 数値 | `10` | 数値の完全一致 |
| 文字列 | `"apple"` | 大文字小文字無視の文字列一致 |
| 比較演算子 | `">5"`, `"<=10"`, `"<>"` | 数値比較 |
| ワイルドカード | `"a*"`, `"?bc"` | `*` = 0文字以上、`?` = 1文字 |

### Application オブジェクト

| プロパティ / メソッド | 説明 | 動作 |
|---|---|---|
| `Application.Calculation = xlCalculationManual` | 手動計算モード | **有効** |
| `Application.Calculation = xlCalculationAutomatic` | 自動計算モード（全数式セルを再評価） | **有効** |
| `Application.ScreenUpdating = False/True` | 画面更新の抑制 / 再開 | **No-op**（画面なし）|
| `Application.EnableEvents = False/True` | イベントの無効化 / 有効化 | **No-op**（イベントなし）|
| `Application.DisplayAlerts = False/True` | ダイアログの抑制 / 解除 | **No-op**（UI なし）|
| `Application.StatusBar = "..."` / `False` | ステータスバーのテキスト設定 | **No-op**（UI なし）|
| `Application.Cursor = xlWait` / `xlDefault` | カーソル形状の変更 | **No-op**（UI なし）|
| `Application.CutCopyMode = False` | クリップボードモードの解除 | **No-op**（クリップボードなし）|

> **No-op** プロパティはパースされてエラーにならず、実行されますが何も起きません。
> これにより、マクロ冒頭の `Application.ScreenUpdating = False` などの高速化イディオムをそのまま動かすことができます。

---

## ステータス凡例

| マーク | 意味 |
|---|---|
| 完了 | 実装・テスト済み |
| 未定 | スケジュール未決定 |

---

## 開発フェーズ

| フェーズ | 内容 | 状況 |
|---|---|---|
| Phase 1 | Rust プロジェクト初期化 + pyo3 バインディング | 完了 |
| Phase 2 | VBA パーサー MVP（Sub/End Sub, 代入, Cells） | 完了 |
| Phase 3 | 仮想 Excel VM（変数, セルストレージ, インタプリタ） | 完了 |
| Phase 3.5 | Excel フォーミュラエンジン（SUM, IF, VLOOKUP, Application.Calculation など） | 完了 |
| Phase 4 | 制御構文（For ループ, If 分岐, 算術式） | 完了 |
| Phase 5 | Python インターフェース（Vm クラス, run_macro, load_workbook, MsgBox） | 完了 |
| Phase 6 | ワークシート関数の大幅拡充（100+ 関数, 118 テスト） | 完了 |
