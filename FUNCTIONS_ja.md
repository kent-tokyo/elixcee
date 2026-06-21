# elixcee — 関数・VBA 対応リファレンス

elixcee がサポートする VBA 構文とワークシート関数の完全リファレンスです。
各ワークシート関数には「どの Excel バージョンから組み込みになったか」を示す最小バージョンを記載しています。

---

## バージョン凡例

| ラベル | 最小 Excel バージョン |
|---|---|
| Classic | Excel 2003 以前から組み込み |
| 2007 | Excel 2007 以降 |
| 2010 | Excel 2010 以降 |
| 2013 | Excel 2013 以降 |
| 2019 | Excel 2019 以降 |
| 365/2021 | Microsoft 365 / Excel 2021 以降 |
| 2024/365 | Excel 2024 / Microsoft 365（最新チャネル） |

---

## VBA 構文

| 構文 | 例 | 状態 |
|---|---|---|
| Sub / End Sub | `Sub MySub() ... End Sub` | 完了 |
| 変数代入 | `a = 10` | 完了 |
| セル書き込み | `Cells(1, 1).Value = a` | 完了 |
| セル読み取り | `x = Cells(1, 1).Value` | 完了 |
| コメント | `' コメント` | 完了 |
| Application.Calculation | `Application.Calculation = xlCalculationAutomatic` | 完了 |
| 算術式 | `Cells(1, 1).Value = a + 1` | 完了 |
| For ループ | `For i = 1 To N ... Next i` | 完了 |
| For ループ（ステップ指定） | `For i = 10 To 1 Step -2` | 完了 |
| If / Else | `If x > 0 Then ... Else ... End If` | 完了 |
| Do While ループ | `Do While x > 0 ... Loop` | 完了 |
| Select Case | `Select Case x ... End Select` | 完了 |
| While / Wend | `While x > 0 ... Wend` | 完了 |
| For Each | `For Each item In collection` | 完了 |
| With ブロック（シート） | `With Sheets("Sheet1") ... End With` | 完了 |
| With ブロック（UDT） | `With p ... .Field = val ... End With` | 完了 |
| On Error Resume Next | `On Error Resume Next` | 完了 |
| On Error GoTo ラベル | `On Error GoTo ErrH` | 完了 |
| Function / Call | `Function Foo() ... End Function` | 完了 |
| Exit For / Exit Sub | `Exit For`, `Exit Sub` | 完了 |
| 配列 / ReDim | `Dim arr(10)`, `ReDim arr(n)` | 完了 |
| Const | `Const PI = 3.14` | 完了 |
| Type ... End Type | ユーザー定義型（UDT） | 完了 |
| ネスト型 | UDT フィールドに別 UDT（`p.Addr.Street`） | 完了 |
| UDT の配列 | `Dim arr(10) As MyType` | 完了 |
| 名前付き範囲 | `Range("A1:B3").Name = "MyData"` | 完了 |
| Debug.Print | `Debug.Print x` | 完了（no-op） |
| Option Explicit | `Option Explicit` | 完了（無視） |

### Application オブジェクト

| プロパティ / メソッド | 動作 |
|---|---|
| `Application.Calculation = xlCalculationManual` | **有効** — 自動計算を無効化 |
| `Application.Calculation = xlCalculationAutomatic` | **有効** — 全数式セルを再評価 |
| `Application.ScreenUpdating = False/True` | **No-op**（画面なし） |
| `Application.EnableEvents = False/True` | **No-op**（イベントなし） |
| `Application.DisplayAlerts = False/True` | **No-op**（UI なし） |
| `Application.StatusBar = "..."` / `False` | **No-op**（UI なし） |
| `Application.Cursor = xlWait` / `xlDefault` | **No-op**（UI なし） |
| `Application.CutCopyMode = False` | **No-op**（クリップボードなし） |

---

## ワークシート関数

### 算術・統計系

| 関数 | 説明 | Excel |
|---|---|---|
| `SUM` | 合計 | Classic |
| `AVERAGE` | 平均 | Classic |
| `MIN` / `MAX` | 最小値 / 最大値 | Classic |
| `COUNT` / `COUNTA` | 数値 / 空白以外のセル数 | Classic |
| `PRODUCT` | 積 | Classic |
| `MEDIAN` | 中央値 | Classic |
| `LARGE` / `SMALL` | K 番目に大きい / 小さい値 | Classic |
| `RANK` | 順位 | Classic |
| `ROUND` / `ROUNDUP` / `ROUNDDOWN` | 四捨五入 / 切り上げ / 切り捨て | Classic |
| `INT` | 床関数（負の無限大方向） | Classic |
| `TRUNC` | 零方向への切り捨て | Classic |
| `MOD` | 余り | Classic |
| `RAND` | 0〜1 の擬似乱数 | Classic |
| `RANDBETWEEN` | 範囲指定の乱数整数 | Classic |
| `SUMPRODUCT` | 要素ごとの積の和 | Classic |
| `COMBIN` | 組み合わせ数 C(n, k) | Classic |
| `COUNTIF` / `SUMIF` / `AVERAGEIF` | 条件付きカウント / 合計 / 平均 | 2007 |
| `COUNTIFS` / `SUMIFS` / `AVERAGEIFS` | 複数条件の集計 | 2007 |
| `SUBTOTAL` | 集計方法を指定して集計（1〜6, 9, 101〜106, 109） | Classic |
| `AGGREGATE` | 拡張 SUBTOTAL（1〜6, 9, 12〜16） | 2010 |
| `PERCENTILE` / `PERCENTILE.INC` | パーセンタイル（inclusive） | Classic / 2010 |
| `PERCENTRANK` / `PERCENTRANK.INC` | パーセントランク | Classic / 2010 |
| `MODE.MULT` | 最頻値 | 2010 |
| `MINIFS` / `MAXIFS` | 条件付き最小値 / 最大値 | 2019 |

### 財務系

| 関数 | 説明 | Excel |
|---|---|---|
| `PMT` | ローンの定期支払額（`rate, nper, pv, [fv], [type]`） | Classic |

### 統計系

| 関数 | 説明 | Excel |
|---|---|---|
| `STDEV` / `STDEV.S` | 標本標準偏差 | Classic / 2010 |
| `STDEVP` / `STDEV.P` | 母標準偏差 | Classic / 2010 |
| `VAR` / `VAR.S` | 標本分散 | Classic / 2010 |
| `VARP` / `VAR.P` | 母分散 | Classic / 2010 |

### 数学・三角関数系

| 関数 | 説明 | Excel |
|---|---|---|
| `ABS` | 絶対値 | Classic |
| `SQRT` | 平方根 | Classic |
| `POWER` | 累乗 | Classic |
| `EXP` | e の累乗 | Classic |
| `LN` | 自然対数 | Classic |
| `LOG` / `LOG10` | 対数（任意底 / 底10） | Classic |
| `PI` | 円周率 π | Classic |
| `SIN` / `COS` / `TAN` | 正弦 / 余弦 / 正接 | Classic |
| `ASIN` / `ACOS` / `ATAN` / `ATAN2` | 逆三角関数 | Classic |
| `DEGREES` / `RADIANS` | 角度変換 | Classic |
| `FLOOR` / `CEILING` | 切り捨て / 切り上げ（整数） | Classic |
| `FLOOR.MATH` / `CEILING.MATH` | 倍数への切り捨て / 切り上げ | 2013 |
| `MROUND` | 最も近い倍数に丸め | Classic |

### 論理系

| 関数 | 説明 | Excel |
|---|---|---|
| `IF` | 条件分岐 | Classic |
| `AND` / `OR` / `NOT` | 論理演算 | Classic |
| `IFERROR` | エラー時の代替値 | 2007 |
| `XOR` | 排他的論理和 | 2013 |
| `IFS` | 複数条件の分岐 | 2019 |
| `SWITCH` | 値に応じた多分岐 | 2019 |

### 文字列系

| 関数 | 説明 | Excel |
|---|---|---|
| `LEFT` / `RIGHT` / `MID` | 文字数ベースの文字列抽出 | Classic |
| `LEFTB` / `RIGHTB` / `MIDB` | DBCS バイト数ベースの文字列抽出 | Classic |
| `LEN` / `LENB` | 文字数 / バイト数 | Classic |
| `UPPER` / `LOWER` / `PROPER` | 大文字化 / 小文字化 / 単語先頭大文字化 | Classic |
| `TRIM` | 余分な空白の削除 | Classic |
| `FIND` | 大文字小文字区別の位置検索 | Classic |
| `SEARCH` | 大文字小文字無視・ワイルドカード検索 | Classic |
| `SUBSTITUTE` | 値による置換 | Classic |
| `REPLACE` | 位置による置換 | Classic |
| `CONCATENATE` | 文字列結合（旧形式） | Classic |
| `TEXT` | 書式付き文字列変換 | Classic |
| `VALUE` | 文字列から数値への変換 | Classic |
| `EXACT` | 大文字小文字区別の完全一致 | Classic |
| `CHAR` | コードから文字へ変換 | Classic |
| `CODE` | 先頭文字のコードポイント | Classic |
| `ASC` | 全角 → 半角変換（DBCS） | Classic |
| `JIS` | 半角 → 全角変換（DBCS） | Classic |
| `UNICHAR` | Unicode コードポイントから文字へ変換 | 2013 |
| `UNICODE` | 先頭文字の Unicode コードポイント | 2013 |
| `CONCAT` | 文字列 / 範囲を結合 | 2019 |
| `TEXTJOIN` | 区切り文字付き結合 | 2019 |
| `TEXTSPLIT` | 区切り文字でテキストを分割して配列を返す | 2024/365 |
| `TEXTBEFORE` | 区切り文字の N 番目より前のテキストを返す | 2024/365 |
| `TEXTAFTER` | 区切り文字の N 番目より後のテキストを返す | 2024/365 |
| `VALUETOTEXT` | 値をテキスト文字列に変換する | 2024/365 |

### 日付・時刻系

| 関数 | 説明 | Excel |
|---|---|---|
| `DATE` | 日付シリアル値の生成（Excel エポック） | Classic |
| `TODAY` / `NOW` | 今日の日付 / 現在日時 | Classic |
| `YEAR` / `MONTH` / `DAY` | 年 / 月 / 日の抽出 | Classic |
| `WEEKDAY` | 曜日番号（戻り値タイプ 1〜3 対応） | Classic |
| `DATEDIF` | 日付差（Y / M / D / MD / YM / YD 単位） | Classic |
| `DATEVALUE` | "YYYY/MM/DD" / "YYYY-MM-DD" のパース | Classic |
| `TIME` | 時刻シリアル値の生成 | Classic |
| `TIMEVALUE` | "HH:MM:SS" のパース | Classic |
| `HOUR` / `MINUTE` / `SECOND` | 時 / 分 / 秒の抽出 | Classic |
| `EOMONTH` | N ヶ月後の月末日 | 2007 |
| `EDATE` | N ヶ月後の同日 | 2007 |
| `NETWORKDAYS` | 稼働日数（土日除外） | 2007 |
| `WORKDAY` | N 稼働日後の日付（土日除外） | 2007 |
| `NETWORKDAYS.INTL` | カスタム週末指定の稼働日数 | 2010 |
| `WORKDAY.INTL` | カスタム週末指定の N 稼働日後 | 2010 |
| `DAYS` | 2つの日付の日数差 | 2013 |

### 検索・参照系

| 関数 | 説明 | Excel |
|---|---|---|
| `VLOOKUP` / `HLOOKUP` | 縦方向 / 横方向検索 | Classic |
| `INDEX` | 行列オフセットで値を取得 | Classic |
| `MATCH` | 値の位置を返す | Classic |
| `CHOOSE` | インデックスで選択 | Classic |
| `INDIRECT` | 文字列からセル参照を生成 | Classic |
| `OFFSET` | 相対オフセットのセル参照 | Classic |
| `ADDRESS` | セルアドレスを文字列で返す（例: `"$A$1"`） | Classic |
| `COUNTBLANK` | 空白セルのカウント | Classic |
| `ROW` / `COLUMN` | セル参照の行番号 / 列番号 | Classic |
| `LOOKUP` | ソート済みベクターの検索 | Classic |
| `TRANSPOSE` | 行列の転置 | Classic |
| `XLOOKUP` | 柔軟な検索（完全一致・以上・以下モード） | 365/2021 |
| `XMATCH` | モード・検索方向指定付きの MATCH | 365/2021 |

### 情報系

| 関数 | 説明 | Excel |
|---|---|---|
| `ISBLANK` | 空白かどうか | Classic |
| `ISERROR` / `ISERR` | エラーかどうか | Classic |
| `ISNA` | #N/A かどうか（常に FALSE — N/A 型未実装） | Classic |
| `ISNUMBER` | 数値かどうか | Classic |
| `ISTEXT` | 文字列かどうか | Classic |
| `ISLOGICAL` | 論理値かどうか | Classic |
| `ISNONTEXT` | 文字列でないかどうか | Classic |

### 配列・スピル関数

| 関数 | 説明 | Excel |
|---|---|---|
| `FILTER` | 条件でフィルタリング | 365/2021 |
| `UNIQUE` | 重複除去 | 365/2021 |
| `SORT` | 並べ替え | 365/2021 |
| `SORTBY` | 外部配列で並べ替え | 365/2021 |
| `SEQUENCE` | 連番の生成 | 365/2021 |
| `RANDARRAY` | 乱数配列の生成 | 365/2021 |
| `TOCOL` | 配列を1列に変換 | 2024/365 |
| `TOROW` | 配列を1行に変換 | 2024/365 |
| `WRAPCOLS` | 1次元配列を列方向に折り返し | 2024/365 |
| `WRAPROWS` | 1次元配列を行方向に折り返し | 2024/365 |
| `TAKE` | 配列の先頭（または末尾）から N 要素を取得 | 2024/365 |
| `DROP` | 配列の先頭（または末尾）から N 要素を除いた残りを返す | 2024/365 |
| `VSTACK` | 複数の配列を縦方向に連結 | 2024/365 |
| `HSTACK` | 複数の配列を横方向に連結 | 2024/365 |
| `CHOOSECOLS` | 配列から指定インデックスの列（要素）を選択 | 2024/365 |
| `CHOOSEROWS` | 配列から指定インデックスの行（要素）を選択 | 2024/365 |

### Lambda・高階関数

| 関数 | 説明 | Excel |
|---|---|---|
| `LET` | 数式内に名前付き変数を定義 | 365/2021 |
| `LAMBDA` | 無名関数の定義 | 365/2021 |
| `MAP` | 各要素に LAMBDA を適用 | 365/2021 |
| `REDUCE` | LAMBDA で配列を単一値に集約 | 365/2021 |
| `SCAN` | REDUCE の中間値を全て返す | 365/2021 |
| `BYROW` | 各行に LAMBDA を適用 | 365/2021 |
| `BYCOL` | 各列に LAMBDA を適用 | 365/2021 |

### データベース系

| 関数 | 説明 | Excel |
|---|---|---|
| `DGET`     | 条件に一致する1行の値をデータベースから抽出 | Classic |
| `DSUM`     | 条件に一致する列の合計 | Classic |
| `DAVERAGE` | 条件に一致する列の平均 | Classic |
| `DCOUNT`   | 条件に一致する数値のカウント | Classic |
| `DCOUNTA`  | 条件に一致するすべての値のカウント | Classic |
| `DMAX`     | 条件に一致する列の最大値 | Classic |
| `DMIN`     | 条件に一致する列の最小値 | Classic |

---

## 条件式の書式（COUNTIF / SUMIF / SUMIFS 等）

| 条件式 | 例 | 意味 |
|---|---|---|
| 数値 | `10` | 数値の完全一致 |
| 文字列 | `"apple"` | 大文字小文字無視の文字列一致 |
| 比較演算子 | `">5"`, `"<=10"`, `"<>"` | 数値比較 |
| ワイルドカード | `"a*"`, `"?bc"` | `*` = 0文字以上、`?` = 1文字 |

---

## 未対応関数

### 対象外（スコープ外）

技術的制約または headless VBA エミュレーターとしての用途との不一致により、以下の関数は実装対象外です。

| 関数 | 理由 |
|---|---|
| `IMAGE(source, ...)` | URL から画像を取得する関数 — headless 環境では非対応 |
| `GROUPBY(row_fields, values, function, ...)` | ピボット集計機能。多次元グループ化エンジンが必要なため複雑 |
| `TRIMRANGE(range)` | 範囲の端から空白行・列を除去する関数。使用頻度が低い |

### 財務系

| 関数 | 説明 | Excel |
|---|---|---|
| `FV` | 将来価値 | Classic |
| `PV` | 現在価値 | Classic |
| `RATE` | 利率 | Classic |
| `NPER` | 期間数 | Classic |
| `IPMT` | 利息支払額 | Classic |
| `PPMT` | 元本返済額 | Classic |
| `NPV` | 正味現在価値（定期キャッシュフロー） | Classic |
| `IRR` | 内部収益率（定期キャッシュフロー） | Classic |
| `MIRR` | 修正内部収益率 | Classic |
| `XNPV` | 正味現在価値（不定期キャッシュフロー） | Classic |
| `XIRR` | 内部収益率（不定期キャッシュフロー） | Classic |

### 数学・組み合わせ論

| 関数 | 説明 | Excel |
|---|---|---|
| `FACT` | 階乗 n! | Classic |
| `PERMUT` | 順列数 P(n, k) | Classic |
| `GCD` | 最大公約数 | Classic |
| `LCM` | 最小公倍数 | Classic |
| `QUOTIENT` | 整数除算（商） | Classic |
| `SIGN` | 数値の符号（−1、0、1） | Classic |

### 統計系

| 関数 | 説明 | Excel |
|---|---|---|
| `NORM.DIST` | 正規分布（累積分布または確率密度） | 2010 |
| `NORM.INV` | 正規分布の逆関数 | 2010 |
| `T.DIST` | t 分布 | 2010 |
| `CORREL` | ピアソン相関係数 | Classic |
| `COVARIANCE.S` / `COVARIANCE.P` | 標本分散 / 母分散 | 2010 |

### 文字列系

| 関数 | 説明 | Excel |
|---|---|---|
| `REPT` | テキスト文字列を N 回繰り返す | Classic |
| `NUMBERVALUE` | ロケール指定で文字列を数値に変換 | 2013 |
| `PHONETIC` | ふりがな（フリガナ）を取り出す | Classic |
| `BAHTTEXT` | 数値をタイバーツ表記に変換 | Classic |

### 日付・時刻系

| 関数 | 説明 | Excel |
|---|---|---|
| `WEEKNUM` | 年内の週番号 | Classic |
| `ISOWEEKNUM` | ISO 8601 週番号 | 2013 |

### 検索・情報系

| 関数 | 説明 | Excel |
|---|---|---|
| `FORMULATEXT` | セルの数式をテキストとして返す | 2013 |
| `CELL` | セルのメタデータを返す（書式・アドレス等） | Classic |
| `N` | 値を数値に変換する | Classic |
| `NA` | #N/A エラーを生成する | Classic |
| `TYPE` | 値の型コードを返す | Classic |
| `ERROR.TYPE` | エラーの種類コードを返す | Classic |

### 動的配列

| 関数 | 説明 | Excel |
|---|---|---|
| `MAKEARRAY` | LAMBDA を使って配列を生成する | 2024/365 |
