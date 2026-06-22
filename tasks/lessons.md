# Lessons Learned

## L1: パーサーのドット対応は `if` ではなく `while` にする

`MODE.MULT` に対応するため「ドットの後にアルファベットが続けば関数名の一部として読む」処理を追加した。
最初は `if` で1回だけ処理したが、`NETWORKDAYS.INTL` や `WORKDAY.INTL` のように複数ドットが入る名前に対応するためには `while` ループにしなければならない。
**教訓:** ドット区切り名はループで消費する。1回限りの `if` は次の拡張で必ず壊れる。

---

## L2: Excel の DBCS バイト数は UTF-8 バイト数ではない

`LENB("日本語")` は UTF-8 では 9 バイト（1文字3バイト）だが、Excel では 6 を返す（1文字2バイト）。
Excel の B 系関数は「ASCII は1バイト、それ以外は2バイト」という DBCS 規則を使う。
**教訓:** `s.len()`（UTF-8 バイト数）を使ってはいけない。`char::is_ascii()` で判定して 1 or 2 を返すヘルパーを用意する。

```rust
fn char_byte_width(c: char) -> usize {
    if (c as u32) <= 0x7F { 1 } else { 2 }
}
```

---

## L3: Excel の日付シリアル値には「1900年うるう年バグ」がある

Excel は1900年を誤ってうるう年として扱い、存在しない「1900年2月29日（シリアル値60）」を含む。
これは互換性のために意図的に残されているバグであり、互換実装でも再現しなければならない。
`date_to_serial` では「1900年3月1日以降のすべての日付を +1 シフト」することで対応した。
**教訓:** `serial 1 = 1900-01-01` は正しいが、内部計算には +1 オフセットが必要。`chrono` のような標準ライブラリと単純に対応しない。

---

## L4: SEARCH の wildcard_match はプレフィックスマッチが必要

`SEARCH("h*o", "Hello")` のテストが最初に失敗した。原因は `wildcard_match` が「テキスト全体とパターンが一致するか」を判定するのに対し、SEARCH は「テキストの途中から始まるサブ文字列にマッチするか」を探す必要があるため。

固定長ウィンドウ（パターン長と同じ長さのスライス）で照合しようとすると、`*` が展開できない。

対策：`wildcard_match_prefix` を別途定義し、パターンを消費しきった時点で `true` を返す（テキストの残りは無視）。

```rust
// 誤: 固定長ウィンドウで照合
h_upper[i..i+n_chars.len()]  // * が展開できない

// 正: 位置 i 以降のテキスト全体を渡してプレフィックスマッチ
wildcard_match_prefix(&h_upper[i..], &n_chars)
```

**教訓:** 「検索」と「完全一致」は異なるマッチングセマンティクスを持つ。関数を分けて設計する。

---

## L5: `String::extend_from_slice` は存在しない

`Vec<u8>` にある `extend_from_slice` を `String` に対して使おうとしてコンパイルエラーになった。
`String` に `&[char]` を追加するには `extend(slice.iter())` を使う。

```rust
// NG
result.extend_from_slice(&chars[start..end]);

// OK
result.extend(chars[start..end].iter());
```

**教訓:** `String` は `Vec<u8>` ではない。文字スライスの追加は `.extend(iter)` で行う。

---

## L6: 条件マッチングヘルパーは早めに共通化する

COUNTIF を実装した時点で、SUMIF・COUNTIFS・SUMIFS・AVERAGEIF・AVERAGEIFS・MAXIFS・MINIFS がすべて同じ「条件文字列のパース→比較」ロジックを必要とすることが明らかだった。
`matches_criteria(val, criteria)` を最初から汎用ヘルパーとして切り出したことで、後続7関数の実装コストが大幅に下がった。
**教訓:** 条件集計系関数が1つ出てきたら、すぐに汎用ヘルパーに切り出す。

---

## L7: IS* 関数は引数を evaluate して `Result` で判定する

`ISERROR(1/0)` は `1/0` を評価しようとするが、その結果がエラーであることを検知して `TRUE` を返さなければならない。
実装は IFERROR と同じパターン：`evaluate(&args[0], cells).is_err()`。
**教訓:** IS* 関数はエラーを握りつぶす（エラーを返さない）。通常の `?` 演算子でエラーを伝播させると動かない。

---

## L8: ROW / COLUMN は引数を評価せずに AST を直接検査する

`ROW(A5)` は A5 セルの「値」ではなく「行番号」（5）を返す。
引数を `evaluate()` してしまうとセルの値が返ってしまう。
AST ノード（`FormulaExpr::CellRef { row, .. }`）を直接パターンマッチして行番号を取り出す。
**教訓:** セル参照そのものを引数に取る関数（ROW, COLUMN, INDIRECT, OFFSET など）は evaluate しない。

---

## L9: 計画した未使用コードはすぐに削除する

`collect_indexed` というヘルパー関数を SUMIFS 用に計画したが、実際には `collect_values` + インデックスアクセスで十分だった。
残したままにすると `dead_code` 警告が出てノイズになる。
**教訓:** 計画段階で書いたが使わなかったコードはコンパイル警告が出る前に削除する。「あとで使うかも」は無用な負債。

---

## L10: AGGREGATE は既存関数の再利用で実装できる

AGGREGATE の各 function_num は他の関数（AVERAGE, COUNT, SUM, MEDIAN, LARGE, SMALL など）と同じロジック。
`func_average(rest, cells)` のように既存関数を直接呼び出すことで重複を避けた。
**教訓:** 「集計関数のディスパッチャ」を実装するときは、既存の関数実装を引数スライスごと渡して再利用する。

---

## L11: テスト中の未使用変数が型推論を壊す

```rust
let c = HashMap::new();
let c2 = cells_from(&[...]);
assert_eq!(calc("=PERCENTILE(A1:A5,0.5)", &c2), ...);
```

`c` が使われないと `HashMap<_, _>` の型パラメータが推論できずコンパイルエラーになった。
使わない `c` は削除するだけでよい。
**教訓:** テスト内の `HashMap::new()` は必ず何かに使うか、型注釈を付けるか、削除する。

---

## L12: 擬似乱数は SystemTime ナノ秒 + LCG で十分

`RAND()` と `RANDBETWEEN()` の実装に外部クレート（`rand`）を使わず、`SystemTime::now().subsec_nanos()` を LCG でミックスする方法を採用した。
テスト用途としては十分な品質で、依存関係を増やさない。
**教訓:** シンプルな擬似乱数が必要な場面では `rand` クレートを追加しなくてもよい。ただし暗号用途には使ってはいけない。

---

## L13: 日付関数の EDATE は月末クランプが必要

`EDATE(DATE(2000,1,31), 1)` は2月31日になるが、2月は31日まで存在しない。
Excel は「その月の最終日にクランプ」するので、`d.min(days_in_month(y, m))` で処理する。
**教訓:** 月を加算する日付関数はすべて月末クランプが必要。EOMONTH は常に月末なので問題ないが、EDATE は d を保持しようとするので注意。

---

## L14: WRAPCOLS/WRAPROWS は引数が Variant::Array のときフラット化が必要

`WRAPCOLS(SEQUENCE(6), 2)` のように配列を返す関数を第1引数に渡すと、`collect_values` が `[Array([1,2,3,4,5,6])]`（配列を1要素として包んだ Vec）を返す。
そのまま `vals.len()` を使うと 1 になり、wrap が全く効かない。
`flatten_array_vals` ヘルパーを挟んで Variant::Array を展開してから処理する。

```rust
let vals = flatten_array_vals(collect_values(&args[0], cells)?);
```

**教訓:** 配列操作関数の第1引数は「セル範囲ではなく別の配列関数の戻り値」になりうる。collect_values の結果をそのまま使わず、Variant::Array を展開するヘルパーを通す。

---

## L15: cells_mut() 呼び出しで dirty フラグを立てる方法

`last_nonempty_row` 等のインデックス検索を O(log n) にするために、`cells_mut()` でインデックスを dirty にしたかった。
Rust のボローチェッカー上、`cells_mut()` が `&mut HashMap` を返す前に `self.cell_index_dirty = true` を書けば、その後の参照はハッシュマップだけになるため問題ない。

```rust
pub fn cells_mut(&mut self) -> &mut HashMap<(u32, u32), CellContent> {
    self.cell_index_dirty = true;  // ← 返却前に設定
    self.sheets.get_mut(&self.active_sheet).expect("active sheet must exist")
}
```

**教訓:** 「`&mut self` メソッドが `&mut` フィールドを返す前に別フィールドを更新する」パターンはボローチェッカーを通る。返却後は通らない。

---

## L17: pyo3 を optional にする場合は `#[cfg(any(feature = "python", test))]` パターンが必要

`pyo3` を optional dependency にして `python` feature でガードすると、pyo3 コードを使うすべての `use` 文・型・関数・`impl` ブロックに `#[cfg(feature = "python")]` が必要になる。
テストモジュールが pyo3 コードに依存しない場合でも、`use super::*` が parent のエクスポートを参照するため、
`CellContent` のような型は `#[cfg(any(feature = "python", test))]` で条件付きにインポートしないと unused import 警告が出る。

```rust
// NG: always imported → unused warning when building without `python` feature
use vm::{CellContent, Variant, Vm};

// OK: Variant/Vm は save_workbook_impl で常に必要。CellContent は pyo3 か test のときだけ必要
use vm::{Variant, Vm};
#[cfg(any(feature = "python", test))]
use vm::CellContent;
```

**教訓:** `#[cfg(feature = "...")]` を追加するとき、その型を使う他の `use` 文も連鎖して条件化が必要になる。`cargo build --lib`（feature なし）でビルドして unused import 警告をすべて潰してから commit する。

---

## L16: Type...End Type の実装で Dim var As TypeName は DimRecord に変換する

`Dim p As Person` は VBA の構文上「型付き宣言」だが、VM は型を知らなければデフォルト初期化できない。
パーサーで「VBA ビルトイン型（Integer/String/Boolean 等）以外の型名は DimRecord を生成する」と決めると、VM 側で type_defs を参照して Record を初期化できる。
ビルトイン型の場合は従来通り no-op（Stmt::Dim）にすればよい。

**教訓:** パーサーとランタイムの責務分担：「ユーザー定義型かどうかの判定」はパーサーが担い（ビルトイン型リストで判定）、「フィールドの初期化」はランタイムが担う。
