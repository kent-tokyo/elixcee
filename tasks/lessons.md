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

---

## L18: 「破棄してよいキーワード引数」も VBA では意味を変えることがある

`Worksheet.Protect` に `ui_only: Option<Expr>` を追加する前は、`.Protect` の全キーワード引数（`Password:=`、`DrawingObjects:=`、`UserInterfaceOnly:=` 等）を `.PasteSpecial` の `Transpose:=` 以外のキーワードと同じ扱いで一律「評価して破棄」していた。ほとんどの引数（`Password:=` 等）はこの単純化で問題ないが、`UserInterfaceOnly:=True` だけは実 Excel の挙動そのものを反転させる——保護は手動 UI 編集だけをブロックし、マクロからの書き込みは許可し続ける。これは「保護したままマクロだけ書き込ませる」という現実の定番パターンであり、単純に破棄すると診断ツールが誤って `SHEET_PROTECTED` を報告する偽陽性になる（コミット前のアドバイザーレビューで発覚）。

**教訓:** 新しいキーワード引数を「値を保持する必要がないから破棄する」と決める前に、そのキーワードが呼び出し先の**分岐そのもの**を変えないか確認する。「大抵は無視できる」という前提は引数ごとに検証すべきで、フラグ名だけで判断しない（`UserInterfaceOnly` のように名前が一見無害でも意味が強いことがある）。

---

## L19: 「担当範囲が非重複」は「相互作用が無い」を意味しない

0.5.0スプリントで2並列subagentに`src/parser`+`src/vm`（VBA structural semantics）と`packages/xlsx`+`crates/elixcee-wasm`（consumer validation）という完全に非重複なファイル範囲を割り当てた。両者ともgit差分の重複はゼロで、それぞれ自分のテストスイートは全てpassしていた。

しかし統合後、README.mdのコード例を手動実行して初めて実バグが発覚した：Subagent Aが`parse_stmt`（ブロック形式文の dispatch）に追加した新規`Tok::Dot`アーム（`With`ブロック内の`.member`文をどこにネストしても認識する機構）を、`parse_stmt`とは別に存在する`parse_single_line_if_branch`（単一行`If`の分岐専用dispatch）は反映しておらず、`With`内の単一行If＋`.member`（`If .Value > 0 Then .Value = .Value + 1`）が無言でno-op化していた。Subagent A自身のテストは全て複数行`If`/ブロック形式でしか検証しておらず、Subagent Bはこのコードパスに一切触れない。

**教訓:** 「担当ファイルが被らない」ことは並列作業の安全性の必要条件であって十分条件ではない。同じコードベース内に「同種の処理を行う独立した第二の dispatch/parser/formatter」が存在する場合（本件のように、ブロック形式と単一行形式で別関数が同じ文法要素を別々に認識している設計）、一方だけを更新すると他方が静かに取り残される。マージ後は必ず、両方の変更が触れうる機能を実際に手で組み合わせて動かすE2E確認を行うこと——各branch自身のCIが緑でも、統合固有のバグは検出できない。

---

## L20: `cargo publish --dry-run`はローカルpathを無視し常に実registryへ依存解決する

ワークスペース内の共有crate（`elixcee-types`）へpublic enumの新規variantを追加し、依存crate（`elixcee`）側の`Cargo.toml`も`version = "0.2.0"`へ追従させた。`cargo build --release --workspace`はローカルpath解決のため問題なく成功する。

ところが`cargo package -p elixcee`/`cargo publish -p elixcee --dry-run`は、`--allow-dirty`/`--locked`/`--no-verify`のいずれを付けても、`elixcee-types`をローカルpathではなく実crates.ioインデックスへ問い合わせて依存解決しようとする。`elixcee-types 0.2.0`がまだ実際には公開されていない段階では「candidate versions found which didn't match: 0.1.0」で必ず失敗する。これはバグではなく、「ローカルでは動くがcrates.io上では壊れる」問題を公開前に検出するためのCargo自身の意図的な仕様。

**教訓:** ワークスペース内の複数crateを段階的に公開する場合、依存される側（`elixcee-types`）を先に実際にpublishし、依存する側（`elixcee`）の`cargo publish --dry-run`による最終確認はその**後**にしか意味を成さない。「両方dry-runしてから両方publish」という順序は成立しない——依存先を公開しない限り依存元のdry-runは原理的に失敗し続ける。この制約はリリース手順の設計段階で織り込んでおくこと（本件では実際に手順を1段階分やり直す形で発覚した）。
