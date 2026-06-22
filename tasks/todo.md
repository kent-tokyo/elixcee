# Todo

---

## 完了済みフェーズ（要約）

| フェーズ | 内容 | テスト |
|---------|------|--------|
| Phase 1 | Rust プロジェクト初期化 + pyo3/maturin 設定 | — |
| Phase 2 | VBA パーサー MVP（Sub/End Sub、代入、Cells 読み書き） | 9 件 |
| Phase 3 | 仮想 Excel VM（Variant 型、セル/変数ストレージ、インタプリタ） | 12 件 |
| Phase 3.5 | Excel フォーミュラエンジン（SUM/IF/VLOOKUP 等、Application.Calculation） | 47 件 |
| Phase 4 | 制御構文（For/If/Else）、Cells 読み取り式 | 67 件 |
| Phase 5 | Python API（Vm クラス、run_macro、load_workbook、MsgBox） | 8 件 |
| Phase 6 | ワークシート関数 118 件拡充（COUNTIF/SUMIF/XLOOKUP/日付/文字列/情報系） | 118 件 |
| Phase 7 高 | ElseIf/Exit/ForEach/OnError/Function/Call/vb 定数/戻り値式 | — |
| Phase 7 中 | While-Wend/Const/配列/ReDim/Split/Join/Format/IIf/TypeName/Empty | — |
| Phase 8 高 | Range 一括書き込み/ClearContents/Clear/Offset/Sheets.Cells | — |
| Phase 8 中 | WorksheetFunction/Delete/Insert/Sort/Find/EntireRow | — |
| Phase 9 | マルチシート対応（sheets HashMap、With Sheets、Sheets.Add/Delete、Python API、load_workbook 全シート） | — |
| Phase 10 | 数式追加: 数学/三角/統計/配列スピル/Lambda（STDEV/POWER/SIN/FILTER/LET/LAMBDA/MAP 等） | — |
| Phase 11 一部 | save_workbook（.xlsx/.ods）、load_workbook .ods 対応、README 競合比較表 | — |
| Phase D1 | 依存削減: `rust_xlsxwriter` 削除、XLSX 書き出しを手書き XML + zip に置き換え（依存: 5→4） | — |
| Phase D2 | 依存削減: `pest` + `pest_derive` 削除、VBA パーサーを手書き再帰下降パーサー（`src/parser/mod.rs`）に全面置き換え。差分テスト 40 件で pest との AST 等価性を検証（依存: 4→3） | 40 件 |
| Phase D3 | 依存削減: `calamine` をランタイム依存から除去、手書き XML パーサー + XLSX/ODS リーダー（`src/reader.rs`）を実装。calamine は `[dev-dependencies]` に移動（テスト oracle として保持）。差分読み込みテスト 3 件で等価性を検証（依存: 3→2） | 3 件 |
| Task A | VBA パーサー強化: `Option Explicit`/`Option Base`/`Attribute VB_Name` をモジュール先頭でスキップ、`Public`/`Private`/`Friend`/`Static` 修飾子付き Sub/Function を通常 Sub として解析、`Debug.Print`/`Debug.Assert` を no-op、Sub 内 `Static` 宣言を no-op | +12 件 |
| Task B | `Variant::Date(serial)` → Python `datetime.date` 変換（`serial_to_ymd_pub` 再利用） | — |
| Task C | ワークシート関数追加: `INDIRECT`（文字列をセル参照として解釈）、`OFFSET`（相対オフセット） | +2 件 |
| Task D | `PyExcelError` Python クラス追加（`code` 属性）、`variant_to_py`/`py_to_variant` で `Variant::Error` ↔ `PyExcelError` 変換、モジュール export | — |
| Task E | 配列スピル関数追加: `FILTER`/`UNIQUE`/`SORT`/`SORTBY`/`SEQUENCE`/`TRANSPOSE` | +6 件 |
| Task F | 数式の依存関係トポロジカルソート: `extract_cell_refs`/`topo_sort_formulas` で循環参照耐性のある再計算順序を保証 | — |
| Task G | 配列整形関数追加: `TOCOL`/`TOROW`/`WRAPCOLS`/`WRAPROWS`、`flatten_array_vals` ヘルパー（SEQUENCE 等の戻り値を正しく展開） | +3 件 |
| Task H | Python API `set_cell_formula_batch({(r,c): formula})` — 一括数式設定 | — |
| Task I | Python API `vm.cells_df()` — pandas DataFrame 変換（pandas 未インストール時は ImportError）、型スタブ `elixcee.pyi` 生成 | — |
| Task J | 名前付き範囲: `Range("A1:B3").Name = "MyData"` → Vm.named_ranges 登録、全 Range 操作（Read/Write/Clear/Sort/Copy/Delete/ForEach）で透過解決 | +2 件 |
| Task K | `RANDARRAY([rows],[cols],[min],[max],[whole])` — thread-local xorshift64 PRNG（依存クレートなし） | +1 件 |
| Task L | `Cells.End` 検索インデックス化 — `BTreeSet` lazy rebuild（sheet 切り替え・cells_mut 呼び出しで dirty、次回 End クエリ時に O(n) rebuild → O(log n) 検索） | — |
| Task M | `Type...End Type` ユーザー定義型: `Program.type_defs`、`Stmt::DimRecord`、型別デフォルト初期化、`Public Type` 対応 | +4 件 |
| Task N | データベース関数追加: `DSUM`/`DAVERAGE`/`DCOUNT`/`DCOUNTA`/`DMAX`/`DMIN`（`db_row_matches_criteria` / `resolve_db_field` を再利用） | +23 件 |
| Task O | 財務関数追加: `FV`/`PV`/`NPER`/`RATE`/`IPMT`/`PPMT`/`NPV`/`IRR`/`MIRR`/`XNPV`/`XIRR`（`annuity_fv` / `compute_pmt` ヘルパーを共有、RATE/IRR/XIRR は Newton-Raphson） | +1 件 |

| Perf Round 4-B | `range_nums_fast!` マクロ: AVERAGE/MIN/MAX の単一 Range 引数で `Vec<Variant>` を省略して直接 `f64` 収集。SUM も同様の直接ループ fast path を追加。 | — |
| Perf Round 4-C | `RangeWrite` / `RangeClear` の dirty フラグ集約: セルごとの `cells_mut()` 呼び出しをやめ、シートマップに直接書き込み後に `cell_index_dirty = true` を1回セット。 | — |
| CLI binary | `src/main.rs` 追加: Python 不要のスタンドアロン CLI バイナリ。`elixcee <vba> <MacroName> [--file xlsx] [--sheet name] [--output xlsx]` | — |
| pyo3 optional | `pyo3` をオプション依存化（`python` feature で有効化）。`#[cfg(feature = "python")]` で Python バインディングを条件コンパイル。maturin は `features = ["python"]` で引き続き動作。 | — |
| GitHub release | `.github/workflows/release.yml`: `bin-v*` タグ push で Windows/Linux/macOS バイナリを GitHub Release に自動アップロード。 | — |

現在のテスト総数: **329 件**

---

## 残タスク（今後の拡張候補）

### パフォーマンス
- [ ] 大規模ループの並列実行（Rayon クレート）— 依存クレート追加が前提

### 品質・テスト
- [x] プロパティベーステスト（`proptest` クレート）
- [x] VBA パーサーのファジングテスト（`cargo-fuzz`）
- [ ] 実 .xlsx / .ods ファイルを使った E2E 統合テスト
- [x] ベンチマーク（`cargo bench`）— 大量ループ・大規模セル書き込みの計測

### VBA 構文（高度）
- [x] ネスト型（`Type A` のフィールドに `Type B` を持つ）
- [x] `Dim arr(10) As MyType`（UDT の配列）
- [x] `With p` ブロック（UDT フィールドへの `.field` アクセス）
