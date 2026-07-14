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
| Milestone A | JSON Agent Contract: `elixcee <file> <macro> --json`。`src/diagnostics.rs`（`ElixceeError`、runtime error のプレフィックス分類、手書き JSON エスケープ／`variant_to_json`）、`Vm.msgbox_log`（MsgBox を stdout 直接出力ではなく `messages` へ集約）、非 `--json` 時の出力は既存動作とバイト単位で同一。serde 非依存（既存のD1〜D3の依存削減方針を踏襲）。 | +13 件 |
| Milestone A.1 | JSON contract hardening。`Vm::msgbox_log` を private 化 + `take_messages()`（drain）+ `run_sub` 先頭でクリア（Vm 使い回し時のリーク防止）、MsgBox は記録してから `error_on_msgbox` で失敗する仕様に固定、`tests/cli_json.rs` を dev-only `serde_json` で構造的に parse する形へ全面刷新（成功/各 stage の失敗/MsgBox/制御文字・引用符・改行・日本語・長大文字列/Excel エラー値/バックスラッシュを含むパスを網羅）、`variant_to_json` に NaN/Infinity ガードを追加、`docs/agent-contract.md` で終了コード・stdout/stderr 契約・エラーコード表・`messages` 仕様を明文化。加えて `on = 0` パーサーバグを修正（`"on"` は後続が `error` の時だけ On Error 文として解析、`"for"`/`"each"` と同じ lookahead パターン）。 | +16 件 |
| Milestone A.5 | Source Location Foundation。`SourceSpan`（char offset）と `SpannedStmt` を AST に追加（`Vec<Stmt>` 12 箇所を `Vec<SpannedStmt>` へ）。トークナイザが並列 spans 配列を出力、`parse_stmts`/`parse_with_body` が各文を span でラップ。既存 `parser::parse`/`Vm::run_sub` のシグネチャは無変更のまま、新規 `parser::parse_with_span`（parse error 用）と `Vm::current_span()`（runtime error 用、`exec_stmt` で毎回更新——ネストした文でも正確）を追加。`diagnostics::locate()`（source 全文を1回線形スキャンして line/column 変換、`check` コマンド登場までインデックス化しない）で `--json` の `error.location` に `{file, line, column}` を付与（`--json` 限定・非JSON出力はバイト単位で不変）。runtime error は文単位の粒度（式単位ではない）。「did you mean」候補は対象外。 | +5 件 |
| Milestone B1 | `check` サブコマンド（静的解析・実行なし）。新規 `src/check.rs`：parse 診断（`parser::parse_with_span` 再利用）、指定 macro の存在確認（`Vm::run_sub` と同じ大文字小文字非依存マッチを静的に再現）、`Stmt::MsgBox` の再帰的 AST 走査による対話操作検出（info severity、新規コード `I1001`）。7項目中この3つのみ実装——欲張らず最小構成から（詳細は下記「Milestone B1.1」参照）。`elixcee check <file> [MacroName] [--json]` を追加、`--json` は `{schema_version, ok, diagnostics: [...]}` という新しい（run モードとは別の）バッチ形式。`docs/agent-contract.md` に `check` 専用セクションを追加。 | +18 件 |
| Milestone B1.1 | `check` に未定義 Sub/Function **呼び出し**検出を追加（entrypoint 名だけでなく本文中の全呼び出し）。当初案（`eval_vba_func`/`eval_wsf` の match arm を手動複製した allowlist）は破棄し、**実際の dispatch table に直接問い合わせる**方式に変更——`src/vm/mod.rs` に `pub fn is_known_builtin_function(name)` を追加、使い捨て `Vm` で `eval_vba_func(name, &[])` を0引数プローブ呼び出しし、`"Unknown VBA function: '"` / `"is not implemented"` のどちらでもない結果なら既知と判定。二次情報源が存在しないため drift のリスクがゼロ（当初懸念していた「VM に関数が増えたら allowlist が古くなり false positive を生む」問題を、ミラーを作らないことで根本的に解消）。`src/check.rs` に `Expr` 全20バリアントを網羅する再帰的 `walk_expr`（wildcard 無し・exhaustive match）と、`Stmt` の Expr フィールドを収集する `collect_stmt_exprs`（同じく exhaustive）を追加——将来 AST にバリアントが増えた際はコンパイルエラーで気づける設計。診断位置は文単位の粒度（式単位の span は存在しないため、Milestone A.5 と同じ制約）。**実装直後にレビューで発見**：`arr(i)` と `func(i)` はこの AST 上で構文的に区別できない（どちらも `Expr::FuncCall` — 専用の配列添字ノードが存在しない）ため、`Split()` の戻り値や `Dim arr(10)` を添字アクセスするだけで「未定義関数」と誤検知する重大な false positive があった。`local_scope_names`（Sub/Function ごとに params・自分自身の名前・本文全体から `Assignment`/`DimArray`/`ReDim`/`For`/`ForEach`/`DimRecord` 等で導入される変数名を収集する exhaustive な `collect_declared_names`）を追加し、`is_resolvable` の判定に「スコープ内変数か」を最優先で含めることで解消。回帰テスト4件を追加（`Split` 添字・`Dim` 配列添字・関数引数の配列添字・配列を使いつつ本物の未定義呼び出しも検出できることの確認）。 | +27 件 |
| Milestone B1.1（後半） | `check` に未対応構文/no-op 検出（`I1002`/`unsupported_construct`）を追加。`src/parser/mod.rs` の no-op fallback 箇所を Explore agent で全数調査した上で分類——「意図的な no-op」（`Option Explicit` 等のモジュール宣言、`Type` ブロック内の余剰行、素の `Dim x`、そして **`Sub` 内の `Static x As Type` も同じ扱い**——`"public"|"private"|"static"|"friend"` の同一 fallback アームを共有しており、テスト `test_static_dim_inside_sub` で意図的な no-op だと確認済み）と「未対応構文」（`Debug.Print`/`Debug.Assert`、`Range`/`Sheets` の未知プロパティ・メソッド、代入を伴わない配列要素/レコードフィールド読み取り、`Call`/括弧なしのベア識別子呼び出し）を明確に区別。後者8箇所のみ `Stmt::Dim` から新設の `Stmt::Unsupported { reason: String }` へ変更（VM 側の実行結果は `Stmt::Dim` と全く同じ no-op のまま——`exec_stmt_inner` に `Stmt::Unsupported {..} => {}` を追加しただけ、既存テストへの影響ゼロを確認済み）。**実装中に発見した独立した既存バグ**：上記8箇所のうち4箇所は `self.skip_to_eol()`（改行トークンも消費する)を使っていたが、呼び出し元が直後にもう一度 `self.eat_eol()?` を呼んでいたため、この4構文（`EntireRow`/`EntireColumn` の未知メソッド、ベア `Sheets.<method>`、代入なしの配列要素読み取り、代入なしのレコードフィールド読み取り）は**後続に他の文があるかどうかに関わらず常に** `"expected newline, got Ident(...)"` という偽の構文エラーでパース自体が失敗していた（最後の文の場合に限らない——`eat_eol()` は次の文の先頭トークンにぶつかって同じ理由で失敗する。fuzzing 未到達の潜在バグ）。個別 lookahead ではなく、既に他の4箇所で使われていた「改行を消費しない」手動スキップ方式に統一して根本修正。モジュールレベルの未対応行（`Program` に文リストが無く、診断を紐づける場所が無い）と `With` ブロック内の同種の no-op（`parse_with_dot_stmt` が現在 `Ok(None)` を返し AST ノード自体を作っていない——`Some` に変えると body の shape が変わり要検証）は今回のスコープ外として明示的に見送り。 | +15 件 |
| Milestone B1.1（残課題） | 上記2つの残課題を解消。(a) モジュールレベル: `Program` に `module_diagnostics: Vec<(String, SourceSpan)>` フィールドを追加し、`parse_program` の2つの skip サイトで診断を収集。**分類の要点**：`Vm::variables` が単一のフラットな `HashMap`（`call_sub_def` はパラメータ名のみ save/restore、それ以外の変数は `Vm` 全体で1つの名前空間を共有）だと確認した上で、素の `Dim`/`Public x`/`Private x`/`Static x`（値を持たない宣言）は Sub 内の場合と同様に無害な no-op として意図的に対象外のまま（フラグを立てると Sub レベルの既存方針と矛盾するノイズになる、と実装前のレビューで指摘されて修正）。一方 `Const FOO = 5` はモジュールレベルでは評価パスが一切存在せず値が本当に失われる実害があるため、修飾子付き（`Public Const` 等）・素の `Const` 両方を `I1002` としてフラグ。(b) With ブロック: `parse_with_dot_stmt` の3箇所の `Ok(None)`（UDT フィールド読み取り、未知の `.property`/`.method`、非識別子ドット文）を `Ok(Some(Stmt::Unsupported{reason}))` に変更。`parse_with_body` の span 収集ロジックはそのまま流用でき、`check.rs` の `collect_declared_names`/`walk_body` は既に `With`/`WithRecord`/`WithSheet` を再帰していたため check.rs 側の変更は一切不要だった。`parse_with` のターゲット未認識フォールバック（`With <非識別子ターゲット>`、文が存在する前に発生するため紐づけ先が無く、かつ極めて稀）のみ引き続きスコープ外として明示的に見送り。 | +12 件 |
| Milestone B2（Phase 1） | 複数 `.bas` モジュール対応。ユーザーとスコープを事前合意——`elixcee.toml` マニフェストは見送り（TOML/serde は実行時依存に一切無く、Phase D1〜D3 でまさにこれを避けるため pest/calamine/rust_xlsxwriter を除去した経緯があるため）、`.cls` クラスモジュールも対象外（既存実装ゼロを確認済み）、`Module.Sub` 修飾エントリポイントは今回実装。`Program` に `module_name: Option<String>` を追加、`Attribute VB_Name = "..."` を実際にキャプチャ（VBA 本来のモジュール命名方式、他の `Attribute` 行は従来通り無視）——無指定時はファイル名の stem にフォールバック。`src/parser/mod.rs` に純粋関数 `resolve_entrypoint`（bare/qualified 両対応、VM 非依存）と `find_cross_module_sub_collisions`/`find_cross_module_func_collisions` を追加。`Vm::run_sub_multi` を新設（`run_sub` は無改変）。**実装前のレビューで設計を根本修正**：当初案「衝突する名前だけフラットマップから除外し ambiguous 集合として追跡する」方式は、実際の VBA セマンティクス（同名 Sub の呼び出しは自モジュール優先で解決され、`Private` はモジュール外から不可視）をそもそも表現できない（どのフラットマージ方式でも「2つのモジューがそれぞれ自分の同名 private ヘルパーを呼ぶ」という正当なケースを壊す）ため撤回し、**衝突は無条件で load 時に拒否**する方式に単純化（`run_sub_multi` は実行前に拒否、`check` は `E1005`/`duplicate_sub_or_function` 診断として報告し他のモジュールの検査は継続）。**副次的な発見**：この設計だと `Module.Sub` 修飾は「衝突の解消」には使えない（衝突がある時点で load 自体が拒否されるため）——修飾が効くのはバレ名でも解決できるが明示性のために使うケースのみ、とテストで確認・ドキュメント化。`check.rs` は `run_check`（既存シグネチャ・既存呼び出し元は無改修）の薄いラッパーとして `run_check_in_project`/`run_check_impl` を追加し、`is_resolvable`/`walk_body`/`walk_expr` に `other_module_names: &HashSet<String>` を1引数追加——他モジュールの Sub/Function 名を渡すことで、モジュールをまたぐ無修飾呼び出しが誤って「未定義」と判定される false positive を防止。**CLI 設計は2段階で修正**：当初案「`check` は `--entry` 省略時は今まで通り位置引数1〜2個」は `elixcee check *.bas`（3ファイル以上・エントリポイント指定なし）で `usage()` に落ちる不具合があり、レビューで指摘されて撤回——最終的に **`check` モードは位置引数を常にすべてファイルとして扱い、エントリポイントは `--entry` でのみ指定**という単純なルールに統一（今セッション内でまだ導入したばかりの機能のため破壊的変更として許容、`tests/cli_check.rs` の該当テストを `--entry` 形式に移行）。run モードは元々 macro name が必須のため曖昧さが無く、「最後の位置引数が macro name・それ以外はファイル」という単純な一般化で済んだ（1ファイル+1macro の既存呼び出しはバイト単位で不変）。マルチモジュール実行時の runtime error の `location` は `null`（`SourceSpan` はどのモジュールの span かを識別する情報を持たない——Milestone A.5 で「1実行につき常に1ファイル」という前提のもと明示的に見送られた `source_id` が必要になるケースだが、今回はそこまで手を広げず単一ファイル実行時の精度はそのまま維持）。 | +23 件 |
| Milestone B3 | 決定的なブラックボックステスト。新規 `tests/blackbox.rs` + `tests/fixtures/blackbox/*.toml`（12 件）——宣言的な `.toml` フィクスチャ（VBA ソース＋CLI 引数＋期待する JSON 全文）を読み、実バイナリの `--json` 標準出力とバイト単位で比較する汎用ハーネス。既存の `tests/cli_json.rs`/`tests/cli_check.rs`（個別 `assert_eq!` による手書きテスト）を置き換えるものではなく追加——新しい回帰ケースは Rust 不要で `.toml` を1つ置くだけで済む形式として並存させる。新規 dev-dependency `toml`（`serde`/derive は追加せず `toml::Value` を `serde_json::Value` と同様にインデックスで読むだけ、既存の `serde_json` dev-only 方針を踏襲）。**非決定的な2箇所を比較前に正規化**：`duration_ms` は JSON 構造として除去、一時ファイルの実パスは（`error.location.file`/`check` 診断の `location.file` だけでなく `E1006` の重複モジュール名メッセージ本文に埋め込まれる場合も含めて）テキストレベルで書き込んだ全パスを `<FILE>` に置換してから JSON パースする方式に統一（当初 JSON キー名 `"file"` 単位での構造的置換を検討したが、`E1006` のメッセージ本文はキーではなく文字列内にパスを埋め込むため対応できず、テキスト置換方式に変更）。フィクスチャごとに1つの `#[test]` を生成する方式（`datatest` 等）は今回見送り——1つの `#[test]` が全フィクスチャを走査し失敗を集約して報告する形にした（フィクスチャ数が今の規模では十分、生成コードの複雑化に見合わない）。**レビューで発見した既知の制約**：`find_cross_module_sub_collisions`/`_func_collisions`（`src/parser/mod.rs`）は `HashMap::into_iter()` で終わるため、1回の実行に**複数の異なる衝突**がある場合はどれが報告されるか（run モード）や `check` の診断順序が prosess-seed 依存で非決定的——今回の全フィクスチャは衝突1件のみに限定することで回避し、この制約をハーネスのドキュメントコメントに明記（将来複数衝突を検証するフィクスチャを追加しないためのガード）。フィクスチャは実バイナリを実際に実行し、出力を確認した上でそのまま `expect_json` に転記——手打ちの期待値ではない。 | +1 件（`all_blackbox_fixtures_match_expected_json`、内部で12フィクスチャを走査） |
| Milestone B4 | Workbook Snapshot（`snapshot` サブコマンド）。実装前に3つの設計論点を `AskUserQuestion` でユーザーと確認（すべて推奨案を選択）：(1) トリガーは独立サブコマンド `elixcee snapshot <file> [--json]`（`check` と同じ「検査のみ・実行しない」立ち位置、run モードの `--json` は無改修）、(2) 安定識別子は XLSX の実 `sheetId` 属性を優先し、無ければ連番フォールバック（`.ods` は常にフォールバック）、(3) 内容は各シートの非空セル（address+value）のみの最小スコープ。`src/reader.rs`: `WorkbookSheet` に `sheet_id: Option<String>` を追加、`xlsx_workbook_sheets` で `sheetId` 属性を捕捉（`.ods` は元々この概念が無く常に `None`）。新規 `src/snapshot.rs`（`pub mod` として `lib.rs` に追加——`check`/`reader` 等と同じ扱いで `cargo test --lib` から独立にテスト可能）：`to_json`/`to_markdown` を実装。**実装中にユーザーからのレビューで命名を修正**：ロードマップ原文は識別子を `code_name` と呼んでいたが、実装直後にユーザーから「`code_name` は VBA 開発者が本物の `CodeName`（`vbaProject.bin` という別のバイナリ OLE 形式に保存される、VBA IDE 側で割り当てられる識別子で、このリーダーは一切パースしない）を連想してしまい紛らわしい」との指摘を受け、フィールド名を `sheet_id`（生の XLSX 属性、null 許容）と `stable_id`（常に存在する計算済み識別子）に分離——`code_name`/`vba_code_name` という名前は将来本物の VBA CodeName を実装する際のために予約し、衝突を避けた。JSON は両方公開（`stable_id` が実ファイル由来か合成フォールバックかを消費者が判別できるように）、Markdown は `stable_id` のみ表示（Markdown は簡略化された表示専用ビューという既存方針を踏襲）。`elixcee snapshot` は `run_check_command` と同じパターンで `main.rs` に追加、失敗時は既存の `ElixceeError::io_error`/`fail_json`/`die` をそのまま再利用（`messages` は snapshot がマクロを一切実行しないため常に空）。テストは B3 の教訓（8 vs 9 カラムの実例）を踏襲し、`tests/cli_snapshot.rs` の期待値は実バイナリを `--output` で実行して得た実データから転記（writer がシート名を小文字化すること、`Sheets.Add` の並び順がアルファベット順であることなど、直感では外しやすい変換を実測で確認済み）。**レビューで発見**：`.xlsx` のみで end-to-end 検証しており、`.ods`（`sheet_id: null` → 合成 `stable_id`、この機能が存在する本質的な理由）が単体テストでしか確認されていなかったため、実バイナリで `.ods` を実際に生成・snapshot する統合テストを追加。あわせて `stable_id` の一意性は「準拠したファイル」でのみ成立する（実 `sheetId` と欠落ケースが1ファイル内に混在する非準拠 `.xlsx` では衝突しうる——起こり得ないケースなので検出はしない）旨を doc に明記。 | +11 件（`reader.rs` 3 件 + `snapshot.rs` 8 件）+ `tests/cli_snapshot.rs` 7 件 |

現在のテスト総数: **436 件**（lib）+ `tests/cli_json.rs` 14 件 + `tests/cli_check.rs` 15 件 + `tests/blackbox.rs` 1 件（内部で `tests/fixtures/blackbox/*.toml` 12 件を走査）+ `tests/cli_snapshot.rs` 7 件

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

---

## Milestone A.1（JSON contract hardening）— 完了

Milestone A のレビューで指摘された残課題。設計のやり直しは不要、地固めのみ。

- [x] 全 JSON 出力を `serde_json`（**dev-dependency のみ**、リリースバイナリの依存には影響しない）で構造的に parse して検証するテストへ強化した（`tests/cli_json.rs` を全面刷新。成功/各 stage の失敗/MsgBox/制御文字・引用符・改行・日本語・空文字列・長大文字列/Excel エラー値/バックスラッシュを含むパスを網羅、計 11 件）
- [x] `Vm::run_sub` の先頭で `msgbox_log` をクリアし、同一 `Vm` を使い回した場合に前回実行の MsgBox メッセージが混入しないようにした。`msgbox_log` は private 化し、`take_messages()`（drain）でのみ読めるようにした。合わせて以下をテスト（`src/vm/mod.rs`）：
  - [x] 2 回目の `run_sub` に 1 回目の `messages` が混ざらない（`test_msgbox_log_does_not_leak_across_runs`）
  - [x] MsgBox の後に runtime error が起きても、そのメッセージは `messages` に残る（`test_msgbox_log_survives_a_later_runtime_error`）
  - [x] `error_on_msgbox = true` のとき「記録してから失敗する」仕様に決定してテストで固定した（`test_msgbox_blocked_is_recorded_before_failing`）— メッセージが失敗の唯一の手がかりになるケースを考慮
  - [x] 非 `--json` 実行時の MsgBox 表示順序が従来と変わらないことを確認（`tests/cli_json.rs::non_json_output_is_unchanged`）
- [x] 終了コード契約（0=成功 / 1=失敗の粗い2値。JSON の `ok`/`error.code` が主な機械可読シグナル）と `--json` 時の stdout/stderr 契約を `docs/agent-contract.md` に明文化した（エラーコード表、`messages` 仕様込み）
- [ ] runtime error の分類を文字列プレフィックス一致（`src/diagnostics.rs::classify_runtime_error`）から VM 側の型付きエラーへ移行する足場を作る。**今回は見送り**（型だけ先に追加すると中途半端な二重経路になるため、次に診断種類を増やす直前が適切）。イメージ：
  ```rust
  pub enum RuntimeErrorKind {
      UndefinedVariable { name: String },
      UndefinedSubOrFunction { name: String },
      SheetNotFound { name: String },
      MsgBoxBlocked,
      UnsupportedFeature { feature: String },
      StepLimitExceeded,
      Other,
  }
  pub struct RuntimeError { pub kind: RuntimeErrorKind, pub message: String, pub span: Option<SourceSpan> }
  ```
  CLI の `ElixceeError` はこれを JSON へ変換するだけにする（人間向けメッセージを機械向け分類の入力にしない）。
- [ ] （要検討・要相談・**今回は見送り**）JSON envelope を成功/失敗で統一する案（`schema_version`/`ok`/`command`/`result`/`messages`/`warnings`/`error` を常時同居させ、未使用時は `null`）。`command` フィールドはサブコマンドが増えてから意味を持つため、B1（`check` コマンド追加）のタイミングで合わせて検討する。現状は成功時と失敗時でトップレベル形状が異なる設計を採用済み。
- [x] 決定的な出力順序 — `cells` は既に (row, col) でソート済み、`messages` は Vec の挿入順で決定的。現状これ以上の対応は不要（`warnings`/変数一覧などを JSON へ追加する際は都度ソートすること）。

### 修正済みの既知バグ
- [x] `tests/prop_tests.rs::prop_vba_assignment_parses` が変数名 `on`（例: `on = 0`）で失敗していた件を修正した。`src/parser/mod.rs:374` の文頭ディスパッチを `"on" if self.is_ident_at(1, "error")` に変更（`"for" if is_ident_at(1, "each")` と同じ lookahead パターン）。`#[ignore]` にはせず根本修正。既存の `On Error` 系テスト（parser/vm 双方）も引き続き通過を確認済み。
- [x] 上と同じバグクラスの姉妹インスタンスを `do`（`do = 0`）で発見（Milestone B1 作業中、proptest が新しいランダム値でヒット。`git stash` で本セッション由来ではなく既存バグだと確認済み）。今回はキーワード1つずつの lookahead 追加ではなく、**根本原因を一箇所で潰す**方式にした：`parse_stmt` の分岐に入る前に「先頭識別子の直後のトークンが `=` なら常に代入文として扱う」というガードを追加（`src/parser/mod.rs`、`parse_stmt` 冒頭）。VBA の文法上、どの文キーワードも自分の直後に `=` が来ることはない（`Dim`/`Const`/`For` 等は必ず名前や式が続く）ため、この判定は `do`/`select`/`with`/`call`/`range` 等、`"on"` 以外の約20個のキーワード全てで同時に安全に効く。個別 lookahead を今後も追加し続ける「もぐら叩き」を避けられた。プロパティテストを5回連続実行して回帰なしを確認済み。

### 新たに判明した既知の技術的負債（今回未着手・要方針決定）
- [ ] リポジトリ全体が `cargo fmt --all -- --check` を通らない（`git stash` で master でも同様に失敗すると確認済み＝今回のパッチ由来ではない）。差分は約 985 箇所、`benches/vm_bench.rs` と `src/formula/eval.rs` 等 pre-existing なコードに集中。今回新規作成した `src/diagnostics.rs` / `tests/cli_json.rs` は `rustfmt` 済みでクリーン。リポジトリ全体を一括 `cargo fmt --all` すると巨大な無関係diffになるため、方針（一括フォーマットする専用コミットを別途作るか、現状維持か）を決めてから着手すること。
- [ ] リポジトリ全体が `cargo clippy --all-targets --all-features -- -D warnings` を通らない（同じく `git stash` で master でも失敗すると確認済み）。67 件、内訳は `src/formula/eval.rs` 48 件・`src/vm/mod.rs` 14 件・`src/reader.rs` 8 件・`src/lib.rs` 7 件・`src/parser/mod.rs` / `src/formula/parser.rs` 各2件（今回の変更行とは無関係な既存コードのみ、行番号を照合して確認済み）。特に `src/lib.rs:24` の pyo3 `#[pyclass]` + `Clone` 由来の非推奨警告は pyo3 側 API 変更への追従が必要で、単純な自動修正では終わらない可能性がある。

## Milestone A.5（Source Location Foundation）— 完了

複数モジュール・プロジェクト形式（Milestone B）を追加する前に、position 情報を通しておくと後続改修量が減る。

- [x] トークナイザに offset を保持させ、`SourceSpan { start: u32, end: u32 }` を AST ノードへ通した。**byte offset ではなく char offset**（当初案からの変更 — 手書きトークナイザが既に `Vec<char>` を走査しており、CJK では列はバイト数より文字数の方が意味を持つため。`SourceSpan { source_id, ... }` の `source_id` も見送り——現状ファイルは1回の実行につき常に1つなので時期尚早）。`Vec<Stmt>` 12 箇所を `SpannedStmt` でラップした `Vec<SpannedStmt>` へ変更。
- [x] `SourceSpan` → `SourceLocation { file, line, column }` へ変換する `diagnostics::locate()`（`end_line`/`end_column` は見送り。ソース全文を1回線形スキャンする方式——1実行で報告するエラーは高々1件なので、行オフセットの事前インデックス化は不要と判断。バッチで多数の診断を返す `check` コマンド登場時に検討）
- [x] parse error（`parser::parse_with_span`）と runtime error（`Vm::current_span()`、`exec_stmt` で毎回更新——ネスト文でも正確）が同じ `SourceSpan`/`locate()` を共有する形にした。**runtime error は文単位の粒度**（式単位ではない——`x = totla + 1` の未定義変数エラーは文の先頭 `x` の位置を指す）。`check`/lint/trace/未対応API警告/refactor解析への適用は Milestone B 以降。
- [x] 「did you mean」候補（近い変数名のサジェスト）は対象外のまま——シンボルテーブルが必要な別課題として `tasks/todo.md` に記録済み。

## Milestone B（優先順・xlflow / xlsm_devkit 由来のアイデアを含む）

- [x] B1: `check` コマンド — parse 診断・指定 macro の存在確認・MsgBox 等の対話操作検出・source location を実装済み（詳細は上表）。未対応 VBA 構文・no-op Excel API・完全未対応 Excel API 検出は B1.1 へ。
- [x] B1.1（前半）: 未定義 Sub/Function **呼び出し**検出（entrypoint 名だけでなく、本文中の呼び出し全て）— 実装済み（詳細は上表）。当初懸念していた allowlist 手動複製の drift/false-positive リスクは、`Vm::is_known_builtin_function` による実 dispatch table への直接プローブ方式で根本解消した（ミラーとなる第二の情報源を作らない設計）。
- [x] B1.1（後半）: 未対応 VBA 構文 / no-op Excel API 検出 — 実装済み（詳細は上表、新規コード `I1002`）。当初 `Static` 宣言も対象と見ていたが、調査の結果 `Sub` 内 `Static x As Type` は素の `Dim x` と同じ「意図的な no-op」（Task A で意図的に追加された挙動）と判明し、対象外のまま維持。完全未対応 Excel API はこの `I1002` 検出に統合済み（「no-op」と「完全未対応」は AST 上・チェックロジック上まったく同じ形なので、reason 文字列の違いのみでコード/kind は分けなかった）。モジュールレベルの未対応行検出と `With` ブロック内の同種 no-op 検出は別途 B1.1 の残課題として維持（下記参照）。
- [x] B1.1（残課題）: (a) モジュールレベルの未対応行検出 — `Program` に `module_diagnostics` フィールドを追加し実装済み（詳細は上表）。ただし全ての module-level skip を機械的にフラグしたわけではない：`Vm::variables` が単一のフラットな namespace だと確認した上で、値を持たない素の宣言（`Dim`/`Public x`/`Private x`/`Static x`）は Sub レベルの既存方針と揃えて対象外のまま維持し、値が本当に失われる `Const` のみをフラグする設計にした（実装前のレビューで「全部フラグするとノイズになる」と指摘されて修正）。 (b) `With` ブロック内の同種 no-op 検出 — `parse_with_dot_stmt` の3箇所を実装済み（詳細は上表）。`check.rs` 側の変更は不要だった（既存の再帰が新しい `Stmt::Unsupported` ノードを自動的に拾う）。`parse_with` のターゲット未認識フォールバック（`With <非識別子ターゲット>`）のみ、文が存在する前に発生し紐づけ先が無いという構造的理由でスコープ外として維持——AST 上の頻出度も極めて低い。
- [x] B2（Phase 1）: 複数 `.bas` モジュール対応・`Module.Sub` エントリポイントを実装済み（詳細は上表）。以下は明示的に Phase 2 以降へ見送り：(a) `elixcee.toml` マニフェスト（現状は複数ファイルを CLI 引数に直接列挙）、(b) `.cls` クラスモジュール対応（`Property Get/Let/Set`・インスタンス化などパーサー未実装、multi-file 対応より大きい別課題）、(c) プロジェクトに「宣言された」fixture workbook（現状は既存の `--file` フラグで代用——マニフェストが無いため「宣言」を持たせる仕組みが無い）、(d) VBA 本来の own-module-first/`Private` スコープ解決（現状は衝突を無条件拒否するのみで、正しいスコープ解決はしていない）、(e) `Type` の cross-module 名前衝突検出（Sub/Function と異なり検出なし・サイレントに last-wins——cross-module UDT は稀なため許容）、(f) `Module.Sub` 修飾はマルチファイル時のみ意味を持つ（単一ファイル実行/check は従来通りバレ名のみで解決、単一ファイル時の後方互換を壊さないための意図的な制約）。
- [x] B3: 決定的なブラックボックステスト（固定期待値 toml + JSON 結果）— 実装済み（詳細は上表）。既存の `tests/cli_json.rs`/`tests/cli_check.rs` は置き換えず、追加の回帰テスト層として並存。
- [x] B4: Workbook Snapshot（`snapshot` サブコマンド）— 実装済み（詳細は上表）。ロードマップ原文の `code_name` という命名は実装中にユーザー指摘を受けて `sheet_id`（生の XLSX 属性）/`stable_id`（常時存在する計算済み識別子）へ変更——VBA の本物の `CodeName` との混同を避けるため。
- [ ] B5: property-based な workbook テスト（境界値・ランダム入力で `no_excel_errors` を検証。既存の `proptest` dev-dependency と相性が良い）
- [ ] B6: `trace` / `diff` コマンド、行列の insert/delete/move に伴う VBA 参照の refactor 影響解析
