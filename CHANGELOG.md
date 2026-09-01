# Changelog

重要な変更を [Keep a Changelog](https://keepachangelog.com/en/1.0.0/) 形式で記録します。

## [0.94.0] - 2026-09-01

- XML要素のタグ名と属性構文を厳格に検証し、属性の`=`欠落、未クォート値、未終了クォート、タグ名欠落を拒否するようにしました。
- 不正な属性を部分的に解釈して成功扱いする経路を閉じ、通常reader・stream reader共通の入力エラー契約を強化しました。

## [0.93.0] - 2026-08-31

- 未終了のコメント、CDATA、終了タグ、属性付きタグをXML入力エラーとして拒否するようにしました。
- 軽量XML反復器が壊れた構文を末尾まで黙って消費せず、通常reader・stream reader共通で部分的な成功を防ぎます。

## [0.92.0] - 2026-08-31

- XML予算検証でルート要素を必須・一意とし、ルート外の非空テキストを拒否するようにしました。
- 複数ルートやルートなしのXMLを部分的な入力として処理せず、通常reader・stream reader共通で入力エラーにします。

## [0.91.0] - 2026-08-31

- XMLの開始タグと終了タグの名前一致をXML予算検証で確認するようにしました。
- 要素の深さだけが一致する壊れたXMLを成功扱いせず、通常reader・stream reader共通で入力エラーにします。

## [0.90.0] - 2026-08-31

- XML要素内の重複属性を通常reader・stream reader共通のXML予算検証で拒否するようにしました。
- 同名属性の先勝ち・後勝ち解釈による入力の曖昧性をなくし、XLSX/XML入力を明示的に失敗させます。

## [0.89.0] - 2026-08-31

- 31文字を超えるworksheet名と、Excelで禁止される`: \\ / ? * [ ]`文字を通常reader・stream readerの双方で拒否するようにしました。
- 無効なworksheet名を曖昧な参照・保存対象として受け入れず、入力エラーとして扱います。

## [0.88.0] - 2026-08-31

- 空のworksheet名・relationship ID、および0/非数値の`sheetId`を通常reader・stream readerの双方で拒否するようにしました。
- Workbook識別子の不正値を既定値へ変換せず、明示的な入力エラーとして扱います。

## [0.87.0] - 2026-08-31

- 必須属性が欠けたworksheet要素や、自己終了形式でないworksheet要素を通常reader・stream readerで拒否するようにしました。
- 不完全な`workbook.xml`のsheet要素を黙って無視せず、部分的なWorkbook成功を防ぎます。

## [0.86.0] - 2026-08-31

- 重複`sheetId`と未知のworksheet `state`値を通常reader・stream readerの双方で拒否するようにしました。
- `visible`、`hidden`、`veryHidden`以外の状態を黙ってvisibleへ変換せず、Workbookメタデータの解釈を明示的に失敗させます。

## [0.85.0] - 2026-08-31

- Workbook内のworksheet名重複（大文字小文字違いを含む）と`r:id`重複を通常reader・stream readerで拒否するようにしました。
- 曖昧なシート名・参照先の後勝ち解決を防ぎ、Workbookメタデータの解釈を決定的にしました。

## [0.84.0] - 2026-08-31

- worksheet relationshipの重複`Id`を通常reader・stream readerの双方で拒否するようにしました。
- relationshipの後勝ち解決を廃止し、同じworksheet参照が入力順に依存しないようにしました。

## [0.83.0] - 2026-08-31

- worksheet relationshipの`TargetMode="External"`を通常reader・stream readerの双方で拒否するようにしました。
- worksheetを外部URLや外部参照として解釈せず、内部ZIP partだけを入力対象にする境界を統一しました。

## [0.82.0] - 2026-08-31

- worksheet relationship targetの正規化とZIPルート脱出拒否を通常readerにも適用しました。
- 通常readerとstream readerが同じ安全なrelationship解決境界を共有し、異常な相対targetを部分成功として扱わないようにしました。

## [0.81.0] - 2026-08-31

- 通常のXLSX readerが、worksheet relationshipまたはworksheet partの欠落をシートの黙ったスキップとして扱わず、明示的な入力エラーを返すようにしました。
- 欠損したシートを含むWorkbookの部分的な成功を防ぎ、通常readerとstream readerの失敗契約を統一しました。

## [0.80.0] - 2026-08-31

- 通常のXLSX readerが、存在する`styles.xml`やworksheet XMLの読み込みエラーを既定値やシート欠落へ変換せず、明示的な入力エラーとして返すようにしました。
- 欠落している任意partだけは従来どおり互換フォールバックし、部分的なWorkbook成功を防止します。

## [0.79.0] - 2026-08-31

- 通常のXLSX readerにも共有文字列インデックスの範囲検証を適用し、stream readerとの入力整合性を統一しました。
- 存在する`sharedStrings.xml`の読み込み・検証エラーを空テーブルとして握りつぶさないようにしました。

## [0.78.0] - 2026-08-31

- ストリーミングreaderが不正な共有文字列インデックスを空セルへ変換せず、明示的な入力エラーとして返すようにしました。
- 共有文字列参照の範囲検証を追加し、通常readerと同様に不完全な入力を成功扱いしないようにしました。

## [0.77.0] - 2026-08-31

- ストリーミングreaderにも共有文字列テーブルの件数・総サイズ上限を適用し、通常readerと同じ資源制限で検証するようにしました。
- 上限超過時に共有文字列を部分的に処理せず、明示的なエラーとして返します。

## [0.76.0] - 2026-08-31

- ストリーミングreaderが、存在する`sharedStrings.xml`の破損・上限超過を空テーブルとして握りつぶさず、明示的なエラーとして返すようにしました。
- `sharedStrings.xml`が本当に存在しない場合だけ、従来どおり空の共有文字列テーブルとして扱います。

## [0.75.0] - 2026-08-31

- ストリーミングreaderが不正UTF-8の行を黙って破棄せず、明示的なエラーとして返すようにしました。
- 終端のないworksheet rowを成功扱いにしない回帰防御を追加しました。

## [0.74.0] - 2026-08-31

- ストリーミングwriterが保存失敗時に保留行を失わないようにし、同じwriterで安全に再試行できるようにしました。
- 保存成功後にだけ保留行・pending bytesを解放する回帰テストを追加しました。

## [0.73.0] - 2026-08-31

- ストリーミングwriterのネスト値サイズ見積もりを飽和加算に変更し、整数オーバーフローによるpending-byte制限の回避を防止しました。
- 行単位のサイズ集計にも同じ保護を適用し、回帰テストを追加しました。

## [0.72.0] - 2026-08-31

- CIに固定版`cargo-deny 0.19.9`の依存監査を追加し、advisory、license、source、duplicate dependencyの検査を継続的に実行するようにしました。
- ローカルの`cargo deny check --disable-fetch`とCIの依存ポリシーを同じ`deny.toml`で検証するようにしました。

## [0.71.0] - 2026-08-31

- Rustワークスペースのclippy警告を解消し、CIの`-D warnings`を実装コード全体へ適用できる状態にしました。
- `cargo-deny`のlicense/sourceポリシーを追加し、依存監査をローカルでも再現できるようにしました。

## [0.67.0] - 2026-08-31

- v0.66で導入した未対応入力拡張子の決定的エラーを、VMのWorkbook読み込みAPIでもそのまま伝播するようにしました。
- 拒否された入力パスを上位APIのエラーメッセージへ再露出しない回帰テストを追加しました。

## [0.68.0] - 2026-08-31

- 実装済みのZIP/XML/Workbook/VBA各種上限をsecurity modelの現行契約へ同期しました。
- 未実装の全体read work budget、defined-name数上限、reader cancellation budgetだけを残課題として明示しました。

## [0.69.0] - 2026-08-31

- Workbookのdefined-name一覧に10万件の上限を追加し、超過時は部分的な一覧を返さず明示的に拒否するようにしました。
- defined-name上限をsecurity modelとresource limits文書へ反映しました。

## [0.70.0] - 2026-08-31

- defined-nameの式文字列を1MiBまでに制限し、超過時は部分結果を返さず明示的に拒否するようにしました。
- defined-nameの件数・式文字列長の上限をセキュリティ文書へ反映しました。

## [0.66.0] - 2026-08-31

- パス指定のWorkbook readerが、未対応または拡張子なしの入力をファイルアクセス前に決定的なエラーとして拒否するようにしました。
- `.xlsx`・`.xlsm`・`.ods`の大文字小文字を区別しない入力形式境界を回帰テストで固定しました。

## [0.48.0] - 2026-08-31

- ZIP入力にentry数、entryごとの展開後サイズ、総展開サイズ、圧縮率の上限を適用しました。
- ZIPエントリ名の絶対パス、親ディレクトリ参照、NUL文字を拒否する回帰テストを追加しました。
- `docs/xlsx-security-model.md`と`docs/limits.md`を実装済みの入力境界に同期しました。
- XMLの要素数、属性数、属性値、テキストノード、入れ子深度に上限を追加しました。
- DTD/ENTITY宣言と不完全なXML文書を拒否する回帰テストを追加しました。

## [0.49.0] - 2026-08-31

- Workbookのシート数、シートごとのセル数・結合範囲数、共有文字列の件数・総サイズに上限を追加しました。
- 共有モデル上限の拒否条件を純粋な検証関数として回帰テストに追加しました。

## [0.50.0] - 2026-08-31

- 数式パーサーに入力長、参照数、ASTノード数、ネスト深度の上限を追加しました。
- 深い括弧・関数ネストと過大な数式入力を安全なパースエラーにしました。

## [0.51.0] - 2026-08-31

- VBA parserにソース長、識別子長、トークン数の上限を追加しました。
- 過大なVBAソースをAST構築前に決定的なパースエラーとして拒否するテストを追加しました。

## [0.52.0] - 2026-08-31

- VBA VMに決定的な命令数上限を追加し、無限ループや過大な実行をwall-clock設定なしでも停止できるようにしました。
- 命令数上限の超過を`BUDGET:`エラーとして返す回帰テストを追加しました。

## [0.53.0] - 2026-08-31

- VBA Sub/Functionの再帰・ネスト呼び出し深度にデフォルト上限を追加しました。
- 呼び出し深度超過を`BUDGET:`エラーとして安全に停止する回帰テストを追加しました。

## [0.54.0] - 2026-08-31

- VBA VMに文字列サイズと配列要素数のデフォルトbudgetを追加しました。
- 上限超過を`BUDGET:`エラーとして安全に停止する回帰テストを追加しました。

## [0.55.0] - 2026-08-31

- Pythonの`Vm.set_budgets()`から、VBAの命令数・呼び出し深度・文字列・配列budgetを個別に設定できるようにしました。
- 既定値は変更せず、明示的な`None`で個別budgetを無制限にできるAPI契約を追加しました。

## [0.56.0] - 2026-08-31

- VBA実行中にワークブック全体へ保持できるセル数のデフォルトbudgetを追加しました。
- `Vm.set_budgets()`からセル数上限も設定でき、超過は`BUDGET:`エラーとして停止します。

## [0.57.0] - 2026-08-31

- `check`でShell、COM/object creation、WScript、ファイル操作相当の外部作用を`E1010`として明示的に拒否するようにしました。
- 既定VMが外部作用を実行しないことを、unsupported情報と区別したエラー契約として文書化しました。

## [0.58.0] - 2026-08-31

- 既定VMでShell、COM/object creation、WScript、ファイル操作相当を実行時にも`SECURITY:`エラーとして拒否するようにしました。
- `On Error Resume Next`で外部作用の遮断エラーを握り潰せないことを回帰テストで固定しました。

## [0.65.0] - 2026-08-31

- 保存APIが`.xlsx`・`.xlsm`・`.ods`以外の拡張子へ誤ってXLSXを書き出さないよう、明示的に拒否するようにしました。
- 未対応拡張子ではファイルを作成しない回帰テストを追加しました。

## [0.64.0] - 2026-08-31

- 既存出力ファイルのpermissionを原子的保存後も保全するようにしました。
- read-only出力をrenameで迂回して上書きせず、明示的に拒否するようにしました。

## [0.63.0] - 2026-08-31

- XLSX/ODSを同一ディレクトリの一時ファイルへ完全に書き込んでから公開する原子的保存に変更しました。
- 保存途中の失敗で既存の出力成果物が部分的に上書きされないようにしました。

## [0.62.0] - 2026-08-31

- XLSX/ODSの保存先パス途中にあるsymbolic linkも検出し、リンク先への意図しない書き込みを拒否するようにしました。
- 保存先と親ディレクトリのsymlink拒否、およびリンク先データ保全を確認する回帰テストを追加しました。

## [0.61.0] - 2026-08-31

- XLSX/ODSの保存先が既存のsymbolic linkの場合、リンク先を上書きせず明示的に拒否するようにしました。
- 保存先symlinkの拒否とリンク先データ保全を確認する回帰テストを追加しました。

## [0.60.0] - 2026-08-31

- `sheet_to_html`のリンクURL判定で前後空白、制御文字、バックスラッシュを拒否するようにしました。
- ブラウザのURL正規化によるscheme/host解釈の揺れを避ける回帰テストを追加しました。

## [0.59.0] - 2026-08-31

- `sheet_to_html`の`cell.h`をデフォルトでescapeし、未信頼値がraw HTMLとして描画されないようにしました。
- 互換性が必要な利用者向けに、`rawHtml: true`の明示opt-inを追加しました。

## [0.46.0] - 2026-08-31

- `Vm.snapshot()`に`hidden_rows`/`hidden_columns`を追加し、行列の非表示状態を取得できるようにしました。

## [0.45.0] - 2026-08-31

- `Vm.snapshot()`に`merged_ranges`を追加し、結合セル範囲をA1記法で取得できるようにしました。

## [0.44.0] - 2026-08-31

- `Vm.snapshot()`に`sheet_states`を追加し、シートの表示状態を取得できるようにしました。

## [0.43.0] - 2026-08-31

- `Vm.snapshot()`に`calculation_mode`を追加し、Automatic/Manual計算状態を取得できるようにしました。

## [0.42.0] - 2026-08-31

- `Vm.snapshot()`に`defined_names`を追加し、名前定義を含むワークブック状態を取得できるようにしました。

## [0.41.0] - 2026-08-31

- `Vm.snapshot()`に`sheet_order`を追加し、ワークシートのタブ順を保持できるようにしました。

## [0.40.0] - 2026-08-31

- `Vm.snapshot(include_formulas=True)`で、セルの計算結果と保存数式を分離して取得できるようにしました。

## [0.39.0] - 2026-08-31

- `diagnose_macro()`を追加し、Pythonから構造化されたVBA診断JSONを取得できるようにしました。

## [0.38.0] - 2026-08-31

- `Vm.snapshot()`を追加し、VMの全シートを独立したPython辞書として取得できるようにしました。

## [0.37.0] - 2026-08-31

- `Vm.fork()`を追加し、ワークブックとVBA実行状態を独立コピーしてバッチ処理に利用できるようにしました。

## [0.36.0] - 2026-08-31

- Pythonの同一`Vm`で同じVBAソースを再実行する際、解析済みASTを再利用するようにしました。

## [0.35.0] - 2026-08-31

- `Vm`/`run_macro`に`timeout_ms`を追加し、PythonからVBA実行期限を設定できるようにしました。

## [0.34.0] - 2026-08-31

- `open_stream`/`StreamReader`に`timeout_ms`を追加し、次の行を待つ時間を制限できるようにしました。

## [0.33.0] - 2026-08-31

- `create_stream`/`StreamWriter`に`max_columns`を追加し、各保留行の列数上限を設定できるようにしました。

## [0.32.0] - 2026-08-31

- `create_stream`/`StreamWriter`に`max_rows`を追加し、保留行数の上限を設定できるようにしました。

- `open_stream`/`StreamReader`に`max_columns`を追加し、1行の列数上限を入力ごとに設定できるようにしました。
- `open_stream`/`StreamReader`に`max_row_bytes`を追加し、1行のXMLバッファ上限を入力ごとに設定できるようにしました。
- `open_stream`/`StreamReader`に`max_rows`を追加し、読み取り行数を上限設定できるようにしました。
- StreamReaderとStreamWriterに明示的な`close()`/`closed`ライフサイクルAPIを追加しました。
- StreamingReaderの契約を安定化し、`include_row_numbers=True`で疎なワークシートの
  元のExcel行番号を返すようにしました。
- 空行・空セル行、worksheet relationship targetの正規化、行バッファ16 MiB上限を追加しました。
- StreamWriterの保留行にも64 MiB上限を設け、超過時は`MemoryError`を返すようにしました。
- StreamWriterに`row_count`、`pending_bytes`、`max_pending_bytes`の読み取り専用状態情報を追加しました。
- `create_stream`/`StreamWriter`で`max_pending_bytes`を指定し、環境ごとのメモリ予算を調整できるようにしました。

## [0.22.0] - 2026-08-30

- Streaming workbook row readerの行メモリ使用量を制限。
- 依存関係とPython/Cargoパッケージのバージョンを更新。

## [0.21.1] - 2026-08-30

- streaming readerのmaturinビルドを修復。

## [0.21.0] - 2026-08-30

- streaming readerの行境界処理を強化。

## [0.20.0] - 2026-08-30

- 大きなワークブックを行単位で読むstreaming APIを追加。
- Python APIとreaderのストリーミング経路を追加。

## [0.19.0] - 2026-08-30

- spreadsheet ZIP入力のセキュリティ対策を強化。
- エントリサイズ、アーカイブ構造、展開処理の検証を追加。

## [0.18.0] - 2026-08-30

- JavaScript/WASM XLSX writerでセルコメントをlegacy Notesとして保存。

## [0.17.0] - 2026-08-30

- JavaScript/WASM XLSX writerで外部・内部ハイパーリンクを保存。
- worksheet relationshipとtooltipの回帰テストを追加。

## [0.16.0] - 2026-08-30

- テーブル、AutoFilter、データ検証の読み取り・編集APIを追加。
- conditional formatting、styled empty cells、default worksheet stylesの保存を修正。

## [0.15.0] - 2026-08-30

- 安全なstyle編集APIを追加（number format、font、fill、border、alignment、protection、
  row/column style、style copy）。
- 既存の共有styleを直接変更せず、style recordを重複排除。

## [0.14.0] - 2026-08-30

- 行列挿入削除、sheet rename、range moveで数式・結合・非表示状態・styleを追従。
- formula cellとAutoFilter metadataの保存を改善。

## [0.12.0] - 2026-08-27

- Pythonのbulk range/row API、sheet管理、merge、hidden row/column、copy sheet、
  defined names、sheet state、row/column size APIを追加。

## [0.11.0] - 2026-08-26

- Python workbook APIとCLIのworksheet操作を拡張。
- 数式・VBA・XLSX round-tripの回帰テストを追加。

## [0.10.1] - 2026-08-24

- XLSX namespaceの`r:` prefix処理を修正し、maturin wheelのround-tripを修復。

## [0.10.0] - 2026-08-24

- worksheet/workbook metadata、freeze panes、selection、style、mergeの保存を改善。
- `.xlsm`のVBA projectと未知のOOXML partを保持するround-trip基盤を追加。

## [0.9.0] - 2026-08-22

- 実Excel作成fixtureによるXLSX round-trip検証を追加。
- VBA project、relationship、content typeの保存を改善。

## [0.1.0]–[0.8.0]

- Rust VBA parser/VM、数式エンジン、Python API、CLI、複数シート、JSON診断、
  property-based workbook test、JavaScript/WASM基盤を段階的に追加。
- 詳細な互換性・セキュリティ方針は [docs/](docs/) を参照。
