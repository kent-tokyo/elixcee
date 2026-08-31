# Changelog

重要な変更を [Keep a Changelog](https://keepachangelog.com/en/1.0.0/) 形式で記録します。

## [0.31.0] - 2026-08-31

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
