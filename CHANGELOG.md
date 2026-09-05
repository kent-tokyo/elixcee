# Changelog

重要な変更を [Keep a Changelog](https://keepachangelog.com/en/1.0.0/) 形式で記録します。

## [Unreleased]

次の開発分はここに記録します。

## [1.0.2] - 2026-09-06

性能値は各リンク先の測定時点（1.0.1表記の開発ツリー）の記録です。公開済み1.0.1との比較や、1.0.2タグそのものの再測定と混同しないでください。

- JSの`sheet_to_html`が既定で`cell.h`を捨てずにエスケープ表示するよう修正しました。生HTMLの出力は引き続き明示的な`rawHtml: true`に限定します。JSパッケージはprivateのままです。
- CLI契約・ロードマップ・各言語README・互換性／資源／配布境界の文書を整理し、0.x変更履歴を別文書へ移動しました。過去の測定条件と未達結果は保持しています。
- 数式校正バイナリもCargo配布対象から除外し、通常のvタグでCLI配布物を作成するよう公開workflowを統一しました。
- 大規模XLSX向けに、保存時の行別tree／配列を単一の座標sortへ変更し、セル番地をスタック上で生成するようにしました。Readerでは属性配列とセル型の作業領域を再利用し、タグ境界・重複属性・禁止宣言の検査コストを削減しました。直前の実装に対する20組の交互測定で、10万／40万／100万セル（25万×4シート）の全体中央値は1.180／1.318／1.196倍でした。厳密な1.2倍目標の達成は40万セルのみです。全ZIP部品の展開後バイト一致、全ターゲットtest／全feature clippyを確認し、保存耐久性・入力制限・versionを維持しました。生データと小規模／文字列混在の回帰測定は `docs/benchmarks/workbook-large-speedup-2026-09-06.md` を参照してください。
- opaque XML要素検索で、不在の要素名を先に判定し、タグ名の一時`String`生成をborrowへ置換しました。共有文字列収集では文字列セルだけを座標順に並べ替え、数値セルの不要なsort／lookupを削減しました。直前の実装を基準とした40組の交互測定で、数値1万／10万セル・混在1万セルは約2倍に高速化し、全ZIP部品の展開後バイト一致を確認しました。小規模は約1.09倍で1.2倍目標未達です。保存耐久性とversionは維持し、詳細は `docs/benchmarks/workbook-speedup120-2026-09-06.md` に記録しました。
- ClosedXML 0.105.1／.NET SDK 10.0.400を固定した比較用workerを追加しました。macOSの同一入力・編集・`F_FULLFSYNC`・atomic rename・再読込条件で、elixcee／ClosedXML／openpyxlを3ケース各30反復×2回測定しました。両runの全ケースでelixceeの中央値がClosedXMLを下回りましたが、高負荷環境の大きなp95変動とラウンド単位の逆転も含め、暫定値として `docs/benchmarks/workbook-closedxml-2026-09-06.md` に記録しました。C#依存は製品の依存関係に追加していません。
- 公開されている低レベルcell map経由で数式本文が変更された場合も、warmな依存計画を無効化して再評価するよう修正しました。
- worksheet XMLを一時的な巨大`String`へ構築せず、保存先の`ZipWriter`へ直接ストリーミングするsink経路へ変更しました。passthrough entry全量保持の削減は引き続き未完了です。
- passthrough entry はメタデータ解析後に所有権移動するようにし、保存用payloadの不要な二重cloneを除去しました。raw ZIP全体の遅延展開は引き続き未完了です。
- dirty/full-rescan の代表値一致を検査しながら p50/p95 を出力する、formula dirty propagation のローカル校正バイナリを追加しました。
- formula dirty propagation の校正を100式／1,000式のcontrolled matrixへ拡張し、各ケースのp50/p95とfull-rescan一致を記録しました。
- formula dirty propagation の測定バイナリに、Unixのpeak RSS・user/system CPUカウンタ出力を追加しました。
- manual→automatic遷移と循環参照のp50/p95校正も追加し、dirty propagationの測定行列を拡張しました。
- release-profile 10反復の dirty/full matrix と resource counter 結果を測定成果物へ保存しました。
- Writerのworksheet sink／passthrough clone削減後に代表XLSXのrelease保存soakを3回実行し、round-tripとallocator結果を保存しました。
- formula dirty propagation の release matrix を30反復へ拡張し、p95の再現性を測定成果物へ保存しました。
- elixcee／openpyxl／LibreOfficeの初回XLSX速度比較と、比較不能条件を含む公開用ベンチマーク成果物を追加しました。
- openpyxl比較に保存後fsyncを追加した初期5反復を記録しました。ただしmacOSのRust `sync_all`（`F_FULLFSYNC`）とは不一致だったため、この初期測定を公平な比較とする説明を撤回しました。
- 標準の耐久性保証を変えず、最終fsyncを省略する明示的な `save_workbook_fast` APIを追加しました。高速経路はクラッシュ耐性が弱いため用途を限定します。
- XLSX writerを圧縮前後の64 KiBバッファとシーク不要のZIP出力へ変更し、標準の `sync_all`／atomic renameを維持したまま細切れ書き込みを削減しました。数値1万／10万セルの保存中央値を変更前より約49%／55%短縮しました。
- macOS `F_FULLFSYNC`まで一致させたopenpyxl 3.1.2比較を、3ケース・30反復＋独立40反復で実測しました。未公開Rust APIのload/edit/save/reload中央値は再測定で1.90〜2.98倍高速でした。条件・p95・全生データは `docs/benchmarks/workbook-equal-durable-2026-09-06.md` を参照してください。
- 純粋な VM 内 `Scripting.Dictionary` adapter を追加しました。`New`、キーの追加・検索・削除、`CompareMode`、挿入順のキー列挙をサポートし、`CreateObject` の外部効果遮断は維持します。
- VBA built-in `Collection`を追加しました。`New Collection`、1-based index、case-insensitive key、`Add`/`Remove`/`Item`/`Count`、値/object `For Each`、object-valued itemの`Set`取得、`With`、参照alias、到達可能性回収、要素budgetに対応します。
- Collectionの`Call c.Add(...)`/`Call c.Remove(...)`構文と、export済みclass moduleの基本object連携（生成、instance field、初期化、Sub/Function、Collection/`For Each`/`With`）を追加しました。
- class moduleの非indexed Property Get/Let/Set、object-valued field/引数、`Class_Terminate`、`Implements`署名検証、Private/Friend境界、class object arrayを追加しました。
- indexed `Property Get/Let/Set`、型付き引数を受けるobject-returning class `Function`、interface宣言型を保持する厳密なSub/Function/Property dispatch（`With`を含む）を追加しました。
- Range objectの相対`.Cells`/`.Range`、1-based default `Item`/`Value`、`With`内の相対参照を追加しました。
- Worksheetの`Name`/`Index`/`Visible`/`UsedRange`/`Activate`/`Select`と、Workbookの`Name`/`ActiveSheet`/`Worksheets.Count`/`Sheets.Count`をdirect/object参照へ追加しました。シート名変更後も保持中のWorksheet/Range参照は同じシートを指し続けます。
- `Range.SpecialCells(Type[, Value])`を全`XlCellType`定数へ拡張し、定数/数式の型mask、空白、最終セル、入力規則、条件付き書式、legacy comment、該当なし1004、走査budgetに対応しました。
- Range、Collection、イベント、外部オブジェクトの残作業を安全境界ごとに分割してロードマップへ追加しました。
- 小容量入力が監視threadの初回pollより先に完了しても、事前作成済みCLI cancel-fileを確実に反映するよう競合を修正しました。

## [1.0.1] - 2026-09-04

- readerの総work budget、deadline、協調キャンセルをRust/Python/CLIへ通し、CLIではSIGINTとcancel-fileから安全に中断できるようにしました。
- ZIP/XMLのchunk境界でキャンセルを確認し、キャンセル・大容量入力・反復資源回収・read/mutate/write経路をローカルfixtureで校正しました。
- XMLイベントをborrow/`Cow`化し、worksheetの構文・資源・shared-string検証とセル構築を単一走査へ統合しました。
- `WorkbookSheet`からVMへの変換を所有権移動に変更し、セル値・数式・シートメタデータの不要なcloneを除去しました。
- macOS arm64のrelease build、dense 400,000セル、3回のローカル比較で、reader→VMロード中央値を約1,386.5 msから650.5 msへ短縮しました（約2.13倍、条件と制約は`docs/measurements/reader-optimization-2026-09-04.md`参照）。
- reader測定成果物のsemantic validator、拒否系self-test、Cargo package境界チェックをCIへ追加しました。

## [1.0.0] - 2026-09-01

- 文書化されたデータ処理サブセットと、安全な失敗動作をv1サポート契約として固定しました。
- 完全なExcel/VBA互換性を主張せず、`supported`、`preserved`、`warned`、`rejected`、`unverified`の境界を公開しました。
- READMEとセキュリティモデルのリリース表記を1.0.0へ更新しました。

## Older releases

[0.xの変更履歴](docs/history/CHANGELOG-0.x.md)へ移動しました。過去の記録は保持しています。
