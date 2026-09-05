# elixcee Roadmap

更新日: 2026-09-06。対象versionは **1.0.2** です。
完了項目は記載した実装・測定の範囲に限ります。公開先の状態はリリースごとに別途確認します。
版ごとの変更は [CHANGELOG](CHANGELOG.md)、実装範囲は
[FUNCTIONS](FUNCTIONS.md)、保証範囲は [v1契約](docs/v1-support-contract.md) を参照してください。

## 方針と完了の定義

Excel不要の、Rust/Pythonによる安全なファイル入出力・数式計算・VBAデータ処理を中心に開発します。
GUI、任意のCOM、完全なExcel/VBA互換、無条件のlossless保存は保証しません。
JavaScript互換APIは別トラックで、`packages/xlsx` はprivate・未公開です。

- `BUILD`: 実装と回帰テスト。`MEASURE`: 固定条件による測定。`GATE`: 配布・互換性・安全性の判定。
- `[x]` は記載した範囲の実装／ローカル検証が完了した意味です。公開・3 OS検証・Excel完全一致を含意しません。
- 互換性は `supported / preserved / warned / rejected / unverified` を区別します。
- 外部効果は既定で遮断します。保存耐久性、入力制限、意味論を速度のために緩めません。
- 測定前に対象・baseline・反復数・成功条件を固定し、未達・p95悪化・除外条件も残します。

## 次の実行順

1. **P0 大規模I/O**: passthrough ZIPの遅延展開とpeak RSSを調べ、同一耐久性・出力同値で比較する。
2. **P0 再測定**: quiet-hostで小規模／大規模の中央値・p95と、style／数式密度別の結果を確認する。
3. **P0 数式校正**: dirty/full再走査の一致と、より大きい依存グラフ・RSS・CPUを検証する。
4. **P1/P2 性能候補**: 下表から、一度に一つの変更を実装して回帰を確認する。
5. **互換性・配布ゲート**: Excel oracle、3 OS、clean-install、安全性の証拠を揃える。

## 性能バックログ（優先順）

| 優先度 | 未完了の作業 | 完了条件 |
|---|---|---|
| P0 | 元ZIPのpassthrough entryを全量展開・保持しないWriter | 大規模partのclone削減、peak RSS・保存時間、全part／relationship同値 |
| P0 | 小規模ファイルの1.2倍目標 | 17セルの固定保存費用を改善し、標準sync・atomic renameを維持。quiet-hostでp95も確認 |
| P0 | 大規模全条件で追加1.2倍 | 10万・40万・100万セルの中央値比で判定。style／数式密度・RSS・単一sheetの上限も校正 |
| P0 | formula dirty propagationの完全校正 | 大規模controlled matrix、full再走査との値一致、p50/p95、CPU/RSS、循環・manual/automatic |
| P1 | VBA bytecode／symbol intern | AST経路との一致、演算主体・セルI/O主体の両方を測定 |
| P1 | canonical packed/tiled cell storage | 公開HashMap APIを保つoverlay。Range・formula・style・save・fork・GCの回帰 |
| P1 | Writerのstyle／shared-string overlay | 全map／文字列cloneを削減し、密度別RSS・CPU・出力同値を確認 |
| P2 | Range中間grid削減 | Sort・転置・広域copyでrow permutation／tile／iteratorを検討。1-based・Empty・formulaを保持 |
| P2 | Python GIL解放・複数VM並列 | Python objectを触らない区間を対象にthroughput・RSS・例外伝播を確認 |
| P2 | byte指向lexer | char配列・一時文字列を削減。Unicode識別子と診断位置を保持 |
| P2 | object GCの世代別／dirty graph化 | Class_Terminateの一度性、Collection cycle、budget、長時間soak |

### 直近の性能実績

- [x] Readerのborrow化・単一走査・VMへの所有権移動、read budget／deadline／協調キャンセル。
- [x] worksheet XMLのZipWriterへの直接出力、passthrough payloadの二重clone除去。
  **raw ZIP全体の遅延展開は未完了**。
- [x] formula dirty/full比較、100／1,000式matrix、manual→automatic・循環参照、CPU／RSS、30反復の校正。
  完全校正ではありません。[測定一覧](docs/measurements/README.md)
- [x] durable saveのbuffering・seek不要ZIP出力、opaque XML検索とshared-string収集の最適化。
- [x] 行別treeから座標sortへの変更、セル番地のstack生成、Readerの属性・型buffer再利用。
  [大規模測定](docs/benchmarks/workbook-large-speedup-2026-09-06.md)
- [x] 固定版openpyxl・ClosedXMLとの同条件比較と、撤回した初期比較を含む生データの保存。
  [比較一覧](docs/benchmarks/README.md)

2026-09-06の直前実装に対する大規模20組の交互測定:

| セル数 | before / afterの全体中央値比 | 厳密な1.2倍目標 |
|---|---:|---|
| 100,000（1 sheet） | 1.180 | 未達 |
| 400,000（1 sheet） | 1.318 | 達成 |
| 1,000,000（250,000 × 4 sheets） | 1.196 | 未達 |

これはローカルの特定入力・特定baselineに限る結果です。1.196を丸めて達成扱いしません。
中央値比とpaired speedupは別の指標です。既存XML要素数制限は緩めていません。
過去の競合比を掛け合わせて、最新実装の競合比とすることもしません。

## XLSX編集互換性トラック

各行は残る評価・拡張範囲です。既存APIがあることと、全ケースの互換性を証明したことを分けます。

| ID | BUILD | MEASURE / GATE |
|---|---|---|
| X0 棚卸し | read／edit／write／round-trip／Excel確認の機能matrixを維持 | 未実装・保持のみ・拒否・未検証をfixtureと結び付ける |
| X1 基本モデル | セル型・数式・style・merge・行列・sheet操作とdirty part管理を整合 | 挿入／削除／移動／copy・大規模疎密入力・参照追従 |
| X2 構造化オブジェクト | table・filter・validation・conditional formatting・defined name等を段階拡張 | owner XMLとrelationshipを含む往復、Excel実機で修復警告・意味を確認 |
| X3 数式・マクロ共存 | A1/RC・sheet間参照・named range・cached value・再計算方針 | エラー／日付／型変換・manual/automatic・VBA binary保持。保持と実行を区別 |
| X4 保存・復旧 | 未知part保持、dirty part再生成、transactional save | 中断・失敗時に元ファイルを保護。content type／relationship／署名等の限界を明示 |
| X5 性能・API | streaming、Python/Rust/JS境界、互換API差分を整理 | 固定版競合、疎密／数式／style／多sheet、CPU・RSS・p50/p95・出力サイズ |
| X6 リリース | サポートmatrix・既知の損失・migration例を固定 | clean-install、3 OS、回帰・ライセンス・配布物・公開情報を別々に確認 |

入力には大きさだけでなくXML・ZIP・work budgetの上限があります。
[limits](docs/limits.md) と [architecture](docs/xlsx-architecture.md) を参照してください。

## VBA・数式の残作業

- [x] Range相対参照、default Item/Value、Worksheet/Workbookの基本member、SpecialCells拡張。
- [x] VM-local Collectionとexport済みclass moduleのobject連携、Property・interface dispatch。
- [x] VM-local Dictionary adapter。キー正規化などの制限は [FUNCTIONS](FUNCTIONS.md#in-memory-dictionary) に明記。
- [ ] Save／Close／外部リンクなど作用を持つmemberは、安全境界とoracle fixtureを先に定義して段階実装。
- [ ] Workbook_Open／Worksheet_Change等のイベントは、EnableEvents・再入抑止・決定的dispatch順・budget付きで設計。
- [ ] runtime errorの文字列分類依存を減らし、型付き診断へ移行。
- [ ] 複数module間のUDT名衝突と診断位置など、解決規則の残差を明確化。
- [ ] Excelとの型変換・丸め・日付・Empty／Error・配列境界・再計算の独立oracle比較を拡張。
- [ ] 実運用macroの回帰と、数式の依存グラフ・循環・volatile／dynamic array等の対応境界を検証。

ローカルsynthetic fixtureの通過数を、実Excelの意味論一致件数として扱いません。
[互換性ハーネス](compat/README.md) と [CLI契約](docs/agent-contract.md) が検証の入口です。

## セキュリティトラック

| ID | 維持・追加する作業 | 残る検証 |
|---|---|---|
| S0 脅威モデル | 入力、出力、保存先、formula／VBA、互換API、秘密情報の境界 | 攻撃面と失敗時の非破壊性をfixtureへ対応付ける |
| S1 ZIP/XML | path／relationship検証、暗号化・特殊entry・不正XML拒否、資源上限 | limit近傍・圧縮爆弾・中断・大規模の3 OS校正 |
| S2 parser/formula/VM | 深さ・式・配列・実行budget、型変換とエラーを検査 | adversarial corpus、fuzz、長時間CPU/RSS、panic／hang／leak |
| S3 出力・互換API | prototype安全性、HTML既定escape、明示的rawHtml、CSV／外部参照境界 | 悪意ある入力の安全な差分、保存失敗・relationship保持 |
| S4 隔離・依存 | 最小権限、外部効果遮断、必要に応じprocess／container隔離 | worker強制終了と協調cancelの区別、依存監査・供給網・clean install |
| S5 継続ゲート | 回帰corpus、fuzz、配布物とsecurity policy | freshなadvisory、license、3 OS、長時間soak。未検証は明記 |

- [x] ZIP/XML構文・entry／part／全体・work budget、協調中断のローカルfixture。
- [x] reader資源回収・read/mutate/write、semantic validatorと拒否self-test、Cargo測定境界の検査。
- [ ] Linux／Windowsを含む資源校正と、長時間fuzz／CPU／RSS・隔離環境検証を完了する。

既存のmacOS測定と安全策は [測定記録](docs/measurements/README.md)、
[脅威モデル](docs/xlsx-security-model.md) に集約します。限界値の緩和を達成条件にしません。

## 配布・サポートゲート

1.0.2のローカル検証（2026-09-06、macOS）:

- [x] Rust全workspace／全targetの1,705 tests、strict clippy／Rustdoc、依存監査、測定記録検証。
- [x] crate検証、wheel／sdist生成、独立Python環境でのimport・VBA・通常／fast保存往復。
- [x] JS型／differential／WASM bundle／packed consumerと、VBA corpus 581件・意味論386件の既存ゲート。

以下は継続的な配布・サポート要件であり、上記のローカル検証だけでは完了しません。

- [ ] Python wheel／sdist、Rust crate、CLIのclean-installと最小利用例を各対象OSで確認。
- [ ] JS packageはutils／read／write／browser／Node／型定義を別々に検証。公開する場合は独立した判断を行う。
- [ ] license／notice／package内容／依存監査／回帰・互換性matrixを確認。
- [ ] サポート外・既知の損失・セキュリティ境界・移行例を公開文書へ反映。
- [ ] tag、registry、workflow、正式Release、worktree状態を別々に確認。

このゲートは将来のリリース判定です。文書の更新だけで完了にしません。
実装・測定は自律的なローカル作業として進め、version変更・push・公開は別の操作として扱います。
