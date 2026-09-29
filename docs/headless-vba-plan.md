# ヘッドレスExcelファイル処理のVBA強化計画

計画日: 2026-09-12。状態: **V0〜V7aおよびV8のローカルCLI計測入口を実装・検証済み／Excel oracle未測定**。
elixcee基準: v1.0.12、`ca9f2cf49b8c491c04fc8678eb28db4aefbf09de`。
優先順位は [ROADMAP](../ROADMAP.md)、現行機能は [FUNCTIONS](../FUNCTIONS.md) を参照。
この計画は公開APIや対応保証を追加しない。

## 1. 目的とスコープ

elixceeの主目的は **Excelをインストールせず、Excelファイルを読み、編集・計算して保存すること**。
VBAは既存の集計・転記・検証ロジックを再利用するための一つの入口であり、必須ではない。
VBAなしのRust／Python API、数式エンジン、OOXML保持、処理速度を犠牲にしない。

| 区分 | 対象 |
|---|---|
| 優先する | XLSX／XLSMの値・数式・対応書式、範囲転記、集計、照合、検証、対応する再計算、出力保存 |
| VBAで補強する | 手続き呼出し、型・配列・ローカルobject、Range操作、ファイル処理に必要なイベントと診断 |
| 必要性を確認して追加 | 埋め込みVBAソースの読込、不足する複数ブック操作、明示許可された入出力 |
| この計画では追わない | VBE／IDE／アドイン、UserForm描画、画面操作、任意COM／ActiveX、Office連携、印刷、共同編集 |

採用条件は「固定したファイル処理fixtureを完遂させるか、結果・保存・安全性を改善するか」。
競合にあるという理由だけでは採用しない。読込時のマクロ自動実行や外部効果の既定許可は追加しない。
GUIを使うマクロは無条件に成功扱いにせず、拒否／明示した代替応答／未対応を記録する。

## 2. xlflowとの比較で混同しないこと

比較対象は [xlflow v0.31.2](https://github.com/harumiWeb/xlflow/releases/tag/v0.31.2)
（2026-09-04公開、commit `0b8ca0eb47f834cb97be6730d73d688a811da676`）に固定する。
xlflowはWindows上のMicrosoft ExcelをCOM経由で利用する。VBA実行の互換性はExcel本体に由来するため、
「独立VBAエンジンの対応関数数」という比較はしない。

| 観点 | xlflow v0.31.2の公開仕様 | elixceeの強化方針 |
|---|---|---|
| 実行環境 | 実Excel、headless／interactive、warm session | Excel不要のnativeファイル処理を維持 |
| 呼出し | 型付き引数、実行前push、保存先指定 | 元の業務Sub／Functionを変えず、正しい引数束縛と明示保存 |
| テスト | テスト発見、setup／teardown、期待エラー、パラメーター、隔離・再試行 | 既存test-workbookに結果／状態検証と再現性を統合 |
| 診断 | 実行phase、source location、XlflowDebug.Log、UI応答wrapper | VMが把握する位置・呼出し・セル変更を上限付きで記録 |
| 失敗復旧 | timeout後の隔離markerとrecovery。VBAが走り続ける場合を扱う | 所有する実行の停止と未完成出力の非公開を検証 |

根拠: [README](https://github.com/harumiWeb/xlflow/blob/v0.31.2/README.md)、
[run](https://github.com/harumiWeb/xlflow/blob/v0.31.2/vitepress/commands/run.md)、
[test](https://github.com/harumiWeb/xlflow/blob/v0.31.2/vitepress/commands/test.md)、
[runtime debugging](https://github.com/harumiWeb/xlflow/blob/v0.31.2/docs/specs/runtime-debugging.md)。

注意: xlflowのパラメーターテスト発見でByRef／Optional／ParamArrayを制限していることは、
Excelでそれらを実行できない意味ではない。
またUI応答は専用wrapperの契約であり、任意のネイティブdialogの再現保証ではない。
[パラメーターテストADR](https://github.com/harumiWeb/xlflow/blob/v0.31.2/docs/adr/ADR-0015-parameterized-vba-tests.md)
を比較時に参照する。

**到達目標**は、事前固定したヘッドレス業務処理でxlflow＋Excel以上の完遂率を保ち、
速度・再現性・停止／保存の安全性で実測上の利点を示すこと。全VBA互換性でExcelを超えるとは主張しない。
比較環境がなければ実装は進められるが、この到達判定は`unverified`のままとする。

## 3. v1.0.12のコードから確認した着手点

| 確認箇所 | 現状 | 計画 |
|---|---|---|
| `src/parser/mod.rs::parse_params` | ByVal／ByRefをASTへ保持。Optional／ParamArrayは拒否 | V1: 黙って異なる結果になる引数意味論を最優先 |
| `src/parser/ast.rs` | 手続きのparams／param_typesとSourceSpanは存在 | V1: 内部の引数metadataとbinderを追加し、公開AST破壊を避ける |
| `src/parser/mod.rs` | Debug.Printはbounded sinkへ保持、Debug.Assertはpolicy no-op、Option Explicitは無視 | V3で観測・診断化、V2で変数宣言規則を段階導入 |
| `src/vm/mod.rs` | On Error／Resume、Err.Sourceを含む状態、部分的な型付きfailure、各種budgetあり | V3／V5: 未移行エラーと復旧境界を補強。ゼロから作り直さない |
| 既存object・event実装 | Collection、限定Dictionary、class／Property／interface、codeNameによるhandler選択あり | V4／V5: 不足する意味論だけを独立oracleへ照合 |
| `compat/corpus` | 581件のローカル生成corpusの実行記録あり | 回帰として維持。実Excel由来の正解や実務完遂率とは分離 |

既存corpusの記録は572 PASS、8 EXPECTED_RUNTIME_ERROR、1 NONDETERMINISTICであり、
「581件のExcel一致」ではない。[記録](measurements/vba-corpus-local-2026-09-10.md)
と [oracleの現状](../compat/README.md) を保持する。

## 4. Phase別実装計画

以下のパスは既存ファイルを除き**新設候補**。既存の実行・テスト経路へ統合し、別製品や別VMを作らない。
各小段階をfixture追加 → 実装 → 既存回帰 → 契約更新で進める。

### V0 — 比較契約と回帰の固定（P0）

進捗: **部分実装**。manifest、結果分類語彙、Transfer／ByRef／Sort・Filter・Findの固定fixture、自己検証、XLSM provenance監査は実装済み。ローカルrunnerはCIの`compat-vba` jobからも実行する。
xlflow＋Excel runner、実測、独立holdoutは未実装・未測定。

- [x] `compat/vba-workflows/`にmanifestを設計する。入力hash、元VBA、entrypoint、引数、初期状態、期待出力、許可効果、上限、出典、licenseを保存する。現状は合成fixtureである。
- [x] 結果を`pass / mismatch / expected_error / unsupported / policy_blocked / timeout / not_measured`へ分類する語彙とmanifest検証を追加した。期待エラーは業務成功の分子に入れない。
- [x] `scripts/audit-xlsm-vba-provenance.py`で、XLSMの保持された`vbaProject.bin`、別渡しの`.bas/.cls/.frm`、未取得のmodule identityを区別する。監査はhashと展開サイズ上限を記録し、OLE/VBAを抽出・変更・実行しない。
- [ ] 既存`compat/corpus`／`vba-semantics`／`vba-diagnostics`を再利用する。既存のExcel COM adapter契約を確認し、xlflow用runnerを分離する。
- [ ] 現在のXLSM読込からVBA source／module identityまでの経路を監査する。別途渡した`.bas`と埋め込みソース、保持しただけの`vbaProject.bin`を区別する。
- [ ] ローカルで回る回帰と、Windows＋Excelを要するoracle／速度測定を別jobにする。未測定を成功扱いにするfallbackは設けない。

完了条件: schema検証、同じ入力の反復で同じ正規化結果、欠測・失敗分類のrunner self-test。
実行コマンド: `python3 scripts/check-vba-workflow-manifest.py`。
baseline測定が未完でもV1のBUILDには進めるが、比較結果は公開しない。

### V1 — 引数と呼出しの正確性（P0、V0の契約後）

進捗: **V1a parser metadataと明示／暗黙ByRefのスカラー／配列要素／UDTフィールドwrite-back（multi-moduleのcaller scope含む）、同一変数を複数ByRef引数へ渡すalias同期、V1bのOptional／ParamArray／IsMissing／名前付き引数が部分実装**。
省略modifierは`DefaultByRef`として区別し、単純変数、基本的な`arr(i[,j])`配列要素、ローカルUDTの`record.field`／ネストフィールド／UDT配列要素フィールドはwrite-back、リテラル／式は既存互換の値渡しとする移行境界に置いている。
Optionalの既定値、ParamArrayの配列化、`IsMissing`のcall-frame追跡、名前付き引数の宣言位置束縛、Sub／Functionの引数個数checkまで追加した。名前付き呼出しで途中のOptionalを省略した場合も、`Empty`とは別の内部欠損マーカーで既定値と`IsMissing`を保持する。OptionalByRefへリテラルを渡す場合は一時値として受け入れ、明示ByRefのlvalue検証は維持する。Sub／Functionとも名前付き引数名を大文字小文字非依存で照合し、Property GetにもOptionalの末尾／名前付き省略、`IsMissing`、OptionalByRef index引数の一時値を接続した。オブジェクトを返すProperty GetのスカラーByRef index引数も、callee内の変更をcallerへwrite-backする経路を追加した。class FunctionにもスカラーByRef、Optional、名前付き引数、`IsMissing`を接続し、object-return class Functionとclass Subにもスカラー／object ByRef write-backを接続した。class Subでcalleeが`Set arg = New Class`のように参照を置換した場合も、caller側のobject identityへ戻す経路を追加した。
aliasの完全な束縛、Function／PropertyのByRef、nested/property/object要素、Missingの公開Variant表現、ParamArrayと名前付き引数の混在、不正な宣言の完全検証は未完で、実行互換性全体は改善済みとは扱わない。

- [x] **V1a 部分 BUILD**: `parser`で渡し方・宣言型・位置を保持し、ユーザー定義Sub／Functionの明示ByRefと、単純変数・基本的なVBA配列要素・ローカルUDTの単純／ネストフィールド・UDT配列要素フィールド・module-level UDTフィールド・module-level UDT配列要素フィールドへ渡す暗黙ByRefをwrite-backする。配列要素targetはcallerのmodule scopeを保持し、multi-module呼出しでもcalleeから同じ配列／recordへ戻す。module-level UDT配列の通常フィールド代入も接続した。スカラーを返すclass PropertyのByRef lvalueも、Property Get→callee→Property Letの経路で接続した。alias slot、object Property Set／複雑なProperty要素の完全な参照束縛は後続作業。
- [ ] コピーして最後に戻す方式にしない。同じ変数を二つの引数へ渡すalias、nested call、途中エラー前の変更も正しく観測できる内部slot／参照を用意する。スカラーaliasはstatement境界の同期と回帰を追加済み。nested callの複雑なalias、途中エラー時のExcel校正、Property／配列要素のlvalueは未完。
- [ ] 括弧で式として渡す場合の一時値と型変換、object参照の値渡しとobject自体の変更を区別する。未対応lvalueは呼出し前に診断し、値渡しへ黙って落とさない。UDTフィールドのローカル／module-level経路は接続済みだが、複雑なobject／Property要素は未完。
- [x] **V1b 部分 BUILD**: Optionalの既定値、末尾ParamArray、引数省略、call-frame単位の`IsMissing`、名前付き引数の宣言位置束縛をSub／Functionへ追加し、compile checkの引数個数判定も可変個数へ更新した。Missingの厳密なVariant表現、Optionalと名前付き引数の混在省略、ParamArray named call、不正な宣言の完全検証は後続作業。
- [ ] Sub／Function／class method／Propertyへ同じbinderを適用し、CLI・Python入口も同じ検証を通す。Sub／Function／class methodとProperty Get／Let／Setのindex引数について、スカラー・基本配列要素ByRefとCLI／Python入口の実行は共通境界へ接続済み。標準Moduleのobject型Sub／Function引数にもobject identity／ByRef write-backを接続し、ネストした同名パラメータが外側の引数を壊さないようcurrent frameだけを同期する。標準Functionのobject returnも`Set result = Make(...)`の経路へ接続した。Property GetのOptional省略と基本ByRef一時値、object returnのOptional index正規化とスカラーByRef write-back、class FunctionのスカラーByRef／Optional／名前付き引数、object-return class Functionとclass Subのスカラー／object ByRef、Property Let／SetのOptional／名前付きindex正規化、値引数ByRef write-backも接続した。class Subのobject ByRef参照置換も接続した。完全な型／lvalue校正、Function呼出しとCollection default memberの曖昧性をExcel oracleで校正する作業は未完。大文字小文字非依存と1-basedセルを維持する。

完了条件: 少なくとも60件の小fixture（alias／ByVal／一時値30、Optional／ParamArray／名前付き20、object／Property10）。
独立Excelで呼出し後の値・型・エラーを照合する。公開候補の必須項目で未対応・不一致が残れば完了にしない。

### V2 — 宣言型、配列、scope（P1、V1と連動）

進捗: **V2aのスカラー引数coercionとV2b整数変換の範囲検証が部分実装**。
`Integer`／`Long`、浮動小数系、`String`、`Boolean`を呼出し境界で扱い、
`CByte`／`CInt`／`CLng`の丸め後範囲も検証する。
Date／Currency／object／配列とoverflowのExcel校正は未完で、これを完全なVBA型互換とは扱わない。

- [x] **V2a 部分 BUILD**: Sub／Functionの呼出し境界で`Byte`／`Integer`／`Long`、浮動小数系、`String`、`Boolean`へ値を変換し、整数の宣言幅を検証する。型変換不能・範囲外はエラーにする。callee内で変更された宣言型付きByRef値をcallerへ戻す際にも同じcoercionと整数幅検証を再適用し、配列要素のwrite-backを回帰化した。whole-dayの`Date` variant、4桁へ量子化する`Currency`、`Null`から宣言型`String`への暗黙変換拒否も追加した。分数を含むDate/time保存、宣言型slotへの全型厳密化、object／配列の完全校正は後続作業。
- [ ] **V2b 値の表現**: VBA Null／Empty／Nothing／MissingとExcelセルerrorを分離する。`CByte`／`CInt`／`CLng`の丸め後範囲検証と、VBA変数に対する`IsObject`判定（`Nothing`を保持するobject変数とEmpty／scalarの分離）は部分実装済み。Currencyの固定小数、Dateの時刻成分、Missingの厳密表現、object式全般の`IsObject`判定、独立oracle照合は未完。
- [ ] **V2c 配列・scope**: Option Base、LBound／UBound、ReDim Preserve、二次元配列、module変数／Static／Const、Option Explicit／Option Compareを監査・補強する。`Option Explicit`の宣言保持と、diagnoseのstrict profileでの未宣言読み取り・代入先・基本object／配列receiver検証、`Option Base`の既存保持、標準moduleのmutableなmodule-level変数の共有・永続化、固定次元およびdynamicなmodule-level配列のbounds／read／write／`ReDim Preserve`は実装済み。object配列の多次元bounds、`LBound`／`UBound`／`IsArray`、未初期化要素の`Nothing`保持、`ReDim Preserve`後の要素identityも回帰化した。module-level `Const`の読込時評価・宣言型変換・read-only代入拒否、手続き内`Static`のpersistent storage、標準moduleの`Option Compare Binary`／`Text`によるVBA文字列演算子の比較モード、基本的な標準module object変数（Worksheet／Range／Collection／Scripting.Dictionary）の`Nothing`初期化・`Set`・修飾アクセス・手続き間identity保持も実装・回帰化した。Dictionaryは外部COMではなくVM-local adapterを使用する。宣言なしの比較は既存互換のcase-insensitive挙動を維持する。object初期化式、Staticのclass／複雑宣言、class moduleのOption Compare、複雑なobject／配列要素の宣言検証は未完。既存UDTのmodule-local解決は維持する。
- [ ] Option Explicitを無視してきた旧挙動との差は、明示互換profileと移行文書で管理する。新しい対応保証では未宣言変数を黙認しない。

完了条件: 正常値だけでなく境界、丸め、空値、型不一致、overflow、配列範囲外を比較。
VBA RoundとWorksheetFunction.Round、VBA Dateとworkbookの日付systemを別契約にする。

### V3 — エラー、ログ、再現実行（P1）

進捗: **V3aのDebug.Print分離sinkとV3cのopt-in trace基盤を部分実装**。`Debug.Print`はVMから取得できる
bounded logへ記録し、MsgBox／stdoutとは分離した。`diagnose`のJSONにも非空時だけ
`debug_output`として出力する。実行時の引数束縛失敗は手続き名・パラメータ名・1-based位置を
`argument_failure`として出力する（事前compile checkの引数個数エラーは既存診断経路）。Debug.Assert policy、
CLI設定、Excel oracle校正は未完であり、診断機能全体の完了とは扱わない。

- [ ] **V3a 呼出し診断**: 引数名・宣言位置・呼出し位置付きの型付きerrorへ接続する。日本語sourceでもSourceSpanの文字単位を壊さない。Debug.Printを上限付きsinkへ流し、stdoutのJSONと分離する。Debug.Printのbounded sink、引数束縛失敗、宣言型coercion失敗の基本evidenceは部分実装済み。
- [ ] **V3b エラー制御**: `Err.Raise`と実行時失敗の`Number`／`Description`／`Source`／`HelpFile`／`HelpContext`を`ErrorEvidence`として取得し、diagnose JSONとCLI `--json`へ接続した。message文字列fallback、On Error／Resumeの手続き間伝播・handler中の再エラー・行番号のExcel校正は未完。
- [ ] **V3c opt-in trace**: `Vm::enable_trace`で明示的に有効化した場合に、entry／exit、cell変更、statement、failureを実行ID・source hash付きで記録し、`take_trace`とdiagnose JSONへ接続した。Python `Vm.enable_trace`／`take_trace`とCLI `--trace <id>`も追加した。CLIのrecalculate／save phase境界も値・パスを含めず記録する。成功・失敗JSONの両方へイベントを出力し、イベント件数・byte上限、値のredaction、打切り表示を実装・回帰化した。event種別の完全網羅、独立Excel校正は未完。
- [ ] Debug.Assertは非GUIの明示policyにする。既定の`ignore`とhostが選べる`error`を実装し、Python APIとCLI `--debug-assert`へ接続した。VBEのbreak／GUIは呼び出さない。clock／RNG／必要なUI代替応答を入力として固定し、seedと設定から再現できるようにする。代替応答ありの結果は通常実行と分ける。

完了条件: trace有無で値・式・errorが変わらず、切詰め・日本語位置・機密値非出力の回帰が通る。
中断・security failureをVBAのOn Errorで無効化できない既存契約を維持する。

### V4 — 範囲転記とデータobject（P1、V1–V2後）

- 進捗: DictionaryのCompareMode、キー大小文字、数値／文字列キーの分離、Add／Exists／Keys／Items、Itemの既存更新・新規追加、Range.Value2のscalar／2D配列／Excel error保持、Collection／Dictionaryの列挙開始時スナップショットを実装・回帰化。標準moduleのCollection／Dictionaryは手続き間identityを保持し、`With`内のCount・method dispatch・GC rootも同じVM-local objectへ解決する。標準module object aliasの`Set`も共通解決へ接続した。Pythonには値gridと分離した`get_range_formulas`を追加し、式保持／置換を観測可能にした。Collection／Dictionary aliasのVM identityは回帰済みだが、Excel oracle照合は未完。
- [ ] **V4a Range**: Value／Value2のscalarと2D配列、形状不一致、空値・error、式の保持／置換をfixture化する。Value2の基本読書きとerror保持、Cells／RangeのValue2書き込み、VBAのrow-major配列転記、空値／Error保持、式の計算結果転記、形状不一致の事前拒否、転記先の`<f>`消去と転記元の`<f>`保持確認を実装・回帰化した。一括読み書きのfixture拡充、式の保持／置換全域、Excel oracle照合は未完。
- [ ] **V4b Dictionary／Collection**: Dictionaryのキー型・比較mode・数値／文字列キー分離・既存項目更新・Keys／Items、列挙開始時スナップショット、module変数を別Subからalias経由で更新するfixtureは実装済み。既存class objectのidentityとaliasのExcel oracle校正は未完。
- [ ] `CreateObject("Scripting.Dictionary")`が対象fixtureに必要な場合に限りVM-local factoryへの厳密なallowlistを検討する。他のProgIDやホストCOMを許可しない。
- [ ] **V4c 実務不足分**: `Cells.Find` の基本検索と `MatchCase`（既定false、明示true）、Sort／Filterの基本経路を実装・回帰化し、V0に表処理fixtureを追加した。次は既存SpecialCells／Paste／非表示／保護を回帰する。Pythonやbrowserの機能がそのままVBA object modelにあるとは扱わない。

完了条件: VBAと直接APIの同じ編集で同じcell／formula state、保存再読込でも一致。
Chart／Pivot等の未知参照を書き換える操作は、G2の更新・保持経路がなければ拒否する。

### V5 — 再計算、イベント、失敗出力の扱い（P1、G3–G4と共通）

- [ ] **V5a 実行順**: Manual／Automatic、一括変更、Worksheet_Change、EnableEvents、再入、codeNameの既存実装を校正する。event内エラーと再計算失敗も記録する。
- [ ] **V5b 明示バッチ境界**: 一時workbook上で実行→再計算→検証→一時出力→publishする経路を設計する。元入力と既存出力を失敗時に変更しない。CLIは構造検証とelixcee reader roundtripを通過した後に成功JSON／成功終了を返し、workflow runnerはさらにstdlibで保存値を照合する。runtime／directory output失敗の注入を回帰化したが、保存障害全体とExcel oracle照合は未完。
  進捗: Rust／CLIの原子的保存に加え、Python `Vm.run_and_save` がfork上で `run` → spillを含む再計算 → 原子的保存を行い、全成功後だけ本体VMへpublishする境界を提供するようにした。保存側はatomic rename前に閉じた一時ZIPの必須OOXML rootを構造検証する。失敗時は出力先と呼び出し元VMを維持する。CLIも成功結果をJSON／保存へ渡す前に同じ再計算境界を通る。CLIの100件failure injection matrixで既存出力保持と後続ジョブの独立性を回帰化した。保存障害全体とExcel oracle照合は未完。
- [ ] 既存のin-memory APIを暗黙に全rollbackへ変更しない。バッチ隔離と、VM上の途中変更を観測するdebug契約を分ける。大規模時はG5のメモリ予算内でcopy-on-write／変更journalを検討する。
- [ ] **V5c 中断**: 読込・VBA・再計算・保存各phaseでbudget／cancelを確認する。process隔離が必要なら所有する子processだけを停止し、partial outputと再利用不可stateを残さない。
  進捗: VMにhost-ownedの協調キャンセルフラグを追加し、ループ実行と再計算の各境界で`CANCELED`を返すようにした。CLIのsignal監視は読込だけでなくVBA実行・再計算完了まで継続し、`run`にもhost-ownedな`--cancel-file`を追加した。Python `Vm.set_cancellation` から同じ`ReadCancellation`を接続できる。保存開始時とatomic publish直前に加え、worksheet／shared stringsの各ストリーム書込み境界でもキャンセルを確認し、Drop guardで要求後の一時出力を残さない。CLIとPython `run_and_save`の事前キャンセル各100件で出力保持・caller VM非汚染・独立後続jobを回帰化し、保存中のwrite境界中断も単体回帰化した。状態汚染の全域回帰は未完。

完了条件: 100件以上のfailure injection（timeout／資源上限／型error／event連鎖／保存失敗）で、
入力不変・失敗出力非公開・次の独立jobへの汚染なし。保存の耐久性はG6のOS別gateを共用する。

### V6 — CLI／Pythonで完結する処理とテスト（P1、L7–L9と統合）

- [ ] 既存CLI／Python実行入口に型付き引数・policy・明示出力を加え、read→VBAまたはAPI編集→calculate→saveを一つの契約で利用できるようにする。新しいCLI構文は実装時に互換性審査する。
  進捗: CLIの統合fixtureでVBA実行→数式再計算→JSON投影→XLSX保存を同一ケースで検証した。V0 workflow corpusのTransfer／ByRef代表ケース、Value2式保持／消去ケース、途中書込み後のruntime errorケースを同じmanifestから実行し、CLI JSON・保存後readback・失敗時既存出力保持を確認した。`run-vba-workflow-local.py --report`でmanifest／入力／source／binary hash、local終了分類、出力hash、bounded traceの件数／digest、失敗一覧を一つのreportへ集約でき、成功・期待失敗の双方で詳細`error.kind`と`termination_class`を検証する。Python wheelでも同じ6ケースを実行し、5成功／1失敗の`last_termination_class`と失敗出力保持を確認した。Python `run_and_save` は分離VM・構造検証・reader再読込を経てpublishする。`test-workbook`へ`equals:<literal>`、`formula_present`、`formula_absent`を追加し、値・式状態の範囲assertionを回帰化した。CLI run JSONとPython `Vm.last_termination_class`へ`termination_class`（success／parse_error／io_error／setup_error／runtime_error／timeout／canceled／policy_blocked）を接続し、既存のPython例外型を維持した。Rust／CLI／Pythonの全入口・wire schema統合とExcel oracleは未完。
- [ ] `test-workbook`へ値・数式・error・変更範囲・保存再読込のassertionを追加する。既存seed／case再現を再利用する。
- [ ] fixtureごとの隔離、setup／teardown、期待エラー、選択再実行を必要順に追加する。再試行だけで成功へ塗り替えずflakyを残す。
- [ ] `docs/agent-contract.md`、Python型stub、CI tutorialを更新する。VBA不要の編集・計算例を主導線に残し、既存マクロ再利用例を追加する。

完了条件: 同じfixtureをRust／Python／CLIから処理し、正規化出力と終了分類が一致。
入力・出力hash、source、環境、traceの参照を一つのreportに集約できる。

### V7 — ファイル内VBAと複数ブック（条件付き、次候補から除外）

- [ ] **V7a source provenance**: V0監査で不足があれば、XLSM内のmodule source・codeName・参照一覧をサイズ／展開／解析budget付きで読む。保護・暗号化・未対応参照は明示拒否する。現状はopaqueな`vbaProject.bin`と別渡しソースの区別、OOXML側の`ThisWorkbook`／worksheet `codeName`とVBA relationship／署名partの存在、module identity未取得、展開サイズ上限、実行未実施を監査し、合成XLSMの回帰テストまで実装済み。
- [x] 元の`vbaProject.bin`保持とソース実行を区別する契約を実装した。p-codeのみ／ソースとの不整合／署名の扱い、VBA編集・再署名・VBProjectへの書込みは未提供で、OLE parserなしのmodule identityは推測しない。
- [ ] **V7b 複数ブック**: ファイル間転記の代表fixtureで必要性が確認された場合のみ、hostが渡したworkbook registryと安定IDを導入する。ThisWorkbook／ActiveWorkbook／active sheetの区別を校正する。
- [ ] Open／Save／Closeは任意pathアクセスに接続せず、hostが明示許可したhandle／output targetに限定する。イベント・外部リンク・中断時の整合性が成立してから有効にする。

着手gate: 対象業務、既存APIで代替できない理由、メモリ／安全性の設計を記録。
既定の外部効果拒否は維持する。Windows COMを本体の依存へ加えない。

### V8 — 性能最適化と比較判定（P2、正しさを先に固定）

- [ ] V0から採るprofileを基に、名前／member解決のcache、共通binder、一括Range転送、依存関係付き再計算、writerを順に最適化する。先にbytecode VMへ全面移行しない。
- [ ] VM単体の実行時間と、read→edit／VBA→calculate→saveのwall timeを別々に測る。VBAなしの編集・計算ベンチマークも維持する。`scripts/benchmark-vba-workflows.py`で成功workflowのCLI全体wall time（p50/p95）と出力hashを測定でき、既定5回のwarmupを除外し、`--measure-rss`では子プロセスRSSも記録できる。測定専用`benchmark_vba_vm`でparse／I/Oを除外したnative VM実行のp50／p95／生sampleをJSON化し、10,000行30反復の初回baselineを記録した。CLI全体・RSS・VMロード・VBAなしread→mutate→durable save→reloadの初回ローカル測定は実施済みだが、ps pollingのオーバーヘッドを含むwall timeは比較値とは分離し、交互順・複数規模、Excel/xlflowの比較は別測定で未完。
- [ ] 下記protocolでxlflow＋Excelと比較し、全caseの結果・速度・メモリ・欠測を出す。G5／G6の測定とhost・境界を揃える。

完了条件: 正しさのgateに加え、旧elixceeに対する非VBAの処理速度・RSS退行も審査する。
対象外や失敗を除いて見かけのspeedupを作らない。

### V9 — native配布と限定横展開（P2、browserは条件付き）

- [ ] native Rust／Pythonの3 OS CI、配布artifact、source／schema互換を確認する。新しい保証・既知差分をサポート契約へ反映する。
- [ ] B5／B11のbrowser Workerが扱える共通fixtureだけを共有する。nativeの対応増加をbrowser対応と表示しない。
- [ ] フルRust VMのWASM同梱は既存サイズgateを満たしていないため、この計画の必須項目にしない。採用にはサイズ・起動・中断・メモリの別測定が必要。

## 5. 受け入れ指標と測定protocol

以下は**目標値であり実績ではない**。V0で対象と除外理由を凍結し、途中で変える場合はprotocol版を上げて両実装を再測定する。

| 指標 | 完了判定の目標 |
|---|---|
| 意味論 | 200件以上（呼出し60、型50、配列／Range40、error／event30、object20）の必須conformanceが一致 |
| 業務処理 | 事前固定40件以上のうち95%以上が正しい出力で完遂、かつxlflow＋Excelの同じ集合に劣らない |
| 入力の独立性 | 業務40件のうち10件以上を独立作成・利用許諾済みのholdoutにする。不在なら合成試験と明記し業務優位を未判定にする |
| 安全な失敗 | V5の100件以上で入力不変・失敗出力非公開・状態汚染なし |
| cold end-to-end | 正しさ合格の固定speed suiteで2倍以上の幾何平均speedupを目標 |
| warm batch | 同suiteで1.2倍以上を目標。95%信頼区間の下限が1を超えることも必要 |
| 退行 | case別p95の10%超悪化とRSS増加を要調査。既存の非VBA編集・計算suiteにも同じ退行審査を適用 |

意味論suiteは対応保証のための必須試験、業務suiteは未対応も含む完遂率評価として分離する。
「原文のまま」は業務moduleのhash一致を意味し、別moduleの測定harnessやfixture生成コードは公開する。
制御用wrapperを書き込んだ差、ソース変換を要するcaseは通常の無変更成功から区別する。

測定方法:

1. 同じWindows host、Excelのbuild／bitness、xlflow tag、elixcee commit／release build、locale／timezone、計算mode、入力hashを記録する。Excelは独立oracleと比較対象のみに必要。
2. 読込、VBA project準備、実行、再計算、保存の境界を双方で揃える。事前push済み／未実施を混ぜない。coldは新規process／Excel、warmは独立初期stateへ戻したsessionとして別集計する。
3. 原則5回warmup、30回の交互順反復。終了前に保存まで待ち、reset費用の包含を記録する。CPU／memory／他processを記録し、静穏条件を外れた測定は理由付きで取り直す。
4. p50／p95、生sample、paired bootstrapの95%信頼区間、peak RSS／commit memoryを出す。Excelを含むprocess treeを測り、CLIだけのmemoryを比較しない。OSごとに定義の異なるmemory値を混ぜない。
5. 値・型・式・対象書式・必要OOXML partを保存再読込で照合する。ZIP全体のbyte一致や実行終了codeだけを正解判定にしない。浮動小数の許容差はcaseごとに事前指定する。
6. 不一致／timeoutにspeedupを付けない。固定speed suiteに失敗があれば総合速度の到達判定を保留する。全件表と除外母数を残す。

Mac／Linux native結果は移植性の証拠であり、そこで動かないxlflowに対する無限大speedupにはしない。
G2 Chart／Pivot、G4数式oracleの限定結果を、VBA全般の互換性へ拡大しない。

## 6. 次の候補までの小さな実装順

1. **V0**: manifest／分類／runner self-test、ByRefの現状を示す失敗fixtureを追加する。
2. **V1a＋V2a**: 内部引数metadataとslot、ByRef／ByVal／alias／型errorを共通化する。公開ASTとの接続を先に決める。
3. **V1b**: Optional／Missing／ParamArray／名前付き引数を同じbinderへ接続する。
4. **V3a**: 呼出し位置・引数診断、上限付きDebug.Printとdiagnose JSON接続を追加する。
5. **候補gate**: V1の60件、既存corpus、Python／CLI、VBAなしの編集・再計算・保存の回帰を実行。小さな転記／集計を保存再読込まで確認する。

この範囲を最初の1.0.12後の候補とし、版番号は実装後のAPI互換性審査で決める。
Excel oracle未取得ならBUILD候補として報告し、V1完了や対xlflow優位は宣言しない。
既知のChart修復問題は別の未解決gateとして記録し、全面互換の表現を避ける。
V2b以降、V7、browser横展開までを一度の候補へ詰め込まない。

## 7. 実装境界と変更管理

- 大きい`parser/mod.rs`／`vm/mod.rs`は、触る領域から`call_binding`、`runtime_error`、`trace`等へ小さく切り出す。移動だけの変更と意味論変更を分ける。
- public ASTの構造体literalや共有Variant enumの変更はRust利用者とPython／WASM schemaへ波及する。内部lowering／side tableで収められるか先に確認し、避けられない破壊変更は別versionと移行案を要する。
- 既存APIにあるイベント、object identity、budget、error sourceを再利用する。旧処理と新処理で二つの意味論を長期維持しない。
- 外部依存なし: parser／VM／fixture／runner自己検証／文書。外部環境が必要: xlflow＋Excel oracle、独立holdout、3 OS実行、静穏host測定、公開処理。後者を前者のテスト通過で代用しない。
- 保存・数式・入力安全性の回帰はVBA機能追加より優先。commit／公開は別の指示に従い、この計画の更新だけでは実施しない。

言語意味論の一次資料:
[ByRef型不一致と括弧](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/byref-argument-type-mismatch)、
[名前付き／省略可能引数](https://learn.microsoft.com/en-us/office/vba/language/concepts/getting-started/understanding-named-arguments-and-optional-arguments)、
[IsMissing](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/ismissing-function)。
文書だけで期待値を捏造せず、独立Excel測定の出典と実行条件をfixtureへ残す。
