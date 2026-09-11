import { aoa_to_sheet, book_append_sheet, book_new, cell_add_comment, encode_cell, decode_range, read, write } from '../../packages/xlsx/src/index.browser.mjs';
import { calculateWorkbook, diagnoseWorkbook } from '../../packages/xlsx/src/runtime.browser.mjs';
import { adjustPivotSourceRef, buildPivotRows } from './pivot-summary.mjs';
import { extractZipEntry, inspectZipBudget } from './zip-budget.js';
import { shiftWorkbookFormulaReferences } from './formula-shift.mjs';
import './style.css';

const copy = {
  en: {
    title: 'elixcee Playground', intro: 'Edit spreadsheet values, recalculate a formula, and download a real XLSX file — entirely in your browser.',
    recipe: 'Choose a sample', sales: 'Sales total', budget: 'Budget summary', grades: 'Class average', multi: 'Multi-sheet workbook', sheetTab: 'Spreadsheet', downloadTab: 'Excel download', input: 'Input data', calculate: 'Recalculate', download: 'Download XLSX', reset: 'Reset', downloadTitle: 'Your workbook is ready', downloadText: 'Download the edited workbook as an XLSX file. Values and the calculated formula are included.',
    formula: 'Formula', result: 'Calculated result', ready: 'Rust/WASM formula engine ready', stale: 'Values changed — recalculate to update', status: 'Status', language: 'Language', diagnostic: 'Diagnostics', operationLog: 'Operation log', dependencies: 'dependencies', formulas: 'formulas', parseErrors: 'parse errors', cycle: 'cycle', yes: 'yes', no: 'no', exported: 'Export verified', undo: 'Undo', redo: 'Redo', vbaTitle: 'VBA sandbox', vbaHint: 'Bounded Worker subset: assignments, Cells(r,c).Value, Dim, and arithmetic.', vbaRun: 'Run VBA', vbaName: 'Macro name', vbaSource: 'VBA source', vbaExample: 'Sub Demo()\n  Cells(2, 3).Value = Cells(2, 2).Value + 5\nEnd Sub', vbaDone: 'VBA completed',
    boundary: 'Browser scope: XLSX read/write, formula recalculation, diagnostics, and a bounded VBA Worker subset. Full VBA data processing runs in the native runtime.',
    guideTitle: 'What this demonstrates', guideText: 'The browser package uses the Rust/WASM core for workbook edits and formula calculation, plus a small isolated VBA Worker. Download the result and continue with the Python API or CLI for full VBA.',
    guideLink: 'Open the quick start', guideHref: '../docs/quickstart.md', upload: 'Open an XLSX or XLSM file', uploadHint: 'Choose a local .xlsx or .xlsm file. XLSM VBA is preserved as an opaque project, not executed in this browser.', imported: 'Imported workbook', selectCell: 'Select a cell', formulaBar: 'Formula bar', copy: 'Copy', paste: 'Paste', addRow: 'Add row', deleteRow: 'Delete row', addColumn: 'Add column', deleteColumn: 'Delete column', fill: 'Fill selection', error: 'Could not process this workbook.', unsupported: 'This workbook could not be opened in the browser.', externalLinksUnsupported: 'Workbooks with external links are blocked in the browser editor to prevent link loss. Use the native runtime.', unsupportedParts: 'This workbook contains OOXML parts that the browser editor cannot preserve yet. Use the native runtime.', macroUnsupported: 'The XLSM VBA project is preserved but not executed in this browser. Use the native runtime.', tooLarge: 'This file is larger than the 20 MiB browser limit.', zipRisk: 'This ZIP exceeds the browser safety budget.', structureBlocked: 'Structure editing is disabled while formulas need reference updates.',
  },
  ja: {
    title: 'elixcee Playground', intro: '表の値を編集し、数式を再計算して、本物のXLSXファイルをブラウザーだけでダウンロードできます。',
    recipe: 'サンプルを選択', sales: '売上合計', budget: '予算集計', grades: 'クラス平均', multi: '複数シート', sheetTab: 'Spreadsheet', downloadTab: 'Excelダウンロード', input: '入力データ', calculate: '再計算', download: 'XLSXをダウンロード', reset: 'リセット', downloadTitle: 'ワークブックの準備ができました', downloadText: '編集したワークブックをXLSXとしてダウンロードできます。値と計算済みの数式を含みます。',
    formula: '数式', result: '計算結果', ready: 'Rust/WASM数式エンジン準備完了', stale: '値を変更しました。再計算してください', status: '状態', language: '言語', diagnostic: '診断', operationLog: '操作ログ', dependencies: '依存関係', formulas: '数式', parseErrors: '解析エラー', cycle: '循環', yes: 'あり', no: 'なし', exported: '出力を検証済み', undo: '元に戻す', redo: 'やり直す', vbaTitle: 'VBAサンドボックス', vbaHint: 'Worker内の限定サブセット：代入、Cells(r,c).Value、Dim、算術演算。', vbaRun: 'VBAを実行', vbaName: 'マクロ名', vbaSource: 'VBAソース', vbaExample: 'Sub Demo()\n  Cells(2, 3).Value = Cells(2, 2).Value + 5\nEnd Sub', vbaDone: 'VBA実行完了',
    boundary: 'ブラウザー版の範囲：XLSXの読み書き、数式再計算、診断、限定VBA Worker。完全なVBAデータ処理はネイティブランタイムで実行します。',
    guideTitle: 'この画面で試せること', guideText: 'ブラウザー版は、ワークブック編集と数式計算にRust/WASMコアを使い、小型のVBA Workerも備えます。結果をダウンロードし、完全なVBAが必要ならPython APIまたはCLIへ進めます。',
    guideLink: 'クイックスタートを読む', guideHref: '../docs/quickstart-ja.md', upload: 'XLSX / XLSMを開く', uploadHint: 'ローカルの.xlsxまたは.xlsmを選択します。XLSMのVBAは不透明なプロジェクトとして保持しますが、ブラウザーでは実行しません。', imported: '読み込んだワークブック', selectCell: 'セルを選択', formulaBar: '数式バー', copy: 'コピー', paste: '貼り付け', addRow: '行を追加', deleteRow: '行を削除', addColumn: '列を追加', deleteColumn: '列を削除', fill: '選択範囲をフィル', error: 'ワークブックを処理できませんでした。', unsupported: 'このワークブックはブラウザーで開けませんでした。', externalLinksUnsupported: '外部リンク付きワークブックは、リンク消失を防ぐためブラウザー編集を拒否します。ネイティブランタイムを使用してください。', unsupportedParts: 'ブラウザー編集でまだ保持できないOOXML partが含まれています。ネイティブランタイムを使用してください。', macroUnsupported: 'XLSMのVBAは保持されますが、ブラウザーでは実行しません。実行にはネイティブランタイムを使用してください。', tooLarge: 'ブラウザーで扱える上限（20 MiB）を超えています。', zipRisk: 'ZIPがブラウザーの安全上限を超えています。', structureBlocked: '数式参照の更新が必要なため、構造編集を無効にしています。',
  },
  zh: {
    title: 'elixcee Playground', intro: '编辑表格数据、重新计算公式，并完全在浏览器中下载真实的 XLSX 文件。',
    recipe: '选择示例', sales: '销售总额', budget: '预算汇总', grades: '班级平均分', multi: '多工作表工作簿', sheetTab: 'Spreadsheet', downloadTab: '下载 Excel', input: '输入数据', calculate: '重新计算', download: '下载 XLSX', reset: '重置', downloadTitle: '工作簿已准备好', downloadText: '将编辑后的工作簿下载为 XLSX 文件，其中包含数据和计算后的公式。',
    formula: '公式', result: '计算结果', ready: 'Rust/WASM 公式引擎已准备就绪', stale: '数值已改变，请重新计算', status: '状态', language: '语言', diagnostic: '诊断', operationLog: '操作日志', dependencies: '依赖关系', formulas: '公式', parseErrors: '解析错误', cycle: '循环', yes: '有', no: '无', exported: '已验证导出', undo: '撤销', redo: '重做', vbaTitle: 'VBA 沙盒', vbaHint: 'Worker 限定子集：赋值、Cells(r,c).Value、Dim 和算术运算。', vbaRun: '运行 VBA', vbaName: '宏名称', vbaSource: 'VBA 源码', vbaExample: 'Sub Demo()\n  Cells(2, 3).Value = Cells(2, 2).Value + 5\nEnd Sub', vbaDone: 'VBA 已完成',
    boundary: '浏览器范围：XLSX 读写、公式计算、诊断和受限 VBA Worker。完整 VBA 数据处理在原生运行时中执行。',
    guideTitle: '此页面展示的功能', guideText: '浏览器包使用 Rust/WASM 核心完成工作簿编辑和公式计算，并提供隔离的小型 VBA Worker。下载结果后，如需完整 VBA 请继续使用 Python API 或 CLI。',
    guideLink: '打开快速开始', guideHref: '../docs/quickstart-zh.md', upload: '打开 XLSX / XLSM 文件', uploadHint: '选择本地 .xlsx 或 .xlsm 文件。XLSM 宏项目会作为不透明数据保留，但不会在浏览器中执行。', imported: '已导入工作簿', selectCell: '选择单元格', formulaBar: '公式栏', copy: '复制', paste: '粘贴', addRow: '添加行', deleteRow: '删除行', addColumn: '添加列', deleteColumn: '删除列', fill: '填充区域', error: '无法处理此工作簿。', unsupported: '此工作簿无法在浏览器中打开。', externalLinksUnsupported: '浏览器编辑器会阻止含外部链接的工作簿，以防链接丢失。请使用原生运行时。', unsupportedParts: '此工作簿包含浏览器编辑器尚无法保留的 OOXML 部件。请使用原生运行时。', macroUnsupported: 'XLSM 宏项目会被保留，但不会在浏览器中执行。请使用原生运行时执行。', tooLarge: '文件超过浏览器限制（20 MiB）。', zipRisk: 'ZIP超过浏览器安全限制。', structureBlocked: '公式需要更新引用，因此已禁用结构编辑。',
  },
};

const recipes = {
  sales: { label: 'sales', headers: ['Item', 'Units', 'Revenue'], rows: [['A', 12, 300], ['B', 8, 240], ['C', 15, 450]], formula: '=SUM(C2:C4)', resultColumn: 3 },
  budget: { label: 'budget', headers: ['Category', 'Amount'], rows: [['Hosting', 120], ['Tools', 80], ['Support', 200], ['Other', 50]], formula: '=SUM(B2:B5)', resultColumn: 2 },
  grades: { label: 'grades', headers: ['Student', 'Score'], rows: [['Aki', 84], ['Bo', 92], ['Chen', 76], ['Dana', 88]], formula: '=AVERAGE(B2:B5)', resultColumn: 2 },
  multi: { label: 'multi', headers: ['Item', 'Units', 'Revenue'], rows: [['A', 12, 300], ['B', 8, 240], ['C', 15, 450]], formula: '=SUM(C2:C4)', resultColumn: 3, extra: { name: 'Notes', rows: [['Workbook', 'Two sheets'], ['Purpose', 'Edit and calculate'], ['Runtime', 'Rust/WASM']]} },
};

const validationLabels = { en: { list: 'Dropdown list', prompt: 'Comma-separated choices', invalid: 'Enter 1–20 non-empty choices.' }, ja: { list: 'ドロップダウン', prompt: '候補をカンマ区切りで入力', invalid: '空でない候補を1〜20個入力してください。' }, zh: { list: '下拉列表', prompt: '请输入逗号分隔的选项', invalid: '请输入1到20个非空选项。' } };
const sheetLabels = {
  en: { sheetAdd: 'Add sheet', sheetDelete: 'Delete sheet', sheetRename: 'Rename sheet', sheetLast: 'A workbook must keep one sheet.', sheetReferenced: 'This sheet is referenced by a formula.', sheetNameInvalid: 'Use a unique sheet name without : \\ / ? * [ ].', freezeRow: 'Freeze top row', unfreezeRow: 'Unfreeze top row', freezeColumn: 'Freeze first column', unfreezeColumn: 'Unfreeze first column', sortAsc: 'Sort A→Z', sortDesc: 'Sort Z→A', sortBlocked: 'Sorting is disabled while formulas need reference updates.', filter: 'Filter', allValues: 'All values', clearFilter: 'Clear filter', hide: 'Hide', resize: 'Resize', resizePrompt: 'Width or height', invalidSize: 'Enter a number from 1 to 255.', unhideAll: 'Show all hidden' },
  ja: { sheetAdd: 'シート追加', sheetDelete: 'シート削除', sheetRename: 'シート名変更', sheetLast: 'ワークブックには1枚以上のシートが必要です。', sheetReferenced: 'このシートは数式から参照されています。', sheetNameInvalid: ': \\ / ? * [ ] を含まない一意のシート名を指定してください。', freezeRow: '先頭行を固定', unfreezeRow: '先頭行の固定を解除', freezeColumn: '先頭列を固定', unfreezeColumn: '先頭列の固定を解除', sortAsc: '昇順ソート', sortDesc: '降順ソート', sortBlocked: '数式参照の更新が必要なため、ソートを無効にしています。', filter: 'フィルター', allValues: 'すべて', clearFilter: '解除', hide: '非表示', resize: 'サイズ変更', resizePrompt: '幅または高さ', invalidSize: '1〜255の数値を入力してください。', unhideAll: '非表示をすべて解除' },
  zh: { sheetAdd: '添加工作表', sheetDelete: '删除工作表', sheetRename: '重命名工作表', sheetLast: '工作簿必须保留至少一个工作表。', sheetReferenced: '此工作表被公式引用。', sheetNameInvalid: '请输入不含 : \\ / ? * [ ] 且不重复的工作表名称。', freezeRow: '冻结首行', unfreezeRow: '取消冻结首行', freezeColumn: '冻结首列', unfreezeColumn: '取消冻结首列', sortAsc: '升序排序', sortDesc: '降序排序', sortBlocked: '公式需要更新引用，因此已禁用。', filter: '筛选', allValues: '全部', clearFilter: '清除筛选', hide: '隐藏', resize: '调整大小', resizePrompt: '宽度或高度', invalidSize: '请输入1到255之间的数字。', unhideAll: '显示全部隐藏' },
};
const mergeLabels = {
  en: { merge: 'Merge cells', unmerge: 'Unmerge cells', mergeBlocked: 'Select a plain, unmerged range without formulas.' },
  ja: { merge: 'セルを結合', unmerge: '結合を解除', mergeBlocked: '数式のない未結合範囲を選択してください。' },
  zh: { merge: '合并单元格', unmerge: '取消合并', mergeBlocked: '请选择没有公式且尚未合并的区域。' },
};
const nameBoxLabels = { en: 'Name box', ja: '名前ボックス', zh: '名称框' };
const definedNameLabels = { en: { add: 'Define name', prompt: 'Name for this selection', invalid: 'Use a unique name such as SalesRange.' }, ja: { add: '名前を定義', prompt: 'この範囲の名前', invalid: 'SalesRangeのような一意の名前を指定してください。' }, zh: { add: '定义名称', prompt: '为此区域命名', invalid: '请输入唯一名称，例如SalesRange。' } };
const commentLabels = { en: { add: 'Add note', prompt: 'Note for this cell', invalid: 'Enter a note.' }, ja: { add: 'メモを追加', prompt: 'このセルのメモ', invalid: 'メモを入力してください。' }, zh: { add: '添加批注', prompt: '此单元格的批注', invalid: '请输入批注。' } };
const chartLabels = { en: { bar: 'Bar chart', line: 'Line chart', edit: 'Edit chart', resize: 'Chart size', remove: 'Remove chart', prompt: 'Chart title', xAxisPrompt: 'Category axis title (optional)', yAxisPrompt: 'Value axis title (optional)', typePrompt: 'Chart type (bar or line)', legendPrompt: 'Show legend? (yes or no)', colorsPrompt: 'Series colors (comma-separated hex, optional)', sourcePrompt: 'Chart source range (e.g. A1:C10)', sizePrompt: 'Width columns, height rows', invalid: 'Select at least two columns and two data rows.', missing: 'Create a chart first.' }, ja: { bar: '棒グラフ', line: '折れ線グラフ', edit: 'グラフ編集', resize: 'グラフサイズ', remove: 'グラフ削除', prompt: 'グラフタイトル', xAxisPrompt: '項目軸タイトル（任意）', yAxisPrompt: '値軸タイトル（任意）', typePrompt: 'グラフ種類（bar または line）', legendPrompt: '凡例を表示しますか？（yes または no）', colorsPrompt: '系列色（カンマ区切りの16進カラー、任意）', sourcePrompt: 'グラフ元範囲（例：A1:C10）', sizePrompt: '幅（列数）,高さ（行数）', invalid: '2列以上、データ行2行以上を選択してください。', missing: '先にグラフを作成してください。' }, zh: { bar: '柱形图', line: '折线图', edit: '编辑图表', resize: '图表大小', remove: '删除图表', prompt: '图表标题', xAxisPrompt: '分类轴标题（可选）', yAxisPrompt: '数值轴标题（可选）', typePrompt: '图表类型（bar 或 line）', legendPrompt: '显示图例？（yes 或 no）', colorsPrompt: '系列颜色（逗号分隔的十六进制颜色，可选）', sourcePrompt: '图表数据范围（例如 A1:C10）', sizePrompt: '宽度列数,高度行数', invalid: '请选择至少两列和两行数据。', missing: '请先创建图表。' } };
const pivotLabels = { en: { add: 'Pivot summary', refresh: 'Refresh pivot', refreshed: 'Pivot summary refreshed', invalid: 'Select a category column, a numeric value column, and at least two data rows.', name: 'Pivot', aggregate: 'Aggregation (sum, count, average)', filter: 'Category filter (comma-separated; blank for all)' }, ja: { add: 'ピボット集計', refresh: 'ピボット更新', refreshed: 'ピボット集計を更新しました', invalid: 'カテゴリ列・数値列・データ行2行以上を選択してください。', name: 'Pivot', aggregate: '集計方法（sum, count, average）', filter: 'カテゴリフィルター（カンマ区切り、空欄は全件）' }, zh: { add: '透视汇总', refresh: '刷新透视', refreshed: '透视汇总已刷新', invalid: '请选择分类列、数值列和至少两行数据。', name: 'Pivot', aggregate: '汇总方式（sum, count, average）', filter: '分类筛选（逗号分隔，留空为全部）' } };
const pasteModeLabels = {
  en: { all: 'All', values: 'Values', formulas: 'Formulas', formats: 'Formats' },
  ja: { all: 'すべて', values: '値', formulas: '数式', formats: '書式' },
  zh: { all: '全部', values: '值', formulas: '公式', formats: '格式' },
};
const formatLabels = {
  en: { label: 'Number format', general: 'General', integer: 'Integer', decimal: '2 decimals', grouped: 'Thousands', percent: 'Percent', date: 'Date' },
  ja: { label: '表示形式', general: '標準', integer: '整数', decimal: '小数2桁', grouped: '桁区切り', percent: 'パーセント', date: '日付' },
  zh: { label: '数字格式', general: '常规', integer: '整数', decimal: '两位小数', grouped: '千位分隔', percent: '百分比', date: '日期' },
};
const findLabels = {
  en: { find: 'Find', replace: 'Replace', findPlaceholder: 'Text or formula', replacePlaceholder: 'Replacement', replaceOne: 'Replace next', replaceAll: 'Replace all', found: 'match(es)' },
  ja: { find: '検索', replace: '置換', findPlaceholder: '値または数式', replacePlaceholder: '置換後の文字列', replaceOne: '次を置換', replaceAll: 'すべて置換', found: '件' },
  zh: { find: '查找', replace: '替换', findPlaceholder: '值或公式', replacePlaceholder: '替换文本', replaceOne: '替换下一个', replaceAll: '全部替换', found: '个匹配' },
};

const styleLabels = { en: { bold: 'Bold', font: 'Text color', fill: 'Fill color', border: 'Bottom border', left: 'Align left', center: 'Center', right: 'Align right', wrap: 'Wrap text', copy: 'Copy format', conditional: 'Conditional format', conditionalPrompt: 'Comparison (greaterThan, lessThan, equal, between, notBetween)', thresholdPrompt: 'Threshold' }, ja: { bold: '太字', font: '文字色', fill: '塗りつぶし', border: '下罫線', left: '左揃え', center: '中央揃え', right: '右揃え', wrap: '折り返し', copy: '書式コピー', conditional: '条件付き書式', conditionalPrompt: '比較（greaterThan, lessThan, equal, between, notBetween）', thresholdPrompt: 'しきい値' }, zh: { bold: '粗体', font: '文字颜色', fill: '填充颜色', border: '下边框', left: '左对齐', center: '居中', right: '右对齐', wrap: '自动换行', copy: '复制格式', conditional: '条件格式', conditionalPrompt: '比较（greaterThan, lessThan, equal, between, notBetween）', thresholdPrompt: '阈值' } };
let language = localStorage.getItem('elixcee-playground-language') || (navigator.language.toLowerCase().startsWith('ja') ? 'ja' : navigator.language.toLowerCase().startsWith('zh') ? 'zh' : 'en');
let recipeKey = 'sales';
let values = [];
let latestBytes;
let result;
let dirty = false;
let activeTab = localStorage.getItem('elixcee-playground-tab') || 'sheet';
let workbook;
let activeSheet = 'Summary';
let selectedRef = 'A1';
let selectionStart = 'A1';
let selectionEnd = 'A1';
let selectionRanges = [{ start: 'A1', end: 'A1' }];
let dragging = false;
let dragMoved = false;
let contextAxis = 'r';
let clipboardCell = '';
let workbookName = '';
let originalBytes;
let vbaStatus = '';
let exportStatus = '';
let freezeTopRow = false;
let freezeFirstColumn = false;
let filterColumn = -1;
let filterValue = '';
let historyPast = [];
let historyFuture = [];
let operationLog = [];
let editSnapshot;
let findQuery = '';
let replaceQuery = '';
let findCursor = 0;
let clipboardData;
let clipboardRanges = [];
let pasteMode = 'all';
let fillDragging = false;
let fillSource;
let activeVbaWorker;
const HISTORY_LIMIT = 50;
const MAX_UPLOAD_BYTES = 20 * 1024 * 1024;

function resetValues() { values = recipes[recipeKey].rows.map((row) => [...row]); dirty = false; }
function resultRef() { const recipe = recipes[recipeKey]; return `${String.fromCharCode(65 + recipe.resultColumn - 1)}${recipe.rows.length + 2}`; }
function firstFormulaRef(ws) { return Object.keys(ws || {}).find((key) => /^[A-Z]+[0-9]+$/.test(key) && ws[key]?.f); }
function columnLabel(index) { let label = ''; let value = index + 1; while (value) { const remainder = (value - 1) % 26; label = String.fromCharCode(65 + remainder) + label; value = Math.floor((value - 1) / 26); } return label; }
function makeWorkbook() {
  const recipe = recipes[recipeKey];
  const rows = [recipe.headers, ...values];
  const summary = Array(recipe.headers.length).fill('');
  summary[0] = 'Result';
  summary[recipe.resultColumn - 1] = { f: recipe.formula.slice(1), t: 'n' };
  rows.push(summary);
  const workbook = book_new();
  book_append_sheet(workbook, aoa_to_sheet(rows), 'Summary');
  if (recipe.extra) book_append_sheet(workbook, aoa_to_sheet(recipe.extra.rows), recipe.extra.name);
  return workbook;
}

function makeLargeSparseWorkbook() {
  return { SheetNames: ['Large sparse'], Sheets: { 'Large sparse': { '!ref': 'A1:X200000', A1: { v: 'Sparse profile', t: 's' }, B2: { v: 42, t: 'n' }, X200000: { v: 'tail', t: 's' } } } };
}
function escapeHTML(value) { return String(value ?? '').replace(/[&<>"']/g, (char) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[char])); }
const iconPaths = {
  calculate: 'M4 3h16v18H4z M8 7h8 M8 11h2 M12 11h2 M16 11h0 M8 15h2 M12 15h2 M16 15h0 M8 19h8',
  undo: 'M9 7 4 12l5 5 M5 12h8a6 6 0 0 1 6 6', redo: 'm15 7 5 5-5 5 M19 12h-8a6 6 0 0 0-6 6',
  reset: 'M4 5v5h5 M4.5 10A8 8 0 1 1 6 17', copy: 'M8 8h11v12H8z M5 16H3V3h12v2', paste: 'M7 5h10v16H7z M10 3h4v4h-4z',
  rowAdd: 'M4 5h16M4 12h16M4 19h16 M12 9v6 M9 12h6', rowDelete: 'M4 5h16M4 12h16M4 19h16 M9 12h6',
  columnAdd: 'M5 4v16M12 4v16M19 4v16 M9 12h6 M12 9v6', columnDelete: 'M5 4v16M12 4v16M19 4v16 M9 12h6',
  freeze: 'M4 4h16v16H4z M4 9h16 M9 4v16', merge: 'M4 6h6v12H4z M14 6h6v12h-6z M10 12h4', unmerge: 'M4 6h6v12H4z M14 6h6v12h-6z M10 9h4 M10 15h4',
  sortAsc: 'M6 18V6 M6 6l-3 3 M6 6l3 3 M12 7h8 M12 12h6 M12 17h4', sortDesc: 'M6 6v12 M6 18l-3-3 M6 18l3-3 M12 7h4 M12 12h6 M12 17h8',
  alignLeft: 'M4 6h16 M4 10h11 M4 14h16 M4 18h9', alignCenter: 'M4 6h16 M7 10h10 M4 14h16 M7 18h10', alignRight: 'M4 6h16 M9 10h11 M4 14h16 M11 18h9', wrap: 'M4 6h16 M4 11h12l-3-3 M4 16h16 M4 20h8', validation: 'M5 5h14v14H5z M8 9h8 M8 13h5 M16 13h0',
  filter: 'M4 5h16l-6 7v6l-4 2v-8z', clear: 'M6 6l12 12 M18 6 6 18', definedName: 'M4 5h16v4H4z M4 12h9 M4 17h7 M17 13v7 M14 17h6', comment: 'M5 5h14v11H9l-4 3z M8 9h8 M8 12h5', bold: 'M7 5h6a4 4 0 0 1 0 8H7z M7 13h7a4 4 0 0 1 0 8H7z M7 5v16', table: 'M4 5h16v14H4z M4 10h16 M10 5v14 M16 5v14', conditional: 'M5 5h14v14H5z M8 9h8 M8 13h5',
  download: 'M12 4v11 M8 11l4 4 4-4 M5 20h14', run: 'm8 5 11 7-11 7z', chartBar: 'M5 20V10h4v10z M10 20V4h4v16z M15 20v-7h4v7z', chartLine: 'M4 18l5-6 4 3 7-9 M18 6h2v2', pivot: 'M4 5h16v14H4z M4 10h16 M10 5v14 M16 5v14 M7 14h1 M12 14h1 M18 14h1', find: 'm15 15 5 5 M17 10a7 7 0 1 1-14 0 7 7 0 0 1 14 0z', replace: 'M5 7h10 M12 4l3 3-3 3 M19 17H9 M12 14l-3 3 3 3', plus: 'M12 5v14 M5 12h14', minus: 'M5 12h14',
};
function iconSvg(name) { const path = iconPaths[name] || iconPaths.calculate; return `<svg class="ribbon-icon" viewBox="0 0 24 24" aria-hidden="true" focusable="false"><path d="${path}" /></svg>`; }
function decorateButton(button, name) { if (!button || button.dataset.iconDecorated) return; const label = button.textContent.trim(); button.classList.add('ribbon-button'); button.innerHTML = `${iconSvg(name)}<span>${escapeHTML(label)}</span>`; button.dataset.iconDecorated = 'true'; }
function buildBytes() { const ws = activeWorksheet(); if (ws) { if (freezeTopRow || freezeFirstColumn) ws['!freezePane'] = { rows: freezeTopRow ? 1 : 0, cols: freezeFirstColumn ? 1 : 0 }; else delete ws['!freezePane']; } return new Uint8Array(write(workbook, { bookType: workbook['!vbaProject'] ? 'xlsm' : 'xlsx', type: 'array' })); }
function sameBytes(left, right) { return left.length === right.length && left.every((value, index) => value === right[index]); }
function historySnapshot() { const bytes = new Uint8Array(buildBytes()); Object.defineProperty(bytes, 'charts', { value: Object.fromEntries((workbook.SheetNames || []).map((name) => [name, JSON.parse(JSON.stringify(workbook.Sheets[name]?.['!charts'] || []))])) }); Object.defineProperty(bytes, 'pivots', { value: Object.fromEntries((workbook.SheetNames || []).map((name) => [name, JSON.parse(JSON.stringify(workbook.Sheets[name]?.['!pivot'] || null))]).filter(([, pivot]) => pivot)) }); if (workbook['!vbaProject']) Object.defineProperty(bytes, 'vbaProject', { value: new Uint8Array(workbook['!vbaProject']) }); return bytes; }
function recordHistory(snapshot, operation = 'edit', context = {}) { if (!snapshot || sameBytes(snapshot, historySnapshot())) return; historyPast.push(snapshot); if (historyPast.length > HISTORY_LIMIT) historyPast.shift(); historyFuture = []; const at = new Date().toISOString(); operationLog = [...operationLog, { operation, at, sheet: context.sheet ?? activeSheet, selection: context.selection ?? selectionLabel(), sequence: operationLog.length + 1 }].slice(-100); }
function beginEdit() { if (!editSnapshot) editSnapshot = historySnapshot(); }
function commitEdit() { if (editSnapshot) { recordHistory(editSnapshot); editSnapshot = undefined; } }
function restoreHistorySnapshot(snapshot) { workbook = read(snapshot, { cellStyles: true }); if (snapshot.vbaProject) workbook['!vbaProject'] = new Uint8Array(snapshot.vbaProject); for (const [name, charts] of Object.entries(snapshot.charts || {})) if (workbook.Sheets[name] && charts.length) workbook.Sheets[name]['!charts'] = charts; for (const [name, pivot] of Object.entries(snapshot.pivots || {})) if (workbook.Sheets[name] && pivot) workbook.Sheets[name]['!pivot'] = pivot; switchActiveSheet(workbook.SheetNames.includes(activeSheet) ? activeSheet : workbook.SheetNames[0]); selectedRef = 'A1'; selectionStart = 'A1'; selectionEnd = 'A1'; selectionRanges = [{ start: 'A1', end: 'A1' }]; recalculate(); render(); }
function undo() { commitEdit(); if (!historyPast.length) return false; historyFuture.push(historySnapshot()); restoreHistorySnapshot(historyPast.pop()); return true; }
function redo() { commitEdit(); if (!historyFuture.length) return false; historyPast.push(historySnapshot()); restoreHistorySnapshot(historyFuture.pop()); return true; }
function cellValue(book, ref) { return book?.Sheets?.[activeSheet]?.[ref]?.v ?? ''; }
function activeWorksheet() { return workbook?.Sheets?.[activeSheet]; }
function switchActiveSheet(name) { activeSheet = name; filterColumn = -1; filterValue = ''; syncFreezeState(); }
function syncFreezeState() { const pane = activeWorksheet()?.['!freezePane']; freezeTopRow = Number(pane?.rows) > 0; freezeFirstColumn = Number(pane?.cols) > 0; }
function worksheetBounds(ws) {
  if (ws?.['!ref']) return decode_range(ws['!ref']);
  let maxRow = 9; let maxCol = 5;
  for (const ref of Object.keys(ws || {})) { if (!/^[A-Z]+[0-9]+$/.test(ref)) continue; const range = decode_range(`${ref}:${ref}`); maxRow = Math.max(maxRow, range.e.r); maxCol = Math.max(maxCol, range.e.c); }
  return { s: { r: 0, c: 0 }, e: { r: maxRow, c: maxCol } };
}
function mergeFor(ws, row, col) { return (ws?.['!merges'] || []).find((merge) => row >= merge.s.r && row <= merge.e.r && col >= merge.s.c && col <= merge.e.c); }
function cellText(cell) { return cell?.f ? `=${cell.f}` : (cell?.v ?? ''); }
function validationFor(ws, ref) { return (ws?.['!dataValidations'] || []).find((validation) => (validation.type || 'list') === 'list' && (validation.sqref || []).some((range) => { try { const bounds = decode_range(range); const point = cellPosition(ref); return point.r >= bounds.s.r && point.r <= bounds.e.r && point.c >= bounds.s.c && point.c <= bounds.e.c; } catch { return false; } })); }
function validationChoices(validation) { const formula = String(validation?.formula1 ?? validation?.formula ?? ''); const raw = formula.replace(/^"|"$/g, ''); return raw.split(',').map((choice) => choice.trim()).filter(Boolean); }
function conditionalFormatsFor(ws, ref) { const point = cellPosition(ref); return (ws?.['!conditionalFormats'] || []).map((rule, index) => ({ rule, index })).filter(({ rule }) => (rule.sqref || []).some((range) => { try { const bounds = decode_range(range); return point.r >= bounds.s.r && point.r <= bounds.e.r && point.c >= bounds.s.c && point.c <= bounds.e.c; } catch { return false; } })).sort((left, right) => (Number.isInteger(left.rule.priority) ? left.rule.priority : Number.MAX_SAFE_INTEGER) - (Number.isInteger(right.rule.priority) ? right.rule.priority : Number.MAX_SAFE_INTEGER) || left.index - right.index).map(({ rule }) => rule); }
function conditionalFormatFor(ws, ref) { return conditionalFormatsFor(ws, ref)[0]; }
function conditionalLiteral(value) { const raw = String(value ?? '').trim(); if (/^TRUE$/i.test(raw)) return true; if (/^FALSE$/i.test(raw)) return false; if (/^-?(?:\d+(?:\.\d*)?|\.\d+)$/.test(raw)) return Number(raw); if (raw.startsWith('"') && raw.endsWith('"')) return raw.slice(1, -1).replace(/""/g, '"'); return undefined; }
function translateConditionalFormula(formula, originRef, targetRef) { const origin = cellPosition(originRef); const target = cellPosition(targetRef); const rowDelta = target.r - origin.r; const colDelta = target.c - origin.c; return String(formula ?? '').replace(/(^|[^A-Za-z0-9_])([$]?)([A-Z]{1,3})([$]?)(\d+)/gi, (match, prefix, colAbsolute, colText, rowAbsolute, rowText, offset) => { const before = String(formula).slice(0, offset + prefix.length); let quoted = false; for (let i = 0; i < before.length; i += 1) { if (before[i] !== '"') continue; if (before[i + 1] === '"') { i += 1; continue; } quoted = !quoted; } if (quoted) return match; const point = { c: colText.toUpperCase().split('').reduce((sum, char) => sum * 26 + char.charCodeAt(0) - 65, 0), r: Number(rowText) - 1 }; if (!colAbsolute) point.c += colDelta; if (!rowAbsolute) point.r += rowDelta; if (point.c < 0 || point.r < 0) return `${prefix}#REF!`; return `${prefix}${colAbsolute ? '$' : ''}${encode_cell({ r: point.r, c: point.c }).replace(/\d+$/, '')}${rowAbsolute ? '$' : ''}${point.r + 1}`; }); }
function conditionalArguments(body) { const args = []; let start = 0; let depth = 0; let quoted = false; for (let i = 0; i < body.length; i += 1) { const char = body[i]; if (char === '"') { if (body[i + 1] === '"') { i += 1; continue; } quoted = !quoted; } else if (!quoted && char === '(') depth += 1; else if (!quoted && char === ')') depth -= 1; else if (!quoted && depth === 0 && char === ',') { args.push(body.slice(start, i).trim()); start = i + 1; } } if (quoted || depth !== 0) return []; args.push(body.slice(start).trim()); return args; }
function conditionalExpressionMatches(ws, ref, formula, rule) { let expression = String(formula ?? '').trim().replace(/^=/, ''); const originRef = String(rule?.sqref?.[0] || '').split(':')[0] || ref; expression = translateConditionalFormula(expression, originRef, ref); const functionCall = /^(AND|OR|NOT)\((.*)\)$/i.exec(expression); if (functionCall) { const args = conditionalArguments(functionCall[2]); if (!args.length || (functionCall[1].toUpperCase() === 'NOT' && args.length !== 1)) return false; const values = args.map((arg) => conditionalExpressionMatches(ws, ref, arg, { ...rule, sqref: [ref] })); return functionCall[1].toUpperCase() === 'AND' ? values.every(Boolean) : functionCall[1].toUpperCase() === 'OR' ? values.some(Boolean) : !values[0]; } const comparison = /^((?:\$?[A-Z]{1,3}\$?\d+))\s*(<>|>=|<=|=|>|<)\s*(.+)$/i.exec(expression); if (!comparison) { const literal = conditionalLiteral(expression); return literal === undefined ? false : Boolean(literal); } const left = ws?.[comparison[1].replace(/\$/g, '').toUpperCase()]?.v; const right = conditionalLiteral(comparison[3]); if (right === undefined) return false; const [lhs, rhs] = typeof left === 'number' && typeof right === 'number' ? [left, right] : [String(left ?? ''), String(right)]; switch (comparison[2]) { case '=': return lhs === rhs; case '<>': return lhs !== rhs; case '>': return lhs > rhs; case '<': return lhs < rhs; case '>=': return lhs >= rhs; case '<=': return lhs <= rhs; default: return false; } }
function conditionalRuleMatches(ws, ref, rule) { const cell = ws?.[ref]; if (!cell || !['cellIs', 'expression'].includes(rule.type || 'cellIs')) return false; if (rule.type === 'expression') return conditionalExpressionMatches(ws, ref, rule.formula, rule); const value = Number(cell.v); const threshold = Number(String(rule.formula ?? '').replace(/^=/, '')); const threshold2 = Number(String(rule.formula2 ?? '').replace(/^=/, '')); if (!Number.isFinite(value) || !Number.isFinite(threshold)) return false; const operator = rule.operator || 'greaterThan'; return operator === 'greaterThan' ? value > threshold : operator === 'greaterThanOrEqual' ? value >= threshold : operator === 'lessThan' ? value < threshold : operator === 'lessThanOrEqual' ? value <= threshold : operator === 'equal' ? value === threshold : operator === 'notEqual' ? value !== threshold : operator === 'between' ? Number.isFinite(threshold2) && value >= Math.min(threshold, threshold2) && value <= Math.max(threshold, threshold2) : operator === 'notBetween' ? Number.isFinite(threshold2) && (value < Math.min(threshold, threshold2) || value > Math.max(threshold, threshold2)) : false; }
function conditionalFormatStyle(ws, ref) { let css = ''; for (const rule of conditionalFormatsFor(ws, ref)) { if (!conditionalRuleMatches(ws, ref, rule)) continue; const dxf = rule.dxf || {}; const fill = dxf.fill?.fgColor?.rgb || dxf.fill?.fg?.rgb; const color = dxf.font?.color?.rgb; css += `${fill ? `background:#${String(fill).slice(-6)};` : ''}${color ? `color:#${String(color).slice(-6)};` : ''}${dxf.font?.bold ? 'font-weight:700;' : ''}${dxf.font?.italic ? 'font-style:italic;' : ''}${dxf.font?.underline ? 'text-decoration:underline;' : ''}`; if (rule.stopIfTrue === true) break; } return css; }
function commentTextFor(cell) { return Array.isArray(cell?.c) ? cell.c.map((comment) => String(comment?.t ?? '')).filter(Boolean).join('\n') : ''; }
function normalizeChartColors(value) { return String(value ?? '').split(',').map((color) => color.trim()).filter(Boolean).map((color) => color.startsWith('#') ? color.toLowerCase() : `#${color.toLowerCase()}`).filter((color) => /^#[0-9a-f]{6}$/.test(color)).slice(0, 65); }
function applyChart(type) { const range = selectionBounds(); if (range.e.c - range.s.c < 1 || range.e.r - range.s.r < 1) return false; const title = window.prompt(chartLabels[language].prompt, type === 'line' ? 'Trend' : 'Values'); if (title === null) return false; const xAxisTitle = window.prompt(chartLabels[language].xAxisPrompt, '') ?? ''; const yAxisTitle = window.prompt(chartLabels[language].yAxisPrompt, '') ?? ''; const colors = normalizeChartColors(window.prompt(chartLabels[language].colorsPrompt, '#3856d9, #1c8b52, #d97706, #b33a8a, #64748b') ?? ''); const before = historySnapshot(); const ws = activeWorksheet(); ws['!charts'] = [...(ws['!charts'] || []), { type, title: title.trim() || (type === 'line' ? 'Trend' : 'Values'), xAxisTitle: xAxisTitle.trim(), yAxisTitle: yAxisTitle.trim(), colors, ref: primarySelectionLabel() }]; recordHistory(before, 'chart-create'); dirty = true; return true; }
function editChart() { const charts = activeWorksheet()?.['!charts']; const chart = charts?.[charts.length - 1]; if (!chart) return false; const title = window.prompt(chartLabels[language].prompt, chart.title || 'Chart'); if (title === null) return false; const type = window.prompt(chartLabels[language].typePrompt, chart.type || 'bar')?.trim().toLowerCase(); if (!['bar', 'line'].includes(type)) return false; const legendAnswer = window.prompt(chartLabels[language].legendPrompt, chart.legend === false ? 'no' : 'yes'); if (legendAnswer === null) return false; const showLegend = ['yes', 'y', 'true'].includes(legendAnswer.trim().toLowerCase()); if (!['yes', 'y', 'true', 'no', 'n', 'false'].includes(legendAnswer.trim().toLowerCase())) return false; const source = window.prompt(chartLabels[language].sourcePrompt, chart.ref || primarySelectionLabel()); if (source === null) return false; const xAxisTitle = window.prompt(chartLabels[language].xAxisPrompt, chart.xAxisTitle || '') ?? ''; const yAxisTitle = window.prompt(chartLabels[language].yAxisPrompt, chart.yAxisTitle || '') ?? ''; const colors = normalizeChartColors(window.prompt(chartLabels[language].colorsPrompt, (chart.colors || []).join(', ')) ?? ''); let range; try { range = decode_range(source.trim()); } catch { return false; } if (range.e.c - range.s.c < 1 || range.e.r - range.s.r < 1) return false; const before = historySnapshot(); chart.title = title.trim() || 'Chart'; chart.type = type; chart.xAxisTitle = xAxisTitle.trim(); chart.yAxisTitle = yAxisTitle.trim(); chart.legend = showLegend; chart.colors = colors; chart.ref = `${encode_cell(range.s)}:${encode_cell(range.e)}`; recordHistory(before, 'chart-edit'); dirty = true; return true; }
function removeChart() { const ws = activeWorksheet(); const charts = ws?.['!charts']; if (!charts?.length) return false; const before = historySnapshot(); charts.pop(); if (!charts.length) delete ws['!charts']; recordHistory(before, 'chart-remove'); dirty = true; return true; }
function resizeChart() { const charts = activeWorksheet()?.['!charts']; const chart = charts?.[charts.length - 1]; if (!chart) return false; const value = window.prompt(chartLabels[language].sizePrompt, `${chart.widthCols || 7},${chart.heightRows || 16}`); if (value === null) return false; const [width, height] = value.split(',').map(Number); if (!Number.isFinite(width) || !Number.isFinite(height) || width < 4 || width > 20 || height < 8 || height > 40) return false; const before = historySnapshot(); chart.widthCols = width; chart.heightRows = height; recordHistory(before, 'chart-resize'); dirty = true; return true; }
function createPivotSummary() {
  const ws = activeWorksheet(); const range = selectionBounds(); if (!ws || range.e.c - range.s.c < 1 || range.e.r - range.s.r < 2) return false;
  const requested = window.prompt(pivotLabels[language].aggregate, 'sum')?.trim().toLowerCase() || 'sum'; if (!['sum', 'count', 'average'].includes(requested)) return false;
  const categoryFilter = window.prompt(pivotLabels[language].filter, ''); if (categoryFilter === null) return false; const normalizedFilter = categoryFilter.trim().toLocaleLowerCase();
  const rows = buildPivotRows(ws, range, requested, normalizedFilter); if (!rows) return false; const before = historySnapshot(); let name = pivotLabels[language].name; let index = 1; while (workbook.SheetNames.includes(name)) name = `${pivotLabels[language].name}${++index}`; const pivotSheet = aoa_to_sheet(rows); pivotSheet['!pivot'] = { sourceSheet: activeSheet, sourceRef: primarySelectionLabel(), aggregate: requested, filter: normalizedFilter }; book_append_sheet(workbook, pivotSheet, name); activeSheet = name; selectedRef = 'A1'; setSingleSelection('A1'); recordHistory(before, 'pivot-create'); dirty = true; return true;
}
function refreshPivotSummaries() {
  let refreshed = 0;
  for (const name of workbook.SheetNames || []) {
    const pivotSheet = workbook.Sheets[name]; const pivot = pivotSheet?.['!pivot']; if (!pivot) continue;
    const source = workbook.Sheets[pivot.sourceSheet]; if (!source) continue;
    let range; try { range = decode_range(pivot.sourceRef); } catch { continue; }
    const rows = buildPivotRows(source, range, pivot.aggregate, pivot.filter); if (!rows) continue;
    const next = aoa_to_sheet(rows); next['!pivot'] = pivot;
    workbook.Sheets[name] = next;
    refreshed += 1;
  }
  return refreshed;
}
function chartPreview(ws) {
  const charts = (ws?.['!charts'] || []).filter((candidate) => candidate && typeof candidate.ref === 'string'); if (!charts.length) return '';
  return charts.map((chart) => chartPreviewOne(ws, chart)).filter(Boolean).join('');
}
function chartPreviewOne(ws, chart) {
  let range; try { range = decode_range(chart.ref); } catch { return ''; }
  const labels = []; const series = [];
  for (let col = range.s.c + 1; col <= range.e.c; col += 1) {
    const points = [];
    for (let row = range.s.r + 1; row <= range.e.r; row += 1) {
      const value = Number(ws?.[encode_cell({ r: row, c: col })]?.v); if (Number.isFinite(value)) points.push({ row, value });
      if (col === range.s.c + 1) labels.push(String(ws?.[encode_cell({ r: row, c: range.s.c })]?.v ?? ''));
    }
    if (points.length) series.push({ name: String(ws?.[encode_cell({ r: range.s.r, c: col })]?.v ?? `Series ${series.length + 1}`), points });
  }
  if (!series.length) return '';
  const max = Math.max(1, ...series.flatMap((entry) => entry.points.map((point) => point.value)));
  const colors = ['#3856d9', '#1c8b52', '#d97706', '#b33a8a', '#64748b']; const colorFor = (index) => /^#[0-9a-f]{6}$/i.test(String(chart.colors?.[index] || '')) ? String(chart.colors[index]).toLowerCase() : colors[index % colors.length];
  const groups = labels.map((label, index) => { const width = Math.max(6, 180 / Math.max(1, labels.length * series.length)); return series.map((entry, seriesIndex) => { const point = entry.points.find((candidate) => candidate.row === range.s.r + 1 + index); if (!point) return ''; const x = 34 + index * (230 / Math.max(1, labels.length)) + seriesIndex * width; const height = Math.max(2, 105 * point.value / max); return chart.type === 'line' ? '' : `<rect x="${x}" y="${145 - height}" width="${width - 1}" height="${height}" rx="1" fill="${colorFor(seriesIndex)}"/>`; }).join('') + `<text x="${34 + index * (230 / Math.max(1, labels.length)) + 3}" y="160">${escapeHTML(label.slice(0, 10))}</text>`; }).join('');
  const lines = chart.type === 'line' ? series.map((entry, seriesIndex) => { const points = entry.points.map((point, index) => `${34 + index * (230 / Math.max(1, entry.points.length - 1))},${145 - Math.max(2, 105 * point.value / max)}`).join(' '); return `<polyline points="${points}" fill="none" stroke="${colorFor(seriesIndex)}" stroke-width="2"/>`; }).join('') : groups;
  const legend = chart.legend === false ? '' : series.map((entry, index) => `<span style="color:${colorFor(index)}">${escapeHTML(entry.name)}</span>`).join(' · ');
  const axisLabels = [chart.xAxisTitle, chart.yAxisTitle].filter(Boolean).map((label) => escapeHTML(label)).join(' · ');
  return `<div class="chart-preview" aria-label="${escapeHTML(chart.title || 'Chart')}"><strong>${escapeHTML(chart.title || 'Chart')}</strong>${axisLabels ? `<small class="chart-axes">${axisLabels}</small>` : ''}<small>${legend}</small><svg viewBox="0 0 280 175" role="img" aria-label="${escapeHTML(chart.title || 'Chart')}"><line x1="28" y1="145" x2="270" y2="145"/><line x1="28" y1="20" x2="28" y2="145"/>${lines}</svg></div>`;
}
function definedNames() { if (!workbook.Workbook) workbook.Workbook = {}; if (!Array.isArray(workbook.Workbook.Names)) workbook.Workbook.Names = []; return workbook.Workbook.Names; }
function definedNameFor(name) { return definedNames().find((entry) => entry.Name.toLocaleLowerCase() === String(name).toLocaleLowerCase()); }
function resolveDefinedName(entry) {
  const raw = String(entry?.Ref ?? '').replace(/\$/g, ''); const bang = raw.lastIndexOf('!');
  const qualifier = Number.isInteger(entry?.Sheet) ? workbook.SheetNames[entry.Sheet] : (bang >= 0 ? raw.slice(0, bang).replace(/^'|'$/g, '').replace(/''/g, "'") : activeSheet);
  const ref = bang >= 0 ? raw.slice(bang + 1) : raw;
  if (!ref || !workbook.SheetNames.includes(qualifier)) return undefined;
  try { const range = decode_range(ref.includes(':') ? ref : `${ref}:${ref}`); return { sheet: qualifier, range }; } catch { return undefined; }
}
function applyDefinedName(name) { const trimmed = String(name ?? '').trim(); const range = primarySelectionLabel(); if (!/^[A-Za-z_][A-Za-z0-9_.]*$/.test(trimmed) || /^[A-Za-z]+[0-9]+$/.test(trimmed) || definedNames().some((entry) => entry.Name.toLocaleLowerCase() === trimmed.toLocaleLowerCase())) return false; const before = historySnapshot(); definedNames().push({ Name: trimmed, Ref: `'${activeSheet.replace(/'/g, "''")}'!${range}` }); recordHistory(before, 'defined-name'); dirty = true; return true; }
function selectedRangeLabels() { return selectionRanges.map(selectionRangeLabel); }
function applyListValidation() { const ws = activeWorksheet(); const choices = window.prompt(validationLabels[language].prompt, 'Open,In progress,Done')?.split(',').map((choice) => choice.trim()).filter(Boolean) || []; if (!choices.length || choices.length > 20 || choices.some((choice) => choice.length > 64)) return false; const ranges = selectedRangeLabels(); const before = historySnapshot(); ws['!dataValidations'] = [...(ws['!dataValidations'] || []).filter((validation) => !(validation.sqref || []).some((range) => ranges.includes(range))), { type: 'list', sqref: ranges, allowBlank: true, formula1: `"${choices.join(',')}"` }]; recordHistory(before, 'validation'); dirty = true; return true; }
function applyConditionalFormat() { const ws = activeWorksheet(); const operator = window.prompt(styleLabels[language].conditionalPrompt, 'greaterThan')?.trim(); const threshold = window.prompt(styleLabels[language].thresholdPrompt, '0')?.trim(); const threshold2 = ['between', 'notBetween'].includes(operator) ? window.prompt(styleLabels[language].thresholdPrompt, '10')?.trim() : undefined; if (!['greaterThan', 'lessThan', 'equal', 'between', 'notBetween'].includes(operator) || !/^[-+]?\d+(?:\.\d+)?$/.test(threshold) || (threshold2 !== undefined && !/^[-+]?\d+(?:\.\d+)?$/.test(threshold2))) return false; const ranges = selectedRangeLabels(); const before = historySnapshot(); ws['!conditionalFormats'] = [...(ws['!conditionalFormats'] || []).filter((rule) => !(rule.sqref || []).some((range) => ranges.includes(range))), { type: 'cellIs', operator, sqref: ranges, formula: threshold, ...(threshold2 === undefined ? {} : { formula2: threshold2 }), dxf: { fill: { patternType: 'solid', fgColor: { rgb: 'FFF2CC' } }, font: { color: { rgb: '9C0006' }, bold: true } } }]; recordHistory(before, 'conditional-format'); dirty = true; return true; }
function applyTable() { const ws = activeWorksheet(); const range = selectionBounds(); if (!ws || range.s.r >= range.e.r || range.s.c > range.e.c) return false; const overlaps = (candidate) => { try { const other = decode_range(candidate.ref); return range.s.r <= other.e.r && range.e.r >= other.s.r && range.s.c <= other.e.c && range.e.c >= other.s.c; } catch { return true; } }; if ((ws['!tables'] || []).some(overlaps)) return false; const defaultName = `Table${Object.values(workbook.Sheets || {}).reduce((count, sheet) => count + (sheet?.['!tables']?.length || 0), 0) + 1}`; const name = window.prompt(language === 'ja' ? 'テーブル名' : language === 'zh' ? '表名称' : 'Table name', defaultName)?.trim(); if (!name || !/^[A-Za-z_][A-Za-z0-9_.]{0,254}$/.test(name) || /^[A-Za-z]+[0-9]+$/.test(name) || Object.values(workbook.Sheets || {}).some((sheet) => sheet?.['!tables']?.some((table) => String(table.name || table.displayName).toLowerCase() === name.toLowerCase()))) return false; const before = historySnapshot(); ws['!tables'] = [...(ws['!tables'] || []), { name, displayName: name, ref: primarySelectionLabel() }]; recordHistory(before, 'table-create'); dirty = true; return true; }
function tableAtCell(ws, ref) { const point = cellPosition(ref); const table = (ws?.['!tables'] || []).find((candidate) => { try { const range = decode_range(candidate.ref); return point.r >= range.s.r && point.r <= range.e.r && point.c >= range.s.c && point.c <= range.e.c; } catch { return false; } }); if (!table) return undefined; try { return { table, range: decode_range(table.ref) }; } catch { return undefined; } }
function setTableValueFilter(ws, column, value) { const match = tableAtCell(ws, selectedRef); if (!match || match.range.s.r === match.range.e.r || column < match.range.s.c || column > match.range.e.c) return false; const filterRef = match.table.autoFilterRef || match.table.ref; let filterRange; try { filterRange = decode_range(filterRef); } catch { filterRange = match.range; } const columns = (match.table.autoFilterColumns || []).filter((entry) => Number(entry?.colId) !== column - filterRange.s.c); if (value) columns.push({ colId: column - filterRange.s.c, hiddenButton: false, showButton: true, criteria: { kind: 'values', values: [String(value)] } }); match.table.autoFilterRef = filterRef; match.table.autoFilterColumns = columns; return true; }
function syncTableFilterState(ws) { if (filterColumn >= 0) return; const match = tableAtCell(ws, selectedRef); const entry = match?.table?.autoFilterColumns?.find((candidate) => candidate?.criteria?.kind === 'values' && candidate.criteria.values?.length); if (!entry) return; filterColumn = match.range.s.c + Number(entry.colId); filterValue = String(entry.criteria.values[0]); }
function workerCells(ws) { return Object.fromEntries(Object.keys(ws || {}).filter((key) => /^[A-Z]+[0-9]+$/.test(key)).map((key) => [key, cellText(ws[key])])); }
function dependencyCount(ws) { return Object.keys(ws || {}).filter((key) => /^[A-Z]+[0-9]+$/.test(key) && ws[key]?.f).reduce((count, key) => count + ((ws[key].f.match(/\b[A-Z]{1,3}\d+(?::[A-Z]{1,3}\d+)?\b/gi) || []).length), 0); }
function diagnosticText(t) { const d = result?.diagnostic || {}; return `${t.diagnostic}: ${d.formulaCount ?? 0} ${t.formulas}, ${dependencyCount(activeWorksheet())} ${t.dependencies}, ${d.formulaParseErrors ?? 0} ${t.parseErrors}, ${t.cycle} ${d.hasFormulaCycle ? t.yes : t.no}`; }
function cellPosition(ref) { return decode_range(`${ref}:${ref}`).s; }
function selectionBounds() { const start = cellPosition(selectionStart); const end = cellPosition(selectionEnd); return { s: { r: Math.min(start.r, end.r), c: Math.min(start.c, end.c) }, e: { r: Math.max(start.r, end.r), c: Math.max(start.c, end.c) } }; }
function selectionRangeBounds(range) { const start = cellPosition(range.start); const end = cellPosition(range.end); return { s: { r: Math.min(start.r, end.r), c: Math.min(start.c, end.c) }, e: { r: Math.max(start.r, end.r), c: Math.max(start.c, end.c) } }; }
function selectionRangeLabel(range) { const bounds = selectionRangeBounds(range); return encode_cell(bounds.s) === encode_cell(bounds.e) ? encode_cell(bounds.s) : `${encode_cell(bounds.s)}:${encode_cell(bounds.e)}`; }
function primarySelectionLabel() { return selectionRangeLabel({ start: selectionStart, end: selectionEnd }); }
function selectionLabel() { return selectionRanges.map(selectionRangeLabel).join(','); }
function setSingleSelection(start, end = start) { selectionStart = start; selectionEnd = end; selectionRanges = [{ start, end }]; }
function selectedCellRefs() { const refs = new Set(); for (const selection of selectionRanges) { const range = selectionRangeBounds(selection); for (let row = range.s.r; row <= range.e.r; row += 1) for (let col = range.s.c; col <= range.e.c; col += 1) refs.add(encode_cell({ r: row, c: col })); } return refs; }
function moveSelection(rowDelta, colDelta, extend = false) { const point = cellPosition(selectedRef); const bounds = worksheetBounds(activeWorksheet()); const row = Math.max(0, Math.min(bounds.e.r, point.r + rowDelta)); const col = Math.max(0, Math.min(bounds.e.c, point.c + colDelta)); selectedRef = encode_cell({ r: row, c: col }); if (extend) { selectionEnd = selectedRef; selectionRanges = [{ start: selectionStart, end: selectionEnd }]; } else setSingleSelection(selectedRef); render(); document.querySelector(`[data-ref="${selectedRef}"]`)?.focus(); }
function applyNumberFormat(format) { const ws = activeWorksheet(); const before = historySnapshot(); let changed = false; for (const ref of selectedCellRefs()) { const current = ws?.[ref]; if (!current && format === 'General') continue; const cell = current || (ws[ref] = { v: '', t: 's' }); const next = format === 'General' ? undefined : format; if (cell.z !== next) { if (next) cell.z = next; else delete cell.z; changed = true; } } if (changed) { recordHistory(before, 'number-format'); dirty = true; } return changed; }
function applyCellStyle(kind, value) { const ws = activeWorksheet(); const before = historySnapshot(); let changed = false; for (const ref of selectedCellRefs()) { const cell = ws?.[ref] || (ws[ref] = { v: '', t: 's' }); const style = { ...(cell.s || {}) }; if (kind === 'bold') style.font = { ...(style.font || {}), bold: !style.font?.bold }; if (kind === 'font') style.font = { ...(style.font || {}), color: { rgb: value } }; if (kind === 'fill') style.fill = { patternType: 'solid', fgColor: { rgb: value } }; if (kind === 'border') style.border = { ...(style.border || {}), bottom: { style: 'thin', color: { rgb: value } } }; if (kind === 'align') style.alignment = { ...(style.alignment || {}), horizontal: value }; if (kind === 'wrap') style.alignment = { ...(style.alignment || {}), wrapText: !style.alignment?.wrapText }; cell.s = style; changed = true; } if (changed) { recordHistory(before, `style-${kind}`); dirty = true; } return changed; }
function copyFormatToSelection() { const ws = activeWorksheet(); const source = ws?.[selectionStart]; if (!ws || !source) return false; const sourceStyle = source.s ? cloneCell(source.s) : undefined; const sourceFormat = source.z; const before = historySnapshot(); let changed = false; for (const ref of selectedCellRefs()) { const cell = ws[ref] || { v: '', t: 's' }; if (sourceStyle) cell.s = cloneCell(sourceStyle); else delete cell.s; if (sourceFormat) cell.z = sourceFormat; else delete cell.z; ws[ref] = cell; changed = true; } if (changed) { recordHistory(before, 'format-copy'); dirty = true; } return changed; }
function findCells(query) { if (!query) return []; const matches = []; for (const name of workbook?.SheetNames || []) for (const [ref, cell] of Object.entries(workbook.Sheets[name] || {})) if (/^[A-Z]+[0-9]+$/.test(ref) && String(cellText(cell)).toLocaleLowerCase().includes(query.toLocaleLowerCase())) matches.push({ name, ref }); return matches; }
function replaceNext() { const matches = findCells(findQuery); if (!matches.length) return false; const match = matches[findCursor % matches.length]; switchActiveSheet(match.name); selectedRef = match.ref; setSingleSelection(match.ref); const current = cellText(workbook.Sheets[match.name][match.ref]); const before = historySnapshot(); const index = current.toLocaleLowerCase().indexOf(findQuery.toLocaleLowerCase()); const replacement = `${current.slice(0, index)}${replaceQuery}${current.slice(index + findQuery.length)}`; workbook.Sheets[match.name][match.ref] = workbook.Sheets[match.name][match.ref].f ? { ...workbook.Sheets[match.name][match.ref], f: replacement.replace(/^=/, '') } : { ...workbook.Sheets[match.name][match.ref], v: replacement, t: 's' }; recordHistory(before, 'replace-one'); dirty = true; findCursor += 1; return true; }
function replaceAll() { const matches = findCells(findQuery); if (!matches.length) return 0; const before = historySnapshot(); let changed = 0; for (const match of matches) { const cell = workbook.Sheets[match.name][match.ref]; const current = cellText(cell); const replacement = current.replace(new RegExp(findQuery.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'gi'), replaceQuery); if (replacement !== current) { workbook.Sheets[match.name][match.ref] = cell.f ? { ...cell, f: replacement.replace(/^=/, '') } : { ...cell, v: replacement, t: 's' }; changed += 1; } } if (changed) { recordHistory(before, 'replace-all'); dirty = true; } return changed; }
function cloneCell(cell) { return cell ? JSON.parse(JSON.stringify(cell)) : undefined; }
function pasteCellsAt(mode, rows, origin) {
  const ws = activeWorksheet(); let changed = false;
  rows.forEach((row, rowOffset) => row.forEach((source, colOffset) => {
    const ref = encode_cell({ r: origin.r + rowOffset, c: origin.c + colOffset }); const target = ws?.[ref];
    if (mode === 'formats') { if (!source) return; const next = { ...(target || {}) }; if (source.s) next.s = cloneCell(source.s); else delete next.s; if (source.z) next.z = source.z; else delete next.z; ws[ref] = next; changed = true; return; }
    if (mode === 'formulas') { if (!source?.f) return; ws[ref] = { ...(target || {}), f: source.f, t: source.t || 'n' }; delete ws[ref].v; changed = true; return; }
    if (mode === 'values') { if (!source) { if (target) { delete ws[ref].f; delete ws[ref].v; changed = true; } return; } const next = { ...(target || {}), v: source.v, t: source.t || (typeof source.v === 'number' ? 'n' : 's') }; delete next.f; ws[ref] = next; changed = true; return; }
    if (source) ws[ref] = cloneCell(source); else if (target) delete ws[ref]; changed = true;
  }));
  return changed;
}
function pasteCells(mode, rows) {
  const before = historySnapshot(); let changed = false;
  const targets = selectionRanges.length > 1 ? selectionRanges.map(selectionRangeBounds) : [{ s: cellPosition(selectionStart) }];
  const payloads = clipboardRanges.length === targets.length ? clipboardRanges : targets.map(() => rows);
  targets.forEach((target, index) => { if (payloads[index]) changed = pasteCellsAt(mode, payloads[index], target.s) || changed; });
  if (changed) { recordHistory(before, 'paste'); dirty = true; }
  return changed;
}
async function copySelection() { const ranges = selectionRanges.length ? selectionRanges : [{ start: selectionStart, end: selectionEnd }]; clipboardRanges = ranges.map((selection) => { const range = selectionRangeBounds(selection); return Array.from({ length: range.e.r - range.s.r + 1 }, (_, r) => Array.from({ length: range.e.c - range.s.c + 1 }, (_, c) => cloneCell(activeWorksheet()?.[encode_cell({ r: range.s.r + r, c: range.s.c + c })]))); }); clipboardData = clipboardRanges[0] || []; clipboardCell = clipboardRanges.map((rows) => rows.map((row) => row.map((cell) => cellText(cell)).join('\t')).join('\n')).join('\n\n'); const clipboardHtml = clipboardRanges.map((rows) => `<table><tbody>${rows.map((row) => `<tr>${row.map((cell) => `<td>${escapeHTML(cellText(cell))}</td>`).join('')}</tr>`).join('')}</tbody></table>`).join(''); try { if (navigator.clipboard.write && typeof ClipboardItem === 'function') await navigator.clipboard.write([new ClipboardItem({ 'text/plain': new Blob([clipboardCell], { type: 'text/plain' }), 'text/html': new Blob([clipboardHtml], { type: 'text/html' }) })]); else await navigator.clipboard.writeText(clipboardCell); } catch {} }
async function readClipboardRows() { try { if (navigator.clipboard.read) { const items = await navigator.clipboard.read(); for (const item of items) if (item.types.includes('text/html')) { const html = await (await item.getType('text/html')).text(); const table = new DOMParser().parseFromString(html, 'text/html').querySelector('table'); const rows = [...(table?.rows || [])].map((row) => [...row.cells].map((cell) => ({ v: cell.textContent || '', t: 's' }))); if (rows.length && rows.some((row) => row.length)) return rows; } } } catch {} let text = ''; try { text = await navigator.clipboard.readText(); } catch {} if (!text) text = clipboardCell; return text ? text.split(/\r?\n/).map((row) => row.split('\t').map((value) => ({ v: value, t: 's' }))) : undefined; }
async function cutSelection() { await copySelection(); const ws = activeWorksheet(); const before = historySnapshot(); const deleted = new Set(); let changed = false; for (const selection of selectionRanges) { const range = selectionRangeBounds(selection); for (let row = range.s.r; row <= range.e.r; row += 1) for (let col = range.s.c; col <= range.e.c; col += 1) { const ref = encode_cell({ r: row, c: col }); if (!deleted.has(ref) && ws?.[ref]) { deleted.add(ref); delete ws[ref]; changed = true; } } } if (changed) { recordHistory(before, 'cut'); dirty = true; render(); } }
function updateSelectionDom() { const ranges = selectionRanges.map(selectionRangeBounds); document.querySelectorAll('.cell-input').forEach((input) => { const pos = cellPosition(input.dataset.ref); input.parentElement.classList.toggle('selected', ranges.some((range) => pos.r >= range.s.r && pos.r <= range.e.r && pos.c >= range.s.c && pos.c <= range.e.c)); }); }
function setCellTextOnWorksheet(ws, ref, text) {
  if (!ws) return;
  const previous = ws[ref] || {};
  const value = String(text ?? '');
  if (!value) delete ws[ref];
  else if (value.startsWith('=')) ws[ref] = { ...previous, f: value.slice(1), t: 'n' };
  else if (value.trim() !== '' && Number.isFinite(Number(value))) ws[ref] = { ...previous, v: Number(value), t: 'n', f: undefined };
  else ws[ref] = { ...previous, v: value, t: 's', f: undefined };
  const bounds = worksheetBounds(ws); ws['!ref'] = `A1:${encode_cell(bounds.e)}`;
}
function setCellText(ref, text) { setCellTextOnWorksheet(activeWorksheet(), ref, text); }
function hasFormula(ws) { return Object.keys(ws || {}).some((key) => /^[A-Z]+[0-9]+$/.test(key) && ws[key]?.f); }
function addSheet() { const before = historySnapshot(); let index = 1; let name = `Sheet${index}`; while (workbook.SheetNames.includes(name)) name = `Sheet${++index}`; book_append_sheet(workbook, aoa_to_sheet([[]]), name); switchActiveSheet(name); freezeTopRow = false; freezeFirstColumn = false; selectedRef = 'A1'; setSingleSelection('A1'); recordHistory(before, 'sheet-add'); dirty = true; }
function deleteActiveSheet() { if (workbook.SheetNames.length <= 1) return 'last'; const formulas = workbook.SheetNames.flatMap((name) => Object.values(workbook.Sheets[name] || {}).filter((cell) => cell?.f).map((cell) => cell.f)); const escaped = activeSheet.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); if (formulas.some((formula) => new RegExp(`(?:'${escaped}'|${escaped})!`, 'i').test(formula)) || pivotSummariesReferencingSheet(activeSheet)) return 'referenced'; const before = historySnapshot(); delete workbook.Sheets[activeSheet]; workbook.SheetNames = workbook.SheetNames.filter((name) => name !== activeSheet); switchActiveSheet(workbook.SheetNames[0]); selectedRef = 'A1'; setSingleSelection('A1'); recordHistory(before, 'sheet-delete'); dirty = true; return true; }
function renameActiveSheet(nextName) {
  const name = String(nextName ?? '').trim();
  if (!name || name.length > 31 || /[:\\/?*\[\]]/.test(name) || workbook.SheetNames.some((sheet) => sheet.toLocaleLowerCase() === name.toLocaleLowerCase() && sheet !== activeSheet)) return false;
  if (name === activeSheet) return true;
  const before = historySnapshot(); const oldName = activeSheet;
  const quote = (value) => `'${value.replace(/'/g, "''")}'`;
  for (const sheet of workbook.SheetNames) for (const cell of Object.values(workbook.Sheets[sheet] || {})) if (cell?.f) {
    cell.f = cell.f.replace(new RegExp(`(?:${quote(oldName)}|${oldName.replace(/[.*+?^${}()|[\\]\\]/g, '\\$&')})!`, 'gi'), `${quote(name)}!`);
  }
  for (const sheet of workbook.SheetNames) { const pivot = workbook.Sheets[sheet]?.['!pivot']; if (pivot?.sourceSheet === oldName) pivot.sourceSheet = name; }
  workbook.Sheets[name] = workbook.Sheets[oldName]; delete workbook.Sheets[oldName]; workbook.SheetNames = workbook.SheetNames.map((sheet) => sheet === oldName ? name : sheet); activeSheet = name; recordHistory(before, 'sheet-rename'); dirty = true; return true;
}
function mergeSelection() { const ws = activeWorksheet(); const range = selectionBounds(); if (range.s.r === range.e.r && range.s.c === range.e.c) return false; if (Object.keys(ws || {}).some((key) => /^[A-Z]+[0-9]+$/.test(key) && ws[key]?.f && cellPosition(key).r >= range.s.r && cellPosition(key).r <= range.e.r && cellPosition(key).c >= range.s.c && cellPosition(key).c <= range.e.c)) return false; if ((ws['!merges'] || []).some((merge) => !(merge.e.r < range.s.r || merge.s.r > range.e.r || merge.e.c < range.s.c || merge.s.c > range.e.c))) return false; const before = historySnapshot(); ws['!merges'] = [...(ws['!merges'] || []), range]; for (let row = range.s.r; row <= range.e.r; row += 1) for (let col = range.s.c; col <= range.e.c; col += 1) if (row !== range.s.r || col !== range.s.c) delete ws[encode_cell({ r: row, c: col })]; recordHistory(before, 'merge'); dirty = true; return true; }
function unmergeSelection() { const ws = activeWorksheet(); const range = selectionBounds(); const merges = ws?.['!merges'] || []; const target = merges.find((merge) => (merge.s.r === range.s.r && merge.s.c === range.s.c && merge.e.r === range.e.r && merge.e.c === range.e.c) || (range.s.r === range.e.r && range.s.c === range.e.c && mergeFor(ws, range.s.r, range.s.c))); if (!target) return false; const before = historySnapshot(); ws['!merges'] = merges.filter((merge) => merge !== target); recordHistory(before, 'unmerge'); dirty = true; return true; }
function sortActiveSheet(direction) { const ws = activeWorksheet(); if (!ws) return false; const bounds = worksheetBounds(ws); const formulaRows = Object.keys(ws).filter((key) => /^[A-Z]+[0-9]+$/.test(key) && ws[key]?.f).map((key) => cellPosition(key).r); const dataEnd = formulaRows.length ? Math.min(...formulaRows) - 1 : bounds.e.r; if (dataEnd < 1 || formulaRows.some((row) => row > 0 && row <= dataEnd)) return false; const before = historySnapshot(); const column = cellPosition(selectedRef).c; const rows = []; for (let row = 1; row <= dataEnd; row += 1) rows.push({ row, cells: Array.from({ length: bounds.e.c + 1 }, (_, col) => ws[encode_cell({ r: row, c: col })]) }); rows.sort((left, right) => { const a = cellText(left.cells[column]); const b = cellText(right.cells[column]); const numeric = Number(a) - Number(b); const comparison = a !== '' && b !== '' && Number.isFinite(numeric) ? numeric : String(a).localeCompare(String(b)); return direction === 'asc' ? comparison : -comparison; }); rows.forEach((entry, offset) => entry.cells.forEach((cell, col) => { const ref = encode_cell({ r: offset + 1, c: col }); if (cell) ws[ref] = cell; else delete ws[ref]; })); recordHistory(before, `sort-${direction}`); dirty = true; return true; }
function adjustTableReference(ref, axis, point, delta, isDelete) { let range; try { range = decode_range(ref); } catch { return undefined; } const index = axis === 'r' ? point.r : point.c; const start = axis === 'r' ? range.s.r : range.s.c; const end = axis === 'r' ? range.e.r : range.e.c; if (delta > 0) { if (index <= start) { range.s[axis] += delta; range.e[axis] += delta; } else if (index <= end) range.e[axis] += delta; } else if (index < start) { range.s[axis] = Math.max(0, range.s[axis] + delta); range.e[axis] = Math.max(0, range.e[axis] + delta); } else if (index <= end) { if (start === end || (isDelete && index === start)) return undefined; range.e[axis] = Math.max(range.s[axis], range.e[axis] + delta); } return `${encode_cell(range.s)}:${encode_cell(range.e)}`; }
function adjustTableRanges(ws, axis, delta, point) { const tables = ws?.['!tables'] || []; const planned = []; for (const table of tables) { const nextRef = adjustTableReference(table.ref, axis, point, delta, delta < 0); if (!nextRef) return false; let nextFilter = table.autoFilterRef; if (nextFilter) { nextFilter = adjustTableReference(nextFilter, axis, point, delta, false); if (!nextFilter) return false; } let nextColumns = table.autoFilterColumns; if (axis === 'c' && Array.isArray(nextColumns) && nextFilter) { const filterRange = decode_range(nextFilter); nextColumns = nextColumns.map((entry) => { const colId = Number(entry?.colId); const absolute = filterRange.s.c + colId; return { ...entry, colId: delta > 0 && point.c <= absolute ? colId + delta : delta < 0 && point.c < absolute ? Math.max(0, colId + delta) : colId }; }); } planned.push({ table, nextRef, nextFilter, nextColumns }); } planned.forEach(({ table, nextRef, nextFilter, nextColumns }) => { table.ref = nextRef; if (nextFilter) table.autoFilterRef = nextFilter; if (nextColumns) table.autoFilterColumns = nextColumns; }); return true; }
function canAdjustChartRanges(ws, axis, delta, point) { return (ws?.['!charts'] || []).every((chart) => typeof chart?.ref !== 'string' || Boolean(adjustTableReference(chart.ref, axis, point, delta, delta < 0))); }
function adjustChartRanges(ws, axis, delta, point) { const charts = ws?.['!charts'] || []; const planned = []; for (const chart of charts) { if (typeof chart?.ref !== 'string') continue; const nextRef = adjustTableReference(chart.ref, axis, point, delta, delta < 0); if (!nextRef) return false; planned.push({ chart, nextRef }); } planned.forEach(({ chart, nextRef }) => { chart.ref = nextRef; }); return true; }
function pivotSummariesReferencingSheet(sheetName) { return workbook.SheetNames.some((name) => workbook.Sheets[name]?.['!pivot']?.sourceSheet === sheetName); }
function canAdjustPivotRanges(sheetName, axis, delta, point) { return workbook.SheetNames.every((name) => { const pivot = workbook.Sheets[name]?.['!pivot']; if (!pivot || pivot.sourceSheet !== sheetName) return true; return Boolean(adjustPivotSourceRef(pivot.sourceRef, axis, point, delta)); }); }
function adjustPivotRanges(sheetName, axis, delta, point) { for (const name of workbook.SheetNames) { const pivot = workbook.Sheets[name]?.['!pivot']; if (!pivot || pivot.sourceSheet !== sheetName) continue; const nextRef = adjustPivotSourceRef(pivot.sourceRef, axis, point, delta); if (!nextRef) return false; pivot.sourceRef = nextRef; } return true; }
function definedNameTarget(entry) {
  const raw = String(entry?.Ref ?? ''); const bang = raw.lastIndexOf('!'); if (bang < 0) return undefined;
  const qualifier = raw.slice(0, bang).replace(/^'|'$/g, '').replace(/''/g, "'");
  if (qualifier !== activeSheet) return undefined;
  const ref = raw.slice(bang + 1); try { return { ref, range: decode_range(ref.includes(':') ? ref : `${ref}:${ref}`) }; } catch { return undefined; }
}
function canAdjustDefinedNames(axis, delta, point) {
  return definedNames().every((entry) => { const target = definedNameTarget(entry); return !target || Boolean(adjustTableReference(target.ref, axis, point, delta, delta < 0)); });
}
function adjustDefinedNames(axis, delta, point) {
  for (const entry of definedNames()) {
    const target = definedNameTarget(entry); if (!target) continue;
    const nextRef = adjustTableReference(target.ref, axis, point, delta, delta < 0); if (!nextRef) return false;
    const bang = String(entry.Ref).lastIndexOf('!'); entry.Ref = `${String(entry.Ref).slice(0, bang)}!${nextRef}`;
  }
  return true;
}
function structureEdit(axis, delta) {
  const ws = activeWorksheet(); if (!ws) return false;
  const point = decode_range(`${selectedRef}:${selectedRef}`).s;
  const before = historySnapshot();
  if (!canAdjustChartRanges(ws, axis, delta, point)) return false;
  if (!canAdjustPivotRanges(activeSheet, axis, delta, point)) return false;
  if (!adjustTableRanges(ws, axis, delta, point)) return false;
  if (!adjustChartRanges(ws, axis, delta, point)) return false;
  if (!canAdjustDefinedNames(axis, delta, point)) return false;
  const cells = Object.entries(ws).filter(([key]) => /^[A-Z]+[0-9]+$/.test(key));
  for (const [key, cell] of cells) {
    delete ws[key]; const pos = decode_range(`${key}:${key}`).s; const index = axis === 'r' ? pos.r : pos.c;
    if (delta < 0 && index === (axis === 'r' ? point.r : point.c)) continue;
    const next = { ...pos }; if (index >= (axis === 'r' ? point.r : point.c)) next[axis === 'r' ? 'r' : 'c'] += delta;
    if (next.r >= 0 && next.c >= 0) ws[encode_cell(next)] = cell;
  }
  shiftWorkbookFormulaReferences(workbook, activeSheet, axis, axis === 'r' ? point.r : point.c, delta);
  if (!adjustDefinedNames(axis, delta, point)) return false;
  if (!adjustPivotRanges(activeSheet, axis, delta, point)) return false;
  const bounds = worksheetBounds(ws); ws['!ref'] = `A1:${encode_cell(bounds.e)}`; recordHistory(before, delta > 0 ? `insert-${axis}` : `delete-${axis}`); return true;
}
function setHidden(axis, hidden) {
  const ws = activeWorksheet(); if (!ws) return false;
  const before = historySnapshot(); const point = cellPosition(selectedRef); const key = axis === 'r' ? point.r : point.c;
  const metadata = axis === 'r' ? (ws['!rows'] || (ws['!rows'] = [])) : (ws['!cols'] || (ws['!cols'] = []));
  if (hidden === Boolean(metadata[key]?.hidden)) return false;
  if (hidden) metadata[key] = { ...(metadata[key] || {}), hidden: true }; else delete metadata[key].hidden;
  if (!hidden && metadata[key] && !Object.keys(metadata[key]).length) delete metadata[key];
  recordHistory(before, hidden ? `hide-${axis}` : `unhide-${axis}`); dirty = true; return true;
}
function resizeAxis(axis, size) {
  const ws = activeWorksheet(); if (!ws) return false;
  const value = Number(size); if (!Number.isFinite(value) || value < 1 || value > 255) return false;
  const before = historySnapshot(); const point = cellPosition(selectedRef); const key = axis === 'r' ? point.r : point.c;
  const metadata = axis === 'r' ? (ws['!rows'] || (ws['!rows'] = [])) : (ws['!cols'] || (ws['!cols'] = []));
  const current = axis === 'r' ? metadata[key]?.hpt : metadata[key]?.wch;
  if (current === value) return false;
  metadata[key] = { ...(metadata[key] || {}), ...(axis === 'r' ? { hpt: value } : { wch: value }) };
  recordHistory(before, `resize-${axis}`); dirty = true; return true;
}
function translateFormula(formula, rowDelta, colDelta) {
  return String(formula).replace(/(^|[^A-Za-z0-9_])([$]?)([A-Z]{1,3})([$]?)(\d+)/gi, (match, prefix, colAbsolute, colText, rowAbsolute, rowText) => {
    const column = colText.toUpperCase().split('').reduce((total, char) => total * 26 + char.charCodeAt(0) - 64, 0) - 1;
    const row = Number(rowText) - 1;
    const nextColumn = colAbsolute ? column : column + colDelta;
    const nextRow = rowAbsolute ? row : row + rowDelta;
    if (nextColumn < 0 || nextRow < 0) return match;
    let label = ''; let value = nextColumn + 1; while (value) { const remainder = (value - 1) % 26; label = String.fromCharCode(65 + remainder) + label; value = Math.floor((value - 1) / 26); }
    return `${prefix}${colAbsolute ? '$' : ''}${label}${rowAbsolute ? '$' : ''}${nextRow + 1}`;
  });
}
function fillSelection(targetRef) {
  const ws = activeWorksheet(); if (!ws || !fillSource) return false;
  const target = decode_range(`${fillSource.sRef}:${targetRef}`); const before = historySnapshot(); let changed = false;
  for (let row = target.s.r; row <= target.e.r; row += 1) for (let col = target.s.c; col <= target.e.c; col += 1) {
    if (row >= fillSource.range.s.r && row <= fillSource.range.e.r && col >= fillSource.range.s.c && col <= fillSource.range.e.c) continue;
    const sourceRow = fillSource.range.s.r + ((row - fillSource.range.s.r) % (fillSource.range.e.r - fillSource.range.s.r + 1));
    const sourceCol = fillSource.range.s.c + ((col - fillSource.range.s.c) % (fillSource.range.e.c - fillSource.range.s.c + 1));
    const source = ws[encode_cell({ r: sourceRow, c: sourceCol })]; const ref = encode_cell({ r: row, c: col });
    if (source) { const next = cloneCell(source); if (source.f) { next.f = translateFormula(source.f, row - sourceRow, col - sourceCol); delete next.v; } ws[ref] = next; } else delete ws[ref]; changed = true;
  }
  fillSource = undefined; fillDragging = false; if (changed) { recordHistory(before, 'fill'); dirty = true; } return changed;
}
function unhideAllAxes() {
  const ws = activeWorksheet(); if (!ws) return false;
  const rows = ws['!rows']; const cols = ws['!cols'];
  const hasHidden = [rows, cols].some((metadata) => metadata?.some((entry) => entry?.hidden));
  if (!hasHidden) return false;
  const before = historySnapshot();
  for (const metadata of [rows, cols]) if (metadata) for (const entry of metadata) if (entry?.hidden) delete entry.hidden;
  recordHistory(before, 'unhide-all'); dirty = true; return true;
}
function recalculate() {
  refreshPivotSummaries();
  const bytes = buildBytes();
  const diagnostic = JSON.parse(diagnoseWorkbook(bytes));
  const calculated = JSON.parse(calculateWorkbook(bytes));
  latestBytes = bytes;
  const ref = workbookName ? (firstFormulaRef(activeWorksheet()) || selectedRef) : (activeSheet === 'Summary' ? resultRef() : selectedRef);
  result = { value: cellValue(calculated, ref), diagnostic };
  for (const name of calculated?.SheetNames || []) for (const [cellRef, next] of Object.entries(calculated.Sheets?.[name] || {})) {
    if (!/^[A-Z]+\d+$/.test(cellRef) || !next) continue;
    workbook.Sheets[name][cellRef] = { ...(workbook.Sheets[name][cellRef] || {}), ...next };
  }
  // The first byte sequence is the calculation input. Rebuild after projecting
  // calculated values so an immediate download cannot return stale formula
  // results while the in-memory grid already shows the new values.
  latestBytes = buildBytes();
  dirty = false;
}
function downloadWorkbook() {
  const bytes = dirty ? buildBytes() : (latestBytes || buildBytes());
  try { inspectZipBudget(bytes); const verified = read(bytes); if (!verified.SheetNames.length) throw new Error('no sheets'); latestBytes = bytes; exportStatus = `${copy[language].exported}: ${bytes.length} bytes, ${verified.SheetNames.length} sheet(s)`; }
  catch { exportStatus = copy[language].error; const status = document.querySelector('#download-status'); if (status) status.textContent = exportStatus; return; }
  const isMacroEnabled = Boolean(workbook['!vbaProject']);
  const blob = new Blob([bytes], { type: isMacroEnabled ? 'application/vnd.ms-excel.sheet.macroEnabled.12' : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const fallbackName = `elixcee-${recipeKey}.${isMacroEnabled ? 'xlsm' : 'xlsx'}`;
  const outputName = workbookName ? workbookName.replace(/\.(xlsx|xlsm)$/i, isMacroEnabled ? '.xlsm' : '.xlsx') : fallbackName;
  const link = document.createElement('a'); link.href = URL.createObjectURL(blob); link.download = outputName; document.body.append(link); link.click(); link.remove(); URL.revokeObjectURL(link.href);
  const status = document.querySelector('#download-status'); if (status) status.textContent = exportStatus;
}
function cellStyleCss(ws, ref, col) {
  const style = ws?.[ref]?.s || {};
  const font = style.font || {};
  const fill = style.fill?.fgColor?.rgb;
  const alignment = style.alignment || {};
  const borderCss = ['top', 'right', 'bottom', 'left'].map((side) => {
    const edge = style.border?.[side];
    return edge?.style
      ? `border-${side}:2px solid ${edge.color?.rgb ? `#${String(edge.color.rgb).slice(-6)}` : '#3856d9'};`
      : '';
  }).join('');
  const width = Number(ws?.['!cols']?.[col]?.wch);
  return `${Number.isFinite(width) ? `width:${Math.max(1, Math.min(255, width))}ch;` : ''}${font.bold ? 'font-weight:700;' : ''}${font.italic ? 'font-style:italic;' : ''}${font.underline ? 'text-decoration:underline;' : ''}${font.color?.rgb ? `color:#${String(font.color.rgb).slice(-6)};` : ''}${fill ? `background:#${String(fill).slice(-6)};` : ''}${borderCss}${alignment.horizontal ? `text-align:${alignment.horizontal};` : ''}${alignment.wrapText ? 'white-space:normal;' : ''}${conditionalFormatStyle(ws, ref)}`;
}

function cellControlHtml(ws, ref) {
  const choices = validationChoices(validationFor(ws, ref));
  const tabIndex = ref === selectedRef ? '0' : '-1';
  if (choices.length) {
    const current = String(cellText(ws?.[ref]));
    return `<select class="cell-input" data-ref="${ref}" tabindex="${tabIndex}" aria-label="${ref}">${choices.map((choice) => `<option value="${escapeHTML(choice)}"${current === choice ? ' selected' : ''}>${escapeHTML(choice)}</option>`).join('')}</select>`;
  }
  return `<input class="cell-input" data-ref="${ref}" tabindex="${tabIndex}" value="${escapeHTML(cellText(ws?.[ref]))}" aria-label="${ref}" />`;
}

function renderSheetCell(ws, row, col, selected) {
  const merge = mergeFor(ws, row, col);
  if (merge && (row !== merge.s.r || col !== merge.s.c)) return '';
  const position = merge ? merge.s : { r: row, c: col };
  const ref = encode_cell(position);
  const inSelection = selected.some((range) => row >= range.s.r && row <= range.e.r && col >= range.s.c && col <= range.e.c);
  const spans = merge ? ` rowspan="${merge.e.r - merge.s.r + 1}" colspan="${merge.e.c - merge.s.c + 1}"` : '';
  const note = commentTextFor(ws?.[ref]);
  const noteAttr = note ? ` title="${escapeHTML(note)}"` : '';
  const activeRange = selected[selected.length - 1] || { e: { r: -1, c: -1 } };
  const handle = inSelection && row === activeRange.e.r && col === activeRange.e.c
    ? '<span class="fill-handle" aria-hidden="true"></span>'
    : '';
  return `<td${spans} role="gridcell" headers="row-header-${row} col-header-${col}" aria-rowindex="${row + 1}" aria-colindex="${col + 1}" aria-selected="${inSelection}" aria-label="${ref}" class="${inSelection ? 'selected' : ''}${col === 0 && freezeFirstColumn ? ' frozen-first' : ''}${note ? ' has-comment' : ''}" style="${cellStyleCss(ws, ref, col)}"${noteAttr}>${cellControlHtml(ws, ref)}${handle}</td>`;
}
function visibleSheetWindow(ws) {
  const bounds = worksheetBounds(ws);
  const point = cellPosition(selectedRef);
  const rowLimit = 20;
  const colLimit = 12;
  const rowStart = Math.max(0, Math.min(Math.max(0, bounds.e.r - rowLimit + 1), point.r - Math.floor(rowLimit / 2)));
  const colStart = Math.max(0, Math.min(Math.max(0, bounds.e.c - colLimit + 1), point.c - Math.floor(colLimit / 2)));
  const rowEnd = Math.min(bounds.e.r, rowStart + rowLimit - 1);
  const colEnd = Math.min(bounds.e.c, colStart + colLimit - 1);
  const rows = Array.from({ length: rowEnd - rowStart + 1 }, (_, offset) => rowStart + offset)
    .filter((row) => !ws?.['!rows']?.[row]?.hidden);
  const cols = Array.from({ length: colEnd - colStart + 1 }, (_, offset) => colStart + offset)
    .filter((col) => !ws?.['!cols']?.[col]?.hidden);
  return { bounds, rows, cols };
}

function renderSheetRows(ws, rows, cols, selected) {
  return rows.map((row) => {
    if (filterColumn >= 0 && row > 0 && String(cellText(ws?.[encode_cell({ r: row, c: filterColumn })])) !== filterValue) return '';
    const height = Number(ws?.['!rows']?.[row]?.hpt);
    const rowStyle = Number.isFinite(height) ? ` style="height:${Math.max(1, Math.min(409, height))}pt"` : '';
    const cells = cols.map((col) => renderSheetCell(ws, row, col, selected)).join('');
    return `<tr role="row" aria-rowindex="${row + 1}"${rowStyle}><th id="row-header-${row}" class="row-label" data-row="${row}" scope="row" aria-label="Row ${row + 1}">${row + 1}</th>${cells}</tr>`;
  }).join('');
}

function renderSheetColumns(ws, cols) {
  return cols.map((col) => {
    const width = Number(ws?.['!cols']?.[col]?.wch);
    const widthStyle = Number.isFinite(width) ? ` style="width:${Math.max(1, Math.min(255, width))}ch"` : '';
    return `<th id="col-header-${col}" class="column-label${col === 0 && freezeFirstColumn ? ' frozen-first' : ''}" data-col="${col}" scope="col" aria-colindex="${col + 1}" aria-label="Column ${columnLabel(col)}"${widthStyle}>${columnLabel(col)}</th>`;
  }).join('');
}

function renderSheet(ws) {
  const { bounds, rows: visibleRows, cols: visibleCols } = visibleSheetWindow(ws);
  const selected = selectionRanges.map(selectionRangeBounds);
  const rows = renderSheetRows(ws, visibleRows, visibleCols, selected);
  const columns = renderSheetColumns(ws, visibleCols);
  return `<div class="sheet-wrap" role="region" tabindex="0" aria-label="Visible worksheet region. Use the scrollbars to view more rows and columns."><table class="sheet${freezeTopRow ? ' freeze-top' : ''}${freezeFirstColumn ? ' freeze-first-column' : ''}" role="grid" aria-multiselectable="true" aria-rowcount="${bounds.e.r + 1}" aria-colcount="${bounds.e.c + 1}"><thead><tr role="row"><th class="corner" aria-hidden="true"></th>${columns}</tr></thead><tbody>${rows}</tbody></table></div>${chartPreview(ws)}`;
}

function revealSelectedCell() {
  const sheetWrap = document.querySelector('.sheet-wrap');
  const cellInput = [...document.querySelectorAll('.cell-input')].find((input) => input.dataset.ref === selectedRef);
  const cell = cellInput?.closest('td');
  if (!sheetWrap || !cell) return;
  const viewport = sheetWrap.getBoundingClientRect();
  const target = cell.getBoundingClientRect();
  const headerHeight = 30;
  const rowHeaderWidth = 34;
  if (target.top < viewport.top + headerHeight) sheetWrap.scrollTop -= viewport.top + headerHeight - target.top;
  else if (target.bottom > viewport.bottom) sheetWrap.scrollTop += target.bottom - viewport.bottom;
  if (target.left < viewport.left + rowHeaderWidth) sheetWrap.scrollLeft -= viewport.left + rowHeaderWidth - target.left;
  else if (target.right > viewport.right) sheetWrap.scrollLeft += target.right - viewport.right;
}

function renderAppHtml(t, recipe, cutLabel) {
  return `<main class="page">
    <header class="header"><div class="brand"><span class="mark">✦</span><span>${t.title}</span></div>
      <label>${t.language}<select id="language"><option value="en" ${language === 'en' ? 'selected' : ''}>English</option><option value="ja" ${language === 'ja' ? 'selected' : ''}>日本語</option><option value="zh" ${language === 'zh' ? 'selected' : ''}>简体中文</option></select></label>
    </header>
    <p class="intro">${t.intro}</p>
    <section class="toolbar"><label>${t.recipe}<select id="recipe"><option value="sales" ${recipeKey === 'sales' ? 'selected' : ''}>${t.sales}</option><option value="budget" ${recipeKey === 'budget' ? 'selected' : ''}>${t.budget}</option><option value="grades" ${recipeKey === 'grades' ? 'selected' : ''}>${t.grades}</option><option value="multi" ${recipeKey === 'multi' ? 'selected' : ''}>${t.multi}</option></select></label><label class="upload"><span>${t.upload}</span><input id="upload" type="file" accept=".xlsx,.xlsm,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel.sheet.macroEnabled.12" /><small>${t.uploadHint}</small></label><span class="badge">${workbookName ? `${t.imported}: ${escapeHTML(workbookName)}` : `${recipe.headers.length} columns · ${recipe.rows.length} rows${recipe.extra ? ` · ${workbook.SheetNames.length} sheets` : ''}`}</span></section>
    <div class="tabs" role="tablist"><button class="tab ${activeTab === 'sheet' ? 'active' : ''}" id="tab-sheet" role="tab" aria-selected="${activeTab === 'sheet'}">▦ ${t.sheetTab}</button><button class="tab ${activeTab === 'download' ? 'active' : ''}" id="tab-download" role="tab" aria-selected="${activeTab === 'download'}">⇩ ${t.downloadTab}</button></div>
    <section id="panel-sheet" class="tab-panel ${activeTab === 'sheet' ? 'active' : ''}" role="tabpanel"><div class="grid"><div class="card editor"><div class="sheet-title"><h2>${t.input}</h2><span>${escapeHTML(activeSheet)}</span></div><div class="sheet-tabs" role="tablist">${workbook.SheetNames.map((name) => `<button class="sheet-tab ${name === activeSheet ? 'active' : ''}" data-sheet="${escapeHTML(name)}" role="tab" aria-selected="${name === activeSheet}">${escapeHTML(name)}</button>`).join('')}</div><div class="formula-bar"><label for="formula-input">${t.formulaBar}</label><input id="formula-input" value="${escapeHTML(cellText(activeWorksheet()?.[selectedRef]))}" aria-label="${t.formulaBar}" /><button class="secondary" id="copy">${t.copy}</button><button class="secondary" id="cut">${cutLabel}</button><button class="secondary" id="paste">${t.paste}</button></div>${renderSheet(activeWorksheet())}<div class="actions"><button id="calculate">${t.calculate}</button><button class="secondary" id="undo">${t.undo}</button><button class="secondary" id="redo">${t.redo}</button><button class="secondary" id="reset">${t.reset}</button><button class="secondary" id="add-row">${t.addRow}</button><button class="secondary" id="delete-row">${t.deleteRow}</button><button class="secondary" id="add-column">${t.addColumn}</button><button class="secondary" id="delete-column">${t.deleteColumn}</button></div><div class="vba-box"><h2>${t.vbaTitle}</h2><p class="hint">${t.vbaHint}</p><label>${t.vbaName}<input id="vba-name" value="Demo" /></label><label>${t.vbaSource}<textarea id="vba-source" rows="5">${escapeHTML(t.vbaExample)}</textarea></label><button id="run-vba">${t.vbaRun}</button><pre id="vba-result" class="vba-result" aria-live="polite">${escapeHTML(vbaStatus)}</pre></div></div><div class="card output"><h2>${t.result}</h2><div class="formula"><span>${t.selectCell}: ${selectionLabel()}</span><code>${escapeHTML(cellText(activeWorksheet()?.[selectedRef])) || '—'}</code></div><div class="answer" id="answer">${result?.value ?? '—'}</div><p class="diagnostics" id="diagnostics">${escapeHTML(diagnosticText(t))}</p><details class="operation-log"><summary>${t.operationLog} (${operationLog.length})</summary><ol>${operationLog.map((entry) => `<li><strong>${escapeHTML(entry.operation || 'edit')}</strong> <time datetime="${escapeHTML(entry.at || '')}">${escapeHTML(entry.at || '')}</time> <span>${escapeHTML(entry.sheet)}!${escapeHTML(entry.selection)}</span></li>`).join('')}</ol></details><p class="status ${dirty ? 'warning' : ''}" id="status">${t.status}: ${dirty ? t.stale : t.ready}</p></div></div></section><div id="context-menu" class="context-menu" hidden><button data-context-action="add">${t.addRow}</button><button data-context-action="delete">${t.deleteRow}</button></div>
    <section id="panel-download" class="tab-panel ${activeTab === 'download' ? 'active' : ''}" role="tabpanel"><div class="card download-card"><div class="download-icon">XLSX</div><h2>${t.downloadTitle}</h2><p>${t.downloadText}</p><div class="file-meta"><span>elixcee-${recipeKey}.xlsx</span><span>${recipe.headers.length} columns · ${recipe.rows.length + 1} rows</span></div><button class="download" id="download">${t.download}</button><p class="export-status" id="download-status" aria-live="polite">${escapeHTML(exportStatus)}</p></div></section>
    <p class="boundary">${t.boundary}</p>
    <section class="card guide"><h2>${t.guideTitle}</h2><p>${t.guideText}</p><a href="${t.guideHref}">${t.guideLink} →</a></section><footer>Rust/WASM · <a href="https://github.com/kent-tokyo/elixcee">source on GitHub</a></footer>
  </main>`;
}

function bindGridKeyboardHandlers() {
  const panel = document.querySelector('#panel-sheet');
  panel.addEventListener('keydown', (event) => {
    const command = event.metaKey || event.ctrlKey;
    if (command && event.key.toLowerCase() === 'z') {
      event.preventDefault();
      if (event.shiftKey) redo(); else undo();
      return;
    }
    if (command && event.key.toLowerCase() === 'y') {
      event.preventDefault();
      redo();
      return;
    }
    if (command && event.key.toLowerCase() === 'c') {
      event.preventDefault();
      void copySelection();
      return;
    }
    if (command && event.key.toLowerCase() === 'v') {
      event.preventDefault();
      document.querySelector('#paste')?.click();
      return;
    }
    if (command && event.key.toLowerCase() === 'x') {
      event.preventDefault();
      void cutSelection();
      return;
    }
    if (event.key === 'Escape' && event.target.matches('.cell-input, #formula-input')) {
      event.preventDefault();
      const snapshot = editSnapshot;
      editSnapshot = undefined;
      if (snapshot) restoreHistorySnapshot(snapshot);
      return;
    }
    if (event.target.matches('.cell-input') && !event.altKey && !command
      && ['ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight', 'Tab', 'Enter'].includes(event.key)) {
      event.preventDefault();
      commitEdit();
      const rowDelta = event.key === 'ArrowUp' ? -1 : event.key === 'ArrowDown' || event.key === 'Enter' ? 1 : 0;
      const colDelta = event.key === 'ArrowLeft' ? -1 : event.key === 'ArrowRight' || event.key === 'Tab'
        ? (event.shiftKey ? -1 : 1)
        : 0;
      moveSelection(rowDelta, colDelta);
    }
  });

  panel.addEventListener('keydown', (event) => {
    if (!event.target.matches('.cell-input') || event.altKey) return;
    const isArrow = ['ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight'].includes(event.key);
    const isEdge = event.key === 'Home' || event.key === 'End';
    const isPage = event.key === 'PageUp' || event.key === 'PageDown';
    if (!isArrow && !isEdge && !isPage) return;
    event.preventDefault();
    event.stopImmediatePropagation();
    commitEdit();
    const point = cellPosition(selectedRef);
    const bounds = worksheetBounds(activeWorksheet());
    let rowDelta = 0;
    let colDelta = 0;
    if (isArrow) {
      rowDelta = event.key === 'ArrowUp' ? -1 : event.key === 'ArrowDown' ? 1 : 0;
      colDelta = event.key === 'ArrowLeft' ? -1 : event.key === 'ArrowRight' ? 1 : 0;
    } else if (event.key === 'Home') {
      if (event.ctrlKey || event.metaKey) {
        rowDelta = -point.r;
        colDelta = -point.c;
      } else colDelta = -point.c;
    } else if (event.key === 'End') {
      if (event.ctrlKey || event.metaKey) {
        rowDelta = bounds.e.r - point.r;
        colDelta = bounds.e.c - point.c;
      } else colDelta = bounds.e.c - point.c;
    }
    if (isPage) rowDelta = (event.key === 'PageDown' ? 1 : -1) * 20;
    moveSelection(rowDelta, colDelta, event.shiftKey);
  }, true);

  panel.addEventListener('keydown', (event) => {
    if (!event.target.matches('.cell-input') || !event.shiftKey || event.altKey
      || !['Tab', 'Enter'].includes(event.key)) return;
    event.preventDefault();
    event.stopPropagation();
    commitEdit();
    moveSelection(event.key === 'Enter' ? -1 : 0, event.key === 'Tab' ? -1 : 0, true);
  }, true);
}

function render() {
  const t = copy[language]; const recipe = recipes[recipeKey]; const cutLabel = language === 'ja' ? '切り取り' : language === 'zh' ? '剪切' : 'Cut'; const tableLabel = language === 'ja' ? 'テーブル化' : language === 'zh' ? '创建表' : 'Create table';
  Object.assign(t, sheetLabels[language], mergeLabels[language]);
  document.documentElement.lang = language; localStorage.setItem('elixcee-playground-language', language);
  syncTableFilterState(activeWorksheet());
  document.querySelector('#app').innerHTML = renderAppHtml(t, recipe, cutLabel);
  const searchBox = document.createElement('div'); searchBox.className = 'search-box'; searchBox.innerHTML = `<label>${findLabels[language].find}<input id="find-input" placeholder="${findLabels[language].findPlaceholder}" value="${escapeHTML(findQuery)}" /></label><label>${findLabels[language].replace}<input id="replace-input" placeholder="${findLabels[language].replacePlaceholder}" value="${escapeHTML(replaceQuery)}" /></label><button class="secondary" id="replace-next">${findLabels[language].replaceOne}</button><button class="secondary" id="replace-all">${findLabels[language].replaceAll}</button><span id="find-count" aria-live="polite"></span>`; document.querySelector('.toolbar').after(searchBox);
  const pasteLabel = document.createElement('label'); pasteLabel.className = 'paste-mode'; pasteLabel.innerHTML = `<span>${pasteModeLabels[language].all}</span>`; const pasteSelect = document.createElement('select'); pasteSelect.id = 'paste-mode'; pasteSelect.innerHTML = `<option value="all">${pasteModeLabels[language].all}</option><option value="values">${pasteModeLabels[language].values}</option><option value="formulas">${pasteModeLabels[language].formulas}</option><option value="formats">${pasteModeLabels[language].formats}</option>`; pasteSelect.value = pasteMode; pasteLabel.append(pasteSelect); document.querySelector('#paste').before(pasteLabel); pasteSelect.addEventListener('change', (event) => { pasteMode = event.target.value; });
  const copyButton = document.querySelector('#copy'); const cutButton = document.querySelector('#cut'); const pasteButton = document.querySelector('#paste'); const copyReplacement = copyButton.cloneNode(true); const cutReplacement = cutButton.cloneNode(true); const pasteReplacement = pasteButton.cloneNode(true); copyButton.replaceWith(copyReplacement); cutButton.replaceWith(cutReplacement); pasteButton.replaceWith(pasteReplacement);
  copyReplacement.addEventListener('click', () => copySelection());
  cutReplacement.addEventListener('click', () => cutSelection());
  pasteReplacement.addEventListener('click', async () => { let rows = clipboardData; let mode = pasteMode; if (!rows) { rows = await readClipboardRows(); mode = 'values'; } if (rows && pasteCells(mode, rows)) { const origin = cellPosition(selectionStart); const last = rows[rows.length - 1] || []; selectedRef = selectionStart; setSingleSelection(selectionStart, encode_cell({ r: origin.r + rows.length - 1, c: origin.c + Math.max(0, last.length - 1) })); render(); } });
  const updateFindCount = () => { findQuery = document.querySelector('#find-input').value; replaceQuery = document.querySelector('#replace-input').value; document.querySelector('#find-count').textContent = findQuery ? `${findCells(findQuery).length} ${findLabels[language].found}` : ''; };
  document.querySelector('#find-input').addEventListener('input', () => { findCursor = 0; updateFindCount(); }); document.querySelector('#replace-input').addEventListener('input', updateFindCount); updateFindCount();
  document.querySelector('#replace-next').addEventListener('click', () => { if (replaceNext()) render(); else updateFindCount(); }); document.querySelector('#replace-all').addEventListener('click', () => { const count = replaceAll(); if (count) render(); else updateFindCount(); });
  const nameBox = document.createElement('input'); nameBox.id = 'name-box'; nameBox.value = selectionLabel(); nameBox.setAttribute('aria-label', nameBoxLabels[language]); nameBox.title = nameBoxLabels[language]; document.querySelector('.formula-bar').prepend(nameBox); nameBox.addEventListener('change', () => { const named = definedNameFor(nameBox.value.trim()); const resolved = named && resolveDefinedName(named); if (resolved) { switchActiveSheet(resolved.sheet); selectedRef = encode_cell(resolved.range.s); setSingleSelection(selectedRef, encode_cell(resolved.range.e)); render(); return; } try { const range = decode_range(`${nameBox.value}:${nameBox.value}`); selectedRef = encode_cell(range.s); setSingleSelection(selectedRef); render(); } catch { nameBox.value = selectionLabel(); } });
  document.querySelector('#language').addEventListener('change', (event) => { language = event.target.value; render(); });
  document.querySelector('#tab-sheet').addEventListener('click', () => { activeTab = 'sheet'; localStorage.setItem('elixcee-playground-tab', activeTab); render(); });
  document.querySelector('#tab-download').addEventListener('click', () => { activeTab = 'download'; localStorage.setItem('elixcee-playground-tab', activeTab); render(); });
  document.querySelector('#undo').addEventListener('click', () => undo());
  document.querySelector('#redo').addEventListener('click', () => redo());
  bindGridKeyboardHandlers();
  document.querySelector('#recipe').addEventListener('change', (event) => { recipeKey = event.target.value; resetValues(); workbook = makeWorkbook(); switchActiveSheet('Summary'); selectedRef = 'A1'; setSingleSelection('A1'); freezeTopRow = false; freezeFirstColumn = false; workbookName = ''; originalBytes = undefined; recalculate(); render(); });
  document.querySelector('#upload').addEventListener('change', async (event) => { const file = event.target.files[0]; if (!file) return; if (file.size > MAX_UPLOAD_BYTES) { document.querySelector('#status').textContent = `${t.status}: ${t.tooLarge}`; return; } try { const candidateBytes = new Uint8Array(await file.arrayBuffer()); const zipInfo = inspectZipBudget(candidateBytes); if (zipInfo.hasExternalLinks) { const error = new Error('external workbook links are not supported in the browser editor'); error.code = 'EXTERNAL_LINKS_UNSUPPORTED'; throw error; } if (zipInfo.unsupportedParts.length) { const error = new Error(`unsupported OOXML parts: ${zipInfo.unsupportedParts.join(', ')}`); error.code = 'UNSUPPORTED_OOXML_PARTS'; throw error; } const candidateWorkbook = read(candidateBytes, { cellStyles: true }); if (!candidateWorkbook.SheetNames?.length) throw new Error('workbook has no worksheets'); if (/\.xlsm$/i.test(file.name)) { const vbaProject = await extractZipEntry(candidateBytes, 'xl/vbaProject.bin'); if (!vbaProject?.length) throw new Error('XLSM has no vbaProject.bin'); candidateWorkbook['!vbaProject'] = vbaProject; } originalBytes = candidateBytes; workbook = candidateWorkbook; switchActiveSheet(workbook.SheetNames[0]); selectedRef = 'A1'; setSingleSelection('A1'); workbookName = file.name; recalculate(); render(); } catch (error) { document.querySelector('#status').textContent = `${t.status}: ${error?.code === 'EXTERNAL_LINKS_UNSUPPORTED' ? t.externalLinksUnsupported : error?.code === 'UNSUPPORTED_OOXML_PARTS' ? t.unsupportedParts : error?.message?.includes('ZIP') ? t.zipRisk : t.unsupported}`; console.error(error); } });
  document.querySelectorAll('.sheet-tab').forEach((button) => button.addEventListener('click', () => { switchActiveSheet(button.dataset.sheet); selectedRef = 'A1'; setSingleSelection('A1'); render(); }));
  document.querySelector('#sheet-context-menu')?.remove(); const sheetContextMenu = document.createElement('div'); sheetContextMenu.id = 'sheet-context-menu'; sheetContextMenu.className = 'context-menu'; sheetContextMenu.hidden = true; sheetContextMenu.innerHTML = `<button data-sheet-context-action="add">${t.sheetAdd}</button><button data-sheet-context-action="rename">${t.sheetRename}</button><button data-sheet-context-action="delete">${t.sheetDelete}</button>`; document.body.append(sheetContextMenu);
  document.querySelectorAll('.sheet-tab').forEach((button) => button.addEventListener('contextmenu', (event) => { event.preventDefault(); switchActiveSheet(button.dataset.sheet); selectedRef = 'A1'; setSingleSelection('A1'); sheetContextMenu.hidden = false; sheetContextMenu.style.left = `${Math.min(event.clientX, window.innerWidth - 170)}px`; sheetContextMenu.style.top = `${Math.min(event.clientY, window.innerHeight - 100)}px`; }));
  sheetContextMenu.addEventListener('click', (event) => { const action = event.target.dataset.sheetContextAction; sheetContextMenu.hidden = true; if (action === 'add') { addSheet(); render(); return; } if (action === 'rename') { const nextName = window.prompt(t.sheetRename, activeSheet); if (nextName !== null) { if (renameActiveSheet(nextName)) render(); else { const status = document.querySelector('#status'); status.textContent = `${t.status}: ${t.sheetNameInvalid}`; status.classList.add('warning'); } } return; } const outcome = deleteActiveSheet(); if (outcome === true) render(); else { const status = document.querySelector('#status'); status.textContent = `${t.status}: ${outcome === 'last' ? t.sheetLast : t.sheetReferenced}`; status.classList.add('warning'); } });
  const sheetTabs = document.querySelector('.sheet-tabs'); const addSheetButton = document.createElement('button'); addSheetButton.className = 'sheet-action'; addSheetButton.id = 'add-sheet'; addSheetButton.title = t.sheetAdd; addSheetButton.textContent = '+'; sheetTabs.append(addSheetButton); addSheetButton.addEventListener('click', () => { addSheet(); render(); }); const deleteSheetButton = document.createElement('button'); deleteSheetButton.className = 'sheet-action'; deleteSheetButton.id = 'delete-sheet'; deleteSheetButton.title = t.sheetDelete; deleteSheetButton.textContent = '−'; sheetTabs.append(deleteSheetButton); deleteSheetButton.addEventListener('click', () => { const outcome = deleteActiveSheet(); if (outcome === true) render(); else { const status = document.querySelector('#status'); status.textContent = `${t.status}: ${outcome === 'last' ? t.sheetLast : t.sheetReferenced}`; status.classList.add('warning'); } });
  const freezeButton = document.createElement('button'); freezeButton.className = 'secondary'; freezeButton.id = 'freeze-row'; freezeButton.textContent = freezeTopRow ? t.unfreezeRow : t.freezeRow; document.querySelector('.actions').append(freezeButton); freezeButton.addEventListener('click', () => { const before = historySnapshot(); freezeTopRow = !freezeTopRow; recordHistory(before, freezeTopRow ? 'freeze-row' : 'unfreeze-row'); dirty = true; render(); });
  const freezeColumnButton = document.createElement('button'); freezeColumnButton.className = 'secondary'; freezeColumnButton.id = 'freeze-column'; freezeColumnButton.textContent = freezeFirstColumn ? t.unfreezeColumn : t.freezeColumn; document.querySelector('.actions').append(freezeColumnButton); freezeColumnButton.addEventListener('click', () => { const before = historySnapshot(); freezeFirstColumn = !freezeFirstColumn; recordHistory(before, freezeFirstColumn ? 'freeze-column' : 'unfreeze-column'); dirty = true; render(); });
  const mergeButton = document.createElement('button'); mergeButton.className = 'secondary'; mergeButton.id = 'merge-cells'; mergeButton.textContent = t.merge; document.querySelector('.actions').append(mergeButton); mergeButton.addEventListener('click', () => { if (mergeSelection()) render(); else { const status = document.querySelector('#status'); status.textContent = `${t.status}: ${t.mergeBlocked}`; status.classList.add('warning'); } });
  const unmergeButton = document.createElement('button'); unmergeButton.className = 'secondary'; unmergeButton.id = 'unmerge-cells'; unmergeButton.textContent = t.unmerge; document.querySelector('.actions').append(unmergeButton); unmergeButton.addEventListener('click', () => { if (unmergeSelection()) render(); else { const status = document.querySelector('#status'); status.textContent = `${t.status}: ${t.mergeBlocked}`; status.classList.add('warning'); } });
  const sortUp = document.createElement('button'); sortUp.className = 'secondary'; sortUp.id = 'sort-asc'; sortUp.textContent = t.sortAsc; document.querySelector('.actions').append(sortUp); sortUp.addEventListener('click', () => { if (sortActiveSheet('asc')) render(); else { const status = document.querySelector('#status'); status.textContent = `${t.status}: ${t.sortBlocked}`; status.classList.add('warning'); } });
  const sortDown = document.createElement('button'); sortDown.className = 'secondary'; sortDown.id = 'sort-desc'; sortDown.textContent = t.sortDesc; document.querySelector('.actions').append(sortDown); sortDown.addEventListener('click', () => { if (sortActiveSheet('desc')) render(); else { const status = document.querySelector('#status'); status.textContent = `${t.status}: ${t.sortBlocked}`; status.classList.add('warning'); } });
  const filterLabel = document.createElement('label'); filterLabel.className = 'filter-control'; filterLabel.textContent = t.filter; const filterSelect = document.createElement('select'); filterSelect.id = 'filter-value'; const filterCol = filterColumn >= 0 ? filterColumn : cellPosition(selectedRef).c; const filterBounds = worksheetBounds(activeWorksheet()); const filterOptions = [...new Set(Array.from({ length: Math.max(0, filterBounds.e.r) }, (_, offset) => String(cellText(activeWorksheet()?.[encode_cell({ r: offset + 1, c: filterCol })]))))].filter((value) => value !== ''); filterSelect.innerHTML = `<option value="">${t.allValues}</option>${filterOptions.map((value) => `<option value="${escapeHTML(value)}">${escapeHTML(value)}</option>`).join('')}`; filterSelect.value = filterValue; filterLabel.append(filterSelect); document.querySelector('.actions').append(filterLabel); filterSelect.addEventListener('change', (event) => { const nextValue = event.target.value; const before = historySnapshot(); const tableChanged = setTableValueFilter(activeWorksheet(), filterCol, nextValue); if (tableChanged) recordHistory(before, 'filter'); filterColumn = filterCol; filterValue = nextValue; render(); }); const clearFilter = document.createElement('button'); clearFilter.className = 'secondary'; clearFilter.id = 'clear-filter'; clearFilter.textContent = t.clearFilter; document.querySelector('.actions').append(clearFilter); clearFilter.addEventListener('click', () => { const before = historySnapshot(); const tableChanged = setTableValueFilter(activeWorksheet(), filterCol, ''); if (tableChanged) recordHistory(before, 'filter-clear'); filterColumn = -1; filterValue = ''; render(); });
  const formatLabel = document.createElement('label'); formatLabel.className = 'filter-control'; formatLabel.textContent = formatLabels[language].label; const formatSelect = document.createElement('select'); formatSelect.id = 'number-format'; formatSelect.innerHTML = `<option value="General">${formatLabels[language].general}</option><option value="0">${formatLabels[language].integer}</option><option value="0.00">${formatLabels[language].decimal}</option><option value="#,##0.00">${formatLabels[language].grouped}</option><option value="0.00%">${formatLabels[language].percent}</option><option value="m/d/yy">${formatLabels[language].date}</option>`; formatLabel.append(formatSelect); document.querySelector('.actions').append(formatLabel); formatSelect.addEventListener('change', (event) => { if (applyNumberFormat(event.target.value)) render(); });
  const styleToolbar = document.createElement('div'); styleToolbar.className = 'style-toolbar'; styleToolbar.innerHTML = `<button class="secondary" id="toggle-bold">${styleLabels[language].bold}</button><button class="secondary" id="copy-format">${styleLabels[language].copy}</button><label>${styleLabels[language].font}<input id="font-color" type="color" value="#1f2937" /></label><label>${styleLabels[language].fill}<input id="fill-color" type="color" value="#ffffff" /></label><label>${styleLabels[language].border}<input id="border-color" type="color" value="#3856d9" /></label><button class="secondary" id="conditional-format">${styleLabels[language].conditional}</button>`; document.querySelector('.actions').after(styleToolbar); document.querySelector('#toggle-bold').addEventListener('click', () => { if (applyCellStyle('bold')) render(); }); document.querySelector('#copy-format').addEventListener('click', () => { if (copyFormatToSelection()) render(); }); document.querySelector('#font-color').addEventListener('change', (event) => { if (applyCellStyle('font', event.target.value.slice(1).toUpperCase())) render(); }); document.querySelector('#fill-color').addEventListener('change', (event) => { if (applyCellStyle('fill', event.target.value.slice(1).toUpperCase())) render(); }); document.querySelector('#border-color').addEventListener('change', (event) => { if (applyCellStyle('border', event.target.value.slice(1).toUpperCase())) render(); }); document.querySelector('#conditional-format').addEventListener('click', () => { if (applyConditionalFormat()) render(); });
  const contextMenu = document.querySelector('#context-menu'); contextMenu.insertAdjacentHTML('beforeend', `<button data-context-action="hide">${t.hide}</button><button data-context-action="resize">${t.resize}</button>`);
  document.querySelectorAll('.row-label, .column-label').forEach((header) => header.addEventListener('click', (event) => { const bounds = worksheetBounds(activeWorksheet()); const isColumn = header.classList.contains('column-label'); const index = Number(header.dataset[isColumn ? 'col' : 'row']); const anchor = cellPosition(selectionStart); const first = isColumn ? Math.min(anchor.c, index) : Math.min(anchor.r, index); const last = isColumn ? Math.max(anchor.c, index) : Math.max(anchor.r, index); const start = isColumn ? encode_cell({ r: bounds.s.r, c: event.shiftKey ? first : index }) : encode_cell({ r: event.shiftKey ? first : index, c: bounds.s.c }); const end = isColumn ? encode_cell({ r: bounds.e.r, c: event.shiftKey ? last : index }) : encode_cell({ r: event.shiftKey ? last : index, c: bounds.e.c }); const next = { start, end }; if (event.ctrlKey || event.metaKey) { if (!selectionRanges.some((range) => selectionRangeLabel(range) === selectionRangeLabel(next))) selectionRanges.push(next); selectedRef = start; selectionStart = start; selectionEnd = end; } else { selectedRef = start; setSingleSelection(start, end); } render(); document.querySelector(`[data-ref="${start}"]`)?.focus(); }));
  document.querySelectorAll('.row-label, .column-label').forEach((header) => header.addEventListener('contextmenu', (event) => { event.preventDefault(); const isColumn = header.classList.contains('column-label'); contextAxis = isColumn ? 'c' : 'r'; const pos = isColumn ? { r: 0, c: Number(header.dataset.col) } : { r: Number(header.dataset.row), c: 0 }; selectedRef = encode_cell(pos); setSingleSelection(selectedRef); contextMenu.querySelector('[data-context-action="add"]').textContent = isColumn ? t.addColumn : t.addRow; contextMenu.querySelector('[data-context-action="delete"]').textContent = isColumn ? t.deleteColumn : t.deleteRow; contextMenu.querySelector('[data-context-action="hide"]').textContent = `${t.hide} ${isColumn ? columnLabel(pos.c) : pos.r + 1}`; contextMenu.hidden = false; contextMenu.style.left = `${Math.min(event.clientX, window.innerWidth - 170)}px`; contextMenu.style.top = `${Math.min(event.clientY, window.innerHeight - 100)}px`; }));
  document.querySelectorAll('.cell-input').forEach((input) => { input.addEventListener('mousedown', (event) => { if (event.button !== 0) return; dragging = true; dragMoved = false; if (!event.ctrlKey && !event.metaKey) setSingleSelection(event.target.dataset.ref); else { selectedRef = event.target.dataset.ref; } selectionStart = event.target.dataset.ref; selectionEnd = selectionStart; selectedRef = selectionStart; updateSelectionDom(); }); input.addEventListener('mouseover', (event) => { if (!dragging) return; const ref = event.target.dataset.ref; if (ref !== selectionEnd) { dragMoved = true; selectionEnd = ref; selectionRanges = [{ start: selectionStart, end: selectionEnd }]; updateSelectionDom(); } }); input.addEventListener('mouseup', () => { dragging = false; }); input.addEventListener('click', (event) => { const ref = event.target.dataset.ref; if (dragMoved) { dragMoved = false; event.preventDefault(); return; } if (event.ctrlKey || event.metaKey) { if (!selectionRanges.some((range) => range.start === ref && range.end === ref)) selectionRanges.push({ start: ref, end: ref }); selectionStart = ref; selectionEnd = ref; } else if (event.shiftKey) { selectionEnd = ref; selectionRanges = [{ start: selectionStart, end: selectionEnd }]; } else setSingleSelection(ref); selectedRef = ref; render(); document.querySelector(`[data-ref="${ref}"]`)?.focus(); }); input.addEventListener('focus', () => { beginEdit(); selectedRef = input.dataset.ref; document.querySelector('#formula-input').value = cellText(activeWorksheet()?.[selectedRef]); }); input.addEventListener('blur', commitEdit); input.addEventListener('input', (event) => { selectedRef = event.target.dataset.ref; setCellText(selectedRef, event.target.value); dirty = true; document.querySelector('#status').textContent = `${t.status}: ${t.stale}`; document.querySelector('#status').classList.add('warning'); }); });
  document.querySelectorAll('.cell-input').forEach((input) => { input.addEventListener('pointerdown', (event) => { if (event.pointerType !== 'touch') return; event.preventDefault(); dragging = true; dragMoved = false; const ref = event.currentTarget.dataset.ref; setSingleSelection(ref); selectionStart = ref; selectionEnd = ref; selectedRef = ref; updateSelectionDom(); }); input.addEventListener('pointerover', (event) => { if (event.pointerType !== 'touch' || !dragging) return; const ref = event.currentTarget.dataset.ref; if (ref !== selectionEnd) { dragMoved = true; selectionEnd = ref; selectedRef = ref; selectionRanges = [{ start: selectionStart, end: selectionEnd }]; updateSelectionDom(); } }); input.addEventListener('pointerup', (event) => { if (event.pointerType !== 'touch' || !dragging) return; dragging = false; if (dragMoved) { dragMoved = false; render(); document.querySelector(`[data-ref="${selectedRef}"]`)?.focus(); } }); input.addEventListener('pointercancel', (event) => { if (event.pointerType !== 'touch') return; dragging = false; dragMoved = false; }); });
  document.querySelectorAll('.fill-handle').forEach((handle) => { handle.addEventListener('mousedown', (event) => { event.preventDefault(); event.stopPropagation(); const range = selectionBounds(); fillSource = { range, sRef: encode_cell(range.s) }; fillDragging = true; }); });
  document.onmouseover = (event) => { if (!fillDragging || !event.target.matches('.cell-input')) return; const ref = event.target.dataset.ref; selectionEnd = ref; selectionRanges = [{ start: selectionStart, end: selectionEnd }]; selectedRef = ref; updateSelectionDom(); };
  document.onmouseup = () => { if (!fillDragging) return; if (fillSelection(selectionEnd)) render(); else { fillSource = undefined; fillDragging = false; } };
  document.querySelector('#formula-input').addEventListener('focus', beginEdit); document.querySelector('#formula-input').addEventListener('blur', commitEdit); document.querySelector('#formula-input').addEventListener('input', (event) => { setCellText(selectedRef, event.target.value); dirty = true; document.querySelector('#status').textContent = `${t.status}: ${t.stale}`; document.querySelector('#status').classList.add('warning'); });
  document.querySelector('#run-vba').addEventListener('click', () => { if (activeVbaWorker) return; const before = historySnapshot(); const vbaSheet = activeSheet; const vbaSelection = selectionLabel(); const worker = new Worker('./assets/vba-worker.js', { type: 'module' }); activeVbaWorker = worker; cancelVbaButton.hidden = false; const resultNode = document.querySelector('#vba-result'); resultNode.textContent = ''; worker.onmessage = (event) => { worker.terminate(); activeVbaWorker = undefined; cancelVbaButton.hidden = true; if (!event.data.ok) { vbaStatus = `Error: ${event.data.error}`; render(); return; } const target = workbook.Sheets[vbaSheet]; event.data.changes.forEach(({ ref, value }) => setCellTextOnWorksheet(target, ref, value)); if (event.data.changes.length) { recordHistory(before, 'vba-run', { sheet: vbaSheet, selection: vbaSelection }); dirty = true; } vbaStatus = `${t.vbaDone}: ${event.data.statements} statement(s), ${event.data.changes.length} cell(s)`; render(); }; worker.onerror = () => { worker.terminate(); activeVbaWorker = undefined; cancelVbaButton.hidden = true; vbaStatus = t.error; render(); }; worker.postMessage({ source: document.querySelector('#vba-source').value, macroName: document.querySelector('#vba-name').value, cells: workerCells(activeWorksheet()) }); });
  const cancelVbaButton = document.createElement('button'); cancelVbaButton.className = 'secondary'; cancelVbaButton.id = 'cancel-vba'; cancelVbaButton.textContent = language === 'ja' ? 'キャンセル' : language === 'zh' ? '取消' : 'Cancel'; cancelVbaButton.hidden = true; document.querySelector('#run-vba').after(cancelVbaButton); cancelVbaButton.addEventListener('click', () => { if (!activeVbaWorker) return; activeVbaWorker.terminate(); activeVbaWorker = undefined; vbaStatus = language === 'ja' ? 'VBAをキャンセルしました' : language === 'zh' ? 'VBA 已取消' : 'VBA cancelled'; cancelVbaButton.hidden = true; render(); });
  for (const [id, axis, delta] of [['add-row', 'r', 1], ['delete-row', 'r', -1], ['add-column', 'c', 1], ['delete-column', 'c', -1]]) document.querySelector(`#${id}`).addEventListener('click', () => { if (structureEdit(axis, delta)) { dirty = true; render(); } else { document.querySelector('#status').textContent = `${t.status}: ${t.structureBlocked}`; document.querySelector('#status').classList.add('warning'); } });
  document.querySelector('#context-menu').addEventListener('click', (event) => { const action = event.target.dataset.contextAction; document.querySelector('#context-menu').hidden = true; if (action === 'hide') { if (setHidden(contextAxis, true)) render(); return; } if (action === 'resize') { const point = cellPosition(selectedRef); const metadata = contextAxis === 'r' ? activeWorksheet()?.['!rows']?.[point.r] : activeWorksheet()?.['!cols']?.[point.c]; const current = contextAxis === 'r' ? metadata?.hpt : metadata?.wch; const value = window.prompt(t.resizePrompt, String(current || (contextAxis === 'r' ? 20 : 12))); if (value !== null) { if (resizeAxis(contextAxis, value)) render(); else { document.querySelector('#status').textContent = `${t.status}: ${t.invalidSize}`; document.querySelector('#status').classList.add('warning'); } } return; } const delta = action === 'add' ? 1 : -1; if (structureEdit(contextAxis, delta)) { dirty = true; render(); } else { document.querySelector('#status').textContent = `${t.status}: ${t.structureBlocked}`; document.querySelector('#status').classList.add('warning'); } });
  const unhideAllButton = document.createElement('button'); unhideAllButton.className = 'secondary'; unhideAllButton.id = 'unhide-all'; unhideAllButton.textContent = t.unhideAll; document.querySelector('.actions').append(unhideAllButton); unhideAllButton.addEventListener('click', () => { if (unhideAllAxes()) render(); });
  const tableButton = document.createElement('button'); tableButton.className = 'secondary'; tableButton.id = 'create-table'; tableButton.textContent = tableLabel; document.querySelector('.actions').append(tableButton); tableButton.addEventListener('click', () => { if (applyTable()) render(); });
  document.querySelector('#calculate').addEventListener('click', () => { try { recalculate(); render(); } catch { document.querySelector('#status').textContent = `${t.status}: ${t.error}`; } });
  document.querySelector('#reset').addEventListener('click', async () => { if (originalBytes && workbookName) { workbook = read(originalBytes, { cellStyles: true }); if (/\.xlsm$/i.test(workbookName)) workbook['!vbaProject'] = await extractZipEntry(originalBytes, 'xl/vbaProject.bin'); switchActiveSheet(workbook.SheetNames[0]); selectedRef = 'A1'; } else { resetValues(); workbook = makeWorkbook(); switchActiveSheet('Summary'); selectedRef = 'A1'; freezeTopRow = false; freezeFirstColumn = false; } recalculate(); render(); });
  document.querySelector('#panel-sheet').addEventListener('keydown', (event) => {
    if (!event.target.matches('.cell-input')) return;
    const isArrow = ['ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight'].includes(event.key);
    const isEdge = event.key === 'Home' || event.key === 'End';
    const isPage = event.key === 'PageUp' || event.key === 'PageDown';
    if (!isArrow && !isEdge && !isPage) return;
    event.preventDefault(); event.stopImmediatePropagation(); commitEdit();
    const point = cellPosition(selectedRef); const bounds = worksheetBounds(activeWorksheet()); let rowDelta = 0; let colDelta = 0;
    if (event.key === 'ArrowUp') rowDelta = event.ctrlKey || event.metaKey ? -point.r : -1;
    if (event.key === 'ArrowDown') rowDelta = event.ctrlKey || event.metaKey ? bounds.e.r - point.r : 1;
    if (event.key === 'ArrowLeft') colDelta = event.ctrlKey || event.metaKey ? -point.c : -1;
    if (event.key === 'ArrowRight') colDelta = event.ctrlKey || event.metaKey ? bounds.e.c - point.c : 1;
    if (event.key === 'Home') { if (event.ctrlKey || event.metaKey) { rowDelta = -point.r; colDelta = -point.c; } else colDelta = -point.c; }
    if (event.key === 'End') { if (event.ctrlKey || event.metaKey) { rowDelta = bounds.e.r - point.r; colDelta = bounds.e.c - point.c; } else colDelta = bounds.e.c - point.c; }
    if (isPage) rowDelta = (event.key === 'PageDown' ? 1 : -1) * 20;
    moveSelection(rowDelta, colDelta, event.shiftKey);
  }, true);
  const validationButton = document.createElement('button'); validationButton.className = 'secondary'; validationButton.id = 'add-validation'; validationButton.textContent = validationLabels[language].list; document.querySelector('.actions').append(validationButton); validationButton.addEventListener('click', () => { if (applyListValidation()) render(); else { const status = document.querySelector('#status'); status.textContent = `${t.status}: ${validationLabels[language].invalid}`; status.classList.add('warning'); } });
  const defineNameButton = document.createElement('button'); defineNameButton.className = 'secondary'; defineNameButton.id = 'define-name'; defineNameButton.textContent = definedNameLabels[language].add; document.querySelector('.actions').append(defineNameButton); defineNameButton.addEventListener('click', () => { const name = window.prompt(definedNameLabels[language].prompt, 'SalesRange'); if (name !== null) { if (applyDefinedName(name)) render(); else { const status = document.querySelector('#status'); status.textContent = `${t.status}: ${definedNameLabels[language].invalid}`; status.classList.add('warning'); } } });
  const commentButton = document.createElement('button'); commentButton.className = 'secondary'; commentButton.id = 'add-comment'; commentButton.textContent = commentLabels[language].add; document.querySelector('.actions').append(commentButton); commentButton.addEventListener('click', () => { const text = window.prompt(commentLabels[language].prompt, ''); if (text !== null) { const note = text.trim(); if (note) { const before = historySnapshot(); const ws = activeWorksheet(); const cell = ws[selectedRef] || { v: '', t: 's' }; ws[selectedRef] = cell; cell_add_comment(cell, note, 'elixcee'); recordHistory(before, 'comment-add'); dirty = true; render(); } else { const status = document.querySelector('#status'); status.textContent = `${t.status}: ${commentLabels[language].invalid}`; status.classList.add('warning'); } } });
  [['add-chart-bar', chartLabels[language].bar, 'bar'], ['add-chart-line', chartLabels[language].line, 'line']].forEach(([id, label, type]) => { const button = document.createElement('button'); button.className = 'secondary'; button.id = id; button.textContent = label; document.querySelector('.actions').append(button); button.addEventListener('click', () => { if (applyChart(type)) render(); else { const status = document.querySelector('#status'); status.textContent = `${t.status}: ${chartLabels[language].invalid}`; status.classList.add('warning'); } }); });
  const editChartButton = document.createElement('button'); editChartButton.className = 'secondary'; editChartButton.id = 'edit-chart'; editChartButton.textContent = chartLabels[language].edit; document.querySelector('.actions').append(editChartButton); editChartButton.addEventListener('click', () => { if (editChart()) render(); else { const status = document.querySelector('#status'); status.textContent = `${t.status}: ${chartLabels[language].missing}`; status.classList.add('warning'); } });
  const resizeChartButton = document.createElement('button'); resizeChartButton.className = 'secondary'; resizeChartButton.id = 'resize-chart'; resizeChartButton.textContent = chartLabels[language].resize; document.querySelector('.actions').append(resizeChartButton); resizeChartButton.addEventListener('click', () => { if (resizeChart()) render(); else { const status = document.querySelector('#status'); status.textContent = `${t.status}: ${chartLabels[language].missing}`; status.classList.add('warning'); } });
  const removeChartButton = document.createElement('button'); removeChartButton.className = 'secondary'; removeChartButton.id = 'remove-chart'; removeChartButton.textContent = chartLabels[language].remove; document.querySelector('.actions').append(removeChartButton); removeChartButton.addEventListener('click', () => { if (removeChart()) render(); else { const status = document.querySelector('#status'); status.textContent = `${t.status}: ${chartLabels[language].missing}`; status.classList.add('warning'); } });
  const pivotButton = document.createElement('button'); pivotButton.className = 'secondary'; pivotButton.id = 'create-pivot'; pivotButton.textContent = pivotLabels[language].add; document.querySelector('.actions').append(pivotButton); pivotButton.addEventListener('click', () => { if (createPivotSummary()) render(); else { const status = document.querySelector('#status'); status.textContent = `${t.status}: ${pivotLabels[language].invalid}`; status.classList.add('warning'); } });
  const refreshPivotButton = document.createElement('button'); refreshPivotButton.className = 'secondary'; refreshPivotButton.id = 'refresh-pivot'; refreshPivotButton.textContent = pivotLabels[language].refresh; document.querySelector('.actions').append(refreshPivotButton); refreshPivotButton.addEventListener('click', () => { const before = historySnapshot(); const refreshed = refreshPivotSummaries(); if (!refreshed) return; recordHistory(before, 'pivot-refresh'); dirty = true; render(); const status = document.querySelector('#status'); status.textContent = `${t.status}: ${pivotLabels[language].refreshed}`; });
  [['align-left', styleLabels[language].left, 'left'], ['align-center', styleLabels[language].center, 'center'], ['align-right', styleLabels[language].right, 'right'], ['toggle-wrap', styleLabels[language].wrap, 'wrap']].forEach(([id, label, value]) => { const button = document.createElement('button'); button.className = 'secondary'; button.id = id; button.textContent = label; styleToolbar.append(button); button.addEventListener('click', () => { if (applyCellStyle(value === 'wrap' ? 'wrap' : 'align', value === 'wrap' ? undefined : value)) render(); }); });
  document.querySelector('#download').addEventListener('click', downloadWorkbook);
  const buttons = { calculate: 'calculate', undo: 'undo', redo: 'redo', reset: 'reset', 'add-row': 'rowAdd', 'delete-row': 'rowDelete', 'add-column': 'columnAdd', 'delete-column': 'columnDelete', 'freeze-row': 'freeze', 'freeze-column': 'freeze', 'merge-cells': 'merge', 'unmerge-cells': 'unmerge', 'sort-asc': 'sortAsc', 'sort-desc': 'sortDesc', 'clear-filter': 'clear', 'add-validation': 'validation', 'create-table': 'table', 'define-name': 'definedName', 'add-comment': 'comment', 'add-chart-bar': 'chartBar', 'add-chart-line': 'chartLine', 'edit-chart': 'chartLine', 'resize-chart': 'chartLine', 'remove-chart': 'clear', 'create-pivot': 'pivot', 'refresh-pivot': 'pivot', 'toggle-bold': 'bold', 'copy-format': 'paste', 'conditional-format': 'conditional', 'align-left': 'alignLeft', 'align-center': 'alignCenter', 'align-right': 'alignRight', 'toggle-wrap': 'wrap', copy: 'copy', cut: 'copy', paste: 'paste', 'replace-next': 'replace', 'replace-all': 'replace', 'run-vba': 'run', download: 'download' };
  Object.entries(buttons).forEach(([id, icon]) => decorateButton(document.querySelector(`#${id}`), icon));
  const sheetWrap = document.querySelector('.sheet-wrap');
  if (sheetWrap) sheetWrap.addEventListener('scroll', () => {
    const point = cellPosition(selectedRef); const bounds = worksheetBounds(activeWorksheet());
    const atBottom = sheetWrap.scrollTop + sheetWrap.clientHeight >= sheetWrap.scrollHeight - 4;
    const atRight = sheetWrap.scrollLeft + sheetWrap.clientWidth >= sheetWrap.scrollWidth - 4;
    const atTop = sheetWrap.scrollTop <= 4; const atLeft = sheetWrap.scrollLeft <= 4;
    if (atBottom && point.r < bounds.e.r) moveSelection(Math.min(10, bounds.e.r - point.r), 0);
    else if (atTop && point.r > 0) moveSelection(-Math.min(10, point.r), 0);
    else if (atRight && point.c < bounds.e.c) moveSelection(0, Math.min(2, bounds.e.c - point.c));
    else if (atLeft && point.c > 0) moveSelection(0, -Math.min(2, point.c));
  }, { passive: true });
  revealSelectedCell();
}

const playgroundProfile = new URLSearchParams(window.location.search).get('profile');
resetValues();
workbook = playgroundProfile === 'large-sparse' || playgroundProfile === 'scroll-soak' ? makeLargeSparseWorkbook() : makeWorkbook();
if (playgroundProfile === 'large-sparse' || playgroundProfile === 'scroll-soak') activeSheet = 'Large sparse';
originalBytes = undefined;
try { const started = performance.now(); recalculate(); render(); document.documentElement.dataset.playgroundProfile = playgroundProfile || 'default'; document.documentElement.dataset.initialRenderMs = String(Math.round(performance.now() - started)); document.documentElement.dataset.viewportWidth = String(window.innerWidth); if (playgroundProfile === 'scroll-soak') { const soakStarted = performance.now(); for (let index = 0; index < 100; index += 1) { const row = 1 + Math.floor((199998 * index) / 99); selectedRef = encode_cell({ r: row, c: index % 24 }); setSingleSelection(selectedRef); render(); } document.documentElement.dataset.virtualScrollSoakMs = String(Math.round(performance.now() - soakStarted)); document.documentElement.dataset.virtualScrollSoakSteps = '100'; } } catch (error) { document.querySelector('#app').innerHTML = `<main class="page"><h1>elixcee Playground</h1><p>${copy[language].error}</p><pre>${escapeHTML(error)}</pre></main>`; }
