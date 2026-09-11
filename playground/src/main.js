import { aoa_to_sheet, book_append_sheet, book_new, encode_cell, decode_range, read, write } from '../../packages/xlsx/src/index.browser.mjs';
import { calculateWorkbook, diagnoseWorkbook } from '../../packages/xlsx/src/runtime.browser.mjs';
import './style.css';

const copy = {
  en: {
    title: 'elixcee Playground', intro: 'Edit spreadsheet values, recalculate a formula, and download a real XLSX file — entirely in your browser.',
    recipe: 'Choose a sample', sales: 'Sales total', budget: 'Budget summary', grades: 'Class average', multi: 'Multi-sheet workbook', sheetTab: 'Spreadsheet', downloadTab: 'Excel download', input: 'Input data', calculate: 'Recalculate', download: 'Download XLSX', reset: 'Reset', downloadTitle: 'Your workbook is ready', downloadText: 'Download the edited workbook as an XLSX file. Values and the calculated formula are included.',
    formula: 'Formula', result: 'Calculated result', ready: 'Rust/WASM formula engine ready', stale: 'Values changed — recalculate to update', status: 'Status', language: 'Language', diagnostic: 'Diagnostics', dependencies: 'dependencies', formulas: 'formulas', parseErrors: 'parse errors', cycle: 'cycle', yes: 'yes', no: 'no', exported: 'Export verified', vbaTitle: 'VBA sandbox', vbaHint: 'Bounded Worker subset: assignments, Cells(r,c).Value, Dim, and arithmetic.', vbaRun: 'Run VBA', vbaName: 'Macro name', vbaSource: 'VBA source', vbaExample: 'Sub Demo()\n  Cells(2, 3).Value = Cells(2, 2).Value + 5\nEnd Sub', vbaDone: 'VBA completed',
    boundary: 'Browser scope: XLSX read/write, formula recalculation, diagnostics, and a bounded VBA Worker subset. Full VBA data processing runs in the native runtime.',
    guideTitle: 'What this demonstrates', guideText: 'The browser package uses the Rust/WASM core for workbook edits and formula calculation, plus a small isolated VBA Worker. Download the result and continue with the Python API or CLI for full VBA.',
    guideLink: 'Open the quick start', guideHref: '../docs/quickstart.md', upload: 'Open an XLSX file', uploadHint: 'Choose a local .xlsx file. It stays in this browser.', imported: 'Imported workbook', selectCell: 'Select a cell', formulaBar: 'Formula bar', copy: 'Copy', paste: 'Paste', addRow: 'Add row', deleteRow: 'Delete row', addColumn: 'Add column', deleteColumn: 'Delete column', error: 'Could not process this workbook.', unsupported: 'This workbook could not be opened in the browser.', tooLarge: 'This file is larger than the 20 MiB browser limit.', structureBlocked: 'Structure editing is disabled while formulas need reference updates.',
  },
  ja: {
    title: 'elixcee Playground', intro: '表の値を編集し、数式を再計算して、本物のXLSXファイルをブラウザーだけでダウンロードできます。',
    recipe: 'サンプルを選択', sales: '売上合計', budget: '予算集計', grades: 'クラス平均', multi: '複数シート', sheetTab: 'Spreadsheet', downloadTab: 'Excelダウンロード', input: '入力データ', calculate: '再計算', download: 'XLSXをダウンロード', reset: 'リセット', downloadTitle: 'ワークブックの準備ができました', downloadText: '編集したワークブックをXLSXとしてダウンロードできます。値と計算済みの数式を含みます。',
    formula: '数式', result: '計算結果', ready: 'Rust/WASM数式エンジン準備完了', stale: '値を変更しました。再計算してください', status: '状態', language: '言語', diagnostic: '診断', dependencies: '依存関係', formulas: '数式', parseErrors: '解析エラー', cycle: '循環', yes: 'あり', no: 'なし', exported: '出力を検証済み', vbaTitle: 'VBAサンドボックス', vbaHint: 'Worker内の限定サブセット：代入、Cells(r,c).Value、Dim、算術演算。', vbaRun: 'VBAを実行', vbaName: 'マクロ名', vbaSource: 'VBAソース', vbaExample: 'Sub Demo()\n  Cells(2, 3).Value = Cells(2, 2).Value + 5\nEnd Sub', vbaDone: 'VBA実行完了',
    boundary: 'ブラウザー版の範囲：XLSXの読み書き、数式再計算、診断、限定VBA Worker。完全なVBAデータ処理はネイティブランタイムで実行します。',
    guideTitle: 'この画面で試せること', guideText: 'ブラウザー版は、ワークブック編集と数式計算にRust/WASMコアを使い、小型のVBA Workerも備えます。結果をダウンロードし、完全なVBAが必要ならPython APIまたはCLIへ進めます。',
    guideLink: 'クイックスタートを読む', guideHref: '../docs/quickstart-ja.md', upload: 'XLSXファイルを開く', uploadHint: 'ローカルの.xlsxを選択します。ファイルはこのブラウザー内で処理します。', imported: '読み込んだワークブック', selectCell: 'セルを選択', formulaBar: '数式バー', copy: 'コピー', paste: '貼り付け', addRow: '行を追加', deleteRow: '行を削除', addColumn: '列を追加', deleteColumn: '列を削除', error: 'ワークブックを処理できませんでした。', unsupported: 'このワークブックはブラウザーで開けませんでした。', tooLarge: 'ブラウザーで扱える上限（20 MiB）を超えています。', structureBlocked: '数式参照の更新が必要なため、構造編集を無効にしています。',
  },
  zh: {
    title: 'elixcee Playground', intro: '编辑表格数据、重新计算公式，并完全在浏览器中下载真实的 XLSX 文件。',
    recipe: '选择示例', sales: '销售总额', budget: '预算汇总', grades: '班级平均分', multi: '多工作表工作簿', sheetTab: 'Spreadsheet', downloadTab: '下载 Excel', input: '输入数据', calculate: '重新计算', download: '下载 XLSX', reset: '重置', downloadTitle: '工作簿已准备好', downloadText: '将编辑后的工作簿下载为 XLSX 文件，其中包含数据和计算后的公式。',
    formula: '公式', result: '计算结果', ready: 'Rust/WASM 公式引擎已准备就绪', stale: '数值已改变，请重新计算', status: '状态', language: '语言', diagnostic: '诊断', dependencies: '依赖关系', formulas: '公式', parseErrors: '解析错误', cycle: '循环', yes: '有', no: '无', exported: '已验证导出', vbaTitle: 'VBA 沙盒', vbaHint: 'Worker 限定子集：赋值、Cells(r,c).Value、Dim 和算术运算。', vbaRun: '运行 VBA', vbaName: '宏名称', vbaSource: 'VBA 源码', vbaExample: 'Sub Demo()\n  Cells(2, 3).Value = Cells(2, 2).Value + 5\nEnd Sub', vbaDone: 'VBA 已完成',
    boundary: '浏览器范围：XLSX 读写、公式计算、诊断和受限 VBA Worker。完整 VBA 数据处理在原生运行时中执行。',
    guideTitle: '此页面展示的功能', guideText: '浏览器包使用 Rust/WASM 核心完成工作簿编辑和公式计算，并提供隔离的小型 VBA Worker。下载结果后，如需完整 VBA 请继续使用 Python API 或 CLI。',
    guideLink: '打开快速开始', guideHref: '../docs/quickstart-zh.md', upload: '打开 XLSX 文件', uploadHint: '选择本地 .xlsx 文件。文件只在此浏览器中处理。', imported: '已导入工作簿', selectCell: '选择单元格', formulaBar: '公式栏', copy: '复制', paste: '粘贴', addRow: '添加行', deleteRow: '删除行', addColumn: '添加列', deleteColumn: '删除列', error: '无法处理此工作簿。', unsupported: '此工作簿无法在浏览器中打开。', tooLarge: '文件超过浏览器限制（20 MiB）。', structureBlocked: '公式需要更新引用，因此已禁用结构编辑。',
  },
};

const recipes = {
  sales: { label: 'sales', headers: ['Item', 'Units', 'Revenue'], rows: [['A', 12, 300], ['B', 8, 240], ['C', 15, 450]], formula: '=SUM(C2:C4)', resultColumn: 3 },
  budget: { label: 'budget', headers: ['Category', 'Amount'], rows: [['Hosting', 120], ['Tools', 80], ['Support', 200], ['Other', 50]], formula: '=SUM(B2:B5)', resultColumn: 2 },
  grades: { label: 'grades', headers: ['Student', 'Score'], rows: [['Aki', 84], ['Bo', 92], ['Chen', 76], ['Dana', 88]], formula: '=AVERAGE(B2:B5)', resultColumn: 2 },
  multi: { label: 'multi', headers: ['Item', 'Units', 'Revenue'], rows: [['A', 12, 300], ['B', 8, 240], ['C', 15, 450]], formula: '=SUM(C2:C4)', resultColumn: 3, extra: { name: 'Notes', rows: [['Workbook', 'Two sheets'], ['Purpose', 'Edit and calculate'], ['Runtime', 'Rust/WASM']]} },
};

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
let dragging = false;
let dragMoved = false;
let contextAxis = 'r';
let clipboardCell = '';
let workbookName = '';
let originalBytes;
let vbaStatus = '';
let exportStatus = '';
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
function escapeHTML(value) { return String(value ?? '').replace(/[&<>"']/g, (char) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[char])); }
function buildBytes() { return new Uint8Array(write(workbook, { bookType: 'xlsx', type: 'array' })); }
function cellValue(book, ref) { return book?.Sheets?.[activeSheet]?.[ref]?.v ?? ''; }
function activeWorksheet() { return workbook?.Sheets?.[activeSheet]; }
function worksheetBounds(ws) {
  if (ws?.['!ref']) return decode_range(ws['!ref']);
  let maxRow = 9; let maxCol = 5;
  for (const ref of Object.keys(ws || {})) { if (!/^[A-Z]+[0-9]+$/.test(ref)) continue; const range = decode_range(`${ref}:${ref}`); maxRow = Math.max(maxRow, range.e.r); maxCol = Math.max(maxCol, range.e.c); }
  return { s: { r: 0, c: 0 }, e: { r: maxRow, c: maxCol } };
}
function cellText(cell) { return cell?.f ? `=${cell.f}` : (cell?.v ?? ''); }
function workerCells(ws) { return Object.fromEntries(Object.keys(ws || {}).filter((key) => /^[A-Z]+[0-9]+$/.test(key)).map((key) => [key, cellText(ws[key])])); }
function dependencyCount(ws) { return Object.keys(ws || {}).filter((key) => /^[A-Z]+[0-9]+$/.test(key) && ws[key]?.f).reduce((count, key) => count + ((ws[key].f.match(/\b[A-Z]{1,3}\d+(?::[A-Z]{1,3}\d+)?\b/gi) || []).length), 0); }
function diagnosticText(t) { const d = result?.diagnostic || {}; return `${t.diagnostic}: ${d.formulaCount ?? 0} ${t.formulas}, ${dependencyCount(activeWorksheet())} ${t.dependencies}, ${d.formulaParseErrors ?? 0} ${t.parseErrors}, ${t.cycle} ${d.hasFormulaCycle ? t.yes : t.no}`; }
function cellPosition(ref) { return decode_range(`${ref}:${ref}`).s; }
function selectionBounds() { const start = cellPosition(selectionStart); const end = cellPosition(selectionEnd); return { s: { r: Math.min(start.r, end.r), c: Math.min(start.c, end.c) }, e: { r: Math.max(start.r, end.r), c: Math.max(start.c, end.c) } }; }
function selectionLabel() { const range = selectionBounds(); return encode_cell(range.s) === encode_cell(range.e) ? encode_cell(range.s) : `${encode_cell(range.s)}:${encode_cell(range.e)}`; }
function updateSelectionDom() { const range = selectionBounds(); document.querySelectorAll('.cell-input').forEach((input) => { const pos = cellPosition(input.dataset.ref); input.parentElement.classList.toggle('selected', pos.r >= range.s.r && pos.r <= range.e.r && pos.c >= range.s.c && pos.c <= range.e.c); }); }
function setCellText(ref, text) {
  const ws = activeWorksheet(); if (!ws) return;
  const value = String(text ?? '');
  if (!value) delete ws[ref];
  else if (value.startsWith('=')) ws[ref] = { f: value.slice(1), t: 'n' };
  else if (value.trim() !== '' && Number.isFinite(Number(value))) ws[ref] = { v: Number(value), t: 'n' };
  else ws[ref] = { v: value, t: 's' };
  const bounds = worksheetBounds(ws); ws['!ref'] = `A1:${encode_cell(bounds.e)}`;
}
function hasFormula(ws) { return Object.keys(ws || {}).some((key) => /^[A-Z]+[0-9]+$/.test(key) && ws[key]?.f); }
function structureEdit(axis, delta) {
  const ws = activeWorksheet(); if (!ws) return false;
  if (hasFormula(ws)) return false;
  const point = decode_range(`${selectedRef}:${selectedRef}`).s;
  const cells = Object.entries(ws).filter(([key]) => /^[A-Z]+[0-9]+$/.test(key));
  for (const [key, cell] of cells) {
    delete ws[key]; const pos = decode_range(`${key}:${key}`).s; const index = axis === 'row' ? pos.r : pos.c;
    if (delta < 0 && index === (axis === 'row' ? point.r : point.c)) continue;
    const next = { ...pos }; if (index >= (axis === 'row' ? point.r : point.c)) next[axis === 'row' ? 'r' : 'c'] += delta;
    if (next.r >= 0 && next.c >= 0) ws[encode_cell(next)] = cell;
  }
  const bounds = worksheetBounds(ws); ws['!ref'] = `A1:${encode_cell(bounds.e)}`; return true;
}
function recalculate() {
  const bytes = buildBytes();
  const diagnostic = JSON.parse(diagnoseWorkbook(bytes));
  const calculated = JSON.parse(calculateWorkbook(bytes));
  latestBytes = bytes;
  const ref = workbookName ? (firstFormulaRef(activeWorksheet()) || selectedRef) : (activeSheet === 'Summary' ? resultRef() : selectedRef);
  result = { value: cellValue(calculated, ref), diagnostic };
  if (calculated?.Sheets?.[activeSheet]?.[ref]) workbook.Sheets[activeSheet][ref] = calculated.Sheets[activeSheet][ref];
  dirty = false;
}
function downloadWorkbook() {
  const bytes = dirty ? buildBytes() : (latestBytes || buildBytes());
  try { const verified = read(bytes); if (!verified.SheetNames.length) throw new Error('no sheets'); latestBytes = bytes; exportStatus = `${copy[language].exported}: ${bytes.length} bytes, ${verified.SheetNames.length} sheet(s)`; }
  catch { exportStatus = copy[language].error; const status = document.querySelector('#download-status'); if (status) status.textContent = exportStatus; return; }
  const blob = new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const link = document.createElement('a'); link.href = URL.createObjectURL(blob); link.download = workbookName || `elixcee-${recipeKey}.xlsx`; document.body.append(link); link.click(); link.remove(); URL.revokeObjectURL(link.href);
  const status = document.querySelector('#download-status'); if (status) status.textContent = exportStatus;
}
function renderSheet(ws) {
  const bounds = worksheetBounds(ws); const maxRow = Math.min(bounds.e.r, 19); const maxCol = Math.min(bounds.e.c, 11);
  const headers = Array.from({ length: maxCol + 1 }, (_, col) => `<th class="column-label">${columnLabel(col)}</th>`).join('');
  const selected = selectionBounds();
  const rows = Array.from({ length: maxRow + 1 }, (_, row) => `<tr><th class="row-label" data-row="${row}">${row + 1}</th>${Array.from({ length: maxCol + 1 }, (_, col) => { const ref = encode_cell({ r: row, c: col }); const inSelection = row >= selected.s.r && row <= selected.e.r && col >= selected.s.c && col <= selected.e.c; return `<td class="${inSelection ? 'selected' : ''}"><input class="cell-input" data-ref="${ref}" value="${escapeHTML(cellText(ws?.[ref]))}" aria-label="${ref}" /></td>`; }).join('')}</tr>`).join('');
  return `<div class="sheet-wrap"><table class="sheet"><thead><tr><th class="corner"></th>${Array.from({ length: maxCol + 1 }, (_, col) => `<th class="column-label" data-col="${col}">${columnLabel(col)}</th>`).join('')}</tr></thead><tbody>${rows}</tbody></table></div>`;
}
function render() {
  const t = copy[language]; const recipe = recipes[recipeKey];
  document.documentElement.lang = language; localStorage.setItem('elixcee-playground-language', language);
  document.querySelector('#app').innerHTML = `<main class="page">
    <header class="header"><div class="brand"><span class="mark">✦</span><span>${t.title}</span></div>
      <label>${t.language}<select id="language"><option value="en" ${language === 'en' ? 'selected' : ''}>English</option><option value="ja" ${language === 'ja' ? 'selected' : ''}>日本語</option><option value="zh" ${language === 'zh' ? 'selected' : ''}>简体中文</option></select></label>
    </header>
    <p class="intro">${t.intro}</p>
    <section class="toolbar"><label>${t.recipe}<select id="recipe"><option value="sales" ${recipeKey === 'sales' ? 'selected' : ''}>${t.sales}</option><option value="budget" ${recipeKey === 'budget' ? 'selected' : ''}>${t.budget}</option><option value="grades" ${recipeKey === 'grades' ? 'selected' : ''}>${t.grades}</option><option value="multi" ${recipeKey === 'multi' ? 'selected' : ''}>${t.multi}</option></select></label><label class="upload"><span>${t.upload}</span><input id="upload" type="file" accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" /><small>${t.uploadHint}</small></label><span class="badge">${workbookName ? `${t.imported}: ${escapeHTML(workbookName)}` : `${recipe.headers.length} columns · ${recipe.rows.length} rows${recipe.extra ? ` · ${workbook.SheetNames.length} sheets` : ''}`}</span></section>
    <div class="tabs" role="tablist"><button class="tab ${activeTab === 'sheet' ? 'active' : ''}" id="tab-sheet" role="tab" aria-selected="${activeTab === 'sheet'}">▦ ${t.sheetTab}</button><button class="tab ${activeTab === 'download' ? 'active' : ''}" id="tab-download" role="tab" aria-selected="${activeTab === 'download'}">⇩ ${t.downloadTab}</button></div>
    <section id="panel-sheet" class="tab-panel ${activeTab === 'sheet' ? 'active' : ''}" role="tabpanel"><div class="grid"><div class="card editor"><div class="sheet-title"><h2>${t.input}</h2><span>${escapeHTML(activeSheet)}</span></div><div class="sheet-tabs" role="tablist">${workbook.SheetNames.map((name) => `<button class="sheet-tab ${name === activeSheet ? 'active' : ''}" data-sheet="${escapeHTML(name)}" role="tab">${escapeHTML(name)}</button>`).join('')}</div><div class="formula-bar"><label for="formula-input">${t.formulaBar}</label><input id="formula-input" value="${escapeHTML(cellText(activeWorksheet()?.[selectedRef]))}" aria-label="${t.formulaBar}" /><button class="secondary" id="copy">${t.copy}</button><button class="secondary" id="paste">${t.paste}</button></div>${renderSheet(activeWorksheet())}<div class="actions"><button id="calculate">${t.calculate}</button><button class="secondary" id="reset">${t.reset}</button><button class="secondary" id="add-row">${t.addRow}</button><button class="secondary" id="delete-row">${t.deleteRow}</button><button class="secondary" id="add-column">${t.addColumn}</button><button class="secondary" id="delete-column">${t.deleteColumn}</button></div><div class="vba-box"><h2>${t.vbaTitle}</h2><p class="hint">${t.vbaHint}</p><label>${t.vbaName}<input id="vba-name" value="Demo" /></label><label>${t.vbaSource}<textarea id="vba-source" rows="5">${escapeHTML(t.vbaExample)}</textarea></label><button id="run-vba">${t.vbaRun}</button><pre id="vba-result" class="vba-result" aria-live="polite">${escapeHTML(vbaStatus)}</pre></div></div><div class="card output"><h2>${t.result}</h2><div class="formula"><span>${t.selectCell}: ${selectionLabel()}</span><code>${escapeHTML(cellText(activeWorksheet()?.[selectedRef])) || '—'}</code></div><div class="answer" id="answer">${result?.value ?? '—'}</div><p class="diagnostics" id="diagnostics">${escapeHTML(diagnosticText(t))}</p><p class="status ${dirty ? 'warning' : ''}" id="status">${t.status}: ${dirty ? t.stale : t.ready}</p></div></div></section><div id="context-menu" class="context-menu" hidden><button data-context-action="add">${t.addRow}</button><button data-context-action="delete">${t.deleteRow}</button></div>
    <section id="panel-download" class="tab-panel ${activeTab === 'download' ? 'active' : ''}" role="tabpanel"><div class="card download-card"><div class="download-icon">XLSX</div><h2>${t.downloadTitle}</h2><p>${t.downloadText}</p><div class="file-meta"><span>elixcee-${recipeKey}.xlsx</span><span>${recipe.headers.length} columns · ${recipe.rows.length + 1} rows</span></div><button class="download" id="download">${t.download}</button><p class="export-status" id="download-status" aria-live="polite">${escapeHTML(exportStatus)}</p></div></section>
    <p class="boundary">${t.boundary}</p>
    <section class="card guide"><h2>${t.guideTitle}</h2><p>${t.guideText}</p><a href="${t.guideHref}">${t.guideLink} →</a></section><footer>Rust/WASM · <a href="https://github.com/kent-tokyo/elixcee">source on GitHub</a></footer>
  </main>`;
  document.querySelector('#language').addEventListener('change', (event) => { language = event.target.value; render(); });
  document.querySelector('#tab-sheet').addEventListener('click', () => { activeTab = 'sheet'; localStorage.setItem('elixcee-playground-tab', activeTab); render(); });
  document.querySelector('#tab-download').addEventListener('click', () => { activeTab = 'download'; localStorage.setItem('elixcee-playground-tab', activeTab); render(); });
  document.querySelector('#recipe').addEventListener('change', (event) => { recipeKey = event.target.value; resetValues(); workbook = makeWorkbook(); activeSheet = 'Summary'; selectedRef = 'A1'; selectionStart = 'A1'; selectionEnd = 'A1'; workbookName = ''; originalBytes = undefined; recalculate(); render(); });
  document.querySelector('#upload').addEventListener('change', async (event) => { const file = event.target.files[0]; if (!file) return; if (file.size > MAX_UPLOAD_BYTES) { document.querySelector('#status').textContent = `${t.status}: ${t.tooLarge}`; return; } try { originalBytes = new Uint8Array(await file.arrayBuffer()); workbook = read(originalBytes, { cellStyles: true }); activeSheet = workbook.SheetNames[0]; selectedRef = 'A1'; workbookName = file.name; recalculate(); render(); } catch (error) { document.querySelector('#status').textContent = `${t.status}: ${t.unsupported}`; console.error(error); } });
  document.querySelectorAll('.sheet-tab').forEach((button) => button.addEventListener('click', () => { activeSheet = button.dataset.sheet; selectedRef = 'A1'; selectionStart = 'A1'; selectionEnd = 'A1'; render(); }));
  document.querySelectorAll('.row-label, .column-label').forEach((header) => header.addEventListener('contextmenu', (event) => { event.preventDefault(); const isColumn = header.classList.contains('column-label'); contextAxis = isColumn ? 'c' : 'r'; const pos = isColumn ? { r: 0, c: Number(header.dataset.col) } : { r: Number(header.dataset.row), c: 0 }; selectionStart = encode_cell(pos); selectionEnd = selectionStart; selectedRef = selectionStart; const menu = document.querySelector('#context-menu'); menu.querySelector('[data-context-action="add"]').textContent = isColumn ? t.addColumn : t.addRow; menu.querySelector('[data-context-action="delete"]').textContent = isColumn ? t.deleteColumn : t.deleteRow; menu.hidden = false; menu.style.left = `${Math.min(event.clientX, window.innerWidth - 170)}px`; menu.style.top = `${Math.min(event.clientY, window.innerHeight - 100)}px`; }));
  document.querySelectorAll('.cell-input').forEach((input) => { input.addEventListener('mousedown', (event) => { if (event.button !== 0) return; dragging = true; dragMoved = false; selectionStart = event.target.dataset.ref; selectionEnd = selectionStart; selectedRef = selectionStart; updateSelectionDom(); }); input.addEventListener('mouseover', (event) => { if (!dragging) return; const ref = event.target.dataset.ref; if (ref !== selectionEnd) { dragMoved = true; selectionEnd = ref; updateSelectionDom(); } }); input.addEventListener('mouseup', () => { dragging = false; }); input.addEventListener('click', (event) => { const ref = event.target.dataset.ref; if (dragMoved) { dragMoved = false; event.preventDefault(); return; } if (event.shiftKey) selectionEnd = ref; else { selectionStart = ref; selectionEnd = ref; } selectedRef = ref; render(); document.querySelector(`[data-ref="${ref}"]`)?.focus(); }); input.addEventListener('focus', () => { selectedRef = input.dataset.ref; document.querySelector('#formula-input').value = cellText(activeWorksheet()?.[selectedRef]); }); input.addEventListener('input', (event) => { selectedRef = event.target.dataset.ref; setCellText(selectedRef, event.target.value); dirty = true; document.querySelector('#status').textContent = `${t.status}: ${t.stale}`; document.querySelector('#status').classList.add('warning'); }); });
  document.querySelector('#formula-input').addEventListener('input', (event) => { setCellText(selectedRef, event.target.value); dirty = true; document.querySelector('#status').textContent = `${t.status}: ${t.stale}`; document.querySelector('#status').classList.add('warning'); });
  document.querySelector('#run-vba').addEventListener('click', () => { const worker = new Worker('./assets/vba-worker.js', { type: 'module' }); const resultNode = document.querySelector('#vba-result'); resultNode.textContent = ''; worker.onmessage = (event) => { worker.terminate(); if (!event.data.ok) { vbaStatus = `Error: ${event.data.error}`; render(); return; } event.data.changes.forEach(({ ref, value }) => setCellText(ref, value)); dirty = event.data.changes.length > 0; vbaStatus = `${t.vbaDone}: ${event.data.statements} statement(s), ${event.data.changes.length} cell(s)`; render(); }; worker.onerror = () => { worker.terminate(); vbaStatus = t.error; render(); }; worker.postMessage({ source: document.querySelector('#vba-source').value, macroName: document.querySelector('#vba-name').value, cells: workerCells(activeWorksheet()) }); });
  document.querySelector('#copy').addEventListener('click', async () => { const range = selectionBounds(); clipboardCell = Array.from({ length: range.e.r - range.s.r + 1 }, (_, r) => Array.from({ length: range.e.c - range.s.c + 1 }, (_, c) => cellText(activeWorksheet()?.[encode_cell({ r: range.s.r + r, c: range.s.c + c })])).join('\t')).join('\n'); try { await navigator.clipboard.writeText(clipboardCell); } catch {} });
  document.querySelector('#paste').addEventListener('click', async () => { let text = clipboardCell; try { const systemText = await navigator.clipboard.readText(); if (systemText) text = systemText; } catch {} if (text !== '') { const rows = text.split(/\r?\n/).map((row) => row.split('\t')); const origin = cellPosition(selectionStart); rows.forEach((row, r) => row.forEach((value, c) => setCellText(encode_cell({ r: origin.r + r, c: origin.c + c }), value))); selectedRef = selectionStart; selectionEnd = encode_cell({ r: origin.r + rows.length - 1, c: origin.c + Math.max(...rows.map((row) => row.length)) - 1 }); dirty = true; render(); } });
  for (const [id, axis, delta] of [['add-row', 'r', 1], ['delete-row', 'r', -1], ['add-column', 'c', 1], ['delete-column', 'c', -1]]) document.querySelector(`#${id}`).addEventListener('click', () => { if (structureEdit(axis, delta)) { dirty = true; render(); } else { document.querySelector('#status').textContent = `${t.status}: ${t.structureBlocked}`; document.querySelector('#status').classList.add('warning'); } });
  document.querySelector('#context-menu').addEventListener('click', (event) => { const delta = event.target.dataset.contextAction === 'add' ? 1 : -1; document.querySelector('#context-menu').hidden = true; if (structureEdit(contextAxis, delta)) { dirty = true; render(); } else { document.querySelector('#status').textContent = `${t.status}: ${t.structureBlocked}`; document.querySelector('#status').classList.add('warning'); } });
  document.querySelector('#calculate').addEventListener('click', () => { try { recalculate(); render(); } catch { document.querySelector('#status').textContent = `${t.status}: ${t.error}`; } });
  document.querySelector('#reset').addEventListener('click', () => { if (originalBytes && workbookName) { workbook = read(originalBytes, { cellStyles: true }); activeSheet = workbook.SheetNames[0]; selectedRef = 'A1'; } else { resetValues(); workbook = makeWorkbook(); activeSheet = 'Summary'; selectedRef = 'A1'; } recalculate(); render(); });
  document.querySelector('#download').addEventListener('click', downloadWorkbook);
}

resetValues();
workbook = makeWorkbook();
originalBytes = undefined;
try { recalculate(); render(); } catch (error) { document.querySelector('#app').innerHTML = `<main class="page"><h1>elixcee Playground</h1><p>${copy[language].error}</p><pre>${escapeHTML(error)}</pre></main>`; }
