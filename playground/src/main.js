import { aoa_to_sheet, book_append_sheet, book_new, write } from '../../packages/xlsx/src/index.browser.mjs';
import { calculateWorkbook, diagnoseWorkbook } from '../../packages/xlsx/src/runtime.browser.mjs';
import './style.css';

const copy = {
  en: {
    title: 'elixcee Playground', intro: 'Edit spreadsheet values, recalculate a formula, and download a real XLSX file — entirely in your browser.',
    recipe: 'Choose a sample', sales: 'Sales total', budget: 'Budget summary', grades: 'Class average', input: 'Input data', calculate: 'Recalculate', download: 'Download XLSX', reset: 'Reset',
    formula: 'Formula', result: 'Calculated result', ready: 'Rust/WASM formula engine ready', stale: 'Values changed — recalculate to update', status: 'Status', language: 'Language',
    boundary: 'Browser scope: XLSX read/write and formula recalculation. Python and data-processing VBA run in the native runtime.',
    guideTitle: 'What this demonstrates', guideText: 'The browser package uses the same Rust/WASM core for typed workbook edits and formula calculation. Download the result and continue with the Python API or CLI when you need VBA.',
    guideLink: 'Open the quick start', guideHref: '../docs/quickstart.md', error: 'Could not process this workbook.',
  },
  ja: {
    title: 'elixcee Playground', intro: '表の値を編集し、数式を再計算して、本物のXLSXファイルをブラウザーだけでダウンロードできます。',
    recipe: 'サンプルを選択', sales: '売上合計', budget: '予算集計', grades: 'クラス平均', input: '入力データ', calculate: '再計算', download: 'XLSXをダウンロード', reset: 'リセット',
    formula: '数式', result: '計算結果', ready: 'Rust/WASM数式エンジン準備完了', stale: '値を変更しました。再計算してください', status: '状態', language: '言語',
    boundary: 'ブラウザー版の範囲：XLSXの読み書きと数式再計算。Pythonとデータ処理VBAはネイティブランタイムで実行します。',
    guideTitle: 'この画面で試せること', guideText: 'ブラウザー版は、型付きワークブック編集と数式計算に同じRust/WASMコアを使います。結果をダウンロードし、VBAが必要ならPython APIまたはCLIへ進めます。',
    guideLink: 'クイックスタートを読む', guideHref: '../docs/quickstart-ja.md', error: 'ワークブックを処理できませんでした。',
  },
  zh: {
    title: 'elixcee Playground', intro: '编辑表格数据、重新计算公式，并完全在浏览器中下载真实的 XLSX 文件。',
    recipe: '选择示例', sales: '销售总额', budget: '预算汇总', grades: '班级平均分', input: '输入数据', calculate: '重新计算', download: '下载 XLSX', reset: '重置',
    formula: '公式', result: '计算结果', ready: 'Rust/WASM 公式引擎已准备就绪', stale: '数值已改变，请重新计算', status: '状态', language: '语言',
    boundary: '浏览器范围：XLSX 读写和公式计算。Python 和数据处理 VBA 在原生运行时中执行。',
    guideTitle: '此页面展示的功能', guideText: '浏览器包使用相同的 Rust/WASM 核心完成类型化工作簿编辑和公式计算。下载结果后，如需 VBA 请继续使用 Python API 或 CLI。',
    guideLink: '打开快速开始', guideHref: '../docs/quickstart-zh.md', error: '无法处理此工作簿。',
  },
};

const recipes = {
  sales: { label: 'sales', headers: ['Item', 'Units', 'Revenue'], rows: [['A', 12, 300], ['B', 8, 240], ['C', 15, 450]], formula: '=SUM(C2:C4)', resultColumn: 3 },
  budget: { label: 'budget', headers: ['Category', 'Amount'], rows: [['Hosting', 120], ['Tools', 80], ['Support', 200], ['Other', 50]], formula: '=SUM(B2:B5)', resultColumn: 2 },
  grades: { label: 'grades', headers: ['Student', 'Score'], rows: [['Aki', 84], ['Bo', 92], ['Chen', 76], ['Dana', 88]], formula: '=AVERAGE(B2:B5)', resultColumn: 2 },
};

let language = localStorage.getItem('elixcee-playground-language') || (navigator.language.toLowerCase().startsWith('ja') ? 'ja' : navigator.language.toLowerCase().startsWith('zh') ? 'zh' : 'en');
let recipeKey = 'sales';
let values = [];
let latestBytes;
let result;
let dirty = false;

function resetValues() { values = recipes[recipeKey].rows.map((row) => [...row]); dirty = false; }
function resultRef() { const recipe = recipes[recipeKey]; return `${String.fromCharCode(65 + recipe.resultColumn - 1)}${recipe.rows.length + 2}`; }
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
  return workbook;
}
function buildBytes() { return new Uint8Array(write(makeWorkbook(), { bookType: 'xlsx', type: 'array' })); }
function cellValue(workbook, ref) { return workbook?.Sheets?.Summary?.[ref]?.v ?? ''; }
function recalculate() {
  const bytes = buildBytes();
  const diagnostic = JSON.parse(diagnoseWorkbook(bytes));
  const calculated = JSON.parse(calculateWorkbook(bytes));
  latestBytes = bytes;
  result = { value: cellValue(calculated, resultRef()), diagnostic };
  dirty = false;
}
function downloadWorkbook() {
  const blob = new Blob([latestBytes || buildBytes()], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const link = document.createElement('a'); link.href = URL.createObjectURL(blob); link.download = `elixcee-${recipeKey}.xlsx`; link.click(); URL.revokeObjectURL(link.href);
}
function render() {
  const t = copy[language]; const recipe = recipes[recipeKey];
  document.documentElement.lang = language; localStorage.setItem('elixcee-playground-language', language);
  document.querySelector('#app').innerHTML = `<main class="page">
    <header class="header"><div class="brand"><span class="mark">✦</span><span>${t.title}</span></div>
      <label>${t.language}<select id="language"><option value="en" ${language === 'en' ? 'selected' : ''}>English</option><option value="ja" ${language === 'ja' ? 'selected' : ''}>日本語</option><option value="zh" ${language === 'zh' ? 'selected' : ''}>简体中文</option></select></label>
    </header>
    <p class="intro">${t.intro}</p>
    <section class="toolbar"><label>${t.recipe}<select id="recipe"><option value="sales" ${recipeKey === 'sales' ? 'selected' : ''}>${t.sales}</option><option value="budget" ${recipeKey === 'budget' ? 'selected' : ''}>${t.budget}</option><option value="grades" ${recipeKey === 'grades' ? 'selected' : ''}>${t.grades}</option></select></label><span class="badge">${recipe.headers.length} columns · ${recipe.rows.length} rows</span></section>
    <section class="grid"><div class="card editor"><div class="sheet-title"><h2>${t.input}</h2><span>Summary</span></div><div class="sheet-wrap"><table class="sheet"><thead><tr><th class="corner"></th>${recipe.headers.map((_, index) => `<th class="column-label">${columnLabel(index)}</th>`).join('')}</tr></thead><tbody>${values.map((row, rowIndex) => `<tr><th class="row-label">${rowIndex + 1}</th>${row.map((value, colIndex) => `<td>${typeof value === 'number' ? `<input class="cell-input" data-row="${rowIndex}" data-col="${colIndex}" type="number" value="${value}" aria-label="${recipe.headers[colIndex]} ${rowIndex + 1}" />` : `<input class="cell-input" data-row="${rowIndex}" data-col="${colIndex}" type="text" value="${value}" aria-label="${recipe.headers[colIndex]} ${rowIndex + 1}" />`}</td>`).join('')}</tr>`).join('')}<tr class="formula-row"><th class="row-label">${values.length + 2}</th>${recipe.headers.map((header, index) => index === 0 ? '<td>Result</td>' : index === recipe.resultColumn - 1 ? `<td class="formula-cell">${result?.value ?? '—'}</td>` : '<td></td>').join('')}</tr></tbody></table></div><div class="actions"><button id="calculate">${t.calculate}</button><button class="secondary" id="reset">${t.reset}</button></div></div>
      <div class="card output"><h2>${t.result}</h2><div class="formula"><span>${t.formula}</span><code>${recipe.formula}</code></div><div class="answer" id="answer">${result?.value ?? '—'}</div><p class="status ${dirty ? 'warning' : ''}" id="status">${t.status}: ${dirty ? t.stale : t.ready}</p></div></section>
    <p class="boundary">${t.boundary}</p><button class="download" id="download">${t.download}</button>
    <section class="card guide"><h2>${t.guideTitle}</h2><p>${t.guideText}</p><a href="${t.guideHref}">${t.guideLink} →</a></section><footer>Rust/WASM · <a href="https://github.com/kent-tokyo/elixcee">source on GitHub</a></footer>
  </main>`;
  document.querySelector('#language').addEventListener('change', (event) => { language = event.target.value; render(); });
  document.querySelector('#recipe').addEventListener('change', (event) => { recipeKey = event.target.value; resetValues(); recalculate(); render(); });
  document.querySelectorAll('.cell-input').forEach((input) => input.addEventListener('input', (event) => { const { row, col } = event.target.dataset; values[row][col] = event.target.type === 'number' ? Number(event.target.value) || 0 : event.target.value; dirty = true; document.querySelector('#status').textContent = `${t.status}: ${t.stale}`; document.querySelector('#status').classList.add('warning'); }));
  document.querySelector('#calculate').addEventListener('click', () => { try { recalculate(); render(); } catch { document.querySelector('#status').textContent = `${t.status}: ${t.error}`; } });
  document.querySelector('#reset').addEventListener('click', () => { resetValues(); recalculate(); render(); });
  document.querySelector('#download').addEventListener('click', downloadWorkbook);
}

resetValues();
try { recalculate(); render(); } catch (error) { document.querySelector('#app').innerHTML = `<main class="page"><h1>elixcee Playground</h1><p>${copy[language].error}</p><pre>${String(error)}</pre></main>`; }
