import { aoa_to_sheet, book_append_sheet, book_new, read, write } from '../../packages/xlsx/src/index.browser.mjs';
import { WorkbookEditor, calculateWorkbook, diagnoseWorkbook } from '../../packages/xlsx/src/runtime.browser.mjs';
import './style.css';

const copy = {
  en: {
    title: 'elixcee Playground', intro: 'Edit a workbook in your browser. Nothing is uploaded.',
    values: 'Sales values', calculate: 'Recalculate', download: 'Download XLSX', reset: 'Reset',
    formula: 'Formula', result: 'Calculated result', status: 'Status', ready: 'Rust/WASM runtime ready',
    working: 'Calculating…', boundary: 'Browser scope: XLSX read/write and typed edits. Python and VBA run in the native/Python runtime.',
    error: 'Could not process this workbook.', language: 'Language', input: 'Value',
    guideTitle: 'Start here', guideText: 'Edit the values, recalculate the formula, then download the workbook. The same runtime can also edit files and run data-processing VBA outside the browser.',
    guideLink: 'Read the English quick start', guideHref: '../docs/quickstart.md',
  },
  ja: {
    title: 'elixcee Playground', intro: 'ブラウザーでワークブックを編集できます。データはアップロードされません。',
    values: '売上値', calculate: '再計算', download: 'XLSXをダウンロード', reset: 'リセット',
    formula: '数式', result: '計算結果', status: '状態', ready: 'Rust/WASMランタイムの準備完了',
    working: '計算中…', boundary: 'ブラウザー版の範囲：XLSXの読み書きと型付き編集。PythonとVBAはネイティブ/Pythonランタイムで実行します。',
    error: 'ワークブックを処理できませんでした。', language: '言語', input: '値',
    guideTitle: 'まずはここから', guideText: '値を編集して数式を再計算し、ワークブックをダウンロードできます。同じランタイムで、ブラウザーの外ではファイル編集とデータ処理VBAも実行できます。',
    guideLink: '日本語クイックスタートを読む', guideHref: '../docs/quickstart-ja.md',
  },
  zh: {
    title: 'elixcee Playground', intro: '直接在浏览器中编辑工作簿。不会上传任何数据。',
    values: '销售数值', calculate: '重新计算', download: '下载 XLSX', reset: '重置',
    formula: '公式', result: '计算结果', status: '状态', ready: 'Rust/WASM 运行时已准备就绪',
    working: '正在计算…', boundary: '浏览器范围：XLSX 读写和类型化编辑。Python 和 VBA 在原生/Python 运行时中执行。',
    error: '无法处理此工作簿。', language: '语言', input: '数值',
    guideTitle: '从这里开始', guideText: '编辑数值、重新计算公式，然后下载工作簿。同一个运行时还可以在浏览器之外编辑文件并运行数据处理 VBA。',
    guideLink: '阅读简体中文快速开始', guideHref: '../docs/quickstart-zh.md',
  },
};

const initialValues = [120, 250, 180];
let values = [...initialValues];
let language = localStorage.getItem('elixcee-playground-language') || (navigator.language.toLowerCase().startsWith('ja') ? 'ja' : navigator.language.toLowerCase().startsWith('zh') ? 'zh' : 'en');
let latestBytes;
let result = null;

function makeWorkbook() {
  const sheet = aoa_to_sheet([
    ['Item', 'Sales'], ['A', values[0]], ['B', values[1]], ['C', values[2]], ['Total', { f: 'SUM(B2:B4)', t: 'n' }],
  ]);
  const workbook = book_new();
  book_append_sheet(workbook, sheet, 'Summary');
  return workbook;
}

function buildBytes() {
  return new Uint8Array(write(makeWorkbook(), { bookType: 'xlsx', type: 'array' }));
}

function cellValue(workbook, ref) {
  return workbook?.Sheets?.Summary?.[ref]?.v ?? '';
}

function runCalculation() {
  const bytes = buildBytes();
  const diagnostic = JSON.parse(diagnoseWorkbook(bytes));
  const calculated = JSON.parse(calculateWorkbook(bytes));
  latestBytes = bytes;
  result = { value: cellValue(calculated, 'B5'), diagnostic };
}

function runEditor(value) {
  const editor = new WorkbookEditor(buildBytes());
  editor.setNumber('Summary', 2, 2, value);
  const snapshot = JSON.parse(editor.recalculate());
  editor.free();
  result = { value: cellValue(snapshot, 'B5'), diagnostic: result?.diagnostic };
}

function downloadWorkbook() {
  const bytes = latestBytes || buildBytes();
  const blob = new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const link = document.createElement('a');
  link.href = URL.createObjectURL(blob);
  link.download = 'elixcee-playground.xlsx';
  link.click();
  URL.revokeObjectURL(link.href);
}

function render() {
  const t = copy[language];
  document.documentElement.lang = language;
  localStorage.setItem('elixcee-playground-language', language);
  document.querySelector('#app').innerHTML = `
    <main class="page">
      <header class="header"><div class="brand"><span class="mark">✦</span><span>${t.title}</span></div>
        <label>${t.language}<select id="language"><option value="en" ${language === 'en' ? 'selected' : ''}>English</option><option value="ja" ${language === 'ja' ? 'selected' : ''}>日本語</option><option value="zh" ${language === 'zh' ? 'selected' : ''}>简体中文</option></select></label>
      </header>
      <p class="intro">${t.intro}</p>
      <section class="grid">
        <div class="card editor"><h2>${t.values}</h2>
          ${values.map((value, i) => `<label>B${i + 2}<input class="value" aria-label="B${i + 2}" data-index="${i}" type="number" value="${value}" /></label>`).join('')}
          <div class="actions"><button id="calculate">${t.calculate}</button><button class="secondary" id="reset">${t.reset}</button></div>
        </div>
        <div class="card output"><h2>${t.result}</h2><div class="formula"><span>${t.formula}</span><code>=SUM(B2:B4)</code></div><div class="answer" id="answer">${result?.value ?? '—'}</div><p class="status" id="status">${t.status}: ${t.ready}</p></div>
      </section>
      <p class="boundary">${t.boundary}</p>
      <button class="download" id="download">${t.download}</button>
      <section class="card guide"><h2>${t.guideTitle}</h2><p>${t.guideText}</p><a href="${t.guideHref}">${t.guideLink} →</a></section>
      <footer>Rust/WASM · <a href="https://github.com/kent-tokyo/elixcee">source on GitHub</a></footer>
    </main>`;
  document.querySelector('#language').addEventListener('change', (e) => { language = e.target.value; render(); });
  document.querySelectorAll('.value').forEach((input) => input.addEventListener('input', (e) => { values[e.target.dataset.index] = Number(e.target.value) || 0; }));
  document.querySelector('#calculate').addEventListener('click', () => { try { runCalculation(); render(); } catch { document.querySelector('#status').textContent = `${t.status}: ${t.error}`; } });
  document.querySelector('#reset').addEventListener('click', () => { values = [...initialValues]; runCalculation(); render(); });
  document.querySelector('#download').addEventListener('click', downloadWorkbook);
}

try { runCalculation(); render(); } catch (error) { document.querySelector('#app').innerHTML = `<main class="page"><h1>elixcee Playground</h1><p>${copy[language].error}</p><pre>${String(error)}</pre></main>`; }
