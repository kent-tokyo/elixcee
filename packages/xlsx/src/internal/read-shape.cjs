'use strict';

// Shapes the WASM bridge's raw parsed JSON (crates/elixcee-wasm's readWorkbook output,
// after JSON.parse) into the WorkBook shape @elixcee/xlsx's read() promises. Kept as its
// own module — not inlined into index.cjs's `read` — so both the Node and browser read()
// entry points (see index.cjs and the browser entry added for the "browser" export
// condition) share one implementation instead of drifting.
//
// ---- !rows / !cols (Milestone read-item 3) ----
//
// crates/elixcee-wasm emits reader.rs's already-parsed hidden-row/col data as internal
// "!hiddenRows"/"!hiddenCols" keys: 1-based inclusive [start,end] intervals — reader.rs's
// own native shape (see src/reader.rs's WorkbookSheet.hidden_rows/hidden_columns), NOT the
// oracle's own !rows/!cols shape (a 0-based array, sparse — a real hole at every
// non-hidden index, not `undefined`/`null` filler — of {hidden:true} objects, one slot per
// row/col up to the last hidden one). expandHiddenIntervals does that expansion.
//
// opts.cellStyles gates whether !rows/!cols are surfaced at all — confirmed live against
// the real oracle (not assumed): XLSX.read(buf, {type:'buffer'}) NEVER returns !rows/!cols,
// even for a file with real hidden rows, unless the caller also passes
// {cellStyles: true} (see compat/node_modules/xlsx/xlsx.js's parse_ws_xml_cols call site
// and parse_ws_xml_data's `if(opts && opts.cellStyles)` guards around rowinfo/colinfo).
// Skipping this gate would silently diverge from the oracle's default-opts behavior on any
// fixture with hidden rows/cols — so it's threaded through here rather than always-on.
// The internal !hiddenRows/!hiddenCols keys are deleted unconditionally either way — they
// must never leak to a caller regardless of opts.
function expandHiddenIntervals(intervals) {
  let maxIdx = -1;
  for (const [, end] of intervals) {
    const idx = end - 1; // 1-based inclusive end -> 0-based index
    if (idx > maxIdx) maxIdx = idx;
  }
  const out = new Array(maxIdx + 1); // starts fully sparse — only hidden slots get set below
  for (const [start, end] of intervals) {
    for (let i = start - 1; i <= end - 1; ++i) out[i] = { hidden: true };
  }
  return out;
}

function shapeSheet(ws, opts) {
  if (ws == null) return ws;
  const hiddenRows = ws['!hiddenRows'];
  const hiddenCols = ws['!hiddenCols'];
  delete ws['!hiddenRows'];
  delete ws['!hiddenCols'];
  if (opts && opts.cellStyles) {
    if (hiddenRows) ws['!rows'] = expandHiddenIntervals(hiddenRows);
    if (hiddenCols) ws['!cols'] = expandHiddenIntervals(hiddenCols);
  }
  return ws;
}

// Mutates and returns `wb` in place — the WASM bridge's JSON.parse output is a fresh
// object with no other owner, so there's no reason to clone before reshaping it.
function shapeWorkBook(wb, opts) {
  for (const name of wb.SheetNames) shapeSheet(wb.Sheets[name], opts);
  return wb;
}

module.exports = { shapeWorkBook };
