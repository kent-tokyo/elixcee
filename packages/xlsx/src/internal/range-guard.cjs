'use strict';

// sheet_to_formulae, sheet_to_csv, sheet_to_txt (which delegates to sheet_to_csv),
// sheet_to_json, and sheet_to_html all iterate every (row, col) pair inside a
// worksheet's `!ref` rectangle, regardless of how sparse the actual data is. Confirmed live against the real oracle (xlsx@0.18.5,
// timeout-guarded subprocess, not assumed): a crafted `!ref` spanning Excel's actual
// grid maximum ('A1:XFD1048576', ~17.18 billion cells) does not return within 25s on
// sheet_to_csv; even much smaller full-rectangle spans are already slow on the oracle
// itself (26,000,000 cells: 12-16s; 16,777,216 cells: ~20s; 2,600,000 cells: ~2.2s) —
// this is a genuine property of the real algorithm, not something introduced by this
// port. Per docs/xlsx-compatibility-goal.md's threat model (malicious/resource-exceeding
// input -> safe deterministic error, never a replicated vulnerability), elixcee rejects
// ranges above this threshold instead of iterating them. Registered as
// ELIXCEE_RANGE_TOO_LARGE in compat/differential/classify.mjs's SAFETY_DIVERGENCE_REGISTRY
// (a caller-argument-shaped issue, not file-parsing — this package has no file reader
// yet — matching the ELIXCEE_NON_FINITE_INDEX / encode_col(Infinity) precedent).
//
// 5,000,000 is not the oracle's own limit (it has none) — it's a threshold chosen from
// the measurements above to stay comfortably under ~4-5s worst case at the observed
// oracle throughput (~1.3M cells/sec) while still permitting worksheets far larger than
// any real spreadsheet is likely to populate via this in-memory API.
const ELIXCEE_RANGE_TOO_LARGE = 'ELIXCEE_RANGE_TOO_LARGE';
const MAX_RANGE_CELLS = 5000000;

function checkRangeSize(range) {
  const rows = range.e.r - range.s.r + 1;
  const cols = range.e.c - range.s.c + 1;
  if (rows > 0 && cols > 0 && rows * cols > MAX_RANGE_CELLS) {
    const err = new RangeError(
      '!ref spans ' + rows * cols + ' cells, exceeding the ' + MAX_RANGE_CELLS + '-cell safety limit'
    );
    err.code = ELIXCEE_RANGE_TOO_LARGE;
    throw err;
  }
}

module.exports = { checkRangeSize, ELIXCEE_RANGE_TOO_LARGE, MAX_RANGE_CELLS };
