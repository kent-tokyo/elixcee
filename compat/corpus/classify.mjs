// The single source of truth for how an elixcee-vs-oracle divergence is classified, for
// the VBA-macro corpus. Reuses the verdict vocabulary and anti-laundering discipline of
// ../differential/classify.mjs (MATCH / INTENTIONAL_SECURITY_DIVERGENCE /
// INTENTIONAL_SAFETY_DIVERGENCE / UNSUPPORTED / BUG / ORACLE_AMBIGUITY / NONDETERMINISTIC
// / UNCLASSIFIED — see that file's doc comment for what each means) rather than importing
// it directly, because its registries (UNSUPPORTED_ALLOWLIST etc.) are keyed to the
// xlsx-npm-API domain (dot-path APIs like "utils.encode_col") and would be meaningless
// here. This file adds exactly one new verdict this domain needs that the xlsx-npm domain
// never does: ORACLE_UNAVAILABLE, for when the oracle backend itself failed to produce
// any comparable output (crash, timeout) — not a divergence in VALUE, a divergence in
// whether a comparison could happen at all. Every result record this produces carries the
// `oracle` field identifying which backend it came from, per this milestone's explicit
// requirement — see run-libreoffice.mjs's records and (future, unimplemented)
// oracle-excel-com's contract.
import { normalizeElixceeCells, normalizeLibreOfficeCells, cellsEqual } from './normalize.mjs';

export const VERDICTS = /** @type {const} */ ([
  'MATCH',
  'INTENTIONAL_SECURITY_DIVERGENCE',
  'INTENTIONAL_SAFETY_DIVERGENCE',
  'UNSUPPORTED',
  'BUG',
  'ORACLE_AMBIGUITY',
  'ORACLE_UNAVAILABLE',
  'NONDETERMINISTIC',
  'UNCLASSIFIED',
]);

// Registered "known-unimplemented" elixcee behavior, keyed by scenario category ->
// Set<scenarioId>. Same anti-laundering shape as ../differential/classify.mjs's
// UNSUPPORTED_ALLOWLIST: a category alone is never enough, the exact scenario id must be
// registered, so an unrelated bug under the same category can't hide behind an
// already-registered entry. Populated from run-elixcee.mjs's actual output (see
// populateUnsupportedFromElixceeErrors below) rather than hand-guessed, so it only ever
// contains cases that were actually observed to fail with an "unknown function" /
// "not implemented" shaped error — not a blanket "this category is probably fine".
export const UNSUPPORTED_ALLOWLIST = new Map();

/**
 * Scans elixcee's own raw results for the specific error shapes that mean "elixcee
 * doesn't implement this VBA feature yet" (as opposed to a crash on a feature it claims
 * to support) and registers those scenario ids into UNSUPPORTED_ALLOWLIST. This is the
 * ONLY way a scenario's elixcee failure classifies as UNSUPPORTED rather than
 * UNCLASSIFIED/BUG — mirrors ../differential/classify.mjs's rule that registration must
 * happen with a reason, not be inferred silently at classification time.
 * @param {Array<{id: string, category: string, ok: boolean, error?: {kind: string, message: string}}>} elixceeResults
 */
export function populateUnsupportedFromElixceeErrors(elixceeResults) {
  for (const r of elixceeResults) {
    if (r.ok) continue;
    const msg = r.error?.message || '';
    const isUnknownFunction = r.error?.kind === 'undefined_sub_or_function' && /Unknown VBA function/i.test(msg);
    const isNotImplemented = /is not implemented/i.test(msg);
    if (isUnknownFunction || isNotImplemented) {
      if (!UNSUPPORTED_ALLOWLIST.has(r.category)) UNSUPPORTED_ALLOWLIST.set(r.category, new Map());
      UNSUPPORTED_ALLOWLIST.get(r.category).set(r.id, msg);
    }
  }
}

/**
 * @param {object} input
 * @param {string} input.id scenario id
 * @param {string} input.category scenario category
 * @param {string} input.oracle oracle backend name, e.g. "libreoffice" or "microsoft_excel"
 * @param {{ok: boolean, cells?: Array, error?: object}} input.elixcee elixcee's raw result for this scenario
 * @param {{ok: boolean, status: string, cells?: Array}} input.oracleResult the oracle backend's raw result for this scenario
 * @returns {{id: string, category: string, oracle: string, classification: string, reason?: string}}
 */
export function classifyScenario({ id, category, oracle, elixcee, oracleResult }) {
  if (category === 'nondeterministic') {
    // Time-dependent by construction (Now()/Timer()) — comparing across two independent
    // process runs is meaningless regardless of engine, so this is never MATCH/BUG.
    return { id, category, oracle, classification: 'NONDETERMINISTIC', reason: 'scenario reads a live clock/timer value' };
  }

  if (!oracleResult || oracleResult.status === 'TIMEOUT' || oracleResult.status === 'NO_OUTPUT' || !oracleResult.ok) {
    return {
      id,
      category,
      oracle,
      classification: 'ORACLE_UNAVAILABLE',
      reason: oracleResult ? `oracle backend status: ${oracleResult.status}` : 'no oracle result recorded',
    };
  }

  if (!elixcee.ok) {
    const registered = UNSUPPORTED_ALLOWLIST.get(category)?.has(id);
    if (registered) {
      return { id, category, oracle, classification: 'UNSUPPORTED', reason: UNSUPPORTED_ALLOWLIST.get(category).get(id) };
    }
    return { id, category, oracle, classification: 'UNCLASSIFIED', reason: `elixcee failed: ${elixcee.error?.message || 'unknown error'}` };
  }

  const elixceeCells = normalizeElixceeCells(elixcee.cells || []);
  const oracleCells = normalizeLibreOfficeCells(oracleResult.cells || []);

  if (elixceeCells.some((c) => c.isPlaceholder)) {
    return {
      id,
      category,
      oracle,
      classification: 'UNCLASSIFIED',
      reason: 'elixcee result contains an "[array]"/"[record]" CLI serialization placeholder — not comparable by value',
    };
  }

  if (cellsEqual(elixceeCells, oracleCells)) {
    return { id, category, oracle, classification: 'MATCH' };
  }

  return {
    id,
    category,
    oracle,
    classification: 'UNCLASSIFIED',
    reason: 'elixcee and oracle both produced cell output, but it differs — needs human triage to become BUG or a registered divergence',
  };
}

// Per-category x per-verdict rollup, same rationale as
// ../differential/classify.mjs's summarizeByApiAndVerdict: every count in a report must
// trace back to this, never hand-typed arithmetic over a results file.
export function summarizeByCategoryAndVerdict(classifications) {
  const byCategory = new Map();
  for (const c of classifications) {
    if (!byCategory.has(c.category)) byCategory.set(c.category, new Map());
    const verdicts = byCategory.get(c.category);
    verdicts.set(c.classification, (verdicts.get(c.classification) || 0) + 1);
  }
  return byCategory;
}

export function summarizeOverall(classifications) {
  const totals = new Map();
  for (const c of classifications) totals.set(c.classification, (totals.get(c.classification) || 0) + 1);
  return totals;
}

// Runnable self-check: `node classify.mjs`
if (import.meta.url === `file://${process.argv[1]}`) {
  const assert = await import('node:assert/strict');

  // MATCH: identical cell output on both sides.
  assert.equal(
    classifyScenario({
      id: 's1',
      category: 'arithmetic',
      oracle: 'libreoffice',
      elixcee: { ok: true, cells: [{ sheet: 'sheet1', address: 'A1', value: 5 }] },
      oracleResult: { ok: true, status: 'DONE', cells: [{ address: '$Sheet1.$A$1', type: 'number', value: 5 }] },
    }).classification,
    'MATCH'
  );

  // ORACLE_UNAVAILABLE: the documented, expected outcome when the oracle backend times out.
  assert.equal(
    classifyScenario({
      id: 's2',
      category: 'range_readwrite',
      oracle: 'libreoffice',
      elixcee: { ok: true, cells: [{ sheet: 'sheet1', address: 'A1', value: 5 }] },
      oracleResult: { ok: false, status: 'TIMEOUT', cells: [] },
    }).classification,
    'ORACLE_UNAVAILABLE'
  );

  // NONDETERMINISTIC: forced regardless of whether cells happen to match, for the
  // nondeterministic category.
  assert.equal(
    classifyScenario({
      id: 's3',
      category: 'nondeterministic',
      oracle: 'libreoffice',
      elixcee: { ok: true, cells: [{ sheet: 'sheet1', address: 'A1', value: 123 }] },
      oracleResult: { ok: true, status: 'DONE', cells: [{ address: 'A1', type: 'number', value: 123 }] },
    }).classification,
    'NONDETERMINISTIC'
  );

  // UNSUPPORTED only after explicit registration from a real elixcee error shape — an
  // elixcee failure is UNCLASSIFIED, never auto-UNSUPPORTED, until registered.
  const before = classifyScenario({
    id: 'fn_0001',
    category: 'type_conversion',
    oracle: 'libreoffice',
    elixcee: { ok: false, error: { kind: 'undefined_sub_or_function', message: "Unknown VBA function: 'fix'" } },
    oracleResult: { ok: true, status: 'DONE', cells: [] },
  });
  assert.equal(before.classification, 'UNCLASSIFIED', 'not registered yet');

  populateUnsupportedFromElixceeErrors([
    { id: 'fn_0001', category: 'type_conversion', ok: false, error: { kind: 'undefined_sub_or_function', message: "Unknown VBA function: 'fix'" } },
  ]);
  const after = classifyScenario({
    id: 'fn_0001',
    category: 'type_conversion',
    oracle: 'libreoffice',
    elixcee: { ok: false, error: { kind: 'undefined_sub_or_function', message: "Unknown VBA function: 'fix'" } },
    oracleResult: { ok: true, status: 'DONE', cells: [] },
  });
  assert.equal(after.classification, 'UNSUPPORTED', 'registered after observing the real elixcee error shape');

  // summarizeByCategoryAndVerdict / summarizeOverall
  const sample = [
    { id: 'a', category: 'arithmetic', classification: 'MATCH' },
    { id: 'b', category: 'arithmetic', classification: 'ORACLE_UNAVAILABLE' },
    { id: 'c', category: 'range_readwrite', classification: 'ORACLE_UNAVAILABLE' },
  ];
  const byCat = summarizeByCategoryAndVerdict(sample);
  assert.deepEqual(Object.fromEntries(byCat.get('arithmetic')), { MATCH: 1, ORACLE_UNAVAILABLE: 1 });
  assert.deepEqual(Object.fromEntries(summarizeOverall(sample)), { MATCH: 1, ORACLE_UNAVAILABLE: 2 });

  console.log('classify.mjs self-check: all assertions passed');
}
