// The single source of truth for how a divergence between the oracle (xlsx@0.18.5) and
// @elixcee/xlsx is classified. Cross-referenced from docs/xlsx-security-model.md rather
// than redefined there. "Roughly the same" is never an acceptable verdict — every
// comparison must resolve to exactly one of these seven values.
//
// Comparisons operate on parsed LOGICAL shape (e.g. sheet_to_json output, cell-object
// shape), never raw bytes. XLSX.write embeds a timestamp in docProps/core.xml, so a raw
// byte comparison would spuriously report NONDETERMINISTIC on every run. Callers must
// normalize (e.g. strip Props timestamps) before calling classify().
//
// UNSUPPORTED and INTENTIONAL_SECURITY_DIVERGENCE are NOT settable by a boolean flag at
// the call site — that would let an unexplained failure be quietly laundered into "not a
// bug" with zero paper trail, which is exactly how a compatibility report ends up lying.
// Both require the divergence to be registered below FIRST, with a reason. Any divergence
// that isn't registered comes back UNCLASSIFIED, which every caller must treat as a test
// failure (see compat/differential/run-demo.mjs and any Phase 1A+ test file).

/**
 * @typedef {'MATCH'|'INTENTIONAL_SECURITY_DIVERGENCE'|'UNSUPPORTED'|'BUG'|'ORACLE_AMBIGUITY'|'NONDETERMINISTIC'|'UNCLASSIFIED'} Verdict
 */

export const VERDICTS = /** @type {const} */ ([
  'MATCH',
  'INTENTIONAL_SECURITY_DIVERGENCE',
  'UNSUPPORTED',
  'BUG',
  'ORACLE_AMBIGUITY',
  'NONDETERMINISTIC',
  'UNCLASSIFIED',
]);

/**
 * MATCH — the oracle's and elixcee's outputs are equal on the normalized comparable shape.
 * INTENTIONAL_SECURITY_DIVERGENCE — elixcee errored with a code registered in
 *   SECURITY_DIVERGENCE_REGISTRY (docs/xlsx-security-model.md's limits).
 * UNSUPPORTED — the oracle API is registered in UNSUPPORTED_ALLOWLIST as not implemented
 *   yet; not a correctness bug.
 * BUG — a divergence a human has triaged and confirmed is a real elixcee defect. classify()
 *   never returns this automatically (see UNCLASSIFIED) — it's assigned by whoever reviews
 *   an UNCLASSIFIED result and files it, the same anti-laundering reasoning as above.
 * ORACLE_AMBIGUITY — the oracle itself is inconsistent/underspecified for this input;
 *   can't be used as ground truth without human review. Requires oracleAmbiguityReason so
 *   the judgment call is on record, not silent.
 * NONDETERMINISTIC — the oracle's own output differs across repeated runs on identical
 *   input (oracleA vs oracleB, both real oracle calls).
 * UNCLASSIFIED — a divergence exists and isn't explained by either registry above. This is
 *   the default for "not yet triaged," not a soft pass: every runner must fail the run on
 *   it.
 */

// Registered "known-unimplemented" oracle APIs, keyed by dot-path (e.g. "utils.encode_col").
// Adding an entry here is the ONLY way a divergence classifies as UNSUPPORTED. Empty by
// default — populate only when an API is genuinely not implemented yet, never as a
// catch-all for "the output doesn't match and I don't know why."
export const UNSUPPORTED_ALLOWLIST = new Map([
  // 'utils.sheet_to_json' => 'Phase 0 demo placeholder — no elixcee implementation exists yet',
]);

// Registered intentional security divergences, keyed by the elixcee-side error code that
// signals them (see docs/xlsx-security-model.md's planned ELIXCEE_* codes). Empty by
// default — populated once real resource limits throw real codes.
export const SECURITY_DIVERGENCE_REGISTRY = new Map([
  // 'ELIXCEE_ZIP_ENTRY_LIMIT' => 'zip bomb protection, see docs/xlsx-security-model.md',
]);

/**
 * @param {object} input
 * @param {string} [input.api] dot-path of the oracle API under test (e.g. "utils.encode_col")
 *   — required for a divergence to ever resolve to UNSUPPORTED.
 * @param {unknown} [input.oracleA] first oracle run (or the only oracle run)
 * @param {unknown} [input.oracleB] a second, independent oracle run on the same input —
 *   used to prove the comparison plumbing itself (oracle vs. itself must MATCH after
 *   normalization). Optional; omit when only comparing oracle vs. elixcee.
 * @param {unknown} [input.elixcee] elixcee's output for the same input
 * @param {string} [input.elixceeErrorCode] the ELIXCEE_* code elixcee's implementation
 *   threw, if any — checked against SECURITY_DIVERGENCE_REGISTRY.
 * @param {boolean} [input.oracleAmbiguous] set when the oracle's own behavior for this
 *   input is known to be inconsistent/underspecified (human-flagged, not inferred here)
 * @param {string} [input.oracleAmbiguityReason] required when oracleAmbiguous is true —
 *   the judgment call must be documented, not silent.
 * @returns {Verdict}
 */
export function classify({
  api,
  oracleA,
  oracleB,
  elixcee,
  elixceeErrorCode,
  oracleAmbiguous = false,
  oracleAmbiguityReason,
} = {}) {
  if (oracleAmbiguous) {
    if (!oracleAmbiguityReason) {
      throw new Error('classify(): oracleAmbiguous requires oracleAmbiguityReason');
    }
    return 'ORACLE_AMBIGUITY';
  }

  if (elixceeErrorCode && SECURITY_DIVERGENCE_REGISTRY.has(elixceeErrorCode)) {
    return 'INTENTIONAL_SECURITY_DIVERGENCE';
  }

  if (oracleB !== undefined) {
    // Oracle-vs-itself comparison: proves the comparison plumbing, not elixcee.
    return deepEqual(oracleA, oracleB) ? 'MATCH' : 'NONDETERMINISTIC';
  }

  if (deepEqual(oracleA, elixcee)) return 'MATCH';

  if (api && UNSUPPORTED_ALLOWLIST.has(api)) return 'UNSUPPORTED';

  return 'UNCLASSIFIED';
}

function deepEqual(a, b) {
  if (Object.is(a, b)) return true;
  if (typeof a !== typeof b) return false;
  if (a === null || b === null) return false;
  if (typeof a !== 'object') return false;
  const aKeys = Object.keys(a);
  const bKeys = Object.keys(b);
  if (aKeys.length !== bKeys.length) return false;
  for (const key of aKeys) {
    if (!Object.prototype.hasOwnProperty.call(b, key)) return false;
    if (!deepEqual(a[key], b[key])) return false;
  }
  return true;
}

// Runnable self-check (no test framework): `node compat/differential/classify.mjs`
if (import.meta.url === `file://${process.argv[1]}`) {
  const assert = await import('node:assert/strict');

  assert.equal(classify({ oracleA: { a: 1 }, oracleB: { a: 1 } }), 'MATCH');
  assert.equal(classify({ oracleA: { a: 1 }, oracleB: { a: 2 } }), 'NONDETERMINISTIC');
  assert.equal(classify({ oracleA: { a: 1 }, elixcee: { a: 1 } }), 'MATCH');
  // An unexplained divergence must NEVER auto-resolve to UNSUPPORTED or BUG.
  assert.equal(classify({ api: 'utils.not_registered', oracleA: { a: 1 }, elixcee: { a: 2 } }), 'UNCLASSIFIED');
  assert.equal(classify({ oracleA: { a: 1 }, elixcee: { a: 2 } }), 'UNCLASSIFIED');
  assert.equal(
    classify({ oracleA: {}, elixcee: {}, oracleAmbiguous: true, oracleAmbiguityReason: 'test' }),
    'ORACLE_AMBIGUITY'
  );
  assert.throws(() => classify({ oracleAmbiguous: true }), /oracleAmbiguityReason/);

  // UNSUPPORTED and INTENTIONAL_SECURITY_DIVERGENCE only fire for registered entries —
  // exercised for real once Phase 1A populates the registries with actual entries.
  assert.equal(UNSUPPORTED_ALLOWLIST.size, 0, 'Phase 0: allowlist should start empty');
  assert.equal(SECURITY_DIVERGENCE_REGISTRY.size, 0, 'Phase 0: security registry should start empty');

  console.log('classify.mjs self-check: all assertions passed');
}
