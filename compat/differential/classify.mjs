// The single source of truth for how a divergence between the oracle (xlsx@0.18.5) and
// @elixcee/xlsx is classified. Cross-referenced from docs/xlsx-security-model.md rather
// than redefined there. "Roughly the same" is never an acceptable verdict — every
// comparison must resolve to exactly one of these eight values.
//
// Comparisons operate on parsed LOGICAL shape (e.g. sheet_to_json output, cell-object
// shape), never raw bytes. XLSX.write embeds a timestamp in docProps/core.xml, so a raw
// byte comparison would spuriously report NONDETERMINISTIC on every run. Callers must
// normalize (e.g. strip Props timestamps) before calling classify() — see
// compat/differential/normalize.mjs for a normalizer that preserves the distinctions a
// naive JSON round-trip silently erases (undefined vs. null, array holes, -0, etc.).
//
// UNSUPPORTED, INTENTIONAL_SECURITY_DIVERGENCE, and INTENTIONAL_SAFETY_DIVERGENCE are NOT
// settable by a boolean flag at the call site — that would let an unexplained failure be
// quietly laundered into "not a bug" with zero paper trail, which is exactly how a
// compatibility report ends up lying. All three require the divergence to be registered
// below FIRST, with a reason. Any divergence that isn't registered comes back
// UNCLASSIFIED, which every caller must treat as a test failure (see
// compat/differential/run-demo.mjs and any Phase 1A+ test file).

/**
 * @typedef {'MATCH'|'INTENTIONAL_SECURITY_DIVERGENCE'|'INTENTIONAL_SAFETY_DIVERGENCE'|'UNSUPPORTED'|'BUG'|'ORACLE_AMBIGUITY'|'NONDETERMINISTIC'|'UNCLASSIFIED'} Verdict
 */

export const VERDICTS = /** @type {const} */ ([
  'MATCH',
  'INTENTIONAL_SECURITY_DIVERGENCE',
  'INTENTIONAL_SAFETY_DIVERGENCE',
  'UNSUPPORTED',
  'BUG',
  'ORACLE_AMBIGUITY',
  'NONDETERMINISTIC',
  'UNCLASSIFIED',
]);

/**
 * MATCH — the oracle's and elixcee's outputs are equal on the normalized comparable shape.
 * INTENTIONAL_SECURITY_DIVERGENCE — elixcee errored with a code registered in
 *   SECURITY_DIVERGENCE_REGISTRY: untrusted-FILE-parsing attack surfaces (zip bombs, XML
 *   blowup, prototype pollution) — see docs/xlsx-security-model.md.
 * INTENTIONAL_SAFETY_DIVERGENCE — elixcee errored with a code registered in
 *   SAFETY_DIVERGENCE_REGISTRY: pure-function robustness against pathological JS values
 *   passed directly by a caller (e.g. Infinity causing an unbounded loop), independent of
 *   any file parsing. Kept distinct from the security registry because the two have
 *   different threat models (untrusted file vs. untrusted argument) even though both are
 *   "don't replicate a DoS" divergences.
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
 * UNCLASSIFIED — a divergence exists and isn't explained by any registry above. This is
 *   the default for "not yet triaged," not a soft pass: every runner must fail the run on
 *   it.
 */

// Registered "known-unimplemented" cases, keyed by dot-path API (e.g. "utils.encode_col")
// -> Map<caseId, reason>. A divergence classifies UNSUPPORTED only when BOTH the api AND
// the specific caseId are registered — registering an api alone (a blanket "anything
// under this api is fine") is deliberately impossible with this shape. caseId must
// uniquely identify the exact input combination (option value + input type), not just
// describe the API in general, so an unrelated bug under the same api cannot hide behind
// an already-registered case. Adding an entry here is the ONLY way a divergence
// classifies as UNSUPPORTED — never a catch-all for "the output doesn't match and I
// don't know why."
export const UNSUPPORTED_ALLOWLIST = new Map([
  [
    'utils.format_cell',
    new Map([
      [
        'z="0.00" (numeric cell, non-General/non-m/d/yy format code)',
        'Phase 1B-1 deliberately implements only a narrow SSF number-format subset ' +
          "('General'/numFmtId 0 and 'm/d/yy'/numFmtId 14 — the only two formats " +
          'sheet_add_aoa actually needs, confirmed by reading the oracle source) rather ' +
          'than the ~900-line SSF_format/eval_fmt engine (the standalone "ssf" npm ' +
          'package, one of the 7 Apache-2.0 deps packages/xlsx deliberately does not ' +
          'take). Throws ELIXCEE_NUMFMT_UNSUPPORTED instead of guessing a rendering. ' +
          "See packages/xlsx/src/index.cjs's ssfFormat.",
      ],
    ]),
  ],
  [
    'utils.sheet_add_aoa',
    new Map([
      [
        'dateNF="yyyy-mm-dd" (Date value, custom format other than "m/d/yy")',
        "sheet_add_aoa's Date branch computes cell.w via ssfFormat immediately " +
          "(confirmed live: the real oracle renders \"2026-01-05\" for this exact " +
          'input) — unlike json_to_sheet/sheet_add_json, which only ever set cell.z ' +
          'and never call the format engine at all, so a custom dateNF is harmless ' +
          'there and needs no registration. sheet_add_aoa throws ' +
          'ELIXCEE_NUMFMT_UNSUPPORTED for any dateNF other than the literal "m/d/yy" ' +
          '(the one format this narrow SSF subset renders), for both cellDates:true ' +
          'and the default numeric-serial mode. See packages/xlsx/src/index.cjs\'s ' +
          'sheetAddAoa Date branch.',
      ],
    ]),
  ],
]);

// Registered intentional SECURITY divergences, keyed either by the elixcee-side error
// code that signals them (see docs/xlsx-security-model.md's planned ELIXCEE_* codes), or
// — for divergences that are a safer VALUE rather than a thrown error — by a descriptive
// string key a test passes explicitly via `securityDivergenceKey`.
export const SECURITY_DIVERGENCE_REGISTRY = new Map([
  [
    'book_append_sheet:proto_key_pollution',
    'A sheet named "__proto__" must be retained as data (docs/xlsx-security-model.md), ' +
      'never rejected, but the real oracle\'s `wb.Sheets[name] = ws` silently reassigns ' +
      'wb.Sheets\'s own prototype instead of storing a retrievable entry (confirmed: ' +
      'oracle\'s wb.Sheets ends up with zero own keys after this). Elixcee uses ' +
      'Object.defineProperty instead, so the sheet stays retrievable and no prototype is ' +
      'touched. See packages/xlsx/src/index.cjs\'s bookAppendSheet.',
  ],
  // 'ELIXCEE_ZIP_ENTRY_LIMIT' => 'zip bomb protection, see docs/xlsx-security-model.md',
]);

// Registered intentional SAFETY divergences (pure-function robustness against
// pathological arguments, no file parsing involved), keyed by the elixcee-side error
// code that signals them.
export const SAFETY_DIVERGENCE_REGISTRY = new Map([
  [
    'ELIXCEE_NON_FINITE_INDEX',
    'utils.encode_col(Infinity) hangs forever on the real oracle (Math.floor(Infinity) ' +
      'never reaches 0) — confirmed by running it to an OOM kill, not assumed. Elixcee ' +
      'rejects non-finite column/row indices instead. See packages/xlsx/src/index.cjs.',
  ],
]);

/**
 * @param {object} input
 * @param {string} [input.api] dot-path of the oracle API under test (e.g. "utils.encode_col")
 *   — required (together with unsupportedCaseId) for a divergence to ever resolve to
 *   UNSUPPORTED.
 * @param {string} [input.unsupportedCaseId] the exact case identifier (option value +
 *   input type, e.g. 'dateNF="yyyy-mm-dd" (Date value, custom format other than
 *   "m/d/yy")') to look up under UNSUPPORTED_ALLOWLIST.get(api). Both api AND this must
 *   match a registered entry — an api with no caseId (or a caseId not registered under
 *   that specific api) never resolves to UNSUPPORTED, by design.
 * @param {unknown} [input.oracleA] first oracle run (or the only oracle run)
 * @param {unknown} [input.oracleB] a second, independent oracle run on the same input —
 *   used to prove the comparison plumbing itself (oracle vs. itself must MATCH after
 *   normalization). Optional; omit when only comparing oracle vs. elixcee.
 * @param {unknown} [input.elixcee] elixcee's output for the same input
 * @param {string} [input.elixceeErrorCode] the ELIXCEE_* code elixcee's implementation
 *   threw, if any — checked against SECURITY_DIVERGENCE_REGISTRY and then
 *   SAFETY_DIVERGENCE_REGISTRY, in that order.
 * @param {string} [input.securityDivergenceKey] a descriptive registry key for
 *   divergences that are a safer RETURN VALUE rather than a thrown error (e.g. elixcee
 *   silently avoiding a prototype-pollution-shaped output the oracle produces) — checked
 *   against SECURITY_DIVERGENCE_REGISTRY only. Same anti-laundering rule: must be
 *   pre-registered with a reason, or this is a no-op.
 * @param {boolean} [input.oracleAmbiguous] set when the oracle's own behavior for this
 *   input is known to be inconsistent/underspecified (human-flagged, not inferred here)
 * @param {string} [input.oracleAmbiguityReason] required when oracleAmbiguous is true —
 *   the judgment call must be documented, not silent.
 * @returns {Verdict}
 */
export function classify({
  api,
  unsupportedCaseId,
  oracleA,
  oracleB,
  elixcee,
  elixceeErrorCode,
  securityDivergenceKey,
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
  if (elixceeErrorCode && SAFETY_DIVERGENCE_REGISTRY.has(elixceeErrorCode)) {
    return 'INTENTIONAL_SAFETY_DIVERGENCE';
  }
  if (securityDivergenceKey && SECURITY_DIVERGENCE_REGISTRY.has(securityDivergenceKey)) {
    return 'INTENTIONAL_SECURITY_DIVERGENCE';
  }

  if (oracleB !== undefined) {
    // Oracle-vs-itself comparison: proves the comparison plumbing, not elixcee.
    return deepEqual(oracleA, oracleB) ? 'MATCH' : 'NONDETERMINISTIC';
  }

  if (deepEqual(oracleA, elixcee)) return 'MATCH';

  if (api && unsupportedCaseId && UNSUPPORTED_ALLOWLIST.get(api)?.has(unsupportedCaseId)) {
    return 'UNSUPPORTED';
  }

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

  // A registered safety divergence classifies correctly even when oracleA/elixcee differ.
  assert.equal(
    classify({ oracleA: '', elixcee: undefined, elixceeErrorCode: 'ELIXCEE_NON_FINITE_INDEX' }),
    'INTENTIONAL_SAFETY_DIVERGENCE'
  );
  // An error code that ISN'T registered must still fall through to UNCLASSIFIED, not
  // silently pass — codes are not a free pass by existing, only by being registered.
  assert.equal(
    classify({ oracleA: '', elixcee: undefined, elixceeErrorCode: 'ELIXCEE_MADE_UP_CODE' }),
    'UNCLASSIFIED'
  );

  // A registered security divergence keyed by a descriptive string (not a thrown error
  // code) — for divergences that are a safer VALUE, not an exception.
  assert.equal(
    classify({
      oracleA: { own_keys: [] },
      elixcee: { own_keys: ['__proto__'] },
      securityDivergenceKey: 'book_append_sheet:proto_key_pollution',
    }),
    'INTENTIONAL_SECURITY_DIVERGENCE'
  );
  assert.equal(
    classify({ oracleA: {}, elixcee: { x: 1 }, securityDivergenceKey: 'not_registered' }),
    'UNCLASSIFIED'
  );

  // UNSUPPORTED only fires when BOTH api and the exact caseId are registered — an api
  // with no caseId, or a caseId that doesn't match, must never resolve to UNSUPPORTED.
  assert.equal(
    classify({ api: 'utils.format_cell', oracleA: '1234.50', elixcee: undefined }),
    'UNCLASSIFIED',
    'api alone (no unsupportedCaseId) must not resolve to UNSUPPORTED — blanket api allowlisting is disallowed'
  );
  assert.equal(
    classify({
      api: 'utils.format_cell',
      unsupportedCaseId: 'not a registered case',
      oracleA: '1234.50',
      elixcee: undefined,
    }),
    'UNCLASSIFIED',
    'an unregistered caseId under a registered api must still fail closed'
  );
  assert.equal(
    classify({
      api: 'utils.format_cell',
      unsupportedCaseId: 'z="0.00" (numeric cell, non-General/non-m/d/yy format code)',
      oracleA: '1234.50',
      elixcee: undefined,
    }),
    'UNSUPPORTED',
    'the exact registered (api, caseId) pair resolves to UNSUPPORTED'
  );

  assert.equal(
    UNSUPPORTED_ALLOWLIST.size,
    2,
    'Phase 1B-2A pre-work: exactly two apis have registered unsupported cases (format_cell, sheet_add_aoa)'
  );
  assert.equal(UNSUPPORTED_ALLOWLIST.get('utils.format_cell').size, 1, 'format_cell: exactly one registered case (narrow SSF subset)');
  assert.equal(UNSUPPORTED_ALLOWLIST.get('utils.sheet_add_aoa').size, 1, 'sheet_add_aoa: exactly one registered case (custom dateNF)');
  assert.equal(SECURITY_DIVERGENCE_REGISTRY.size, 1, 'Phase 1A: exactly one security divergence registered (book_append_sheet proto-key)');
  assert.equal(SAFETY_DIVERGENCE_REGISTRY.size, 1, 'Phase 1A: exactly one safety divergence registered (ELIXCEE_NON_FINITE_INDEX)');

  console.log('classify.mjs self-check: all assertions passed');
}
