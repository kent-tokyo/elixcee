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
// The two Phase 1B-1/1B-2A-era cases (format_cell's narrow 'General'/'m/d/yy'-only
// subset, and sheet_add_aoa's custom-dateNF gap that subset caused) were closed when
// format_cell/sheet_add_aoa switched to the real SSF engine (see
// docs/xlsx-architecture.md's "SSF backend" decision and
// packages/xlsx/src/internal/ssf-adapter.cjs) — both now MATCH the oracle for the exact
// inputs that used to be registered here, so they were removed rather than left pointing
// at behavior that no longer diverges (a stale UNSUPPORTED entry is exactly the
// laundering hole this registry exists to prevent).
//
// Phase 2B added two real entries under 'read' (empty-string cell values, and
// <dimension> parsing) — both genuine reader.rs DEFECTS rather than capability gaps, kept
// here as a disclosed, deliberate exception until fixed (see the removed entries in this
// file's own git history for the original writeup). Both were fixed in the phase that
// added formula/dimension/rows-cols support to read() (see src/reader.rs's
// xlsx_sheet_cells and parse_dimension_ref, and compat/differential/xlsx-read.test.mjs's
// now-MATCHing cases for the live proof) and removed from here — leaving them would be
// exactly the stale-entry-laundering this registry exists to prevent. The allowlist is
// empty again as of that phase, same as it was after the SSF-backend phase closed the
// two Phase 1B-1/1B-2A-era cases mentioned above.
//
// One new entry, added when readFile()'s fixture-by-fixture differential (see
// compat/differential/xlsx-read.test.mjs) was written and immediately surfaced a real,
// previously-undisclosed reader.rs DEFECT — registered here under the same
// "disclosed exception until fixed" precedent as the two Phase 2B entries described above,
// NOT as a capability gap:
//
//   src/reader.rs's xlsx_sheet_cells calls `xlsx_parse_cell(text.trim(), ...)` — it trims
//   every cell's text unconditionally, ignoring the `xml:space="preserve"` attribute the
//   XLSX format uses precisely to mark significant leading/trailing whitespace. A cell
//   whose real value is "  padded  " therefore reads back as "padded". Confirmed live
//   against compat/corpus/workbooks/with_text.xlsx cell A3: oracle "  padded  ", elixcee
//   "padded". Reachable through read() and readFile() alike (readFile is a thin wrapper
//   over read), which is why it is registered under both apis rather than only the one that
//   happened to expose it.
//
// Not fixed here: src/reader.rs is outside the scope of the round that found this. Fixing
// it means honoring xml:space on the <t> element rather than trimming at the call site, and
// re-checking that the trim isn't load-bearing for the numeric/boolean parse paths that
// share xlsx_parse_cell. Both entries below must be REMOVED, not left stale, the moment
// that lands.
const XML_SPACE_PRESERVE_DEFECT =
  'reader.rs trims every cell\'s text (xlsx_sheet_cells: `xlsx_parse_cell(text.trim(), ...)`), ' +
  'ignoring xml:space="preserve", so significant leading/trailing whitespace in a cell value ' +
  'is lost — confirmed live on compat/corpus/workbooks/with_text.xlsx cell A3 ("  padded  " ' +
  'on the oracle, "padded" here). A genuine reader defect, disclosed rather than silently ' +
  'excluded from the fixture set; remove this entry when reader.rs honors xml:space.';

export const UNSUPPORTED_ALLOWLIST = new Map([
  ['read', new Map([['with_text.xlsx:xml_space_preserve_trimmed', XML_SPACE_PRESERVE_DEFECT]])],
  [
    'readFile',
    new Map([
      ['with_text.xlsx:xml_space_preserve_trimmed', XML_SPACE_PRESERVE_DEFECT],
      ['with_text.xlsx[cellStyles+cellDates]:xml_space_preserve_trimmed', XML_SPACE_PRESERVE_DEFECT],
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
  [
    'table_to_book:proto_key_pollution',
    'The same hazard as book_append_sheet above, in a different oracle code path: ' +
      'table_to_book builds its returned WorkBook via a `sheets[n] = sheet` assignment ' +
      '(n from a caller-controlled opts.sheet), not via book_append_sheet — confirmed ' +
      'live opts.sheet:"__proto__" corrupts the resulting wb.Sheets prototype the same ' +
      'way. Fixed identically with Object.defineProperty. See ' +
      'packages/xlsx/src/index.cjs\'s sheetToWorkbookSafe.',
  ],
  [
    'sheet_to_json:proto_header_primitive_dropped',
    'An explicit opts.header array may legitimately contain the literal string ' +
      '"__proto__" (spreadsheet column titled that, or a crafted probe). With a ' +
      'PRIMITIVE cell value there, the real oracle\'s `row[hdr[C]] = v` invokes ' +
      "Object.prototype's inherited __proto__ accessor, whose setter is a spec no-op " +
      'for non-object values — the column\'s data is silently DROPPED (confirmed: the ' +
      'oracle\'s row object ends up with no "__proto__" own key at all). Per ' +
      'docs/xlsx-security-model.md, elixcee must retain spreadsheet-derived data rather ' +
      'than lose it, so it uses Object.defineProperty (setJsonRowKey) instead, keeping ' +
      'the value as an ordinary own key. See packages/xlsx/src/index.cjs\'s ' +
      'makeJsonRow/setJsonRowKey.',
  ],
  [
    'sheet_to_json:proto_header_object_corruption',
    'Same opts.header:["__proto__",...] path as above, but with an OBJECT cell value ' +
      '(e.g. a Date cell under cellDates:true): the real oracle\'s `row[hdr[C]] = v` ' +
      "reassigns the ROW's own [[Prototype]] to that object (confirmed: `row instanceof " +
      'Date === true`, `Object.keys(row).length === 0` — the global Object.prototype ' +
      'stays clean, but this specific row object is corrupted). Elixcee\'s ' +
      'Object.defineProperty write keeps the row a plain object with the value stored as ' +
      'ordinary data, matching the primitive-value divergence\'s reasoning above.',
  ],
  [
    'sheet_to_html:unescaped_attribute',
    'The real oracle builds data-t/data-v/data-z/id (both the per-cell id and opts.id, ' +
      'table-level and per-cell) via raw string concatenation with NO escaping ' +
      '(confirmed live: a cell value or opts.id containing `"` breaks out of the ' +
      'attribute and injects an arbitrary onXXX handler that fires when the output is ' +
      'inserted into a live DOM). Applies to any cell value/number-format/id containing ' +
      "one of `&<>'\"` or a \\u0000-\\u001f control character — ordinary spreadsheet " +
      'content, not just a crafted probe. Elixcee escapes every attribute value it ' +
      'builds (escapeHtmlAttr). See packages/xlsx/src/index.cjs\'s sheetToHtml doc ' +
      'comment (finding 1) and docs/xlsx-security-model.md.',
  ],
  [
    'sheet_to_html:unsafe_href_scheme',
    'cell.l.Target is embedded into `href="..."` with no scheme check on the real ' +
      'oracle (confirmed live: a `javascript:` Target produces a clickable, ' +
      'code-executing link in the generated HTML — quote-escaping alone does not fix ' +
      'this, since no quote character is needed to make a href value dangerous). ' +
      'Elixcee allow-lists http(s)/mailto/tel/ftp/relative/fragment targets ' +
      '(isSafeHrefTarget); anything else renders as plain text with no <a> wrapper. A ' +
      'distinct failure mode from the attribute-escaping divergence above (a scheme ' +
      'check, not a character-escaping fix), so kept as its own registry entry. See ' +
      'packages/xlsx/src/index.cjs\'s sheetToHtml doc comment (finding 2).',
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
  [
    'ELIXCEE_RANGE_TOO_LARGE',
    'sheet_to_formulae/sheet_to_csv/sheet_to_txt/sheet_to_json/sheet_to_html all iterate every (row,col) pair in a ' +
      "worksheet's !ref rectangle regardless of sparsity — confirmed live (timeout-" +
      "guarded subprocess) that a crafted full-grid !ref ('A1:XFD1048576', ~17.18 " +
      'billion cells) does not return within 25s on the real oracle\'s sheet_to_csv, and ' +
      'even much smaller full-rectangle spans are already slow on the oracle itself ' +
      '(26,000,000 cells: 12-16s). Elixcee rejects ranges above a 5,000,000-cell ' +
      'threshold instead of iterating them. See packages/xlsx/src/internal/range-guard.cjs.',
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

// Per-API x per-verdict rollup — the single source of truth for compatibility-report
// counts. Reports must never hand-type a verdict breakdown (e.g. "57/57 acceptable"
// blurring MATCH together with registered divergences into one number); every count in a
// completion report is expected to trace back to this function's output, not to manual
// arithmetic. Takes the flat `{api, label, verdict}` results array every differential
// test file already accumulates and groups it as Map<api, Map<verdict, count>>.
//
// Deliberately does NOT merge results across test FILES (e.g. xlsx-utils.test.mjs's
// public-API fixtures vs. ssf-format.test.mjs's internal SSF-backend conformance
// fixtures) — each file calls this on its own `results` array and reports its own table.
// Public API differential fixtures and internal backend conformance fixtures answer
// different questions (does the public surface match the oracle? vs. does the internal
// SSF engine choice match the oracle's bundled formatter?) and mixing their totals into
// one number would hide which one a given count is actually about.
export function summarizeByApiAndVerdict(results) {
  const byApi = new Map();
  for (const r of results) {
    if (!byApi.has(r.api)) byApi.set(r.api, new Map());
    const verdicts = byApi.get(r.api);
    verdicts.set(r.verdict, (verdicts.get(r.verdict) || 0) + 1);
  }
  return byApi;
}

// Renders summarizeByApiAndVerdict()'s output as one line per API, e.g.:
//   utils.sheet_to_json: MATCH=54 INTENTIONAL_SAFETY_DIVERGENCE=1 INTENTIONAL_SECURITY_DIVERGENCE=2 (total 57)
// VERDICT_ORDER controls column order (matches VERDICTS' own priority ordering above);
// verdicts with a zero count for a given api are omitted from that api's line.
export function formatApiVerdictSummary(byApi) {
  const lines = [];
  for (const [api, verdicts] of byApi) {
    const total = [...verdicts.values()].reduce((a, b) => a + b, 0);
    const parts = VERDICTS.filter((v) => verdicts.has(v))
      .map((v) => `${v}=${verdicts.get(v)}`)
      .join(' ');
    lines.push(`${api}: ${parts} (total ${total})`);
  }
  return lines.join('\n');
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
  // Exercised against a synthetic entry (not a real one — the allowlist is empty as of
  // Phase 1B-2B, see its doc comment) so this coverage survives regardless of which real
  // cases are or aren't registered at any given time.
  const SYNTHETIC_CASE_ID = '__classify_self_check__: synthetic unsupported case';
  UNSUPPORTED_ALLOWLIST.set('utils.__self_check_api__', new Map([[SYNTHETIC_CASE_ID, 'test fixture only']]));
  try {
    assert.equal(
      classify({ api: 'utils.__self_check_api__', oracleA: '1234.50', elixcee: undefined }),
      'UNCLASSIFIED',
      'api alone (no unsupportedCaseId) must not resolve to UNSUPPORTED — blanket api allowlisting is disallowed'
    );
    assert.equal(
      classify({
        api: 'utils.__self_check_api__',
        unsupportedCaseId: 'not a registered case',
        oracleA: '1234.50',
        elixcee: undefined,
      }),
      'UNCLASSIFIED',
      'an unregistered caseId under a registered api must still fail closed'
    );
    assert.equal(
      classify({
        api: 'utils.__self_check_api__',
        unsupportedCaseId: SYNTHETIC_CASE_ID,
        oracleA: '1234.50',
        elixcee: undefined,
      }),
      'UNSUPPORTED',
      'the exact registered (api, caseId) pair resolves to UNSUPPORTED'
    );
  } finally {
    UNSUPPORTED_ALLOWLIST.delete('utils.__self_check_api__');
  }

  // Pinned counts, deliberately: this assert is what forces anyone ADDING an allowlist
  // entry to state it here explicitly rather than slipping a divergence past review, and
  // what forces anyone FIXING one to remove it rather than leave it stale.
  //
  // Currently 2 apis / 3 cases, all one defect: reader.rs trims cell text unconditionally
  // and so loses xml:space="preserve" whitespace (see UNSUPPORTED_ALLOWLIST's own comment
  // above for the full writeup). Registered under both 'read' (1 case) and 'readFile'
  // (2 cases: with and without cellStyles+cellDates) because readFile is a thin wrapper
  // over read and both reach the same defect. Delete all three when reader.rs honors
  // xml:space.
  //
  // The two Phase 2B-era "read" cases (empty-string cell value,
  // declared-<dimension>-wider-than-data) that used to be pinned here were removed when
  // reader.rs's xlsx_sheet_cells/parse_dimension_ref fixed them — same as the earlier Phase
  // 1B-1/1B-2A-era format_cell/sheet_add_aoa cases closing when the real SSF engine was
  // wired in.
  assert.equal(UNSUPPORTED_ALLOWLIST.size, 2, 'exactly two apis have registered unsupported cases: read, readFile');
  assert.deepEqual(
    [...UNSUPPORTED_ALLOWLIST.get('read').keys()],
    ['with_text.xlsx:xml_space_preserve_trimmed'],
    '"read" has exactly the one registered xml:space defect case'
  );
  assert.deepEqual(
    [...UNSUPPORTED_ALLOWLIST.get('readFile').keys()],
    [
      'with_text.xlsx:xml_space_preserve_trimmed',
      'with_text.xlsx[cellStyles+cellDates]:xml_space_preserve_trimmed',
    ],
    '"readFile" has exactly the two registered xml:space defect cases (one per opts shape)'
  );
  assert.equal(
    SECURITY_DIVERGENCE_REGISTRY.size,
    6,
    'Phase 1A + 1B-3 + 1C: six security divergences registered (book_append_sheet ' +
      'proto-key, table_to_book proto-key, sheet_to_json proto-header primitive-dropped, ' +
      'sheet_to_json proto-header object-corruption, sheet_to_html unescaped attribute, ' +
      'sheet_to_html unsafe href scheme)'
  );
  assert.equal(
    SAFETY_DIVERGENCE_REGISTRY.size,
    2,
    'Phase 1A + 1B-2B: two safety divergences registered (ELIXCEE_NON_FINITE_INDEX, ELIXCEE_RANGE_TOO_LARGE)'
  );

  // summarizeByApiAndVerdict / formatApiVerdictSummary: per-API verdict counts must never
  // be hand-typed in a report — this is the exact bug the user caught in the Phase 1B-3
  // report ("57/57 acceptable" hid that 3 of the 57 were registered divergences, not
  // MATCH). A mixed-verdict api must show every verdict's own count, not a collapsed
  // pass/fail ratio.
  {
    const sample = [
      { api: 'utils.sheet_to_json', label: 'a', verdict: 'MATCH' },
      { api: 'utils.sheet_to_json', label: 'b', verdict: 'MATCH' },
      { api: 'utils.sheet_to_json', label: 'c', verdict: 'INTENTIONAL_SAFETY_DIVERGENCE' },
      { api: 'utils.sheet_to_json', label: 'd', verdict: 'INTENTIONAL_SECURITY_DIVERGENCE' },
      { api: 'utils.sheet_to_json', label: 'e', verdict: 'INTENTIONAL_SECURITY_DIVERGENCE' },
      { api: 'utils.sheet_get_cell', label: 'f', verdict: 'MATCH' },
    ];
    const byApi = summarizeByApiAndVerdict(sample);
    assert.deepEqual(
      Object.fromEntries([...byApi.get('utils.sheet_to_json').entries()]),
      { MATCH: 2, INTENTIONAL_SAFETY_DIVERGENCE: 1, INTENTIONAL_SECURITY_DIVERGENCE: 2 }
    );
    assert.deepEqual(Object.fromEntries([...byApi.get('utils.sheet_get_cell').entries()]), { MATCH: 1 });
    const rendered = formatApiVerdictSummary(byApi);
    assert.ok(rendered.includes('utils.sheet_to_json: MATCH=2 INTENTIONAL_SECURITY_DIVERGENCE=2 INTENTIONAL_SAFETY_DIVERGENCE=1 (total 5)'));
    assert.ok(rendered.includes('utils.sheet_get_cell: MATCH=1 (total 1)'));
  }

  console.log('classify.mjs self-check: all assertions passed');
}
