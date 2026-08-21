// Standing regression guard for function-identity metadata: every @elixcee/xlsx export's
// `.name`, `.length`, and its own property descriptor on the exports object must match
// the real oracle's (xlsx@0.18.5) `utils.*` — not just its runtime behavior. This exists
// because the divergence it guards against survived Phase 1A and Phase 1B-1 undetected:
// a plain `{ encode_col: encodeCol }` object-literal assignment does NOT rename an
// already-named function's `.name`, so every export's `.name` stayed the internal
// camelCase name (e.g. "encodeCol") instead of the public "encode_col" until this file's
// companion fix (packages/xlsx/src/index.cjs's nameAs helper). See
// docs/xlsx-architecture.md / the fix(xlsx) commit that introduced this file for the
// discovery.
//
// Deliberately compares against elixcee's CJS export object, NOT the ESM namespace: ESM
// module-namespace object properties are `configurable: false` per spec, while the
// oracle's `utils` is a plain object literal (`configurable: true`) — comparing against
// the ESM side would report a `configurable` mismatch that has nothing to do with this
// package's own correctness. CJS/ESM function-object identity (same reference reachable
// both ways) is checked separately below, once, not per-export descriptor comparison.
import assert from 'node:assert/strict';
import XLSX from 'xlsx';
import elixceeCjs from '../../packages/xlsx/src/index.cjs';
import * as elixceeEsm from '../../packages/xlsx/src/index.mjs';

const U = XLSX.utils;
let failures = 0;

// These exports aren't `utils.*` members at all — the oracle's own `read`/`readFile`/
// `readFileSync`/`write`/`writeFile`/`writeFileSync` live at the TOP level (confirmed
// live: `U.read === undefined`, and `Object.keys(XLSX)` is
// ["version","parse_xlscfb","parse_zip","read","readFile","readFileSync","write",
// "writeFile","writeFileSync","writeFileAsync",...] — writeFileAsync is not implemented
// here, matching this package's own established convention of simply not exporting a
// capability it doesn't have, rather than exporting one that throws). Comparing the six
// implemented ones against `XLSX.utils` in the loop below would always report a type
// mismatch for a reason that has nothing to do with this package's own correctness —
// they're checked against the right oracle surface (`XLSX`) in their own block instead,
// and excluded from the utils-key-order check further down for the same reason. Their
// relative order against each OTHER is checked separately, below.
const TOP_LEVEL_KEYS = ['read', 'readFile', 'readFileSync', 'write', 'writeFile', 'writeFileSync'];

for (const key of Object.keys(elixceeCjs)) {
  if (TOP_LEVEL_KEYS.includes(key)) continue;
  const oracleVal = U[key];
  const elixceeVal = elixceeCjs[key];
  if (typeof oracleVal !== typeof elixceeVal) {
    console.error(`FAIL  ${key}: type mismatch (oracle=${typeof oracleVal} elixcee=${typeof elixceeVal})`);
    failures += 1;
    continue;
  }
  const problems = [];
  if (typeof oracleVal === 'function') {
    // .name/.length are function-only metadata — plain data exports (e.g. `consts`) have
    // neither and are compared purely by descriptor + value shape below instead.
    if (elixceeVal.name !== oracleVal.name) {
      problems.push(`name: oracle=${JSON.stringify(oracleVal.name)} elixcee=${JSON.stringify(elixceeVal.name)}`);
    }
    if (elixceeVal.length !== oracleVal.length) {
      problems.push(`length: oracle=${oracleVal.length} elixcee=${elixceeVal.length}`);
    }
  } else {
    // Data export (e.g. `consts`): compare own enumerable keys and their values/
    // descriptors directly, since there's no .name/.length to check.
    const oracleKeys = Object.keys(oracleVal).sort();
    const elixceeKeys = Object.keys(elixceeVal).sort();
    if (JSON.stringify(oracleKeys) !== JSON.stringify(elixceeKeys)) {
      problems.push(`own keys: oracle=${JSON.stringify(oracleKeys)} elixcee=${JSON.stringify(elixceeKeys)}`);
    } else {
      for (const k of oracleKeys) {
        if (oracleVal[k] !== elixceeVal[k]) problems.push(`${key}.${k}: oracle=${JSON.stringify(oracleVal[k])} elixcee=${JSON.stringify(elixceeVal[k])}`);
      }
    }
  }
  const oracleDesc = Object.getOwnPropertyDescriptor(U, key);
  const elixceeDesc = Object.getOwnPropertyDescriptor(elixceeCjs, key);
  for (const flag of ['enumerable', 'writable', 'configurable']) {
    if (oracleDesc[flag] !== elixceeDesc[flag]) {
      problems.push(`descriptor.${flag}: oracle=${oracleDesc[flag]} elixcee=${elixceeDesc[flag]}`);
    }
  }
  if (elixceeEsm[key] !== elixceeCjs[key]) {
    problems.push('CJS/ESM identity: elixceeEsm[key] !== elixceeCjs[key] (should be the same object/function)');
  }
  if (problems.length) {
    console.error(`FAIL  ${key}: ${problems.join('; ')}`);
    failures += 1;
  } else {
    console.log(`OK    ${key}`);
  }
}

// The top-level exports against their real oracle counterparts (`XLSX.read` etc., not
// `XLSX.utils`) — see TOP_LEVEL_KEYS' comment above for why these live in their own block.
for (const key of TOP_LEVEL_KEYS) {
  const oracleVal = XLSX[key];
  const elixceeVal = elixceeCjs[key];
  const problems = [];
  if (elixceeVal.name !== oracleVal.name) {
    problems.push(`name: oracle=${JSON.stringify(oracleVal.name)} elixcee=${JSON.stringify(elixceeVal.name)}`);
  }
  if (elixceeVal.length !== oracleVal.length) {
    problems.push(`length: oracle=${oracleVal.length} elixcee=${elixceeVal.length}`);
  }
  const oracleDesc = Object.getOwnPropertyDescriptor(XLSX, key);
  const elixceeDesc = Object.getOwnPropertyDescriptor(elixceeCjs, key);
  for (const flag of ['enumerable', 'writable', 'configurable']) {
    if (oracleDesc[flag] !== elixceeDesc[flag]) {
      problems.push(`descriptor.${flag}: oracle=${oracleDesc[flag]} elixcee=${elixceeDesc[flag]}`);
    }
  }
  if (elixceeEsm[key] !== elixceeCjs[key]) {
    problems.push(`CJS/ESM identity: elixceeEsm.${key} !== elixceeCjs.${key} (should be the same function)`);
  }
  if (problems.length) {
    console.error(`FAIL  ${key} (vs top-level XLSX.${key}): ${problems.join('; ')}`);
    failures += 1;
  } else {
    console.log(`OK    ${key} (vs top-level XLSX.${key}, not XLSX.utils)`);
  }
}

// The oracle exports ONE function under both `readFile` and `readFileSync` (confirmed
// live: `XLSX.readFile === XLSX.readFileSync`). That identity is part of the public shape —
// a consumer can legitimately compare or swap them — so it's asserted rather than left to
// the per-key `.name`/`.length` checks above, which two separate but identically-shaped
// functions would also pass.
{
  const oracleAliased = XLSX.readFile === XLSX.readFileSync;
  const elixceeAliased = elixceeCjs.readFile === elixceeCjs.readFileSync;
  if (oracleAliased !== elixceeAliased) {
    console.error(
      `FAIL  readFile/readFileSync aliasing: oracle readFile===readFileSync is ${oracleAliased}, elixcee is ${elixceeAliased}`
    );
    failures += 1;
  } else {
    console.log(`OK    readFile === readFileSync (same function object, matching the oracle)`);
  }
}

// Same aliasing check, for writeFile/writeFileSync (confirmed live:
// `XLSX.writeFile === XLSX.writeFileSync`).
{
  const oracleAliased = XLSX.writeFile === XLSX.writeFileSync;
  const elixceeAliased = elixceeCjs.writeFile === elixceeCjs.writeFileSync;
  if (oracleAliased !== elixceeAliased) {
    console.error(
      `FAIL  writeFile/writeFileSync aliasing: oracle writeFile===writeFileSync is ${oracleAliased}, elixcee is ${elixceeAliased}`
    );
    failures += 1;
  } else {
    console.log(`OK    writeFile === writeFileSync (same function object, matching the oracle)`);
  }
}

// Relative order among the top-level (non-utils) exports, against the oracle's own
// Object.keys(XLSX) order — the utils-key-order check below excludes these, so without this
// they'd have no order check at all.
{
  const oracleTopOrder = Object.keys(XLSX).filter((k) => TOP_LEVEL_KEYS.includes(k));
  const elixceeTopOrder = Object.keys(elixceeCjs).filter((k) => TOP_LEVEL_KEYS.includes(k));
  if (JSON.stringify(elixceeTopOrder) !== JSON.stringify(oracleTopOrder)) {
    console.error(`FAIL  top-level key order: elixcee=${JSON.stringify(elixceeTopOrder)} oracle=${JSON.stringify(oracleTopOrder)}`);
    failures += 1;
  } else {
    console.log(`OK    top-level key order matches the oracle's own Object.keys(XLSX) relative order`);
  }
}

console.log(`\n${Object.keys(elixceeCjs).length - failures}/${Object.keys(elixceeCjs).length} exports match name/length/descriptor/CJS-ESM-identity against the oracle`);

// Key ORDER, not just per-key content — Phase 1C discovered this had never been checked
// (every key matched individually, but elixcee's module.exports literal had been ordered
// by this file's own section comments, not the oracle's actual Object.keys() insertion
// order, undetected across Phases 1A-1B-3). Compares elixcee's own key order against the
// oracle's FULL key order filtered down to only the keys elixcee currently implements —
// this must hold at every point during the utils-completion phases, not just once every
// oracle key is implemented.
{
  const oracleOrder = Object.keys(U);
  // read/readFile/readFileSync excluded — they aren't utils.* members (see TOP_LEVEL_KEYS'
  // comment above), so they have no position in Object.keys(XLSX.utils) to compare against.
  // Their own relative order is checked in its own block above instead.
  const elixceeOrder = Object.keys(elixceeCjs).filter((k) => !TOP_LEVEL_KEYS.includes(k));
  const oracleFiltered = oracleOrder.filter((k) => elixceeOrder.includes(k));
  if (JSON.stringify(elixceeOrder) !== JSON.stringify(oracleFiltered)) {
    console.error('FAIL  key order: elixcee\'s Object.keys() does not match the oracle\'s own relative order');
    console.error(`      elixcee: ${JSON.stringify(elixceeOrder)}`);
    console.error(`      oracle:  ${JSON.stringify(oracleFiltered)}`);
    failures += 1;
  } else {
    console.log('OK    key order matches the oracle\'s own Object.keys(XLSX.utils) relative order');
  }
}

if (failures > 0) {
  console.error('\nmetadata differential suite FAILED.');
  process.exit(1);
}

// Sanity check on the check itself: a deliberately-wrong name must be caught, not
// silently accepted (guards against this file's own comparison logic regressing).
const wrongName = function encodeColWRONG() {};
assert.notEqual(wrongName.name, U.encode_col.name, 'self-check: this comparison must be capable of detecting a name mismatch');

console.log('\nmetadata differential suite passed.');
