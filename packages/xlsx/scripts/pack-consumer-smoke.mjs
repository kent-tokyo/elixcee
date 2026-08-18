// Proves the PUBLISHED TARBALL works standalone for a real consumer — a claim no other
// check in this repo makes. Every existing test (packages/xlsx/test/, compat/differential/,
// scripts/wasm-smoke.mjs) reaches this package through a relative path into
// packages/xlsx/src/, which exercises the source tree, never npm's own resolution of an
// installed package. scripts/audit-pack-contents.mjs is the closest existing check and it
// only inspects `npm pack --dry-run`'s FILE LIST; it never installs or executes anything.
//
// So this script does the whole real thing: `npm pack` a real .tgz, `npm install` it into a
// throwaway package under os.tmpdir() (deliberately outside this repo, so no parent
// node_modules can satisfy the import), and run every check from INSIDE that install.
//
// The single assertion that makes this honest rather than decorative is step 3: both
// `require.resolve("@elixcee/xlsx")` and `import.meta.resolve("@elixcee/xlsx")`, evaluated
// in the consumer, must land under the consumer's own node_modules/@elixcee/xlsx. Without
// it, a resolution that silently walked back into this repo would still "pass" every other
// step. The resolved paths are printed, not just asserted, so a CI log shows the evidence.
// `npm install <tarball>` (never `npm install <path>`, which npm symlinks back to the
// source directory) is the other half of that guarantee.
//
// Run: `node scripts/pack-consumer-smoke.mjs` from packages/xlsx/ (needs `npm ci` first —
// the TypeScript check below reuses this package's OWN typescript devDependency by absolute
// path rather than installing a second copy into the throwaway project, so the version
// under test is exactly the version `npm run typecheck` uses).
//
// Network: `npm install` of the tarball fetches its one runtime dependency (`ssf`) from the
// registry. This check is not offline-safe, unlike every other script here.
import { execFileSync } from 'node:child_process';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const DIR = path.dirname(fileURLToPath(import.meta.url));
const PKG_DIR = path.resolve(DIR, '..');
const REPO_ROOT = path.resolve(PKG_DIR, '..', '..');
const FIXTURE = path.join(REPO_ROOT, 'tests', 'fixtures', 'e2e', 'source.xlsx');
const TSC = path.join(PKG_DIR, 'node_modules', 'typescript', 'bin', 'tsc');

function step(name, fn) {
  console.log(`\n[pack-consumer] ${name}`);
  const out = fn();
  console.log(`[pack-consumer] ok: ${name}`);
  return out;
}

// Children print exactly one `__RESULT__ <json>` line; everything else they log is passed
// through for the CI log. Keeps the parent's assertions in one place instead of scattering
// process.exit calls across generated scripts.
function runInConsumer(dir, file, extraNodeArgs = []) {
  const stdout = execFileSync(process.execPath, [...extraNodeArgs, file], {
    cwd: dir,
    encoding: 'utf8',
    stdio: ['ignore', 'pipe', 'inherit'],
  });
  const line = stdout.split('\n').find((l) => l.startsWith('__RESULT__ '));
  process.stdout.write(
    stdout
      .split('\n')
      .filter((l) => l && !l.startsWith('__RESULT__ '))
      .map((l) => `  ${l}\n`)
      .join('')
  );
  if (!line) throw new Error(`${file} produced no __RESULT__ line`);
  return JSON.parse(line.slice('__RESULT__ '.length));
}

function assertUnder(label, resolved, root) {
  const real = fs.realpathSync(resolved.startsWith('file:') ? fileURLToPath(resolved) : resolved);
  const expectedRoot = fs.realpathSync(root);
  if (!real.startsWith(expectedRoot + path.sep)) {
    throw new Error(`${label} resolved OUTSIDE the throwaway install: ${real} (expected under ${expectedRoot})`);
  }
  if (!real.includes(path.join('node_modules', '@elixcee', 'xlsx'))) {
    throw new Error(`${label} did not resolve through node_modules/@elixcee/xlsx: ${real}`);
  }
  console.log(`  ${label} -> ${real}`);
  return real;
}

if (!fs.existsSync(FIXTURE)) {
  console.error(`missing fixture: ${FIXTURE}`);
  process.exit(1);
}
if (!fs.existsSync(TSC)) {
  console.error(`missing ${TSC} -- run \`npm ci\` in packages/xlsx first`);
  process.exit(1);
}

const tmpRoot = fs.mkdtempSync(path.join(os.tmpdir(), 'elixcee-pack-consumer-'));
const consumerDir = path.join(tmpRoot, 'consumer');
fs.mkdirSync(consumerDir);
console.log(`[pack-consumer] throwaway consumer: ${consumerDir}`);

const tarball = step('1. npm pack (real tarball, sizes recorded)', () => {
  const dryRun = JSON.parse(
    execFileSync('npm', ['pack', '--dry-run', '--json'], { cwd: PKG_DIR, encoding: 'utf8' })
  )[0];
  console.log(
    `  packed: ${dryRun.size} bytes (${(dryRun.size / 1024).toFixed(1)} KiB), ` +
      `unpacked: ${dryRun.unpackedSize} bytes (${(dryRun.unpackedSize / 1024).toFixed(1)} KiB), ` +
      `${dryRun.files.length} files`
  );
  // --pack-destination keeps the .tgz out of the repo working tree entirely.
  const name = execFileSync('npm', ['pack', '--pack-destination', tmpRoot], {
    cwd: PKG_DIR,
    encoding: 'utf8',
  })
    .trim()
    .split('\n')
    .pop();
  const tgz = path.join(tmpRoot, name);
  if (!fs.existsSync(tgz)) throw new Error(`npm pack reported ${name} but ${tgz} does not exist`);
  console.log(`  tarball: ${tgz} (${fs.statSync(tgz).size} bytes on disk)`);
  return tgz;
});

step('2. npm install the tarball into a throwaway package', () => {
  fs.writeFileSync(
    path.join(consumerDir, 'package.json'),
    JSON.stringify({ name: 'elixcee-pack-consumer', version: '0.0.0', private: true, type: 'commonjs' }, null, 2)
  );
  // The tarball path, never `npm install ../../packages/xlsx` — a directory spec makes npm
  // symlink straight back into this repo, which would defeat the entire point of this file.
  execFileSync('npm', ['install', tarball, '--no-audit', '--no-fund'], {
    cwd: consumerDir,
    stdio: 'inherit',
  });
  const installed = path.join(consumerDir, 'node_modules', '@elixcee', 'xlsx');
  if (!fs.existsSync(installed)) throw new Error(`not installed at ${installed}`);
  if (fs.lstatSync(installed).isSymbolicLink()) {
    throw new Error(`${installed} is a SYMLINK — npm linked the source dir instead of unpacking the tarball`);
  }
  // The fixture is data, not code, but copying it in keeps the consumer entirely
  // self-contained so nothing it runs reads from this repo at all.
  fs.copyFileSync(FIXTURE, path.join(consumerDir, 'fixture.xlsx'));
});

const CJS_PROBE = `
const path = require('node:path');
const fs = require('node:fs');
const resolved = require.resolve('@elixcee/xlsx');
const XLSX = require('@elixcee/xlsx');
const wb = XLSX.read(fs.readFileSync(path.join(__dirname, 'fixture.xlsx')));
const ws = wb.Sheets[wb.SheetNames[0]];
const rows = XLSX.sheet_to_json(ws, { header: 1 });
console.log('__RESULT__ ' + JSON.stringify({
  resolved,
  keys: Object.keys(XLSX).sort(),
  sheetNames: wb.SheetNames,
  ref: ws['!ref'],
  firstRows: rows.slice(0, 3),
}));
`;

const ESM_PROBE = `
import path from 'node:path';
import fs from 'node:fs';
import { fileURLToPath } from 'node:url';
import * as XLSX from '@elixcee/xlsx';
const here = path.dirname(fileURLToPath(import.meta.url));
const resolved = import.meta.resolve('@elixcee/xlsx');
const wb = XLSX.read(fs.readFileSync(path.join(here, 'fixture.xlsx')));
const ws = wb.Sheets[wb.SheetNames[0]];
// 'default' is Node's own synthesized CJS namespace default, and '__esModule' its interop
// marker — neither is a declared export of this package, so both are filtered out before
// comparing against the CJS require() key set.
const keys = Object.keys(XLSX).filter(k => k !== 'default' && k !== '__esModule').sort();
console.log('__RESULT__ ' + JSON.stringify({
  resolved,
  keys,
  sheetNames: wb.SheetNames,
  ref: ws['!ref'],
  firstRows: XLSX.sheet_to_json(ws, { header: 1 }).slice(0, 3),
}));
`;

const cjs = step('3. require("@elixcee/xlsx") from inside the install (CJS)', () => {
  fs.writeFileSync(path.join(consumerDir, 'probe.cjs'), CJS_PROBE);
  const r = runInConsumer(consumerDir, 'probe.cjs');
  assertUnder('require.resolve', r.resolved, consumerDir);
  if (!Array.isArray(r.sheetNames) || r.sheetNames.length === 0) {
    throw new Error(`read() returned no sheets: ${JSON.stringify(r.sheetNames)}`);
  }
  console.log(`  SheetNames: ${JSON.stringify(r.sheetNames)}  !ref: ${r.ref}`);
  console.log(`  first rows: ${JSON.stringify(r.firstRows)}`);
  return r;
});

const esm = step('4. import * as XLSX from "@elixcee/xlsx" from inside the install (ESM)', () => {
  fs.writeFileSync(path.join(consumerDir, 'probe.mjs'), ESM_PROBE);
  const r = runInConsumer(consumerDir, 'probe.mjs');
  assertUnder('import.meta.resolve', r.resolved, consumerDir);
  console.log(`  SheetNames: ${JSON.stringify(r.sheetNames)}  !ref: ${r.ref}`);
  return r;
});

step('5. CJS and ESM expose the identical export set, and read() the same data', () => {
  const a = JSON.stringify(cjs.keys);
  const b = JSON.stringify(esm.keys);
  if (a !== b) {
    const only = (x, y) => x.filter((k) => !y.includes(k));
    throw new Error(
      `export sets differ — CJS-only: ${JSON.stringify(only(cjs.keys, esm.keys))}, ` +
        `ESM-only: ${JSON.stringify(only(esm.keys, cjs.keys))}`
    );
  }
  console.log(`  ${cjs.keys.length} exports, identical in both entry points`);
  if (JSON.stringify(cjs.firstRows) !== JSON.stringify(esm.firstRows)) {
    throw new Error('CJS and ESM read() produced different data from the same fixture');
  }
});

step('6. "browser" export condition resolves from the INSTALLED tarball and runs', () => {
  // Same probe body as ESM, plus the resolution-target assertion below: "the browser
  // condition resolved" is only meaningful if it resolved to index.browser.mjs specifically,
  // rather than falling through to the "import" condition and quietly passing.
  fs.writeFileSync(path.join(consumerDir, 'probe-browser.mjs'), ESM_PROBE);
  const r = runInConsumer(consumerDir, 'probe-browser.mjs', ['--conditions=browser']);
  const resolved = assertUnder('import.meta.resolve (--conditions=browser)', r.resolved, consumerDir);
  if (!resolved.endsWith('index.browser.mjs')) {
    throw new Error(`browser condition resolved to ${resolved}, expected .../src/index.browser.mjs`);
  }
  if (JSON.stringify(r.sheetNames) !== JSON.stringify(cjs.sheetNames)) {
    throw new Error(
      `browser entry read() disagrees with CJS: ${JSON.stringify(r.sheetNames)} vs ${JSON.stringify(cjs.sheetNames)}`
    );
  }
  // index.mjs and index.browser.mjs are two hand-maintained re-export lists; comparing them
  // is nearly free here and is the only place a drift between them would be caught.
  if (JSON.stringify(r.keys) !== JSON.stringify(cjs.keys)) {
    throw new Error('browser entry export set differs from the CJS entry');
  }
  console.log(`  browser-entry SheetNames: ${JSON.stringify(r.sheetNames)}, export set identical`);
});

step('7. a TypeScript consumer snippet compiles against the installed types', () => {
  fs.writeFileSync(
    path.join(consumerDir, 'consumer.ts'),
    [
      `import * as XLSX from '@elixcee/xlsx';`,
      ``,
      `// Declared, not imported from node:fs — the tsconfig below sets "types": [] so that`,
      `// no ambient @types package can be picked up, which is exactly what makes this a`,
      `// check of the INSTALLED package's own declarations and nothing else.`,
      `declare const bytes: Uint8Array;`,
      ``,
      `const wb: XLSX.WorkBook = XLSX.read(bytes);`,
      `const ws: XLSX.WorkSheet = wb.Sheets[wb.SheetNames[0]];`,
      `const rows: unknown[] = XLSX.sheet_to_json(ws);`,
      `const csv: string = XLSX.sheet_to_csv(ws);`,
      `const addr: XLSX.CellAddress = XLSX.decode_cell('B2');`,
      `const ref: string = XLSX.encode_cell(addr);`,
      `const fresh: XLSX.WorkBook = XLSX.book_new();`,
      `XLSX.book_append_sheet(fresh, XLSX.aoa_to_sheet([[1, 'two', true]]), 'S1');`,
      `const hidden: 2 = XLSX.consts.SHEET_VERY_HIDDEN;`,
      `export { wb, rows, csv, ref, fresh, hidden };`,
    ].join('\n')
  );
  fs.writeFileSync(
    path.join(consumerDir, 'tsconfig.json'),
    JSON.stringify(
      {
        compilerOptions: {
          target: 'ES2020',
          // nodenext (not the legacy "node"): only a resolver that reads `exports` proves
          // the package's own exports["."].types entry works. Classic "node" resolution
          // would silently fall back to the top-level "types" field — a weaker claim.
          module: 'nodenext',
          moduleResolution: 'nodenext',
          strict: true,
          noEmit: true,
          esModuleInterop: true,
          skipLibCheck: true,
          // Nothing ambient: no @types package from any parent node_modules can be picked
          // up, so this compiles against the installed package's own declarations only.
          types: [],
        },
        include: ['consumer.ts'],
      },
      null,
      2
    )
  );
  console.log(`  tsc: ${TSC} (packages/xlsx's own typescript devDependency, by absolute path)`);
  execFileSync(process.execPath, [TSC, '--project', 'tsconfig.json'], {
    cwd: consumerDir,
    stdio: 'inherit',
  });
});

console.log('\n[pack-consumer] all checks passed against the real packed tarball.');
console.log(`[pack-consumer] leaving ${tmpRoot} in place for inspection (under os.tmpdir()).`);
