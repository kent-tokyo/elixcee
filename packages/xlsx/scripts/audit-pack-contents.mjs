// Asserts what `npm pack` would actually publish, using `npm pack --dry-run --json`'s own
// machine-readable file list rather than trusting the `files` field in package.json to mean
// what it says — this is a real dry-run of npm's own packing logic (which also honors
// .npmignore/.gitignore and npm's built-in always-included files like LICENSE/README.md),
// not a re-implementation of npm's own inclusion rules.
//
// Written because this exact check didn't exist anywhere before (confirmed by a dedicated
// investigation: a manual dry-run was clean, but nothing asserted it in CI) — a future
// change to `files`/`exports`/build output could silently drop something required (a
// license file, a public entry point) or start shipping something that shouldn't be public
// (test files, source maps pointing at internal build paths, a stray node_modules/ entry)
// with no CI signal either way.
//
// Run: `node scripts/audit-pack-contents.mjs` from packages/xlsx/.
import { execFileSync } from 'node:child_process';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const DIR = path.dirname(fileURLToPath(import.meta.url));
const PKG_DIR = path.resolve(DIR, '..');

// Every one of these must appear in the tarball, by exact path. Sourced from
// package.json's own main/module/types/exports/files fields plus the two
// attribution files npm's own docs require for a licensed package — not an
// arbitrary list, each entry traces back to something package.json already
// promises.
const REQUIRED_FILES = [
  'LICENSE',
  'README.md',
  'THIRD_PARTY_NOTICES.md',
  'package.json',
  'src/index.cjs',        // "main"
  'src/index.mjs',        // "module" / exports["."].import
  'src/index.d.ts',       // "types" / exports["."].types
  'src/index.browser.mjs', // exports["."].browser
];

// Path prefixes that must never appear in the tarball, regardless of what
// files/.npmignore currently say — a regression in either could let one of
// these leak into a real publish.
const FORBIDDEN_PREFIXES = [
  'node_modules/',
  'test/',
  '.git',
  'tsconfig',       // tsconfig.json, tsconfig.no-dom.json — build-time only
  'scripts/',        // this script itself and any siblings — dev-only tooling
];

function main() {
  const stdout = execFileSync('npm', ['pack', '--dry-run', '--json'], {
    cwd: PKG_DIR,
    encoding: 'utf8',
  });
  const [pack] = JSON.parse(stdout);
  const paths = pack.files.map(f => f.path);
  const pathSet = new Set(paths);

  const problems = [];

  for (const required of REQUIRED_FILES) {
    if (!pathSet.has(required)) {
      problems.push(`missing required file: ${required}`);
    }
  }

  for (const p of paths) {
    for (const forbidden of FORBIDDEN_PREFIXES) {
      if (p === forbidden || p.startsWith(forbidden)) {
        problems.push(`forbidden path present: ${p} (matches "${forbidden}")`);
      }
    }
  }

  // Every internal-runtime file (src/internal/**) must be a real, expected
  // asset (a .cjs adapter module or the vendored WASM artifacts) -- not a
  // hard allowlist of exact filenames (those legitimately change as the
  // bridge evolves), but a shape check: anything under src/internal/ that
  // isn't .cjs or under src/internal/wasm/ is unexpected.
  for (const p of paths) {
    if (!p.startsWith('src/internal/')) continue;
    const isAdapterModule = p.endsWith('.cjs') && !p.startsWith('src/internal/wasm/');
    const isWasmArtifact = p.startsWith('src/internal/wasm/');
    if (!isAdapterModule && !isWasmArtifact) {
      problems.push(`unexpected file under src/internal/: ${p}`);
    }
  }

  console.log(`npm pack --dry-run: ${paths.length} files, ${pack.size} bytes packed, ${pack.unpackedSize} bytes unpacked`);
  for (const p of paths) console.log(`  ${p}`);

  if (problems.length > 0) {
    console.log(`\n${problems.length} problem(s):`);
    for (const p of problems) console.log(`  - ${p}`);
    process.exitCode = 1;
    return;
  }

  console.log('\naudit-pack-contents: all required files present, nothing forbidden, nothing unexpected under src/internal/.');
}

main();
