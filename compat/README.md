# compat/

A Node.js project living inside an otherwise-Rust repository. Its whole job is verifying
`@elixcee/xlsx` and elixcee's own VBA semantics against real, independent ground truth —
never guessing, never asserting elixcee's own output as if it were the spec. Four
independent suites, each answering a different question:

```sh
cd compat
npm install
```

## `differential/` — does `@elixcee/xlsx` match the real `xlsx@0.18.5` package?

Installs the real `xlsx@0.18.5` npm package (the "oracle") as a devDependency and runs it
side by side with `@elixcee/xlsx` (imported only via a relative path into
`../packages/xlsx/src`, never as an npm dependency — see `docs/xlsx-architecture.md`'s
"Non-negotiable" section). `differential/classify.mjs` defines the six-value verdict every
comparison resolves to: `MATCH` / `INTENTIONAL_SECURITY_DIVERGENCE` /
`INTENTIONAL_SAFETY_DIVERGENCE` / `UNSUPPORTED` / `BUG` / `NONDETERMINISTIC` — see its own
module doc comment for definitions and the comparison-normalization rule (compare parsed
logical shape, never raw bytes — the oracle embeds a timestamp in `docProps/core.xml`).

```sh
npm run classifier:self-check   # classify.mjs / normalize.mjs self-checks
npm run differential:utils      # all 33 utils.* exports — 512 MATCH + 14 disclosed
npm run differential:ssf-format # SSF number-format conformance — 1831/1831
npm run differential:read       # XLSX.read()/readFile()/readFileSync() — 30 MATCH + 3 disclosed
npm run differential:metadata   # CJS/ESM export identity, key order — 36/36
```

`UNSUPPORTED`/`INTENTIONAL_*` verdicts are pinned to an explicit allowlist in
`classify.mjs` (count and reason both asserted by `classifier:self-check`) — a divergence
can't silently appear or silently get fixed without the allowlist itself changing, which
forces it to be a reviewed, visible diff.

## `corpus/` — does elixcee run 581 real-world-shaped VBA scenarios the way it's supposed to?

A different axis from `differential/`: not "matches SheetJS" but "matches elixcee's own
documented pass/fail expectations," scenario ID by scenario ID.

```sh
npm run corpus:run       # run all 581 scenarios against the elixcee CLI binary
npm run corpus:outcomes  # classify against corpus/expected-outcomes.json
```

Verdicts: `PASS` / `EXPECTED_RUNTIME_ERROR` / `EXPECTED_UNSUPPORTED` / `NONDETERMINISTIC` /
`MISMATCH` / `UNEXPLAINED`. The last two must both be 0 — an `UNEXPLAINED` result (a
crash/timeout/hang, or a pass/fail that doesn't match what's registered) is never silently
tolerated. Requires `cargo build --release` first (drives the real CLI binary, not a mock).

`corpus/run-libreoffice.mjs` exists as a second, LibreOffice-based oracle for this same
scenario set, but is currently unusable for most scenarios (578/581 are
`ORACLE_UNAVAILABLE` — headless UNO hangs on `Range`/`Cells` access) — see `ROADMAP.md`'s
"Known gaps" for why this hasn't been fixed.

## `vba-semantics/` — is the VALUE elixcee produces the one real VBA semantics says it should be?

The suite `differential/` and `corpus/` both can't answer: a function that runs without
error and returns a plausible-but-wrong value is invisible to either of them. Needs no
oracle at all — expected values are computed from `reference/*.mjs`, small,
independently-checkable pure-JS reference implementations of *documented* real VBA
semantics (sourced from Microsoft's own VBA language reference,
learn.microsoft.com/en-us/office/vba/language/reference/, fetched live when adding a case
— never recalled from memory and never elixcee's own output laundered into looking like
the spec).

```sh
npm run semantics:generate  # only needed after editing generate-cases.mjs itself
cd .. && cargo build --release --bin elixcee && cd compat
npm run semantics:run
npm run semantics:report
```

386 cases, 0 `BUG`, 0 `UNCLASSIFIED`, 19 `KNOWN_LIMITATION` (disclosed, non-gating —
see `vba-semantics/README.md`'s "Current state" for the full, root-cause-grouped
breakdown). Deterministic: two consecutive runs produce a byte-identical `report.json`.

## `oracle/` — what does xlsx@0.18.5 actually expose?

```sh
npm run oracle:manifest
```

Regenerates `oracle/api-manifest.json`: a machine-derived record of the oracle's public
API surface (top-level exports, `.utils`, `.stream`, for both its CJS and ESM
entrypoints). Commit the regenerated file whenever the pinned `xlsx` version changes.

`differential:demo` (`node differential/run-demo.mjs`) is the original Phase 0 plumbing
check (oracle-vs-oracle → `MATCH`, oracle-vs-placeholder → `UNSUPPORTED`) — superseded by
the real suites above for actual coverage, kept only as a minimal sanity check of the
classification machinery itself.

## `oracle-excel-com/` — a written, unimplemented Excel COM adapter contract

**Interface definition only — nothing here has ever run against real Microsoft Excel.**
See `oracle-excel-com/CONTRACT.md` and `UNVERIFIED.md` for exactly what a real Windows/COM
environment would need to confirm before any of this could be treated as verified. This is
the single largest gap blocking a claim of Excel-validated compatibility anywhere in this
project — see `ROADMAP.md`'s "Known gaps" #1.

## CI

All four `differential/` suites and the classifier self-checks run in
`.github/workflows/ci.yml`'s `node-js` job on every push (Node 20/22 matrix). `corpus` and
`vba-semantics` run in that same workflow's `compat-vba` job (single Node 22, since both
need a release build of the elixcee CLI binary first, which `node-js` doesn't do) — run
them locally after `cargo build --release` whenever touching `src/parser/`/`src/vm/` for
faster iteration than waiting on CI.
