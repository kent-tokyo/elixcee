# compat/corpus/ — VBA-macro oracle-agnostic corpus and runners

Builds and runs a corpus of VBA-macro scenarios against elixcee and against
**LibreOffice** (a secondary reference implementation, explicitly **not** an "Excel
oracle" — see the framing note below) using a schema, runner interface, and
classification vocabulary designed so a future **real-Excel** backend can be plugged in
without redesigning anything. See `../oracle-excel-com/` for that (currently
unimplemented, contract-only) adapter.

## Framing note — read before quoting any number from this directory

LibreOffice's VBA support is its own independently-implemented compatibility layer, not
Microsoft Excel. Every result file in `results/` produced by `run-libreoffice.mjs` is
tagged `"oracle": "libreoffice"`. Nothing in this directory should be read, quoted, or
summarized as "Excel COM validation," "verified against Excel," or "effectively
Excel-equivalent" — that validation has not happened (see `../oracle-excel-com/
UNVERIFIED.md` for the complete, itemized list of what a real Excel run would still need
to confirm) and this project has a standing discipline against claiming verification
against a source that wasn't actually used.

## Layout

- `SCHEMA.md` — the scenario JSON schema (oracle-agnostic; no backend-specific fields).
- `generate-scenarios.mjs` → `scenarios.json` — ~580 scenarios, generated from templates
  (not hand-typed one at a time — see the file's own doc comment for why).
- `workbooks/generate-workbooks.mjs` → `workbooks/*.xlsx` — five small named base
  workbooks scenarios reference by name, using the `xlsx` npm package already a
  devDependency in `compat/package.json` (no new dependency).
- `run-elixcee.mjs` — drives `scenarios.json` against the real `elixcee` CLI binary
  (built from this repo's own Rust source; not modified) → `results/elixcee-results.json`.
- `run-libreoffice.mjs` — drives `scenarios.json` against LibreOffice via
  `soffice --headless` CLI macro invocation → `results/libreoffice-results*.json`.
- `normalize.mjs` — canonical cell-array shape + equality, shared by the classifier.
- `classify.mjs` — the verdict vocabulary and anti-laundering registries (extends
  `../differential/classify.mjs`'s pattern — see that file's doc comment for the
  vocabulary this reuses).
- `run-classify.mjs` — joins the two results files on scenario id, classifies every
  scenario, writes `results/classify-results.json`, prints the summary table.

## How to re-run everything

```sh
cd compat
npm install                                    # installs xlsx (already a devDependency)
cd ../..
cargo build --release --bin elixcee             # build the CLI once

cd compat/corpus
node workbooks/generate-workbooks.mjs           # regenerate the 5 base .xlsx fixtures
node generate-scenarios.mjs                     # regenerate scenarios.json (deterministic)
node run-elixcee.mjs                            # real execution against elixcee, ~seconds
node run-libreoffice.mjs                        # real execution against LibreOffice —
                                                 # see timing note below before running
                                                 # the full corpus serially
node run-classify.mjs                           # joins + classifies, prints the summary
```

Self-checks (no framework, run directly): `node normalize.mjs`, `node classify.mjs`.

### Running the LibreOffice leg in parallel shards

`run-libreoffice.mjs [count] [startIndex] [outSuffix]` runs a slice of `scenarios.json`
and writes `results/libreoffice-results<outSuffix>.json`, so multiple instances can run
concurrently over disjoint slices (each gets its own temp profile/workdir via
`mkdtempSync` — no shared mutable state). This matters because of the timeout behavior
described next: running the full corpus serially at 8s/scenario would take over an hour.
Example, 8-way split of a 580-scenario corpus:

```sh
for i in 0 1 2 3 4 5 6 7; do
  node run-libreoffice.mjs 73 $((i * 73)) "_shard$i" &
done
wait
```

`run-classify.mjs` globs every `results/libreoffice-results*.json` file, so shard output
needs no merging step.

## Known, reproducible limitation: VBA object-model access hangs under headless LibreOffice

`run-libreoffice.mjs`'s harness macro invokes each scenario's Sub via
`oDoc.getScriptProvider().getScript(uri).invoke(Array(), Array(), Array())` — the
standard, documented way to run a Basic/VBA macro embedded in a document from another
running macro. This works and completes quickly for VBA that doesn't touch the Excel
object model (confirmed with a trivial `x = 1 + 1` body). It **hangs indefinitely** — not
slow, not eventually-completing; confirmed still hung after 90+ seconds with no CPU
activity change — as soon as the invoked code touches `Range(...)` or `Cells(...)`,
whether reading or writing, confirmed identically across:

- an in-memory document (`private:factory/scalc`) vs. a document saved to real `.xlsm`
  (via the `Calc MS Excel 2007 VBA XML` filter) and reopened fresh,
- `Hidden:=True` vs. `Hidden:=False` on load,
- a bare `Range("A1")` read vs. `Cells(1,1)`.

A second invocation path was tried — triggering execution via the document-load filter
itself (`Auto_Open`, with `MacroExecutionMode = ALWAYS_EXECUTE_NO_WARN`) instead of a
nested `invoke()` call — which does **not** hang, but the `Auto_Open`
convention did not appear to fire on load in this environment either (no observable
side effect from the macro body), and this was not investigated further given time
constraints. Root cause of the original hang (a lock held by the invoking macro that
VBA-runtime initialization needs, per the working theory) was never confirmed with a
stack sample of the hung process — that's the concrete next step for anyone picking this
back up, not a re-guess at the property name or invocation path (both already
double-checked; see git history on this file for the isolation trail if useful).

**Practical consequence**: because this corpus is built specifically to exercise
`Range`/`Cells` (that's the VBA object-model surface elixcee implements), nearly every
scenario times out under LibreOffice, not because the scenario is wrong but because of
this invocation-path limitation. `run-libreoffice.mjs` uses an 8-second per-scenario
timeout (well above the ~1-3s observed completion time for non-hanging code, well below
a serial full-corpus run's practical wall-clock budget) and records `status: "TIMEOUT"`ed
scenarios honestly — `classify.mjs` maps these to the `ORACLE_UNAVAILABLE` verdict, not
silently as a MATCH or a skip. See the top-level task report for the actual measured
counts from the real run.

### Why this invocation shape (soffice CLI macro URI) rather than pyuno/unohttpd

Two other automation paths from the milestone's brief were tried first:

- **pyuno via LibreOffice's own bundled Python 3.12** — the interpreter binary
  (`.../LibreOfficePython.framework/Versions/3.12/bin/python3.12`) is killed (SIGKILL,
  exit 137) immediately on invocation in this sandboxed macOS environment, even after
  removing its `com.apple.quarantine` xattr, even for `--version`. Not pursued further —
  the CLI macro-invocation path below doesn't depend on this binary at all.
- **A pyuno socket bridge from system Python** (`soffice --accept=socket,...` +
  `import uno` from `/usr/bin/python3` or similar) was not attempted once the CLI-macro
  path proved viable — `pyuno.so` is compiled against LibreOffice's bundled CPython 3.12
  ABI specifically, and the system Python here is 3.13, a mismatch that would need its
  own investigation for uncertain benefit over the working approach.

`soffice --headless "vnd.sun.star.script:Library.Module.Sub?language=Basic&location=..."`
— invoking a Basic macro by URI from the command line — worked reliably once two
non-obvious pitfalls were found and fixed:

1. **`Environ()`, called more than once per macro invocation, crashes the whole `soffice`
   process** (confirmed: works fine called once; a second call anywhere in the same Sub —
   even reading an unrelated, even unset, variable name — brings down the process with no
   error surfaced anywhere). Fix: never read scenario parameters via `Environ()` inside
   the harness macro at all — `run-libreoffice.mjs` instead generates a fresh
   `Module1.xba` per scenario with every value (file paths, VBA source, entrypoint name)
   baked in as Basic string literals.
2. **`.xba` module files are XML.** A raw, unescaped `&` anywhere in the Basic source —
   including inside a string literal, including as VBA's own `&` string-concatenation
   operator — produces a malformed XML document that LibreOffice loads as if the module
   were empty, with no error surfaced anywhere (the macro URI just silently resolves to
   nothing, `soffice` exits 0). This was misdiagnosed at length as "the `&` operator is
   broken in this LibreOffice build" before the actual cause (XML well-formedness, not a
   Basic-language issue) was found. Fix: `run-libreoffice.mjs`'s `xmlEscape()` escapes
   `&`/`<`/`>` in the fully-assembled Basic source before writing the `.xba` file — this
   correctly handles both the harness's own concatenation and any `&` the *scenario's*
   VBA source itself contains once XML-decoded back to real Basic source at load time.

## Silent-wrong-result checking

`normalize.mjs`'s LibreOffice-side normalizer deliberately reads each cell's type via
`getType()` (VALUE/TEXT/FORMULA) rather than a bare `getValue()` — `getValue()` returns
`0` for a text cell, and elixcee can legitimately also produce a real `0` for a numeric
cell at the same address; comparing via `getValue()` alone would silently manufacture a
false MATCH on that address. See the task's final report for the actual count of
scenarios (if any) where elixcee produced a plausible-looking but incorrect result.
