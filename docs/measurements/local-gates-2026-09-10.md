# Self-contained local gate record

Date: 2026-09-10 (Asia/Tokyo)

This record captures the repository-local candidate gate on macOS arm64. It
does not claim Excel compatibility, three-OS support, external review, or
publication readiness.

## Environment and command

- Branch: `release/0.21.0`
- Version check: `1.0.5`
- Command: `bash scripts/check-local-gates.sh`
- Cargo mode: offline
- Fuzz mode: nightly, four targets, 5 seconds each, 1 GiB RSS limit
- Browser: Google Chrome 152.0.7977.83, headless real-browser smoke

Before the successful run, the repository's regenerable Cargo build cache was
removed with `cargo clean` after a fixture test hit `No space left on device`.
The command reported 10.2 GiB removed; no repository source or user workbook
was removed. The successful run started with approximately 4.4 GiB free and
ended with approximately 3.0 GiB free.

## Result

All self-contained gates passed:

- version, measurement-boundary, formula-dispatch, OOXML-matrix, and stream
  measurement-validator checks;
- formatting and diff checks;
- Rust workspace `--all-targets` tests: 1,679 library tests passed, including Drawing flip, fill, line-color, line-width, line-dash, and worksheet code-name event regressions;
- benchmark smoke, blackbox/CLI/property/XLSX round-trip targets;
- workspace Clippy with warnings denied, Rustdoc, and offline audit;
- formula parser/evaluator, VBA parser, and XLSX-reader fuzz smoke, all with
  exit status 0 and RSS below the 1 GiB limit;
- WASM smoke, including Node/browser conditions, CJS/ESM bundles, typed
  WorkbookEditor operations, operation-plan undo, and the 10% payload gate;
- real packed npm tarball consumer for CJS, ESM, browser conditions, file
  writes, export parity, and TypeScript declarations;
- real Chrome browser smoke over HTTP.

This rerun covered commit `e3d8e27` (`feat: edit drawing line dash`) and its
ancestors through worksheet code-name event resolution; all self-contained gates
completed successfully. The documentation-only follow-up is `9ff3236`.

The gate output recorded a WASM payload growth of 9.95%, within the 10% local
baseline limit. The browser smoke returned `ok: true`, read the sample sheet,
and completed write/read round-trip checks.

## Follow-up after the snapshot change

On the same host, a fresh run after the Python snapshot/date1904 changes passed
version, measurement-boundary, formula-dispatch, OOXML-matrix,
stream-validator, diff, and format checks. The Rust all-target build then
stopped at the linker with `errno=28` (`No space left on device`). The
reproducible command was:

```sh
env TMPDIR=/private/tmp bash scripts/check-local-gates.sh
```

This is an environmental incomplete run, not a passing full-gate claim. The
regenerable `target` cache was removed with `cargo clean` afterward; source
files and user workbooks were not removed.

After removing that cache, the Rust gate was rerun with
`CARGO_INCREMENTAL=0 RUSTFLAGS='-C debuginfo=0'` to reduce temporary linker
storage. Workspace all-target tests passed (1,679 library tests plus all
workspace targets and benchmark smoke), and warnings-denied Clippy and
Python-feature Rustdoc also passed. This is a valid low-disk Rust verification
run; the canonical all-gates command remains incomplete because the original
default-profile link step still exceeded the host's available disk space.

The same follow-up also passed `cargo audit --no-fetch --stale` and all
`packages/xlsx` checks: TypeScript (DOM and no-DOM), operation-plan tests,
package audit, WASM smoke, the real packed-tarball CJS/ESM/browser consumer,
and real Chrome browser smoke. Nightly fuzz smoke was not rerun after the
cache cleanup because rebuilding its targets would exceed the remaining disk
budget; it remains an explicit open gate.

The four fuzz targets were subsequently rerun sequentially with the same
5-second / 1 GiB RSS limit and low-disk build settings. Results were:

- `fuzz_formula_parser`: 213,931 runs, peak reported RSS 538 MiB
- `fuzz_formula_eval`: 145,288 runs, peak reported RSS 538 MiB
- `fuzz_vba_parser`: 209,211 runs, peak reported RSS 522 MiB
- `fuzz_xlsx_reader`: 21,408 runs, peak reported RSS 300 MiB

All four exited with status 0 and reported no crash artifact. This closes the
local macOS fuzz-smoke rerun; long-duration fuzzing and other OS/resource
profiles remain open.

## Boundaries

This is local BUILD and consumer-smoke evidence only. It does not complete the
following roadmap items:

- independent Excel oracle and Excel reopen/repair-warning checks;
- Linux/Windows clean-install, resource, and long-running verification;
- fixed-version comparisons with LogiSheets, EPPlus, or Aspose.Cells;
- external review, registry publication, tag, or formal release.
