# Compatibility harnesses

These suites answer different questions. A passing synthetic fixture, JS reference
case, workbook round-trip, and real Excel execution are not interchangeable evidence.
The Node project is private and uses the fixed `xlsx@0.18.5` development oracle.

The native OOXML feature boundary is tracked in the machine-readable
[feature matrix](ooxml-feature-matrix.json). It separates read, preservation,
editing, recalculation, and Excel-reopen evidence; `preserved` does not mean
editable, and `unverified` is not counted as a compatibility success.

From this directory, install dependencies with `npm ci`.
Native VBA suites also require `cargo build --release --bin elixcee` at the repository root.

## JavaScript differential tests

Compare the local [JS package](../packages/xlsx/README.md) to the installed reference.
Normalize logical workbook contents rather than ZIP timestamps.

```sh
npm run classifier:self-check
npm run differential:utils
npm run differential:ssf-format
npm run differential:read
npm run differential:write
npm run differential:metadata
```

Verdicts: MATCH, INTENTIONAL_SECURITY_DIVERGENCE, INTENTIONAL_SAFETY_DIVERGENCE,
UNSUPPORTED, BUG, NONDETERMINISTIC. Allowed differences have explicit registry
entries in [classify.mjs](differential/classify.mjs), checked by classifier:self-check.
Use current run output for counts; old counts are not current coverage claims.

## VBA corpus

The [corpus](corpus/README.md) checks declared pass/fail expectations against the
native CLI. Its synthetic, real-world-shaped scenarios do not prove Excel equivalence.

```sh
npm run corpus:run
npm run corpus:outcomes
```

PASS, EXPECTED_RUNTIME_ERROR, EXPECTED_UNSUPPORTED, and NONDETERMINISTIC are distinct
from MISMATCH and UNEXPLAINED. Do not silently treat a crash, hang, or unknown result
as expected. The LibreOffice runner's recorded availability limitations are in the
corpus README; its outcomes are not Microsoft Excel measurements.

## VBA semantic reference cases

The [semantic suite](vba-semantics/README.md) compares values with independent JS
reference implementations derived from documented VBA semantics. It is not a live
Excel oracle and has explicitly classified known limitations.

```sh
npm run semantics:generate   # only after changing case generation
npm run semantics:run
npm run semantics:report
```

## Runtime API inventory

`npm run oracle:manifest` regenerates [api-manifest.json](oracle/api-manifest.json)
from the installed reference's CJS/ESM runtime exports. Review changes when changing
the pinned oracle. The old differential:demo remains only a classifier plumbing smoke.

## Microsoft Excel evidence

The Windows/COM scenario adapter remains
[unverified scaffolding](oracle-excel-com/UNVERIFIED.md); see its
[contract](oracle-excel-com/CONTRACT.md) and [execution plan](oracle-excel-com/WINDOWS_EXECUTION.md).

Separately, the [0.9.0-A round-trip record](oracle-excel-com/results/0.9.0-A_summary.md)
documents selected workbooks authored/reopened in Excel for Mac.
It does **not** verify the whole VBA corpus or post-save macro execution.
Its version-specific losses must not be presented as a current implementation inventory.

## Performance and CI

[Workbook benchmarks](../docs/benchmarks/README.md) and
[ClosedXML worker](benchmarks/closedxml/README.md) measure performance separately
from these compatibility verdicts.

[CI configuration](../.github/workflows/ci.yml) is authoritative for enabled jobs,
platforms, and Node versions. Run the affected suites locally when changing their
reader, writer, parser, VM, or wrapper paths. [Documentation map](../docs/README.md)
