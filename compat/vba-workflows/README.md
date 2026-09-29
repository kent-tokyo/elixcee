# Headless VBA workflow corpus

This directory defines the V0 comparison contract for workbook-processing
workflows. It is intentionally separate from `compat/vba-diagnostics`, which
contains focused diagnostic cases, and from `compat/vba-semantics`, which is a
local semantics suite.

`manifest.json` is a metadata contract, not a claim that the cases have been
measured. Each case records the workbook, VBA source, operation boundary,
expected output, allowed effects, resource limits, provenance, and current
measurement status. A missing Excel/xlflow runner must remain
`not_measured`.
Optional `local` expectations may add a local CLI error class and an existing
output sentinel; these do not change the external oracle status. The local
runner verifies both the detailed error kind and stable `termination_class`,
and stores the observed classification in its report. It also enables the
bounded CLI trace and stores only its event count and digest, not trace values.

Validate the contract with:

```sh
python3 scripts/check-vba-workflow-manifest.py
```

After building the local CLI, run the elixcee-only self-check with:

```sh
python3 scripts/run-vba-workflow-local.py --binary target/debug/elixcee
```

Add `--report path.json` to retain a machine-readable report containing the
manifest/source/workbook hashes, local result class, output hash, and failures.
The report records the external oracle status separately and never promotes a
local self-check to an Excel/xlflow measurement.

For a local wall-time baseline, run `scripts/benchmark-vba-workflows.py` with
at least three measured repetitions. It performs five warmup runs by default
and excludes them from p50/p95; use `--warmup 0` only when that is intentional.
It measures the complete CLI process boundary for successful cases only; RSS
and Excel/xlflow remain separate axes.

For the VM-only boundary, build and run the measurement-only binary:

```sh
cargo run --release --bin benchmark_vba_vm -- --rows 10000 --repetitions 30
```

Its `elixcee.vba-vm-benchmark.v1` JSON excludes parse and workbook I/O. It is
useful for elixcee before/after regression checks, not for an Excel/xlflow
speed claim.

This verifies the declared cell results both in the CLI JSON and after reading
the published XLSX with Python's standard library; it does not replace the
Excel/xlflow oracle and leaves `measurement_status` as `not_measured`.
For cases declaring `formulas_absent` or `formulas_present`, it also checks the
saved worksheet XML to ensure formula elements were removed or retained as
declared.

The eventual runners must keep these result classes distinct:
`pass`, `mismatch`, `expected_error`, `unsupported`, `policy_blocked`,
`timeout`, and `not_measured`. A wrapper used only by the measurement harness
does not count as an unchanged-source success.

For XLSM provenance, run the dependency-free auditor:

```sh
python3 scripts/audit-xlsm-vba-provenance.py workbook.xlsm --source Module1.bas
```

It reports the workbook hash, preserved `vbaProject.bin` parts, OOXML-side
`ThisWorkbook`/worksheet `codeName` and VBA relationship presence, separately
supplied source hashes, and an explicit
`not_available_without_ole_source_parser` module-identity state. It never
extracts, modifies, or executes embedded VBA.
