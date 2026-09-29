# V6 Python workflow self-check — 2026-09-12

This is an elixcee-only Python binding check. It is not an Excel or xlflow
oracle measurement.

## Protocol

- package: locally built `elixcee` v1.0.12 wheel for CPython 3.13 / macOS arm64
- input: the six cases in `compat/vba-workflows/manifest.json`
- operation: `load_workbook` → `Vm.run_and_save` → inspect
  `last_termination_class`
- successful cases: 5
- expected failure cases: 1
- output: each successful case wrote a separate temporary XLSX; the failure
  case started with an existing sentinel output

## Result

| Cases | Observed classification | Output behavior |
|---:|---|---|
| 5 success workflows | `success` | 5 XLSX outputs created |
| 1 failure-boundary workflow | `runtime_error` | existing sentinel remained byte-identical |

The failure case also confirmed that the caller VM did not publish the partial
macro result. This validates the Python binding boundary locally; it does not
prove Excel compatibility, full VBA coverage, cross-platform parity, or
competitor performance.
