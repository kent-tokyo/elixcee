# Migration and boundary guide

This guide describes the safe first step when replacing an Excel-dependent
batch job with elixcee 1.0.12. It is a migration guide, not a claim of full
Excel compatibility.

## Choose the runtime surface

- Use the native Python API when the job needs `.xlsx`/`.xlsm` read, supported
  formula recalculation, VBA data processing, and a saved workbook.
- Use the CLI when the job needs a reproducible headless run, JSON diagnostics,
  or an explicit external-link policy.
- Use the private JavaScript package only for its documented Node/browser/WASM
  read, data-only edit, and bounded editor operations. Its preservation and VBA
  contract is not identical to the native runtime.

## Minimal migration pattern

1. Keep the source workbook unchanged and write to a new destination path.
2. Start with `external_links="reject"` when external references are not part
   of the job; use `preserve` only when round-tripping the relationship is
   required, and never expect elixcee to fetch or recalculate an external URL.
3. Run the smallest supported macro or formula operation first, then inspect
   the structured result and diagnostics.
4. Treat charts, drawings, pivots, VBA project bytes, and unknown OOXML parts as
   preservation concerns unless the compatibility matrix marks the exact edit
   as supported.
5. Promote the output only after the relevant fixture and platform/oracle checks
   have passed. A local round-trip is not evidence of Microsoft Excel reopen
   compatibility.

Python sketch:

```python
import elixcee

vm = elixcee.load_workbook("input.xlsm", external_links="reject")
vm.set_cell_value(2, 1, 42)
vm.recalculate_all()
vm.save_workbook("output.xlsm")
```

The exact available methods and VBA/formula coverage are defined by
[FUNCTIONS.md](../FUNCTIONS.md) and [elixcee.pyi](../elixcee.pyi).

## Known loss and refusal boundaries

- A structural edit is refused when chart/drawing or pivot references cannot be
  rewritten safely. The narrow loaded-workbook sheet-rename path updates chart
  qualified formulas and pivot worksheet sources. Existing chart-series formulas
 and worksheet-backed Pivot source fields can also be edited explicitly. Chart
 title text and two-cell Drawing anchor markers are supported as bounded OOXML
 edits; these APIs do not create charts, perform general shape editing,
 regenerate caches, or recalculate pivots.
- Supported formula and VBA subsets are finite. Unsupported expressions should
  remain an explicit diagnostic, not be treated as an Excel result.
- GUI effects, Shell/COM/network/file side effects, UserForms, ActiveX, and
  arbitrary plugin code are outside the default execution boundary.
- The writer is bounded streaming, not a blanket constant-memory guarantee.
  Input size, XML/ZIP size, cell count, formula work, and execution budgets still
  apply.
- Unknown or unverified OOXML behavior is classified separately from preserved,
  supported, and rejected behavior in the [feature matrix](../compat/ooxml-feature-matrix.json).

## Evidence rule

The local gate proves build, tests, static contracts, package smoke, and the
documented local security checks. It does not prove Excel reopen behavior,
another OS's resource profile, or speed against a third-party library. Keep
those claims tied to dated records in [measurements](measurements/README.md).
