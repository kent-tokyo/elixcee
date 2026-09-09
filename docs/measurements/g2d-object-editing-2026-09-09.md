# G2d object editing local verification

Date: 2026-09-09 (Asia/Tokyo)

This record covers only the bounded local BUILD for existing Chart-series,
Chart-series-cache, Chart-title, two-cell Drawing-anchor, and worksheet-backed
Pivot-source edits.
It is not evidence of
Excel reopening, Pivot recalculation, or general OOXML object compatibility.

## Scope

- `Vm.set_chart_series_formulas(chart_part, series_index, categories, values)`
  rewrites only selected series category/value `<c:f>` elements.
- `Vm.set_pivot_worksheet_source(cache_part, sheet, reference)` rewrites only
  the first worksheet-backed `worksheetSource` `sheet` and/or A1 `ref`.
- `Vm.set_chart_title(chart_part, text)` rewrites only the first `<a:t>` inside
  the first `<c:title>` element.
- `Vm.set_chart_series_name_formula(chart_part, series_index, name_formula)`
  rewrites only the `<c:f>` inside the selected series' `<c:tx>` element.
- `Vm.set_chart_series_cache(chart_part, series_index, categories, values)`
  rewrites existing category/value cache points while preserving the cache
  kind, formula, format code, and surrounding Chart XML.
- `Vm.set_drawing_anchor(drawing_part, anchor_index, from_row, from_col,
  to_row, to_col)` rewrites only the `<xdr:from>` and `<xdr:to>` cell markers
  of the selected two-cell anchor. Public cell coordinates are 1-based.
- Missing source parts, series, references, malformed attributes, control
  characters, and invalid A1 ranges are rejected before a successful save.
- Explicit Pivot source edits also require the requested worksheet to exist in
  the loaded workbook.
- Existing Drawing shape content, relationship parts, unrelated Chart caches,
  Pivot cache records, and PivotTable layout are preserved or left opaque by
  this scope.

## Reproduction

From the repository root, run:

```text
cargo test --test xlsx_roundtrip pivot_worksheet_source --offline -- --nocapture
cargo test --test xlsx_roundtrip edit_chart_series --offline -- --nocapture
cargo test chart_title_rewriter --offline -- --nocapture
cargo test chart_series_rewriter --offline -- --nocapture
cargo test chart_series_cache --offline -- --nocapture
cargo test drawing_anchor --offline -- --nocapture
cargo test --workspace --all-targets --offline --quiet
cargo clippy --workspace --all-targets --offline -- -D warnings
cargo fmt --all -- --check
git diff --check
bash scripts/check-local-gates.sh
```

## Result

On the recorded macOS arm64 environment, the Chart title, series-name,
series-cache, Drawing anchor, and real-fixture targeted tests passed. The full
Rust workspace passed 1,608 tests,
including 54 XLSX round-trip tests. Strict
clippy, formatting, version/formula/OOXML checks, offline audit, four five-second
fuzz smoke targets, TypeScript checks, WASM smoke, packed npm consumer smoke,
and a real Chrome browser smoke all passed through
`scripts/check-local-gates.sh`.

The title and series-name rewriter unit tests replaced XML-escaped text while
retaining unrelated text runs, series references, and the surrounding plot
area. The Pivot fixture changed
`ref="A1:B2"` to `ref="A1:C3"` while the renamed
worksheet source became `sheet="Data &amp; 2026"`. The Chart fixture changed
the selected series' name, category, and value formulas and retained
`xl/drawings/drawing1.xml` and its relationship part.
The same real fixture moved the selected two-cell anchor's from/to markers
while preserving its offsets, shape content, and drawing relationship. Its
existing string and numeric series caches were updated with escaped text and
new point counts while their formulas and format code remained unchanged.

## Not measured or claimed

- Excel reopen, repair-warning absence, or recalculated Chart/Pivot caches.
- Chart or Drawing creation, general Drawing editing beyond two-cell anchor
  markers, multiple title-run editing, cache creation/regeneration beyond
  existing caches, or table-backed Pivot sources.
- Linux/Windows clean-install, resource calibration, external review, or
  LogiSheets/EPPlus/Aspose.Cells comparison.
