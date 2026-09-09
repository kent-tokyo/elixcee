# G2d object editing local verification

Date: 2026-09-09 (Asia/Tokyo)

This record covers only the bounded local BUILD for existing Chart-series,
Chart-series-cache, Chart-title/style/axis-title/legend-overlay/data-label, two-cell Drawing-anchor/shape-name/hidden, rotation/flip/fill/line-color/line-width/line-dash, worksheet-backed
Pivot-source, Pivot refresh-policy, Pivot field-caption, Drawing
alternative-text edits, and existing DrawingML text-run editing.
It is not evidence of
Excel reopening, Pivot recalculation, or general OOXML object compatibility.

## Scope

- `Vm.set_chart_series_formulas(chart_part, series_index, categories, values)`
  rewrites only selected series category/value `<c:f>` elements.
- `Vm.set_pivot_worksheet_source(cache_part, sheet, reference)` rewrites only
  the first worksheet-backed `worksheetSource` `sheet` and/or A1 `ref`.
- `Vm.set_pivot_cache_refresh_on_load(cache_part, enabled)` adds or replaces
  only the root cache definition's `refreshOnLoad` flag.
- `Vm.set_pivot_cache_field_caption(cache_part, field_index, caption)` rewrites
  only an existing `cacheField@name` attribute.
- `Vm.set_chart_title(chart_part, text)` rewrites only the first `<a:t>` inside
  the first `<c:title>`, or adds a minimal rich-text title before `plotArea`
  when the chart has no title.
- `Vm.set_chart_legend_position(chart_part, position)` rewrites only the first
  `<c:legendPos@val>` for the allowed values `b`, `tr`, `r`, `l`, or `t`, or
  adds a minimal legend before `plotArea` when the chart has no legend.
- `Vm.set_chart_style(chart_part, style)` rewrites or adds only the chart-space
  `<c:style val>` with an Excel style number from 1 through 48.
- `Vm.set_chart_axis_title(chart_part, axis_index, text)` rewrites only the
  first text run of an existing axis title, or adds a minimal title when absent;
  axes are zero-based in document order.
- `Vm.set_chart_legend_overlay(chart_part, overlay)` rewrites or adds only the
  first legend's `overlay` flag; a missing legend is rejected.
- `Vm.set_chart_data_labels_show_value(chart_part, show_value)` rewrites or
  adds only the first existing `<c:dLbls showVal>` flag; a missing data-label
  element is rejected.
- `Vm.set_chart_data_labels_show_category(chart_part, show_category)` rewrites
  or adds only the first existing `<c:dLbls showCat>` flag and composes with
  `showVal`; a missing data-label element is rejected.
- `Vm.set_chart_data_labels_show_series_name(chart_part, show_series_name)`
  rewrites or adds only the first existing `<c:dLbls showSerName>` flag and
  composes with `showVal`/`showCat`; a missing data-label element is rejected.
- `Vm.set_chart_data_labels_show_percent(chart_part, show_percent)` rewrites
  or adds only the first existing `<c:dLbls showPercent>` flag; a missing
  data-label element is rejected.
- `Vm.set_chart_data_labels_show_leader_lines(chart_part, show_leader_lines)`
  rewrites or adds only the first existing `<c:dLbls showLeaderLines>` flag;
  a missing data-label element is rejected.
- `Vm.set_chart_data_labels_show_bubble_size(chart_part, show_bubble_size)`
  rewrites or adds only the first existing `<c:dLbls showBubbleSize>` flag;
  a missing data-label element is rejected.
- `Vm.set_chart_data_labels_show_legend_key(chart_part, show_legend_key)`
  rewrites or adds only the first existing `<c:dLbls showLegendKey>` flag;
  a missing data-label element is rejected.
- `Vm.set_chart_data_labels_position(chart_part, position)` rewrites or adds
  the first existing `<c:dLblPos val>` element using the bounded OOXML
  vocabulary, preserving other data-label flags and child elements.
- `Vm.set_chart_data_labels_number_format(chart_part, number_format)` rewrites
  or adds the first existing `<c:numFmt formatCode>` under `c:dLbls`, with
  XML escaping and control-character/size validation.
- `Vm.set_chart_data_labels_separator(chart_part, separator)` rewrites or adds
  the first existing `<c:separator val>` under `c:dLbls`, with XML escaping
  and control-character/size validation.
- `Vm.set_chart_data_labels_number_format(chart_part, number_format)` rewrites
  or adds the first existing `<c:numFmt formatCode>` under `c:dLbls`, with
  XML escaping and control-character/size validation.
- `Vm.set_chart_series_name_formula(chart_part, series_index, name_formula)`
  rewrites only the `<c:f>` inside the selected series' `<c:tx>` element.
- `Vm.set_chart_series_cache(chart_part, series_index, categories, values)`
  rewrites existing category/value cache points while preserving the cache
  kind, formula, format code, and surrounding Chart XML.
- `Vm.set_drawing_anchor(drawing_part, anchor_index, from_row, from_col,
  to_row, to_col)` rewrites only the `<xdr:from>` and `<xdr:to>` cell markers
  of the selected two-cell anchor. Public cell coordinates are 1-based.
- `Vm.set_drawing_shape_name(drawing_part, anchor_index, name)` rewrites only
  the selected drawing anchor's `<xdr:cNvPr name>` attribute. The index spans
  two-cell, one-cell, and absolute anchors in document order.
- `Vm.set_drawing_shape_description(drawing_part, anchor_index, description)`
  adds or replaces only the optional `<xdr:cNvPr descr>` attribute.
- `Vm.set_drawing_shape_title(drawing_part, anchor_index, title)` adds or
  replaces only the optional `<xdr:cNvPr title>` attribute.
- `Vm.set_drawing_shape_text(drawing_part, anchor_index, text)` replaces only
  the first existing `<a:t>` in the selected anchor. Missing text runs are
  rejected; other runs and shape/relationship content remain opaque.
- `Vm.set_drawing_shape_hidden(drawing_part, anchor_index, hidden)` adds or
  replaces only the `<xdr:cNvPr hidden>` flag.
- `Vm.set_drawing_shape_rotation(drawing_part, anchor_index, degrees)` updates
  only an existing shape transform's integer-degree rotation.
- `Vm.set_drawing_shape_flip(drawing_part, anchor_index, flip_horizontal,
  flip_vertical)` updates only the supplied existing transform flip flags.
- `Vm.set_drawing_shape_fill(drawing_part, anchor_index, color)` updates or
  adds a solid RGB/ARGB fill while preserving the shape property container.
- `Vm.set_drawing_shape_line_color(drawing_part, anchor_index, color)` updates
  only an existing line's solid RGB/ARGB color.
- `Vm.set_drawing_shape_line_width(drawing_part, anchor_index, width_points)`
  updates only an existing line width after points-to-EMU conversion.
- `Vm.set_drawing_shape_line_dash(drawing_part, anchor_index, dash)` updates
  only an existing `<a:prstDash>` using the DrawingML preset vocabulary.
  The save path is covered by `drawing_shape_line_dash_edit_survives_xlsx_save`
  using a minimal Drawing-backed XLSX fixture. The same save round-trip now
  covers `Vm.set_drawing_shape_text`, XML escaping (`&` to `&amp;`), and
  retention of a second text run.
- `Vm.set_chart_series_cache(...)` updates an existing cache or creates the
  matching cache when the selected series has only a `strRef`/`numRef` formula.
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
cargo test chart_legend_overlay --offline -- --nocapture
cargo test chart_data_labels_rewriter --offline -- --nocapture
cargo test chart_series_rewriter --offline -- --nocapture
cargo test chart_series_cache --offline -- --nocapture
cargo test pivot_refresh_on_load --offline -- --nocapture
cargo test pivot_cache_field_caption --offline -- --nocapture
cargo test drawing_anchor --offline -- --nocapture
cargo test drawing_shape_line_dash_edit_survives_xlsx_save --offline -- --nocapture
cargo test drawing_shape_text_rewriter --offline -- --nocapture
cargo test --workspace --all-targets --offline --quiet
cargo clippy --workspace --all-targets --offline -- -D warnings
cargo fmt --all -- --check
git diff --check
bash scripts/check-local-gates.sh
```

## Result

On the recorded macOS arm64 environment, the Chart title, legend-overlay,
data-label, series-name, series-cache, Pivot refresh-policy/field-caption,
Drawing anchor/shape-name/text-run, line-dash persistence, and real-fixture targeted tests passed. The full Rust workspace passed the current workspace test set,
including 54 XLSX round-trip tests. Strict
clippy, formatting, version/formula/OOXML checks, offline audit, four five-second
fuzz smoke targets, TypeScript checks, WASM smoke, packed npm consumer smoke,
and a real Chrome browser smoke all passed through
`scripts/check-local-gates.sh`.

The new text-run rewriter test passed, and a fresh offline maturin build
produced a CPython 3.13 arm64 wheel exposing
`Vm.set_drawing_shape_text`. The normal `cargo test --features python` test
binary is not used as Python binding evidence because it requires embedding
Python symbols; the extension build is the supported boundary check.

After the implementation commit, a clean low-disk rebuild reran
`cargo test --offline --lib drawing_shape_text_rewriter` (1 passed) and
warnings-denied `cargo clippy --offline --lib --all-targets` (success). The
wheel was installed into the isolated test environment and the binding was
resolved as `Vm.set_drawing_shape_text`.

The subsequent low-disk workspace all-target run passed 1,680 library tests
and all integration/benchmark targets, including the new text-run regression.

The minimal Drawing-backed XLSX save test passed with both line-dash and
text-run edits. After reload from the saved ZIP, the selected first run was
`Updated &amp; text`, while the unrelated second run remained `Keep`.

The title and series-name rewriter unit tests replaced XML-escaped text while
retaining unrelated text runs, series references, and the surrounding plot
area. The Pivot fixture changed
`ref="A1:B2"` to `ref="A1:C3"` while the renamed
worksheet source became `sheet="Data &amp; 2026"`. The Chart fixture changed
the selected series' name, category, and value formulas and retained
`xl/drawings/drawing1.xml` and its relationship part.
The same real fixture moved the selected two-cell anchor's from/to markers
while preserving its offsets, shape content, and drawing relationship. Its
non-visual shape name was XML-escaped and updated. Its
existing string and numeric series caches were updated with escaped text and
new point counts while their formulas and format code remained unchanged. The
data-label unit test changed `c:dLbls@showVal` from `0` to `1` and back, then
added the missing attribute while preserving an unrelated label child.

## Not measured or claimed

- Excel reopen, repair-warning absence, or recalculated Chart/Pivot caches.
- Excel-side execution of `refreshOnLoad` and external source retrieval.
- Chart or Drawing creation, general Drawing editing beyond two-cell anchor
  markers, shape metadata/style, and the bounded first text run, multiple
  text-run editing, cache recalculation beyond
  writing cached points, or table-backed Pivot sources.
- Linux/Windows clean-install, resource calibration, external review, or
  LogiSheets/EPPlus/Aspose.Cells comparison.
