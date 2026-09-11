# G2d object editing local verification

Date: 2026-09-10 (Asia/Tokyo; refreshed)

This record covers the bounded local BUILD for new and existing Chart content,
including existing Chart-series,
Chart-series-cache/smooth/invert-if-negative/visibility/line-color/fill-color, Chart-title/style/axis-title/legend-overlay/data-label, two-cell Drawing-anchor/shape-name/hidden, rotation/flip/fill/line-color/line-width/line-dash, worksheet-backed
Pivot-source, Pivot refresh-policy, Pivot field-caption, Drawing
alternative-text edits, and existing DrawingML text-run editing.
It is not evidence of Pivot recalculation or general OOXML object compatibility.

## Excel reopen status

The generated line-chart-only smoke case opened in Microsoft Excel for Mac
without a recovery warning. The complete regression case containing a second
Chart/barChart still opened as `修復済み` (Excel recovery). Its cells survived,
but the Chart result is not accepted as a valid Excel artifact. This is a
negative compatibility result, not a release claim; `compat/ooxml-feature-matrix.json`
therefore keeps `excel_reopen` as `unverified`.

The same result was reproduced on 2026-09-11 with the current v1.0.10
generator output (`target/tmp/g2d-chart-bar-smoke.xlsm`) from
`create_chart_connects_new_part_to_existing_drawing`. Excel displayed its
recovery dialog before opening the workbook. The warning was dismissed without
accepting recovery, so no repaired workbook was used as evidence.

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
  or adds only the first existing `<c:dLbls showCatName>` flag and composes with
  `showVal`; a missing data-label element is rejected.
- `Vm.set_chart_data_labels_show_series_name(chart_part, show_series_name)`
  rewrites or adds only the first existing `<c:dLbls showSerName>` flag and
  composes with `showVal`/`showCatName`; a missing data-label element is rejected.
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
- `Vm.set_chart_series_marker_symbol(chart_part, series_index, symbol)` updates
  only an existing series marker's `<c:symbol val>`. Supported symbols are
  `circle`, `dash`, `diamond`, `dot`, `none`, `picture`, `plus`, `square`,
  `star`, `triangle`, and `x`; a missing marker is rejected.
- `Vm.set_chart_series_marker_size(chart_part, series_index, size)` updates only
  an existing marker's `<c:size val>` with the bounded OOXML range 2..72.
- `Vm.set_chart_series_smooth(chart_part, series_index, smooth)` updates only
  an existing series `<c:smooth val>` flag (`0` or `1`); a missing flag is
  rejected rather than synthesized.
- `Vm.set_chart_series_invert_if_negative(chart_part, series_index, enabled)`
  updates only an existing `<c:invertIfNegative val>` flag (`0` or `1`); a
  missing flag is rejected rather than synthesized.
- `Vm.set_chart_series_deleted(chart_part, series_index, deleted)` updates only
  an existing `<c:delete val>` flag (`0` or `1`); the series is never removed
  from the chart XML and a missing flag is rejected.
- `Vm.set_chart_series_cache(chart_part, series_index, categories, values)`
  rewrites existing category/value cache points while preserving the cache
  kind, formula, format code, and surrounding Chart XML.
- `Vm.set_chart_series_line_color(chart_part, series_index, color)` updates
  only an existing series solid RGB line color. The six-digit RGB value may
  include `#`; theme colors, gradients, missing style chains, and unsupported
  series indexes are rejected rather than synthesized.
- `Vm.set_chart_series_fill_color(chart_part, series_index, color)` updates
  only an existing series solid RGB fill color. Theme colors, gradients,
  missing fill chains, and unsupported series indexes are rejected.
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
- `Vm.set_drawing_shape_text_run(drawing_part, anchor_index, run_index, text)`
  replaces one existing `<a:t>` selected by its zero-based run index within
  the anchor. Missing runs are rejected and unrelated runs remain opaque.
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
- `Vm.set_drawing_shape_geometry(drawing_part, anchor_index, preset)` updates
  only an existing `<a:prstGeom prst>`. A bounded DrawingML preset vocabulary
  is accepted; custom geometry and missing geometry are rejected.
  The save path is covered by `drawing_shape_line_dash_edit_survives_xlsx_save`
  using a minimal Drawing-backed XLSX fixture. The same save round-trip now
  covers `Vm.set_drawing_shape_text`, XML escaping (`&` to `&amp;`), and
  retention of a second text run.
- `Vm.set_chart_series_cache(...)` updates an existing cache or creates the
  matching cache when the selected series has only a `strRef`/`numRef` formula.
- `Vm.add_chart(...)` and `Vm.add_chart_series(...)` are BUILD-stage APIs for
  bounded line/bar/area/pie creation in an existing worksheet Drawing. They
  generate a Chart part, two-cell anchor, Drawing relationship, and content
  type override. New-sheet Drawing creation and Excel-reopen compatibility are
  not covered by this record.
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
cargo test --test xlsx_roundtrip edit_chart_series_smooth_survives_real_fixture_save --offline -- --nocapture
cargo test --test xlsx_roundtrip edit_chart_series_visibility_survives_real_fixture_save --offline -- --nocapture
cargo test --test xlsx_roundtrip edit_chart_series_line_color_survives_real_fixture_save --offline -- --nocapture
cargo test chart_title_rewriter --offline -- --nocapture
cargo test chart_legend_overlay --offline -- --nocapture
cargo test chart_data_labels_rewriter --offline -- --nocapture
cargo test chart_series_rewriter --offline -- --nocapture
cargo test chart_series_marker --offline -- --nocapture
cargo test chart_series_smooth --offline -- --nocapture
cargo test chart_series_marker_smooth_and_visibility --offline -- --nocapture
cargo test chart_series_cache --offline -- --nocapture
cargo test pivot_refresh_on_load --offline -- --nocapture
cargo test pivot_cache_field_caption --offline -- --nocapture
cargo test drawing_anchor --offline -- --nocapture
cargo test drawing_shape_line_dash_edit_survives_xlsx_save --offline -- --nocapture
cargo test drawing_shape_text_rewriter --offline -- --nocapture
cargo test drawing_shape_text_run_rewriter --offline -- --nocapture
cargo test --workspace --all-targets --offline --quiet
cargo clippy --workspace --all-targets --offline -- -D warnings
cargo fmt --all -- --check
git diff --check
bash scripts/check-local-gates.sh
```

## Result

On the recorded macOS arm64 environment, the Chart title, legend-overlay,
data-label, series-name, marker-symbol, series-cache, Pivot refresh-policy/field-caption,
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
and all integration/benchmark targets, including the first text-run regression.
After the indexed-run addition, the same low-disk command passed 1,681 library
tests and all integration/benchmark targets.

The indexed-run unit test and the minimal Drawing-backed XLSX save test passed
after the follow-up API addition. The saved ZIP contained
`Updated &amp; text` and `Second &amp; text`, confirming that an explicit run
index can update a non-first run without reconstructing the text body. The
follow-up full workspace run completed with 1,681 library tests; the command
used `TMPDIR=/private/tmp CARGO_INCREMENTAL=0 RUSTFLAGS='-C debuginfo=0'
cargo test --offline --workspace --all-targets --quiet`.

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

The follow-up Chart-series smooth-flag addition passed its focused rewriter
regression (1 test), warnings-denied library clippy, and the full workspace
offline all-target gate (1,683 library tests, 54 XLSX round-trip tests, and all
integration/benchmark targets). This remains BUILD evidence only and does not
establish Excel reopen or chart rendering compatibility.

The save-path follow-up also passed `edit_chart_series_smooth_survives_real_fixture_save`:
a temporary copy of the real Excel-authored fixture was loaded, its existing
series smooth flag changed from `0` to `1`, saved, and verified after ZIP
reload. The test validates the elixcee save path, not Excel's own reopen.

The corresponding `edit_chart_series_visibility_survives_real_fixture_save`
regression also passed with the same temporary-copy method, changing an
existing series `<c:delete val>` from `0` to `1` while retaining its formula.

On 2026-09-11, Excel for macOS was used to open the multi-chart creation output
`g2d-chart-bar-smoke-no-style-rels.xlsm`. Excel displayed its repair warning;
accepting recovery opened the sheet but removed the newly-created Charts. A
follow-up experiment that omitted the generated chart parts' optional
style/color relationship files produced the same warning and the same loss.
This rules out dangling style/color relationships as the sole cause; Chart
creation remains an Excel-reopen failure until the Chart XML/Drawing structure
is corrected. A bar-only output was then opened after adding the OOXML
`<c:crossBetween val="between"/>` value-axis element; Excel still showed the
same warning. That element is therefore not sufficient to fix the failure.
Adding an explicit literal `c:tx` series name to the generated bar chart was
also insufficient; Excel emitted the same warning. The candidate failure is
still in the broader Chart/Drawing package structure.

## Not measured or claimed

- Excel reopen, repair-warning absence, or recalculated Chart/Pivot caches.
- Excel-side execution of `refreshOnLoad` and external source retrieval.
- Chart or Drawing creation, general Drawing editing beyond two-cell anchor
  markers, shape metadata/style, and the bounded first text run, multiple
  text-run editing, cache recalculation beyond
  writing cached points, or table-backed Pivot sources.
- Linux/Windows clean-install, resource calibration, external review, or
  LogiSheets/EPPlus/Aspose.Cells comparison.
