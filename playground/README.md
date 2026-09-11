# elixcee Playground

Try elixcee in the browser: open an `.xlsx` or `.xlsm` file, switch between sheets, edit
cells and ranges, format and filter data, create tables and bar/line charts, run
a bounded VBA subset in a Web Worker, recalculate with Rust/WASM, inspect
diagnostics, and download a verified `.xlsx` or `.xlsm` file. XLSM's
`xl/vbaProject.bin` is preserved as opaque bytes but never executed in the browser;
use the native runtime for full VBA. The page is client-only;
workbook bytes are not uploaded.
Workbooks containing `xl/externalLinks/` parts are rejected by the browser editor
because this package does not yet guarantee lossless external-link preservation;
the same applies to PivotTables/PivotCaches, embedded media, threaded comments,
slicers, and custom XML. Use the native runtime when those parts must be retained.

Open the live [GitHub Pages playground](https://kent-tokyo.github.io/elixcee/playground/)
or read the [English quick start](../docs/quickstart.md) first.
Use the in-page language selector to switch the interface and quick-start link
between English, Japanese, and Simplified Chinese; translations are not shown
side by side.

The browser page covers the private `@elixcee/xlsx` browser package: in-memory
XLSX read/write, multiple sheet tabs, typed workbook edits, formula
recalculation, charts, tables, and structured preflight diagnostics. Its VBA
sandbox supports a bounded subset of `Sub`, `Dim`, scalar assignment,
`Cells(row,col).Value`, arithmetic, `If`, `For`, `Do`, and `Select Case` in a
Worker. File, network, COM, Shell, and UI effects are rejected. For full
headless VBA data processing, use the native/Python runtime shown in the
[beginner tutorial](../docs/tutorial-beginners.md).

The **Pivot summary** button creates a bounded, worksheet-backed category summary (sum, count, or average) and refreshes it when its source changes. It is not an OOXML PivotTable/PivotCache editor; native PivotTable files are rejected by the browser preflight.

## Run the Python sample locally

```bash
python playground/run.py
python playground/run.py --output /tmp/elixcee-playground.xlsx
```

Install the Python package first if needed:

```bash
python -m pip install elixcee
```

Translations: [日本語](README-ja.md) · [简体中文](README-zh.md)
