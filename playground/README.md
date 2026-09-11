# elixcee Playground

Try elixcee in the browser: open an `.xlsx` file, switch between sheets, edit
cells and ranges, run a bounded VBA subset in a Web Worker, recalculate with
Rust/WASM, inspect diagnostics, and download a verified `.xlsx` file. The page
is client-only; workbook bytes are not uploaded.

Open the live [GitHub Pages playground](https://kent-tokyo.github.io/elixcee/playground/)
or read the [English quick start](../docs/quickstart.md) first.
Use the in-page language selector to switch the interface and quick-start link
between English, Japanese, and Simplified Chinese; translations are not shown
side by side.

The browser page covers the private `@elixcee/xlsx` browser package: in-memory
XLSX read/write, multiple sheet tabs, typed workbook edits, formula
recalculation, and structured preflight diagnostics. Its VBA sandbox is
intentionally small: `Sub`, `Dim`, scalar assignment, `Cells(row,col).Value`,
and arithmetic only. File, network, COM, Shell, UI, and control-flow effects
are rejected. For full headless VBA data processing, use the native/Python
runtime shown in the [beginner tutorial](../docs/tutorial-beginners.md).

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
