# elixcee Playground

Try elixcee in the browser: switch between sales, budget, and grade samples;
edit a small table; recalculate `SUM` or `AVERAGE` with the Rust/WASM engine;
and download the resulting `.xlsx` file. The page is client-only; workbook
bytes are not uploaded.

Open the live [GitHub Pages playground](https://kent-tokyo.github.io/elixcee/playground/)
or read the [English quick start](../docs/quickstart.md) first.
Use the in-page language selector to switch the interface and quick-start link
between English, Japanese, and Simplified Chinese; translations are not shown
side by side.

The browser page covers the private `@elixcee/xlsx` browser package: in-memory
XLSX read/write and typed workbook edits. It does not run Python or VBA in a
browser. For headless VBA data processing, use the native/Python runtime shown
in the [beginner tutorial](../docs/tutorial-beginners.md).

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
