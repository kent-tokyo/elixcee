# 0.20.0 Streaming API

This milestone adds a separate forward-only row pipeline. `Vm` remains the complete,
random-access workbook model; `StreamReader` owns a worksheet ZIP entry in a worker and
emits one row at a time through a bounded channel. `StreamWriter` accepts rows through an
append-only API and emits a single-sheet XLSX on close.

The first slice deliberately excludes formulas, VBA, styles, merges, charts, and arbitrary
cell mutation. Empty cells inside a row are represented as `None` at the Python boundary.
ZIP validation from 0.19.0 is applied before any worksheet bytes are consumed.
