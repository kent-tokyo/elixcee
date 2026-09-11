# In-process reader soak and allocator measurement

Date: 2026-09-03 (Asia/Tokyo)

This local, offline measurement repeatedly calls the Rust buffer reader in one
process. It is distinct from the earlier CLI child-process reclamation test.

## Protocol

- Host: macOS/Darwin, `aarch64-apple-darwin`
- Rust: `rustc 1.97.0`
- Binary: `target/debug/measure_reader_inprocess`
- API: `elixcee::reader::read_workbook_from_bytes`
- Input bytes are loaded once; each returned `BufferWorkbook` is dropped before
  the next iteration
- Profiles: `dense`, `mixed`, and `multi-sheet-4`
- 10 iterations per profile, 30 API calls total
- macOS default malloc-zone statistics are sampled after each drop

Raw observations are in `reader-inprocess-soak-2026-09-03.json`.

## Summary

| profile | iterations | cells/call | total cells | total wall | allocator size-in-use range | allocator allocated range |
|---|---:|---:|---:|---:|---:|---:|
| dense | 10/10 | 400,000 | 4,000,000 | 136.5 s | 15.9–15.9 MiB | 31.9 MiB |
| mixed | 10/10 | 320,000 | 3,200,000 | 119.5 s | 16.2–16.2 MiB | 40.1 MiB |
| multi-sheet-4 | 10/10 | 400,000 | 4,000,000 | 123.5 s | 15.5–15.5 MiB | 39.4 MiB |

All 30 in-process calls returned the expected cell count and completed without
an error. Within each profile, allocator size-in-use varied by at most 3,456
bytes after the workbook was dropped. The allocator's allocated total stayed
constant within each profile. This is evidence of stable post-drop allocator
observations on this host, not a proof that all allocators or workloads are
leak-free. Longer soak, allocator instrumentation on other platforms, and
in-process API loops with mutation/write paths remain separate gates.
