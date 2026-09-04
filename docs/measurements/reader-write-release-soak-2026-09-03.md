# Release in-process read/mutate/write soak

Date: 2026-09-03 (Asia/Tokyo)

This local, offline measurement exercises the Rust read and write path in one
release process. It is not an Excel compatibility or all-allocator leak proof.

## Protocol

- Host: macOS/Darwin, `aarch64-apple-darwin`
- Rust: `rustc 1.97.0`
- Binary: `target/release/measure_reader_write_inprocess`
- Per iteration: `Vm::load_workbook_file` → `write_rect` →
  `set_cell_formula` → `save_workbook` → `reader::read_workbook` verification
- Profiles: dense, mixed, multi-sheet-4
- 5 iterations per profile, 15 read/mutate/write cycles total
- Each VM, reloaded workbook, and temporary output is dropped/removed before
  the next iteration; macOS malloc statistics are sampled after cleanup

Raw observations are in `reader-write-release-soak-2026-09-03.json`.

## Summary

| profile | cycles | cells/cycle | total wall | allocator size-in-use range | allocated range | outcome |
|---|---:|---:|---:|---:|---:|---|
| dense | 5 | 400,000 | 19.4 s | 139.5–140.1 KiB | 24.0 MiB | 5/5 round-tripped |
| mixed | 5 | 320,000 | 18.0 s | 187.5–188.1 KiB | 32.0–36.0 MiB | 5/5 round-tripped |
| multi-sheet-4 | 5 | 400,000 | 20.2 s | 139.5–140.1 KiB | 32.0 MiB | 5/5 round-tripped |

All 15 mutations were observed after save/reload: the integer write matched
the iteration number and the formula cell survived. Post-cleanup allocator
size-in-use remained within 640 bytes per profile. The allocated-total value
is allocator behavior, not live-object usage; other allocators, platforms,
longer soaks, and richer mutation operations remain unmeasured.
