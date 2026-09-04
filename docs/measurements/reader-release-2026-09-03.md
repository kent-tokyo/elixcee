# Release binary reader calibration

Date: 2026-09-03 (Asia/Tokyo)

This local, offline measurement reruns the large-input matrix with the
optimized `--release` CLI binary. It is not a cross-machine performance claim.

## Protocol

- Host: macOS/Darwin, `aarch64-apple-darwin`
- Rust: `rustc 1.97.0`
- Binary: `target/release/elixcee`
- Profiles: dense, dense-5col, sparse, style-heavy, formula-heavy, mixed,
  multi-sheet-4
- Repetitions: 3 per profile, sequential
- Correctness: the same ZIP/XML structure and snapshot cell-count checks as the
  debug calibration

Raw observations are in `reader-release-2026-09-03.json`.

## Summary

| profile | cells | median wall | max RSS | outcome |
|---|---:|---:|---:|---|
| dense | 400,000 | 1.758 s | 86.3 MiB | 3/3 correct |
| dense-5col | 400,000 | 1.500 s | 69.6 MiB | 3/3 correct |
| sparse | 100,000 | 0.505 s | 19.6 MiB | 3/3 correct |
| style-heavy | 200,000 | 0.925 s | 40.9 MiB | 3/3 correct |
| formula-heavy | 240,000 | 1.111 s | 62.6 MiB | 3/3 correct |
| mixed | 320,000 | 1.173 s | 92.2 MiB | 3/3 correct |
| multi-sheet-4 | 400,000 | 1.396 s | 67.8 MiB | 3/3 correct |

The measurements calibrate the release binary on this host and fixture set;
they are not a debug-versus-release benchmark claim or a general throughput
claim. Cross-platform, independent oracle, and repeated-process reclamation
remain separate gates.
