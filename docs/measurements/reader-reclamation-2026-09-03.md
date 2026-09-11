# Repeated reader process reclamation measurement

Date: 2026-09-03 (Asia/Tokyo)

This local, offline measurement checks sequential CLI child-process termination
and resource drift. It is not an in-process allocator or leak-proof claim.

## Protocol

- Host: macOS/Darwin, `aarch64-apple-darwin`
- Rust: `rustc 1.97.0`
- Binary: `target/debug/elixcee`
- Profiles: `dense`, `mixed`, and `multi-sheet-4`
- Repetitions: 10 sequential child processes per profile, 30 total
- Each child writes JSON to a temporary file, is waited on to exit, and only
  then is the next child started
- Recorded: wall time, child peak RSS, exact snapshot cell count, and parent
  measurement-process RSS before/after each profile

Raw observations are in `reader-reclamation-2026-09-03.json`.

## Summary

| profile | repetitions | correctness | child peak RSS range | median wall | parent RSS delta |
|---|---:|---|---:|---:|---:|
| dense | 10 | 10/10 | 69.6–71.6 MiB | 15.3 s | +8.1 MiB |
| mixed | 10 | 10/10 | 76.5–83.9 MiB | 14.6 s | +7.8 MiB |
| multi-sheet-4 | 10 | 10/10 | 49.8–54.1 MiB | 14.6 s | +9.8 MiB |

All 30 child processes exited successfully; no child remained running between
iterations and no timeout or malformed JSON occurred. The parent RSS deltas
are observed process-level drift in the Python harness and OS allocator/cache,
not evidence of a Rust leak or proof of its absence. A future in-process API
loop, allocator instrumentation, and longer soak run are separate gates.
