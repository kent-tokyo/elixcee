# Resource limits

This file records the empirical basis for elixcee's own resource limits — each one added
only after measuring a real, confirmed cost on the real oracle (or, for `src/reader.rs`,
a crafted malicious input), never speculatively. See
[`docs/xlsx-security-model.md`](xlsx-security-model.md) for the full limits inventory and
threat model context; this file is where the *sizing* decision for each limit is worked
through and kept up to date as new measurements arrive.

## `packages/xlsx`: `MAX_ITERATED_RANGE_CELLS` (`ELIXCEE_RANGE_TOO_LARGE`)

`sheet_to_formulae`, `sheet_to_csv`, and `sheet_to_txt` (which delegates to `sheet_to_csv`)
all walk every `(row, col)` pair inside a worksheet's `!ref` rectangle regardless of
sparsity. `packages/xlsx/src/internal/range-guard.cjs` rejects ranges above
`MAX_ITERATED_RANGE_CELLS = 5,000,000` cells before iterating them.

### Measurement (2026-08-16, one fixed pass — not repeated)

A sparse worksheet (only 2 populated cells: top-left and bottom-right) with `!ref` sized
to exactly the target cell count, measured in a fresh subprocess per case
(`process.memoryUsage().rss` sampled immediately after the call returns; wall time via
`Date.now()`):

| Range size | Function | Oracle: time | Oracle: RSS | elixcee: time | elixcee: RSS |
|---|---|---|---|---|---|
| 100,000 | `sheet_to_formulae` | 38 ms | 86 MB | 40 ms | 58 MB |
| 100,000 | `sheet_to_csv` | 39 ms | 89 MB | 67 ms | 60 MB |
| 100,000 | `sheet_to_txt` | 36 ms | 90 MB | 72 ms | 69 MB |
| 1,000,000 | `sheet_to_formulae` | 365 ms | 122 MB | 385 ms | 91 MB |
| 1,000,000 | `sheet_to_csv` | 385 ms | 126 MB | 404 ms | 97 MB |
| 1,000,000 | `sheet_to_txt` | 382 ms | 130 MB | 457 ms | 147 MB |
| 5,000,000 | `sheet_to_formulae` | 2,176 ms | 248 MB | not measured (blocked by the guard — see below) | |
| 5,000,000 | `sheet_to_csv` | 2,391 ms | 250 MB | not measured | |
| 5,000,000 | `sheet_to_txt` | 2,168 ms | 260 MB | not measured | |
| 10,000,000 | `sheet_to_formulae` | 4,897 ms | 229 MB | not measured | |
| 10,000,000 | `sheet_to_csv` | 5,950 ms | 329 MB | not measured | |
| 10,000,000 | `sheet_to_txt` | 5,277 ms | 346 MB | not measured | |

elixcee wasn't measured at 5M/10M directly (the guard itself blocks those calls through
the public API — measuring would require bypassing it). At 100K/1M, elixcee's wall time
and RSS are in the same order of magnitude as the oracle's (same O(rows × cols) walk
shape, same language), which is the basis for treating the oracle's 5M/10M numbers as a
reliable proxy for what elixcee's own cost would be at those sizes if unguarded.

### Decision: keep `MAX_ITERATED_RANGE_CELLS = 5,000,000`

At the threshold itself, the cost is ~2.2-2.4s and ~250MB RSS — noticeably slow for a
single synchronous call but not a severe hang, and the point beyond which cost keeps
climbing linearly with no natural ceiling (10M already reaches 5-6s / up to ~345MB, and
the original full-grid probe — `A1:XFD1048576`, ~17.18 billion cells — did not return
within 25s at all). 5,000,000 cells is also far beyond what any realistic populated
worksheet needs for these in-memory JS APIs — the number is chosen to reject
pathologically large ranges specifically, not to constrain normal use. No adjustment from
the original value.

## `src/reader.rs`: `ZIP_ENTRY_MAX_BYTES`

See [`docs/xlsx-security-model.md`](xlsx-security-model.md#existing-limits-as-of-phase-0)
— 64 MB per ZIP entry. Sizing rationale not yet written up separately here; this section
is a placeholder for when that limit's own measurement basis is documented.
