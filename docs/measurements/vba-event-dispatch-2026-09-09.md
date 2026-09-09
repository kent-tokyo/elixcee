# VBA event dispatch local verification

Date: 2026-09-09 (Asia/Tokyo)

This record covers the bounded event-dispatch slice. It is not an Excel
event-compatibility claim.

## Scope

- `Vm.run_event` and Python `Vm.run_event` dispatch one explicitly named,
  zero-argument handler from a parsed program.
- `Vm.run_sub_with_events` and Python `Vm.run_with_events` opt in to running
  `Workbook_Open` before the selected entrypoint; a failing open handler stops
  the entrypoint.
- `Vm.run_sub_multi_with_events` applies the same opt-in behavior across
  multiple standard modules and rejects duplicate `Workbook_Open` handlers;
  same-program duplicates are rejected as well.
- `Vm.run_worksheet_change` binds an explicit A1 range on the active sheet to
  a single `As Range` parameter for `Worksheet_Change`.
- The bound target exposes `Value`, `Address`, `Row`, `Column`, and the
  single-area `Rows.Count` / `Columns.Count` members.
- `run_sub_with_events` and `run_sub_multi_with_events` automatically dispatch
  a unique handler after supported VBA value/formula writes. Handler writes
  are queued and bounded to 64 cumulative dispatches.
- Supported names are `Workbook_Open`, `Workbook_BeforeClose`,
  `Worksheet_Change`, `Worksheet_Calculate`, and `Worksheet_SelectionChange`.
- `Application.EnableEvents = False` suppresses dispatch, and a handler cannot
  recursively re-enter the event dispatcher.
- The normal instruction, call-depth, string, array, and cell budgets remain
  active during dispatch.

## Reproduction

From the repository root, run:

```text
cargo test explicit_event_dispatch --offline -- --nocapture
cargo test run_sub_multi_with_events --offline -- --nocapture
cargo test worksheet_change --offline -- --nocapture
cargo test --workspace --all-targets --offline --quiet
cargo clippy --workspace --all-targets --offline -- -D warnings
cargo fmt --all -- --check
git diff --check
```

## Result

The target metadata regression and event regression tests passed. The full
Rust target suite passed 1,663 tests on macOS arm64, including blackbox and
benchmark targets. Strict clippy and diff checks passed.

## Not measured or claimed

- Automatic event discovery after workbook load, beyond the opt-in event
  runner.
- Worksheet-scoped selection among multiple `Worksheet_Change` handlers.
- Excel reopen, Excel oracle agreement, other OS behavior, or external review.
