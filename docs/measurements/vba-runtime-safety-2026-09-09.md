# VBA runtime safety local verification

Date: 2026-09-09 (Asia/Tokyo)

This record covers the local safety boundary for VBA workbook `Save` and
`Close` members and the bounded UDT-resolution diagnostics. It does not
implement persistence or claim Excel semantic equivalence.

## Scope

- `ThisWorkbook.Save` and `ThisWorkbook.Close` are rejected by default during
  headless execution as blocked external effects.
- The failure remains a human-readable `SECURITY:` runtime error and is also
  exposed as the structured `SecurityBlockedExternalEffect` category.
- The VM records `SecurityBlockedExternalEffect` and `MsgBoxBlocked` at the
  error origin; message-pattern classification is retained only as a fallback
  for older or not-yet-migrated runtime paths.
- CLI JSON diagnostics prefer the VM category and expose it as `E1011`.
- `On Error Resume Next` cannot suppress a blocked external effect.

## Reproduction

From the repository root, run:

```text
cargo test workbook_save_and_close_are_blocked_as_external_effects --offline -- --nocapture
cargo test type_collisions --offline -- --nocapture
cargo test run_sub_multi_rejects_a_genuine_type_collision_before_binding --offline -- --nocapture
cargo test module_qualified_udt_resolution --offline -- --nocapture
cargo test --workspace --all-targets --offline --quiet
cargo clippy --workspace --all-targets --offline -- -D warnings
cargo fmt --all -- --check
git diff --check
python3 scripts/check-formula-dispatch.py --check-contracts --check-docs
python3 scripts/check-ooxml-feature-matrix.py
```

## Result

The targeted safety, UDT-collision, and module-qualified UDT tests passed. The
full local Rust workspace passed 1,613 tests; strict clippy, formatting, diff,
formula-dispatch, and OOXML matrix checks also passed on macOS arm64.

## Not measured or claimed

- Actual output-path selection, persistence, `Close` post-state, or event
  dispatch.
- Excel reopen, repair-warning absence, or Excel oracle agreement.
- Linux/Windows behavior, long-running isolation, or external review.
