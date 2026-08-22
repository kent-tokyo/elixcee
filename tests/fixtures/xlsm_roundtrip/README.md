# Real-`.xlsm` slot for the safe-round-trip milestone — superseded

**A real Microsoft-Excel-authored `.xlsm` now exists** (5 of them, as of `0.9.0`'s real-Excel
round-trip validation) under `compat/oracle-excel-com/fixtures/pristine/` — see that
directory's own `README.md` and `compat/oracle-excel-com/results/0.9.0-A_summary.md`.
`tests/xlsx_roundtrip.rs`'s `real_excel_xlsm_roundtrip_preserves_vba_project_and_relationships`
test reads `fixture2_vba_macro.xlsm` directly from there rather than a copy under this
directory — deliberately not duplicated here, to avoid two committed copies of the same
bytes drifting apart.

This directory is kept empty (this README only) as the documented reason the earlier
synthetic-only tests existed in the first place, and as a pointer to where the real thing
actually landed — not as an active "slot waiting to be filled" anymore.

## What `tests/xlsx_roundtrip.rs`'s tests actually cover now

- Three tests (`xlsm_roundtrip_preserves_vba_project_and_declares_macro_enabled_content_types`,
  `xlsx_roundtrip_passes_through_unknown_parts_without_macro_content_type`,
  `xlsm_roundtrip_in_place_save_preserves_vba_project`) still use hand-built, in-test
  synthetic fixtures (`zip::write::ZipWriter`) — kept because a hand-built fixture is more
  reviewable as source than an opaque binary blob, and can pin an exact deterministic byte
  pattern (e.g. a non-trivial, non-all-zero `vbaProject.bin` stand-in) a real file's
  incidental bytes can't guarantee across an Excel version change.
- `real_excel_xlsm_roundtrip_preserves_vba_project_and_relationships` (new as of `0.9.0`)
  runs the same class of assertions against the real fixture instead — including a
  regression guard for the two relationship-carry-over bugs (`xl/theme/theme1.xml`,
  `docProps/{app,core}.xml`) that this exact real fixture found and fixed. See
  `CHANGELOG.md`'s `[0.9.0]`.
- Real Excel itself opening the round-tripped output without a repair prompt is **not**
  something any `cargo test` run can check (no Excel install in CI) — that's what
  `compat/oracle-excel-com/mechanical_check.py` plus the manual real-Excel verification in
  `0.9.0-A_summary.md` covers instead.
