# 0.9.0-A real-Excel round-trip fixtures

`pristine/` -- real Microsoft-Excel-authored `.xlsx`/`.xlsm` files, committed as-is. Never
written to by any test or script; every run copies one into `work/` first
(`mechanical_check.py`'s callers are expected to do this, not open a pristine file directly
-- an in-place-save test would otherwise silently consume a fixture that cost real Excel
authoring effort to create).

`work/` -- gitignored scratch copies elixcee reads/edits/saves into. Safe to delete anytime;
regenerate by copying from `pristine/` again.

## Fixtures

- `fixture1_values_styles_merge_hidden.xlsm` -- values, a formula (`A3 = A2*2`), a number
  format, a merged range (`B1:C1`), a hidden column (D), a hidden row (5). No VBA project
  (`has vb project` is false) -- this fixture isolates the data/style/layout preservation
  question from the VBA-project-passthrough question, which fixture 2 (not yet authored)
  covers instead. Authored live in Microsoft Excel for Mac (16.108) via AppleScript,
  confirmed via direct ZIP/XML inspection to be genuine Excel-producer output (`xr:uid`,
  `calcChain.xml`, `theme1.xml`, etc.), not hand-built.

See `ROADMAP.md`'s `0.9.0-A` section and `compat/oracle-excel-com/mechanical_check.py` for
how these are used.
