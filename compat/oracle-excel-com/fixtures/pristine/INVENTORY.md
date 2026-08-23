# Fixture inventory (0.10.0-A, step 1)

**Purpose**: a fixture's filename is a description written at authoring time, not a
verified manifest — `fixture5_chart_image_freeze_print.xlsm`'s own name claims freeze
panes it does not contain (see `docs/xlsx-worksheet-preservation-0.10.0-design.md`, §3).
This file replaces filename-trust with a script-generated scan of each fixture's actual
ZIP parts and worksheet/workbook XML, so 0.10.0-A's negative self-tests and the
relationship-type mapping table (step 2, `mechanical_check.py`) are built against what is
actually here, not what a name implies.

**Regeneration**: produced by unzipping each `.xlsm` and regex-scanning
`xl/worksheets/sheet1.xml` / `xl/workbook.xml` for the tags below, plus a `namelist()` scan
for parts (`xl/vbaProject.bin`, `xl/printerSettings/*`, `xl/embeddings/*` for OLE). Re-run
if a fixture is ever replaced or a new one added — this file is not auto-checked against
the fixtures, so it can go stale like any other manifest.

## fixture1_values_styles_merge_hidden.xlsm
- Parts: workbook.xml, worksheets/sheet1.xml, theme, styles, sharedStrings, calcChain — no VBA.
- Worksheet features present: `<sheetViews>`, `<mergeCells>`, hidden row/col (`hidden="1"`).
- No table/validation/conditionalFormatting/hyperlinks/drawing/legacyDrawing/pageSetup/pane/definedNames.

## fixture2_vba_macro.xlsm
- Parts: workbook.xml, worksheets/sheet1.xml, theme, styles, sharedStrings, **vbaProject.bin**.
- Worksheet features present: `<sheetViews>` only.
- No table/validation/conditionalFormatting/hyperlinks/drawing/legacyDrawing/pageSetup/pane/definedNames.

## fixture3_table_validation_conditional.xlsm
- Parts: workbook.xml, worksheets/sheet1.xml (+ `worksheets/_rels/sheet1.xml.rels`), theme,
  styles, sharedStrings, vbaProject.bin, **tables/table1.xml**.
- Worksheet features present: `<tableParts>`, `<dataValidations>`, `<conditionalFormatting>`,
  `<sheetViews>`.
- This is the fixture used for the design round's original empirical `SOURCE_REFERENCE_LOSS`
  proof (`xl/worksheets/_rels/sheet1.xml.rels` → `../tables/table1.xml` survives a save, but
  the regenerated `sheet1.xml` never re-emits `<tableParts><tablePart r:id="rId1"/></tableParts>`).
  **Confirmed to reproduce identically in fixture4 (hyperlink + vmlDrawing, both lost in the
  same save) and fixture5 (drawing lost)** once `mechanical_check.py`'s
  `check_source_references()` existed to check for it — every fixture with a worksheet-level
  relationship at all shows this bug; fixture1/fixture2 (no worksheet-level relationships)
  correctly report `CLEAN`.
- No hyperlinks/drawing/legacyDrawing/pageSetup/pane/definedNames.

## fixture4_hyperlink_comment_name.xlsm
- Parts: workbook.xml, worksheets/sheet1.xml (+ `.rels`), theme, styles, sharedStrings,
  vbaProject.bin, `drawings/vmlDrawing1.vml`, **comments1.xml**,
  `threadedComments/threadedComment1.xml`, `persons/person.xml`.
- Worksheet features present: `<conditionalFormatting>`, `<hyperlinks>` (**`r:id` form only**
  — `Target="https://yahoo.co.jp/" TargetMode="External"`, no `location`-attribute example),
  `<legacyDrawing>`, `<sheetViews>`.
- Workbook-level: `<definedNames>` — one ordinary named range (`name="test"`,
  `Sheet1!$F$5`, carries a `comment` attribute), **not** a builtin `_xlnm.*` name.
- No table/dataValidations/drawing/pageSetup/pane.
- `location`-attribute (same-workbook, no-relationship) hyperlink: **not present in any
  fixture** — confirmed absent, not merely unchecked.

## fixture5_chart_image_freeze_print.xlsm
- Parts: workbook.xml, worksheets/sheet1.xml (+ `.rels`), theme, styles, sharedStrings,
  `media/image1.png`, `drawings/drawing1.xml` (+ `.rels`), `charts/chart1.xml` (+ `.rels`),
  `charts/style1.xml`, `charts/colors1.xml`, `metadata.xml`, `richData/*` (4 parts + `.rels`).
  No VBA project.
- Worksheet features present: `<drawing r:id>` (chart/image), `<pageSetup>`, `<sheetViews>`.
- **`<pane>` (freeze pane): confirmed absent**, despite the filename — re-confirmed here
  with the same regex scan used for every other fixture, not a one-off `grep`.
- Workbook-level: `<definedNames>` — but this one is the **builtin** `_xlnm.Print_Area`
  (`localSheetId="0"`, `Sheet1!$E$3`), a *different* case from fixture4's ordinary named
  range. Both belong to 0.10.0-C ("print area・print titles" / "defined names" are the
  same XML container, `<definedNames>`, but a builtin `_xlnm.*` name has spreadsheet-app
  semantics an ordinary name doesn't — worth keeping distinct in any future structured
  handling, even though the opaque-fragment-passthrough design in the main doc's §7 treats
  both the same way at the XML level).
- One cell carries `t="e" vm="1"` — a rich-value-linked error cell (the `richData` feature,
  already declared out of scope in the main design doc's §3).

## What is confirmed absent across all 5 fixtures (not merely unchecked)

- `<pane>` (freeze pane / split view)
- `<hyperlink location="...">` (same-workbook, relationship-free hyperlink)
- `xl/printerSettings/*` (printer settings binary + its `<pageSetup r:id>` backing)
- Any OLE object / ActiveX control part (`xl/embeddings/*`, `<oleObjects>`/`<controls>`)
- `<dataValidations>` with a source **outside** the current sheet, or a `<dataValidation>`
  whose `type` differs from fixture3's `list`
- `<autoFilter>` as a standalone worksheet element (fixture3's table has its own
  `<autoFilter>` *inside* `xl/tables/table1.xml`, which is a different part, already
  passthrough-safe today per the main design doc's §5)

Per the main design doc's hard gate (§10): **no writer code for any of the above may be
written until a real fixture containing it exists.** This is 0.10.0-A prerequisite work,
not yet done as of this file.
