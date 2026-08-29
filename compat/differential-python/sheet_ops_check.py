#!/usr/bin/env python3
"""Differential check: elixcee's sheet management + workbook-metadata API
(P1 core 3 + remainder + P2 hidden row/col + copy_sheet + defined_names +
sheet_state + row height/column width) vs. openpyxl.

openpyxl is used PURELY as a test-only oracle here -- never a runtime dependency
of the shipped `elixcee` package (see pyproject.toml, which declares none).
Requires a `maturin develop --features python` build of elixcee and
`pip install openpyxl` in the active environment; see this directory's
README.md for the one-time setup.

Compares rename_sheet()/move_sheet()/copy_sheet()/merged_cells()/
merge_cells()/unmerge_cells()/hidden_rows()/hidden_columns()/
set_row_hidden()/set_column_hidden()/defined_names() against openpyxl's own
read of the same real fixture after a save/reload round trip. Must build
from FIXTURE via elixcee.load_workbook, not a bare elixcee.Vm() -- a
from-scratch VM's minimal styles.xml emits a bare <fill/> that openpyxl's
own reader rejects on reopen (a real, pre-existing, unrelated bug -- see
ROADMAP.md's known gaps).

sheet_state() coverage uses an openpyxl-AUTHORED fixture, not a real
Excel-authored one under compat/oracle-excel-com (none has a hidden/
veryHidden sheet), and reads only -- no elixcee save() round trip, since
elixcee has no writer support for `state="..."` yet.

row_height()/column_width() coverage also uses an openpyxl-AUTHORED fixture
(no real fixture has a genuine custom row height or column width), reads
only, same no-writer-support caveat -- a loaded row height/column width is
dropped on EVERY elixcee save today, not just sometimes (the writer
regenerates <row>/<cols> from hidden-row/column state alone).

defined_names() coverage uses FIXTURE4 (fixture1 has no <definedNames> at
all), and reads directly rather than through a save/reload round trip --
it's a passthrough-only feature this round doesn't create/modify.

Row/col insert-delete is deliberately given NO differential coverage here:
the disclosed fidelity gap (merges/hidden markers/styles/formats not shifted)
means an openpyxl comparison would correctly fail on exactly the cases worth
testing -- Rust unit + integration tests already cover the values-only
behavior that IS correct.

sort_range() also gets NO differential coverage: openpyxl has no sort
primitive of its own to compare against. Its PyO3-layer bound checks are
pinned directly (no openpyxl comparison needed) in
SortRangeAndMergeCellsRejectOversizedOrInvalidInput below.

Run standalone: `python3 compat/differential-python/sheet_ops_check.py`
"""

from __future__ import annotations

import os
import tempfile
import unittest
import zipfile

import elixcee
import openpyxl

FIXTURE = os.path.join(
    os.path.dirname(__file__),
    "..",
    "oracle-excel-com",
    "fixtures",
    "pristine",
    "fixture1_values_styles_merge_hidden.xlsm",
)

# The one real Excel-authored fixture with genuine <definedNames> content --
# fixture1 has none, so defined_names() coverage needs this one instead.
FIXTURE4 = os.path.join(
    os.path.dirname(__file__),
    "..",
    "oracle-excel-com",
    "fixtures",
    "pristine",
    "fixture4_hyperlink_comment_name.xlsm",
)


class RenameSheetAgreesWithOpenpyxl(unittest.TestCase):
    def test_rename_sheet_agrees_with_openpyxl_after_a_round_trip(self):
        vm = elixcee.load_workbook(FIXTURE)
        vm.rename_sheet("Sheet1", "Renamed")

        with tempfile.TemporaryDirectory() as d:
            path = os.path.join(d, "out.xlsx")
            vm.save_workbook(path)

            reloaded_vm = elixcee.load_workbook(path)
            wb = openpyxl.load_workbook(path)

            self.assertEqual(wb.sheetnames, ["Renamed"])
            # elixcee's own sheet_names() convention is lowercase, unlike
            # openpyxl's display-cased sheetnames.
            self.assertEqual(reloaded_vm.sheet_names(), ["renamed"])

            ws = wb["Renamed"]
            self.assertEqual(
                [str(r) for r in ws.merged_cells.ranges],
                reloaded_vm.merged_cells(sheet="Renamed"),
            )


class CopySheetAgreesWithOpenpyxl(unittest.TestCase):
    def test_copy_sheet_agrees_with_openpyxl_after_a_round_trip(self):
        vm = elixcee.load_workbook(FIXTURE)
        vm.copy_sheet("Sheet1", "Copy")

        with tempfile.TemporaryDirectory() as d:
            path = os.path.join(d, "out.xlsx")
            vm.save_workbook(path)

            reloaded_vm = elixcee.load_workbook(path)
            wb = openpyxl.load_workbook(path)

            self.assertEqual(wb.sheetnames, ["Sheet1", "Copy"])
            # elixcee's own sheet_names() is alphabetically sorted, NOT
            # sheet_order/tab-position order (a pre-existing, undocumented
            # quirk unrelated to copy_sheet -- discovered while writing this
            # test). "copy" < "sheet1" alphabetically, so this is reversed
            # from wb.sheetnames' tab-position order above.
            self.assertEqual(reloaded_vm.sheet_names(), ["copy", "sheet1"])

            # The copy must carry the source's merge and hidden row/column
            # state -- both elixcee's own reload and openpyxl's independent
            # read of the same saved file must agree on it.
            ws_copy = wb["Copy"]
            self.assertEqual(
                [str(r) for r in ws_copy.merged_cells.ranges],
                reloaded_vm.merged_cells(sheet="Copy"),
            )
            self.assertIn("B1:C1", reloaded_vm.merged_cells(sheet="Copy"))
            self.assertIn(5, reloaded_vm.hidden_rows(sheet="Copy"))
            self.assertTrue(ws_copy.row_dimensions[5].hidden)


class MoveSheetAgreesWithOpenpyxl(unittest.TestCase):
    def test_move_sheet_to_the_front_matches_openpyxls_own_sheetnames(self):
        vm = elixcee.load_workbook(FIXTURE)
        vm.set_sheet("Second")  # creates a second, empty sheet
        vm.move_sheet("Second", 0)

        with tempfile.TemporaryDirectory() as d:
            path = os.path.join(d, "out.xlsx")
            vm.save_workbook(path)
            wb = openpyxl.load_workbook(path)
            self.assertEqual(wb.sheetnames[0], "Second")


class MergedCellsAgreesWithOpenpyxl(unittest.TestCase):
    def test_merged_cells_matches_openpyxl_on_the_real_fixture(self):
        vm = elixcee.load_workbook(FIXTURE)
        wb = openpyxl.load_workbook(FIXTURE)
        ws = wb.active
        self.assertEqual(
            vm.merged_cells(), [str(r) for r in ws.merged_cells.ranges]
        )


class MergeCellsAndUnmergeCellsAgreeWithOpenpyxl(unittest.TestCase):
    def test_a_newly_created_merge_matches_openpyxl_after_a_round_trip(self):
        vm = elixcee.load_workbook(FIXTURE)
        vm.merge_cells("D1:E1")  # non-overlapping with fixture1's pre-existing B1:C1

        with tempfile.TemporaryDirectory() as d:
            path = os.path.join(d, "out.xlsx")
            vm.save_workbook(path)

            reloaded_vm = elixcee.load_workbook(path)
            wb = openpyxl.load_workbook(path)
            ws = wb.active
            self.assertEqual(
                sorted(reloaded_vm.merged_cells()),
                sorted(str(r) for r in ws.merged_cells.ranges),
            )
            self.assertIn("D1:E1", reloaded_vm.merged_cells())

    def test_removing_the_fixtures_pre_existing_merge_matches_openpyxl(self):
        vm = elixcee.load_workbook(FIXTURE)
        vm.unmerge_cells("B1:C1")

        with tempfile.TemporaryDirectory() as d:
            path = os.path.join(d, "out.xlsx")
            vm.save_workbook(path)

            reloaded_vm = elixcee.load_workbook(path)
            wb = openpyxl.load_workbook(path)
            ws = wb.active
            self.assertEqual(reloaded_vm.merged_cells(), [])
            self.assertEqual(list(ws.merged_cells.ranges), [])


class HiddenRowsAndColumnsAgreeWithOpenpyxl(unittest.TestCase):
    # openpyxl's ws.row_dimensions/column_dimensions are sparse dicts,
    # populated only for rows/columns it actually parsed a <row>/<col>
    # element for -- these tests check agreement on the fixture's known
    # hidden units specifically, not an assumption that openpyxl reports
    # every hidden row/column elixcee does.
    def test_the_fixtures_pre_existing_hidden_row_and_column_match_openpyxl(self):
        vm = elixcee.load_workbook(FIXTURE)
        wb = openpyxl.load_workbook(FIXTURE)
        ws = wb.active

        self.assertIn(5, vm.hidden_rows())
        self.assertTrue(ws.row_dimensions[5].hidden)

        self.assertIn(4, vm.hidden_columns())  # column D
        self.assertTrue(ws.column_dimensions["D"].hidden)

    def test_a_newly_hidden_row_and_column_match_openpyxl_after_a_round_trip(self):
        vm = elixcee.load_workbook(FIXTURE)
        vm.set_row_hidden(20)
        vm.set_column_hidden(6)  # column F

        with tempfile.TemporaryDirectory() as d:
            path = os.path.join(d, "out.xlsx")
            vm.save_workbook(path)

            reloaded_vm = elixcee.load_workbook(path)
            wb = openpyxl.load_workbook(path)
            ws = wb.active

            self.assertIn(20, reloaded_vm.hidden_rows())
            self.assertTrue(ws.row_dimensions[20].hidden)

            self.assertIn(6, reloaded_vm.hidden_columns())
            self.assertTrue(ws.column_dimensions["F"].hidden)

    def test_unhiding_the_fixtures_pre_existing_hidden_row_matches_openpyxl(self):
        vm = elixcee.load_workbook(FIXTURE)
        vm.set_row_hidden(5, hidden=False)

        with tempfile.TemporaryDirectory() as d:
            path = os.path.join(d, "out.xlsx")
            vm.save_workbook(path)

            reloaded_vm = elixcee.load_workbook(path)
            wb = openpyxl.load_workbook(path)
            ws = wb.active

            self.assertNotIn(5, reloaded_vm.hidden_rows())
            # A <row> element with no hidden attribute at all is possible
            # (the row still exists for its cell data), so check the
            # attribute's truthiness rather than the row's mere presence.
            self.assertFalse(bool(ws.row_dimensions[5].hidden))


class DefinedNamesAgreesWithOpenpyxl(unittest.TestCase):
    def test_defined_names_matches_openpyxl_on_the_real_fixture(self):
        # FIXTURE4, not FIXTURE -- fixture1 has no <definedNames> at all.
        vm = elixcee.load_workbook(FIXTURE4)
        wb = openpyxl.load_workbook(FIXTURE4)

        elixcee_names = vm.defined_names()
        openpyxl_names = {name: dn.value for name, dn in wb.defined_names.items()}
        self.assertEqual(elixcee_names, openpyxl_names)
        self.assertEqual(elixcee_names.get("test"), "Sheet1!$F$5")

    def test_defined_names_is_empty_on_a_fixture_with_none(self):
        vm = elixcee.load_workbook(FIXTURE)
        wb = openpyxl.load_workbook(FIXTURE)
        self.assertEqual(vm.defined_names(), {})
        self.assertEqual(dict(wb.defined_names), {})


class SheetStateAgreesWithOpenpyxl(unittest.TestCase):
    # No real fixture in this repo has a hidden/veryHidden sheet (see
    # docs/openpyxl-gap-audit.md), so the fixture here is built WITH
    # openpyxl itself rather than loaded from compat/oracle-excel-com --
    # this compares reads only, no elixcee save() round trip involved.
    def test_sheet_state_matches_openpyxl_on_an_openpyxl_authored_fixture(self):
        with tempfile.TemporaryDirectory() as d:
            path = os.path.join(d, "states.xlsx")
            wb = openpyxl.Workbook()
            wb.active.title = "Visible"
            wb.create_sheet("Hidden").sheet_state = "hidden"
            wb.create_sheet("VeryHidden").sheet_state = "veryHidden"
            wb.save(path)

            vm = elixcee.load_workbook(path)
            wb2 = openpyxl.load_workbook(path)

            for name in ("Visible", "Hidden", "VeryHidden"):
                self.assertEqual(vm.sheet_state(name), wb2[name].sheet_state)

    def test_sheet_state_does_not_yet_survive_an_elixcee_save(self):
        # Disclosed known gap, not a hypothetical: elixcee has no writer
        # support for `state="..."` yet (no real fixture to validate the
        # writer shape against -- see ROADMAP.md's known gaps), so a loaded
        # hidden sheet currently reverts to visible on ANY save, even a
        # no-op one. Pinned here so a future writer fix is a deliberate,
        # visible change to this test, not a silent behavior shift.
        with tempfile.TemporaryDirectory() as d:
            src = os.path.join(d, "states_src.xlsx")
            wb = openpyxl.Workbook()
            wb.active.title = "Visible"
            wb.create_sheet("Hidden").sheet_state = "hidden"
            wb.save(src)

            vm = elixcee.load_workbook(src)
            self.assertEqual(vm.sheet_state("Hidden"), "hidden")

            out = os.path.join(d, "states_out.xlsx")
            vm.save_workbook(out)
            wb2 = openpyxl.load_workbook(out)
            self.assertEqual(
                wb2["Hidden"].sheet_state,
                "visible",
                "known gap: update this test (and ship set_sheet_state) once "
                "the writer preserves state=... on save",
            )


class RowHeightAndColumnWidthAgreeWithOpenpyxl(unittest.TestCase):
    # No real fixture in this repo has a genuine custom row height or column
    # width (fixture1's only <col> is a hidden column with width="0", not
    # real data) -- fixture built with openpyxl itself, same approach as
    # SheetStateAgreesWithOpenpyxl. Reads only, no elixcee save() round trip.
    def test_row_height_and_column_width_match_openpyxl_on_an_openpyxl_authored_fixture(
        self,
    ):
        with tempfile.TemporaryDirectory() as d:
            path = os.path.join(d, "dims.xlsx")
            wb = openpyxl.Workbook()
            ws = wb.active
            ws["A1"] = 1
            ws.row_dimensions[5].height = 30.5
            ws.column_dimensions["B"].width = 12.5
            wb.save(path)

            vm = elixcee.load_workbook(path)
            wb2 = openpyxl.load_workbook(path)
            ws2 = wb2.active

            self.assertEqual(vm.row_height(5), ws2.row_dimensions[5].height)
            self.assertEqual(vm.column_width(2), ws2.column_dimensions["B"].width)
            # A row with no explicit height is None on both sides. Column is
            # NOT compared the same way: openpyxl's column_dimensions[letter]
            # auto-vivifies a default-width (13.0) entry on first `[]` access
            # even for a column the file never set -- an openpyxl
            # implementation artifact, not something elixcee's own correct
            # None-for-unset behavior should be judged against.
            self.assertIsNone(vm.row_height(1))
            self.assertIsNone(ws2.row_dimensions[1].height)
            self.assertIsNone(vm.column_width(1))

    def test_row_height_and_column_width_survive_an_elixcee_save(self):
        # Was a disclosed known gap (see this test's old name/history): the
        # writer used to unconditionally regenerate <row>/<cols> from
        # hidden-row/column state alone, dropping a loaded row height/column
        # width on EVERY save. Fixed -- see CHANGELOG.md. write support
        # (set_row_height/set_column_width) is still deferred separately (no
        # real fixture has genuine data to validate that API's own writer
        # shape against, a different concern from preserving an
        # already-loaded value through a save).
        with tempfile.TemporaryDirectory() as d:
            src = os.path.join(d, "dims_src.xlsx")
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.row_dimensions[5].height = 30.5
            ws.column_dimensions["B"].width = 12.5
            wb.save(src)

            vm = elixcee.load_workbook(src)
            self.assertEqual(vm.row_height(5), 30.5)
            self.assertEqual(vm.column_width(2), 12.5)

            out = os.path.join(d, "dims_out.xlsx")
            vm.save_workbook(out)
            wb2 = openpyxl.load_workbook(out)
            ws2 = wb2.active
            self.assertEqual(ws2.row_dimensions[5].height, 30.5)
            self.assertEqual(ws2.column_dimensions["B"].width, 12.5)

            vm2 = elixcee.load_workbook(out)
            self.assertEqual(vm2.row_height(5), 30.5)
            self.assertEqual(vm2.column_width(2), 12.5)

            # A second save must not re-drop or duplicate the attributes.
            out2 = os.path.join(d, "dims_out2.xlsx")
            vm2.save_workbook(out2)
            vm3 = elixcee.load_workbook(out2)
            self.assertEqual(vm3.row_height(5), 30.5)
            self.assertEqual(vm3.column_width(2), 12.5)


class AutoFilterSurvivesAnElixceeSave(unittest.TestCase):
    # Was a disclosed known gap (ROADMAP.md known-gap 28, found while scoping
    # 0.14.0-C): `<autoFilter>` has no `r:id` at all, unlike tableParts/drawing, so
    # it needed no rels_survived gate -- but it was ALSO not one of the
    # unconditionally-restored opaque fragments (sheetFormatPr/dataValidations/
    # pageMargins/...), so it was silently destroyed on every elixcee save, not
    # merely stale. Fixed as a byte-preservation-only passthrough fragment -- see
    # CHANGELOG.md. elixcee has no auto_filter read/write API of its own (this is
    # preservation only, not new feature support -- create/remove/filter-type API
    # is 0.16.0), so this reads back through openpyxl on both sides rather than
    # comparing against an elixcee-side getter.
    def test_autofilter_ref_survives_an_elixcee_save(self):
        with tempfile.TemporaryDirectory() as d:
            src = os.path.join(d, "af_src.xlsx")
            wb = openpyxl.Workbook()
            ws = wb.active
            ws["A1"] = "h"
            ws.auto_filter.ref = "A1:C10"
            wb.save(src)

            vm = elixcee.load_workbook(src)
            out = os.path.join(d, "af_out.xlsx")
            vm.save_workbook(out)

            wb2 = openpyxl.load_workbook(out)
            self.assertEqual(wb2.active.auto_filter.ref, "A1:C10")

            # A second save must not re-drop or duplicate it.
            vm2 = elixcee.load_workbook(out)
            out2 = os.path.join(d, "af_out2.xlsx")
            vm2.save_workbook(out2)
            wb3 = openpyxl.load_workbook(out2)
            self.assertEqual(wb3.active.auto_filter.ref, "A1:C10")

    def test_autofilter_with_a_filter_column_survives_an_elixcee_save(self):
        # Child <filterColumn>/<filters>/<filter> content must round-trip too, not
        # just the container's own `ref` attribute -- extract_raw_element captures
        # the whole element span verbatim, so this also guards against a future
        # regression that tried to reconstruct the container from parsed parts.
        with tempfile.TemporaryDirectory() as d:
            src = os.path.join(d, "af_cols_src.xlsx")
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.auto_filter.ref = "A1:C10"
            ws.auto_filter.add_filter_column(0, ["foo", "bar"], blank=False)
            wb.save(src)

            vm = elixcee.load_workbook(src)
            out = os.path.join(d, "af_cols_out.xlsx")
            vm.save_workbook(out)

            ws2 = openpyxl.load_workbook(out).active
            self.assertEqual(ws2.auto_filter.ref, "A1:C10")
            self.assertEqual(len(ws2.auto_filter.filterColumn), 1)
            col = ws2.auto_filter.filterColumn[0]
            self.assertEqual(col.colId, 0)
            self.assertEqual(sorted(col.filters.filter), ["bar", "foo"])

    def test_autofilter_survives_an_unrelated_cell_edit_and_a_real_merge(self):
        # Confirms schema position too: CT_Worksheet orders autoFilter BEFORE
        # mergeCells (verified against openpyxl's own writer, worksheet/_writer.py's
        # write_tail) -- an otherwise-correct-looking swap would produce a file
        # real Excel silently "repairs" by dropping the misordered element, even
        # though every byte is present.
        with tempfile.TemporaryDirectory() as d:
            src = os.path.join(d, "af_merge_src.xlsx")
            wb = openpyxl.Workbook()
            ws = wb.active
            ws["A1"] = "h"
            ws.merge_cells("D1:E1")
            ws.auto_filter.ref = "A1:C10"
            wb.save(src)

            vm = elixcee.load_workbook(src)
            vm.set_range("A2", [["unrelated edit"]])
            out = os.path.join(d, "af_merge_out.xlsx")
            vm.save_workbook(out)

            xml = zipfile.ZipFile(out).read("xl/worksheets/sheet1.xml").decode()
            af_pos = xml.find("<autoFilter")
            mc_pos = xml.find("<mergeCells")
            self.assertNotEqual(af_pos, -1)
            self.assertNotEqual(mc_pos, -1)
            self.assertLess(af_pos, mc_pos)

            ws2 = openpyxl.load_workbook(out).active
            self.assertEqual(ws2.auto_filter.ref, "A1:C10")
            self.assertEqual(ws2["A2"].value, "unrelated edit")
            self.assertIn("D1:E1", [str(r) for r in ws2.merged_cells.ranges])


class SortRangeAndMergeCellsRejectOversizedOrInvalidInput(unittest.TestCase):
    # Pins the PyO3-layer bound checks that have no Rust unit test of their
    # own (they live in #[cfg(feature = "python")] glue, not Vm-core logic) --
    # an oversized address here would otherwise write real geometry spanning
    # the whole sheet into the saved file, unlike get_range/iter_rows where
    # the only cost of an oversized address is a large allocation.
    def test_sort_range_rejects_an_address_beyond_the_sheet_bounds(self):
        vm = elixcee.Vm()
        with self.assertRaises(ValueError):
            vm.sort_range("A1:A1048577", key_col=1)

    def test_merge_cells_rejects_an_address_beyond_the_sheet_bounds(self):
        vm = elixcee.Vm()
        with self.assertRaises(ValueError):
            vm.merge_cells("A1:XFE1")

    def test_sort_range_rejects_a_key_col_outside_the_ranges_own_span(self):
        vm = elixcee.Vm()
        vm.set_range("A1:B2", [[1, 2], [3, 4]])
        with self.assertRaises(ValueError):
            vm.sort_range("A1:B2", key_col=3)


if __name__ == "__main__":
    unittest.main(verbosity=2)
