#!/usr/bin/env python3
"""Differential check: elixcee's sheet management API (P1 core 3 + remainder +
P2 hidden row/col) vs. openpyxl.

openpyxl is used PURELY as a test-only oracle here -- never a runtime dependency
of the shipped `elixcee` package (see pyproject.toml, which declares none).
Requires a `maturin develop --features python` build of elixcee and
`pip install openpyxl` in the active environment; see this directory's
README.md for the one-time setup.

Compares rename_sheet()/move_sheet()/merged_cells()/merge_cells()/
unmerge_cells()/hidden_rows()/hidden_columns()/set_row_hidden()/
set_column_hidden() against openpyxl's own read of the same real fixture
after a save/reload round trip. Must build from FIXTURE via
elixcee.load_workbook, not a bare elixcee.Vm() -- a from-scratch VM's
minimal styles.xml emits a bare <fill/> that openpyxl's own reader rejects
on reopen (a real, pre-existing, unrelated bug -- see ROADMAP.md's known
gaps).

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
