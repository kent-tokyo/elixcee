#!/usr/bin/env python3
"""Differential check: elixcee's sheet management API (P1 core 3) vs. openpyxl.

openpyxl is used PURELY as a test-only oracle here -- never a runtime dependency
of the shipped `elixcee` package (see pyproject.toml, which declares none).
Requires a `maturin develop --features python` build of elixcee and
`pip install openpyxl` in the active environment; see this directory's
README.md for the one-time setup.

Compares rename_sheet()/move_sheet()/merged_cells() against openpyxl's own
read of the same real fixture after a save/reload round trip. Must build from
FIXTURE via elixcee.load_workbook, not a bare elixcee.Vm() -- a from-scratch
VM's minimal styles.xml emits a bare <fill/> that openpyxl's own reader
rejects on reopen (a real, pre-existing, unrelated bug -- see
ROADMAP.md's known gaps).

Row/col insert-delete is deliberately given NO differential coverage here:
the disclosed fidelity gap (merges/hidden markers/styles/formats not shifted)
means an openpyxl comparison would correctly fail on exactly the cases worth
testing -- Rust unit + integration tests already cover the values-only
behavior that IS correct.

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


if __name__ == "__main__":
    unittest.main(verbosity=2)
