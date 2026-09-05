"""Regression checks for the comparator's durability and validation contracts."""
import tempfile
import types
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

import openpyxl
import equal_workbook_speed as bench


class BenchmarkContractTests(unittest.TestCase):
    def test_apple_uses_full_sync_not_fsync(self):
        file = Mock()
        file.fileno.return_value = 42
        fcntl = types.SimpleNamespace(F_FULLFSYNC=51, fcntl=Mock())
        with patch.object(bench.sys, 'platform', 'darwin'), \
             patch.dict('sys.modules', fcntl=fcntl), \
             patch.object(bench.os, 'fsync') as ordinary_sync:
            bench.durable_sync(file)
        fcntl.fcntl.assert_called_once_with(42, 51)
        ordinary_sync.assert_not_called()

    def test_non_apple_sync_failure_is_not_ignored(self):
        with patch.object(bench.sys, 'platform', 'linux'), \
             patch.object(bench.os, 'fsync', side_effect=OSError('sync failed')):
            with self.assertRaises(OSError):
                bench.durable_sync(Mock())

    def test_validation_rejects_untouched_cell_and_formula_loss(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'source.xlsx'
            workbook = openpyxl.Workbook()
            sheet = workbook.active
            sheet['A1'] = 123
            sheet['B1'] = '=1+2'
            sheet['C2'] = 'untouched 日本語 &<>'
            workbook.save(path)
            workbook.close()
            expected = bench.logical(path)
            bench.check(path, expected)
            expected[sheet.title]['C2'] = 'lost'
            with self.assertRaises(AssertionError):
                bench.check(path, expected)
            expected = bench.logical(path)
            expected[sheet.title]['B1'] = 3
            with self.assertRaises(AssertionError):
                bench.check(path, expected)


if __name__ == '__main__':
    unittest.main()
