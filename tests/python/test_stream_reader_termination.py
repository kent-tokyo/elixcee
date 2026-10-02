"""Run against an installed wheel: python -I tests/python/test_stream_reader_termination.py."""

import pathlib
import tempfile
import unittest

import elixcee


class StreamReaderTermination(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory(prefix="elixcee-reader-termination-")
        self.addCleanup(self.directory.cleanup)

    def workbook(self, name, row_count):
        path = pathlib.Path(self.directory.name) / name
        with elixcee.create_stream(str(path)) as writer:
            for row in range(1, row_count + 1):
                writer.append([row])
        return path

    def read(self, path, max_rows):
        reader = elixcee.open_stream(str(path), max_rows=max_rows)
        rows = list(reader)
        return reader, rows

    def test_below_limit_reports_clean_eof(self):
        reader, rows = self.read(self.workbook("below.xlsx", 1), 2)
        self.assertEqual(rows, [[1]])
        self.assertEqual(reader.termination_reason, "eof")
        self.assertFalse(reader.limit_reached)
        self.assertEqual(reader.rows_read, 1)
        self.assertTrue(reader.closed)

    def test_exact_limit_reports_clean_eof(self):
        reader, rows = self.read(self.workbook("exact.xlsx", 2), 2)
        self.assertEqual(rows, [[1], [2]])
        self.assertEqual(reader.termination_reason, "eof")
        self.assertFalse(reader.limit_reached)
        self.assertEqual(reader.rows_read, 2)

    def test_over_limit_reports_confirmed_truncation_without_returning_extra_row(self):
        reader, rows = self.read(self.workbook("over.xlsx", 3), 2)
        self.assertEqual(rows, [[1], [2]])
        self.assertEqual(reader.termination_reason, "max_rows")
        self.assertTrue(reader.limit_reached)
        self.assertEqual(reader.rows_read, 2)

    def test_explicit_close_has_its_own_reason(self):
        reader = elixcee.open_stream(str(self.workbook("close.xlsx", 1)))
        self.assertIsNone(reader.termination_reason)
        reader.close()
        self.assertEqual(reader.termination_reason, "closed")
        self.assertFalse(reader.limit_reached)

    def test_cancellation_is_typed_and_recorded(self):
        cancellation = elixcee.ReadCancellation()
        reader = elixcee.open_stream(
            str(self.workbook("cancel.xlsx", 3)), cancellation=cancellation
        )
        cancellation.cancel()
        with self.assertRaises(InterruptedError):
            next(reader)
        self.assertEqual(reader.termination_reason, "canceled")
        self.assertFalse(reader.limit_reached)
        self.assertTrue(reader.closed)


if __name__ == "__main__":
    unittest.main()
