"""Run against an installed wheel: python -I tests/python/test_stream_writer_limits.py."""

import pathlib
import tempfile
import unittest
import zipfile

import elixcee


class StreamWriterLimits(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory(prefix="elixcee-writer-limits-")
        self.addCleanup(self.directory.cleanup)
        self.path = pathlib.Path(self.directory.name) / "output.xlsx"

    def writer(self, **options):
        writer = elixcee.create_stream(str(self.path), **options)
        self.addCleanup(writer.close)
        return writer

    def test_column_limit_stops_before_converting_excess_item(self):
        seen = []

        def row():
            for item in (1, 2, object()):
                seen.append(item)
                yield item
            self.fail("writer consumed beyond the first excess column")

        writer = self.writer(max_columns=2)
        with self.assertRaisesRegex(MemoryError, "2 columns"):
            writer.append(row())
        self.assertEqual(len(seen), 3)
        self.assertEqual((writer.row_count, writer.pending_bytes), (0, 0))
        writer.append([42])
        writer.close()
        self.assertEqual(elixcee.load_workbook(str(self.path)).get_cell(1, 1), 42)

    def test_row_limit_does_not_open_or_consume_iterator(self):
        class NeverIterate:
            def __iter__(self):
                raise AssertionError("iterator must not be opened")

        writer = self.writer(max_rows=1)
        writer.append([42])
        before = writer.pending_bytes
        with self.assertRaisesRegex(MemoryError, "1 rows"):
            writer.append(NeverIterate())
        self.assertEqual((writer.row_count, writer.pending_bytes), (1, before))

    def test_empty_strings_cannot_evade_byte_budget(self):
        seen = []

        def row():
            for i in range(65):
                seen.append(i)
                yield ""
            self.fail("zero-length payloads bypassed per-cell accounting")

        writer = self.writer(max_pending_bytes=64)
        with self.assertRaisesRegex(MemoryError, "64 bytes"):
            writer.append(row())
        self.assertLessEqual(len(seen), 65)
        self.assertEqual((writer.row_count, writer.pending_bytes), (0, 0))

    def test_byte_limit_rejects_row_without_consuming_tail(self):
        def row():
            yield "small"
            yield "x" * 2048
            self.fail("writer consumed past byte-budget rejection")

        writer = self.writer(max_pending_bytes=1024)
        writer.append([1])
        before = writer.pending_bytes
        with self.assertRaises(MemoryError):
            writer.append(row())
        self.assertEqual((writer.row_count, writer.pending_bytes), (1, before))
        writer.append([2])
        writer.close()
        workbook = elixcee.load_workbook(str(self.path))
        self.assertEqual(workbook.get_cell(1, 1), 1)
        self.assertEqual(workbook.get_cell(2, 1), 2)
        with zipfile.ZipFile(self.path) as archive:
            xml = archive.read("xl/worksheets/sheet1.xml")
        self.assertNotIn(b"small", xml)
        self.assertNotIn(b'<row r="3"', xml)

    def test_conversion_and_iterator_errors_do_not_commit_partial_rows(self):
        def broken():
            yield 1
            raise RuntimeError("iterator failure")

        writer = self.writer()
        for row, error in (([1, object()], TypeError), (broken(), RuntimeError)):
            with self.assertRaises(error):
                writer.append(row)
            self.assertEqual((writer.row_count, writer.pending_bytes), (0, 0))
        writer.append([7])
        writer.close()
        self.assertEqual(elixcee.load_workbook(str(self.path)).get_cell(1, 1), 7)

    def test_empty_row_leaves_counters_unchanged(self):
        writer = self.writer()
        with self.assertRaisesRegex(ValueError, "row must not be empty"):
            writer.append(iter(()))
        self.assertEqual((writer.row_count, writer.pending_bytes), (0, 0))

    def test_budget_includes_utf8_and_cell_overhead(self):
        writer = self.writer()
        writer.append([""])
        cell_overhead = writer.pending_bytes
        self.assertGreater(cell_overhead, 0)
        value = '日本<&"'
        writer.append([value])
        self.assertEqual(writer.pending_bytes, cell_overhead * 2 + len(value.encode()))
        writer.close()
        self.assertEqual(writer.pending_bytes, 0)
        self.assertEqual(elixcee.load_workbook(str(self.path)).get_cell(2, 1), value)

    def test_exact_cumulative_budget_is_accepted_then_rejected(self):
        probe = self.writer()
        probe.append([1])
        cost = probe.pending_bytes
        probe.close()
        writer = self.writer(max_pending_bytes=cost * 2)
        writer.append([1])
        writer.append([2])
        self.assertEqual(writer.pending_bytes, cost * 2)
        with self.assertRaises(MemoryError):
            writer.append([3])
        self.assertEqual((writer.row_count, writer.pending_bytes), (2, cost * 2))

    def test_rejection_does_not_publish_over_existing_destination(self):
        self.path.write_bytes(b"original destination")
        writer = self.writer(max_columns=1)
        with self.assertRaises(MemoryError):
            writer.append([1, 2])
        self.assertEqual(self.path.read_bytes(), b"original destination")
        writer.append([9])
        self.assertEqual(self.path.read_bytes(), b"original destination")
        writer.close()
        self.assertEqual(elixcee.load_workbook(str(self.path)).get_cell(1, 1), 9)

    def test_bounded_api_exposes_separate_row_and_work_limits(self):
        writer = elixcee.create_stream_bounded(
            str(self.path), max_row_bytes=128, max_work_bytes=512, max_rows=4, max_columns=2
        )
        self.addCleanup(writer.close)
        self.assertEqual(writer.max_row_bytes, 128)
        self.assertEqual(writer.max_work_bytes, 512)
        writer.append([1, "ok"])
        with self.assertRaisesRegex(MemoryError, "max_row_bytes"):
            writer.append(["x" * 200])
        self.assertEqual(writer.row_count, 1)


if __name__ == "__main__":
    unittest.main()
