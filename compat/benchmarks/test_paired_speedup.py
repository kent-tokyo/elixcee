import tempfile
import unittest
import warnings
import zipfile
from pathlib import Path

from paired_speedup import compare_parts, zip_parts


class ZipEqualityTests(unittest.TestCase):
    def test_compression_can_differ_but_not_cell_formula_or_member_list(self):
        with tempfile.TemporaryDirectory() as directory:
            before, after = (Path(directory)/name for name in ('before.zip', 'after.zip'))
            source = '<worksheet><c><f>1+2</f><v>3</v></c></worksheet>'
            with zipfile.ZipFile(before, 'w', compression=zipfile.ZIP_STORED) as archive:
                archive.writestr('sheet.xml', source)
            with zipfile.ZipFile(after, 'w', compression=zipfile.ZIP_DEFLATED) as archive:
                archive.writestr('sheet.xml', source)
            compare_parts(before, after)
            with zipfile.ZipFile(after, 'w') as archive:
                archive.writestr('sheet.xml', source.replace('<f>1+2</f>', ''))
            with self.assertRaises(AssertionError):
                compare_parts(before, after)
            with zipfile.ZipFile(after, 'w') as archive:
                archive.writestr('other.xml', source)
            with self.assertRaises(AssertionError):
                compare_parts(before, after)

    def test_duplicate_zip_members_are_rejected(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory)/'duplicate.zip'
            with zipfile.ZipFile(path, 'w') as archive, warnings.catch_warnings():
                warnings.simplefilter('ignore', UserWarning)
                archive.writestr('sheet.xml', 'one')
                archive.writestr('sheet.xml', 'two')
            with self.assertRaises(AssertionError):
                zip_parts(path)


if __name__ == '__main__':
    unittest.main()
