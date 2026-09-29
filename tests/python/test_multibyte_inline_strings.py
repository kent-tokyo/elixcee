"""Regression tests for inline-string UTF-8 boundary handling.

Run against an installed wheel with:
    python -I tests/python/test_multibyte_inline_strings.py
"""

import pathlib
import tempfile
import unittest
import zipfile

import elixcee

DATA_ROWS = 146
BAD_ROW = 12


def cell_text(row):
    if row == 1:
        return "取引先"
    if row == BAD_ROW:
        return "A&フフフ 商事"
    return f"株式会社テスト{row}"


def escape(text):
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def write_erp_style_xlsx(path):
    rows = "".join(
        f'<row r="{row}"><c r="A{row}" t="inlineStr"><is><t>{escape(cell_text(row))}</t></is></c>'
        f'<c r="B{row}"><v>{row}</v></c></row>'
        for row in range(1, DATA_ROWS + 2)
    )
    parts = {
        "[Content_Types].xml": (
            '<?xml version="1.0" encoding="UTF-8"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
            '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            "</Types>"
        ),
        "_rels/.rels": (
            '<?xml version="1.0" encoding="UTF-8"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
            "</Relationships>"
        ),
        "xl/workbook.xml": (
            '<?xml version="1.0" encoding="UTF-8"?>'
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>'
        ),
        "xl/_rels/workbook.xml.rels": (
            '<?xml version="1.0" encoding="UTF-8"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
            "</Relationships>"
        ),
        "xl/worksheets/sheet1.xml": (
            '<?xml version="1.0" encoding="UTF-8"?>'
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            f"<sheetData>{rows}</sheetData></worksheet>"
        ),
    }
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for name, payload in parts.items():
            archive.writestr(name, payload.encode("utf-8"))


class MultibyteInlineStrings(unittest.TestCase):
    def setUp(self):
        directory = tempfile.TemporaryDirectory(prefix="elixcee-multibyte-")
        self.addCleanup(directory.cleanup)
        self.path = str(pathlib.Path(directory.name) / "erp_export.xlsx")
        write_erp_style_xlsx(self.path)

    def test_load_workbook_reads_ampersand_followed_by_japanese(self):
        vm = elixcee.load_workbook(self.path)
        self.assertEqual(vm.get_cell(BAD_ROW, 1), "A&フフフ 商事")
        self.assertEqual(vm.get_cell(DATA_ROWS + 1, 1), cell_text(DATA_ROWS + 1))
        self.assertEqual(vm.get_cell(DATA_ROWS + 1, 2), DATA_ROWS + 1)

    def test_open_stream_returns_every_row(self):
        with elixcee.open_stream(self.path, include_row_numbers=True) as reader:
            rows = list(reader)
        self.assertEqual(len(rows), DATA_ROWS + 1)
        self.assertEqual(rows[BAD_ROW - 1], (BAD_ROW, ["A&フフフ 商事", BAD_ROW]))
        self.assertEqual(rows[-1][0], DATA_ROWS + 1)

    def test_internal_error_is_catchable_as_exception(self):
        self.assertTrue(issubclass(elixcee.InternalError, RuntimeError))
        self.assertTrue(issubclass(elixcee.InternalError, Exception))


if __name__ == "__main__":
    unittest.main()
