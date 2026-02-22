import tempfile
import unittest
from pathlib import Path
from zipfile import ZipFile

from src.excel.xlsx_reader import read_xlsx_rows


class TestXlsxReader(unittest.TestCase):
    def test_read_rows(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            xlsx_path = Path(tmpdir) / "sample.xlsx"
            with ZipFile(xlsx_path, "w") as zf:
                zf.writestr("[Content_Types].xml", """<?xml version=\"1.0\" encoding=\"UTF-8\"?>
<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">
  <Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>
  <Default Extension=\"xml\" ContentType=\"application/xml\"/>
  <Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>
  <Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>
  <Override PartName=\"/xl/sharedStrings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/>
</Types>""")
                zf.writestr("xl/workbook.xml", """<?xml version=\"1.0\" encoding=\"UTF-8\"?>
<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">
  <sheets><sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/></sheets>
</workbook>""")
                zf.writestr("xl/_rels/workbook.xml.rels", """<?xml version=\"1.0\" encoding=\"UTF-8\"?>
<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">
  <Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>
</Relationships>""")
                zf.writestr("xl/sharedStrings.xml", """<?xml version=\"1.0\" encoding=\"UTF-8\"?>
<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"2\" uniqueCount=\"2\">
  <si><t>ID</t></si>
  <si><t>Item</t></si>
</sst>""")
                zf.writestr("xl/worksheets/sheet1.xml", """<?xml version=\"1.0\" encoding=\"UTF-8\"?>
<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">
  <sheetData>
    <row r=\"1\"><c r=\"A1\" t=\"s\"><v>0</v></c><c r=\"B1\" t=\"s\"><v>1</v></c></row>
    <row r=\"2\"><c r=\"A2\"><v>1</v></c><c r=\"B2\" t=\"inlineStr\"><is><t>Alpha</t></is></c></row>
  </sheetData>
</worksheet>""")

            rows = read_xlsx_rows(str(xlsx_path))
            self.assertEqual(rows[0], ["ID", "Item"])
            self.assertEqual(rows[1], ["1", "Alpha"])


if __name__ == "__main__":
    unittest.main()
