"""Reading and writing .xlsx files with the standard library."""
import tempfile
import unittest
import zipfile
from pathlib import Path

from parity_check.xlsx import Cell, read_workbook, write_workbook


class RoundTrip(unittest.TestCase):
    def setUp(self):
        self.dir = tempfile.TemporaryDirectory()
        self.path = Path(self.dir.name) / "book.xlsx"

    def tearDown(self):
        self.dir.cleanup()

    def test_numbers_text_and_formulas_survive_a_round_trip(self):
        write_workbook(self.path, {
            "Club total": {"A1": Cell(value="Week 1"), "B1": Cell(value=1000), "C1": Cell(formula="B1*2", value=2000)},
            "Youth": {"A1": Cell(value=2.5)},
        })
        wb = read_workbook(self.path)
        self.assertEqual(list(wb), ["Club total", "Youth"])
        self.assertEqual(wb["Club total"]["A1"].value, "Week 1")
        self.assertEqual(wb["Club total"]["B1"].value, 1000)
        self.assertEqual(wb["Club total"]["C1"].formula, "B1*2")
        self.assertEqual(wb["Youth"]["A1"].value, 2.5)

    def test_saved_value_of_a_formula_is_kept_even_when_stale(self):
        write_workbook(self.path, {"S": {"A1": Cell(value=1), "B1": Cell(formula="A1", value=777)}})
        self.assertEqual(read_workbook(self.path)["S"]["B1"].value, 777)

    def test_written_file_asks_excel_to_recalculate_on_open(self):
        write_workbook(self.path, {"S": {"A1": Cell(value=1)}})
        with zipfile.ZipFile(self.path) as z:
            self.assertIn('fullCalcOnLoad="1"', z.read("xl/workbook.xml").decode())

    def test_text_with_xml_characters_is_escaped(self):
        write_workbook(self.path, {"S": {"A1": Cell(value="Kit & <balls>")}})
        self.assertEqual(read_workbook(self.path)["S"]["A1"].value, "Kit & <balls>")


class FilesSavedByExcel(unittest.TestCase):
    """Excel stores text in a shared table and copies formulas down as 'shared formulas'."""

    def make(self, sheet_xml: str, shared_strings: str = "") -> Path:
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        path = Path(tmp.name) / "excel.xlsx"
        ns = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
        rel_ns = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
        with zipfile.ZipFile(path, "w") as z:
            z.writestr("xl/workbook.xml", f'<workbook {ns} {rel_ns}><sheets>'
                                          '<sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>')
            z.writestr("xl/_rels/workbook.xml.rels",
                       '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                       '<Relationship Id="rId1" Target="worksheets/sheet1.xml" '
                       'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"/>'
                       '</Relationships>')
            z.writestr("xl/worksheets/sheet1.xml", f"<worksheet {ns}><sheetData>{sheet_xml}</sheetData></worksheet>")
            if shared_strings:
                z.writestr("xl/sharedStrings.xml", f"<sst {ns}>{shared_strings}</sst>")
        return path

    def test_shared_strings_are_resolved(self):
        path = self.make('<row r="1"><c r="A1" t="s"><v>1</v></c></row>',
                         "<si><t>zero</t></si><si><r><t>Month </t></r><r><t>total</t></r></si>")
        self.assertEqual(read_workbook(path)["S"]["A1"].value, "Month total")

    def test_copied_down_shared_formula_is_flagged_not_guessed(self):
        path = self.make('<row r="1"><c r="A1"><f t="shared" ref="A1:A2" si="0">B1+1</f><v>2</v></c></row>'
                         '<row r="2"><c r="A2"><f t="shared" si="0"/><v>3</v></c></row>')
        wb = read_workbook(path)
        self.assertEqual(wb["S"]["A1"].formula, "B1+1")
        self.assertIsNone(wb["S"]["A2"].formula)
        self.assertIn("shared formula", wb["S"]["A2"].problem)


if __name__ == "__main__":
    unittest.main()
