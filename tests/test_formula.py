"""Formula engine: recalculates a cell the way Excel would when it opens the file."""
import unittest

from parity_check.formula import Engine, FormulaError
from parity_check.xlsx import Cell
from tests.helpers import workbook


class CellValues(unittest.TestCase):
    def test_number_is_read_as_is(self):
        self.assertEqual(Engine(workbook({"S!A1": 42.5})).value("S", "A1"), 42.5)

    def test_missing_empty_or_text_cell_counts_as_zero_like_excel_sum(self):
        wb = workbook({"S!A2": None, "S!A3": "hello"}, sheets=["S"])
        engine = Engine(wb)
        self.assertEqual(engine.value("S", "A1"), 0)
        self.assertEqual(engine.value("S", "A2"), 0)
        self.assertEqual(engine.value("S", "A3"), 0)

    def test_number_stored_as_text_is_converted(self):
        self.assertEqual(Engine(workbook({"S!A1": "17"})).value("S", "A1"), 17)

    def test_saved_value_of_a_formula_is_ignored(self):
        wb = workbook({"S!A1": 10})
        wb["S"]["B1"] = Cell(formula="A1*2", value=999)  # stale saved value
        self.assertEqual(Engine(wb).value("S", "B1"), 20)


class Arithmetic(unittest.TestCase):
    base = {"S!A1": 10, "S!A2": 3, "S!A3": 2}

    def calc(self, formula):
        return Engine(workbook({**self.base, "S!B1": formula})).value("S", "B1")

    def test_plus_and_minus(self):
        self.assertEqual(self.calc("=A1+A2-A3"), 11)

    def test_multiplication_before_addition(self):
        self.assertEqual(self.calc("=A1+A2*A3"), 16)

    def test_parentheses_and_unary_minus(self):
        self.assertEqual(self.calc("=-(A1+A2)*A3"), -26)

    def test_division(self):
        self.assertEqual(self.calc("=A1/A3"), 5)

    def test_absolute_references(self):
        self.assertEqual(self.calc("=$A$1+A$2"), 13)

    def test_lowercase_reference(self):
        self.assertEqual(self.calc("=a1+a2"), 13)

    def test_division_by_zero_is_an_error(self):
        with self.assertRaises(FormulaError):
            self.calc("=A1/0")


class Functions(unittest.TestCase):
    def test_sum_of_a_column_range(self):
        wb = workbook({"S!A1": 1, "S!A2": 2, "S!A3": 4, "S!B1": "=SUM(A1:A3)"})
        self.assertEqual(Engine(wb).value("S", "B1"), 7)

    def test_sum_of_a_rectangle_with_empty_cells(self):
        wb = workbook({"S!A1": 1, "S!B1": 2, "S!A2": 4, "S!C1": "=SUM(A1:B2)"})
        self.assertEqual(Engine(wb).value("S", "C1"), 7)

    def test_sum_mixes_ranges_cells_and_numbers(self):
        wb = workbook({"S!A1": 1, "S!A2": 2, "S!B1": 10, "S!C1": "=SUM(A1:A2,B1,5)"})
        self.assertEqual(Engine(wb).value("S", "C1"), 18)

    def test_unsupported_function_is_an_error_never_a_silent_zero(self):
        wb = workbook({"S!A1": 1, "S!B1": "=ROUND(A1,0)"})
        with self.assertRaisesRegex(FormulaError, "ROUND"):
            Engine(wb).value("S", "B1")

    def test_garbage_after_the_formula_is_an_error(self):
        wb = workbook({"S!A1": 1, "S!B1": "=A1 A1"})
        with self.assertRaises(FormulaError):
            Engine(wb).value("S", "B1")


class SheetReferences(unittest.TestCase):
    def test_reference_to_another_sheet(self):
        wb = workbook({"Other!D5": 9, "S!B1": "=Other!D5*2"})
        self.assertEqual(Engine(wb).value("S", "B1"), 18)

    def test_quoted_sheet_name_with_spaces(self):
        wb = workbook({"Club total!B2": 5, "S!B1": "='Club total'!B2+1"})
        self.assertEqual(Engine(wb).value("S", "B1"), 6)

    def test_quoted_sheet_name_with_an_escaped_quote(self):
        wb = workbook({"Coach's fund!A1": 4, "S!B1": "='Coach''s fund'!A1"})
        self.assertEqual(Engine(wb).value("S", "B1"), 4)

    def test_sum_over_a_range_on_another_sheet(self):
        wb = workbook({"Other!A1": 1, "Other!A2": 2, "S!B1": "=SUM(Other!A1:A2)"})
        self.assertEqual(Engine(wb).value("S", "B1"), 3)

    def test_reference_to_an_external_file_is_an_error(self):
        wb = workbook({"S!B1": "=[1]Other!A1"})
        with self.assertRaisesRegex(FormulaError, "external"):
            Engine(wb).value("S", "B1")

    def test_missing_sheet_is_an_error(self):
        wb = workbook({"S!B1": "=Missing!A1"})
        with self.assertRaisesRegex(FormulaError, "Missing"):
            Engine(wb).value("S", "B1")


class CellsTheReaderCouldNotExpand(unittest.TestCase):
    def test_cell_flagged_by_the_reader_is_an_error(self):
        wb = {"S": {"A1": Cell(value=3, problem="shared formula not expanded")}}
        with self.assertRaisesRegex(FormulaError, "shared formula"):
            Engine(wb).value("S", "A1")


class CachingAndCycles(unittest.TestCase):
    def test_cycle_is_an_error(self):
        wb = workbook({"S!A1": "=B1", "S!B1": "=A1"})
        with self.assertRaisesRegex(FormulaError, "cycle"):
            Engine(wb).value("S", "A1")

    def test_shared_cell_is_computed_once(self):
        reads = []

        class CountingSheet(dict):
            def get(self, key, default=None):
                reads.append(key)
                return super().get(key, default)

        wb = workbook({"S!A1": 2, "S!B1": "=A1", "S!C1": "=A1", "S!D1": "=B1+C1+A1"})
        wb["S"] = CountingSheet(wb["S"])
        self.assertEqual(Engine(wb).value("S", "D1"), 6)
        self.assertEqual(reads.count("A1"), 1)


if __name__ == "__main__":
    unittest.main()
