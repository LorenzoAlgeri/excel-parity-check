"""The gate: every checked cell is MATCH, KNOWN or UNEXPLAINED; one UNEXPLAINED fails it."""
import unittest

from parity_check.gate import (
    KnownDifference, classify, known_from_list, run_gate, spec_from_dict,
)
from parity_check.layout import build_workbook
from parity_check.xlsx import Cell

SPEC = spec_from_dict({
    "sheets": {
        "Youth": {
            "Fees": {"weeks": [1000, 2000], "budget": 2000},
            "Events": {"weeks": [0, 1000], "budget": 2000},
        },
        "Seniors": {
            "Fees": {"weeks": [3000, 0], "budget": 4000},
            "Events": {"weeks": [1000, 1000], "budget": 1000},
        },
    },
    "totals": {"Club total": ["Youth", "Seniors"]},
})

# Layout built by parity_check.layout for this spec:
#   row 1 headers, rows 2-3 weeks, 4 month total, 5 budget, 6 variance
#   column A labels, B Fees, C Events, D Total


def find(report, sheet, column, row):
    return next(c for c in report.cells if (c.sheet, c.column, c.row) == (sheet, column, row))


class Classify(unittest.TestCase):
    def test_equal_within_one_cent_is_match(self):
        self.assertEqual(classify(100.0, 100.004, None, None).status, "MATCH")

    def test_difference_without_a_declared_cause_is_unexplained(self):
        result = classify(100.0, 90.0, None, None)
        self.assertEqual(result.status, "UNEXPLAINED")
        self.assertIn("expected 100", result.reason)

    def test_declared_difference_is_known(self):
        kd = KnownDifference("KD-9", ["S"], ["Fees"], ["Week 1"], -10, "booked late")
        result = classify(100.0, 90.0, None, kd)
        self.assertEqual(result.status, "KNOWN")
        self.assertIn("KD-9", result.reason)

    def test_recalculation_error_is_unexplained_even_if_a_known_difference_exists(self):
        kd = KnownDifference("KD-9", ["S"], ["Fees"], ["Week 1"], -10, "booked late")
        result = classify(100.0, None, "function not supported: ROUND()", kd)
        self.assertEqual(result.status, "UNEXPLAINED")
        self.assertIn("cannot recalculate", result.reason)


class CorrectFile(unittest.TestCase):
    def test_correct_file_passes_with_every_cell_matching(self):
        report = run_gate(build_workbook(SPEC), SPEC)
        self.assertEqual(report.verdict, "PASS")
        self.assertEqual(report.counts, {"MATCH": 45, "KNOWN": 0, "UNEXPLAINED": 0})

    def test_totals_sheet_expects_the_sum_of_the_areas(self):
        report = run_gate(build_workbook(SPEC), SPEC)
        cell = find(report, "Club total", "Total", "Month total")
        self.assertEqual(cell.expected, 9000)
        self.assertEqual(cell.ref, "D4")


class UnexplainedDifference(unittest.TestCase):
    def test_one_wrong_input_fails_the_gate(self):
        wb = build_workbook(SPEC)
        wb["Youth"]["B2"] = Cell(value=1500)  # Fees, Week 1: expected 1000
        report = run_gate(wb, SPEC)
        self.assertEqual(report.verdict, "FAIL")
        self.assertEqual(find(report, "Youth", "Fees", "Week 1").status, "UNEXPLAINED")
        # the error travels up: month total, variance, total column, club total
        self.assertEqual(find(report, "Club total", "Total", "Month total").status, "UNEXPLAINED")


class KnownDifferences(unittest.TestCase):
    def wb_with_late_booking(self):
        wb = build_workbook(SPEC)
        wb["Seniors"]["C3"] = Cell(value=0)  # Events, Week 2: expected 1000
        return wb

    def known(self, delta):
        return known_from_list([{
            "id": "KD-1", "sheets": ["Seniors", "Club total"], "columns": ["Events", "Total"],
            "rows": ["Week 2", "Month total", "Variance vs budget"], "delta": delta,
            "reason": "payment booked after the file was generated",
        }])

    def test_declared_difference_passes_and_is_reported_as_known(self):
        report = run_gate(self.wb_with_late_booking(), SPEC, self.known(-1000))
        self.assertEqual(report.verdict, "PASS")
        self.assertEqual(report.counts["KNOWN"], 12)
        self.assertIn("KD-1", find(report, "Seniors", "Events", "Week 2").reason)

    def test_declared_difference_does_not_cover_a_different_amount(self):
        report = run_gate(self.wb_with_late_booking(), SPEC, self.known(-500))
        self.assertEqual(report.verdict, "FAIL")
        self.assertEqual(report.counts["KNOWN"], 0)

    def test_declared_difference_needs_a_reason(self):
        with self.assertRaises(ValueError):
            known_from_list([{"id": "KD-2", "sheets": ["Youth"], "columns": ["Fees"],
                              "rows": ["Week 1"], "delta": 5, "reason": ""}])


class RecalculatedNotSaved(unittest.TestCase):
    def test_wrong_formula_with_a_correct_saved_value_fails(self):
        wb = build_workbook(SPEC)
        # Month total should sum B2:B3 (3000); the formula forgets week 2,
        # but the saved value still says 3000.
        wb["Youth"]["B4"] = Cell(formula="SUM(B2:B2)", value=3000)
        report = run_gate(wb, SPEC)
        cell = find(report, "Youth", "Fees", "Month total")
        self.assertEqual(cell.actual, 1000)
        self.assertEqual(cell.status, "UNEXPLAINED")

    def test_right_formula_with_a_wrong_saved_value_passes(self):
        wb = build_workbook(SPEC)
        wb["Youth"]["B4"] = Cell(formula="SUM(B2:B3)", value=0)
        self.assertEqual(run_gate(wb, SPEC).verdict, "PASS")

    def test_formula_that_cannot_be_recalculated_is_unexplained(self):
        wb = build_workbook(SPEC)
        wb["Youth"]["B4"] = Cell(formula="ROUND(SUM(B2:B3),0)", value=3000)
        report = run_gate(wb, SPEC)
        cell = find(report, "Youth", "Fees", "Month total")
        self.assertEqual(cell.status, "UNEXPLAINED")
        self.assertIn("ROUND", cell.reason)


class InputsOnlyIsNotEnough(unittest.TestCase):
    """The real story: right inputs, empty subtotal, input check green."""

    def setUp(self):
        self.wb = build_workbook(SPEC)
        self.wb["Seniors"]["B4"] = Cell()  # Fees month total: empty, no formula

    def test_input_check_alone_passes(self):
        report = run_gate(self.wb, SPEC, inputs_only=True)
        self.assertEqual(report.verdict, "PASS")
        self.assertEqual(report.counts, {"MATCH": 12, "KNOWN": 0, "UNEXPLAINED": 0})

    def test_full_recalculation_fails_on_the_same_file(self):
        report = run_gate(self.wb, SPEC)
        self.assertEqual(report.verdict, "FAIL")
        self.assertEqual(find(report, "Seniors", "Fees", "Month total").actual, 0)
        self.assertEqual(find(report, "Club total", "Total", "Month total").status, "UNEXPLAINED")


class BrokenStructure(unittest.TestCase):
    def test_missing_row_label_makes_its_cells_unexplained(self):
        wb = build_workbook(SPEC)
        wb["Youth"]["A4"] = Cell(value="Monthly sum")  # label changed
        report = run_gate(wb, SPEC)
        cell = find(report, "Youth", "Fees", "Month total")
        self.assertEqual(cell.status, "UNEXPLAINED")
        self.assertIn("row 'Month total' not found", cell.reason)

    def test_missing_sheet_makes_its_cells_unexplained(self):
        wb = build_workbook(SPEC)
        del wb["Seniors"]
        report = run_gate(wb, SPEC)
        self.assertEqual(report.verdict, "FAIL")
        self.assertIn("sheet 'Seniors' not found", find(report, "Seniors", "Fees", "Week 1").reason)


class Summary(unittest.TestCase):
    def test_summary_line(self):
        report = run_gate(build_workbook(SPEC), SPEC)
        self.assertEqual(report.summary(), "PASS: 45 cells checked, 45 match, 0 known, 0 unexplained")


class MissingColumn(unittest.TestCase):
    def test_missing_column_header_makes_its_cells_unexplained(self):
        wb = build_workbook(SPEC)
        wb["Youth"]["C1"] = Cell(value="Trips")  # header "Events" renamed
        report = run_gate(wb, SPEC)
        self.assertIn("column 'Events' not found", find(report, "Youth", "Events", "Week 1").reason)


class SpecValidation(unittest.TestCase):
    def test_sheets_must_list_the_same_lines(self):
        with self.assertRaisesRegex(ValueError, "same lines"):
            spec_from_dict({"sheets": {"A": {"X": {"weeks": [1], "budget": 1}},
                                       "B": {"Y": {"weeks": [1], "budget": 1}}}})

    def test_every_line_needs_the_same_number_of_weeks(self):
        with self.assertRaisesRegex(ValueError, "weeks"):
            spec_from_dict({"sheets": {"A": {"X": {"weeks": [1], "budget": 1}},
                                       "B": {"X": {"weeks": [1, 2], "budget": 1}}}})

    def test_totals_sheet_can_only_add_known_sheets(self):
        with self.assertRaisesRegex(ValueError, "unknown sheets"):
            spec_from_dict({"sheets": {"A": {"X": {"weeks": [1], "budget": 1}}},
                            "totals": {"T": ["A", "Nope"]}})

if __name__ == "__main__":
    unittest.main()
