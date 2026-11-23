"""End to end: the command line on the four example files in examples/."""
import subprocess
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
EXAMPLES = ROOT / "examples"


def gate(*args):
    result = subprocess.run([sys.executable, "-B", "-m", "parity_check", *args],
                            cwd=ROOT, capture_output=True, text=True)
    return result.returncode, result.stdout + result.stderr


class ExampleFiles(unittest.TestCase):
    def test_good_file_passes_with_the_declared_difference_shown_as_known(self):
        code, out = gate("examples/good.xlsx")
        self.assertEqual(code, 0, out)
        self.assertIn("PASS: 84 cells checked, 72 match, 12 known, 0 unexplained", out)
        self.assertIn("[KD-1]", out)

    def test_empty_subtotal_fails(self):
        code, out = gate("examples/empty_subtotal.xlsx")
        self.assertEqual(code, 1, out)
        self.assertIn("FAIL: 84 cells checked, 68 match, 8 known, 8 unexplained", out)
        self.assertIn("Seniors!C6", out)

    def test_empty_subtotal_passes_when_only_inputs_are_checked(self):
        # The real incident: the input check was green while the file was wrong.
        code, out = gate("examples/empty_subtotal.xlsx", "--inputs-only")
        self.assertEqual(code, 0, out)
        self.assertIn("PASS: 30 cells checked, 29 match, 1 known, 0 unexplained", out)

    def test_wrong_reference_fails_even_though_its_saved_values_look_right(self):
        code, out = gate("examples/wrong_reference.xlsx")
        self.assertEqual(code, 1, out)
        self.assertIn("FAIL: 84 cells checked, 70 match, 12 known, 2 unexplained", out)
        self.assertIn("Club total!C3", out)

    def test_unsupported_function_fails_instead_of_counting_zero(self):
        code, out = gate("examples/unsupported_function.xlsx")
        self.assertEqual(code, 1, out)
        self.assertIn("FAIL: 84 cells checked, 66 match, 10 known, 8 unexplained", out)
        self.assertIn("ROUND", out)

    def test_missing_file_is_a_usage_error(self):
        code, out = gate("examples/nope.xlsx")
        self.assertEqual(code, 2, out)


class ExamplesAreReproducible(unittest.TestCase):
    def test_make_examples_rebuilds_the_committed_files(self):
        # Compare the unzipped parts, not the zip bytes: zip headers and
        # compression differ between operating systems.
        with tempfile.TemporaryDirectory() as tmp:
            subprocess.run([sys.executable, "-B", str(EXAMPLES / "make_examples.py"), tmp],
                           cwd=ROOT, check=True, capture_output=True)
            for name in ["good", "empty_subtotal", "wrong_reference", "unsupported_function"]:
                self.assertEqual(parts(Path(tmp) / f"{name}.xlsx"), parts(EXAMPLES / f"{name}.xlsx"), name)


class CommandLineOptions(unittest.TestCase):
    def test_show_limits_the_listed_cells(self):
        code, out = gate("examples/empty_subtotal.xlsx", "--show", "2")
        self.assertEqual(code, 1, out)
        self.assertIn("... and 6 more", out)

    def test_reason_of_a_known_difference_is_printed_once(self):
        code, out = gate("examples/good.xlsx")
        self.assertEqual(out.count("Sponsor payment of week 4"), 1, out)
        self.assertIn("[KD-1] file is off by -2000 as declared", out)


def parts(path):
    with zipfile.ZipFile(path) as z:
        return {name: z.read(name) for name in sorted(z.namelist())}


if __name__ == "__main__":
    unittest.main()
