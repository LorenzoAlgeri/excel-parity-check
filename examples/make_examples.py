"""Rebuild the four example files from expected.json.

    python examples/make_examples.py [output folder]

Every file starts from the same correct layout. The saved values of the
formulas always come from the correct file, so a tool that trusts saved
values would think all four files are fine. Only recalculation tells them
apart.
"""
import copy
import json
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parent))

from parity_check.gate import spec_from_dict  # noqa: E402
from parity_check.layout import build_workbook, fill_saved_values  # noqa: E402
from parity_check.xlsx import Cell, write_workbook  # noqa: E402


def main(out: Path) -> None:
    spec = spec_from_dict(json.loads((HERE / "expected.json").read_text("utf-8")))
    good = build_workbook(spec)

    # The declared difference (KD-1): one sponsor payment of week 4 is not in
    # the file yet. Saved values are refreshed so this file is fully coherent.
    good["Seniors"]["D5"] = Cell(value=0)
    fill_saved_values(good)

    files = {"good": good}

    # The real incident, rebuilt: inputs are right, one month total is empty.
    broken = copy.deepcopy(good)
    broken["Seniors"]["C6"] = Cell()
    files["empty_subtotal"] = broken

    # Club total, Events, Week 2 adds the Seniors cell of week 1 by mistake.
    broken = copy.deepcopy(good)
    broken["Club total"]["C3"] = Cell(formula="Youth!C3+Seniors!C2", value=good["Club total"]["C3"].value)
    files["wrong_reference"] = broken

    # A function the engine does not support must stop the gate.
    broken = copy.deepcopy(good)
    broken["Youth"]["B6"] = Cell(formula="ROUND(SUM(B2:B5),0)", value=good["Youth"]["B6"].value)
    files["unsupported_function"] = broken

    for name, workbook in files.items():
        write_workbook(out / f"{name}.xlsx", workbook)
        print(f"wrote {name}.xlsx")


if __name__ == "__main__":
    main(Path(sys.argv[1]) if len(sys.argv) > 1 else HERE)
