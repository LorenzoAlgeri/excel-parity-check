"""Small helpers shared by the tests."""
from parity_check.xlsx import Cell


def workbook(cells: dict, sheets=None) -> dict:
    """Build an in-memory workbook from {"Sheet!A1": value}.

    A string that starts with "=" is a formula, any other string is text,
    a number is a number. `sheets` lists sheets that exist but may be empty.
    """
    wb: dict = {name: {} for name in (sheets or [])}
    for key, raw in cells.items():
        sheet, addr = key.rsplit("!", 1)
        if isinstance(raw, str) and raw.startswith("="):
            cell = Cell(formula=raw[1:])
        else:
            cell = Cell(value=raw)
        wb.setdefault(sheet, {})[addr] = cell
    return wb
