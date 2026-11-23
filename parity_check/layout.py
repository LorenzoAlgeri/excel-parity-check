"""The shape of the example file, and a builder for a correct copy of it.

Every sheet looks like this (one column per budget line, one Total column):

          A                    B          C        D          E
    1                          Fees       Events   Sponsors   Total
    2     Week 1               (input)    ...                 =SUM(B2:D2)
    ...   Week N
    N+2   Month total          =SUM(B2:B(N+1))
    N+3   Budget               (input)
    N+4   Variance vs budget   =B(N+2)-B(N+3)

An area sheet holds inputs. A totals sheet has the same shape, but every
line cell is a formula that adds the same cell of each area sheet.

The builder stands in for the real exporter: the tests and the example
files start from a correct copy and then break one thing on purpose.
"""
from .formula import Engine, column_letters
from .xlsx import Cell

MONTH_TOTAL = "Month total"
BUDGET = "Budget"
VARIANCE = "Variance vs budget"
TOTAL = "Total"


def week_label(n: int) -> str:
    return f"Week {n}"


def row_labels(weeks: int) -> list:
    return [week_label(n) for n in range(1, weeks + 1)] + [MONTH_TOTAL, BUDGET, VARIANCE]


def build_workbook(spec) -> dict:
    """Correct workbook for a Spec, with saved values filled in."""
    lines = spec.lines
    labels = row_labels(spec.weeks)
    row_of = {label: i + 2 for i, label in enumerate(labels)}
    first_week, last_week = row_of[week_label(1)], row_of[week_label(spec.weeks)]
    total_col = column_letters(len(lines) + 2)
    last_line_col = column_letters(len(lines) + 1)

    def frame() -> dict:
        cells = {"A1": Cell(value="")}
        for i, line in enumerate(lines):
            cells[f"{column_letters(i + 2)}1"] = Cell(value=line)
        cells[f"{total_col}1"] = Cell(value=TOTAL)
        for label, r in row_of.items():
            cells[f"A{r}"] = Cell(value=label)
            cells[f"{total_col}{r}"] = Cell(formula=f"SUM(B{r}:{last_line_col}{r})")
        return cells

    workbook = {}
    for sheet, data in spec.sheets.items():
        cells = frame()
        for i, line in enumerate(lines):
            col = column_letters(i + 2)
            for n, amount in enumerate(data[line].weeks, start=1):
                cells[f"{col}{row_of[week_label(n)]}"] = Cell(value=amount)
            cells[f"{col}{row_of[MONTH_TOTAL]}"] = Cell(formula=f"SUM({col}{first_week}:{col}{last_week})")
            cells[f"{col}{row_of[BUDGET]}"] = Cell(value=data[line].budget)
            cells[f"{col}{row_of[VARIANCE]}"] = Cell(formula=f"{col}{row_of[MONTH_TOTAL]}-{col}{row_of[BUDGET]}")
        workbook[sheet] = cells

    for sheet, parts in spec.totals.items():
        cells = frame()
        for i in range(len(lines)):
            col = column_letters(i + 2)
            for r in row_of.values():
                cells[f"{col}{r}"] = Cell(formula="+".join(f"{_quote(p)}!{col}{r}" for p in parts))
        workbook[sheet] = cells

    fill_saved_values(workbook)
    return workbook


def fill_saved_values(workbook: dict) -> None:
    """Store the current result of each formula as its saved value."""
    engine = Engine(workbook)
    for sheet, cells in workbook.items():
        for addr, cell in cells.items():
            if cell.formula:
                cell.value = engine.value(sheet, addr)


def _quote(sheet: str) -> str:
    if sheet.replace("_", "").isalnum():
        return sheet
    return "'" + sheet.replace("'", "''") + "'"
