"""The parity gate: compare a generated workbook with the values it should show.

For every checked cell:
  expected  comes from the source data (here a JSON file, in the real system
            a database), never from the file;
  actual    is the value Excel would show when it opens the file, recalculated
            by our own engine, never the value saved in the file;
  status    is MATCH (within one cent), KNOWN (a declared difference with a
            written cause and the exact amount) or UNEXPLAINED.

One UNEXPLAINED cell fails the gate. A cell we cannot recalculate is
UNEXPLAINED too: if nobody knows what the file will show, it does not pass.
"""
from dataclasses import dataclass, field
from typing import Optional

from .formula import Engine, FormulaError
from .layout import BUDGET, MONTH_TOTAL, TOTAL, VARIANCE, row_labels, week_label

TOLERANCE = 0.01


# ---------------------------------------------------------------------------
# Inputs: the expected data and the declared differences
# ---------------------------------------------------------------------------

@dataclass
class Line:
    weeks: list
    budget: float


@dataclass
class Spec:
    sheets: dict            # area sheet -> {line name -> Line}
    totals: dict            # totals sheet -> [area sheets it adds up]
    lines: list = field(default_factory=list)
    weeks: int = 0


def spec_from_dict(data: dict) -> Spec:
    sheets = {
        sheet: {name: Line(list(v["weeks"]), v["budget"]) for name, v in lines.items()}
        for sheet, lines in data["sheets"].items()
    }
    first = next(iter(sheets.values()))
    spec = Spec(sheets=sheets, totals=dict(data.get("totals", {})),
                lines=list(first), weeks=len(next(iter(first.values())).weeks))
    for sheet, lines in sheets.items():
        if list(lines) != spec.lines:
            raise ValueError(f"sheet {sheet!r} must list the same lines in the same order")
        for name, line in lines.items():
            if len(line.weeks) != spec.weeks:
                raise ValueError(f"{sheet}/{name}: expected {spec.weeks} weeks")
    for sheet, parts in spec.totals.items():
        missing = [p for p in parts if p not in sheets]
        if missing:
            raise ValueError(f"totals sheet {sheet!r} adds unknown sheets {missing}")
    return spec


@dataclass
class KnownDifference:
    """A difference someone investigated. It covers a cell only if the cell
    sits in these sheets, columns and rows AND is off by exactly `delta`
    (file minus expected). A new, different error in the same cell still fails."""
    id: str
    sheets: list
    columns: list
    rows: list
    delta: float
    reason: str

    def covers(self, sheet, column, row, expected, actual) -> bool:
        return (sheet in self.sheets and column in self.columns and row in self.rows
                and actual is not None and abs((actual - expected) - self.delta) <= TOLERANCE)


def known_from_list(items: list) -> list:
    out = []
    for item in items:
        if not str(item.get("reason", "")).strip():
            raise ValueError(f"known difference {item.get('id')!r} has no reason: write down the cause")
        out.append(KnownDifference(item["id"], list(item["sheets"]), list(item["columns"]),
                                   list(item["rows"]), float(item["delta"]), item["reason"]))
    return out


# ---------------------------------------------------------------------------
# Classification
# ---------------------------------------------------------------------------

@dataclass
class Verdict:
    status: str
    reason: str = ""


def classify(expected: float, actual: Optional[float], error: Optional[str],
             known: Optional[KnownDifference]) -> Verdict:
    if error:
        return Verdict("UNEXPLAINED", f"cannot recalculate: {error}")
    if abs(actual - expected) <= TOLERANCE:
        return Verdict("MATCH")
    if known is not None:
        return Verdict("KNOWN", f"[{known.id}] {known.reason}")
    return Verdict("UNEXPLAINED", f"expected {_fmt(expected)}, file shows {_fmt(actual)}")


def _fmt(x: float) -> str:
    return f"{x:,.2f}".rstrip("0").rstrip(".")


# ---------------------------------------------------------------------------
# The gate
# ---------------------------------------------------------------------------

@dataclass
class CheckedCell:
    sheet: str
    column: str
    row: str
    ref: str
    expected: float
    actual: Optional[float]
    status: str
    reason: str = ""


@dataclass
class Report:
    cells: list
    inputs_only: bool = False

    @property
    def counts(self) -> dict:
        counts = {"MATCH": 0, "KNOWN": 0, "UNEXPLAINED": 0}
        for c in self.cells:
            counts[c.status] += 1
        return counts

    @property
    def verdict(self) -> str:
        return "FAIL" if self.counts["UNEXPLAINED"] else "PASS"

    def summary(self) -> str:
        c = self.counts
        return (f"{self.verdict}: {len(self.cells)} cells checked, {c['MATCH']} match, "
                f"{c['KNOWN']} known, {c['UNEXPLAINED']} unexplained")


def expected_value(spec: Spec, sheet: str, column: str, row: str) -> float:
    """What the cell must show, computed from the source data only."""
    if sheet in spec.totals:
        return sum(expected_value(spec, part, column, row) for part in spec.totals[sheet])
    if column == TOTAL:
        return sum(expected_value(spec, sheet, line, row) for line in spec.lines)
    line = spec.sheets[sheet][column]
    if row == MONTH_TOTAL:
        return sum(line.weeks)
    if row == BUDGET:
        return line.budget
    if row == VARIANCE:
        return sum(line.weeks) - line.budget
    return line.weeks[int(row.split()[-1]) - 1]


def _find_labels(cells: dict) -> tuple:
    """Rows by their label in column A, columns by their header in row 1."""
    rows, cols = {}, {}
    for addr, cell in cells.items():
        if not isinstance(cell.value, str) or cell.formula:
            continue
        letters = addr.rstrip("0123456789")
        number = int(addr[len(letters):])
        if letters == "A" and number > 1:
            rows.setdefault(cell.value.strip(), number)
        elif number == 1 and letters != "A":
            cols.setdefault(cell.value.strip(), letters)
    return rows, cols


def run_gate(workbook: dict, spec: Spec, known: Optional[list] = None,
             inputs_only: bool = False) -> Report:
    """Check every cell the spec describes.

    inputs_only=True reproduces the old, weaker check: it reads only the
    literal input cells of the area sheets and recalculates nothing.
    """
    known = known or []
    engine = Engine(workbook)
    input_rows = [week_label(n) for n in range(1, spec.weeks + 1)] + [BUDGET]
    cells = []

    sheets = list(spec.sheets) + ([] if inputs_only else list(spec.totals))
    for sheet in sheets:
        if inputs_only:
            columns, rows = spec.lines, input_rows
        else:
            columns, rows = spec.lines + [TOTAL], row_labels(spec.weeks)
        found_rows, found_cols = _find_labels(workbook.get(sheet, {}))

        for row in rows:
            for column in columns:
                expected = expected_value(spec, sheet, column, row)
                ref, actual, error = "", None, None
                if sheet not in workbook:
                    error = f"sheet '{sheet}' not found"
                elif row not in found_rows:
                    error = f"row '{row}' not found in column A"
                elif column not in found_cols:
                    error = f"column '{column}' not found in row 1"
                else:
                    ref = f"{found_cols[column]}{found_rows[row]}"
                    try:
                        actual = engine.value(sheet, ref)
                    except FormulaError as exc:
                        error = str(exc)
                match = next((k for k in known if k.covers(sheet, column, row, expected, actual)), None)
                verdict = classify(expected, actual, error, match)
                cells.append(CheckedCell(sheet, column, row, ref, expected, actual,
                                         verdict.status, verdict.reason))
    return Report(cells, inputs_only)
