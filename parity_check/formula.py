"""Formula engine: the value Excel would show when it opens the file.

A generated file can carry stale saved values next to its formulas (and a
flag that tells Excel to recalculate everything on open). Reading the saved
values tells you nothing about what the user will see. So we recalculate.

Supported on purpose, and nothing more:
  - references A1, $A$1, Sheet!A1, 'Sheet name'!A1, ranges A1:B9
  - SUM over ranges, cells and numbers
  - + - * / with the usual precedence, parentheses, unary minus
  - empty, missing or text cells count as 0, like Excel's SUM

Everything else raises FormulaError. A function we do not know, a link to
another file, a cycle or a missing sheet would otherwise turn into a quiet
zero, which is exactly the kind of result this check exists to catch.
"""
import re

from .xlsx import Cell

_CELL = re.compile(r"\$?([A-Za-z]{1,3})\$?(\d+)")
_EXTERNAL = re.compile(r"\[[^\]]*\]")
_NUMBER = re.compile(r"\d+(\.\d+)?([eE][+-]?\d+)?")
_FUNCTION = re.compile(r"([A-Za-z][A-Za-z0-9.]*)\(")
_UNQUOTED_SHEET = re.compile(r"([A-Za-z_][A-Za-z0-9_.]*)!")


class FormulaError(Exception):
    """A cell whose value cannot be recalculated with confidence."""


def column_index(letters: str) -> int:
    """'A' -> 1, 'Z' -> 26, 'AA' -> 27."""
    n = 0
    for ch in letters.upper():
        n = n * 26 + ord(ch) - 64
    return n


def column_letters(index: int) -> str:
    """1 -> 'A', 27 -> 'AA'."""
    out = ""
    while index > 0:
        index, rem = divmod(index - 1, 26)
        out = chr(65 + rem) + out
    return out


def split_address(addr: str) -> tuple:
    """'$B$7' -> (2, 7) as (column, row), both 1-based."""
    m = _CELL.fullmatch(addr)
    if not m:
        raise FormulaError(f"not a cell address: {addr}")
    return column_index(m.group(1)), int(m.group(2))


def normalize(addr: str) -> str:
    col, row = split_address(addr)
    return f"{column_letters(col)}{row}"


def expand_range(start: str, end: str) -> list:
    c1, r1 = split_address(start)
    c2, r2 = split_address(end)
    return [
        f"{column_letters(c)}{r}"
        for r in range(min(r1, r2), max(r1, r2) + 1)
        for c in range(min(c1, c2), max(c1, c2) + 1)
    ]


class Engine:
    """Recalculates cells of an in-memory workbook: {sheet: {address: Cell}}."""

    def __init__(self, workbook: dict):
        self.workbook = workbook
        self._done: dict = {}
        self._running: set = set()

    def value(self, sheet: str, addr: str) -> float:
        addr = normalize(addr)
        key = (sheet, addr)
        if key in self._done:
            return self._done[key]
        if key in self._running:
            raise FormulaError(f"cycle of formulas through {sheet}!{addr}")
        if sheet not in self.workbook:
            raise FormulaError(f"sheet not found: {sheet}")
        self._running.add(key)
        try:
            cell = self.workbook[sheet].get(addr)
            result = self._cell_value(cell, sheet, addr)
        finally:
            self._running.discard(key)
        self._done[key] = result
        return result

    def _cell_value(self, cell, sheet: str, addr: str) -> float:
        if cell is None:
            return 0.0
        if cell.problem:
            raise FormulaError(f"{sheet}!{addr}: {cell.problem}")
        if cell.formula:
            return _Parser(self, cell.formula, sheet).run()
        if isinstance(cell.value, bool) or cell.value is None:
            return 0.0
        try:
            return float(cell.value)
        except (TypeError, ValueError):
            return 0.0  # text counts as 0, like SUM


class _Parser:
    """Recursive descent over one formula, evaluated as it is read."""

    def __init__(self, engine: Engine, formula: str, sheet: str):
        if _EXTERNAL.search(formula):
            raise FormulaError(f"link to an external file cannot be recalculated: ={formula}")
        self.engine = engine
        self.text = formula[1:] if formula.startswith("=") else formula
        self.sheet = sheet
        self.i = 0

    # --- helpers -----------------------------------------------------------
    def fail(self, message: str):
        raise FormulaError(f"{message} in ={self.text}")

    def skip_spaces(self):
        while self.i < len(self.text) and self.text[self.i].isspace():
            self.i += 1

    def peek(self) -> str:
        self.skip_spaces()
        return self.text[self.i] if self.i < len(self.text) else ""

    def run(self) -> float:
        result = self.expression()
        if self.peek():
            self.fail(f"unexpected text {self.text[self.i:]!r}")
        return result

    # --- grammar -----------------------------------------------------------
    def expression(self) -> float:
        value = self.term()
        while self.peek() in ("+", "-"):
            op = self.text[self.i]
            self.i += 1
            right = self.term()
            value = value + right if op == "+" else value - right
        return value

    def term(self) -> float:
        value = self.atom()
        while self.peek() in ("*", "/"):
            op = self.text[self.i]
            self.i += 1
            right = self.atom()
            if op == "*":
                value *= right
            elif right == 0:
                self.fail("division by zero")
            else:
                value /= right
        return value

    def atom(self) -> float:
        ch = self.peek()
        if ch == "(":
            self.i += 1
            value = self.expression()
            if self.peek() != ")":
                self.fail("missing closing parenthesis")
            self.i += 1
            return value
        if ch == "-":
            self.i += 1
            return -self.atom()
        if ch == "+":
            self.i += 1
            return self.atom()
        rest = self.text[self.i:]
        fn = _FUNCTION.match(rest)
        if fn:
            name = fn.group(1).upper()
            if name != "SUM":
                self.fail(f"function not supported: {name}()")
            self.i += len(fn.group(0))
            return sum(self.arguments())
        ref = self.reference()
        if ref is not None:
            sheet, addresses = ref
            return sum(self.engine.value(sheet, a) for a in addresses)
        num = _NUMBER.match(rest)
        if num:
            self.i += len(num.group(0))
            return float(num.group(0))
        self.fail(f"cannot read {rest[:16]!r}")

    def arguments(self) -> list:
        values = []
        while True:
            if self.peek() == ")":
                self.i += 1
                return values
            if not self.peek():
                self.fail("missing closing parenthesis")
            ref = self.reference()
            if ref is not None:
                sheet, addresses = ref
                values.extend(self.engine.value(sheet, a) for a in addresses)
            else:
                values.append(self.expression())
            if self.peek() in (",", ";"):
                self.i += 1

    def sheet_prefix(self):
        """Consume 'Sheet name'! or Sheet! if present; return the name or None."""
        start = self.i
        if self.peek() == "'":
            j = self.i + 1
            name = ""
            while j < len(self.text):
                if self.text[j] == "'":
                    if self.text[j + 1:j + 2] == "'":  # escaped quote
                        name += "'"
                        j += 2
                        continue
                    break
                name += self.text[j]
                j += 1
            if self.text[j + 1:j + 2] == "!":
                self.i = j + 2
                return name
            self.i = start
            return None
        m = _UNQUOTED_SHEET.match(self.text, self.i)
        if m:
            self.i = m.end()
            return m.group(1)
        return None

    def reference(self):
        """Consume a cell or a range; return (sheet, [addresses]) or None."""
        start = self.i
        sheet = self.sheet_prefix() or self.sheet
        first = _CELL.match(self.text, self.i)
        if not first:
            self.i = start
            return None
        end = first.end()
        if self.text[end:end + 1] == ":":
            second = _CELL.match(self.text, end + 1)
            if second:
                self.i = second.end()
                return sheet, expand_range(first.group(0), second.group(0))
        # "LOG10(" or a defined name must not be read as a cell reference
        if re.match(r"[A-Za-z0-9_(]", self.text[end:end + 1]):
            self.i = start
            return None
        self.i = end
        return sheet, [normalize(first.group(0))]
