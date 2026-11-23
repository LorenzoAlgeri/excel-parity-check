"""Read and write .xlsx files with the standard library only.

An .xlsx file is a zip of XML parts. We need very little of it: sheet names,
and for each cell its formula and its saved value. Keeping this in the
standard library means the project has no dependencies, and lets the writer
put a stale saved value next to a formula, which the tests need.
"""
import posixpath
import re
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Optional, Union
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape

MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
_NS = {"m": MAIN}


@dataclass
class Cell:
    """One cell as stored in the file.

    formula: formula text without the leading "=", or None.
    value:   the literal value, or for a formula the value SAVED in the file
             (which can be stale; the gate never trusts it).
    problem: set by the reader when it could not read the cell faithfully.
             The formula engine refuses such a cell instead of guessing.
    """
    formula: Optional[str] = None
    value: Union[float, str, bool, None] = None
    problem: Optional[str] = None


# ---------------------------------------------------------------------------
# Reading
# ---------------------------------------------------------------------------

def _text(node) -> str:
    """All <t> text under a node (handles rich text runs)."""
    return "".join(t.text or "" for t in node.iter(f"{{{MAIN}}}t"))


def _number(raw: str):
    n = float(raw)
    return int(n) if n.is_integer() else n


def read_workbook(path) -> dict:
    """Return {sheet name: {address: Cell}}, sheets in workbook order."""
    with zipfile.ZipFile(path) as z:
        names = set(z.namelist())
        shared = []
        if "xl/sharedStrings.xml" in names:
            root = ET.fromstring(z.read("xl/sharedStrings.xml"))
            shared = [_text(si) for si in root.findall("m:si", _NS)]

        rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        targets = {r.get("Id"): r.get("Target") for r in rels}
        book = ET.fromstring(z.read("xl/workbook.xml"))

        workbook = {}
        for sheet in book.find("m:sheets", _NS):
            target = targets[sheet.get(f"{{{REL}}}id")]
            part = target.lstrip("/") if target.startswith("/") else posixpath.normpath(f"xl/{target}")
            workbook[sheet.get("name")] = _read_sheet(ET.fromstring(z.read(part)), shared)
        return workbook


def _read_sheet(root, shared: list) -> dict:
    cells = {}
    for c in root.iter(f"{{{MAIN}}}c"):
        kind = c.get("t", "n")
        f = c.find("m:f", _NS)
        v = c.find("m:v", _NS)
        raw = v.text if v is not None else None

        if kind == "s" and raw is not None:
            value = shared[int(raw)]
        elif kind == "inlineStr":
            node = c.find("m:is", _NS)
            value = _text(node) if node is not None else ""
        elif kind in ("str", "e"):
            value = raw
        elif kind == "b":
            value = raw == "1"
        else:
            value = _number(raw) if raw not in (None, "") else None

        cell = Cell(value=value)
        if f is not None:
            if f.text:
                cell.formula = f.text
            else:
                # A copy of a "shared formula": the text lives on another
                # cell with relative references to shift. We do not shift
                # them; the engine will refuse this cell.
                cell.problem = "shared formula (copied down in Excel) is not expanded by this reader"
        cells[c.get("r")] = cell
    return cells


# ---------------------------------------------------------------------------
# Writing (used to build the example files and the test files)
# ---------------------------------------------------------------------------

_ADDR = re.compile(r"([A-Z]+)(\d+)")


def _col(letters: str) -> int:
    n = 0
    for ch in letters:
        n = n * 26 + ord(ch) - 64
    return n


def write_workbook(path, sheets: dict) -> None:
    """Write {sheet name: {address: Cell}} to a minimal, valid .xlsx file.

    A formula cell keeps the value you give it as its SAVED value, even if it
    is wrong. The workbook asks Excel to recalculate everything on open, as
    a generated file usually does.
    """
    names = list(sheets)
    parts = {
        "[Content_Types].xml": _content_types(len(names)),
        "_rels/.rels": f'<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="{PKG_REL}">'
                       f'<Relationship Id="rId1" Type="{REL}/officeDocument" Target="xl/workbook.xml"/>'
                       "</Relationships>",
        "xl/workbook.xml": _workbook_xml(names),
        "xl/_rels/workbook.xml.rels": _workbook_rels(len(names)),
    }
    for i, name in enumerate(names, start=1):
        parts[f"xl/worksheets/sheet{i}.xml"] = _sheet_xml(sheets[name])

    Path(path).parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for name, text in parts.items():
            info = zipfile.ZipInfo(name, date_time=(2026, 1, 1, 0, 0, 0))  # same bytes on every run
            info.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(info, text)


def _content_types(count: int) -> str:
    sheets = "".join(
        f'<Override PartName="/xl/worksheets/sheet{i}.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        for i in range(1, count + 1)
    )
    return (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        f"{sheets}</Types>"
    )


def _workbook_xml(names: list) -> str:
    sheets = "".join(
        f'<sheet name="{escape(n, {chr(34): "&quot;"})}" sheetId="{i}" r:id="rId{i}"/>'
        for i, n in enumerate(names, start=1)
    )
    return (
        f'<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="{MAIN}" xmlns:r="{REL}">'
        f'<sheets>{sheets}</sheets><calcPr fullCalcOnLoad="1"/></workbook>'
    )


def _workbook_rels(count: int) -> str:
    rels = "".join(
        f'<Relationship Id="rId{i}" Type="{REL}/worksheet" Target="worksheets/sheet{i}.xml"/>'
        for i in range(1, count + 1)
    )
    return f'<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="{PKG_REL}">{rels}</Relationships>'


def _sheet_xml(cells: dict) -> str:
    rows: dict = {}
    for addr, cell in cells.items():
        letters, row = _ADDR.fullmatch(addr).groups()
        rows.setdefault(int(row), []).append((_col(letters), addr, cell))

    out = []
    for r in sorted(rows):
        parts = []
        for _, addr, cell in sorted(rows[r]):
            parts.append(_cell_xml(addr, cell))
        out.append(f'<row r="{r}">{"".join(parts)}</row>')
    return (f'<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="{MAIN}">'
            f'<sheetData>{"".join(out)}</sheetData></worksheet>')


def _cell_xml(addr: str, cell: Cell) -> str:
    v = cell.value
    if cell.formula:
        saved = "" if v is None else f"<v>{escape(str(v))}</v>"
        kind = ' t="str"' if isinstance(v, str) else ""
        return f'<c r="{addr}"{kind}><f>{escape(cell.formula)}</f>{saved}</c>'
    if v is None:
        return f'<c r="{addr}"/>'
    if isinstance(v, str):
        return f'<c r="{addr}" t="inlineStr"><is><t xml:space="preserve">{escape(v)}</t></is></c>'
    return f'<c r="{addr}"><v>{v}</v></c>'
