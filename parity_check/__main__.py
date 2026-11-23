"""Command line: python -m parity_check FILE.xlsx [options]

Exit code 0 = PASS, 1 = FAIL, 2 = the command could not run.
"""
import argparse
import json
import sys
from pathlib import Path

from .gate import known_from_list, run_gate, spec_from_dict
from .xlsx import read_workbook


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(prog="python -m parity_check",
                                     description="Check a generated .xlsx file against the values it should show.")
    parser.add_argument("file", type=Path, help="the .xlsx file to check")
    parser.add_argument("--expected", type=Path,
                        help="JSON with the expected values (default: expected.json next to the file)")
    parser.add_argument("--known", type=Path,
                        help="JSON list of declared differences (default: known_differences.json next to the file)")
    parser.add_argument("--inputs-only", action="store_true",
                        help="only read the input cells, recalculate nothing (the old, weaker check)")
    parser.add_argument("--show", type=int, default=20, help="how many cells to list per class (default 20)")
    args = parser.parse_args(argv)

    expected_path = args.expected or args.file.parent / "expected.json"
    known_path = args.known or args.file.parent / "known_differences.json"
    try:
        workbook = read_workbook(args.file)
        spec = spec_from_dict(json.loads(expected_path.read_text("utf-8")))
        known = known_from_list(json.loads(known_path.read_text("utf-8"))) if known_path.exists() else []
    except (OSError, ValueError, KeyError) as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 2

    report = run_gate(workbook, spec, known, inputs_only=args.inputs_only)

    print(f"file:     {args.file.as_posix()}")
    print(f"expected: {expected_path.as_posix()}")
    print(f"known:    {known_path.as_posix() if known_path.exists() else '(none)'} ({len(known)} declared)")
    print(f"mode:     {'inputs only, no recalculation' if args.inputs_only else 'full recalculation'}")
    print()
    for status in ("UNEXPLAINED", "KNOWN"):
        rows = [c for c in report.cells if c.status == status]
        if not rows:
            continue
        print(f"{status} ({len(rows)}):")
        for c in rows[:args.show]:
            where = f"{c.sheet}!{c.ref}" if c.ref else c.sheet
            # a known difference: show its id here, its reason once below
            note = c.reason.split("]")[0] + "]" if status == "KNOWN" else c.reason
            print(f"  {where:<16} {c.column} / {c.row}: {note}")
        if len(rows) > args.show:
            print(f"  ... and {len(rows) - args.show} more")
        if status == "KNOWN":
            used = {c.reason.split("]")[0].lstrip("[") for c in rows}
            for k in known:
                if k.id in used:
                    print(f"  [{k.id}] file is off by {k.delta:g} as declared: {k.reason}")
        print()
    print(report.summary())
    return 0 if report.verdict == "PASS" else 1


if __name__ == "__main__":
    sys.exit(main())
