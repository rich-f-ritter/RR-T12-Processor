#!/usr/bin/env python3
"""
dev/compare.py — regression check: freshly built workbook vs a golden workbook.

Compares every sheet shared by both files cell-by-cell on computed VALUES
(data_only), reporting:
  - sheets present in only one file
  - per-sheet dimension differences
  - numeric cells differing beyond --tol (absolute) / --rtol (relative)
  - text cells that differ

Exits non-zero if any difference exceeds tolerance, so it can gate CI.

Usage:
    python3 dev/compare.py [fresh.xlsx] [golden.xlsx] [--tol 0.5] [--rtol 0.001]
"""
import argparse, sys
from pathlib import Path
import openpyxl

ROOT = Path(__file__).resolve().parent.parent
DEF_FRESH = ROOT / "dev" / "out" / "Canyon_Ridge__Underwriting_Intake.xlsx"
DEF_GOLD = ROOT / "dev" / "golden" / "Canyon_Ridge__Underwriting_Intake.golden.xlsx"


def load(p):
    return openpyxl.load_workbook(p, data_only=True)


def num(v):
    return v if isinstance(v, (int, float)) else None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("fresh", nargs="?", default=str(DEF_FRESH))
    ap.add_argument("golden", nargs="?", default=str(DEF_GOLD))
    ap.add_argument("--tol", type=float, default=0.5, help="abs tolerance for numbers")
    ap.add_argument("--rtol", type=float, default=0.001, help="rel tolerance for numbers")
    ap.add_argument("--max", type=int, default=40, help="max diffs to print per sheet")
    args = ap.parse_args()

    a, b = load(args.fresh), load(args.golden)
    print(f"fresh : {args.fresh}")
    print(f"golden: {args.golden}\n")

    sa, sb = set(a.sheetnames), set(b.sheetnames)
    if sa - sb:
        print("  sheets only in FRESH :", sorted(sa - sb))
    if sb - sa:
        print("  sheets only in GOLDEN:", sorted(sb - sa))

    total_diffs = 0
    for sn in a.sheetnames:
        if sn not in sb:
            continue
        wa, wb = a[sn], b[sn]
        rows = max(wa.max_row, wb.max_row)
        cols = max(wa.max_column, wb.max_column)
        if (wa.max_row, wa.max_column) != (wb.max_row, wb.max_column):
            print(f"[{sn}] dims fresh={wa.max_row}x{wa.max_column} "
                  f"golden={wb.max_row}x{wb.max_column}")
        diffs = []
        for r in range(1, rows + 1):
            for c in range(1, cols + 1):
                va, vb = wa.cell(r, c).value, wb.cell(r, c).value
                na, nb = num(va), num(vb)
                if na is not None and nb is not None:
                    d = abs(na - nb)
                    if d > args.tol and d > abs(nb) * args.rtol:
                        diffs.append((r, c, va, vb))
                elif (va or "") != (vb or "") and (na is None and nb is None):
                    diffs.append((r, c, va, vb))
        if diffs:
            total_diffs += len(diffs)
            print(f"[{sn}] {len(diffs)} differing cell(s):")
            for r, c, va, vb in diffs[:args.max]:
                cell = openpyxl.utils.get_column_letter(c) + str(r)
                print(f"    {cell}: fresh={va!r}  golden={vb!r}")
            if len(diffs) > args.max:
                print(f"    ... +{len(diffs) - args.max} more")

    print()
    if total_diffs == 0 and not (sa ^ sb):
        print("✓ PARITY: fresh matches golden within tolerance.")
        sys.exit(0)
    print(f"✗ {total_diffs} cell diff(s); sheet set "
          f"{'matches' if not (sa ^ sb) else 'DIFFERS'}.")
    sys.exit(1)


if __name__ == "__main__":
    main()
