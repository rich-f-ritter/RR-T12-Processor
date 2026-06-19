#!/usr/bin/env python3
"""
dev/run.py — one-command build + recalc for the rediq-replacement skill.

Builds the intake workbook from the Canyon Ridge fixtures (or any files you
pass through), then runs the mandatory recalc step exactly as the skill spec
requires. This is the local dev loop; the shipped skill is unchanged.

Usage:
    python3 dev/run.py                      # default Canyon Ridge fixtures
    python3 dev/run.py --rr <rentroll.xlsx> --t12 <a.xlsx> <b.xlsx> ...
    python3 dev/run.py --no-recalc          # build only (skip recalc)

Recalc resolution order:
    1. $RECALC_PY if set
    2. /mnt/skills/public/xlsx/scripts/recalc.py   (Claude app / this env)
    3. error with instructions
"""
import argparse, os, subprocess, sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SKILL = ROOT / "rediq-replacement"
TESTDATA = ROOT / "dev" / "testdata"
OUT = ROOT / "dev" / "out"


def find_recalc():
    env = os.environ.get("RECALC_PY")
    if env and Path(env).exists():
        return env
    canonical = "/mnt/skills/public/xlsx/scripts/recalc.py"
    if Path(canonical).exists():
        return canonical
    return None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--t12", nargs="+",
                    default=[str(TESTDATA / "Canyon_Ridge__T12_2026.03.xlsx")])
    ap.add_argument("--rr",
                    default=str(TESTDATA / "Canyon_Ridge__Rent_Roll_2026.05.31.xlsx"))
    ap.add_argument("--hd", default=None)
    ap.add_argument("--name", default="Canyon Ridge")
    ap.add_argument("--out", default=str(OUT / "Canyon_Ridge__Underwriting_Intake.xlsx"))
    ap.add_argument("--no-recalc", action="store_true")
    ap.add_argument("--timeout", default="120")
    args = ap.parse_args()

    OUT.mkdir(parents=True, exist_ok=True)

    build = [sys.executable, str(SKILL / "scripts" / "build_intake.py"),
             "--t12", *args.t12, "--rr", args.rr, "--name", args.name,
             "--out", args.out]
    if args.hd:
        build += ["--hd", args.hd]

    print("→ build:", " ".join(build))
    r = subprocess.run(build)
    if r.returncode != 0:
        sys.exit(r.returncode)

    if args.no_recalc:
        print("✓ built (recalc skipped):", args.out)
        return

    recalc = find_recalc()
    if not recalc:
        print("✗ recalc.py not found. Set $RECALC_PY or run in an environment "
              "with /mnt/skills/public/xlsx/scripts/recalc.py.", file=sys.stderr)
        sys.exit(2)

    print("→ recalc:", recalc, args.out, args.timeout)
    r = subprocess.run([sys.executable, recalc, args.out, args.timeout])
    sys.exit(r.returncode)


if __name__ == "__main__":
    main()
