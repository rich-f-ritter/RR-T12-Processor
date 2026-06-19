# Test fixtures

These are **real property financials** and are **git-ignored by default** (see
repo `.gitignore`) so they are not pushed to GitHub. The container is ephemeral,
so they must be re-supplied each session unless you choose to commit them.

## `dev/testdata/` — inputs to the build

| File | Role |
|------|------|
| `Canyon_Ridge__T12_2026.03.xlsx` | T12 operating statement (Apr 2025 – Mar 2026) |
| `Canyon_Ridge__T12_2026.05.xlsx` | T12 operating statement (Jun 2025 – May 2026) |
| `Canyon_Ridge__Rent_Roll_2026.05.13.xlsx` | rent roll, as-of 05/13 |
| `Canyon_Ridge__Rent_Roll_2026.05.31.xlsx` | rent roll, as-of 05/31 (newest) |
| `hellodata_unit_details_2026.06.17.csv` | HelloData "Unit Details" export |

With the full set, the harness reproduces the golden output to full value parity:

```
python3 dev/run.py \
  --t12 dev/testdata/Canyon_Ridge__T12_2026.03.xlsx dev/testdata/Canyon_Ridge__T12_2026.05.xlsx \
  --rr  dev/testdata/Canyon_Ridge__Rent_Roll_2026.05.31.xlsx \
  --hd  dev/testdata/hellodata_unit_details_2026.06.17.csv \
  --name "Canyon Ridge Apartments"
python3 dev/compare.py        # -> ✓ PARITY
```

## `dev/reference/` — context, not inputs

The RedIQ outputs we are replacing (operating statement + rent roll) and the
destination `TMG_Acquisition_Model` live here for reference. They are NOT consumed
by the build; they are the gold standard for categorization and the paste target.

## Committing fixtures

Everything here is git-ignored. To commit fixtures for cross-session
reproducibility, remove the matching lines from the repo `.gitignore` — only if
this repo is private and storing the financials there is acceptable.
