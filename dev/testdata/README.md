# Test fixtures

These are **real property financials** and are **git-ignored by default** (see
repo `.gitignore`) so they are not pushed to GitHub. The container is ephemeral,
so they must be re-supplied each session unless you choose to commit them.

The dev harness (`dev/run.py`, `dev/compare.py`) expects:

| File | Role |
|------|------|
| `Canyon_Ridge__T12_2026.03.xlsx` | T12 operating statement (Apr 2025 – Mar 2026) |
| `Canyon_Ridge__Rent_Roll_2026.05.13.xlsx` | rent roll, as-of 05/13 |
| `Canyon_Ridge__Rent_Roll_2026.05.31.xlsx` | rent roll, as-of 05/31 (newest) |

The golden output in `../golden/` was built from a **larger** input set than the
above and currently CANNOT be reproduced bit-for-bit from these fixtures alone:

- a **second T12** covering through **May 2026** (`...T12_2026.05.xlsx`) — needed
  for the 14-month stitch the golden shows; and
- a **HelloData "Unit Details" CSV** — needed for the `HelloData` tab and the
  T90/T365 market-rent reads.

Drop those two files in here to enable full-parity regression against the golden.

To commit fixtures for cross-session reproducibility, remove the matching lines
from `.gitignore` (only if this repo is private and the data may be stored there).
