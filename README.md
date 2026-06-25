# RR-T12-Processor

Development home for the **`rr-t12-processor`** Claude skill (display name
"RR-T12 Processor") — a RedIQ replacement for multifamily underwriting intake.
It turns raw operating statements (T12s /
monthlies) and a rent roll into a standardized, editable chart of accounts, a
stitched multi-period operating history, a one-lined rent roll, an enhanced unit
mix, lease-trend analysis, and a rent-roll-to-T12 reconciliation — all paste-ready
for the TMG acquisition model.

The skill is developed and round-trip tested here, then packaged and shipped to the
Claude app (where it runs against the public `xlsx` recalc helper and the sibling
`multifamily-deal-analysis` methodology skill).

## Layout

```
.claude/skills/
  rr-t12-processor/         ← the skill (auto-loads in Claude Code web sessions on this repo)
    SKILL.md
    references/             ← account_mapping.md, hd_fee_detection.md, model_paste_targets.md
    scripts/                ← account_map.py, intake_lib.py, build_intake.py
dev/                        ← local dev harness (NOT part of the shipped skill)
  run.py                    ← build + recalc in one command
  recalc_local.py           ← deterministic, LibreOffice-free formula evaluator
  compare.py                ← regression: fresh build vs golden, cell-by-cell
  testdata/                 ← input fixtures (git-ignored; real financials)
  reference/                ← RedIQ outputs + TMG model (git-ignored; context)
  golden/                   ← known-good output to diff against (git-ignored)
  out/                      ← generated workbooks (git-ignored)
```

## Install / use the skill

**Claude Code on the web — two ways:**

- **Repo-scoped (zero setup):** the skill lives at `.claude/skills/rr-t12-processor/` on
  `main`, so any web session started on this repo loads it automatically. Just upload a
  T12 + rent roll (+ HelloData CSV) and ask for the underwriting intake.
- **Account-wide:** upload the packaged skill zip in **claude.ai → Customize → Skills**
  and enable it; it then loads in every cloud session, on any repo. (`SKILL.md` sits at
  the zip root, as the uploader expects.) Cut a fresh zip from the latest `main` with:

  ```bash
  bash dev/package_skill.sh      # -> rr-t12-processor-skill.zip (SKILL.md at root)
  ```

  Re-upload that whenever you want the account-wide copy to catch up to `main`.

Note: `/plugin` is **not** available in Claude Code on the web — the `.claude/skills/`
location above is the correct web mechanism (desktop/CLI users could alternatively wrap
it as a plugin).

Runtime needs: Python 3.8+, `openpyxl`, and LibreOffice (for the mandatory recalc via the
public `xlsx` skill's `recalc.py`). The `dev/` harness is for maintaining the skill and is
not required to run it.



## Dev loop

```bash
# build the intake workbook from the Canyon Ridge fixtures + recalc
python3 dev/run.py

# check the result against the golden workbook (exits non-zero on drift)
python3 dev/compare.py        # -> ✓ PARITY
```

Pass your own inputs with `--t12 a.xlsx b.xlsx --rr rr.xlsx [--hd hello.csv]`.

### Recalc note

The skill's mandatory recalc step (`/mnt/skills/public/xlsx/scripts/recalc.py`,
LibreOffice) is what runs in the **Claude app** and is the production path. In this
**sandbox** that LibreOffice path is unreliable — the recalc macro hangs and the
script reports success off an unmodified file — so the dev harness defaults to
`recalc_local.py`, a small Python evaluator covering the exact formula vocabulary
the skill emits (`SUM`, `SUMIFS`, `VLOOKUP`, `IFERROR`, `IF`, `OR`, arithmetic). It
reproduces the golden's LibreOffice-computed values exactly (~21k cells, <1s).

- `python3 dev/run.py --recalc local`  (default) — Python evaluation, works anywhere
- `python3 dev/run.py --recalc office` — the shipped LibreOffice path (flaky here)
- `dev/compare.py` evaluates the fresh workbook via `recalc_local` and diffs it
  against the golden's cached values.

## Packaging for the Claude app

The skill ships as the `rr-t12-processor/` directory only. The `dev/` harness,
fixtures, and golden output stay in the repo and are not part of the package.

See `dev/testdata/README.md` for which input fixtures the harness expects and what
is currently missing to reproduce the golden output exactly.
