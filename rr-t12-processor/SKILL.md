---
name: rr-t12-processor
description: "RR-T12 Processor — replicate and improve on RedIQ for multifamily underwriting intake — turn one or more raw operating statements (T12s/monthlies) and a rent roll into a standardized, EDITABLE chart of accounts plus a stitched multi-period operating history, one-lined rent roll, enhanced unit mix (new-lease + HelloData T90 market-rent indicators), lease-trend analysis, and a rent-roll-to-T12 reconciliation with AGPR tie-out — all paste-ready for the TMG model. Use whenever the user uploads T12s/operating statements and/or a rent roll and wants them standardized, categorized, stitched, one-lined, or reconciled (the work RedIQ does before underwriting). Trigger on: 'run RedIQ', 'replace RedIQ', 'standardize this T12', 'stitch these statements', 'categorize the operating statement', 'one-line the rent roll', 'unit mix', 'reconcile rent roll to T12', 'underwriting intake', or an uploaded Yardi/RealPage/Entrata statement + rent roll. Produces the intake workbook only; does NOT populate the model."
---

# RR-T12 Processor — Underwriting Intake

Turn raw operating statements and a rent roll into the standardized, paste-ready inputs
an analyst needs to start underwriting — the job RedIQ does, but with the full statement
visible, **every line's category editable in one place**, the standardized OS Summary
rolling up **live** as codes change, and **multiple statements stitched into one
continuous operating history**.

Read the **`multifamily-deal-analysis`** skill first — it defines the methodology (RedIQ
code definitions, contract-rent = AGPR terminology, market-rent data hierarchy,
expense-to-income netting) that this skill operationalizes. This skill **stops at
intake**: it produces the workbook and the model paste targets; it does **not** populate
the underwriting model. Deep submarket comp work belongs to
`submarket-comp-template-execution`.

## Required inputs

1. **One or more T12s / monthly operating statements** (.xlsx or binary .xls) —
   Yardi/RealPage/Entrata native exports. Upload **as many periods as you have**: years of
   T12s, a couple of overlapping T12s, or a string of monthly statements. They are
   **stitched into one continuous monthly series** over the union of all their months (see
   *Stitching*). Prefer the **most detailed** versions (separate GL lines for concessions,
   loss-to-lease, RUBS, payroll burden, insurance); a summary statement still reconciles
   at EGR/Opex/NOI but flattens those sub-codes to zero for the months it owns.
   Full GL / trial-balance exports are fine too: **non-operating sections** (Debt Service /
   interest & principal, Depreciation & Amortization, Partnership/Owner, and grab-bag
   "Adjustments" carrying distributions, contributions, loan & mortgage balances) are routed
   to the non-op codes so they stay **out of NOI** — never operating G&A or vacancy.
2. **Rent roll(s)** (.xlsx or binary .xls) — native export (Yardi/CBREI/RealPage/Entrata/
   OneSite), as-of a recent month. Carries **move-in** and ideally **lease start** dates
   (used to split new vs renewal — see *New vs renewal*); the parser auto-detects the column
   layout (two-row headers, and both **charge-code sub-row** "block" rolls and **flat/wide**
   rolls where each charge is its own column) and the unit-id format. OneSite-style rolls that
   list an upcoming lease (Applicant / Pending renewal / Future) as a **second row** per unit
   are collapsed to one row per physical unit, so unit counts and occupancy tie to the
   operator's own totals.
   You may pass **several rent rolls** (`--rr a.xlsx b.xlsx …`). The **newest** (by as-of date)
   is primary — it alone drives the dashboard, unit mix, occupancy and lease-type reads. Each
   roll gets its own **one-line tab** (dated), and **older rolls extend the Lease Trend's
   new-lease history** back in time — they capture signings that have since turned over and so
   are gone from the current roll.
3. **HelloData "Unit Details" CSV** *(optional)* — supplies clean bed/bath, the **T90
   executed** market-rent indicator per floor plan, and the monthly market-rent trend.
   It is joined to the rent roll **by unit number**, so HelloData's marketing floor-plan
   names need not match the rent roll's internal plan codes (the website names are also
   shown on the unit mix). Also feeds the model's `HD Dump`.
4. **Charge-code lookup** (.xlsx) *(optional, `--charge-codes`)* — for rent rolls that bill
   by **numeric operator code** rather than a name (e.g. CBREI/Yardi). A sheet with
   Account / Name / Type columns; the parser decodes each charge to its name so it
   categorizes correctly. Not needed when the rent roll already names its charges.

State which files you used and the operating period they cover.

## What it produces

One workbook, `<Property>__Underwriting_Intake.xlsx`, with up to **eight** tabs (the
HelloData tab appears only when a HelloData CSV is provided):

- **Dashboard** — snapshot (units, occupancy, SF, new/renewal counts), a **data-vintage**
  line (rent-roll as-of + latest financial statement), the **true-market-rent indicators**
  (HelloData asking/effective **T90** mix-weighted + new-lease contract T90), T12
  EGR/Opex/NOI, **the unit mix** (see below — its "Market Rent" column is the rent roll's
  asking figure and is *not* a market-rent signal), all auto-raised flags, and copy/paste
  instructions.
- **T12 Categorized** — the **entire** stitched statement set in **one clean section**:
  an editable **Code** column (amber dropdown), Category via VLOOKUP, raw line item, **one
  column per month across the union of all uploaded periods**, and a Total. Every distinct
  GL line appears **once** (merged across statements), with each month sourced from the
  freshest statement that owns it — so a SUMIFS by code reproduces the de-duplicated series
  with **no double-counting**, and the months flow continuously left-to-right rather than
  repeating per file. **This is the control surface** — change a code and everything
  downstream updates.
- **OS Summary** — the standardized chart of accounts, an exact clone of RedIQ's Overview
  template (A1:Q77), covering the **most recent 12 months** (the model paste target).
  Monthly cells are **live SUMIFS** over the `T12 Categorized` code column. The full
  multi-period detail lives line-by-line on `T12 Categorized` and in summary on `Lease
  Trend`, so there is no separate "operating history" tab to duplicate it.
- **Rent Roll (One-Line)** — one row per unit. Core columns A–M match the model's `RR
  Dump`; then, after **one blank spacer column**, a **per-unit charge-code block** (one
  column per scheduled charge — Rent, RUBS, amenity, parking, pet, valet trash, etc.,
  ordered by total $) showing exactly what each unit is billed. The charge block is
  reference detail, not a model paste target.
- **Lease Trend** — a **monthly** (left-to-right) grid over the **full HelloData + new-lease
  history** (back as far as the data goes; financials populate the overlap):
  **market rent** (HelloData executed asking/effective per unit, mix-weighted; HD concession
  %; executed counts; new-lease contract/unit), **occupancy & rent position** (economic
  occupancy, AGPR contract/unit, and **loss-to-lease backed into as HD market − AGPR**),
  **concessions** (T12 % of AGPR vs HelloData new-lease %, with the new-lease-vs-portfolio
  dilution explained), the **operating trend** (monthly EGR/Opex/NOI + T12/T6/T3
  annualizations — folded in from the old Trends tab), and the last-5 new leases / T90 by
  floor plan.
- **HelloData** *(if provided)* — CSV pass-through matching the model's `HD Dump`.
- **Reconciliation** — rent-roll ↔ T12 control tie-outs including the **T1 AGPR** tie
  (RR contract rent ↔ latest-month annualized AGPR) and the **amenity-rent verification**;
  a charge-code map flagging which charges are **in contract rent**, plus a **"Charge → T12
  Placement"** test that decides contract-rent membership *empirically* — matching each charge
  to the T12 line it actually lands on (both sides, since RUBS recoveries are booked as
  contra-expenses) and flagging disagreements with the categorization; a **"HelloData Market
  Rent: Fee Netting"** disclosure (gross→net HD asking **and** effective, the HD-vs-base gap,
  and the rent-roll candidate fees); **correlated
  cross-checks** (parking spaces billed ↔ T12 parking income; utility expense ↔ RUBS/billback
  income = **% recaptured**); an **NOI tie-out** that compares the standardized NOI to each
  source statement's own reported "Net Operating Income" subtotal (Δ should be ~0 — the RUBS
  gross-up nets to zero at NOI, so a gap means a line crossed the NOI boundary and is flagged);
  and flags. (The RR-vs-T12 market-rent gap is shown as informational only — both are
  seller-set asking and not a market-rent signal.)
- **Codes** — the standardized code legend and the dropdown's source list.

The **unit mix** (on the Dashboard, beside the snapshot) covers each floor plan: bed/bath,
occ/vac, avg SF, avg market & contract rent, **new vs renewal counts**, the **last 5
new-lease rents** and their average, and the HelloData executed market-rent reads —
**T90 and T365 asking/effective** plus **HD90 year-over-year asking/effective** (trailing
90 days vs the same window a year earlier). All portfolio totals are **mix-weighted** by
unit count. New-lease rents and the HelloData executed reads are the market-rent signals
underwriting actually trusts; renewals are excluded.

## Execution

### 1. Build the workbook

```bash
python3 <skill_dir>/scripts/build_intake.py \
  --t12 "<statement1.xlsx>" ["<statement2.xlsx>" ...] \
  --rr  "<RentRoll.xlsx>" ["<OlderRentRoll.xlsx>" ...] \
  [--hd "<hello.csv>"] [--charge-codes "<lease_charge_codes.xlsx>"] [--hd-fee-offset <$/mo>] [--name "Property Name"] \
  --out /home/claude/<Property>__Underwriting_Intake.xlsx
```

`--t12` accepts **any number** of statement files; pass them in any order (they are
sorted and stitched by date). `scripts/` holds three self-contained modules:
`account_map.py` (categorizer + chart of accounts), `intake_lib.py` (parsing, stitching,
reconciliation, unit-mix, lease-trend — no Excel), and `build_intake.py` (the writer).
Run from `scripts/` or add it to `sys.path`.

### 2. Recalculate (MANDATORY — the standardized statements are all formulas)

```bash
python3 /mnt/skills/public/xlsx/scripts/recalc.py /home/claude/<Property>__Underwriting_Intake.xlsx 120
```

Expect `status: success`, `total_errors: 0`. The workbook ships with ~1,000–2,000+
formulas (more with a long history); any `#REF!`/`#NAME?` means a tab reference broke —
fix before delivering.

### 3. Review the categorization

Open `T12 Categorized` and scan the **Code** column. The categorizer is a strong first
pass (validated at ~96–100% vs RedIQ on the test deal); confirm the operator-specific,
reliably contentious lines — see `references/account_mapping.md` (trash → `cont`,
workers' comp → `Pay`, employee-apartment concession → `PBo`, month-to-month →
`Rentinc`, **amenity rent → `Rentinc`** *only when it folds into Rental Income* — an amenity
*fee* booked as Other Income is not contract rent; the Reconciliation "Charge → T12 Placement"
test confirms which — insurance split from taxes, late fees → `OI`). Re-map by picking a new
code from the dropdown; Category and the OS Summary re-roll automatically.

### 4. Check the reconciliation

On `Reconciliation`, confirm **Contract Rent** ties between the rent roll and the T12's
latest month and review the **T1 AGPR** tie-out — contract rent / AGPR is the number that
"cannot be bullshitted," so that is the tie that matters. The **Gross Market Rent (asking)**
line is informational only: both the RR Market column and the T12 GPR are seller-set asking
rents, *not* a market-rent signal, so a gap there is expected and is **not** flagged — read
true market rent from the unit mix (new-lease + HelloData executed) and Lease Trend tabs.
Then read the charge-code map — especially any RUBS/valet-trash recovery that could net
against an expense or book as Other Income. Resolve the flags before underwriting.

The **"Charge → T12 Placement (empirical contract-rent test)"** section settles *what is in
contract rent* by evidence, not by charge name: each rent-roll charge is matched (by monthly
$, and name where the code is mnemonic) to the T12 line it actually lands on — searching
**both sides** of the T12, since RUBS recoveries (valet trash, water/sewer rebill, reimbursed
pest) are booked as **contra-expenses**, not revenue. A charge is called contract rent only
if it ties into a Rental Income line; if it ties to Other Income or a rebill it is not. The
match **overrides** the name categorization only on strong evidence (a tight $ tie, or a
decent tie corroborated by the line name) — a coincidental same-magnitude match cannot flip a
charge, and a charge folded into Rental Income with no distinct T12 line (e.g. amenity *rent*)
correctly stays contract. Disagreements between the evidence and the categorization are
flagged in red for review of the Code column. (This is why, at Aura, the $10 **amenity fee**
is *not* contract rent — it ties to the T12 "Amenities Income" line in Other Income — while
Canyon Ridge's **amenity rent** is, since it folds into Rental Income.) The **"Contract Rent"**
tie-out label adapts: it reads "incl. amenity rent" only when amenity rent actually rolls into
contract rent.

### 5. Hand off to the model

Use the paste targets in `references/model_paste_targets.md`. In short: `OS Summary`
A1:Q77 → `OS Summary Dump` (Paste **Values**, most recent 12 months); `T12 Categorized`
Code/Category/Line Item + the most recent 12 month columns → `T12 Dump`; `Rent Roll
(One-Line)` core columns A2:M → `RR Dump` (the per-unit charge-code block to the right of
the spacer is reference detail, not a paste target); `HelloData` A2:U → `HD Dump`. **Do
not** populate the model's underwriting tabs — these dumps are the only paste targets.

## Stitching (multiple statements → one continuous history)

- The union of every statement's months becomes the timeline. Each month is **owned** by
  the **freshest extract** that contains it (latest window-end, then most detailed), so
  later books supersede earlier ones and **each (code, month) is counted exactly once**.
- The `T12 Categorized` tab is built by **merging** every statement's GL lines into one
  ordered set (each distinct line once) and sourcing each month's value from that month's
  owner; non-owned months are simply blank. The live SUMIFS therefore reproduces the
  resolved series while the tab reads as one clean statement with months flowing across.
  Validated: on the test deal a 14-month stitch of two T12s rolls up to the standalone
  trailing-12 T12 to the dollar, and a SUMIFS by code over the merged tab equals the
  resolved union across all months and codes exactly.
- **Flags** are raised automatically: overlapping months that **disagree** between
  statements (restated financials) and months sourced from a **summary-level** statement
  (sub-line detail not itemized → reads as 0 in those codes).

## New vs renewal & true market rent

- **New lease** = lease start **on/equal to** move-in (the first lease on the unit);
  **renewal** = move-in is **older than** lease start (the resident moved in earlier and
  re-signed). Renewals are never market-tested, so only **new** leases inform market rent.
  When the rent roll carries **move-in but no lease start**, a move-in within **~12 months**
  of the as-of date is treated as a **new** lease (still on its original lease); an older
  move-in implies a **likely renewal** and is excluded from the new-lease reads. With only a
  lease-start (no move-in), or neither, the lease is left **unknown**.
- **HelloData ↔ rent roll join + bundled-fee netting.** HelloData is matched to the rent
  roll **by unit number**, so HD's marketing floor-plan names (e.g. *Barnsley*, *Phoenix*)
  resolve to the rent roll's internal plan codes; the website names are shown on the unit
  mix. HelloData scrapes the website **"Total Monthly"** price, which can fold **mandatory
  flat fees** (pest, amenity, valet trash, utility-billing admin, tech) into the asking
  number — inflating it vs base/contract rent. **Policy: gross + flag + confirm.** The skill
  does **not** auto-apply a fee — the true bundle lives on the **property website**, which this
  data (rent roll + T12 + HD executed) cannot see, and inferring it is unreliable (charge-code
  keying grabs per-unit "Varies" charges the website excludes and misses fees not itemized per
  unit; per-unit HD-vs-signed deltas are swamped by concession/term/listing-timing noise — see
  `references/hd_fee_detection.md`). So HD is shown **gross**, and the skill **flags** when HD
  gross sits materially above the new-lease base rent (likely website-bundled fees) and lists
  the rent-roll candidate fees as **evidence**. A fee is netted **only** when
  **`--hd-fee-offset <$/mo>`** supplies a confirmed amount. The **Reconciliation tab →
  "HelloData Market Rent: Fee Netting"** section discloses the full derivation (gross HD T90
  asking AND effective, base, gap, fee netted + source, net asking + net effective,
  candidates). The fee is removed from **both asking and effective** (and from T365 + the YoY
  reads) — anywhere HD market rent is shown — not just asking. **YoY note:** the HD90 YoY
  columns are computed on the **net-of-fee** rents (both the current and the year-ago figure),
  so the fee offset slightly shifts the growth rate (~0.1%). This is deliberate — it measures
  true rent growth with the constant fee stripped from both periods; it is *not* a $-for-$
  subtraction from the percentage.
  **When this gap is flagged, ASK the user** to check the property website's "Total Monthly"
  breakdown for a unit; if it bundles fees, re-run with `--hd-fee-offset` to net them. Do not
  guess the amount — the gap also contains genuine market premium. (Aura's site bundles
  pest $5 + amenity $10 = **$15**; the rent roll's `trtra`/`amfee` codes would have mis-detected
  $35 — which is exactly why the skill flags rather than auto-applies.)
- The **last 5 new-lease contract rents** per floor plan (in the Dashboard unit mix) and
  **HelloData executed** rents — **T90/T365 asking & effective**, plus **HD90 YoY** — are
  the preferred market-rent reads; cross-check them against each other. The T12 market-rent
  line is unreliable — do not use it.
- On a lease-up, total-revenue YoY reflects the occupancy ramp, not rent growth — read
  market-rent movement from the **Lease Trend** tab (per-unit, mix-weighted), not the
  operating series.

## Critical lessons

- **Contract rent = AGPR, never "in-place rent."** Contract rent = base Rent + **amenity
  rent** *where amenity rent folds into Rental Income* — i.e. it appears nowhere else on the
  T12, so including it ties the rent roll's contract rent tighter to the latest-month AGPR.
  But **do not assume every "amenity" line is contract rent**: an amenity *fee* booked as
  **Other Income** (e.g. Aura's $10/mo, which ties to the T12 "Amenities Income" line) is
  **not** contract rent. Let the Reconciliation **"Charge → T12 Placement"** test decide by
  where the charge actually lands on the T12, not the charge name.
- **Faithfully surface, don't normalize.** Raw statements carry GL noise (one-off vacancy
  or concession reclasses, lease-up ramps). Report it on Lease Trend / flags; let the analyst
  normalize in the model.
- **Detailed beats summary; recent owns overlaps.** A summary statement ties at subtotals
  but flattens `ltl`/`conc`/`RF`/`park` to zero for the months it owns. The freshest
  statement wins overlapping months.
- **The categorizer is a first pass, not gospel.** The editable code column + live rollup
  is the deliverable — match RedIQ where unambiguous to minimize edits; the user owns the
  judgment calls.
- **Reimbursement lines booked on the expense side flip sign.** Utility rebill / reimbursement
  lines that an operator parks in the expense section (a recovery booked as a negative
  contra-expense, or a rebilling service-fee cost booked positive) are reclassified to the RUBS
  revenue codes (RWS/RT/RF) and their sign is **negated** so they read as income — matching how
  RedIQ presents them, and keeping EGR/NOI correct. This **grosses up** RUBS: EGR and OpEx each
  rise by the recovery amount vs an operator statement that nets reimbursements inside expenses,
  but **NOI is unchanged** — which is why the standardized NOI still ties to the operator's own
  reported NOI line (see the Reconciliation tab's NOI tie-out).
- **Recalc is non-negotiable** — the standardized statements are formulas, so an
  un-recalced file shows zeros until opened.

## Output expectations

A clean-recalced `.xlsx` (0 errors) in the workspace with the model paste ranges ready
and the flags/reconciliation reviewed. Present it with `present_files` and a short
summary of the operating period stitched, what tied, the market-rent indicators, and any
open flags.
