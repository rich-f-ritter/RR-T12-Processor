# Account Mapping Reference

How `account_map.py` assigns each raw T12 line and each rent-roll charge code to a
standardized RedIQ-style code. This is a **first pass** — the workbook's `T12
Categorized` tab exposes every code in an editable dropdown, and the `OS Summary`
re-rolls live via SUMIFS, so a human can override any debatable line in one place.

The code list and ordering are an **exact** match to RedIQ's Overview template
(13 revenue codes, 19 operating-expense codes, 1 non-operating revenue, 22
non-operating expense). Do not reorder or rename codes — the `OS Summary` tab is
laid out to paste 1:1 into the model's `OS Summary Dump` (A1:Q77).

## The standardized chart of accounts

**Effective Gross Revenue** (rows 7–19): `Rentinc` Rental Income · `OI` Other
Income · `RWS` Rubs Water/Sewage · `RT` Rubs Trash · `RF` Rubs Fees · `park`
Garage/Parking · `CI` Corporate Income · `cable` Cable Income · `ltl` (Loss to
Lease)/Gain to Lease · `vac` Vacancy · `cl` Collection Loss/Bad Debt · `nr`
Non-Revenue Units · `conc` Concessions.

**Operating Expenses** (rows 23–41): `Pay` Payroll · `PB` Payroll Burden · `PBo`
Payroll Bonuses · `adv` Marketing/Advertising · `GA` General & Administrative ·
`turn` Turnover · `inter/exte` R&M Interior/Exterior · `cont` Contract · `Safe`
Life Safety/Elevators · `HAF` HOA Fees · `UT` Utilities Trash · `UWS` Utilities
Water Sewer · `UC` Utilities Common · `UF` Utilities Fees · `mgt` Property
Management Fee · `ins` Insurance · `ret` Real Estate Taxes · `NI` Not Included ·
`site` R&M Site.

**Non-operating** (rarely present in a T12): `dispro` (revenue); and expenses
`rd capx aex ece tird tirw tiae intex prin draw repay othdf ord icd scd entex
pfex purch acqex disex depex onoe`.

`ltl`, `vac`, `conc`, `cl`, `nr` carry **negative** values (they reduce revenue).
Signs come straight from the GL — the rollup is a plain `SUM`, never re-signed.

## How a T12 line gets its code

`categorize_t12_line(name, side, section_hint, acct_number)` runs in this order:

1. **Cross-section overrides** (operators reliably misfile these):
   - *forced-placed / tenant-liability / legal-liability insurance* → **GA** (it is a
     recovery cost, not property insurance — per the underwriting methodology).
   - *uniforms* → **GA** (a supply, even when filed under payroll).

2. **Section-aware family** — `section_hint` is the operator's own native
   subsection label, which the parser tracks. It is normalized loosely so it works
   across Yardi / RealPage / Entrata. Section → family:
   payroll/personnel/compensation→payroll · advertising/marketing→adv ·
   G&A/office→GA · utilities→utilities · maintenance/R&M/turnover→maintenance ·
   management fee→mgt · tax|insurance→taxins · contract→cont · other income→otherinc ·
   rubs/reimbursement/recovery→rubs · concession→conc · rental
   adjustment/vacancy/loss→rentadj · gross potential/rental income→rent · write-off→cl.

3. **Within-family splitters** refine by keyword once the family is known. The ones
   that matter most:
   - **payroll**: bonus/commission/incentive **and** employee-apartment-concession /
     housing-allowance → **PBo**; fica/payroll-tax/health/401/benefit/insurance/
     unemployment/burden → **PB**; everything else (salaries, wages, manager,
     leasing, maintenance, **workers' comp**) → **Pay**.
   - **utilities**: **trash/garbage/refuse/valet → `cont`** (RedIQ files trash hauling
     under Contract, not a utility); water/sewer → `UWS`; billing/transfer/audit/fee →
     `UF`; electric/gas/common/cable/energy → `UC`.
   - **maintenance**: turn/make-ready/carpet/paint/vinyl/resurface/**key/lock** →
     `turn`; landscape/pest/snow/elevator/fire/sprinkler/courtesy patrol/security/
     valet/monitor → `cont`; all other physical R&M → `inter/exte`.
   - **taxins**: insurance → `ins`; tax → `ret`.
   - **otherinc**: **month-to-month / short-term-lease fee → `Rentinc`** (it's a rent
     premium); parking/storage → `park`; water-reimb → `RWS`; trash-reimb/valet-trash →
     `RT`; utility-reimb/rubs → `RF`; cable/bulk → `cable`; corporate-housing → `CI`;
     else `OI`.
   - **rentadj**: non-revenue/down/model/employee unit → `nr`; vacancy → `vac`;
     concession/free rent → `conc`; loss-to-lease → `ltl`; bad debt/collection/skip →
     `cl`; else `vac`.

4. **Pure-keyword fallback** (`REVENUE_RULES` / `EXPENSE_RULES`) only when no section
   is recognized. Last rule is catch-all `OI` (revenue) or `GA` (expense).

### RedIQ quirks worth knowing
- **Trash hauling → `cont`** (Contract), not a utility. Valet-trash *recovery* on the
  income side → `RT`.
- **"Rebill" / "Reimbursement" — recovery (revenue) vs billing-program cost (expense).**
  A utility *recovery* billed to residents ("Water/Sewer Rebill", "Electric Rebill",
  "Utility Reimbursement", "Reimbursed Pest") is RUBS **revenue** → `RWS`/`RT`/`RF` (even when
  the operator parks it on the expense side as a contra — it's sign-flipped up to income). But a
  **"rebill service" or "billing fee"** line ("Utility Rebill Services", "Utility Rebill Service
  Fees") is the **cost the property pays a third party to run the RUBS program** — a genuine
  utility **expense → `UF`** (Utilities Fees), not income. The tell is the word *service*/*fee*.
- **Workers' comp → `Pay`** (RedIQ keeps it in base payroll, not burden).
- **Employee apartment concession → `PBo`** (a payroll bonus, not a rent concession).
- **Month-to-month / STL fee → `Rentinc`** (pulled up into rental income).
- **Insurance** always splits out of a combined "Taxes & Insurance" line → `ins`.
- **Late Fees**: if the line sits in the *Other Income* section it is revenue → `OI`.
  RedIQ sometimes buckets it to `GA`; the editable column lets you match house style.
- **A NEGATIVE value in an EXPENSE line is often INCOME.** Operators frequently book a
  resident recovery / pass-through / rebill as a *contra* inside the expense section
  (e.g. "Resident Utility Passthrough", a tenant rebill of a service contract). A
  consistently-negative expense line usually belongs in revenue (a RUBS recovery → `RF`/
  `RWS`/`RT`, or `OI`). The build flags materially- and consistently-negative expense lines
  with a **`CONFIRM categorization`** note so they get reclassified. (Alta Berry Creek:
  "Boiler Rebill" — the tenant rebill of the boiler maintenance contract — is income `RF`,
  not a contract expense; RedIQ had it as `cont`, a model error.)

## RedIQ cross-check learnings (The Preserve, Jun 2026)
Validated a full build against the operator's own **RedIQ** Operating Statement (same
standardized A1:Q77 code template the OS Summary uses). **NOI tied to the dollar
($5,695,294)**; 140 GL lines shared, ~81% coded identically. The divergences were all
NOI-neutral reclassifications and surfaced these lessons:
- **Marketing GL block `54xxx` → `adv` (FIXED).** Hierarchical T12s put a parent GL line
  in the section column ("54005-000 - Ad Performance Fees", "54025-000 - Property
  Website"); `_family_from_section` now strips the account prefix and recognizes
  marketing parent-sections (ad performance, property website, online presence, search
  engine, internet listing/ILS, locator). Fixed 8 lines / ~$64k that had scattered into
  `PB`/`GA` (incl. "Search Engine Marketing" → `PB`). `adv` now matches RedIQ exactly.
- **RedIQ's `UF` is a utilities catch-all (OPEN — house-style choice).** RedIQ folds
  common-area & vacant-unit electric, internet access, and cable-TV contract into `UF`
  (Utilities Fees). This skill splits them into `UC` (common), `GA` (internet), and
  `cont` (cable). Both tie at NOI; the editable code column lets you match RedIQ if
  preferred. (~$66k of reclass across `UF`/`UC`/`GA`/`cont` on this deal.)
- **Utility-rebill *reimbursement* → `RF`, not `UF`.** The rebill *service fee* (cost) is
  `UF`, but a line titled a "…Service Fee **Reimbursement**" is the resident **recovery**
  → RedIQ codes it `RF` (RUBS revenue, grossed up). This skill currently nets it into
  `UF`; refine if matching RedIQ. NOI-neutral.
- **"Centralization Fees" / "Contract Staffing" → `PB`** in RedIQ (payroll burden), where
  this skill reads them as base `Pay`. Minor; arguable.
- **Month-to-Month premiums:** RedIQ booked these to `OI` here, not `Rentinc` — counter to
  the quirk note above. Operator-dependent; low-dollar.

## RedIQ cross-check learnings (Alta Berry Creek, Jun 2026)
Validated against the operator's **corrected** RedIQ export (after the Boiler Rebill fix).
**EGR / OpEx / NOI tied to the dollar ($4,289,688 / $3,742,062 / $547,626).** All 13 revenue
codes matched **exactly**, and the divergences were 5 NOI-neutral OpEx reclasses totaling well
under 1% — no errors in either direction:
- **Boiler Rebill → `RF` confirmed on both sides (`RF` = $277,625 exact).** The one real model
  error caught on this deal — the tenant rebill of the boiler maintenance contract is RUBS
  **revenue** (`RF`), not a `cont` expense. The corrected RedIQ now agrees. (The matching boiler
  *cost* — "Boiler Contract" $49,500 — stays in `cont`.) This is the canonical "a negative in
  the expense section is often income" case; see the quirk note above.
- **Paint *supplies* vs paint *contractor* (OPEN — $10,354, largest divergence).** RedIQ keeps
  **Paint Contractor** (labor) in `turn` (make-ready) but **Painting Supplies** in `inter/exte`
  (general R&M). This skill's maintenance splitter routes anything containing "paint" → `turn`,
  so it pulls Painting Supplies into turnover. Defensible to refine: `paint` + supply/material
  context → `inter/exte`; `paint` + contractor/labor/make-ready → `turn`. NOI-neutral.
- **Internet Access → `GA` vs `UC` (house-style, $4,320).** Same internet split as The Preserve,
  except RedIQ folded internet into `UF` there and `UC` here — RedIQ is internally inconsistent
  on this, so not worth chasing. NOI-neutral.
- **Minor (<$1k each, NOI-neutral):** Fuel & Propane (mine `inter/exte` / RedIQ `UC`);
  Pool Supplies and Preventative Maintenance (mine `inter/exte` / RedIQ `cont`); Safety & Fire
  Supplies (mine `cont` / RedIQ `inter/exte`); Other Recreational Amenities (mine `GA` / RedIQ
  `inter/exte`). All judgment calls between R&M / contract / utilities buckets.

## Detailed vs. summary T12
A **detailed** T12 has separate GL lines for concessions, loss-to-lease, RUBS fees,
parking, payroll burden/bonuses, insurance, etc. — the categorizer separates each
into its own code. A **summary** T12 collapses these into parent lines (e.g. "Rental
Income" already net of LtL/concessions; "Taxes & Insurance" as one line). On a summary
T12 the EGR / Opex / NOI subtotals still tie exactly, but codes like `ltl`, `conc`,
`RF`, `park` will read 0 because the detail isn't in the file. **Prefer the most
detailed T12 export available.**

## Rent-roll charge codes
`categorize_charge(name)` → `(code, is_contract_rent, is_recurring)`.
**Contract rent = base Rent + Amenity/premium rent** (the only two `is_contract_rent`
charges). Everything else — parking, RUBS, pet, technology, late fees, application,
admin, insurance/renters-liability — is recurring or one-time **Other Income** and is
flagged accordingly in the `Reconciliation` tab's "In Contract Rent?" column.

## Validation against RedIQ (Canyon Ridge)
- Per-line codes on the **detailed** T12 vs RedIQ's own line-level mapping: **96/97**
  (the single difference is the Late-Fees call above).
- **EGR / Operating Expenses / NOI** reconcile to RedIQ **exactly** (Δ = $0) on the
  matching trailing-12 window.

---

## Amenity rent → Rental Income (contract rent / AGPR)

"Amenity Rent" (and unit premiums: view/floor/upgrade premiums) maps to **`Rentinc`** and
is treated as **contract rent**, not Other Income. Rationale, verified on the Reconciliation
tab:

- Amenity rent is a recurring, lease-level rent premium tied to the specific unit — it is
  part of what the existing lease pays (the definition of contract rent = AGPR).
- It appears **nowhere else on the T12** (no separate amenity line), so excluding it from
  contract rent would lose it entirely.
- Including it ties the rent roll's contract rent **tighter to the latest-month AGPR**
  (Rentinc + loss-to-lease, annualized). On the test deal, including amenity narrowed the
  gap to T1 AGPR materially vs excluding it.

The Reconciliation tab prints this check (the "T1 AGPR, annualized" tie-out plus an
amenity-verification flag) on every run, so the treatment is re-validated per deal rather
than assumed. Pet rent, parking, and technology/valet packages are **not** contract rent
(they are Other Income / their own codes).

**Amenity *rent* vs amenity *fee* — they are not the same.** The above applies to an amenity
**rent premium** that folds into Rental Income (no separate T12 line). A flat monthly **amenity
fee** is different: operators book it as its **own Other Income line** (e.g. Aura's "Amenities
Income," ~$10/unit/mo), so it is **not** contract rent. The two are easy to confuse by name.
The Reconciliation **"Charge → T12 Placement"** test resolves it empirically — it matches the
charge to the T12 line it actually lands on, so an amenity item that ties to an Other Income
line is correctly excluded from contract rent, while one folded into Rental Income stays in.
Decide by where it lands on the T12, not by the word "amenity."

## New vs renewal (lease-date rule)

Per-unit: **new lease** = lease start date **on or before** move-in date; **renewal** =
lease start **after** move-in. Only new leases are market-tested, so the unit mix's
"last 5 new-lease rents" and the lease-trend new-lease series use new leases only.
