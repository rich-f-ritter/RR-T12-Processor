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
- **Workers' comp → `Pay`** (RedIQ keeps it in base payroll, not burden).
- **Employee apartment concession → `PBo`** (a payroll bonus, not a rent concession).
- **Month-to-month / STL fee → `Rentinc`** (pulled up into rental income).
- **Insurance** always splits out of a combined "Taxes & Insurance" line → `ins`.
- **Late Fees**: if the line sits in the *Other Income* section it is revenue → `OI`.
  RedIQ sometimes buckets it to `GA`; the editable column lets you match house style.

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

## New vs renewal (lease-date rule)

Per-unit: **new lease** = lease start date **on or before** move-in date; **renewal** =
lease start **after** move-in. Only new leases are market-tested, so the unit mix's
"last 5 new-lease rents" and the lease-trend new-lease series use new leases only.
