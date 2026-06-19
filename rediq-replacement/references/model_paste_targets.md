# Model Paste Targets

Exact column maps from the intake workbook's tabs into the TMG acquisition model's
four "Dump" tabs. The model's downstream columns (anything past the paste range) are
model-computed — **do not** overwrite them. Paste **values** for the OS Summary;
plain paste is fine for the others.

The model's dump tabs (confirmed in the blank `TMG_Acquisition_Model`):
`RR Dump` A1:T1720 · `HD Dump` A1:U294 · `OS Summary Dump` A1:Q77 · `T12 Dump` A2:O222.

---

## 1. OS Summary  →  `OS Summary Dump`  (A1:Q77)

Copy the **whole** `OS Summary` sheet `A1:Q77` and Paste-Special → **Values** onto
`OS Summary Dump!A1`. The layout is a 1:1 clone of RedIQ's Overview template, so the
model's references line up exactly.

| Col | Contents |
|----|----------|
| A | Standardized code (blank on subtotal/section rows) |
| B | Category label / section header |
| C, D, E | Annual columns — left blank (T12-only intake) |
| F … Q | 12 monthly values (oldest month in F) |

Row anchors: header row 6 · revenue 7–19 · **EGR row 20** · expenses 23–41 · **Opex
row 42** · **NOI row 44** · non-op revenue 47 / subtotal 48 · non-op expenses 51–72 /
subtotal 73 · **Net Income row 75**.

Because the cells are live SUMIFS in the intake file, paste as **values** so the model
doesn't inherit a cross-workbook link.

---

## 2. T12 Categorized  →  `T12 Dump`  (A2:O222)

Copy `T12 Categorized!A:O` (i.e. through the last month column — **not** the `Total`
column, which falls in P) and paste onto `T12 Dump!A2`.

| Col | Contents |
|----|----------|
| A | Code (editable; blank on subtotal/section rows so SUMIFS skips them) |
| B | Category (VLOOKUP of the code) |
| C | Line Item — the raw GL account name; also holds native section/subtotal labels |
| D … O | 12 monthly values (oldest in D). **No total column** — the model recomputes |

This preserves the T12's native section/subtotal structure exactly as RedIQ's detailed
mapping sheet does (subtotal and header rows ride along with a blank code).

> The paste assumes a standard 12-month T12 (months D–O). If a T12 has fewer than 12
> months, the month block is shorter and the analyst should align columns manually.

---

## 3. Rent Roll (One-Line)  →  `RR Dump`  (A2:M…)

Copy `Rent Roll (One-Line)!A2:M{last unit row}` (exclude the TOTAL/AVG row) onto
`RR Dump!A2`.

| Col | Header | Source |
|----|--------|--------|
| A | Unit No. | rent roll unit id |
| B | Floor Plan | "Unit Type:" header above each block |
| C | Net sf | rent roll SQFT |
| D | Bed | HelloData → "NxM" in plan → letter convention (A=1, B=2, C=3) |
| E | Bath | same source as Bed |
| F | Lease Type | parsed (Market / Vacant / …) |
| G | Occupancy Status | Occupied / Vacant / Non-Revenue |
| H | Market Rent | rent roll Market column |
| I | Contractual Rent | **base Rent + Amenity** scheduled (never "in-place rent") |
| J | Net Effective Rent | contract net of recurring concessions |
| K | Lease Start Date | |
| L | Lease Expiration | |
| M | Move In Date | |

Model columns N–T (assumed dates, new/renewal split, expiry timing) are computed from
these — leave them alone.

---

## 4. HelloData  →  `HD Dump`  (A2:U…)

Only present if a HelloData CSV was supplied. Copy `HelloData!A2:U{last}` onto
`HD Dump!A2`.

| Col | Contents |
|----|----------|
| A … T | HelloData CSV passed through in the model's column order |
| U | Floorplan Mapped — the plan id normalized (trailing "NxM" bed/bath token stripped) |

---

## Always
- **"Contract rent," never "in-place rent."**
- Do **not** populate the model's underwriting tabs from here — these four dumps are
  the only paste targets. Underwriting happens in the model after the paste.

---

## Multi-period stitching & the trailing-12 paste

When several statements are uploaded they are stitched into one continuous monthly series
over the union of their months. For paste targets:

- **`OS Summary` and the `T12 Dump` are always the most recent 12 months.** The model's
  `OS Summary Dump` (A1:Q77) and `T12 Dump` expect a trailing-12 window, so the OS Summary
  tab shows exactly the last 12 months of the stitched series. Copy its A1:Q77 / the 12
  most-recent month columns as usual.
- **The full multi-period detail is on `T12 Categorized`** — every GL line once, with all
  stitched months flowing across — and in summary on the `Trends` tab. There is no separate
  "operating history" tab; the line-level history and the trailing-12 paste come from the
  same `T12 Categorized` surface.

Overlapping months are owned by the freshest statement, so the stitched series never
double-counts. If an overlap-disagreement or summary-granularity flag is raised on the
Dashboard, resolve it before pasting.

## Rent roll charge-code block (reference, not a paste target)

The `Rent Roll (One-Line)` tab adds, after the core columns A–M and **one blank spacer
column**, a per-unit charge-code block — one column per scheduled charge (Rent, RUBS,
amenity, parking, pet, valet trash, M2M, concessions, etc.), ordered by total dollars,
showing exactly what each unit is billed by code. Only the core columns A2:M feed `RR
Dump`; the charge block is underwriting reference (e.g. confirming which charges roll into
contract rent and which are Other Income), not a model paste target.
