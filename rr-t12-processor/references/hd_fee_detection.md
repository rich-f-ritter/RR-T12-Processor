# HelloData bundled-fee detection — design spec (future)

## Problem
HelloData scrapes a property website's headline **"Total Monthly"** price. That number
can fold **mandatory flat fees** (pest, amenity, valet trash, utility-billing admin, tech)
on top of base rent, so HD asking/effective is inflated versus the rent roll's
base/contract rent. To compare HD market reads against signed rents we must net out that
bundle — but **only the real bundle**, not noise.

### Why charge-code keying fails
The first implementation inferred the bundle from the rent roll: any non-contract charge
billed on ≥80% of units at a flat amount was summed and netted. This is unreliable:

- It **grabs charges the website excludes.** Aura's rent roll carries `trtra` (valet
  trash $25), but the website lists Trash as *"Varies / per unit"* — **not** in the
  $1,620 headline. We wrongly netted $25.
- It **misses fees not itemized per unit.** Aura's website bundles **Pest Control $5**,
  which is not a distinct rent-roll charge code, so we missed it.

Ground truth from the Aura unit 5104 cost estimate: `$1,620 = Rent $1,605 + Pest $5 +
Amenity $10`, i.e. the real bundle is **$15** — but the rent-roll heuristic produced $35.

Only the **website / HelloData** knows what is actually in "Total Monthly." So the current
skill shows HD **gross** by default and nets only an **explicit, disclosed** `--hd-fee-offset`
(see SKILL.md and the Reconciliation tab's "HelloData Asking: Fee Netting" section).

## The detector (unit-by-unit delta clustering)
Infer the bundle from the **distribution of per-unit deltas** between what HelloData
reported for a unit and what that unit actually signed for. The signal is a *constant*
offset, not a charge code.

### Inputs (already available)
- HelloData rows joined to the rent roll **by unit number** (the unit-mix join already
  does this). For each matched, occupied unit `i`:
  - `hd_i` = HelloData **Last Effective Rent** (preferred) or **Last Asking Rent**.
  - `signed_i` = rent-roll **base/contract rent** (contract rent, fees excluded).
  - `term_i`, lease start/`off-market` dates — for windowing/term adjustment.

### Method
1. **Window & match.** Keep units where HD's off-market date is reasonably close to the
   signed lease date (e.g. within ~120 days) so we compare the same lease event. Drop
   unmatched units.
2. **Per-unit delta.** `d_i = hd_i − signed_i`.
3. **Classify the deltas:**
   - **Ties** (`|d_i| ≲ $5`) → HD ≈ signed: no bundle on that unit (or fee already
     excluded). 
   - **Scattered** non-zero deltas → lease-term / concession / timing noise — ignore.
   - **A tight cluster at a common positive value** → that shared offset is the bundled
     fee. This is the mode of the delta distribution.
4. **Estimate the fee** as the **modal cluster center** (round to the nearest $1; e.g. a
   cluster at $14–$16 → **$15**), and report:
   - the estimate, the cluster's **share of matched units** (confidence), the
     within-cluster spread, and the count of ties vs scattered.
5. **Gate.** Only propose a netting offset when the cluster covers a meaningful share of
   matched units (e.g. ≥ 50–60%) and is tight (spread ≲ $5). Otherwise propose **$0**
   (gross) and surface the histogram for manual review.

### Robustness notes
- Use **effective** rent on both sides where possible so concessions don't masquerade as
  a fee. If only asking is available, expect more scatter.
- Term adjustment: a 6- vs 12-month lease premium will shift `d_i` for short-term units;
  either restrict to the dominant term or model the premium before clustering.
- A bimodal delta distribution can mean **two fee tiers** (e.g. with vs without a parking
  add-on) — report both modes rather than averaging them away.
- This is a **detector**, not an oracle: its output should populate the **disclosed**
  netting cell as a *default*, still overridable, and always shown on the Reconciliation
  tab with its confidence so the math can be double-checked.

### Output contract (proposed)
```
detect_hd_fee(rr, hd) -> {
    "fee": float,            # modal cluster center, rounded; 0.0 if no confident cluster
    "confidence": float,     # cluster share of matched units (0–1)
    "spread": float,         # within-cluster max−min
    "n_matched": int,
    "n_ties": int,
    "histogram": [(bucket, count), ...],   # for disclosure on the Reconciliation tab
}
```
Wire `fee` as the default for the netting offset (still overridable by `--hd-fee-offset`),
and render `confidence`/`histogram` in the "HelloData Asking: Fee Netting" section.
