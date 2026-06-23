# HelloData bundled fees — why the skill flags, and does NOT auto-detect

## Problem
HelloData scrapes a property website's headline **"Total Monthly"** price, which can fold
**mandatory flat fees** (pest, amenity, valet trash, utility-billing admin, tech) on top of
base rent. So HD asking/effective is inflated versus the rent roll's base/contract rent, and
to compare HD market reads against signed rents the bundle must be netted out — but **only the
real bundle**, not noise or genuine market premium.

## Decision: gross + flag + confirm (no auto-apply)
The skill shows HD **gross**, **flags** when HD gross sits materially above the new-lease base
rent, lists the rent-roll candidate fees as **evidence**, and nets a fee **only** when
`--hd-fee-offset <$/mo>` supplies a confirmed amount. The agent should **ask the user** to
confirm the website's "Total Monthly" breakdown when the gap is flagged.

This is deliberate. Across thousands of deals, auto-applying an inferred fee produces
*confidently wrong* adjustments, and the data the skill has does not contain a clean fee
signal. The two inference approaches both fail:

### 1. Charge-code keying fails
Summing non-contract rent-roll charges billed on most units:
- **Grabs charges the website excludes.** Aura's rent roll carries `trtra` (valet trash $25),
  but the website lists Trash as *"Varies / per unit"* — not in the $1,620 headline.
- **Misses fees not itemized per unit.** Aura's website bundles **Pest Control $5**, which is
  not a distinct rent-roll charge code.

Ground truth (Aura unit 5104): `$1,620 = Rent $1,605 + Pest $5 + Amenity $10`, so the real
bundle is **$15** — but the charge-code heuristic produces **$35**.

### 2. Per-unit HD-vs-signed delta clustering fails (empirically)
The intuitive detector — cluster `HD rent − signed rent` per unit and read the constant offset
as the fee — was tested on Aura (462 HD listings, 428 joined to the rent roll by unit):
- The **dominant cluster is delta = $0** (87 listings tie exactly), **not** a clean +$15 cluster.
- The rest are **scattered negatives** (−$250, −$239, −$185, …), median **−$143**.
- **No visible +$15 cluster.**

The signal is buried, for reasons that recur on most deals:
1. HD lists a unit **multiple times over its history**; comparing an old listing to the current
   signed rent without matching the listing to its lease (by date) adds noise.
2. The bundle lives in HD's **asking / "Total Monthly,"** but **effective** rent is net of
   concessions — a different number.
3. **Concessions and lease-term** differences swamp a ~$15 fee.

Recovering it robustly would require listing-to-lease timing matches, rent-field selection, and
term/concession adjustment before clustering — not simple, and **still** unable to recover the
truth, which lives on the **property website** (not in the rent roll, T12, or HD executed data).

## What the skill CAN do reliably (and does)
- Compute, mix-weighted, the **gross HD T90 asking**, the **new-lease base rent (T90)**, and the
  **gap** between them.
- List the **candidate flat fees** on the rent roll (as evidence, never auto-applied).
- **Flag** when the gap is material (HD gross > base by more than ~$10 or ~0.5% of base),
  prompting the user to confirm the website bundle and pass `--hd-fee-offset`.
- Disclose all of it on the Reconciliation tab's **"HelloData Market Rent: Fee Netting"** section.

The gap is a *flag*, not a fee: it also contains genuine market premium, so it must not be
assumed to be all fees.
