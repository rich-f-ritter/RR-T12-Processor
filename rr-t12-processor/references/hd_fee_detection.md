# HelloData & website-bundled fees — why the skill flags, confirms, and does NOT auto-detect

## What HelloData actually reports
HelloData reflects the **price the property website advertises**. **Usually that is the base
asking rent.** But some operators advertise **all-in pricing** — base rent **plus** mandatory
flat fees (pest, amenity, valet trash, utility-billing admin, tech) — as the headline
"Total Monthly," and HelloData then captures that inflated number. Which one you get is
**operator-dependent**, so the skill must not assume either way.

When HD carries an all-in number, HD asking/effective is inflated versus the rent roll's
base/contract rent, and the bundle should be netted out for a like-for-like comparison — but
**only the real bundle**, confirmed per deal, never an inferred guess.

### The phenomenon is real — but a bundling website does NOT guarantee HD bundled
**Worked example — Aura Beacon Island (a false alarm):** the property website itemizes a
$15 bundle on its cost estimate — unit 5104 shows `$1,620 Total Monthly = Rent $1,605 +
Pest $5 + Amenity $10`. It *looks* like HD should be $15 high. But HelloData's asking for
unit 5104 came through at **$1,605 — the base rent**, not the $1,620 total. So **HD was not
carrying the fee, and netting $15 would have understated it.** Verified by tying HD's
unit-level asking to the website's base rent, and corroborated by HD's current on-market
asking matching the website's advertised "starting" rents (1BR from $1,535, 2BR from $2,085).

The lesson: **confirm at the unit level before netting.** A website that bundles does not mean
HD picked the bundle up. On other properties the website's advertised number *is* the all-in
figure and HD *does* carry it — there the netting is correct. Decide per deal.

**Worked example — The Preserve, Grapevine TX (a confirmed bundle → net it):** a **Greystar**
property, and Greystar advertises a **"Total Monthly Leasing Price" = Base Rent + fixed,
mandatory monthly fees** (it's their standard pricing model, stated right on the listing). HD
scrapes that all-in number. The rent roll carries exactly two flat, near-universal mandatory
fees — **Trash Rebill – Door to Door $25** (386 of 399 units) and **Package Valet $15** (≈all
units) = **$40**. HD T90 asking ($1,905, mix-wtd) sat **$75** over the new-lease base ($1,830);
netting the confirmed **$40** bundle left a ~$35 residual, i.e. ordinary asking-over-signed
premium (~2%). Built with `--hd-fee-offset 40`. Contrast with Aura: same flag fired on both,
but the Aura website didn't bundle and Greystar's does — which is exactly why the call is made
per deal, not by rule. The **"Total Monthly Leasing Price"** phrasing (Greystar and some other
operators) is the most reliable tell that HD is carrying an all-in number.

## Why the skill does not auto-detect the fee
The truth lives on the **property website**, and the data the skill has (rent roll + T12 + HD
executed) does not contain a clean fee signal. Every inference route is unreliable:

### 1. Charge-code keying fails
Summing non-contract rent-roll charges billed on most units grabs charges the website excludes
(Aura's `trtra` valet trash $25 is "Varies/per-unit," not in the headline) and misses fees not
itemized per unit (Aura's pest $5). It would have produced a $35 "bundle" for Aura — which is
doubly wrong, since HD wasn't bundling anything there to begin with.

### 2. Per-unit HD-vs-signed delta clustering fails (empirically)
Clustering `HD rent − signed rent` per unit was tested on Aura (462 listings, 428 joined): the
dominant delta cluster is **$0**, the rest scattered negatives (median −$143), no clean fee
signal. HD lists a unit many times over its history, "effective" is net of concessions, and
lease-term/concession noise swamps any flat fee.

### 3. Scraping the website automatically is blocked
Tested live on Aura (June 2026): the property's RentCafe site, the RentCafe backend,
apartments.com, Lighthouse, and HAR **all return HTTP 403** (server-side bot protection; the
agent proxy was healthy). And the cost-estimate breakdown is a dynamic JS widget that would not
render through a markdown fetch even on a 200. So an automated, baked-in scraper is not viable
across deals.

## What the skill does instead: gross + flag + confirm
- Show HD **gross**; never auto-apply a fee.
- Compute, mix-weighted, the **gross HD T90 asking/effective**, the **new-lease base rent
  (T90)**, and the **gap** between them.
- **Flag** when HD asking sits materially above base (HD gross > base by ~$10 or ~0.5%).
  The flag is explicit that the gap is **usually ordinary market premium** (current asking
  above recently-signed rents) and only *sometimes* bundled fees.
- List the rent-roll **candidate flat fees** as evidence (never auto-applied).
- Net a fee **only** via `--hd-fee-offset <$/mo>`, after the user/agent **confirms** by
  comparing HD's asking for a *currently-listed* unit to that unit's base vs "Total Monthly"
  on the property website.
- Disclose all of it on the Reconciliation tab's **"HelloData Market Rent: Fee Netting"**
  section (gross→net for both asking and effective).

## Cleaner future input
If HelloData's export or API exposes a **fees/concessions/amenities breakdown** (beyond the
bundled `Last Asking/Effective Rent` in the Unit Details CSV), that is a deterministic source
and beats both scraping and inference — worth asking the HD rep.
