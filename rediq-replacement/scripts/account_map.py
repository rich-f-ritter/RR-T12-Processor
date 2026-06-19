"""
account_map.py
==============
The categorization "brain" for the RedIQ-replacement intake engine.

Jobs:
  1. Define the canonical standardized chart of accounts (the OS Summary rows, in
     order, with section + sign). Mirrors RedIQ's "Overview" sheet and the TMG
     model's 'OS Summary Dump' tab exactly.
  2. Auto-categorize each raw T12 line item to a standardized code, and each raw
     rent-roll charge code to a standardized code.

CATEGORIZATION STRATEGY (why it's built this way):
  RedIQ -- and good analysts -- largely RESPECT the operator's own statement
  structure. A line sitting in the "General & Administrative" section is almost
  always G&A, even if its name ("Courtesy Patrol") might keyword-match Contract.
  So the categorizer is SECTION-AWARE: the native T12 subsection (passed by the
  parser as `section_hint`) picks the CODE FAMILY, then keywords REFINE within the
  family. When no section is known (other PM systems, odd layouts) it falls back to
  pure keyword rules. A short list of cross-section OVERRIDES handles the handful of
  lines operators reliably misfile (e.g. forced-placed insurance -> G&A).

  Auto-categorization is only a FIRST PASS. The intake workbook's Code column is
  editable and OS Summary re-rolls-up via SUMIFS, so the goal is "mostly right and
  never crashes," not perfection. Unknowns fall back to OI (revenue) / GA (expense)
  and are flagged for human review.
"""

import re

# ---------------------------------------------------------------------------
# 1. CANONICAL STANDARDIZED CHART OF ACCOUNTS (order/labels = OS Summary template)
# ---------------------------------------------------------------------------
REVENUE_CODES = [
    ("Rentinc", "Rental Income"),
    ("OI",      "Other Income"),
    ("RWS",     "Rubs Water / Sewage"),
    ("RT",      "Rubs Trash"),
    ("RF",      "Rubs Fees"),
    ("park",    "Garage / Parking"),
    ("CI",      "Corporate Income"),
    ("cable",   "Cable Income"),
    ("ltl",     "(Loss to Lease) / Gain to Lease"),
    ("vac",     "Vacancy"),
    ("cl",      "Collection Loss / Bad Debt"),
    ("nr",      "Non-Revenue Units"),
    ("conc",    "Concessions"),
]
EXPENSE_CODES = [
    ("Pay",         "Payroll"),
    ("PB",          "Payroll Burden"),
    ("PBo",         "Payroll Bonuses"),
    ("adv",         "Marketing / Advertising"),
    ("GA",          "General & Administrative"),
    ("turn",        "Turnover"),
    ("inter/exte",  "R&M Interior / Exterior"),
    ("cont",        "Contract"),
    ("Safe",        "Life Safety / Elevators"),
    ("HAF",         "Homeowners Association Fees"),
    ("UT",          "Utilities Trash"),
    ("UWS",         "Utilities Water Sewer"),
    ("UC",          "Utilities Common"),
    ("UF",          "Utilities Fees"),
    ("mgt",         "Property Management Fee"),
    ("ins",         "Insurance"),
    ("ret",         "Real Estate Taxes"),
    ("NI",          "Not Included"),
    ("site",        "R&M Site"),
]
NONOP_REV_CODES = [("dispro", "Disposition Proceeds")]
NONOP_EXP_CODES = [
    ("rd","Replacement Reserve Deposits"),("capx","Capital Improvements"),
    ("aex","Appliance Expenditures"),("ece","Extraordinary Capital Expenditure"),
    ("tird","TI&LC Reserve Deposits"),("tirw","TI&LC Reserve Withdrawls"),
    ("tiae","TI&LC Actual Expenditures"),("intex","Interest"),("prin","Principal"),
    ("draw","Loan Draws"),("repay","Loan Repayments"),("othdf","Other Debt Fees"),
    ("ord","Other Reserves Funded / Used / Released"),
    ("icd","Investor Contributions and Distributions"),
    ("scd","Sponsor Contributions and Distributions"),("entex","Entity Expenses"),
    ("pfex","Partnership Fees and Expenses"),("purch","Purchase Price"),
    ("acqex","Acquisition Expenses"),("disex","Disposition Expenses"),
    ("depex","Depreciation"),("onoe","Other Non-Operating Expense"),
]
CODE_TO_CATEGORY = {c: l for c, l in
                    REVENUE_CODES + EXPENSE_CODES + NONOP_REV_CODES + NONOP_EXP_CODES}
CODE_SECTION = {}
for c, l in REVENUE_CODES:   CODE_SECTION[c] = "rev"
for c, l in EXPENSE_CODES:   CODE_SECTION[c] = "opex"
for c, l in NONOP_REV_CODES: CODE_SECTION[c] = "nonop-rev"
for c, l in NONOP_EXP_CODES: CODE_SECTION[c] = "nonop-exp"
NEGATIVE_REVENUE_CODES = {"ltl", "vac", "conc", "cl", "nr"}
ALL_CODES = list(CODE_TO_CATEGORY.keys())


# ---------------------------------------------------------------------------
# 2. NATIVE SECTION  ->  CODE FAMILY
#    `section_hint` is the operator's own subsection label (the parser tracks it).
#    We normalize loosely so it works across Yardi/RealPage/Entrata naming.
# ---------------------------------------------------------------------------
def _family_from_section(section_hint):
    s = (section_hint or "").strip().lower()
    if not s:
        return None
    if re.search(r"payroll|personnel|salaries|compensation|labor", s):           return "payroll"
    if re.search(r"advertis|marketing", s):                                       return "marketing"
    if re.search(r"general.*admin|administ|g\s*&\s*a|g/a|office", s):             return "admin"
    if re.search(r"utilit", s):                                                   return "utilities"
    if re.search(r"maintenance|repairs|r\s*&\s*m|make.?ready|turnover", s):       return "maintenance"
    if re.search(r"management fee|mgmt fee|management$", s):                       return "mgmt"
    if re.search(r"tax|insurance", s):                                            return "taxins"
    if re.search(r"contract", s):                                                 return "contract"
    if re.search(r"other income|other revenue|misc.*income|ancillary", s):       return "otherinc"
    if re.search(r"utility billback|rubs|reimburs|recover", s):                   return "rubs"
    if re.search(r"concession", s):                                               return "concession"
    if re.search(r"rental adjustment|adjustments|vacancy|loss", s):              return "rentadj"
    if re.search(r"gross potential|gross rent|rental (income|revenue)|scheduled rent", s): return "rent"
    if re.search(r"write.?off", s):                                               return "writeoff"
    return None


# ---------------------------------------------------------------------------
# 3. WITHIN-FAMILY SPLITTERS (keyword refinement once the family is known)
# ---------------------------------------------------------------------------
def _split_payroll(n):
    # Employee apartment concession / housing allowance = value of a free/discounted
    # employee unit. RedIQ books this as a payroll bonus (PBo), not base payroll.
    if re.search(r"concession|apartment allowance|employee (apartment|unit|housing)|"
                 r"housing allowance", n):                                       return "PBo"
    if re.search(r"bonus|commission|incentive", n):                              return "PBo"
    if re.search(r"fica|payroll tax|medicare|social security|health|401|benefit|"
                 r"insurance|unemploy|sui|futa|suta|burden|pension|retirement", n): return "PB"
    return "Pay"  # salaries, wages, manager/leasing/maintenance, workers comp -> Pay (RedIQ convention)

def _split_utilities(n):
    if re.search(r"trash|garbage|refuse|waste|valet", n):                        return "cont"  # operator-specific; RedIQ files trash hauling in Contract
    if re.search(r"water|sewer", n):                                             return "UWS"
    if re.search(r"billing|transfer|audit|util.*fee|admin fee", n):              return "UF"
    if re.search(r"electric|gas|common|vacant|cable|energy|power", n):           return "UC"
    return "UC"

def _split_maintenance(n):
    # turnover: make-ready / unit turn work
    if re.search(r"turn|make.?ready|carpet clean|paint|interior repairs?.*unit|"
                 r"unit clean|vinyl|resurfac|key|lock", n):                      return "turn"
    # contract services (recurring vendor contracts)
    if re.search(r"landscap|pest|extermin|snow|elevator|fire alarm|fire suppl|"
                 r"sprinkler|courtesy patrol|security|contract clean|package|"
                 r"valet|trash haul|fitness|pool service|monitor", n):           return "cont"
    # everything else physical R&M -> interior/exterior
    return "inter/exte"

def _split_taxins(n):
    if re.search(r"insurance", n):                                               return "ins"
    if re.search(r"tax", n):                                                     return "ret"
    return "ret"

def _split_otherinc(n):
    # Month-to-month / short-term-lease premiums are a rent premium; RedIQ pulls
    # these up into Rental Income rather than leaving them in Other Income.
    if re.search(r"month.?to.?month|short.?term lease|\bmtm\b.*(fee|premium|rent)|"
                 r"(fee|premium).*\bmtm\b", n):                                   return "Rentinc"
    if re.search(r"parking|carport|garage|storage", n):                          return "park"
    if re.search(r"water.*reimb|sewer.*reimb|rubs.*water|water.*billback", n):    return "RWS"
    if re.search(r"trash.*reimb|rubs.*trash|trash.*billback|valet.*trash|trash.*income", n): return "RT"
    if re.search(r"utility reimburs|util.*billback|rubs|electric.*reimb|gas.*reimb|"
                 r"utility management|vacant recovery|corporate utilit", n):      return "RF"
    if re.search(r"cable|revenue shar|internet income|bulk.*income", n):         return "cable"
    if re.search(r"corporate (apartment|housing|unit) income", n):               return "CI"
    return "OI"

def _split_rubs(n):
    if re.search(r"water|sewer", n):                                             return "RWS"
    if re.search(r"trash|garbage", n):                                           return "RT"
    return "RF"

def _split_rentadj(n):
    if re.search(r"non.?revenue|down unit|model unit|employee unit|office unit|"
                 r"admin unit|staff unit|mgr unit|manager unit", n):             return "nr"
    if re.search(r"vacanc", n):                                                  return "vac"
    if re.search(r"concession|free rent|chargeback", n):                         return "conc"
    if re.search(r"loss to lease|gain to lease|loss/gain|gain/loss|market loss|market gain", n): return "ltl"
    if re.search(r"bad debt|collection|write.?off|uncollect|skip", n):           return "cl"
    return "vac"

def _split_rent(n):
    if re.search(r"loss to lease|gain to lease|loss/gain|gain/loss|loss to old lease|market loss|market gain", n): return "ltl"
    return "Rentinc"


# ---------------------------------------------------------------------------
# 4. PURE-KEYWORD FALLBACK (used when section is unknown)
# ---------------------------------------------------------------------------
REVENUE_RULES = [
    (r"loss to lease|gain to lease|loss/gain|gain/loss|market loss|market gain|loss to old lease", "ltl"),
    (r"non.?revenue|down unit|model unit|employee unit|office unit|admin unit|staff unit", "nr"),
    (r"vacanc", "vac"),
    (r"concession|free rent|chargeback concession", "conc"),
    (r"bad debt|collection loss|write.?off.*rent|rent.*write.?off|collected write|skip|uncollect", "cl"),
    (r"gross market rent|market rent|gross potential|gross rent|rental income|scheduled rent|base rent|^rent$", "Rentinc"),
    (r"month.?to.?month|mtm|short.?term lease|stl fee", "Rentinc"),
    (r"water.*reimb|sewer.*reimb|rubs.*water|water.*billback|sewer.*billback", "RWS"),
    (r"trash.*reimb|rubs.*trash|trash.*billback|valet.*trash|trash.*recovery|trash.*income", "RT"),
    (r"utility reimburs|utility recover|rubs|util.*billback|electric.*billback|gas.*billback|"
     r"utility management|vacant recovery|corporate utilit", "RF"),
    (r"parking|carport|garage|storage", "park"),
    (r"cable|revenue shar|bulk internet income|internet income", "cable"),
    (r"corporate (apartment|housing|unit) income", "CI"),
    (r".*", "OI"),
]
EXPENSE_RULES = [
    (r"bonus|commission|incentive", "PBo"),
    (r"fica|payroll tax|medicare|health|401|benefit|workers? comp|work comp|unemploy|sui|futa|burden", "PB"),
    (r"salary|salaries|wage|payroll|manager|leasing agent|maintenance (supervisor|technician|tech)|courtesy officer", "Pay"),
    (r"advertis|marketing|signage|brochure|locator|ils|promotion|referral|resident retention|"
     r"resident event|reputation|website|seo|model expense", "adv"),
    (r"management fee|mgmt fee|asset management fee", "mgt"),
    (r"make.?ready|carpet clean|paint(ing)? suppl|interior repairs?.*unit|unit clean|vinyl|resurfac|turn", "turn"),
    (r"water.?/?sewer|water and sewer|^water$|sewer", "UWS"),
    (r"utilit.*trash|trash.*utilit", "UT"),
    (r"utility billing|utility transfer|utility audit|billing fee", "UF"),
    (r"electric|gas|vacant unit utilit|common area electric", "UC"),
    (r"landscap|pest|extermin|courtesy patrol|security|alarm monitor|contract clean|snow removal|"
     r"pool service|fitness|package (service|locker)|valet|elevator|fire alarm|fire suppl|trash removal|trash haul", "cont"),
    (r"life safety", "Safe"),
    (r"hoa|homeowner|association fee", "HAF"),
    (r"repair|maintenance|hvac|plumb|electrical|appliance|window|door|screen|flooring|equipment|"
     r"hardware|lighting|key|lock|pool|spa|fountain|janitor|tools|building", "inter/exte"),
    (r"insurance", "ins"),
    (r"real estate tax|property tax|^tax|re tax|ad valorem", "ret"),
    (r".*", "GA"),
]
def _match(name, rules):
    low = (name or "").strip().lower()
    for pat, code in rules:
        if re.search(pat, low):
            return code
    return rules[-1][1]


# ---------------------------------------------------------------------------
# 5. CROSS-SECTION OVERRIDES (operators reliably misfile these)
# ---------------------------------------------------------------------------
def _override(name):
    n = (name or "").strip().lower()
    # forced-placed / tenant-liability insurance EXPENSE -> G&A (per methodology), not Contract/Insurance
    if re.search(r"forced.?placed|tenant liability insurance|legal liability insurance", n):
        return "GA"
    # uniforms -> G&A even when filed under payroll (RedIQ convention; it's a supply)
    if re.search(r"uniform", n):
        return "GA"
    return None


def categorize_t12_line(name, side, section_hint=None, acct_number=None):
    """Best-guess standardized code for a raw T12 line item.

    name         : raw account name (column A on the T12)
    side         : "rev" or "exp" -- which side of the T12 the line sits on
    section_hint : the operator's native subsection label, if known (preferred path)
    acct_number  : optional GL number (unused by default; available for tie-breaks)
    """
    ov = _override(name)
    if ov:
        return ov
    n = (name or "").strip().lower()
    fam = _family_from_section(section_hint)
    if fam == "payroll":     return _split_payroll(n)
    if fam == "marketing":   return "adv"
    if fam == "admin":       return "GA"
    if fam == "utilities":   return _split_utilities(n)
    if fam == "maintenance": return _split_maintenance(n)
    if fam == "mgmt":        return "mgt"
    if fam == "taxins":      return _split_taxins(n)
    if fam == "contract":    return "cont"
    if fam == "otherinc":    return _split_otherinc(n)
    if fam == "rubs":        return _split_rubs(n)
    if fam == "concession":  return "conc"
    if fam == "rentadj":     return _split_rentadj(n)
    if fam == "rent":        return _split_rent(n)
    if fam == "writeoff":    return "cl"
    # no recognized section -> pure keyword fallback by side
    return _match(name, REVENUE_RULES if side == "rev" else EXPENSE_RULES)


# ---------------------------------------------------------------------------
# 6. RENT-ROLL CHARGE-CODE -> (code, is_contract_rent, is_recurring)
#    Contract rent = recurring rent charges (base Rent + Amenity/premium rent).
#    Everything else is other income / reimbursements (separate T12 lines).
# ---------------------------------------------------------------------------
CHARGE_RULES = [
    (r"amenity rent|premium|view premium|floor premium|upgrade premium", "Rentinc", True,  True),
    (r"^rent$|base rent|market rent|gross rent|apartment rent",          "Rentinc", True,  True),
    (r"loss to lease|gain to lease|loss/gain",                           "ltl",  False, True),
    (r"concession|free rent|employee concession",                        "conc", False, True),
    (r"parking|carport|garage|storage",                                  "park", False, True),
    (r"valet trash|trash service|^trash",                                "RT",   False, True),
    (r"water|sewer",                                                     "RWS",  False, True),
    (r"utility reimburs|rubs|util.*billback|utility management|electric|gas", "RF", False, True),
    (r"technology|tech package|cable|internet|media",                    "OI",   False, True),
    (r"pet rent",                                                        "OI",   False, True),
    (r"pet fee|pet charge|pet deposit",                                  "OI",   False, False),
    (r"month to month|mtm",                                              "OI",   False, True),
    (r"insurance|renters? liab|renters? insurance|homebody|deposit alternative", "OI", False, True),
    (r"application fee",                                                 "OI",   False, False),
    (r"admin|holding fee",                                               "OI",   False, False),
    (r"late fee|nsf",                                                    "OI",   False, False),
    (r"lease termination|lease buy|early term",                          "OI",   False, False),
    (r"attorney|legal|filing fee|eviction",                             "OI",   False, False),
    (r"maintenance charge|resident maintenance|damage|key|lock",         "OI",   False, False),
    (r"referral",                                                        "OI",   False, False),
    (r"clubroom|facility rental|clubhouse",                              "OI",   False, False),
    (r".*",                                                              "OI",   False, False),
]
def categorize_charge(name):
    low = (name or "").strip().lower()
    for pat, code, is_cr, is_rec in CHARGE_RULES:
        if re.search(pat, low):
            return code, is_cr, is_rec
    return "OI", False, False
