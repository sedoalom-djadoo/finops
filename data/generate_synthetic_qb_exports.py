"""
Generates FICTIONAL QuickBooks-style exports for the Financial Dashboard demo.

The files deliberately reproduce the awkward layout of real QuickBooks Online and
billing-team exports (title rows to skip, account header rows that must be
filled down, totals mixed with detail, inconsistent sponsor spellings, typos in
study codes, USD invoices on a separate A/R account) so that the Power Query
cleaning logic of the model can be demonstrated on realistic input.

Every company, clinic, sponsor, study, lender and amount is invented.
Deterministic: fixed seed. Usage: python generate_synthetic_qb_exports.py
"""
from collections import defaultdict
from datetime import date, timedelta
from pathlib import Path
import calendar

import numpy as np
from openpyxl import Workbook

SEED = 7
rng = np.random.default_rng(SEED)
OUT = Path(__file__).parent / "raw"
OUT.mkdir(exist_ok=True)
(OUT / "billing").mkdir(exist_ok=True)

COMPANY = "Maplecrest Clinical Research Inc."
FY_STARTS = [2022, 2023, 2024, 2025]                    # FY2023 .. FY2026 (Sep-Aug)
SNAPSHOT_END = date(2026, 7, 9)                          # last data date (FY2026 is partial)
# Same three clinics as the clinical-operations dashboard. Vancouver opened in October 2024.
SITES = {"TOR": "Toronto", "CAL": "Calgary", "VAN": "Vancouver"}
SITE_CODE = {v: k for k, v in SITES.items()}
SITE_YEARS = {2022: ["TOR", "CAL"], 2023: ["TOR", "CAL"], 2024: ["VAN", "TOR", "CAL"], 2025: ["VAN", "TOR", "CAL"]}
SITE_ACTIVE = {"TOR": (date(2022, 9, 1), date(2026, 8, 31)), "CAL": (date(2022, 9, 1), date(2026, 8, 31)),
               "VAN": (date(2024, 10, 1), date(2026, 8, 31))}
COST_SCALE = 0.27
USD_CAD = {}  # filled from the model's public Bank of Canada table (approximation for synthetic use)
for y, m, r in [(2022, 9, 1.3389), (2022, 10, 1.3722), (2022, 11, 1.3419), (2022, 12, 1.3577), (2023, 1, 1.3414), (2023, 2, 1.3475),
                (2023, 3, 1.3696), (2023, 4, 1.3493), (2023, 5, 1.3527), (2023, 6, 1.3295), (2023, 7, 1.3219), (2023, 8, 1.3491),
                (2023, 9, 1.3524), (2023, 10, 1.3741), (2023, 11, 1.3720), (2023, 12, 1.3421), (2024, 1, 1.3430), (2024, 2, 1.3501),
                (2024, 3, 1.3555), (2024, 4, 1.3700), (2024, 5, 1.3681), (2024, 6, 1.3700), (2024, 7, 1.3720), (2024, 8, 1.3624),
                (2024, 9, 1.3541), (2024, 10, 1.3783), (2024, 11, 1.3987), (2024, 12, 1.4224), (2025, 1, 1.4408), (2025, 2, 1.4291),
                (2025, 3, 1.4381), (2025, 4, 1.3982), (2025, 5, 1.3873), (2025, 6, 1.3686), (2025, 7, 1.3724), (2025, 8, 1.3810),
                (2025, 9, 1.3838), (2025, 10, 1.3987), (2025, 11, 1.4059), (2025, 12, 1.3825), (2026, 1, 1.3808), (2026, 2, 1.3651),
                (2026, 3, 1.3717), (2026, 4, 1.3751), (2026, 5, 1.3723), (2026, 6, 1.4040), (2026, 7, 1.3700), (2026, 8, 1.3700)]:
    USD_CAD[(y, m)] = r

# ------------------------------------------------------------------ sponsors, studies
# canonical name, currency, billing-file spellings, QB customer names (first one used for invoices), CRO
# Same seven sponsors as the clinical-operations dashboard.
# canonical name, currency, billing-file spellings, QB customer names (first one used for invoices), CRO
SPONSORS = {
    "Aurelis Pharma":         ("CAD", ["Aurelis Pharma", "AURELIS", "Aurelis"], ["Aurelis Pharma Canada Inc.", "Aurelis - Trial Alias"], "Cambria CRO"),
    "Nordvale Therapeutics":  ("USD", ["Nordvale Therapeutics", "NORDVALE", "Nordvale"], ["Nordvale Therapeutics, USD", "Nordvale Therapeutics Inc. (deleted)"], ""),
    "Corvant Pharma":         ("CAD", ["Corvant Pharma"], ["Corvant Pharma ULC"], "Vantage Clinical"),
    "Helix Ridge Bio":        ("USD", ["Helix Ridge Bio", "Helix Ridge", "HelixRidge"], ["Helix Ridge Bio Inc., USD", "HELIX RIDGE BIO"], ""),
    "Cobalt Neuro":           ("CAD", ["Cobalt Neuro", "Cobalt"], ["Cobalt Neuro Ltd."], "Northstar Research"),
    "Larkspur Biotech":       ("CAD", ["Larkspur Biotech", "Larkspur"], ["Larkspur Biotech Inc."], "Vantage Clinical"),
    "Northbridge Memory Lab": ("CAD", ["Northbridge Memory Lab", "Northbridge"], ["Northbridge University - Memory Lab"], ""),
}
NON_SPONSORS = [("PayrollCo Services", "Payroll vendor recorded as a customer"),
                ("Global Grants Admin Inc.", "Expense re-invoicing (dry ice, monitoring visits), not clinical revenue"),
                ("Memory Research Foundation", "Grant income, not a trial sponsor")]
# Studies that closed before the CTMS was in place (not in the clinical-operations dataset)
LEGACY_STUDIES = {  # code -> (sponsor, site, start, end, monthly CAD mean)
    "STUDY-901": ("Aurelis Pharma", "TOR", date(2022, 9, 1), date(2025, 8, 31), 26000),
    "STUDY-902": ("Nordvale Therapeutics", "CAL", date(2022, 9, 1), date(2025, 6, 30), 15000),
    "STUDY-903": ("Corvant Pharma", "TOR", date(2022, 10, 1), date(2024, 6, 30), 13000),
    "STUDY-904": ("Helix Ridge Bio", "CAL", date(2022, 9, 1), date(2025, 9, 30), 11000),
    "STUDY-905": ("Cobalt Neuro", "TOR", date(2022, 11, 1), date(2025, 5, 31), 9000),
    "STUDY-906": ("Larkspur Biotech", "CAL", date(2023, 5, 1), date(2025, 7, 31), 8000),
    "STUDY-907": ("Aurelis Pharma", "CAL", date(2023, 1, 1), date(2025, 10, 31), 18000),
    "STUDY-908": ("Northbridge Memory Lab", "TOR", date(2023, 9, 1), date(2025, 6, 30), 7000),
    "STUDY-909": ("Corvant Pharma", "VAN", date(2024, 10, 1), date(2025, 8, 31), 9000),
}
# Studies from 2025 onward: revenue by study, site and month taken from the clinical-operations dataset
CLINOPS_REVENUE = Path(__file__).parent / "source" / "clinops_revenue_by_study_site_month.csv"
LEGACY_SCALE = 1.35
BAD_STUDY = ["-", "??", "two protocols?", "which cohort ??"]


def billing_spelling(code):
    """Billing-team spellings of protocol codes: STUDY-007 may be typed 'Study-007', 'STUDY 007' or 'STUDY-0007'."""
    r = rng.random()
    n = code.split("-")[1]
    if r < 0.12: return f"Study-{n}"
    if r < 0.20: return f"STUDY {n}"
    if r < 0.25: return f"STUDY-0{n}"
    return code


def months(start, end):
    y, m = start.year, start.month
    while date(y, m, 1) <= end:
        yield y, m
        y, m = (y + 1, 1) if m == 12 else (y, m + 1)


def fy_start(y, m):
    return y if m >= 9 else y - 1


ALL_MONTHS = list(months(date(2022, 9, 1), SNAPSHOT_END))
MONTH_HDR = ["Sept", "Oct", "Nov", "Dec", "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug"]

# ------------------------------------------------------------------ billing revenue (study x month, CAD)
billing = []  # (fy_start, site, sponsor, study, y, m, amount)
for code, (sp, site, st_, en, mean) in LEGACY_STUDIES.items():
    st_ = max(st_, SITE_ACTIVE[site][0]); en = min(en, SITE_ACTIVE[site][1], SNAPSHOT_END)
    ramp = 0
    span = list(months(st_, en))
    for k, (y, m) in enumerate(span):
        ramp = min(ramp + 0.25, 1.0)
        taper = min(1.0, (len(span) - k) / 6)          # close-out: fewer visits in the last six months
        amt = 0.0 if rng.random() < 0.06 else max(0.0, rng.normal(mean * LEGACY_SCALE, mean * LEGACY_SCALE * 0.28) * ramp * taper)
        billing.append((fy_start(y, m), site, sp, code, y, m, round(amt, 2)))
import csv
with open(CLINOPS_REVENUE, newline="") as fh:
    for r in csv.DictReader(fh):
        y, m = int(r["year"]), int(r["month"])
        if date(y, m, 1) > SNAPSHOT_END:
            continue
        billing.append((fy_start(y, m), SITE_CODE[r["site"]], r["sponsor"], r["study_id"], y, m, round(float(r["revenue"]), 2)))

# ------------------------------------------------------------------ chart of accounts
COA = [  # number, name, QB account type, detail type, currency
    ("10100", "Operating bank account", "Bank", "Chequing", "CAD"),
    ("10200", "USD bank account", "Bank", "Chequing", "USD"),
    ("12000", "Accounts Receivable (A/R)", "Accounts receivable (A/R)", "Accounts Receivable (A/R)", "CAD"),
    ("12001", "Accounts Receivable (A/R) - USD", "Accounts receivable (A/R)", "Accounts Receivable (A/R)", "USD"),
    ("13000", "Prepaid expenses", "Current assets", "Prepaid Expenses", "CAD"),
    ("13100", "SR&ED credit receivable", "Current assets", "Other Current Assets", "CAD"),
    ("15000", "Medical equipment", "Property, plant and equipment", "Machinery and equipment", "CAD"),
    ("15100", "Computer equipment", "Property, plant and equipment", "Machinery and equipment", "CAD"),
    ("15200", "Leasehold improvements", "Property, plant and equipment", "Leasehold Improvements", "CAD"),
    ("15900", "Accumulated depreciation", "Property, plant and equipment", "Accumulated Depreciation", "CAD"),
    ("20105", "Security deposit - Toronto", None, None, "CAD"),
    ("20109", "Security deposit - Vancouver", None, None, "CAD"),
    ("20112", "Advances to employees", None, None, "CAD"),
    ("20500", "Long-term deposit", None, None, "CAD"),
    ("21000", "Accounts Payable (A/P)", "Accounts payable (A/P)", "Accounts Payable (A/P)", "CAD"),
    ("21500", "Corporate credit card", "Credit Card", "Credit Card", "CAD"),
    ("22000", "Accrued liabilities", "Other Current Liabilities", "Accrued Liabilities", "CAD"),
    ("22100", "Payroll liabilities", "Other Current Liabilities", "Payroll Clearing", "CAD"),
    ("22500", "Deferred revenue", "Other Current Liabilities", "Other Current Liabilities", "CAD"),
    ("30115", "Line of credit", None, None, "CAD"),
    ("30171", "Current portion of long-term debt", None, None, "CAD"),
    ("30184", "Term loan - Harbour Bank", "Long-term Liabilities", "Notes Payable", "CAD"),
    ("30185", "Equipment loan - Harbour Bank", "Long-term Liabilities", "Notes Payable", "CAD"),
    ("30189", "Growth loan - Summit Capital", None, None, "CAD"),
    ("30190", "Equipment loan - Summit Capital", "Long-term Liabilities", "Notes Payable", "CAD"),
    ("30191", "Leasehold loan - Summit Capital", "Long-term Liabilities", "Notes Payable", "CAD"),
    ("31000", "Share capital", "Equity", "Common Stock", "CAD"),
    ("32000", "Retained earnings", "Equity", "Retained Earnings", "CAD"),
    ("40001", "Income", "Income", "Service/Fee Income", "CAD"),
    ("40098", "Interest income", "Other Income", "Interest Earned", "CAD"),
    ("40200", "Grant income", "Other Income", "Other Miscellaneous Income", "CAD"),
    ("40210", "Re-invoiced expenses", "Other Income", "Other Miscellaneous Income", "CAD"),
    ("50100", "Lab and diagnostic services", "Cost of Goods Sold", "Cost of labour - COS", "CAD"),
    ("50120", "Imaging services", "Cost of Goods Sold", "Other Costs of Services - COS", "CAD"),
    ("50150", "Study supplies", "Cost of Goods Sold", "Supplies & Materials - COS", "CAD"),
    ("50155", "Courier and dry ice", "Cost of Goods Sold", "Shipping - COS", "CAD"),
    ("50160", "Patient stipends", "Cost of Goods Sold", "Other Costs of Services - COS", "CAD"),
    ("50184", "SR&ED tax credit", "Cost of Goods Sold", "Other Costs of Services - COS", "CAD"),
    ("50185", "SR&ED tax credit - prior year", "Cost of Goods Sold", "Other Costs of Services - COS", "CAD"),
    ("60005", "Subcontractors", "Expenses", "Contract labour", "CAD"),
    ("60010", "Salaries - clinical staff", "Expenses", "Payroll Expenses", "CAD"),
    ("60020", "Salaries - administration", "Expenses", "Payroll Expenses", "CAD"),
    ("60030", "Employee benefits", "Expenses", "Payroll Expenses", "CAD"),
    ("60040", "Group insurance", "Expenses", "Payroll Expenses", "CAD"),
    ("60050", "DPSP employer contributions", "Expenses", "Payroll Expenses", "CAD"),
    ("60110", "Advertising", "Expenses", "Advertising/Promotional", "CAD"),
    ("60120", "Business meals", "Expenses", "Meals and entertainment", "CAD"),
    ("60200", "Travel", "Expenses", "Travel", "CAD"),
    ("60250", "Parking", "Expenses", "Travel", "CAD"),
    ("60300", "Conferences", "Expenses", "Travel", "CAD"),
    ("60501", "Rent or lease", "Expenses", "Rent or Lease of Buildings", "CAD"),
    ("70100", "Office supplies", "Expenses", "Office/General Administrative Expenses", "CAD"),
    ("70125", "Software", "Expenses", "Office/General Administrative Expenses", "CAD"),
    ("70130", "Telecommunications", "Expenses", "Utilities", "CAD"),
    ("70140", "Old storage account (deleted)", "Expenses", "Office/General Administrative Expenses", "CAD"),
    ("70300", "Equipment maintenance", "Expenses", "Repair & Maintenance", "CAD"),
    ("70310", "Premises maintenance", "Expenses", "Repair & Maintenance", "CAD"),
    ("70400", "Business insurance", "Expenses", "Insurance", "CAD"),
    ("70600", "Permits and municipal taxes", "Expenses", "Taxes Paid", "CAD"),
    ("70800", "Accounting fees", "Expenses", "Legal & Professional Fees", "CAD"),
    ("70830", "Legal fees", "Expenses", "Legal & Professional Fees", "CAD"),
    ("70900", "Memberships", "Expenses", "Dues & subscriptions", "CAD"),
    ("80001", "Severance", "Expenses", "Other Miscellaneous Service Cost", "CAD"),
    ("80010", "Depreciation - furniture", "Expenses", "Depreciation", "CAD"),
    ("80020", "Depreciation - computers", "Expenses", "Depreciation", "CAD"),
    ("80030", "Amortization - leasehold improvements", "Expenses", "Depreciation", "CAD"),
    ("80505", "Interest on long-term debt", "Expenses", "Interest Paid", "CAD"),
    ("80520", "Bank fees", "Expenses", "Bank Charges", "CAD"),
    ("80540", "Payroll processing fees", "Expenses", "Bank Charges", "CAD"),
    ("81001", "Exchange gain or loss", "Other Expense", "Exchange Gain or Loss", "CAD"),
    ("90900", "Income taxes - current", "Other Expense", "Other Miscellaneous Expense", "CAD"),
]
NAME = {a[0]: a[1] for a in COA}
NAME["70120"] = "Cloud storage"         # exists in the GL but missing from the chart of accounts (orphan)

# ------------------------------------------------------------------ invoices (QB) from billing
inv_rows = []          # (entity, date, type, number, posted, memo, ar_account, amount_in_invoice_ccy, sponsor, cad_amount)
by_sponsor_month = defaultdict(float)
for fy, site, sp, code, y, m, amt in billing:
    by_sponsor_month[(sp, y, m)] += amt
inv_no = 10450
for (sp, y, m), cad in sorted(by_sponsor_month.items(), key=lambda k: (k[0][1], k[0][2], k[0][0])):
    if cad <= 0:
        continue
    lag = int(rng.choice([0, 0, 0, 1, 1, 2]))                      # billing team books the month, QB invoices later
    iy, im = (y + (m + lag - 1) // 12, (m + lag - 1) % 12 + 1)
    last = calendar.monthrange(iy, im)[1]
    d = date(iy, im, int(rng.integers(max(1, last - 9), last + 1)))
    if d > SNAPSHOT_END:
        continue
    ccy, _, qb_names, _ = SPONSORS[sp]
    entity = qb_names[0] if rng.random() < 0.85 or len(qb_names) == 1 else qb_names[1]
    ar = "12001 Accounts Receivable (A/R) - USD" if ccy == "USD" else "12000 Accounts Receivable (A/R)"
    amount = round(cad / USD_CAD[(iy, im)], 2) if ccy == "USD" else round(cad, 2)
    inv_no += 1
    inv_rows.append((entity, d, "Invoice", str(inv_no), "Yes", f"Site visits {calendar.month_abbr[m]} {y}", ar, amount, sp, round(cad, 2)))
for name, _ in NON_SPONSORS:                                       # noise that must be filtered out
    for _ in range(int(rng.integers(3, 8))):
        y, m = ALL_MONTHS[int(rng.integers(0, len(ALL_MONTHS)))]
        inv_no += 1
        inv_rows.append((name, date(y, m, 15), "Invoice", str(inv_no), "Yes", "Re-invoicing", "12000 Accounts Receivable (A/R)",
                         round(float(rng.uniform(800, 9000)), 2), None, None))

# ------------------------------------------------------------------ general ledger
gl = defaultdict(list)    # account -> [(date, type, ref, name, class, description, split, amount)]
ref = 5000


def post(acct, d, typ, name, desc, amount, cls=None, split="10100 Operating bank account"):
    global ref
    ref += 1
    gl[acct].append((d, typ, str(ref), name, cls, desc, split, round(amount, 2)))


for e, d, typ, no, _, memo, ar, amt, sp, cad in inv_rows:
    if sp is None:
        post("40210", d, typ, e, memo, amt, split=ar)                 # re-invoicing / grants are not clinical revenue
    else:
        post("40001", d, typ, e, memo, cad, split=ar)


def active_sites(y, m):
    d = date(y, m, 1)
    return [s for s, (a, b) in SITE_ACTIVE.items() if a <= d <= b]


for y, m in ALL_MONTHS:
    last = min(calendar.monthrange(y, m)[1], SNAPSHOT_END.day if (y, m) == (2026, 7) else 31)
    d15, dend = date(y, m, min(15, last)), date(y, m, last)
    sites = active_sites(y, m)
    rev = sum(a for f, s, sp, c, yy, mm, a in billing if (yy, mm) == (y, m))
    # direct research costs: vendors, tagged with a Class only ~8% of the time (data quality finding)
    for acct, share, vendor in [("50100", 0.11, "Lab vendor"), ("50120", 0.05, "Imaging partner"), ("50150", 0.03, "Medical supplier"),
                                ("50155", 0.012, "Courier"), ("50160", 0.02, "Patient stipends")]:
        for s in sites:
            amt = rev * share / len(sites) * rng.uniform(0.75, 1.25)
            post(acct, dend, "Bill", vendor, f"{NAME[acct]} - {SITES[s]}", amt, SITES[s] if rng.random() < 0.08 else None, "21000 Accounts Payable (A/P)")
    post("60005", dend, "Bill", "Freelance coordinators", "Contract coordinators", rev * 0.04 * rng.uniform(0.6, 1.4), None, "21000 Accounts Payable (A/P)")
    clinical = 55000 * COST_SCALE * len(sites) * rng.uniform(0.95, 1.05)
    for half in (d15, dend):
        post("60010", half, "Payroll Check", "Payroll", "Clinical staff payroll", clinical / 2, SITES[sites[0]] if rng.random() < 0.08 else None, "22100 Payroll liabilities")
        post("60020", half, "Payroll Check", "Payroll", "Administration payroll", 21000 * COST_SCALE / 2 * rng.uniform(0.97, 1.03), None, "22100 Payroll liabilities")
    post("60030", dend, "Journal Entry", "Payroll", "Benefits", clinical * 0.09, None, "22100 Payroll liabilities")
    post("60040", dend, "Bill", "Group insurer", "Group insurance", 6500 * COST_SCALE * len(sites) / 2, None, "21000 Accounts Payable (A/P)")
    post("60050", dend, "Journal Entry", "Payroll", "DPSP employer contribution", clinical * 0.03, None, "22100 Payroll liabilities")
    for s in sites:
        post("60501", date(y, m, 1), "Bill", f"{SITES[s]} landlord", f"Rent - {SITES[s]}", {"TOR": 9000, "CAL": 6000, "VAN": 7000}[s], SITES[s], "21000 Accounts Payable (A/P)")
    for acct, base, payee in [("60110", 2200, "Ad agency"), ("60120", 900, "Restaurants"), ("60200", 2600, "Airline"), ("60250", 450, "Parking"),
                              ("70100", 1800, "Office supplier"), ("70125", 4200, "Software vendor"), ("70130", 2100, "Telecom"),
                              ("70300", 1500, "Biomed services"), ("70310", 1100, "Facilities"), ("70400", 3300, "Insurer"),
                              ("70800", 3800, "Accounting firm"), ("70900", 600, "Associations"), ("70120", 350, "Cloud provider"),
                              ("80520", 380, "Bank"), ("80540", 620, "PayrollCo Services")]:
        post(acct, dend, "Expense", payee, NAME[acct], base * COST_SCALE * rng.uniform(0.6, 1.5), None, "21500 Corporate credit card")
    if m in (10, 4):
        post("60300", dend, "Expense", "Conference organizer", "Scientific conference", 9000 * COST_SCALE * rng.uniform(0.7, 1.3))
        post("70830", dend, "Bill", "Law firm", "Contract review", 7000 * COST_SCALE * rng.uniform(0.5, 1.5), None, "21000 Accounts Payable (A/P)")
    if m == 1:
        post("70600", dend, "Bill", "Municipality", "Permits and municipal taxes", 14000 * COST_SCALE, None, "21000 Accounts Payable (A/P)")
    for acct, amt in [("80010", 800), ("80020", 1100), ("80030", 2200)]:
        post(acct, dend, "Journal Entry", "", "Monthly depreciation", amt, None, "15900 Accumulated depreciation")
    post("80505", dend, "Journal Entry", "Harbour Bank", "Loan interest", (7400 - (y - 2022) * 900) * COST_SCALE)
    post("40098", dend, "Deposit", "Bank", "Interest earned", 400 * rng.uniform(0.5, 1.5))
    usd_ar = sum(a for e_, d_, t_, n_, p_, me_, ar_, a, sp_, c_ in inv_rows if ar_.startswith("12001") and (d_.year, d_.month) == (y, m))
    post("81001", dend, "Journal Entry", "", "FX revaluation", usd_ar * rng.normal(0, 0.012))
    if m == 8 or (y, m) == (2026, 7):
        post("50184", dend, "Journal Entry", "", "SR&ED credit accrual", -rev * 0.0 - 95000 * rng.uniform(0.8, 1.2), None, "13100 SR&ED credit receivable")
        post("40200", dend, "Deposit", "Memory Research Foundation", "Research grant", 25000 * COST_SCALE * rng.uniform(0.6, 1.4))
    if (y, m) == (2024, 2):
        post("50185", dend, "Journal Entry", "", "SR&ED reassessment FY2022", -42000, None, "13100 SR&ED credit receivable")
    if m == 8:
        post("90900", dend, "Journal Entry", "", "Income tax provision", 0)        # set below from pre-tax profit

# income tax per FY (15 % of positive pre-tax profit)
SIGN = {"Revenue": 1, "OtherIncome": 1}


def category_type(n):
    n = int(n)
    if n == 40001: return "Revenue"
    if n in (40098, 40200, 40210): return "OtherIncome"
    if n == 81001: return "NonOperating"
    if n in (90900, 90910): return "Tax"
    if 80010 <= n <= 80100 and n not in (80001,): return "Depreciation"
    if 80505 <= n <= 80540: return "Financial"
    return "Expense"


def fy_net(fy, upto_tax=True):
    tot = 0
    for acct, rows in gl.items():
        t = category_type(acct)
        for r in rows:
            if fy_start(r[0].year, r[0].month) != fy or (t == "Tax" and upto_tax):
                continue
            tot += r[7] if t in ("Revenue", "OtherIncome") else -r[7]
    return tot


for fy in FY_STARTS:
    pre = fy_net(fy)
    for i, r in enumerate(gl["90900"]):
        if fy_start(r[0].year, r[0].month) == fy:
            gl["90900"][i] = r[:7] + (round(max(pre, 0) * 0.15, 2),)

# ------------------------------------------------------------------ writers
def wb_save(wb, path):
    if "Sheet" in wb.sheetnames and len(wb.sheetnames) > 1:
        del wb["Sheet"]
    wb.save(path)


# 1. Chart of accounts (Sheet2, 4 title rows)
wb = Workbook(); ws = wb.active; ws.title = "Sheet2"
for r in [[COMPANY], ["Account List"], ["All dates"], []]:
    ws.append(r)
ws.append(["Account #", "Full name", "Type", "Detail type", "Currency", "Description"])
for a in COA:
    ws.append(list(a))
ws.append([None, None])
ws.append(["Total accounts", len(COA)])
wb.save(OUT / "Chart_of_Accounts.xlsx")

# 2. General ledger (QB "Transaction Detail by Account": account header rows, detail, totals)
wb = Workbook(); ws = wb.active; ws.title = "Sheet1"
ws.append([COMPANY]); ws.append(["Transaction Detail by Account"]); ws.append([f"September 1, 2022 - {SNAPSHOT_END:%B %-d, %Y}"]); ws.append([])
ws.append([None, "Transaction date", "Transaction type", "Num", "Name", "Class full name", "Memo/Description", "Split account", "Amount", "Balance"])
for acct in sorted(gl, key=int):
    label = f"{acct} {NAME[acct]}"
    ws.append([label])
    bal = 0
    for d, typ, refno, name, cls, desc, split, amt in sorted(gl[acct], key=lambda r: r[0]):
        bal += amt
        ws.append([None, d, typ, refno, name, cls, desc, split, amt, round(bal, 2)])
    ws.append([f"Total for {label}", None, None, None, None, None, None, None, round(bal, 2)])
ws.append([]); ws.append([f"Generated {SNAPSHOT_END:%Y-%m-%d} (synthetic data)"])
wb.save(OUT / "General_Ledger_Transaction_Detail.xlsx")

# 3. P&L by month (Sheet1, 4 title rows, header row with blank first cell)
pl = defaultdict(float)
for acct, rows in gl.items():
    for r in rows:
        pl[(acct, r[0].year, r[0].month)] += r[7]
labels = [f"{calendar.month_abbr[m]} {y}" for y, m in ALL_MONTHS]
wb = Workbook(); ws = wb.active; ws.title = "Sheet1"
ws.append([COMPANY]); ws.append(["Profit and Loss by Month"]); ws.append([f"September 2022 - {SNAPSHOT_END:%B %Y}"]); ws.append([])
ws.append([" "] + labels + ["Total"])
sections = [("Income", lambda a: category_type(a) in ("Revenue", "OtherIncome") and int(a) < 40200 or a == "40001"),
            ("Cost of Goods Sold", lambda a: a.startswith("50")),
            ("Expenses", lambda a: a[0] in "678" and a != "81001"),
            ("Other Income", lambda a: a in ("40200", "40210")),
            ("Other Expenses", lambda a: a in ("81001", "90900"))]
done = set()
for title, pred in sections:
    ws.append([title])
    for acct in sorted(gl, key=int):
        if acct in done or not pred(acct):
            continue
        done.add(acct)
        vals = [round(pl.get((acct, y, m), 0), 2) or None for y, m in ALL_MONTHS]
        ws.append([f"{acct} {NAME[acct]}"] + vals + [round(sum(v or 0 for v in vals), 2)])
    ws.append([f"Total for {title}"])
ws.append(["Net Income"])
wb.save(OUT / "Profit_and_Loss_by_Month.xlsx")

# 4. P&L statement per fiscal year (one sheet per FY, used for the unrealized FX line)
AS_OF = {2022: date(2023, 8, 31), 2023: date(2024, 8, 31), 2024: date(2025, 8, 31), 2025: SNAPSHOT_END}
wb = Workbook()
for fy in FY_STARTS:
    ws = wb.create_sheet(str(fy + 1))
    ws.append([COMPANY]); ws.append(["Profit and Loss"]); ws.append([f"Fiscal year ending {AS_OF[fy]:%Y-%m-%d}"]); ws.append([])
    ws.append(["Net Operating Income", round(fy_net(fy, upto_tax=False) + sum(r[7] for r in gl["81001"] if fy_start(r[0].year, r[0].month) == fy), 2)])
    ws.append(["Other Expenses"])
    ws.append(["Exchange Gain or Loss", round(sum(r[7] for r in gl["81001"] if fy_start(r[0].year, r[0].month) == fy), 2)])
    ws.append(["Unrealized Gain or Loss", round(float(rng.normal(0, 9000)), 2)])
    ws.append(["Net Income", round(fy_net(fy, upto_tax=False), 2)])
wb_save(wb, OUT / "Profit_and_Loss_by_FY.xlsx")

# 5. Balance sheet per fiscal year (5 title rows; "Profit for the year" has no account number)
wb = Workbook()
retained = 480_000.0
ppe = {"15000": 270_000.0, "15100": 90_000.0, "15200": 220_000.0}
acc_dep = -160_000.0
debt = {"30184": 380_000.0, "30185": 75_000.0, "30189": 190_000.0, "30190": 60_000.0, "30191": 110_000.0}
for fy in FY_STARTS:
    profit = round(fy_net(fy, upto_tax=False), 2)
    acc_dep -= sum(r[7] for a in ("80010", "80020", "80030") for r in gl[a] if fy_start(r[0].year, r[0].month) == fy)
    for k in debt:
        debt[k] = round(debt[k] * (0.84 if k in ("30184", "30185") else 0.88), 2)
    asof = AS_OF[fy]
    last2 = [(asof.year, asof.month), ((asof.year if asof.month > 1 else asof.year - 1), (asof.month - 1 or 12))]
    ar_cad = round(sum(c for e, d, t, n, p, me, ar, a, sp, c in inv_rows if sp and not ar.startswith("12001") and (d.year, d.month) in last2), 2)
    ar_usd = round(sum(c for e, d, t, n, p, me, ar, a, sp, c in inv_rows if sp and ar.startswith("12001") and (d.year, d.month) in last2), 2)
    monthly_cost = sum(r[7] for a in gl if a[0] in "56" for r in gl[a] if (r[0].year, r[0].month) == last2[0])
    liabilities = {"21000": round(monthly_cost * 0.45, 2), "21500": round(float(rng.uniform(15000, 30000)), 2),
                   "22000": round(monthly_cost * 0.2, 2), "22100": round(monthly_cost * 0.12, 2), "22500": round(float(rng.uniform(60000, 140000)), 2),
                   "30115": round(float(rng.choice([0, 0, 75000])), 2), "30171": round(sum(debt.values()) * 0.12, 2), **debt}
    equity = {"31000": 100.0, "32000": round(retained, 2)}
    assets_wo_bank = {"12000": ar_cad, "12001": ar_usd, "13000": round(float(rng.uniform(40000, 70000)), 2),
                      "13100": round(-sum(r[7] for r in gl["50184"] if fy_start(r[0].year, r[0].month) == fy), 2),
                      **{k: v for k, v in ppe.items()}, "15900": round(acc_dep, 2),
                      "20105": 42000.0, "20109": 31000.0 if fy >= 2024 else 0.0, "20112": round(float(rng.uniform(2000, 6000)), 2), "20500": 25000.0}
    total_le = sum(liabilities.values()) + sum(equity.values()) + profit
    bank = round(total_le - sum(assets_wo_bank.values()), 2)
    usd_bank = round(bank * 0.15, 2)
    assets = {"10100": round(bank - usd_bank, 2), "10200": usd_bank, **assets_wo_bank}
    ws = wb.create_sheet(str(fy + 1))
    ws.append([COMPANY]); ws.append(["Balance Sheet"]); ws.append([f"As of {asof:%B %-d, %Y}"]); ws.append([]); ws.append([None, "Total"])
    ws.append(["Assets"])
    for k, v in assets.items():
        if v:
            ws.append([f"{k} {NAME[k]}", v])
    ws.append(["Total for Assets", round(sum(assets.values()), 2)])
    ws.append(["Liabilities and Equity"])
    for k, v in liabilities.items():
        if v:
            ws.append([f"{k} {NAME[k]}", v])
    ws.append(["Total for Liabilities", round(sum(liabilities.values()), 2)])
    for k, v in equity.items():
        ws.append([f"{k} {NAME[k]}", v])
    ws.append(["Profit for the year", profit])
    ws.append(["Total for Equity", round(sum(equity.values()) + profit, 2)])
    ws.append(["Total for Liabilities and Equity", round(total_le, 2)])
    retained += profit
    if fy == 2024:
        ppe["15200"] += 130_000.0                                  # Vancouver fit-out
wb_save(wb, OUT / "Balance_Sheet_by_FY.xlsx")

# 6. Transaction list by customer (two sheets, 4 title rows, entity header rows, dates as US text)
by_entity = defaultdict(list)
for row in inv_rows:
    by_entity[row[0]].append(row)
for e in by_entity:                                                  # payments that must be filtered out
    for row in list(by_entity[e])[::3]:
        by_entity[e].append((e, row[1] + timedelta(days=30), "Payment", "", "Yes", "Payment received", row[6], -row[7], None, None))
entities = sorted(by_entity)
wb = Workbook()
for i, chunk in enumerate([entities[: len(entities) // 2], entities[len(entities) // 2:]], 1):
    ws = wb.create_sheet(f"Sheet{i}")
    ws.append([COMPANY]); ws.append(["Transaction List by Customer"]); ws.append(["All dates"]); ws.append([])
    ws.append([None, "Date", "Transaction type", "#", "Posting (Y/N)", "Memo", "Account full name", "Amount"])
    for e in chunk:
        ws.append([e])
        rows = sorted(by_entity[e], key=lambda r: r[1])
        for r in rows:
            ws.append([None, r[1].strftime("%m/%d/%Y"), r[2], r[3], r[4], r[5], r[6], r[7]])
        ws.append([f"Total for {e}", None, None, None, None, None, None, round(sum(r[7] for r in rows), 2)])
wb_save(wb, OUT / "Transaction_List_by_Customer.xlsx")

# 7. Sponsor mapping (title row, two note rows, header on row 4)
wb = Workbook(); ws = wb.active; ws.title = "DimSponsorMapping"
ws.append(["DimSponsorMapping - billing-file spellings and QB customer names mapped to one sponsor"])
ws.append(["Maintained by Finance. IncludeInRevenue = FALSE for vendors, grants and re-invoicing entities."])
ws.append([])
ws.append(["SponsorRaw_Billing", "SponsorRaw_QB", "SponsorNameNormalized", "IncludeInRevenue", "Notes", "Status"])
for canon, (ccy, bill_names, qb_names, cro) in SPONSORS.items():
    for b in bill_names:
        for q in qb_names:
            ws.append([b, q, canon, "TRUE", "USD invoices on A/R 12001" if ccy == "USD" else None, "Validated"])
for name, note in NON_SPONSORS:
    ws.append([None, name, name, "FALSE", note, "Excluded"])
wb.save(OUT / "DimSponsorMapping.xlsx")

# 8. Billing revenue files (one workbook per FY, one sheet per site, messy spellings)
for fy in FY_STARTS:
    wb = Workbook()
    for site in SITE_YEARS[fy]:
        ws = wb.create_sheet(f"{site} Revenues")
        ws.append([f"{SITES[site]} Memory Clinic - total revenues FY {fy}-{fy + 1}"]); ws.append(["Prepared by the billing team"]); ws.append([])
        with_pharma = fy >= 2024
        ws.append(["CRO", "Sponsor", "Study"] + (["Pharma"] if with_pharma else []) + MONTH_HDR + ["TOTAL"])
        rows = defaultdict(lambda: [0.0] * 12)
        for f, s, sp, code, y, m, amt in billing:
            if f == fy and s == site:
                rows[(sp, code)][MONTH_HDR.index("Sept" if m == 9 else calendar.month_abbr[m])] += amt
        grand = [0.0] * 12
        for (sp, code), vals in sorted(rows.items()):
            ccy, bill_names, _, cro = SPONSORS[sp]
            spon = bill_names[int(rng.integers(0, len(bill_names)))]
            study = billing_spelling(code)
            cro_txt = {"Cambria CRO": rng.choice(["Cambria CRO", "Cambira", "cambria"]), "Vantage Clinical": rng.choice(["Vantage Clinical", "Vantage"])}.get(cro, cro or "N/A")
            cells = [round(v, 2) if v else None for v in vals]
            ws.append([cro_txt, spon, study] + (["Yes"] if with_pharma else []) + cells + [round(sum(vals), 2)])
            grand = [g + v for g, v in zip(grand, vals)]
        if rng.random() < 0.7:                                       # a line nobody could attribute to a protocol
            sp = list(SPONSORS)[int(rng.integers(0, 5))]
            v = round(float(rng.uniform(1500, 6000)), 2)
            cells = [None] * 12; cells[int(rng.integers(0, 6))] = v
            ws.append(["N/A", SPONSORS[sp][1][0], rng.choice(BAD_STUDY)] + ([None] if with_pharma else []) + cells + [v])
        ws.append([None, None, "TOTAL"] + ([None] if with_pharma else []) + [round(g, 2) for g in grand] + [round(sum(grand), 2)])
        ws.append([]); ws.append([None, None, "Legend: blank month = not billed"])
    wb_save(wb, OUT / "billing" / f"Billing_Revenue_FY{fy + 1}.xlsx")

print("billing rows:", len(billing), "| invoices:", len(inv_rows), "| GL accounts:", len(gl),
      "| GL rows:", sum(len(v) for v in gl.values()))
for fy in FY_STARTS:
    print(f"FY{fy + 1}: billing revenue {sum(a for f, *_, a in billing if f == fy):,.0f}  net income {fy_net(fy, upto_tax=False):,.0f}")
