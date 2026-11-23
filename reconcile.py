"""Prove the migration: read Business Central back and compare it to QuickBooks.

Usage:
    python reconcile.py                    # read BC from Excel exports in bc_export/ (default)
    python reconcile.py --break            # break test: tamper with one expected number and prove a check FAILS
    python reconcile.py --api [--break]    # read BC live through its API (needs a tenant where the app can sign in)

Excel exports (see docs/RECONCILE.md): from each BC list page, Share -> Open in Excel, saved into bc_export/ as
    coa.xlsx        Chart of Accounts   (No., Balance)
    customers.xlsx  Customers           (No., Balance)
    vendors.xlsx    Vendors             (No., Balance)
    items.xlsx      Items               (No., Inventory)

API mode needs BC_TENANT_ID, BC_CLIENT_ID, [BC_CLIENT_SECRET], BC_ENVIRONMENT, BC_COMPANY.

Checks (plain version):
1. Record counts: same number of accounts, customers, vendors and items as went in.
2. A/R total and per customer: what each customer owes in BC = what they owed in QBO.
3. A/P total and per vendor: what we owe each vendor in BC = QBO.
4. Inventory: quantity on hand per item in BC = QBO.
5. Trial balance: every account's balance in BC = QBO's trial balance on go-live day,
   and the clearing account 3900 is exactly zero.
"""
import csv
import json
import os
import sys
from pathlib import Path

from openpyxl import load_workbook

BREAK = "--break" in sys.argv
USE_API = "--api" in sys.argv
EXPORT_DIR = Path("bc_export")
TOL = 0.005
CLEARING_NO = "3900"


def num(v):
    if v in (None, ""):
        return 0.0
    if isinstance(v, str):
        v = v.replace(",", "").replace("$", "").strip()
        if v.startswith("(") and v.endswith(")"):
            v = "-" + v[1:-1]
    return round(float(v), 2)


# ---------------------------------------------------------------- BC side: Excel exports
def read_export(filename, value_header):
    """Return {No.: value} from a BC 'Open in Excel' file.

    value_header is matched loosely ("Balance" also matches "Balance ($)"), and never
    matches "Balance Due" / "Balance at Date" style columns by accident.
    """
    path = EXPORT_DIR / filename
    if not path.exists():
        sys.exit(f"Missing {path}. Export it from BC (see docs/RECONCILE.md).")
    wb = load_workbook(path, data_only=True, read_only=True)
    for ws in wb.worksheets:
        rows = list(ws.iter_rows(values_only=True))
        for h, header in enumerate(rows[:15]):
            names = [str(c).strip() if c is not None else "" for c in header]
            if "No." not in names:
                continue
            no_col = names.index("No.")
            val_col = next((i for i, n in enumerate(names)
                            if n == value_header or n.startswith(value_header + " (")), None)
            if val_col is None:
                sys.exit(f"{path}: no '{value_header}' column. Columns are: {[n for n in names if n]}\n"
                         f"Add it to the page (Personalize) and export again.")
            out = {}
            for r in rows[h + 1:]:
                if r[no_col] not in (None, ""):
                    out[str(r[no_col]).strip()] = num(r[val_col])
            return out
    sys.exit(f"{path}: couldn't find a header row with 'No.'")


def bc_from_excel():
    coa = read_export("coa.xlsx", "Balance")
    customers = read_export("customers.xlsx", "Balance")
    vendors = read_export("vendors.xlsx", "Balance")
    items = read_export("items.xlsx", "Inventory")
    return {
        "source": f"Excel exports in {EXPORT_DIR}/",
        "counts": {"accounts": len(coa), "customers": len(customers), "vendors": len(vendors), "items": len(items)},
        "ar": customers, "ap": vendors, "qty": items, "tb": coa,
    }


# ---------------------------------------------------------------- BC side: live API
def bc_from_api():
    import msal
    import requests

    tenant = os.environ["BC_TENANT_ID"]
    authority = f"https://login.microsoftonline.com/{tenant}"
    scopes = ["https://api.businesscentral.dynamics.com/.default"]
    if os.environ.get("BC_CLIENT_SECRET"):
        app = msal.ConfidentialClientApplication(os.environ["BC_CLIENT_ID"], authority=authority,
                                                 client_credential=os.environ["BC_CLIENT_SECRET"])
        token = app.acquire_token_for_client(scopes=scopes)
    else:
        app = msal.PublicClientApplication(os.environ["BC_CLIENT_ID"], authority=authority)
        flow = app.initiate_device_flow(scopes=scopes)
        if "user_code" not in flow:
            sys.exit(f"Could not start sign-in: {flow.get('error_description', flow)}")
        print(flow["message"], flush=True)
        token = app.acquire_token_by_device_flow(flow)
    if "access_token" not in token:
        sys.exit(f"Could not get a BC token: {token.get('error_description', token)}")
    session = requests.Session()
    session.headers.update({"Authorization": f"Bearer {token['access_token']}", "Accept": "application/json"})
    base = (f"https://api.businesscentral.dynamics.com/v2.0/{tenant}/"
            f"{os.environ.get('BC_ENVIRONMENT', 'Production')}/api/v2.0")

    def get_all(url):
        rows = []
        while url:
            r = session.get(url)
            if r.status_code != 200:
                sys.exit(f"BC API error {r.status_code} on {url}\n{r.text[:500]}")
            body = r.json()
            rows += body.get("value", [])
            url = body.get("@odata.nextLink")
        return rows

    name = os.environ["BC_COMPANY"]
    companies = get_all(f"{base}/companies")
    company = next((c for c in companies if c["name"] == name), None)
    if not company:
        sys.exit(f"Company '{name}' not found. BC has: {[c['name'] for c in companies]}")
    api = f"{base}/companies({company['id']})"

    accounts, customers = get_all(f"{api}/accounts"), get_all(f"{api}/customers")
    vendors, items = get_all(f"{api}/vendors"), get_all(f"{api}/items")
    tb = {}
    for r in get_all(f"{api}/trialBalance"):
        if r.get("number"):
            tb[r["number"]] = round(num(r.get("balanceAtDateDebit")) - num(r.get("balanceAtDateCredit")), 2)
    return {
        "source": f"BC API, company {name}",
        "counts": {"accounts": len(accounts), "customers": len(customers), "vendors": len(vendors), "items": len(items)},
        "ar": {r["customerNumber"]: num(r["balanceDue"]) for r in get_all(f"{api}/agedAccountsReceivables") if r.get("customerNumber")},
        "ap": {r["vendorNumber"]: num(r["balanceDue"]) for r in get_all(f"{api}/agedAccountsPayables") if r.get("vendorNumber")},
        "qty": {i["number"]: num(i.get("inventory")) for i in items},
        "tb": tb,
    }


# ---------------------------------------------------------------- QBO side (expected)
def load(entity):
    with open(f"data/{entity}.json", encoding="utf-8") as f:
        return json.load(f)["QueryResponse"].get(entity, [])


q_customers, q_vendors, q_items = load("Customer"), load("Vendor"), load("Item")
with open("data/TrialBalance.json", encoding="utf-8") as f:
    q_tb = json.load(f)["rows"]

xref = {}  # (entity, QBO Id) -> BC No.
with open("out/crossref.csv", encoding="utf-8") as f:
    for row in csv.DictReader(f):
        xref[(row["Entity"], row["QBO Id"])] = row["BC No."]

exp_customer = {xref[("Customer", c["Id"])]: num(c.get("Balance")) for c in q_customers}
exp_vendor = {xref[("Vendor", v["Id"])]: num(v.get("Balance")) for v in q_vendors}
exp_qty = {xref[("Item", i["Id"])]: num(i.get("QtyOnHand")) for i in q_items if i["Type"] == "Inventory"}
exp_tb = {}  # BC account No. -> debit-positive balance
for r in q_tb:
    no = xref[("G/L Account", r["id"])]
    exp_tb[no] = round(exp_tb.get(no, 0) + r["debit"] - r["credit"], 2)
exp_tb[CLEARING_NO] = 0.0
gl_loaded = load_workbook("out/bc_load.xlsx", read_only=True)["15 G-L Account"].max_row - 1

if BREAK:
    # Pretend QBO said one customer owed $1.00 more. A working reconciliation must catch this.
    victim = next(no for no, bal in exp_customer.items() if bal > 0)
    exp_customer[victim] = round(exp_customer[victim] + 1, 2)
    print(f"BREAK TEST: expected balance for {victim} raised by $1.00\n")

bc = bc_from_api() if USE_API else bc_from_excel()

# ---------------------------------------------------------------- checks
results = []


def check(name, got, want):
    ok = abs(got - want) < TOL if isinstance(want, float) else got == want
    results.append(ok)
    print(f"{'PASS' if ok else 'FAIL'}  {name}: BC {got} vs QBO {want}")


def compare_map(name, actual, expected):
    """Compare per-record balances; report the total and every mismatch.

    Only records that exist on one side with a non-zero amount, or on both sides, are compared;
    BC lists can include rows (headings, totals, zero-balance records) that QBO's side doesn't carry.
    """
    keys = set(expected) | {k for k, a in actual.items() if abs(a) >= TOL}
    bad = [(k, actual.get(k, 0.0), expected.get(k, 0.0)) for k in sorted(keys)
           if abs(actual.get(k, 0.0) - expected.get(k, 0.0)) >= TOL]
    check(f"{name} total", round(sum(actual.get(k, 0.0) for k in keys), 2), round(sum(expected.values()), 2))
    results.append(not bad)
    print(f"{'PASS' if not bad else 'FAIL'}  {name} per record: {len(keys) - len(bad)}/{len(keys)} match")
    for k, a, e in bad:
        print(f"        {k}: BC {a} vs QBO {e}")


print(f"Reading BC from: {bc['source']}\n")
print("-- record counts")
check("G/L accounts (incl. headings and totals)", bc["counts"]["accounts"], gl_loaded)
check("Customers", bc["counts"]["customers"], len(q_customers))
check("Vendors", bc["counts"]["vendors"], len(q_vendors))
check("Items", bc["counts"]["items"], len(q_items))

print("\n-- receivables and payables")
compare_map("A/R by customer", bc["ar"], exp_customer)
compare_map("A/P by vendor", bc["ap"], exp_vendor)

print("\n-- inventory")
compare_map("Quantity on hand", {k: v for k, v in bc["qty"].items() if k in exp_qty or v}, exp_qty)

print("\n-- trial balance")
# Begin-Total / End-Total lines in the chart of accounts show subtotals; compare posting accounts only.
posting = {k for k in exp_tb} | {r["BC No."] for r in csv.DictReader(open("out/crossref.csv", encoding="utf-8"))
                                  if r["Entity"] == "G/L Account"}
compare_map("Trial balance by account", {k: v for k, v in bc["tb"].items() if k in posting}, exp_tb)
check("Clearing account 3900 is zero", bc["tb"].get(CLEARING_NO, 0.0), 0.0)

passed = sum(results)
print(f"\n{passed}/{len(results)} checks passed")
if BREAK:
    print("Break test " + ("WORKED: the tampered number was caught." if passed < len(results)
                           else "DID NOT WORK: nothing failed, so the checks are too weak."))
    sys.exit(0 if passed < len(results) else 1)
sys.exit(0 if passed == len(results) else 1)
