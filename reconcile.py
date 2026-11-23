"""Prove the migration: read Business Central back through its API and compare to QuickBooks.

Usage:
    python reconcile.py            # run every check
    python reconcile.py --break    # break test: tamper with one expected number and prove a check FAILS

Needs these set (see docs/RECONCILE.md). Leave BC_CLIENT_SECRET unset to sign in as yourself:
    BC_TENANT_ID, BC_CLIENT_ID, [BC_CLIENT_SECRET], BC_ENVIRONMENT (e.g. Production), BC_COMPANY (e.g. Craigs Landscaping MIG)

Checks (plain version):
1. Record counts: same number of accounts, customers, vendors and items as went in.
2. A/R total and A/R per customer: what each customer owes in BC = what they owed in QBO.
3. A/P per vendor: what we owe each vendor in BC = QBO.
4. Inventory: quantity on hand per item in BC = QBO.
5. Trial balance: every account's balance in BC = QBO's trial balance on go-live day,
   and the clearing account 3900 is exactly zero.
"""
import csv
import json
import os
import sys

import msal
import requests

BREAK = "--break" in sys.argv
TOL = 0.005
CLEARING_NO = "3900"

# ---------------------------------------------------------------- BC API
tenant = os.environ["BC_TENANT_ID"]
authority = f"https://login.microsoftonline.com/{tenant}"
scopes = ["https://api.businesscentral.dynamics.com/.default"]
if os.environ.get("BC_CLIENT_SECRET"):
    # App-only sign-in (needs an admin to grant consent).
    app = msal.ConfidentialClientApplication(os.environ["BC_CLIENT_ID"], authority=authority,
                                             client_credential=os.environ["BC_CLIENT_SECRET"])
    token = app.acquire_token_for_client(scopes=scopes)
else:
    # Sign in as yourself with a one-time code (no admin needed). Reads BC with your own BC permissions.
    app = msal.PublicClientApplication(os.environ["BC_CLIENT_ID"], authority=authority)
    flow = app.initiate_device_flow(scopes=scopes)
    if "user_code" not in flow:
        sys.exit(f"Could not start sign-in: {flow.get('error_description', flow)}")
    print(flow["message"], flush=True)  # "To sign in, open https://microsoft.com/devicelogin and enter the code ..."
    token = app.acquire_token_by_device_flow(flow)
if "access_token" not in token:
    sys.exit(f"Could not get a BC token: {token.get('error_description', token)}")
session = requests.Session()
session.headers.update({"Authorization": f"Bearer {token['access_token']}", "Accept": "application/json"})

base = (f"https://api.businesscentral.dynamics.com/v2.0/{tenant}/"
        f"{os.environ.get('BC_ENVIRONMENT', 'Production')}/api/v2.0")


def get_all(url, params=None):
    rows = []
    while url:
        r = session.get(url, params=params)
        if r.status_code != 200:
            sys.exit(f"BC API error {r.status_code} on {url}\n{r.text[:500]}")
        body = r.json()
        rows += body.get("value", [])
        url, params = body.get("@odata.nextLink"), None
    return rows


company_name = os.environ["BC_COMPANY"]
companies = get_all(f"{base}/companies")
company = next((c for c in companies if c["name"] == company_name), None)
if not company:
    sys.exit(f"Company '{company_name}' not found. BC has: {[c['name'] for c in companies]}")
api = f"{base}/companies({company['id']})"


def num(v):
    return round(float(v or 0), 2)


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

if BREAK:
    # Pretend QBO said one customer owed $1.00 more. A working reconciliation must catch this.
    victim = next(no for no, bal in exp_customer.items() if bal > 0)
    exp_customer[victim] = round(exp_customer[victim] + 1, 2)
    print(f"BREAK TEST: expected balance for {victim} raised by $1.00\n")

# ---------------------------------------------------------------- BC side (actual)
bc_accounts = get_all(f"{api}/accounts")
bc_customers = get_all(f"{api}/customers")
bc_vendors = get_all(f"{api}/vendors")
bc_items = get_all(f"{api}/items")
bc_ar = {r["customerNumber"]: num(r["balanceDue"]) for r in get_all(f"{api}/agedAccountsReceivables") if r.get("customerNumber")}
bc_ap = {r["vendorNumber"]: num(r["balanceDue"]) for r in get_all(f"{api}/agedAccountsPayables") if r.get("vendorNumber")}
bc_tb = {}
for r in get_all(f"{api}/trialBalance"):
    if r.get("number"):
        bc_tb[r["number"]] = round(num(r.get("balanceAtDateDebit")) - num(r.get("balanceAtDateCredit")), 2)

# ---------------------------------------------------------------- checks
results = []


def check(name, got, want, detail=""):
    ok = abs(got - want) < TOL if isinstance(want, float) else got == want
    results.append(ok)
    print(f"{'PASS' if ok else 'FAIL'}  {name}: BC {got} vs QBO {want}{('  ' + detail) if detail else ''}")


def compare_map(name, actual, expected):
    """Compare per-record balances; report the total and every mismatch."""
    bad = [(k, actual.get(k, 0.0), v) for k, v in expected.items() if abs(actual.get(k, 0.0) - v) >= TOL]
    extra = [(k, a) for k, a in actual.items() if k not in expected and abs(a) >= TOL]
    check(f"{name} total", round(sum(actual.values()), 2), round(sum(expected.values()), 2))
    ok = not bad and not extra
    results.append(ok)
    print(f"{'PASS' if ok else 'FAIL'}  {name} per record: {len(expected) - len(bad)}/{len(expected)} match")
    for k, a, e in bad:
        print(f"        {k}: BC {a} vs QBO {e}")
    for k, a in extra:
        print(f"        {k}: in BC ({a}) but not in QBO")


print(f"Company: {company_name}\n")
print("-- record counts")
from openpyxl import load_workbook  # noqa: E402  (only needed for this one count)
gl_loaded = load_workbook("out/bc_load.xlsx", read_only=True)["15 G-L Account"].max_row - 1
check("G/L accounts (incl. headings and totals)", len(bc_accounts), gl_loaded)
check("Customers", len(bc_customers), len(q_customers))
check("Vendors", len(bc_vendors), len(q_vendors))
check("Items", len(bc_items), len(q_items))

print("\n-- receivables and payables")
compare_map("A/R by customer", bc_ar, exp_customer)
# BC reports what we owe vendors as a positive balance due; QBO does too.
compare_map("A/P by vendor", bc_ap, exp_vendor)

print("\n-- inventory")
bc_qty = {i["number"]: num(i.get("inventory")) for i in bc_items if i["number"] in exp_qty}
compare_map("Quantity on hand", bc_qty, exp_qty)

print("\n-- trial balance")
compare_map("Trial balance by account", {k: v for k, v in bc_tb.items() if k in exp_tb or abs(v) >= TOL}, exp_tb)
check("Clearing account 3900 is zero", bc_tb.get(CLEARING_NO, 0.0), 0.0)

passed = sum(results)
print(f"\n{passed}/{len(results)} checks passed")
if BREAK:
    print("Break test " + ("WORKED: the tampered number was caught." if passed < len(results)
                           else "DID NOT WORK: nothing failed, so the checks are too weak."))
    sys.exit(0 if passed < len(results) else 1)
sys.exit(0 if passed == len(results) else 1)
