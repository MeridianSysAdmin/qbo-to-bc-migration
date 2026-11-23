"""Turn the QuickBooks JSON pulls into Business Central load data.

Usage:
    python transform.py 2026-11-23                      # writes out/bc_load.xlsx + out/crossref.csv
    python transform.py 2026-11-23 package.xlsx         # also fills a Configuration Package export

Every rule here comes from mapping.csv. The go-live date is the posting date of every opening entry.

How the opening balances work (plain version):
- Each open customer invoice becomes one journal line, so aging still works (A/R).
- Each vendor with a balance becomes one journal line (A/P).
- Inventory comes in through the item journal (quantity x cost), which posts to the inventory account.
- Every other non-zero trial balance line becomes one G/L journal line.
- Every line balances against 3900 Opening Balance Clearing. If the migration is right, 3900 ends at exactly 0.
"""
import csv
import datetime as dt
import json
import re
import sys
from collections import defaultdict
from pathlib import Path

from openpyxl import Workbook, load_workbook

GO_LIVE = sys.argv[1] if len(sys.argv) > 1 else "2026-11-23"
TEMPLATE = sys.argv[2] if len(sys.argv) > 2 else None
OUT = Path("out")

CLEARING_NO = "3900"
UOM = "EA"                 # decision: QBO items have no unit of measure
GEN_TEMPLATE, GEN_BATCH = "GENERAL", "DEFAULT"
ITEM_TEMPLATE, ITEM_BATCH = "ITEM", "DEFAULT"


def load(entity):
    with open(f"data/{entity}.json", encoding="utf-8") as f:
        return json.load(f)["QueryResponse"].get(entity, [])


accounts = load("Account")
customers = load("Customer")
vendors = load("Vendor")
items = load("Item")
invoices = load("Invoice")
with open("data/TrialBalance.json", encoding="utf-8") as f:
    trial_balance = json.load(f)["rows"]

# ---------------------------------------------------------------- G/L accounts
# Numbering block per QBO account type, in the order they appear within the block.
BLOCKS = [
    (1000, ["Bank", "Accounts Receivable", "Other Current Asset", "Fixed Asset", "Other Asset"]),
    (2000, ["Accounts Payable", "Credit Card", "Other Current Liability", "Long Term Liability"]),
    (3000, ["Equity"]),
    (4000, ["Income"]),
    (5000, ["Cost of Goods Sold"]),
    (6000, ["Expense"]),
    (7000, ["Other Income", "Other Expense"]),
]
CATEGORY = {
    "Bank": "Assets", "Accounts Receivable": "Assets", "Other Current Asset": "Assets",
    "Fixed Asset": "Assets", "Other Asset": "Assets",
    "Accounts Payable": "Liabilities", "Credit Card": "Liabilities",
    "Other Current Liability": "Liabilities", "Long Term Liability": "Liabilities",
    "Equity": "Equity", "Income": "Income", "Other Income": "Income",
    "Cost of Goods Sold": "Cost of Goods Sold", "Expense": "Expense", "Other Expense": "Expense",
}
# Accounts BC fills through sub-ledgers; no direct posting allowed.
SUBLEDGER_TYPES = {"Accounts Receivable", "Accounts Payable"}
NO_DIRECT_POSTING_NAMES = {"Inventory Asset"}   # fed by item postings only

by_id = {a["Id"]: a for a in accounts}
children = defaultdict(list)
for a in accounts:
    if a.get("SubAccount"):
        children[a["ParentRef"]["value"]].append(a)
for kids in children.values():
    kids.sort(key=lambda a: a["Name"])

tb_by_id = {r["id"]: round(r["debit"] - r["credit"], 2) for r in trial_balance}  # + = debit
inventory_account_id = next(a["Id"] for a in accounts if a["Name"] == "Inventory Asset")

gl_rows = []          # BC G/L Account rows, in number order
qbo_to_bc = {}        # QBO account Id -> BC posting account No. (where its balance lands)


def add_account(no, name, acct, account_type, totaling=""):
    gl_rows.append({
        "No.": str(no),
        "Name": name[:100],
        "Income/Balance": "Balance Sheet" if acct["Classification"] in ("Asset", "Liability", "Equity") else "Income Statement",
        "Account Category": CATEGORY[acct["AccountType"]],
        "Account Type": account_type,
        "Totaling": totaling,
        "Direct Posting": "Yes" if account_type == "Posting" and acct["AccountType"] not in SUBLEDGER_TYPES
                          and acct["Name"] not in NO_DIRECT_POSTING_NAMES else "No",
        "QBO Id": acct["Id"] if account_type == "Posting" else "",
    })


def place(acct, counter):
    """Number an account and (if it has sub-accounts) its whole subtree. Returns the next free number."""
    kids = children.get(acct["Id"])
    if not kids:
        add_account(counter, acct["Name"], acct, "Posting")
        qbo_to_bc[acct["Id"]] = str(counter)
        return counter + 10
    begin = counter
    add_account(counter, acct["Name"], acct, "Begin-Total")
    counter += 10
    if tb_by_id.get(acct["Id"]):
        # The parent itself carries a balance in QBO (QBO shows it as "<Parent> - Other").
        add_account(counter, f"{acct['Name']} - Other", acct, "Posting")
        qbo_to_bc[acct["Id"]] = str(counter)
        counter += 10
    for kid in kids:
        counter = place(kid, counter)
    add_account(counter, f"Total {acct['Name']}", acct, "End-Total", f"{begin}..{counter}")
    return counter + 10


for start, types in BLOCKS:
    counter = start
    top = [a for a in accounts if not a.get("SubAccount") and a["AccountType"] in types]
    top.sort(key=lambda a: (types.index(a["AccountType"]), a["Name"]))
    for acct in top:
        counter = place(acct, counter)
    if start == 3000:
        gl_rows.append({"No.": CLEARING_NO, "Name": "Opening Balance Clearing", "Income/Balance": "Balance Sheet",
                        "Account Category": "Equity", "Account Type": "Posting", "Totaling": "",
                        "Direct Posting": "Yes", "QBO Id": ""})
    assert counter < start + 1000, f"block {start} overflowed"

missing = [a["FullyQualifiedName"] for a in accounts if a["Id"] not in qbo_to_bc and not children.get(a["Id"])]
assert not missing, f"accounts not placed: {missing}"

ar_no = next(r["No."] for r in gl_rows if r["QBO Id"] and by_id[r["QBO Id"]]["AccountType"] == "Accounts Receivable")
ap_no = next(r["No."] for r in gl_rows if r["QBO Id"] and by_id[r["QBO Id"]]["AccountType"] == "Accounts Payable")
inv_no = qbo_to_bc[inventory_account_id]
sales_no = qbo_to_bc[next(a["Id"] for a in accounts if a["Name"] == "Sales of Product Income")]
cogs_no = qbo_to_bc[next(a["Id"] for a in accounts if a["AccountType"] == "Cost of Goods Sold")]

# ---------------------------------------------------------------- setup tables
setup = {
    "Unit of Measure": [{"Code": UOM, "Description": "Each"}],
    "Customer Posting Group": [{"Code": "DOMESTIC", "Receivables Account": ar_no}],
    "Vendor Posting Group": [{"Code": "DOMESTIC", "Payables Account": ap_no}],
    "Gen. Product Posting Group": [{"Code": "RETAIL", "Description": "Inventory items"},
                                   {"Code": "SERVICES", "Description": "Service items"}],
    "Inventory Posting Group": [{"Code": "RESALE", "Description": "Resale items"}],
    "General Posting Setup": [
        # The inventory opening entry balances against the clearing account.
        {"Gen. Bus. Posting Group": "", "Gen. Prod. Posting Group": "RETAIL",
         "Sales Account": sales_no, "COGS Account": cogs_no, "Inventory Adjmt. Account": CLEARING_NO},
        {"Gen. Bus. Posting Group": "", "Gen. Prod. Posting Group": "SERVICES",
         "Sales Account": sales_no, "COGS Account": cogs_no, "Inventory Adjmt. Account": CLEARING_NO},
    ],
    "Inventory Posting Setup": [{"Location Code": "", "Invt. Posting Group Code": "RESALE", "Inventory Account": inv_no}],
}

# ---------------------------------------------------------------- customers / vendors / items
def party_row(p, prefix, posting_group_field):
    addr = p.get("BillAddr", {})
    return {
        "No.": f"{prefix}{int(p['Id']):04d}",
        "Name": p["DisplayName"][:100],
        "Name 2": "",
        "Address": addr.get("Line1", "")[:100],
        "City": addr.get("City", "")[:30],
        "Post Code": addr.get("PostalCode", "")[:20],
        "County": addr.get("CountrySubDivisionCode", "")[:30],
        "Phone No.": p.get("PrimaryPhone", {}).get("FreeFormNumber", "")[:30],
        "E-Mail": p.get("PrimaryEmailAddr", {}).get("Address", "")[:80],
        posting_group_field: "DOMESTIC",
        "QBO Id": p["Id"],
    }


cust_by_id = {c["Id"]: c for c in customers}
customer_rows = []
for c in sorted(customers, key=lambda c: int(c["Id"])):
    row = party_row(c, "C", "Customer Posting Group")
    if c.get("Job"):  # decision: sub-customers become their own customers, parent kept in Name 2
        row["Name 2"] = cust_by_id[c["ParentRef"]["value"]]["DisplayName"][:50]
    customer_rows.append(row)
cust_no = {r["QBO Id"]: r["No."] for r in customer_rows}

vendor_rows = [party_row(v, "V", "Vendor Posting Group") for v in sorted(vendors, key=lambda v: int(v["Id"]))]
vend_no = {r["QBO Id"]: r["No."] for r in vendor_rows}

item_rows = []
for it in sorted(items, key=lambda i: int(i["Id"])):
    inventory = it["Type"] == "Inventory"
    item_rows.append({
        "No.": f"I{int(it['Id']):04d}",
        "Description": it["Name"][:100],
        "Description 2": (it.get("Description") or "")[:50],
        "Type": "Inventory" if inventory else "Service",
        "Base Unit of Measure": UOM,
        "Unit Price": it.get("UnitPrice", 0),
        "Unit Cost": it.get("PurchaseCost", 0),
        "Gen. Prod. Posting Group": "RETAIL" if inventory else "SERVICES",
        "Inventory Posting Group": "RESALE" if inventory else "",
        "QBO Id": it["Id"],
    })
item_no = {r["QBO Id"]: r["No."] for r in item_rows}

# ---------------------------------------------------------------- opening journals
gen_lines, line_no = [], 0


def gen_line(**kw):
    global line_no
    line_no += 10000
    base = {"Journal Template Name": GEN_TEMPLATE, "Journal Batch Name": GEN_BATCH, "Line No.": line_no,
            "Posting Date": GO_LIVE, "Document Date": GO_LIVE, "Document Type": "", "Document No.": "",
            "External Document No.": "", "Account Type": "G/L Account", "Account No.": "", "Description": "",
            "Amount": 0, "Due Date": "", "Bal. Account Type": "G/L Account", "Bal. Account No.": CLEARING_NO}
    base.update(kw)
    gen_lines.append(base)


# A/R: one line per open invoice (decision: open items only)
for inv in sorted(invoices, key=lambda i: (i["TxnDate"], i.get("DocNumber", ""))):
    if inv["Balance"] > 0:
        gen_line(**{"Document Type": "Invoice", "Document No.": f"OB-{inv['DocNumber']}",
                    "External Document No.": inv["DocNumber"], "Document Date": inv["TxnDate"],
                    "Due Date": inv.get("DueDate", ""), "Account Type": "Customer",
                    "Account No.": cust_no[inv["CustomerRef"]["value"]],
                    "Description": f"Opening balance QBO invoice {inv['DocNumber']}", "Amount": inv["Balance"]})

# A/P: one line per vendor with a balance
for v in sorted(vendors, key=lambda v: int(v["Id"])):
    if v.get("Balance"):
        gen_line(**{"Document Type": "Invoice", "Document No.": f"OB-{vend_no[v['Id']]}",
                    "Account Type": "Vendor", "Account No.": vend_no[v["Id"]],
                    "Description": "Opening balance (QBO vendor balance)", "Amount": -round(v["Balance"], 2)})

# Everything else on the trial balance, straight to G/L
sub_ledger_ids = {a["Id"] for a in accounts if a["AccountType"] in SUBLEDGER_TYPES} | {inventory_account_id}
for r in trial_balance:
    amount = round(r["debit"] - r["credit"], 2)
    if amount and r["id"] not in sub_ledger_ids:
        gen_line(**{"Document No.": "OB-GL", "Account No.": qbo_to_bc[r["id"]],
                    "Description": f"Opening balance {r['account']}"[:100], "Amount": amount})

# Inventory: quantity x cost through the item journal
item_lines = []
for i, it in enumerate([i for i in items if i["Type"] == "Inventory" and i.get("QtyOnHand")], start=1):
    item_lines.append({"Journal Template Name": ITEM_TEMPLATE, "Journal Batch Name": ITEM_BATCH, "Line No.": i * 10000,
                       "Posting Date": GO_LIVE, "Entry Type": "Positive Adjmt.", "Document No.": "OB-INV",
                       "Item No.": item_no[it["Id"]], "Quantity": it["QtyOnHand"], "Unit Amount": it["PurchaseCost"]})

# ---------------------------------------------------------------- self-checks
tb_ar = tb_by_id[next(a["Id"] for a in accounts if a["AccountType"] == "Accounts Receivable")]
tb_ap = tb_by_id[next(a["Id"] for a in accounts if a["AccountType"] == "Accounts Payable")]
tb_inv = tb_by_id[inventory_account_id]
loaded_ar = round(sum(l["Amount"] for l in gen_lines if l["Account Type"] == "Customer"), 2)
loaded_ap = round(sum(l["Amount"] for l in gen_lines if l["Account Type"] == "Vendor"), 2)
loaded_inv = round(sum(l["Quantity"] * l["Unit Amount"] for l in item_lines), 2)
clearing = round(sum(l["Amount"] for l in gen_lines) + loaded_inv, 2)

checks = [
    ("A/R lines = trial balance A/R", loaded_ar, tb_ar),
    ("A/P lines = trial balance A/P", loaded_ap, tb_ap),
    ("Inventory qty x cost = trial balance inventory", loaded_inv, tb_inv),
    ("Clearing account 3900 nets to zero", clearing, 0.0),
]
ok = True
for name, got, want in checks:
    passed = abs(got - want) < 0.005
    ok &= passed
    print(f"{'PASS' if passed else 'FAIL'}  {name}: {got} vs {want}")

# ---------------------------------------------------------------- output
def cell(v):
    """Excel needs real dates, not 'YYYY-MM-DD' text, or BC reads them in the wrong format."""
    if isinstance(v, str) and re.fullmatch(r"\d{4}-\d{2}-\d{2}", v):
        return dt.date.fromisoformat(v)
    return v


TABLES = [  # (BC table ID, table name, rows) in the order BC must import them
    (204, "Unit of Measure", setup["Unit of Measure"]),
    (15, "G/L Account", gl_rows),
    (251, "Gen. Product Posting Group", setup["Gen. Product Posting Group"]),
    (94, "Inventory Posting Group", setup["Inventory Posting Group"]),
    (92, "Customer Posting Group", setup["Customer Posting Group"]),
    (93, "Vendor Posting Group", setup["Vendor Posting Group"]),
    (252, "General Posting Setup", setup["General Posting Setup"]),
    (5813, "Inventory Posting Setup", setup["Inventory Posting Setup"]),
    (18, "Customer", customer_rows),
    (23, "Vendor", vendor_rows),
    (27, "Item", item_rows),
    (81, "Gen. Journal Line", gen_lines),
    (83, "Item Journal Line", item_lines),
]
OUT.mkdir(exist_ok=True)

wb = Workbook()
wb.remove(wb.active)
for table_id, name, rows in TABLES:
    ws = wb.create_sheet(f"{table_id} {name}".replace("/", "-")[:31])
    headers = [h for h in rows[0] if h != "QBO Id"]
    ws.append(headers)
    for r in rows:
        ws.append([cell(r[h]) for h in headers])
wb.save(OUT / "bc_load.xlsx")

with open(OUT / "crossref.csv", "w", newline="", encoding="utf-8") as f:
    w = csv.writer(f)
    w.writerow(["Entity", "QBO Id", "BC No.", "Name"])
    for r in gl_rows:
        if r["QBO Id"]:
            w.writerow(["G/L Account", r["QBO Id"], r["No."], r["Name"]])
    for ent, rows in (("Customer", customer_rows), ("Vendor", vendor_rows)):
        for r in rows:
            w.writerow([ent, r["QBO Id"], r["No."], r["Name"]])
    for r in item_rows:
        w.writerow(["Item", r["QBO Id"], r["No."], r["Description"]])

print(f"\n{len(gl_rows)} G/L lines, {len(customer_rows)} customers, {len(vendor_rows)} vendors, {len(item_rows)} items, "
      f"{len(gen_lines)} journal lines, {len(item_lines)} item journal lines")
print(f"Wrote {OUT / 'bc_load.xlsx'} and {OUT / 'crossref.csv'}")


# ---------------------------------------------------------------- optional: fill a BC package export
def fill_template(path):
    """Fill the Excel file exported from a BC Configuration Package.

    Each sheet in the export is one table. We find the header row by matching field captions,
    then write our rows underneath. Columns we don't have are left blank.
    """
    def key(s):  # sheet titles lose characters like "/" and get cut at 31 chars
        return "".join(ch for ch in s.lower() if ch.isalnum())[:20]

    tpl = load_workbook(path)
    by_name = {key(name): rows for _, name, rows in TABLES}
    for ws in tpl.worksheets:
        rows = by_name.get(key(ws.title))
        if rows is None:
            print(f"  skip sheet '{ws.title}' (no data for it)")
            continue
        wanted = {h for h in rows[0] if h != "QBO Id"}
        header_row, cols = None, {}
        for r in range(1, 6):
            found = {ws.cell(r, c).value: c for c in range(1, ws.max_column + 1) if ws.cell(r, c).value in wanted}
            if len(found) > len(cols):
                header_row, cols = r, found
        if not cols:
            print(f"  ! sheet '{ws.title}': no matching headers, check the package field list")
            continue
        missing_cols = wanted - set(cols)
        if missing_cols:
            print(f"  ! sheet '{ws.title}': package is missing fields {sorted(missing_cols)}")
        for i, row in enumerate(rows, start=header_row + 1):
            for h, c in cols.items():
                ws.cell(i, c).value = cell(row[h])
        print(f"  filled '{ws.title}': {len(rows)} rows")
    out = OUT / "package_filled.xlsx"
    tpl.save(out)
    print(f"Wrote {out}")


if TEMPLATE:
    fill_template(TEMPLATE)

sys.exit(0 if ok else 1)
