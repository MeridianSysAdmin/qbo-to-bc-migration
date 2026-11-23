"""Pull the QuickBooks Trial Balance report (the source for opening balances).

Usage: python extract_tb.py 2026-11-23      (the go-live date)
Needs QBO_ACCESS_TOKEN and QBO_REALM_ID set, same as extract.py.
"""
import json
import os
import sys

import requests

go_live = sys.argv[1]
token = os.environ["QBO_ACCESS_TOKEN"]
realm_id = os.environ["QBO_REALM_ID"]

url = f"https://sandbox-quickbooks.api.intuit.com/v3/company/{realm_id}/reports/TrialBalance"
params = {"end_date": go_live, "minorversion": "75"}
headers = {"Authorization": f"Bearer {token}", "Accept": "application/json"}

response = requests.get(url, params=params, headers=headers)
print("Status:", response.status_code)
response.raise_for_status()
report = response.json()

# Flatten to [{account, debit, credit}] so the other scripts don't need to know the report format.
rows = []
for row in report.get("Rows", {}).get("Row", []):
    cols = row.get("ColData")
    if not cols or row.get("type") == "Section":
        continue
    name = cols[0].get("value", "")
    account_id = cols[0].get("id")
    debit = float(cols[1].get("value") or 0)
    credit = float(cols[2].get("value") or 0)
    rows.append({"account": name, "id": account_id, "debit": debit, "credit": credit})

total_dr = round(sum(r["debit"] for r in rows), 2)
total_cr = round(sum(r["credit"] for r in rows), 2)
print(f"{len(rows)} accounts | debits {total_dr} | credits {total_cr} | {'BALANCED' if total_dr == total_cr else 'NOT BALANCED'}")

with open("data/TrialBalance.json", "w", encoding="utf-8") as f:
    json.dump({"go_live": go_live, "rows": rows, "raw": report}, f, indent=2)
print("Saved to data/TrialBalance.json")
