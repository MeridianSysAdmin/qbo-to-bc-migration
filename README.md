# qbo-to-bc-migration

Moves a small company from QuickBooks Online to Microsoft Dynamics 365 Business Central: chart of accounts, customers, vendors, items, open invoices, vendor balances, inventory and opening balances. A mapping sheet records every field decision, and a reconciliation reads Business Central back through its API to prove every number matches QuickBooks to the cent.

## How it works

| Step | File | What it does |
|---|---|---|
| 1. Extract | `extract.py`, `extract_tb.py` | Pulls Account, Customer, Vendor, Item and Invoice from the QuickBooks Online API (sandbox), plus the Trial Balance report at go-live |
| 2. Map | `mapping.csv` | Every QuickBooks field → Business Central table and field, the transform rule, and each decision with its reason |
| 3. Transform | `transform.py` | Builds the BC load: numbered chart of accounts with totals, customers, vendors, items, posting setup and the two opening-balance journals. Fills a BC Configuration Package export directly. Four self-checks run before anything is loaded |
| 4. Load | `docs/LOAD.md` | Configuration Package import into a new, empty BC company, then post the general and item journals |
| 5. Reconcile | `reconcile.py` | Reads BC through API v2.0 (or Excel exports) and compares against QuickBooks: record counts, A/R and A/P per customer and vendor, inventory per item, every account balance, and the clearing account. `--break` proves the checks catch a $1.00 error |

## Results

```
-- record counts
PASS  G/L accounts (incl. headings and totals): BC 108 vs QBO 108
PASS  Customers: BC 29 vs QBO 29
PASS  Vendors: BC 26 vs QBO 26
PASS  Items: BC 18 vs QBO 18
-- receivables and payables
PASS  A/R by customer total: BC 5281.52 vs QBO 5281.52
PASS  A/R by customer per record: 29/29 match
PASS  A/P by vendor total: BC 1602.67 vs QBO 1602.67
PASS  A/P by vendor per record: 26/26 match
-- inventory
PASS  Quantity on hand per record: 4/4 match
-- trial balance
PASS  Trial balance by account per record: 47/47 match
PASS  Clearing account 3900 is zero: BC 0.0 vs QBO 0.0
13/13 checks passed
```

## Decisions

| Decision | Choice | Why |
|---|---|---|
| Account numbers (QuickBooks has none) | Numbered by type: 1000s assets, 2000s liabilities, 3000s equity, 4000s income, 5000s COGS, 6000s expenses, 7000s other | BC requires a number on every account; grouping by type keeps reports readable |
| Sub-accounts (30 of 89, up to 3 levels) | Kept as Begin-Total / End-Total groups | BC reports subtotal the same way QuickBooks does |
| Sub-customers ("jobs", 3 with $837 open) | Separate BC customers, parent name in Name 2 | BC has no sub-customers; this keeps each job's balance and aging |
| History | Open items only: 20 of 31 invoices, one opening line each | Old paid invoices are rarely used after go-live; QuickBooks stays read-only for history |
| Vendor balances | One opening line per vendor with a balance | A/P ties per vendor without migrating each bill |
| Inventory | Quantity on hand × cost through the item journal | Item ledger and G/L agree from day one ($596.25) |
| Units of measure | EA for all items | QuickBooks items have no unit; BC requires one |
| Balancing | Every opening line balances against 3900 Opening Balance Clearing | If 3900 ends at exactly 0, nothing was dropped or doubled |

## Findings

What came up during the migration, and how it was handled:

1. **The account list isn't a trial balance.** QuickBooks' Account API returns balances for balance-sheet accounts only, so they don't net to zero (this year's profit is missing). Opening balances come from the Trial Balance report instead, which balances at $41,421.60.
2. **Parent accounts can carry their own balance.** 5 parents (for example Landscaping Services and Automobile) had amounts posted directly to them. BC can't post to a total line, so each got a "<Parent> - Other" posting account, the same way QuickBooks reports show it.
3. **Emails.** QuickBooks stores several emails comma-separated; BC only accepts semicolons. One vendor failed validation; the transform now converts all of them.
4. **External document numbers.** BC requires one on vendor invoices by default, and the whole journal refused to post without it. The transform now supplies one; for this load the setting was turned off during posting and back on after.
5. **The package's XML map.** BC's Configuration Package Excel file carries a hidden XML map. Saving it with a normal Excel library strips the map and BC rejects the file, so the transform writes the sheet XML directly.
6. **Stale validation errors.** Validating before the master data existed left errors on the journal lines, and Apply skips any line with a saved error. Apply the master data first, then validate again.
7. **Too many fields.** A package includes every field by default, and blank cells would overwrite defaults (on the journals, they can zero the amounts). The journals were cut to only the fields loaded. On a client project every table would be trimmed the same way.

## Run it

```
python extract.py Account   # also Customer, Vendor, Item, Invoice (needs QBO_REALM_ID and QBO_ACCESS_TOKEN)
python extract_tb.py 2026-11-23
python transform.py 2026-11-23 package.xlsx
# load in BC: docs/LOAD.md
python reconcile.py --api   # API setup: docs/RECONCILE.md
python reconcile.py --api --break
```

Source data is the QuickBooks Online sandbox company. Built with Python, the QuickBooks Online API and the Business Central API v2.0.
