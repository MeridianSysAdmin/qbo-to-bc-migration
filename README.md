# qbo-to-bc-migration (Demo 2)

Moves a small company from QuickBooks Online to Business Central: chart of accounts, customers, vendors, items and open invoices. A mapping sheet records every field decision, and a reconciliation report proves the numbers in Business Central match QuickBooks.

**Due Dec 11, 2026.** Recorded later with the other demos.

## Why this demo matters (plain version)

A small business outgrows QuickBooks and buys Business Central. Someone has to move the data across without losing a customer or a dollar. That "someone" is what a Business Central partner subcontracts out: the work is tedious, the mistakes are expensive, and their senior people don't want to do it. This demo shows you can do it and prove it.

## The plan

| Step | What | Done when | Target |
|---|---|---|---|
| 1. Extract | Pull Vendor and Item too (Customer, Invoice, Account are already in `data/`, copied from Demo 1) | 5 JSON files in `data/` | **Today** |
| 2. Mapping sheet | `mapping.csv`: every QBO field → BC table and field, with the rule. Answer every "YES" in the last column | No open decisions | **Today / weekend** |
| 3. Transform | `transform.py` turns the JSON into Excel files in the layout of a BC configuration package | 5 Excel files | Nov 30 – Dec 4 |
| 4. Load | Import through **Configuration Packages** in your BC trial (Search → "Configuration Packages"), then post the open-invoice journal | Data visible in BC | Dec 7–9 |
| 5. Reconcile | `reconcile.py` reads BC back through its API and compares to the QBO JSON: counts, open A/R total, A/R per customer, trial balance. Break test like Demo 1 | All PASS, break test caught | Dec 10–11 |
| 6. README findings | Same style as Demo 1 | Pushed to GitHub | Dec 11 |

## Today (about 1 hour)

1. **Refresh your QBO token** in the OAuth 2.0 Playground (it expires after about an hour), set `QBO_REALM_ID` and `QBO_ACCESS_TOKEN` like Demo 1, then from this folder:
   ```
   python extract.py Vendor
   python extract.py Item
   ```
2. **Open `mapping.csv`** in Excel. The Account, Customer and Invoice rows are drafted. Read every row and fix anything you disagree with.
3. **Fill in the Vendor and Item rows** now that you have the data (look at one record in each JSON file to see the field names).
4. **Answer the "YES" decisions** in the last column. Write your answer and one line of why. These decisions are the real consulting work; they're what you'll talk through on the recording.
5. `git init`, commit, and create the GitHub repo `qbo-to-bc-migration`.

## The decisions (and the usual answer)

| Decision | Usual answer | Why |
|---|---|---|
| Account numbers (QBO has none) | Number by type: 1000 assets, 2000 liabilities, 3000 equity, 4000 income, 5000 COGS, 6000+ expenses | BC requires a No. on every G/L account; numbering by type makes reports readable |
| Full history or open items? | **Open items only**, plus opening balances. Keep QuickBooks read-only for history | Moving years of history is slow, expensive and rarely used |
| Sub-customers (QBO "jobs") | Separate BC customers, or merge into the parent | BC has no sub-customer concept |
| Inactive accounts and customers | Skip, unless they have a balance | Less clutter; nothing is lost because QBO stays read-only |
| Items with no unit of measure | Default to `EA` (each) | BC requires a base unit of measure |

## Words you'll need

- **Configuration package:** BC's built-in import tool. You pick a table, export an Excel template, fill it, import it, and BC checks it before applying.
- **Opening balance:** the amount owed on the go-live date. You don't re-create each old sale; you post one line per open invoice so aging still works.
- **Cross-reference:** a column keeping the old QBO Id next to the new BC No., so every record can be traced back.
- **Trial balance:** the list of every account's balance. If QBO and BC trial balances match on go-live day, the migration is right.
