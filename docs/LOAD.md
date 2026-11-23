# Loading into Business Central (step 4)

This is done by hand in your BC trial. It takes about 45-60 minutes the first time. Write down every error you hit; those become part of the "findings" on the recording.

## A. New, empty company (5 min)

1. Search (Alt+Q) → **Companies** → **New**. Choose the option with **no data** (in the wizard it's "Production - No data", or under "Advanced"). Name it `Craigs Landscaping MIG`.
2. Switch to it: Settings (gear icon) → **My Settings** → Company.

The sample company (CRONUS) already has its own chart of accounts and would clash with ours, which is why you need an empty one.

## B. Small setup before importing (10 min)

| Do | Where (Alt+Q) | Why |
|---|---|---|
| Open the page once, then close it | **General Journals** | BC creates the `GENERAL` template and `DEFAULT` batch. Check that the names match. If they don't, change `GEN_TEMPLATE` / `GEN_BATCH` at the top of transform.py |
| Open the page once, then close it | **Item Journals** | Creates `ITEM` / `DEFAULT` |
| **Create Year** starting 01/01/2026, 12 periods | **Accounting Periods** | Posting needs an open fiscal year |
| Turn on **Automatic Cost Posting** | **Inventory Setup** | So the inventory opening entry posts to G/L account 1030 |

## C. Configuration package (20 min)

1. Alt+Q → **Configuration Packages** → **New**. Code `QBO-MIG`, name "QuickBooks migration".
2. On the **Tables** lines, add these table IDs **in this order**: `204, 15, 251, 94, 92, 93, 252, 5813, 18, 23, 27, 81, 83`.
3. For each table, choose **Fields** → **Clear Included**. Then tick only the fields that are column headers on that table's sheet in `out/bc_load.xlsx`. If you leave every field included, blank cells will wipe out BC's default values.
4. **Export to Excel** → save the file as `package.xlsx` in this folder.
5. Run:
   ```
   python transform.py 2026-11-23 package.xlsx
   ```
   This writes `out/package_filled.xlsx`. Read what it prints: any `!` line names a field the package is missing.
6. Back in the package: **Import from Excel** → choose `out/package_filled.xlsx`.
7. **Validate Package**. Fix every error, either in BC or by changing transform.py and importing again.
8. **Apply Package**.

## D. Post and check (10 min)

1. **Chart of Accounts** → action **Indent Chart of Accounts**. This sets indentation and fills in the totals from the Begin-Total and End-Total lines.
2. **General Journals**, batch DEFAULT → you should see 67 lines. Choose **Post**.
3. **Item Journals**, batch DEFAULT → you should see 4 lines. Choose **Post**.
4. Quick look: **Trial Balance** report as of 11/23/2026. Account **3900 should be 0.00**, A/R (1020) should be **5,281.52**, and A/P (2000) should be **-1,602.67**.

Then step 5: `reconcile.py` proves all of this automatically through the BC API.
