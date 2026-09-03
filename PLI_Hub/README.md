# PLI Hub — Monarch export column break (2026-09-03)

## Symptom

PLI Hub's `CleanedData` sheet was only populating **Ship Date, Promised
Date, Job ID, and Week Start**. Every job on every row had a blank
Customer, Description, QTY, Location Date, Location, and Work Center —
100% of ~4,970 rows in the workbook as uploaded.

## Root cause

`Module4.ProcessAndDistributeAllData` reads the Monarch "Open Jobs By
Work Center" export and pulls each field from a **hardcoded column
number** (`keepCols.Add 6, "Customer"`, `keepCols.Add 10, "Description"`,
etc. — see the original code below).

Monarch's export changed its column layout. Comparing the two attached
exports:

| Field | Old column (`Full_Monarch_Pull_9.1.1.xls`) | New column (`New.OJBWC.9.3.xls`) |
|---|---|---|
| Ship | 1 | 1 (unchanged) |
| Prom | 3 | 3 (unchanged) |
| Job ID | 5 | 5 (unchanged) |
| Customer | 6 | 7 |
| Description | 10 | 11 |
| QTY | 16 | 17 |
| Location Date | 23 | 24 |
| Location | 25 | 26 |
| Work Center | 29 | 31 |

A new blank spacer column was inserted right after Job ID (shifting
Customer through Location by one column), a second spacer was added
between Location and Work Center (shifting Work Center by two columns
total), and a brand-new **"To Perso"** field was appended at the far
right (columns 32–33, populated on about 1 in 5 job rows — looks like a
"date sent to personalization" field). Every hardcoded column number
after Job ID was now pointing at the wrong field (or a blank spacer),
which is why those columns came back empty instead of erroring — nothing
in the old code checked that it was reading the field it thought it was.

This is not the first time Monarch's export has drifted — the module's
own comments reference prior migrations for sheet names and machine
columns — so a fixed column number is a bug waiting to happen again the
next time Monarch adds or removes a column.

## The fix

`PLI_Hub/Module4.bas` in this folder is a corrected copy of the Hub's
`Module4`. Instead of hardcoded column numbers, it now **finds each
column by its header text** on the export's header row (the row whose
first cell reads "Ship"), with a one-column search window and a
data-driven tie-breaker to handle two additional Crystal Reports quirks
confirmed against both attached exports:

- **Customer Name**'s data has been seen both directly under its header
  and one column to its right (aligned in the old export, offset by one
  in the new one) — not a fixed relationship.
- **Last Location** and **Gang Location**'s header labels consistently
  sit one column to the *right* of their actual data in both exports.
- **Work Center has no header label at all.** It's detected as the first
  column after Location that's populated on ~100% of real job rows —
  which reliably distinguishes it from both blank spacer columns (~0%
  filled) and the new "To Perso" column (~22% filled, and further right
  regardless).

If Monarch reshuffles columns again, this keeps working as long as the
header text itself doesn't change. If a header goes missing entirely,
**Update Data now fails with a clear message** naming the problem instead
of silently writing an empty `CleanedData` column — that silent failure
is exactly what happened this time and took a while to notice.

The "To Perso" column is not captured — the Hub wasn't reading it before,
and nothing asked for it. It's flagged in the code comments in case it's
wanted later.

I validated the column-detection algorithm in Python against both
attached export files (`Full_Monarch_Pull_9.1.1.xls` and
`New.OJBWC.9.3.xls`) — it recovers the exact expected column for every
field in both files, including the shifted ones. I could not run the
actual VBA in Excel from this environment, so **please do one test run
of "Update Data" and check `CleanedData`** after applying the fix, before
relying on it.

## How to apply it

1. Open `PLI_Hub.xlsm` in Excel.
2. Press **Alt+F11** to open the VBA editor.
3. In the Project Explorer, expand **VBAProject (PLI_Hub.xlsm) → Modules**
   and double-click **Module4**.
4. Click inside the code, press **Ctrl+A** then **Delete** to clear it.
5. Open `Module4.bas` from this folder, copy its entire contents, and
   paste them into the now-empty Module4.
6. Press **Ctrl+S** to save. Excel will ask to keep the Macro-Enabled
   Workbook format (.xlsm) — choose **Yes**.
7. Close the VBA editor and click the **Update Data** button (or run
   `ProcessAndDistributeAllData`), pointed at a folder containing the new
   Monarch export. Confirm `CleanedData` now has Customer, Description,
   QTY, Location Date, Location, and Work Center filled in again.

## Files

- `Module4.bas` — corrected module, ready to paste in as described above.
