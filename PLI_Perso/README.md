# PLI Perso — To Perso column added before Ship Date (2026-09-03)

## What changed

Exactly one thing, as asked: the **Tracie** sheet gets a new **To Perso**
column, placed as column A, immediately before Ship Date (which moves to
column B). It shows the raw To Perso date pulled from Hub's CleanedData —
blank when Hub doesn't have one for that job.

No scheduling, sorting, filtering, or capacity-math behavior changed. This
is a straight column insertion, not a redesign — unlike PLI Press, nothing
here reads To Perso as a scheduling override.

## What "just add the column" actually touched

Tracie has no named column constants (unlike Hub/Press) — every column is
a bare literal number scattered through `Module1`. Inserting a column at
the very front means everything after it shifts right by exactly one, so
"just add the column" still meant finding and shifting every one of those
literals so the sheet keeps working as a whole, including two pieces that
aren't obviously "the data columns":

- **The K-purge column** (stale conditional formatting the original
  workbook left behind) is now column L.
- **The capacity table** — the "Capacity - QTY by Week Start x Machine"
  spill grid to the right of the job list — now starts at column M
  instead of L, and its formulas (which reference the data columns
  directly, e.g. `$G$2:$G$...` for Machine) were repointed to the columns
  Machine/Week Start/QTY actually moved to (H/K/F). The Update Data
  button's parked position (it sits "above the capacity table") moved
  with it.

Nothing outside that shifted: the AZ1/AZ2 checkbox-format stash and the
column-50 safety clamp are arbitrary far-right anchors, unrelated to
where the data starts, and were left alone.

## Current column layout

**A** To Perso · **B** Ship Date · **C** Job ID · **D** Customer Name ·
**E** Description · **F** QTY · **G** Location · **H** Machine ·
**I** Cards · **J** Rdy · **K** Week Start · (buffer) **L** ·
**M** onward: the capacity table.

## One-time transition note

Cards/Rdy tick state survives a rebuild by reading the sheet's own
*current* (pre-rebuild) layout at the start of the build and restoring it
after. On the very first Update Data run after this update, that read
happens against last run's OLD-layout cells but at the NEW column
positions — so it won't find a match, and every job comes back with
Cards/Rdy **unticked once**. Every run after that is fully self-consistent
again. This is an unavoidable, one-time side effect of any column-adding
change to a sheet that snapshots and restores its own state like this —
not a bug, and not something a "don't change anything else" scope could
avoid without leaving Cards/Rdy tracking broken permanently instead.

## What I could not test

I don't have Excel available in this environment. I traced every literal
column reference in the file by hand (there's no `COL_OUT_*` constant
scheme here to lean on, so this took a careful full pass through
`BuildTracieTab`, `ApplyMachineStyles`, `ApplyTracieConditionalFormatting`,
`CapacityLastCol`, `EnsureCapacityTable`, `FormatCapacityTable`,
`EnsureCheckboxFormats`, `CompactPersoSheet`, and the Update Data button's
placement code) and re-verified the Sub/Function/If/For counts balance
exactly against the original file. **Please do one Update Data run and
check the Tracie tab** — column order, a spot-check of a couple of dates,
and that the capacity table still spills correctly at its new position —
before relying on it.

## How to apply it

Same as Hub and Press:

1. Open `PLI_Perso.xlsm` in Excel.
2. **Alt+F11** → Project Explorer → **VBAProject (PLI_Perso.xlsm) →
   Modules → Module1** (this workbook's module is named Module1, not
   Module4).
3. Click inside the code, **Ctrl+A** then **Delete**.
4. Paste in the entire contents of `Module1.bas` from this folder.
5. **Ctrl+S** → keep Macro-Enabled Workbook (.xlsm) format.
6. Click **Update Data** and check the Tracie tab.

## Files

- `Module1.bas` — corrected module, ready to paste in as described above.
