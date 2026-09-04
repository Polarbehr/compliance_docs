# PLI Perso — column layout changes (2026-09-03)

## What changed, most recent first

**Cards removed, Week Start hidden.** Cards (the checkbox column) is
fully removed — not needed any more, per the user. Rdy takes its old
place; the K-purge column, the capacity table, and the Update Data
button's parked position each shift one column left to close the gap.
Week Start is **hidden**, not removed — it's not needed for viewing, but
the odd-week-bold conditional format still reads it (`ISODD(WEEKNUM($J2,
21))`), so the column stays and `Columns(10).Hidden = True` keeps it out
of sight.

**Ship Date and To Perso swapped back.** Ship Date is column A again, To
Perso is B.

**To Perso added.** The **Tracie** sheet got a new **To Perso** column,
placed as column A, immediately before Ship Date. It shows the raw To
Perso date pulled from Hub's CleanedData — blank when Hub doesn't have
one for that job.

No scheduling, sorting, filtering, or capacity-math behavior changed
across any of this. These are straight column layout changes, not a
redesign — unlike PLI Press, nothing here reads To Perso as a scheduling
override.

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

**A** Ship Date · **B** To Perso · **C** Job ID · **D** Customer Name ·
**E** Description · **F** QTY · **G** Location · **H** Machine ·
**I** Rdy · **J** Week Start (hidden) · (buffer) **K** ·
**L** onward: the capacity table.

Column history, most recent first:

- **2026-09-03 (later same day again):** Cards removed entirely (not
  needed any more, per the user) and Week Start hidden (not needed for
  viewing, but still read by the odd-week-bold conditional format).
  Rdy moved from J to I; the buffer/capacity table each shifted one
  column left to close the gap Cards left behind — net result, the
  layout right of Machine is now exactly what it was before To Perso
  was ever added.
- **2026-09-03 (later same day):** Ship Date and To Perso swapped back —
  Ship Date is A, To Perso is B. The spacer-row check in
  `ApplyTracieConditionalFormatting` moved back to `$A2`.
- **2026-09-03:** To Perso added as column A, before Ship Date (B).

## One-time transition notes (apply on the first run after each change above)

Rdy (and Cards, while it existed) tick state survives a rebuild by reading
the sheet's own *current* (pre-rebuild) layout at the start of the build
and restoring it after. On the very first Update Data run after a column
layout change, that read happens against last run's OLD-layout cells but
at the NEW column positions — so it won't find a match, and every job
comes back with its checkbox(es) **unticked once**. Every run after that
is fully self-consistent again. This is an unavoidable, one-time side
effect of any column-shifting change to a sheet that snapshots and
restores its own state like this — not a bug, and not something a
"don't change anything else" scope could avoid without leaving the
tracking broken permanently instead. It already happened once for the To
Perso addition; it will happen again, once, for the Cards removal.

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

0. If you're carrying forward from an earlier paste of this file, you can
   skip straight to step 2 — the file below already includes every
   change described above; there's no need to apply them one at a time.
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
