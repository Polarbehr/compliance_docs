# PLI Press — To Perso column + scheduling override (2026-09-03)

## What changed

Two things, both in `Module4`:

1. **"To Perso" is now a column on every machine tab**, placed as column A,
   immediately before Ship Date (which moves to column B). It shows the raw
   To Perso date pulled from Hub's CleanedData — blank when Hub doesn't
   have one for that job.
2. **The scheduler uses To Perso as a job's effective deadline whenever
   Hub has one.** Queue order, the "same ship date" gap-fill in Dicetrax's
   QTY scheduler, and the late/on-time colour check now all key off that
   date instead of Ship Date for that job. When a job has no To Perso
   date, everything works exactly as before, off Ship Date.

Ship Date's own column is untouched — it always shows the job's actual
ship date. Only the *scheduling math* substitutes To Perso in for it; the
two dates are never blended or overwritten on display.

## Where "effective schedule date" is used

All of it flows through one new function, `EffectiveScheduleDate` (a few
lines above `UpdatePressData`): To Perso if Hub has a valid date there,
otherwise Ship Date. Every place that used to read Ship Date directly for
*scheduling* purposes (as opposed to *display*) now calls this instead:

- The dated/undated split — a job with a To Perso date but no Ship Date is
  now schedulable (previously it would have fallen to the bottom of the
  tab as "No Ship Date").
- The sort that sets queue order (`CompareRowsShipThenJobID`).
- Dicetrax's QTY scheduler: the "same ship date" backfill rule now means
  "same effective date."
- The late / on-time / "ships today" row colour check.

`WritePressRow` (what actually gets written to each cell) was **not**
changed to use the effective date — it still writes the raw Ship Date and
the raw To Perso, each to its own column, so the tab always shows the
truth for both fields.

## Column layout

Every machine tab gained one column at the front; everything else shifted
right by one. Full current layout, left to right:

**A** To Perso · **B** Ship Date · **C** Job ID · **D** Customer Name ·
**E** Job Description · **F** Qty · **G** RFID · **H** Last Location ·
**I** Work Center · **J** Production Day · **K** Total Qty (merged) ·
**L** Remove (checkbox) · **M** RFID Type list · **N** RFID bulk-remove
tick · (hidden) **O** row-format state marker.

This is a pure shift — I did not rename or reorder anything else. Every
`COL_OUT_*` constant in the module was renumbered to match (plus one new
one, `COL_OUT_LOCATION`, naming what was previously a bare literal `7` for
"Last Location"), and I went through the file for every place a column
letter was hardcoded — including inside the operator-facing text on the
Onboarding and Rules tabs (e.g. "tick the REMOVE box in column K" →
"column L") — since those sheets are rebuilt from this code on every run
and are supposed to always match what the code actually does.

## Hub compatibility

This reads Hub's CleanedData column L (To Perso) the same guarded way
Hub's own fix reads Monarch's columns: `UBound(rawData, 2) >= COL_TOPERSO`
before touching it. If Hub hasn't been updated yet (still shipping 11
columns), Press runs exactly as it did before — To Perso just shows blank
everywhere and scheduling falls back to Ship Date for every job. Nothing
breaks either way; you get the new behavior once Hub's `Module4.bas` fix
is applied and re-run.

## What I could not test

Same caveat as the Hub fix: I don't have Excel available in this
environment to actually run the macro. I traced every column reference in
the file by hand — the code itself (via the `COL_OUT_*`/`COL_*`
constants, which the vast majority of the file already routed through)
and the handful of places that hardcoded a column letter in a text
string — and re-verified the Sub/Function/If/For counts balance against
the original file. But **please do one Update Data run and eyeball a
machine tab** (column order, a job or two with a real To Perso date
scheduled ahead of its Ship Date, the Rules/Onboarding text) before
relying on it for real scheduling.

## How to apply it

Same as the Hub fix:

1. Open `PLI_Press.xlsm` in Excel.
2. **Alt+F11** → Project Explorer → **VBAProject (PLI_Press.xlsm) →
   Modules → Module4**.
3. Click inside the code, **Ctrl+A** then **Delete**.
4. Paste in the entire contents of `Module4.bas` from this folder.
5. **Ctrl+S** → keep Macro-Enabled Workbook (.xlsm) format.
6. Click **Update Data** (needs Hub's own fix applied and re-run first, so
   Hub is actually shipping the To Perso column) and check a machine tab.

## Files

- `Module4.bas` — corrected module, ready to paste in as described above.
