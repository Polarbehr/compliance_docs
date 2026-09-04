Option Explicit

' =============================================================================
' PLI PERSO WORKBOOK -- Module1 (v3.15, 2026-09-03)
'
' The version is ALSO held in the MODULE_VERSION constant below and printed at
' the end of every Update Data run, so the build a workbook is actually running
' can be read without opening the VBA editor. The banner said "v1" through the
' whole of the v2 change set below -- the workbook on the share was running v2
' code under a v1 label, which is exactly the confusion Press v43 added its
' version banner to end.
'
'   ---------------------------------------------------------------------
'   REVISION HISTORY
'   ---------------------------------------------------------------------
'   v3.15  2026-09-03
'       * Spacer-row checkboxes FIXED, this time from evidence rather than
'         from a guess. DiagnoseTracieCheckboxes on the live sheet reported
'         for spacer cell I5: value empty, NumberFmt General, Interior
'         black, Validation none, no shapes on the row (1 on the whole
'         sheet -- the Update Data button), and CellControl: PRESENT.
'         So the box is a cell-native CellControl that survives BOTH
'         ClearFormats (v3.11) and a full Clear (v3.14), and is not a
'         shape (v3.14's other branch, now removed as dead code). It is
'         removed through the CellControl API itself -- see the new
'         RemoveCellCheckbox, which is late-bound and verifies the control
'         is actually gone rather than assuming the call worked, and which
'         reports into gSetupWarnings if it ever cannot.
'       * Same bug fixed in EnsureCheckboxFormats' below-the-data trim,
'         which has been calling ClearFormats since the feature shipped and
'         therefore never actually removed a checkbox in its life. Live
'         boxes have been accumulating below the data footprint that whole
'         time; they get cleared properly now.
'   v3.14  2026-09-03
'       * AutoFilter removed from the Tracie header row (user request:
'         the dropdown arrows were not wanted). Nothing in this module
'         reads the filter, and filtering breaks the machine grouping the
'         sheet is laid out to show. Turned OFF explicitly, not merely
'         left unset, so a filter from an earlier build does not keep its
'         arrows forever.
'       * THIRD go at the spacer-row checkboxes, after v3.11 (ClearFormats)
'         and v3.13 (clone a clean format over it) both failed against the
'         live workbook. Both were guesses at which mechanism Excel is
'         using; this one stops guessing and covers every candidate at
'         once -- full Clear, NumberFormat reset, black fill restored, AND
'         removal of any Form Control / ActiveX checkbox SHAPE anchored on
'         a spacer row. Shapes are the one candidate neither earlier fix
'         could ever have touched, and they fit the symptom exactly: a
'         shape is positioned by geometry, so one parked over a job row on
'         a previous run sits over a spacer row after a rebuild changes
'         the row count.
'       * NEW, manual: DiagnoseTracieCheckboxes (Alt+F8). Read-only. If a
'         box still survives, this reports what is actually on a spacer
'         row's Rdy cell -- value, number format, validation, CellControl,
'         and any anchored shapes -- so the next fix is aimed instead of
'         guessed. Two wrong guesses is enough.
'   v3.13  2026-09-03
'       * v3.11's spacer-row checkbox fix did not actually work (confirmed
'         by the user) -- ClearFormats does not remove whatever
'         PasteSpecial xlPasteFormats propagates for this workbook's
'         checkbox, unlike Press's CellControl.SetCheckbox-created ones,
'         which ClearFormats does remove. Fixed properly: rather than try
'         to strip the checkbox back off, each spacer row's Rdy cell now
'         gets its format overwritten by cloning from that SAME row's
'         Machine cell (column 8), which is already correctly plain black
'         with no checkbox -- the same paste-cloning technique this
'         workbook already uses everywhere else, applied to a source that
'         is guaranteed checkbox-free instead of to a removal API that
'         did not actually reach it.
'   v3.12  2026-09-03
'       * Fixed: stray dates the user spotted sitting in K1:K5. The
'         column-K purge block (see BuildTracieTab) only ever cleared
'         FORMATTING there (conditional-format rules, fills, borders),
'         never CONTENTS, even though K is documented as holding no data
'         at all in this design -- nothing had ever cleared literal
'         values left over from an earlier layout. Added
'         Columns(11).ClearContents. Safe across the whole column, every
'         row, unlike the capacity zone's own row-1-5 strip starting at
'         L, which stays deliberately untouched as user territory.
'   v3.11  2026-09-03
'       * First attempt at the checkbox fix in v3.13 -- collected each
'         spacer row during the write loop and ClearFormats'd its Rdy
'         cell afterward. Did not remove the checkbox; see v3.13.
'   v3.10  2026-09-03
'       * Cards removed (not needed any more, per the user). Snapshot/
'         restore, checkbox-format cloning, headers, and the hide-blank-
'         text CF rule all simplified from a Cards+Rdy pair down to Rdy
'         alone; every column from Rdy onward (Week Start, the K-purge
'         column, the capacity table, the Update Data button's parked
'         position) shifted one column LEFT to close the gap.
'       * Week Start no longer shown -- not removed, HIDDEN
'         (ws.Columns(...).Hidden = True), because the odd-week-bold
'         conditional format still reads it. Same one-time transition
'         cost as before applies to the Rdy tick restore (see the v3.8
'         note): the very first run after this update reads last run's
'         Cards-column leftovers at Rdy's new position and won't match,
'         so every job's Rdy comes back unticked once.
'   v3.9  2026-09-03
'       * Ship Date and To Perso swapped back: Ship Date is column A again,
'         To Perso is B. Only those two columns' positions changed --
'         everything from Job ID (C) onward, the K-purge column, and the
'         capacity table stayed exactly where v3.8 put them. The spacer-row
'         detection in ApplyTracieConditionalFormatting ($A2<>"" /
'         $A2="") moved from $B2 back to $A2, since Ship Date -- reliably
'         filled, unlike To Perso -- is column A again.
'   v3.8  2026-09-03
'       * Hub started shipping a 12th CleanedData column, To Perso (column
'         L). Tracie gains it as column A, before Ship Date -- every other
'         data column, the K purge column, and the capacity table (still
'         starting one column past the data, "Capacity - QTY by Week Start
'         x Machine") all shift right by exactly one to make room. Nothing
'         about the scheduling, sorting, filtering, or capacity math
'         changed -- this is a straight column insertion, so every literal
'         column reference in BuildTracieTab, ApplyMachineStyles,
'         ApplyTracieConditionalFormatting, CapacityLastCol,
'         EnsureCapacityTable (values, formulas AND the L6 label),
'         FormatCapacityTable, EnsureCheckboxFormats, CompactPersoSheet and
'         the Update Data button's parked position moved with it. The AZ1/
'         AZ2 checkbox-format stash and the column-50 (AX) safety clamp
'         did NOT move -- both exist purely to stay clear of each other and
'         of the far edge, independent of where the data starts.
'       * One-time transition note: Cards/Rdy tick state is restored by
'         reading the sheet's OWN current (pre-rebuild) layout, so the
'         very first Update Data run after this update reads last run's
'         OLD-layout cells at the NEW column positions and will not find a
'         match -- every job comes back with Cards/Rdy unticked once. Any
'         run after that is fully self-consistent again.
'   v3  2026-08-31
'       * DiscoverWorkCenters no longer WRITES over the rows below the roster.
'         It wrote at RosterLastRow + 1, which is the blank separator row.
'         One discovered machine consumed that separator, after which the
'         roster reader ran on into the "Location Exclude" label and its
'         instruction paragraph and registered them as machines; a second
'         discovered machine OVERWROTE the "Location Exclude" label itself,
'         at which point ReadLabeledList found no anchor, LocationIsExcluded
'         returned False for every row, and the whole Shipping exclusion
'         silently stopped working. It now INSERTS a row, so every section
'         below shifts down intact. This is the same "never lastRow + N on
'         the Preferences sheet" rule Press already carries.
'       * The three Preferences block readers (ReadLabeledList,
'         ReadPersoRoster, RosterLastRow) stopped at the FIRST BLANK ROW.
'         They now skip blanks and stop at the next SECTION LABEL, so a gap
'         left by hand while adding an entry no longer silently truncates the
'         list. Press hardened its equivalent reader after the column-A,
'         first-blank version came one Update Data away from cutting Dicetrax
'         from 665 jobs to 119; Perso never got the same treatment.
'       * Two hard-coded ceilings removed. EnsureCheckboxFormats cleared H:I
'         to a fixed row 2000; the capacity zone was number-formatted to a
'         fixed L8:L500 / M8:AD500. Both are the same silent-ceiling class the
'         v2 note below congratulates itself for removing from the capacity
'         FORMULAS -- the formatting three lines under them still had it.
'         Extents are now derived from the used range and the roster.
'       * Capacity number formats moved out of EnsureCapacityTable into
'         FormatCapacityTable, which is the only place that knows how far the
'         spills actually reached. Formatting a fixed L8:AD500 block was not
'         just a ceiling, it INSTANTIATED ~9,400 empty cells in the file every
'         run; sizing it to the job count instead would have kept the bloat,
'         because there are only ever as many week rows as there are distinct
'         weeks. The reset window returns the zone to General first, so a
'         shrinking table now cleans up after itself.
'       * Update Data button gained a SECOND CAPTION LINE showing the pull
'         stamp, and a state colour: amber when the Hub data behind the pull
'         is older than STALE_HUB_HOURS, red after a failure, the user's own
'         colour otherwise. The caption repeats the state with a leading "!",
'         so it still reads without colour vision.
'         The button stays user-owned. Line 1 is read back and re-used
'         verbatim, so renaming it survives; position and shape are never
'         touched; height grows once, and only while the button is still at
'         exactly its created 160x34. The user's own fill is remembered in the
'         shape's AlternativeText and handed back when the warning clears --
'         and if the fill on entry is not the one this code last set, that is
'         taken as a deliberate recolour and becomes the new base.
'         Caption colour is picked from the fill by WCAG relative luminance,
'         NOT a brightness average: on the amber state the naive test picks
'         white (contrast 3.49) where the real formula picks near-black
'         (6.02). Every state clears AA at the 9pt line size.
'       * NEW, manual: CompactPersoSheet (Alt+F8). Reclaims the empty rows
'         earlier builds left below the data -- Tracie's sheet1.xml was 432 KB
'         holding 212 rows of jobs. One-off; not wired into Update Data,
'         because a routine refresh should never delete rows on its own.
'       * Version banner + MODULE_VERSION, printed in the finish message.
'
' (Module1, not Module4: this workbook's VBA project rejects newly INSERTED
'  modules under the LibreOffice build pipeline -- any insertByName breaks
'  the xlsm store, so the build REPLACES the original Module1 in place.
'  Future rebuilds must use replaceByName('Module1', ...) too.)
' Satellite workbook: pulls CleanedData (values only) from the Hub workbook
' on demand via the Update Data button -- EXACTLY the same Hub contract as
' the Press workbook (same Hub Workbook Path preference, same Hub Status
' gate: REFRESHING/ERROR block the pull, READY proceeds) -- then rebuilds
' the "Tracie" tab: Personalization/Thermal jobs grouped by machine with a
' short display name, a Rdy checkmark column, a hidden Monday Week Start
' column (kept for the odd-week-bold formatting rule, not for viewing --
' see BuildTracieTab), and per-machine row colors driven by the
' Preferences Machine Roster. Cards was removed 2026-09.03 (not needed
' any more, per the user).
'
' This module REPLACES the workbook's original Power Query chain
' (Monarch Import -> ConvertRawEmail -> Tracie query). The query rules it
' reproduces, confirmed against the extracted M source (Section1.m):
'   - Columns kept: Ship Date, Job ID, Customer Name, Description, QTY,
'     Location, Machine (short name), plus added Cards / Rdy / Week Start.
'   - Location rule: a job is EXCLUDED when its Location contains any
'     "Location Exclude" phrase (seed: "Shipping") UNLESS the whole
'     Location exactly equals a "Location Exclude Exceptions" row (seeds:
'     "Partial in Shipping", "Staged Receiving") -- the M code kept those
'     two exact values and dropped every other "shipping" location.
'   - Machine short names: Personalization 12 -> RFID 12, 13 -> RFID 13,
'     4 -> DOD 4, 5 -> DOD 5, 6 -> DOD 6, 7 -> DOD 7,
'     Thermal Printer 1 -> Thermal 1 (all seeded in the Machine Roster;
'     add rows there for new machines -- no VBA change needed).
'   - Sort: Machine asc, then Week Start asc, then Ship Date asc (undated
'     rows first within their machine, matching Power Query's nulls-first
'     ascending order), then numeric-aware Job ID as a stable tiebreak.
'   - One blank spacer row (filled black) between machine groups.
'
' Improvements over the query version, per the user's 2026-08.12 request:
'   - Rdy values now SURVIVE a refresh (the query wiped them every run):
'     preserved across updates keyed on Job ID + Machine. (Cards got the
'     same treatment until it was removed 2026-09.03.)
'   - Conditional formatting is VBA-managed: rebuilt on every run from the
'     Machine Roster's Row Color cells (paint a roster color cell any fill
'     you like and the machine's rows follow on the next update). Also
'     rebuilt: the odd-week BOLD banding rule and the hide-checkbox-text-
'     on-blank-rows rule from the original workbook.
'   - The Update Data button is USER-OWNED after creation: move it or
'     recolor it freely; updates only re-wire its macro, never its look.
'
' First-run bootstrap: same limitation as Hub/Press -- the Update Data
' button does not exist until the first run creates it, so the very first
' run must be started via Alt+F8 -> UpdatePersoData -> Run.
' =============================================================================

' -----------------------------------------------------------------------------
' Column positions within Hub's CleanedData sheet (stable, unchanged since
' the Hub split -- identical to the Press module's constants): A Ship Date,
' B Promised Date, C Job ID, D Customer, E Description, F QTY,
' G Location Date, H Location, I Work Center, J Week Start, K RFID Type,
' L To Perso.
'
' 2026-08-31: K/L previously documented here as "Effective QTY / Job Hours".
' That was the pre-split Dashboard layout. Hub v51 writes K = RFID Type and
' has no column L. Nothing in this module reads past column 9, so the error
' was inert -- but it was wrong for anyone reaching for K next.
'
' 2026-09.03: Hub added column L, To Perso. This module now reads it (see
' COL_TOPERSO) purely to copy it to Tracie's own new first column -- no
' other behavior depends on it.
' -----------------------------------------------------------------------------
Private Const COL_SHIP As Long = 1
Private Const COL_PROMISED As Long = 2
Private Const COL_JOBID As Long = 3
Private Const COL_CUSTOMER As Long = 4
Private Const COL_DESC As Long = 5
Private Const COL_QTY As Long = 6
Private Const COL_LOCDATE As Long = 7
Private Const COL_LOCATION As Long = 8
Private Const COL_WORKCENTER As Long = 9
Private Const COL_TOPERSO As Long = 12

' Output tab name -- the original query landed on "Tracie"; kept as-is.
Private Const OUT_SHEET As String = "Tracie"

' -----------------------------------------------------------------------------
' v2 (2026-08.20)
'   * Machine Roster gains a SHOW tick box (column D) -- untick a machine and
'     its section is not built. Form Controls, snapshot/restore, so a rebuild
'     can never wipe a tick (the mistake that reset every box in Press v36).
'   * Work Center and Location matching moved from CONTAINS to EXACT WITH
'     WILDCARDS. Contains made the lists impossible to control: it is the same
'     defect that let "TI" match the "ti" inside "To Imposition" in Press.
'   * New "Work Center Discovery" section: work centers in Hub matching one of
'     its patterns but absent from the roster are APPENDED, unticked, and named
'     in the completion message. Seeded Personalization* / Thermal*, so Casing,
'     Punch, QC and the other 34 never appear.
'   * Column C of the roster is now a full STYLE swatch, not just a fill:
'     fill, font, size, bold, italic and colour all follow it.
'   * Scripting.Dictionary REPLACED with a native Collection. It is a
'     Windows-only COM object; with it in the build path LibreOffice could not
'     execute BuildTracieTab at all ("323 invalid format"), so none of this
'     could be tested. A Collection is intrinsic, behaves identically here, and
'     makes the whole build verifiable.
'   * Perso Status now states TWO facts -- when Perso pulled, and how old the
'     Hub data was -- instead of one ambiguous "Updated At". Warns when the Hub
'     report is more than STALE_HUB_HOURS old.
'   * Capacity grid formulas are sized to the ACTUAL last row each run instead
'     of a hard-coded 5000, which was a silent ceiling.
'   * No silent failures: setup problems are recorded and reported.
' -----------------------------------------------------------------------------

' Machine Roster columns.
Private Const RC_MATCH As Long = 1
Private Const RC_DISPLAY As Long = 2
Private Const RC_STYLE As Long = 3
Private Const RC_SHOW As Long = 4

' Tick boxes owned by this module. Anything without the prefix is the user's
' (the Update Data button included) and is never touched.
Private Const ROSTER_CHK_PREFIX As String = "MRchk_"

' How old the Hub report may be before a successful pull says so.
Private Const STALE_HUB_HOURS As Double = 12#

' Printed at the end of a run. Keep in step with the banner at the top.
'
' 2026-08-31: this shipped as a bare "v3" three times in one afternoon while
' the module was still growing -- readers/discovery, then CompactPersoSheet,
' then the button -- under one filename each time. Three different builds,
' one version string: precisely the confusion a version banner exists to
' prevent, and it cost a round of "the macro isn't in Alt+F8". Point releases
' from here, so the finish message identifies the BUILD, not just the series.
'   v3.1  block readers, discovery insert, ceilings, K/L column comment
'   v3.2  + CompactPersoSheet
'   v3.3  + two-line button with contrast-picked caption colour; the
'         L1:AD500 migration narrowed off the user's top strip
'   v3.8  + To Perso column inserted before Ship Date; every column
'         reference after it shifted right by one
'   v3.9  Ship Date and To Perso swapped back (A/B), nothing else moved
'   v3.10 Cards removed; Week Start hidden, not removed; capacity table
'         and everything after Rdy shifted one column left
'   v3.11 First (unsuccessful) attempt at the spacer-row checkbox fix
'   v3.12 Fixed: stray dates in K1:K5 -- column K purge now clears content
'   v3.13 Second (also unsuccessful) attempt at the spacer-row checkbox
'   v3.14 AutoFilter arrows removed; third spacer-row checkbox attempt,
'         covering cell AND shape mechanisms; + DiagnoseTracieCheckboxes
'   v3.15 Spacer-row checkboxes fixed for real, from the diagnostic's
'         evidence: removed via the CellControl API, the only thing that
'         touches them. Same fix applied below the data footprint.
Private Const MODULE_VERSION As String = "v3.15"

' -----------------------------------------------------------------------------
' UPDATE DATA BUTTON -- second caption line and state colour (v3).
'
' The button stays USER-OWNED: position, size, shape and the line-1 label are
' yours, and a run never touches them. What a run DOES manage is line 2 (the
' pull stamp) and, in a warning state, the fill.
'
' The user's own fill is remembered in the shape's AlternativeText so a warning
' colour can be handed back afterwards. If the fill on entry is not the one this
' code last set, it is taken as a deliberate recolour and becomes the new base.
' -----------------------------------------------------------------------------
Private Const BTN_NAME As String = "PersoUpdateDataButton"
Private Const BTN_STATE_OK As String = "OK"
Private Const BTN_STATE_STALE As String = "STALE"
Private Const BTN_STATE_ERROR As String = "ERROR"

' Created size. A button still at EXACTLY this size is untouched and may be
' grown to fit two lines; anything else the user has sized themselves.
Private Const BTN_DEF_W As Single = 160
Private Const BTN_DEF_H As Single = 34
Private Const BTN_TWO_LINE_H As Single = 46

' VBA cannot call RGB() in a Const, so these are the literal Long values.
' RGB(r,g,b) = r + g*256 + b*65536 -- the triple is given so the number can be
' checked rather than trusted.
Private Const BTN_FILL_BASE As Long = 7884319      ' RGB(31, 78, 120)  default blue
Private Const BTN_FILL_STALE As Long = 31932       ' RGB(188, 124, 0)  amber
Private Const BTN_FILL_ERROR As Long = 1315990     ' RGB(150, 20, 20)  red
Private Const BTN_TEXT_LIGHT As Long = 16777215    ' RGB(255, 255, 255)
Private Const BTN_TEXT_DARK As Long = 1513239      ' RGB(23, 23, 23)
Private Const BTN_TAG As String = "PersoBtn|"

' How many consecutive blank rows a Preferences block reader will step over
' before it accepts that the block has ended. Blocks normally end at the next
' SECTION LABEL (see IsPersoSectionLabel); this is only the backstop for the
' LAST block on the sheet, which has no label after it. Generous on purpose --
' a gap left while editing must never truncate a list -- but bounded, so a
' stray value far down column A is never read as an entry.
Private Const MAX_PREF_BLANK_RUN As Long = 25

' One-time migration marker: roster match values carry wildcards and the
' Location Exclude list has been rewritten for Like-matching.
Private Const PERSO_MARK_COL As Long = 26
Private Const PERSO_MARK As String = "PERSOV2"

' Non-fatal setup problems, reported at the end of a run. Never swallowed in
' silence -- a feature that fails invisibly is worse than one that errors.
Private gSetupWarnings As String

' The Hub report timestamp for the run in progress.
Private gHubDataStamp As String

' =============================================================================
' ENTRY POINT -- wired to the "Update Data" button on Preferences.
' =============================================================================
Public Sub UpdatePersoData()
    Dim t0 As Single: t0 = Timer

    ' WRONG-WORKBOOK GUARD (v3.6). Refuse to run anywhere that is not Perso.
    '
    ' 2026-08-31: this module was imported into PLI Hub.xlsm by mistake and
    ' sat there for an afternoon. Nothing objected -- the four procedure names
    ' it shares with Hub's Module4 are Private in both, so there is no
    ' ambiguity error to notice. But EnsurePreferencesSheet, the FIRST setup
    ' step below, does this:
    '
    '     If A1 <> "Perso (Hub Satellite)" Then ws.Cells.Clear
    '
    ' Hub's Preferences A1 is empty. One Alt+F8 on UpdatePersoData in the Hub
    ' workbook would therefore have wiped Hub's Preferences sheet outright --
    ' the Monarch Report Folder path, the Hub Status flag every satellite
    ' checks before it pulls, and the whole RFID technology token list. A
    ' module that destroys a sheet when dropped in the wrong file is not
    ' allowed to rely on nobody making that mistake.
    If Not IsPersoWorkbook() Then
        MsgBox "This is Module1 from the PLI PERSO workbook, and this is not " & _
               "the Perso workbook." & vbCrLf & vbCrLf & _
               "Workbook : " & ThisWorkbook.name & vbCrLf & _
               "Expected : a '" & OUT_SHEET & "' sheet, or Preferences!A1 = " & _
               """Perso (Hub Satellite)""" & vbCrLf & vbCrLf & _
               "Nothing has been changed. Remove Module1 from this workbook " & _
               "and import it into PLI Perso.xlsm instead.", _
               vbCritical, "Wrong workbook"
        Exit Sub
    End If

    ' The setup phase used to run with NO error handler armed. Anything that
    ' threw in here produced a raw VBA dialog and killed the run before the
    ' status cell was even written, leaving the workbook looking untouched
    ' with nothing recorded anywhere.
    gSetupWarnings = ""
    gHubDataStamp = ""
    Dim setupStep As String
    On Error GoTo SetupFailed
    setupStep = "EnsurePreferencesSheet": EnsurePreferencesSheet
    setupStep = "EnsureUpdateButton": EnsureUpdateButton
    On Error GoTo 0

    Dim hubPath As String
    hubPath = Trim(CStr(GetPrefValue("Hub Workbook Path")))
    If hubPath = "" Then
        MsgBox "Set 'Hub Workbook Path' on Preferences to the PLI Hub.xlsm file first.", vbExclamation, "Update Data"
        ShowPreferencesWhenDone
        Exit Sub
    End If

    SetPersoStatus "Not Ready", "pull in progress"

    On Error GoTo PullFailed

    Dim rawData As Variant
    Dim hubStatusWord As String, hubStatusDetail As String
    If Not PullCleanedDataFromHub(hubPath, rawData, hubStatusWord, hubStatusDetail) Then
        GoTo CleanExit  ' PullCleanedDataFromHub already showed the reason
    End If

    ' Discovery runs BEFORE the roster is read for the build, so a machine
    ' found this run is already in the list (unticked) when the tab is built.
    Dim discovered As String
    Dim roster0 As Variant
    roster0 = ReadPersoRoster()
    If Not IsEmpty(roster0) Then
        discovered = DiscoverWorkCenters(rawData, roster0)
        If discovered <> "" Then EnsureRosterCheckboxes ThisWorkbook.Sheets("Preferences")
    End If

    Dim roster As Variant
    roster = ReadPersoRoster()
    If IsEmpty(roster) Then
        SetPersoStatus "Error", "Machine Roster table on Preferences is empty -- nothing to build"
        MsgBox "The Machine Roster table on Preferences has no rows. Add at least one " & _
               "(Work Center Match | Display Name | Row Color) and try again.", vbExclamation
        GoTo CleanExit
    End If

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    Dim builtRows As Long
    builtRows = BuildTracieTab(rawData, roster)

    EnforcePersoWorksheetOrder

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True

    ' v2: state TWO facts separately. "Updated At: <hub timestamp>" looked
    ' like Perso's own run time and was not -- the same confusion that cost a
    ' round of diagnosis on Press.
    Dim exportTs As String, pulledTs As String, statusDetail As String
    exportTs = ExtractHubExportTimestamp(hubStatusDetail)
    gHubDataStamp = exportTs
    pulledTs = Format(Now, "yyyy-mm-dd hh:nn")
    If exportTs = "" Then
        statusDetail = "Pulled " & pulledTs & " (Hub data timestamp unreadable)"
    Else
        statusDetail = "Pulled " & pulledTs & " (Hub data as of " & exportTs & ")"
    End If
    SetPersoStatus "Ready", statusDetail

    ' A clean pull here proves nothing about the freshness of what was pulled.
    Dim staleNote As String, hubWhen As Date, hubAgeHrs As Double
    hubWhen = ParseHubTimestamp(exportTs)
    If hubWhen > 0 Then
        hubAgeHrs = (Now - hubWhen) * 24#
        If hubAgeHrs > STALE_HUB_HOURS Then
            staleNote = vbCrLf & vbCrLf & "WARNING -- the Hub's own data is " & _
                Format(hubAgeHrs, "0.0") & " hours old (" & exportTs & ")." & vbCrLf & _
                "This pull succeeded, but Perso can only ever get what the Hub last loaded. " & _
                "Run Update Data in the Hub workbook first if you need today's jobs."
        End If
    End If

    ' Second line on the Update Data button. Written HERE, not in
    ' EnsureUpdateButton, because that runs before the pull and cannot know
    ' the timestamp yet.
    RefreshUpdateButton IIf(staleNote <> "", BTN_STATE_STALE, BTN_STATE_OK), _
                        pulledTs, hubAgeHrs

    ' Land on the Tracie tab -- it's the deliverable here (unlike Press,
    ' where the user asked to land on Preferences). Status is still visible
    ' in the message below and on Preferences whenever needed.
    On Error Resume Next
    ThisWorkbook.Sheets(OUT_SHEET).Activate
    On Error GoTo 0

    MsgBox "Perso data updated in " & Format(Timer - t0, "0.00") & " seconds." & vbCrLf & _
           "Module1 " & MODULE_VERSION & vbCrLf & vbCrLf & _
           "Hub status at pull time: " & hubStatusWord & IIf(hubStatusDetail <> "", " (" & hubStatusDetail & ")", "") & vbCrLf & _
           "Jobs on " & OUT_SHEET & ": " & builtRows & _
           IIf(discovered <> "", vbCrLf & vbCrLf & "New work centers added to the Machine Roster " & _
               "(unticked -- tick to show them):" & discovered, "") & _
           SetupWarningBlock() & staleNote, vbInformation
    Exit Sub

SetupFailed:
    Dim sErrN As Long, sErrD As String
    sErrN = Err.Number: sErrD = Err.Description
    On Error Resume Next
    SetPersoStatus "Error", "Setup failed in " & setupStep & " (" & sErrN & ": " & sErrD & ")"
    On Error GoTo 0
    MsgBox "Perso could not finish setting up the Preferences sheet." & vbCrLf & vbCrLf & _
           "Step: " & setupStep & vbCrLf & _
           "Error " & sErrN & ": " & sErrD & vbCrLf & vbCrLf & _
           "No data was pulled. The " & OUT_SHEET & " tab has not been changed.", _
           vbCritical, "Update Data"
    RefreshUpdateButton BTN_STATE_ERROR, Format(Now, "yyyy-mm-dd hh:nn"), 0
    ShowPreferencesWhenDone
    Exit Sub

CleanExit:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ShowPreferencesWhenDone
    Exit Sub

PullFailed:
    Dim errMsg As String, errNum As Long
    errMsg = Err.Description: errNum = Err.Number
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    SetPersoStatus "Error", "Build failed (" & errNum & ": " & errMsg & ") at " & Format(Now, "yyyy-mm-dd hh:nn")
    RefreshUpdateButton BTN_STATE_ERROR, Format(Now, "yyyy-mm-dd hh:nn"), 0
    ShowPreferencesWhenDone
    MsgBox "Update Data failed: " & errNum & " - " & errMsg, vbCritical, "Perso Update Data Failed"
End Sub

' =============================================================================
' HUB PULL -- byte-for-byte the Press workbook's contract: opens Hub
' read-only, checks Hub Status BEFORE trusting CleanedData (REFRESHING = a
' build is mid-flight, ERROR = last build failed -- both block the pull),
' reads CleanedData as plain values, closes without saving.
' =============================================================================
Private Function PullCleanedDataFromHub(ByVal hubPath As String, ByRef rawData As Variant, _
        ByRef hubStatusWord As String, ByRef hubStatusDetail As String) As Boolean
    PullCleanedDataFromHub = False
    hubStatusWord = "": hubStatusDetail = ""

    If Dir(hubPath) = "" Then
        SetPersoStatus "Error", "Hub file not found at " & hubPath
        MsgBox "Could not find the Hub workbook at:" & vbCrLf & hubPath & vbCrLf & vbCrLf & _
               "Check 'Hub Workbook Path' on Preferences.", vbExclamation, "Update Data"
        Exit Function
    End If

    Dim prevSU As Boolean
    prevSU = Application.ScreenUpdating
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Dim wbHub As Workbook
    On Error Resume Next
    Set wbHub = Workbooks.Open(Filename:=hubPath, ReadOnly:=True, UpdateLinks:=0)
    On Error GoTo 0

    If wbHub Is Nothing Then
        Application.DisplayAlerts = True
        Application.ScreenUpdating = prevSU
        SetPersoStatus "Error", "Could not open Hub workbook (locked or blocked)"
        MsgBox "Found the Hub workbook but could not open it (it may be mid-save or locked by another user).", _
               vbExclamation, "Update Data"
        Exit Function
    End If

    Dim hubPrefs As Worksheet
    On Error Resume Next
    Set hubPrefs = wbHub.Sheets("Preferences")
    On Error GoTo 0

    Dim statusRaw As String
    If Not hubPrefs Is Nothing Then
        statusRaw = Trim(CStr(GetLabeledValueFromSheet(hubPrefs, "Hub Status")))
    End If

    Dim pipePos As Long
    pipePos = InStr(statusRaw, "|")
    If pipePos > 0 Then
        hubStatusWord = Trim(Left(statusRaw, pipePos - 1))
        hubStatusDetail = Trim(Mid(statusRaw, pipePos + 1))
    Else
        hubStatusWord = statusRaw
        hubStatusDetail = ""
    End If

    If UCase(hubStatusWord) = "REFRESHING" Then
        wbHub.Close SaveChanges:=False
        Application.DisplayAlerts = True
        Application.ScreenUpdating = prevSU
        SetPersoStatus "Not Ready", "Hub is REFRESHING -- retry in a moment"
        MsgBox "Hub is currently refreshing (someone clicked its Update Data and it hasn't finished). " & _
               "Wait a moment and click Update again.", vbInformation, "Hub Is Refreshing"
        Exit Function
    ElseIf UCase(hubStatusWord) = "ERROR" Then
        wbHub.Close SaveChanges:=False
        Application.DisplayAlerts = True
        Application.ScreenUpdating = prevSU
        SetPersoStatus "Error", "Hub's last build failed: " & hubStatusDetail
        MsgBox "The Hub workbook's last Update Data run failed:" & vbCrLf & hubStatusDetail & vbCrLf & vbCrLf & _
               "Fix that in the Hub workbook and re-run its Update Data before pulling here -- " & _
               "CleanedData may be incomplete.", vbCritical, "Hub Reported An Error"
        Exit Function
    ElseIf UCase(hubStatusWord) <> "READY" Then
        MsgBox "Hub Status is unrecognized ('" & statusRaw & "') -- proceeding anyway, but verify " & _
               "the Hub workbook has been run at least once.", vbExclamation, "Unrecognized Hub Status"
    End If

    Dim cdSheet As Worksheet
    On Error Resume Next
    Set cdSheet = wbHub.Sheets("CleanedData")
    On Error GoTo 0

    If cdSheet Is Nothing Then
        wbHub.Close SaveChanges:=False
        Application.DisplayAlerts = True
        Application.ScreenUpdating = prevSU
        SetPersoStatus "Error", "Hub workbook has no CleanedData sheet"
        MsgBox "The Hub workbook does not have a CleanedData sheet.", vbCritical, "Update Data"
        Exit Function
    End If

    rawData = cdSheet.UsedRange.Value
    wbHub.Close SaveChanges:=False
    Application.DisplayAlerts = True
    Application.ScreenUpdating = prevSU

    If IsEmpty(rawData) Then
        SetPersoStatus "Error", "Hub CleanedData was empty"
        MsgBox "Hub's CleanedData sheet has no data.", vbExclamation, "Update Data"
        Exit Function
    End If

    PullCleanedDataFromHub = True
End Function

' =============================================================================
' MACHINE ROSTER -- label-anchored on Preferences, same shape as Press's
' roster but Perso's columns are: Work Center Match | Display Name |
' Row Color. The Row Color column is read by CELL FILL, not cell value --
' paint it any color and that machine's rows follow on the next update; a
' no-fill color cell means "no row color for this machine."
' Returns a 2D Variant: (i,1)=match keyword, (i,2)=display name,
' (i,3)=color Long or -1 for none.
' =============================================================================
Private Function ReadPersoRoster() As Variant
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Preferences")
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    Dim anchorRow As Long: anchorRow = 0
    Dim r As Long
    For r = 1 To 100
        If Trim(CStr(ws.Cells(r, 1).Value)) = "Machine Roster" Then
            anchorRow = r
            Exit For
        End If
    Next r
    If anchorRow = 0 Then Exit Function

    Dim dataStartRow As Long: dataStartRow = anchorRow + 3  ' anchor, instructions, header, THEN data

    Dim n As Long: n = 0
    Dim results() As Variant
    ReDim results(1 To 200, 1 To 9)

    ' Blanks are stepped over, not treated as the end of the roster. A machine
    ' row inserted by DiscoverWorkCenters sits directly under the last existing
    ' one, so the separator below the block survives -- but a user who leaves a
    ' gap while typing a new machine must not lose every machine after it.
    Dim rosterCell As String
    Dim rBlankRun As Long: rBlankRun = 0
    r = dataStartRow
    Do While r < dataStartRow + 200
        rosterCell = Trim(CStr(ws.Cells(r, RC_MATCH).Value))
        If IsPersoSectionLabel(rosterCell) Then Exit Do
        If rosterCell = "" Then
            rBlankRun = rBlankRun + 1
            If rBlankRun > MAX_PREF_BLANK_RUN Then Exit Do
            GoTo NextRosterRow
        End If
        rBlankRun = 0
        If n >= 200 Then Exit Do
        n = n + 1
        results(n, 1) = rosterCell                                 ' Work Center pattern
        results(n, 2) = Trim(CStr(ws.Cells(r, RC_DISPLAY).Value))  ' Display name (short)
        ' Column D is the SHOW tick. EMPTY reads as SHOWN, so a roster that
        ' predates the tick boxes behaves exactly as it always did; only an
        ' explicit FALSE hides a machine.
        results(n, 4) = True
        If VarType(ws.Cells(r, RC_SHOW).Value) = vbBoolean Then
            results(n, 4) = CBool(ws.Cells(r, RC_SHOW).Value)
        End If
        ' Column C is a STYLE swatch, not just a fill: whatever that cell
        ' looks like, the machine's rows look like. Fill still reads -1 when
        ' the cell has none, so "no fill = no row colour" still holds.
        If ws.Cells(r, RC_STYLE).Interior.ColorIndex = xlNone Then
            results(n, 3) = -1&
        Else
            results(n, 3) = CLng(ws.Cells(r, RC_STYLE).Interior.Color)
        End If
        results(n, 5) = CStr(ws.Cells(r, RC_STYLE).Font.name)
        results(n, 6) = CDbl(ws.Cells(r, RC_STYLE).Font.Size)
        results(n, 7) = CBool(ws.Cells(r, RC_STYLE).Font.Bold)
        results(n, 8) = CBool(ws.Cells(r, RC_STYLE).Font.Italic)
        results(n, 9) = CLng(ws.Cells(r, RC_STYLE).Font.Color)
NextRosterRow:
        r = r + 1
    Loop

    If n = 0 Then Exit Function

    Dim finalArr() As Variant
    ReDim finalArr(1 To n, 1 To 9)
    Dim i As Long, j As Long
    For i = 1 To n
        For j = 1 To 9
            finalArr(i, j) = results(i, j)
        Next j
    Next i
    ReadPersoRoster = finalArr
End Function

' =============================================================================
' LOCATION EXCLUSION -- two label-anchored single-column lists on
' Preferences (both same anchor shape as Press's Location Filter):
'   "Location Exclude"            -- hide a job if Location CONTAINS any of
'                                    these (case/spacing don't matter).
'   "Location Exclude Exceptions" -- ...UNLESS the whole Location exactly
'                                    equals one of these (case/spacing
'                                    don't matter).
' This reproduces the original Power Query rule: keep "Partial in Shipping"
' and "Staged Receiving" exactly, drop every other "shipping" location,
' keep everything else.
' =============================================================================
Private Function ReadLabeledList(ByVal anchorLabel As String) As Variant
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Preferences")
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    Dim anchorRow As Long: anchorRow = 0
    Dim r As Long
    For r = 1 To 200
        If Trim(CStr(ws.Cells(r, 1).Value)) = anchorLabel Then
            anchorRow = r
            Exit For
        End If
    Next r
    If anchorRow = 0 Then Exit Function

    Dim dataStartRow As Long: dataStartRow = anchorRow + 2  ' anchor, instructions, THEN data

    Dim n As Long: n = 0
    Dim results() As String
    ReDim results(1 To 200)

    ' Blanks are stepped over, not treated as the end of the block. See
    ' IsPersoSectionLabel for why.
    Dim cellText As String
    Dim blankRun As Long: blankRun = 0
    r = dataStartRow
    Do While r < dataStartRow + 200
        cellText = Trim(CStr(ws.Cells(r, 1).Value))
        If IsPersoSectionLabel(cellText) Then Exit Do
        If cellText = "" Then
            blankRun = blankRun + 1
            If blankRun > MAX_PREF_BLANK_RUN Then Exit Do
        Else
            blankRun = 0
            If n < 200 Then
                n = n + 1
                results(n) = UCase(cellText)
            End If
        End If
        r = r + 1
    Loop

    If n = 0 Then Exit Function

    Dim finalArr() As String
    ReDim finalArr(1 To n)
    Dim i As Long
    For i = 1 To n
        finalArr(i) = results(i)
    Next i
    ReadLabeledList = finalArr
End Function

' True when this Location should be HIDDEN: contains an exclude phrase and
' is not an exact-match exception. Compare uses UCase + collapsed spacing
' on the phrase side, plain UCase/Trim on the location, matching the
' "case/spacing don't matter" promise made in the Preferences text.
Private Function LocationIsExcluded(ByVal rawLocation As String, ByRef exclList As Variant, _
        ByRef exceptList As Variant) As Boolean
    LocationIsExcluded = False
    If IsEmpty(exclList) Then Exit Function

    Dim ucLoc As String: ucLoc = UCase(Trim(rawLocation))

    If Not IsEmpty(exceptList) Then
        Dim e As Variant
        For Each e In exceptList
            If ucLoc = CStr(e) Then Exit Function  ' exact-match exception: keep
        Next e
    End If

    ' Like, not InStr: the seeded "Shipping" is migrated to "*Shipping*" so
    ' behaviour is identical, but the list can now express "exactly this" as
    ' well as "anything containing this" -- and deleting an entry actually
    ' removes it instead of being masked by a shorter one.
    Dim k As Variant
    For Each k In exclList
        If ucLoc Like CStr(k) Then
            LocationIsExcluded = True
            Exit Function
        End If
    Next k
End Function

' =============================================================================
' TRACIE TAB BUILD
' =============================================================================
Private Function BuildTracieTab(ByRef rawData As Variant, ByRef roster As Variant) As Long
    Dim totalRows As Long
    totalRows = UBound(rawData, 1)

    Dim exclList As Variant, exceptList As Variant
    exclList = ReadLabeledList("Location Exclude")
    exceptList = ReadLabeledList("Location Exclude Exceptions")

    ' ---- 1. Snapshot Rdy so it survives the rebuild (query never did this
    ' -- user-requested improvement). Keyed Job ID + "|" + Machine.
    '
    ' 2026-09.03: Cards was removed (not needed any more, per the user) --
    ' this used to snapshot Cards and Rdy as a pair; now it's Rdy alone.
    Dim ws As Worksheet
    Set ws = GetOrCreateSheet(OUT_SHEET)

    ' A native Collection, not Scripting.Dictionary. The Dictionary is a
    ' Windows-only COM object; with it here LibreOffice could not execute this
    ' procedure at all, so the entire Tracie build was untestable.
    Dim checkState As Collection
    Set checkState = New Collection
    Dim lastOld As Long
    lastOld = ws.Cells(ws.Rows.Count, 3).End(xlUp).Row
    Dim r As Long, key As String
    If lastOld >= 2 Then
        For r = 2 To lastOld
            key = Trim(CStr(ws.Cells(r, 3).Value)) & "|" & Trim(CStr(ws.Cells(r, 8).Value))
            If key <> "|" And Not CollHasKey(checkState, key) Then
                checkState.Add ws.Cells(r, 9).Value, key
            End If
        Next r
    End If

    ' ---- 2. Filter CleanedData: Work Center matches a roster row (contains,
    ' case/spacing-insensitive) and Location is not excluded.
    Dim matchIdx() As Long, matchMachine() As String
    ReDim matchIdx(1 To totalRows)
    ReDim matchMachine(1 To totalRows)
    Dim matchCount As Long: matchCount = 0

    ' Matching is EXACT WITH WILDCARDS (VBA Like), not "contains". Contains
    ' let roster entries silently swallow one another -- the same defect that
    ' made Press's Location Filter uncontrollable, where "TI" matched the "ti"
    ' inside "To Imposition" and removing an entry changed nothing. "*" and "?"
    ' mean what they mean in the Hub's Exclusion table.
    '
    ' Note this compares TRIMMED UPPER text, not CanonText: CanonText strips
    ' spaces, and a pattern like "Personalization 12*" needs its space.
    Dim i As Long, m As Long
    Dim wcUp As String, kwUp As String
    For i = 2 To totalRows  ' row 1 = CleanedData header
        If Not IsEmpty(rawData(i, COL_JOBID)) And Trim(CStr(rawData(i, COL_JOBID))) <> "" Then
            wcUp = UCase(Trim(CStr(rawData(i, COL_WORKCENTER))))
            For m = 1 To UBound(roster, 1)
                kwUp = UCase(Trim(CStr(roster(m, 1))))
                If kwUp <> "" And CBool(roster(m, 4)) And wcUp Like kwUp Then
                    If Not LocationIsExcluded(CStr(rawData(i, COL_LOCATION)), exclList, exceptList) Then
                        matchCount = matchCount + 1
                        matchIdx(matchCount) = i
                        matchMachine(matchCount) = CStr(roster(m, 2))
                    End If
                    Exit For  ' first roster match wins
                End If
            Next m
        End If
    Next i

    ' ---- 3. Sort: Machine asc, Week Start asc, Ship Date asc (undated rows
    ' first within their machine, matching Power Query nulls-first), then
    ' numeric-aware Job ID. Insertion sort -- row counts are small.
    Dim j As Long, keyIdx As Long, keyMach As String
    For i = 2 To matchCount
        keyIdx = matchIdx(i)
        keyMach = matchMachine(i)
        j = i - 1
        Do While j >= 1
            If ComparePersoRows(rawData, matchIdx(j), matchMachine(j), keyIdx, keyMach) > 0 Then
                matchIdx(j + 1) = matchIdx(j)
                matchMachine(j + 1) = matchMachine(j)
                j = j - 1
            Else
                Exit Do
            End If
        Loop
        matchIdx(j + 1) = keyIdx
        matchMachine(j + 1) = keyMach
    Next i

    ' ---- 4. Clear and rewrite. ClearContents (NOT .Clear) keeps the
    ' checkbox cell formatting on I (Rdy) alive across rebuilds; stale
    ' direct fills (last run's black spacer rows) are reset explicitly,
    ' and stale row heights reset too (same .Cells.Clear-doesn't-reset-
    ' RowHeight lesson as Press v14).
    ' Every clear below is SCOPED TO COLUMNS A:J (was A:K with the Cards
    ' column, before it was removed 2026-09.03; A:J again now, same as
    ' before To Perso was added). The user keeps a hand-built weekly
    ' summary grid (dynamic-array SORT/UNIQUE/FILTER spills plus a SUMIFS
    ' week-by-machine table) to the RIGHT of the data columns on this same
    ' sheet -- a whole-sheet clear would destroy it (2026-08.12, found
    ' while fixing the sheet1 metadata repair). The side grid's own
    ' conditional formatting is likewise left alone.
    Dim clearRange As Range
    Set clearRange = ws.Range(ws.Columns(1), ws.Columns(10))
    clearRange.ClearContents
    clearRange.Interior.Pattern = xlNone
    clearRange.FormatConditions.Delete

    ' Column K (was L, was J, was K again before that -- see the column
    ' history in the revision notes above) carries STALE conditional-
    ' formatting rules inherited from the original workbook (machine-color
    ' rules on K121:K351 pointing at long-gone row offsets -- the blue
    ' cells the user spotted at K121:K123, back when this really was
    ' column K). It holds no data in this design; purge its rules, fills
    ' AND CONTENTS every run.
    '
    ' 2026-09.03: the ClearContents line is new -- everything else here
    ' only ever purged FORMATTING, never VALUES, so stale dates the user
    ' spotted sitting in K1:K5 (left over from some earlier layout, before
    ' K was this run's dedicated no-data buffer) had nothing that would
    ' ever clear them. This is safe across the whole column, unlike the
    ' capacity zone's own row-1-5 strip starting at L -- that one is
    ' deliberately left as user territory (see EnsureCapacityTable); K
    ' never was, by design, at any row.
    ws.Columns(11).ClearContents
    ws.Columns(11).FormatConditions.Delete
    ws.Columns(11).Interior.Pattern = xlNone
    ws.Columns(11).Borders.LineStyle = xlNone

    ' Stale manual borders inherited from the original workbook run across
    ' row 1 past the data (user-spotted 2026-08.12: a border line through
    ' K1..S1, back when the data ended at J). The capacity-table zone reset
    ' starts at row 6, so clear the top strip K1:AD5 explicitly every run.
    ws.Range("K1:AD5").Borders.LineStyle = xlNone

    ' A:J borders are rebuilt to the data footprint below -- wipe them
    ' first so a shrinking queue never leaves stale lines behind.
    clearRange.Borders.LineStyle = xlNone
    On Error Resume Next
    clearRange.UnMerge
    On Error GoTo 0

    Dim headers As Variant
    headers = Array("Ship Date", "To Perso", "Job ID", "Customer Name", "Description", "QTY", _
                    "Location", "Machine", "Rdy", "Week Start")
    Dim c As Long
    For c = 0 To UBound(headers)
        ws.Cells(1, c + 1).Value = headers(c)
    Next c
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, UBound(headers) + 1))
        .Font.Bold = True
        .Interior.Color = RGB(31, 78, 120)
        .Font.Color = RGB(255, 255, 255)
    End With

    ' Week Start (J) is HIDDEN, not removed (user request 2026-09.03: not
    ' needed for viewing, but ApplyTracieConditionalFormatting's odd-week-
    ' bold rule still reads it via $J2). Set on the column every run so it
    ' can never drift back to visible from a stray unhide.
    ws.Columns(10).Hidden = True

    ' Checkbox cell formatting on I (Rdy only -- Cards was removed
    ' 2026-09.03) sized to the data (user request 2026-08.12: the original
    ' file had it statically painted down to row 380). The expected
    ' footprint = data rows + one spacer row between machine groups;
    ' EnsureCheckboxFormats stamps the checkbox format down exactly that
    ' far and clears any leftover below.
    Dim grpCount As Long: grpCount = 0
    For i = 1 To matchCount
        If i = 1 Then
            grpCount = 1
        ElseIf matchMachine(i) <> matchMachine(i - 1) Then
            grpCount = grpCount + 1
        End If
    Next i
    Dim expectedLast As Long
    expectedLast = 1 + matchCount + IIf(grpCount > 1, grpCount - 1, 0)
    EnsureCheckboxFormats ws, expectedLast

    Dim outRow As Long: outRow = 2
    Dim prevMach As String: prevMach = ""

    Dim rowIdx As Long, sd As Variant, wkStart As Variant

    ' Rows painted black as machine-group spacers -- EnsureCheckboxFormats
    ' (above) already stamped the checkbox format down the WHOLE
    ' contiguous I2:I<expectedLast> span before this loop even ran, which
    ' includes these rows, so a spacer row shows an empty checkbox glyph
    ' unless something removes it. Collected here and cleaned up in one
    ' pass right after the loop (2026-09.03 fix).
    Dim sepRows As Collection
    Set sepRows = New Collection

    For i = 1 To matchCount
        rowIdx = matchIdx(i)

        ' Blank spacer row (filled black) between machine groups -- the
        ' original workbook achieved this look with a CF rule on blank
        ' Machine rows; done directly here since the spacer is structural.
        If i > 1 And matchMachine(i) <> prevMach Then
            ws.Range(ws.Cells(outRow, 1), ws.Cells(outRow, UBound(headers) + 1)).Interior.Color = RGB(0, 0, 0)
            sepRows.Add outRow
            outRow = outRow + 1
        End If
        prevMach = matchMachine(i)

        sd = rawData(rowIdx, COL_SHIP)
        ws.Cells(outRow, 1).Value = sd

        ' To Perso comes straight from Hub's CleanedData column L. Blank on
        ' a Hub that hasn't been updated yet (fewer than 12 columns) or on
        ' a job that hasn't reached that stage -- guarded the same way
        ' Press guards its own read of this column.
        If UBound(rawData, 2) >= COL_TOPERSO Then
            ws.Cells(outRow, 2).Value = rawData(rowIdx, COL_TOPERSO)
        End If
        ws.Cells(outRow, 3).Value = rawData(rowIdx, COL_JOBID)
        ws.Cells(outRow, 4).Value = rawData(rowIdx, COL_CUSTOMER)
        ws.Cells(outRow, 5).Value = rawData(rowIdx, COL_DESC)
        ws.Cells(outRow, 6).Value = rawData(rowIdx, COL_QTY)
        ws.Cells(outRow, 7).Value = Trim(CStr(rawData(rowIdx, COL_LOCATION)))
        ws.Cells(outRow, 8).Value = matchMachine(i)

        ' Rdy: restore the previous run's tick if this job was on the tab
        ' before; default to FALSE (renders as an unticked checkbox where
        ' the checkbox cell format is applied, plain FALSE otherwise).
        key = Trim(CStr(rawData(rowIdx, COL_JOBID))) & "|" & matchMachine(i)
        If CollHasKey(checkState, key) Then
            ws.Cells(outRow, 9).Value = SafeBool(checkState(key))
        Else
            ws.Cells(outRow, 9).Value = False
        End If

        ' Week Start = Monday of the Ship Date's week (blank when undated).
        ' Column kept but HIDDEN (user request 2026-09.03: not needed for
        ' viewing, still needed by the odd-week-bold conditional format in
        ' ApplyTracieConditionalFormatting) -- see the Columns(10).Hidden
        ' line right after the header row is written, above.
        If IsDate(sd) Then
            wkStart = CDate(sd) - Weekday(CDate(sd), vbMonday) + 1
            ws.Cells(outRow, 10).Value = wkStart
        End If

        outRow = outRow + 1
    Next i

    ' Strip the checkbox back off every spacer row's Rdy cell.
    '
    ' 2026-09.03 -- settled by DiagnoseTracieCheckboxes against the live
    ' sheet, after two fixes that each guessed a mechanism and each failed.
    ' What the diagnostic reported for a spacer row's Rdy cell (I5):
    '     value (empty) | NumberFmt General | Interior 0 (black)
    '     Validation none | Shapes on the row: none (1 on the whole sheet,
    '     the Update Data button) | CellControl: PRESENT
    ' So the box is a cell-native CellControl, and it survives BOTH
    ' ClearFormats and a full Clear -- the cell was already emptied, already
    ' General, already black, and the box was still there. Press's inherited
    ' "ClearFormats is the only way to take a cell control off" simply does
    ' not hold here. The only thing that can remove it is the CellControl
    ' API itself; see RemoveCellCheckbox.
    '
    ' Done as ONE call over a Union of every spacer cell rather than a loop:
    ' CellControl works on a multi-cell range (it is how EnsureCheckboxFormats'
    ' counterpart in Press sets a whole column span at once).
    Dim sr As Variant
    Dim sepCells As Range
    For Each sr In sepRows
        If sepCells Is Nothing Then
            Set sepCells = ws.Cells(CLng(sr), 9)
        Else
            Set sepCells = Union(sepCells, ws.Cells(CLng(sr), 9))
        End If
    Next sr

    If Not sepCells Is Nothing Then
        If Not RemoveCellCheckbox(sepCells) Then
            gSetupWarnings = gSetupWarnings & vbCrLf & _
                "- The Rdy checkboxes on the black spacer rows could not be removed. " & _
                "Run DiagnoseTracieCheckboxes (Alt+F8) and send the result on."
        End If
        ' The checkbox paste in EnsureCheckboxFormats also brought the
        ' template cell's own look with it; the spacer's black fill was
        ' painted over that during the loop above, so only re-assert the
        ' pieces the control could have carried.
        sepCells.NumberFormat = "General"
        sepCells.Interior.Color = RGB(0, 0, 0)
    End If

    Dim lastRow As Long: lastRow = outRow - 1

    If lastRow >= 2 Then
        ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, 1)).NumberFormat = "m/d/yyyy"
        ws.Range(ws.Cells(2, 2), ws.Cells(lastRow, 2)).NumberFormat = "m/d/yyyy"
        ws.Range(ws.Cells(2, 10), ws.Cells(lastRow, 10)).NumberFormat = "m/d/yyyy"
        ws.Range(ws.Cells(2, 6), ws.Cells(lastRow, 6)).NumberFormat = "#,##0"

        ' Bottom border on every A:J cell, sized to the populated rows
        ' (user request 2026-08.12) -- inside-horizontal lines plus the
        ' block's bottom edge give each row its underline; nothing is drawn
        ' past column J (Week Start -- Cards was removed 2026-09.03, To
        ' Perso added 2026-09.03 earlier the same day, net back to J) or
        ' below the last populated row.
        With ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, 10))
            .Borders(xlInsideHorizontal).LineStyle = xlContinuous
            .Borders(xlInsideHorizontal).Weight = xlThin
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).Weight = xlThin
        End With
    End If

    ApplyMachineStyles ws, lastRow, roster
    ApplyTracieConditionalFormatting ws, lastRow, roster

    EnsureCapacityTable ws, lastRow, roster

    FormatCapacityTable ws, roster

    ' NO AutoFilter (2026-09.03, user request: "Each of the header rows has
    ' a drop down arrow like a table. If they are not needed remove them.").
    ' Nothing in this module reads or depends on the filter -- it was purely
    ' a UI affordance, and on a sheet that is grouped by machine with black
    ' spacer rows, filtering breaks the grouping it is laid out to show.
    ' Turned OFF explicitly rather than just not set: a filter applied by an
    ' earlier build (or by hand) survives ClearContents and would otherwise
    ' keep its arrows forever.
    ws.AutoFilterMode = False

    On Error Resume Next
    ws.Activate
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = False
    ActiveWindow.SplitColumn = 0
    ActiveWindow.SplitRow = 1
    ActiveWindow.FreezePanes = True
    On Error GoTo 0

    BuildTracieTab = matchCount
End Function

' Machine asc -> Week Start asc (undated first) -> Ship Date asc -> Job ID
' (numeric-aware base/suffix, same comparator family as Press).
Private Function ComparePersoRows(ByRef rawData As Variant, ByVal rowA As Long, ByVal machA As String, _
        ByVal rowB As Long, ByVal machB As String) As Long
    If machA < machB Then ComparePersoRows = -1: Exit Function
    If machA > machB Then ComparePersoRows = 1: Exit Function

    Dim aDated As Boolean, bDated As Boolean
    aDated = IsDate(rawData(rowA, COL_SHIP))
    bDated = IsDate(rawData(rowB, COL_SHIP))
    If aDated And Not bDated Then ComparePersoRows = 1: Exit Function   ' undated first
    If bDated And Not aDated Then ComparePersoRows = -1: Exit Function

    If aDated And bDated Then
        Dim dA As Date, dB As Date
        dA = CDate(rawData(rowA, COL_SHIP))
        dB = CDate(rawData(rowB, COL_SHIP))
        ' Week Start asc is implied by Ship Date asc (same Monday-of-week
        ' function is monotonic in the ship date), so one date compare
        ' covers both sort keys.
        If dA < dB Then ComparePersoRows = -1: Exit Function
        If dA > dB Then ComparePersoRows = 1: Exit Function
    End If

    Dim baseA As String, baseB As String
    Dim sufA As Long, sufB As Long
    ParseJobIdBaseSuffix CStr(rawData(rowA, COL_JOBID)), baseA, sufA
    ParseJobIdBaseSuffix CStr(rawData(rowB, COL_JOBID)), baseB, sufB
    If baseA < baseB Then ComparePersoRows = -1: Exit Function
    If baseA > baseB Then ComparePersoRows = 1: Exit Function
    If sufA < sufB Then ComparePersoRows = -1: Exit Function
    If sufA > sufB Then ComparePersoRows = 1: Exit Function
    ComparePersoRows = 0
End Function

' =============================================================================
' CONDITIONAL FORMATTING -- rebuilt from scratch on every update (the
' original workbook's CF rules died the moment a rebuild shifted their
' ranges; VBA-managed rules can't go stale). Three rule families, matching
' the original workbook's visuals:
'   1. Per-machine row color over A2:J<last> (was A2:K<last> with Cards,
'      before it was removed 2026-09.03; A2:J<last> before that with To
'      Perso added but Cards still present), formula
'      =ISNUMBER(SEARCH("<display>",$H2)), fill = that roster row's Row
'      Color cell fill. No fill on the roster cell = no rule.
'   2. Odd-week BOLD banding: =AND($A2<>"",ISODD(WEEKNUM($J2,21))) --
'      alternate Ship-Date weeks read bold, same as the original ($A2 is
'      Ship Date; $J2 is Week Start, hidden but still read here).
'   3. Hide Rdy text on non-job rows: column I white font when $A2="" --
'      was I:J (Cards/Rdy) until Cards was removed 2026-09.03.
' Rule 1 is added FIRST so machine fills sit above later rules in priority,
' mirroring the original rule order.
' =============================================================================
' Paints each job row with its machine's Style swatch: fill, font name, size,
' bold, italic and colour. Rows whose swatch has no fill keep the sheet
' default, matching the old "no fill = no row colour" rule.
Private Sub ApplyMachineStyles(ByRef ws As Worksheet, ByVal lastRow As Long, ByRef roster As Variant)
    If lastRow < 2 Then Exit Sub
    On Error Resume Next
    Dim r As Long, m As Long
    Dim disp As String, mach As String
    Dim rng As Range
    For r = 2 To lastRow
        mach = Trim(CStr(ws.Cells(r, 8).Value))
        If mach <> "" Then
            For m = 1 To UBound(roster, 1)
                disp = Trim(CStr(roster(m, 2)))
                If disp <> "" And StrComp(disp, mach, vbTextCompare) = 0 Then
                    Set rng = ws.Range(ws.Cells(r, 1), ws.Cells(r, 10))
                    If CLng(roster(m, 3)) <> -1 Then rng.Interior.Color = CLng(roster(m, 3))
                    If CStr(roster(m, 5)) <> "" Then rng.Font.name = CStr(roster(m, 5))
                    If CDbl(roster(m, 6)) > 0 Then rng.Font.Size = CDbl(roster(m, 6))
                    rng.Font.Bold = CBool(roster(m, 7))
                    rng.Font.Italic = CBool(roster(m, 8))
                    rng.Font.Color = CLng(roster(m, 9))
                    Exit For
                End If
            Next m
        End If
    Next r
    Err.Clear
    On Error GoTo 0
End Sub

Private Sub ApplyTracieConditionalFormatting(ByRef ws As Worksheet, ByVal lastRow As Long, ByRef roster As Variant)
    If lastRow < 2 Then Exit Sub

    Dim dataRange As Range
    Set dataRange = ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, 10))

    ' The per-machine look is applied DIRECTLY by ApplyMachineStyles, not as a
    ' conditional format. Conditional formatting cannot carry a font name or
    ' size, so the roster swatch could only ever have been a fill; direct
    ' formatting lets the whole style follow it. Safe here because the tab is
    ' rebuilt from scratch on every run, which is the very reason the ORIGINAL
    ' workbook's CF rules went stale.
    ' $A2/$J2 are Ship Date and Week Start -- $A2<>"" is "this is a real job
    ' row, not a spacer". Ship Date is filled far more consistently than To
    ' Perso, so keying the spacer check on it (rather than whichever column
    ' happens to be A) is deliberate, not incidental. Week Start (J) is a
    ' hidden column (2026-09.03) but still perfectly readable by a formula.
    Dim fc As Object

    Set fc = dataRange.FormatConditions.Add(Type:=xlExpression, _
        Formula1:="=AND($A2<>"""",ISODD(WEEKNUM($J2,21)))")
    fc.Font.Bold = True
    fc.StopIfTrue = False

    ' Column I only (Rdy) -- was I:J (Cards/Rdy) before Cards was removed
    ' 2026-09.03.
    Dim checkRange As Range
    Set checkRange = ws.Range(ws.Cells(2, 9), ws.Cells(lastRow, 9))
    Set fc = checkRange.FormatConditions.Add(Type:=xlExpression, Formula1:="=$A2=""""")
    fc.Font.Color = RGB(255, 255, 255)
    fc.StopIfTrue = False
End Sub

' =============================================================================
' CAPACITY TABLE (columns L+ on Tracie -- was M+ until Cards was removed
' 2026-09.03, and L+ before that again with the To Perso column but Cards
' still present) -- 2026-08.12 'fully dynamic' rebuild, replacing the
' original hand-built grid: 158 per-cell SUMIFS formulas hardcoded to rows
' 2:381 plus two header spills capped at row 381/382 (the hardcoding meant
' any queue past 381 rows silently undercounted, and Excel's repair of the
' v1 file wiped the header spills entirely). Recreated as THREE self-sizing
' spill formulas:
'   M7  machines across:  TRANSPOSE(SORT(UNIQUE(FILTER(H2:H5000,...))))
'   L8  week starts down: SORT(UNIQUE(FILTER(J2:J5000,...)))
'   M8  the whole grid in ONE formula: SUMIFS with the two spill refs
'       ($L$8# rows x $M$7# columns) broadcasting into a 2D result that
'       grows/shrinks with both header spills; zeros render blank.
' (Machine stayed at H throughout; Week Start moved K -> J when Cards was
' removed 2026-09.03, having moved there from J -> K when To Perso was
' added earlier the same day -- net unchanged from before either change.)
' New machines, more weeks, longer queues: nothing to maintain. The 5000-
' row bound on the header FILTERs is ~13x the largest queue seen to date
' (SUMIFS itself scans whole columns, so totals stay exact regardless).
'
' Written ONCE (self-healing): if M7/M8 already hold the dynamic formulas
' nothing is touched, so manual tweaks survive updates. The one-time
' migration clears the legacy grid area (still L1:AD500 -- see the note at
' its check, which moved to L6, right back where it was before the To
' Perso/Cards changes) first -- stale fixed-range formulas or cached spill
' values there would otherwise block the new spills with #SPILL! errors.
' Right-hand edge of the capacity zone, in columns. The table starts at L
' (12) with the week column, so machines run from M (13). It can never
' hold more machines than the roster does, plus a small margin so a
' machine added by Work Center Discovery is already formatted on the run
' that finds it.
'
' Clamped at column 50 (AX) on purpose: AZ1/AZ2 hold the checkbox format
' template that EnsureCheckboxFormats clones from, and a border/fill reset
' reaching those cells could strip the XF complement that carries Excel's
' cell control -- which is the whole mechanism behind the Rdy box (Cards
' was removed 2026-09.03).
Private Function CapacityLastCol(ByRef roster As Variant) As Long
    Dim n As Long: n = 0
    If Not IsEmpty(roster) Then n = UBound(roster, 1)
    CapacityLastCol = 12 + n + 2
    If CapacityLastCol < 13 Then CapacityLastCol = 13
    ' Clamped at column 50 (AX) on purpose regardless of where the table
    ' starts: AZ1/AZ2 hold the checkbox format template (EnsureCheckbox-
    ' Formats), and a border/fill reset reaching those cells could strip
    ' the XF complement that carries Excel's cell control.
    If CapacityLastCol > 50 Then CapacityLastCol = 50
End Function

Private Sub EnsureCapacityTable(ByRef ws As Worksheet, ByVal lastDataRow As Long, ByRef roster As Variant)
    ' Rewritten on every run and SIZED TO THE DATA.
    '
    ' v1 hard-coded $G$2:$G$5000 and skipped the rewrite whenever the formulas
    ' merely looked right. It was dynamic exactly until Tracie passed 5,000
    ' rows, at which point work would silently stop being counted with nothing
    ' to indicate it -- the worst kind of ceiling. The bound now comes from the
    ' row count this run actually produced, so there is no ceiling to hit and
    ' no oversized range to slow the spills down.
    Dim lastRef As Long
    lastRef = lastDataRow
    If lastRef < 2 Then lastRef = 2

    On Error Resume Next

    ' One-time migration: wipe the legacy fixed grid so the spills have room.
    '
    ' 2026-08-31: this cleared L1:AD500 -- rows 1 to 5 INCLUDED. The legacy
    ' grid never occupied those rows (the zone reset below has always started
    ' at row 6), but anything a user parked in that top strip was inside the
    ' blast radius, and the only thing holding the trigger down is L6 still
    ' containing the word "Capacity". Retitle that label, delete the row, or
    ' let it get overwritten, and the next Update Data silently ClearContents
    ' across the whole block. Narrowed to the capacity zone proper, so rows
    ' 1-5 to the right of the data are now genuinely user territory.
    '
    ' 2026-09.03: the check cell moved L6 -> M6 (To Perso column added) ->
    ' L6 again (Cards column removed, later the same day). The clear's left
    ' edge tracks the CURRENT buffer/purge column (see BuildTracieTab) each
    ' time, one column left of wherever the table starts -- so it always
    ' sweeps up whatever the previous build last wrote at its own label
    ' cell, on the one run where the label cell is genuinely empty again.
    If InStr(CStr(ws.Range("L6").Value), "Capacity") = 0 Then
        ws.Range(ws.Cells(6, 11), ws.Cells(500, 30)).ClearContents
    End If

    ws.Range("L6").Value = "Capacity - QTY by Week Start x Machine (auto-sizing)"
    ws.Range("L6").Font.Bold = True
    ws.Range("L6").Font.Italic = True

    ' H = Machine (unchanged throughout), J = Week Start, F = QTY.
    Dim hRef As String, jRef As String
    hRef = "$H$2:$H$" & lastRef
    jRef = "$J$2:$J$" & lastRef

    ws.Range("M7").Formula2 = "=IFERROR(TRANSPOSE(SORT(UNIQUE(FILTER(" & hRef & "," & hRef & "<>"""")))),"""")"
    ws.Range("L8").Formula2 = "=IFERROR(SORT(UNIQUE(FILTER(" & jRef & "," & jRef & "<>""""))),"""")"
    ws.Range("M8").Formula2 = "=LET(s,SUMIFS($F$2:$F$" & lastRef & ",$J$2:$J$" & lastRef & _
                              ",$L$8#,$H$2:$H$" & lastRef & ",$M$7#),IF(s=0,"""",s))"

    ' The header bold is safe to set here -- one row, bounded by the roster.
    ' NUMBER FORMATS ARE NOT SET HERE ANY MORE (2026-08-31).
    '
    ' They were fixed at L8:L500 / M8:AD500, which is both a ceiling (past AD
    ' a spilled machine loses its thousands separator, past row 500 a spilled
    ' week renders as a raw date serial) and a bloat source: formatting 493 x
    ' 19 empty cells instantiates all of them in the file. Sizing them to the
    ' DATA row count instead would remove the ceiling but keep the bloat --
    ' there are only ever as many week rows as there are distinct weeks, which
    ' is a tiny fraction of the job count.
    '
    ' The formats now live in FormatCapacityTable, which runs immediately
    ' after this and is the only place that knows how far the spills actually
    ' reached. That also makes them self-correcting when the table shrinks.
    Dim zoneLastCol As Long: zoneLastCol = CapacityLastCol(roster)
    ws.Range(ws.Cells(7, 13), ws.Cells(7, zoneLastCol)).Font.Bold = True

    If Err.Number <> 0 Then
        gSetupWarnings = gSetupWarnings & vbCrLf & _
            "- The capacity grid formulas could not be written (" & Err.Number & ": " & Err.Description & ")."
        Err.Clear
    End If
    On Error GoTo 0
End Sub

' Capacity-table FORMATTING -- reapplied on EVERY update (unlike the
' formulas above, which are written once and then user-owned): thin
' border grid, bold header row and week column, and per-machine column
' fills, all sized to the CURRENT spill extent so the look tracks the
' data exactly -- however many weeks and machines spill (user request
' 2026-08.12: rows past the old hand-formatted block showed unformatted).
' Column fills come from the Machine Roster's Row Color cells -- the same
' single source of truth that colors the job rows, so changing a roster
' color updates both on the next run. The zone L6:AD500 (was M6:AD500
' while To Perso was in and Cards was still present, and L6:AD500 again
' now that Cards is removed) is treated as VBA-owned for FORMATTING: stale
' borders/fills there are reset each run (values, formulas, and number
' formats are never touched).
Private Sub FormatCapacityTable(ByRef ws As Worksheet, ByRef roster As Variant)
    ' Force this sheet to calculate so the spills are current (the build
    ' runs under Calculation = Manual).
    On Error Resume Next
    ws.Calculate
    On Error GoTo 0

    ' Measure the spill extent by walking the calculated values --
    ' works on any Excel that can calculate the formulas at all.
    ' Walk limits are derived, not fixed at 500/30 (2026-08-31). The row walk
    ' cannot run past the sheet's own used extent; the column walk cannot run
    ' past the roster. Both were hard caps that silently stopped formatting a
    ' table that outgrew them.
    Dim usedLastRow As Long
    usedLastRow = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    If usedLastRow < 8 Then usedLastRow = 8
    Dim maxMachCol As Long: maxMachCol = CapacityLastCol(roster)

    Dim lastWkRow As Long: lastWkRow = 7
    Do While lastWkRow < usedLastRow
        If Trim(CStr(ws.Cells(lastWkRow + 1, 12).Value)) = "" Then Exit Do
        lastWkRow = lastWkRow + 1
    Loop
    Dim lastMachCol As Long: lastMachCol = 12
    Do While lastMachCol < maxMachCol
        If Trim(CStr(ws.Cells(7, lastMachCol + 1).Value)) = "" Then Exit Do
        lastMachCol = lastMachCol + 1
    Loop

    ' Reset the zone's look first so a SHRINKING table never leaves stale
    ' borders or fills behind. The reset window is the CURRENT extent plus a
    ' margin rather than a fixed L6:AD500 block: it has to cover last run's
    ' table, but painting it across every unused row below would instantiate
    ' thousands of empty cells in the file for nothing. A table that loses
    ' more than RESET_MARGIN_ROWS weeks in a single run could leave stale
    ' borders below it -- cosmetic, and one further run clears them.
    Const RESET_MARGIN_ROWS As Long = 120
    Dim resetLastRow As Long: resetLastRow = lastWkRow + RESET_MARGIN_ROWS
    If resetLastRow > usedLastRow Then resetLastRow = usedLastRow
    If resetLastRow < 8 Then resetLastRow = 8

    Dim zone As Range
    Set zone = ws.Range(ws.Cells(6, 12), ws.Cells(resetLastRow, maxMachCol))
    zone.Borders.LineStyle = xlNone
    zone.Interior.Pattern = xlNone
    ws.Range(ws.Cells(7, 12), ws.Cells(resetLastRow, maxMachCol)).Font.Bold = False
    ws.Range(ws.Cells(8, 12), ws.Cells(resetLastRow, maxMachCol)).NumberFormat = "General"

    If lastWkRow < 8 Or lastMachCol < 13 Then Exit Sub  ' nothing spilled

    ' Number formats, sized to what actually spilled (moved here from
    ' EnsureCapacityTable on 2026-08-31 -- see the note there). The reset
    ' above has already returned the whole window to General, so a table that
    ' shrank does not leave formatted empties behind.
    ws.Range(ws.Cells(8, 12), ws.Cells(lastWkRow, 12)).NumberFormat = "m/d/yyyy"
    ws.Range(ws.Cells(8, 13), ws.Cells(lastWkRow, lastMachCol)).NumberFormat = "#,##0"

    Dim tbl As Range
    Set tbl = ws.Range(ws.Cells(7, 12), ws.Cells(lastWkRow, lastMachCol))
    tbl.Borders.LineStyle = xlContinuous
    tbl.Borders.Weight = xlThin
    ws.Range(ws.Cells(7, 12), ws.Cells(7, lastMachCol)).Font.Bold = True
    ws.Range(ws.Cells(8, 12), ws.Cells(lastWkRow, 12)).Font.Bold = True

    ' Machine column fills from the roster's Row Color cells.
    Dim c As Long, m As Long, hdr As String
    For c = 13 To lastMachCol
        hdr = Trim(CStr(ws.Cells(7, c).Value))
        For m = 1 To UBound(roster, 1)
            If StrComp(hdr, CStr(roster(m, 2)), vbTextCompare) = 0 Then
                If CLng(roster(m, 3)) <> -1 Then
                    ws.Range(ws.Cells(8, c), ws.Cells(lastWkRow, c)).Interior.Color = CLng(roster(m, 3))
                End If
                Exit For
            End If
        Next m
    Next c
End Sub

' Keeps the I (Rdy) checkbox cell formatting sized to the CURRENT data
' (user request 2026-08.12; was H:I, then I:J for Cards+Rdy, before Cards
' was removed 2026-09.03). VBA cannot CREATE Excel's checkbox cell
' control, but it can copy the format of a cell that already has it: on
' the first run the original I2 checkbox format is stashed in AZ1 (marked
' at AZ2, far right of anything this workbook uses); every run then
' pastes that stashed format down I2:I<last data row> and CLEARS formats
' below it, so checkboxes always end exactly where the data ends instead
' of at the old hand-painted row 380. AZ1/AZ2 themselves did not move --
' they are an arbitrary far-right stash, not part of the data/capacity
' layout.
Private Sub EnsureCheckboxFormats(ByRef ws As Worksheet, ByVal lastDataRow As Long)
    If Trim(CStr(ws.Range("AZ2").Value)) <> "chk-template" Then
        ws.Range("I2").Copy
        ws.Range("AZ1").PasteSpecial xlPasteFormats
        Application.CutCopyMode = False
        ws.Range("AZ2").Value = "chk-template"
        ws.Range("AZ2").Font.Color = RGB(217, 217, 217)
    End If

    If lastDataRow >= 2 Then
        ws.Range("AZ1").Copy
        ws.Range(ws.Cells(2, 9), ws.Cells(lastDataRow, 9)).PasteSpecial xlPasteFormats
        Application.CutCopyMode = False
    End If

    ' Trim stale checkbox formatting below the data footprint.
    '
    ' 2026-08-31: this was hard-coded to row 2000. Tracie is nowhere near
    ' that today, but Hub's CleanedData is already 5,522 rows and widening
    ' the roster is a two-click change on Preferences -- past 2000 the
    ' cleanup would simply stop covering the tail, with no error and no
    ' symptom except checkboxes below the data. Same silent-ceiling class as
    ' the $G$2:$G$5000 bound v2 removed from the capacity formulas.
    ' UsedRange is the right bound here because it counts cells that carry
    ' FORMATTING but no value -- which is precisely what is being cleaned.
    '
    ' 2026-09.03: ClearFormats alone never actually did this job. The
    ' checkbox is a cell-native CellControl and ClearFormats does not touch
    ' it (proven by DiagnoseTracieCheckboxes -- see the note in
    ' BuildTracieTab), so every run since this feature shipped has been
    ' leaving live checkboxes below the data. RemoveCellCheckbox first,
    ' then ClearFormats for everything else that IS a format.
    Dim clearTo As Long
    clearTo = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    If clearTo > lastDataRow Then
        Dim tail As Range
        Set tail = ws.Range(ws.Cells(lastDataRow + 1, 9), ws.Cells(clearTo, 9))
        RemoveCellCheckbox tail
        tail.ClearFormats
    End If
End Sub

' =============================================================================
' REMOVE AN IN-CELL CHECKBOX (2026-09.03)
'
' The one thing that actually removes this workbook's checkboxes. They are
' cell-native CellControls: ClearFormats does not remove them, a full Clear
' does not remove them, and pasting a checkbox-free format over them does
' not remove them either -- all three were tried against the live sheet and
' all three left CellControl reporting PRESENT.
'
' Every call is LATE-BOUND on purpose. CellControl only exists on recent
' builds of Excel, and the method that clears one is not something this can
' verify from here, so a name this Excel does not have has to be a catchable
' runtime error rather than a compile error that would break the whole
' module. SetNone is the documented counterpart to the SetCheckbox that
' Press uses; the rest are tried in case this build names it differently,
' and each attempt is verified by re-reading the control rather than
' assumed to have worked.
'
' Returns True if the range ends up with no checkbox -- including on an
' Excel too old to have CellControl at all, which therefore cannot have put
' one there in the first place.
' =============================================================================
Private Function RemoveCellCheckbox(ByRef target As Range) As Boolean
    If target Is Nothing Then
        RemoveCellCheckbox = True
        Exit Function
    End If

    Dim cc As Object
    On Error Resume Next

    Set cc = Nothing
    Err.Clear
    Set cc = target.CellControl
    If Err.Number <> 0 Or cc Is Nothing Then
        Err.Clear
        On Error GoTo 0
        RemoveCellCheckbox = True
        Exit Function
    End If

    Err.Clear: cc.SetNone
    If Not CellCheckboxGone(target) Then Err.Clear: cc.Remove
    If Not CellCheckboxGone(target) Then Err.Clear: cc.Delete
    If Not CellCheckboxGone(target) Then Err.Clear: cc.Clear
    Err.Clear
    On Error GoTo 0

    RemoveCellCheckbox = CellCheckboxGone(target)
End Function

' True when the range carries no in-cell checkbox: either no CellControl at
' all, or one whose Type reads as none (0). Both are checked because a build
' may hand back a live object that merely describes an empty control rather
' than dropping the object entirely.
Private Function CellCheckboxGone(ByRef target As Range) As Boolean
    Dim cc As Object
    Dim t As Variant

    CellCheckboxGone = True
    On Error Resume Next

    Set cc = Nothing
    Err.Clear
    Set cc = target.CellControl
    If Err.Number <> 0 Or cc Is Nothing Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If

    t = Empty
    Err.Clear
    t = cc.Type
    If Err.Number = 0 Then
        If Not IsEmpty(t) Then CellCheckboxGone = (CLng(t) = 0)
    Else
        CellCheckboxGone = False
    End If

    Err.Clear
    On Error GoTo 0
End Function

' =============================================================================
' PREFERENCES -- same label-anchored, self-healing pattern as Hub/Press.
' =============================================================================
Private Sub EnsurePreferencesSheet()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Preferences")
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(1))
        ws.name = "Preferences"
    End If

    If Trim(CStr(ws.Cells(1, 1).Value)) <> "Perso (Hub Satellite)" Then
        ws.Cells.Clear
        ws.Cells(1, 1).Value = "Perso (Hub Satellite)"
        ws.Cells(1, 1).Font.Bold = True
        ws.Cells(1, 1).Font.Size = 14
        ws.Cells(2, 1).Value = "Point 'Hub Workbook Path' at PLI Hub.xlsm, then click Update Data below."
        ws.Cells(2, 1).Font.Italic = True
    End If

    EnsureOnePrefRow ws, "Hub Workbook Path", "path", "P:\Production Dashboard\PLI Hub.xlsm", 0, 0
    EnsureOnePrefRow ws, "Perso Status", "status", "Not Ready - not yet run", RGB(226, 239, 218), RGB(0, 97, 0)
    EnsureMachineRosterTable ws
    EnsureLocationExcludeLists ws
    EnsureDiscoverySection ws
    MigratePersoV2 ws
    EnsureRosterHeaderShow ws
    ' Guarded at the CALL as well as inside: failing to draw a tick box must
    ' never stop the Preferences sheet being built. Two separate concerns and
    ' only one of them is fragile -- the same split that stopped a checkbox
    ' failure blocking a grid resize in Press v39.
    On Error Resume Next
    EnsureRosterCheckboxes ws
    If Err.Number <> 0 Then
        gSetupWarnings = gSetupWarnings & vbCrLf & _
            "- Machine Roster tick boxes could not be drawn (" & Err.Number & ": " & Err.Description & _
            "). Showing/hiding a machine still works by typing TRUE/FALSE in the Show column."
        Err.Clear
    End If
    On Error GoTo 0
End Sub

' The roster header gained a SHOW column in v2. Added to sheets that predate
' it; the instruction row is rewritten too, because the old wording promises
' "look for anywhere in a job's Work Center", which is no longer what happens.
Private Sub EnsureRosterHeaderShow(ByRef ws As Worksheet)
    Dim a As Long: a = FindLabelRow(ws, "Machine Roster")
    If a = 0 Then Exit Sub
    Dim headerRow As Long: headerRow = a + 2
    If Trim(CStr(ws.Cells(headerRow, RC_SHOW).Value)) = "" Then
        ws.Cells(headerRow, RC_SHOW).Value = "Show"
        ws.Cells(headerRow, RC_SHOW).Font.Bold = True
        ws.Cells(headerRow, RC_SHOW).Interior.Color = RGB(220, 230, 241)
    End If
    ws.Cells(headerRow, RC_STYLE).Value = "Style"
    ws.Cells(a + 1, 1).Value = "Add a row below to add a machine -- no VBA change needed. " & _
        "Work Center Match is EXACT, with * and ? as wildcards, e.g. 'Personalization 12*' matches " & _
        "'Personalization 12 - 687.10'. Display Name = the short name shown in the Machine column. " & _
        "Style = format THIS CELL however you like (fill, font, size, bold, italic, colour) and that " & _
        "machine's rows follow on the next update. Show = untick to leave a machine off the Tracie tab."
    ws.Cells(a + 1, 1).Font.Size = 9
    ws.Cells(a + 1, 1).Font.Italic = True
    ws.Columns(RC_SHOW).ColumnWidth = 7
End Sub

' Form Control tick boxes down the roster's Show column.
'
' SNAPSHOT -> rebuild -> RESTORE, and the restore is unskippable. A Form
' Control is created UNCHECKED and assigning .LinkedCell pushes that state
' into the cell, so rebuilding is exactly what destroys the ticks -- the bug
' that reset every box in Press v36. From the snapshot onward nothing may
' abort this procedure.
'
' Errors are swallowed so a tick box can never fail a data pull, but they are
' RECORDED. Silence is what made the same failure take three rounds to find in
' Press.
Private Sub EnsureRosterCheckboxes(ByRef ws As Worksheet)
    Dim a As Long: a = FindLabelRow(ws, "Machine Roster")
    If a = 0 Then Exit Sub
    Dim firstRow As Long: firstRow = a + 3
    Dim lastRow As Long: lastRow = RosterLastRow(ws)
    If lastRow < firstRow Then Exit Sub

    Dim r As Long
    Dim wasOn() As Boolean
    ReDim wasOn(firstRow To lastRow)
    For r = firstRow To lastRow
        wasOn(r) = True
        If VarType(ws.Cells(r, RC_SHOW).Value) = vbBoolean Then
            wasOn(r) = CBool(ws.Cells(r, RC_SHOW).Value)
        End If
    Next r

    Dim errN As Long, errD As String
    Dim made As Long: made = 0
    Dim misaligned As String
    Dim prevName As String
    Dim c As Range, cb As Object, i As Long

    On Error Resume Next            ' <-- stays on until after the restore

    ' Form Control creation is in the same family as .Select and fails on a
    ' sheet that is not active.
    prevName = ThisWorkbook.ActiveSheet.name
    If prevName <> ws.name Then ws.Activate
    Err.Clear

    For i = ws.CheckBoxes.Count To 1 Step -1
        If Left(ws.CheckBoxes(i).name, Len(ROSTER_CHK_PREFIX)) = ROSTER_CHK_PREFIX Then
            ws.CheckBoxes(i).Delete
        End If
    Next i
    Err.Clear

    For r = firstRow To lastRow
        Set c = ws.Cells(r, RC_SHOW)
        c.NumberFormat = ";;;"          ' hide the linked TRUE/FALSE under the box
        Set cb = Nothing
        Set cb = ws.CheckBoxes.Add(c.Left + 3, c.Top, 16, c.Height)
        If Not cb Is Nothing Then
            cb.name = ROSTER_CHK_PREFIX & CStr(r)
            cb.caption = ""
            cb.LinkedCell = "'" & ws.name & "'!" & c.Address
            cb.Placement = xlMoveAndSize
            cb.Top = c.Top
            cb.Height = c.Height
            cb.Left = c.Left + 3
            If cb.TopLeftCell.Row <> r Then misaligned = misaligned & " " & c.Address(False, False)
            made = made + 1
        End If
        If Err.Number <> 0 And errN = 0 Then
            errN = Err.Number
            errD = Err.Description
        End If
        Err.Clear
    Next r

    ' THE RESTORE. Unconditional, and the last thing to touch these cells.
    For r = firstRow To lastRow
        ws.Cells(r, RC_SHOW).Value = wasOn(r)
        Err.Clear
    Next r

    If prevName <> "" And prevName <> ws.name Then ThisWorkbook.Sheets(prevName).Activate
    Err.Clear
    On Error GoTo 0

    Dim want As Long: want = lastRow - firstRow + 1
    If made = 0 Then
        gSetupWarnings = gSetupWarnings & vbCrLf & _
            "- No Machine Roster tick boxes could be created" & _
            IIf(errN <> 0, " (" & errN & ": " & errD & ")", "") & _
            ". Showing/hiding a machine still works by typing TRUE/FALSE in the Show column."
    ElseIf made < want Then
        gSetupWarnings = gSetupWarnings & vbCrLf & _
            "- Only " & made & " of " & want & " Machine Roster tick boxes were created."
    End If
    If misaligned <> "" Then
        gSetupWarnings = gSetupWarnings & vbCrLf & _
            "- A roster tick box did not land on the cell it controls (" & Trim(misaligned) & ")."
    End If
End Sub

' Appends work centers Hub knows about that match a Discovery pattern and no
' roster row -- UNTICKED. Returns what it added, for the completion message.
Private Function DiscoverWorkCenters(ByRef rawData As Variant, ByRef roster As Variant) As String
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Preferences")
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    Dim pats As Variant
    pats = ReadLabeledList("Work Center Discovery")
    If IsEmpty(pats) Then Exit Function

    Dim seen As Collection
    Set seen = New Collection

    Dim totalRows As Long: totalRows = UBound(rawData, 1)
    Dim i As Long, m As Long, p As Variant
    Dim wcRaw As String, wcUp As String
    Dim added As String
    Dim addRow As Long: addRow = RosterLastRow(ws)
    If addRow = 0 Then Exit Function

    For i = 2 To totalRows
        wcRaw = Trim(CStr(rawData(i, COL_WORKCENTER)))
        If wcRaw <> "" Then
            wcUp = UCase(wcRaw)
            If Not CollHasKey(seen, wcUp) Then
                seen.Add wcUp, wcUp
                Dim wanted As Boolean: wanted = False
                For Each p In pats
                    If wcUp Like CStr(p) Then wanted = True
                Next p
                If wanted Then
                    Dim known As Boolean: known = False
                    For m = 1 To UBound(roster, 1)
                        If UCase(Trim(CStr(roster(m, 1)))) <> "" Then
                            If wcUp Like UCase(Trim(CStr(roster(m, 1)))) Then known = True
                        End If
                    Next m
                    If Not known Then
                        ' INSERT, never write. Writing at RosterLastRow + 1
                        ' lands on the blank separator row; the machine after
                        ' that lands on the "Location Exclude" LABEL, and
                        ' losing that anchor silently disables the entire
                        ' Shipping exclusion (ReadLabeledList returns Empty,
                        ' LocationIsExcluded then returns False for every
                        ' row). Inserting shifts every section below down
                        ' intact. Rows inherit formatting from the row above,
                        ' which is the roster row -- exactly what is wanted.
                        addRow = addRow + 1
                        ws.Rows(addRow).Insert Shift:=xlDown
                        ws.Cells(addRow, RC_MATCH).Value = wcRaw
                        ws.Cells(addRow, RC_DISPLAY).Value = ShortMachineName(wcRaw)
                        ws.Cells(addRow, RC_SHOW).Value = False
                        added = added & vbCrLf & "    " & wcRaw
                    End If
                End If
            End If
        End If
    Next i
    DiscoverWorkCenters = added
End Function

' "Personalization 8 - 687.4" -> "Personalization 8". A starting point for the
' Display Name; rename it on the sheet to whatever the floor calls it.
Private Function ShortMachineName(ByVal wc As String) As String
    Dim p As Long
    p = InStr(wc, " - ")
    If p > 0 Then
        ShortMachineName = Trim(Left(wc, p - 1))
    Else
        ShortMachineName = Trim(wc)
    End If
End Function

' =============================================================================
' COMPACT TRACIE  --  run manually: Alt+F8 > CompactPersoSheet > Run
'
' Reclaims the empty rows and stray formatting that earlier builds left behind
' below the data. As found on 2026-08-31, Tracie's sheet1.xml was 432 KB
' carrying 212 rows of data: rows 214-2000 each held a pair of empty H/I cells
' (the old EnsureCheckboxFormats cleared H:I to a hard-coded row 2000, and
' ClearFormats MATERIALISES every cell it touches), and L:AD carried ~9,400
' number-formatted empties from the old fixed L8:L500 / M8:AD500 block.
'
' v3 stops both from re-accumulating. This routine removes what is already
' there. It is a ONE-OFF -- run it once after importing v3, then only if the
' sheet ever bloats again. Not called from Update Data: deleting rows is not
' something a routine refresh should ever do on its own.
'
' SAFETY. The delete floor is the LOWEST row that anything real occupies:
'   - the last job row (max over columns A:J -- was A:K with the Cards
'     column, before it was removed 2026-09.03 -- so a black spacer row
'     cannot shorten it),
'   - the bottom of the capacity table's week spill in column L (was M
'     while Cards was still present alongside To Perso),
'   - the bottom of every shape on the sheet (the Update Data button sits
'     at row 1 col L today -- was col M -- but it is user-owned and may
'     have been dragged),
'   - a floor of row 10 regardless.
' Nothing at or above that row is touched. Below it, whole rows are deleted,
' which is what actually removes the cell elements from the file -- clearing
' them would leave them exactly where they are.
'
' Idempotent: a second run finds nothing below the floor and says so.
' =============================================================================
Public Sub CompactPersoSheet()
    If Not IsPersoWorkbook() Then
        MsgBox "This is Module1 from the PLI PERSO workbook, and '" & _
               ThisWorkbook.name & "' is not it. Nothing has been changed.", _
               vbCritical, "Wrong workbook"
        Exit Sub
    End If

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(OUT_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "No '" & OUT_SHEET & "' sheet in this workbook.", vbExclamation, "Compact"
        Exit Sub
    End If

    Dim beforeAddr As String
    beforeAddr = ws.UsedRange.Address(False, False)

    ' ---- Floor 1: the last job row, taken across the data columns. A black
    ' spacer row is blank in every one of them, so column A alone would be
    ' enough -- but taking the max costs nothing and cannot be wrong.
    Dim lastDataRow As Long, c As Long, rr As Long
    lastDataRow = 1
    For c = 1 To 10
        rr = ws.Cells(ws.Rows.Count, c).End(xlUp).Row
        If rr > lastDataRow Then lastDataRow = rr
    Next c

    ' ---- Floor 2: the bottom of the capacity table's week spill.
    Dim capBottom As Long: capBottom = 7
    Do While capBottom < ws.Rows.Count - 1
        If Trim(CStr(ws.Cells(capBottom + 1, 12).Value)) = "" Then Exit Do
        capBottom = capBottom + 1
    Loop

    ' ---- Floor 3: every shape's lowest row. The Update Data button is
    ' user-owned and explicitly allowed to be dragged anywhere.
    Dim shapeBottom As Long: shapeBottom = 1
    Dim shp As Shape
    On Error Resume Next
    For Each shp In ws.Shapes
        rr = 0
        rr = shp.BottomRightCell.Row
        If rr > shapeBottom Then shapeBottom = rr
    Next shp
    On Error GoTo 0

    Dim floorRow As Long
    floorRow = lastDataRow
    If capBottom > floorRow Then floorRow = capBottom
    If shapeBottom > floorRow Then floorRow = shapeBottom
    If floorRow < 10 Then floorRow = 10

    Dim usedLast As Long
    usedLast = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1

    If usedLast <= floorRow Then
        MsgBox "Nothing to compact." & vbCrLf & vbCrLf & _
               "Used range: " & beforeAddr & vbCrLf & _
               "Last real row: " & floorRow & " (jobs " & lastDataRow & _
               ", capacity " & capBottom & ", shapes " & shapeBottom & ")", _
               vbInformation, "Compact " & OUT_SHEET
        Exit Sub
    End If

    If MsgBox("Compact the '" & OUT_SHEET & "' sheet?" & vbCrLf & vbCrLf & _
              "Used range now : " & beforeAddr & vbCrLf & _
              "Last real row  : " & floorRow & vbCrLf & _
              "   jobs " & lastDataRow & " | capacity table " & capBottom & _
              " | shapes " & shapeBottom & vbCrLf & vbCrLf & _
              "Rows " & (floorRow + 1) & ":" & usedLast & " will be DELETED." & vbCrLf & _
              "They hold no values -- only leftover formatting from earlier builds." & vbCrLf & vbCrLf & _
              "Nothing at or above row " & floorRow & " is touched. Save first if you " & _
              "want a way back.", vbYesNo + vbQuestion, "Compact " & OUT_SHEET) <> vbYes Then Exit Sub

    Dim prevCalc As Long, prevScreen As Boolean
    prevCalc = Application.Calculation
    prevScreen = Application.ScreenUpdating
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim errN As Long, errD As String
    On Error Resume Next

    ' Stray formatting in the capacity columns that sits BESIDE the job rows
    ' (between the bottom of the week spill and the last job row). Row
    ' deletion cannot reach it -- those rows carry real data in A:J (was
    ' A:K with Cards, before it was removed 2026-09.03).
    If lastDataRow > capBottom Then
        ws.Range(ws.Cells(capBottom + 1, 11), ws.Cells(lastDataRow, 50)).ClearFormats
    End If

    ws.Range(ws.Rows(floorRow + 1), ws.Rows(usedLast)).Delete

    ' Force Excel to recompute the last cell. Reading UsedRange after a row
    ' delete is what actually resets it; without this the scroll bar and
    ' Ctrl+End keep pointing at the old bottom until the file is reopened.
    Dim dummy As String: dummy = ws.UsedRange.Address(False, False)

    errN = Err.Number: errD = Err.Description
    Err.Clear
    On Error GoTo 0

    Application.Calculation = prevCalc
    Application.ScreenUpdating = prevScreen

    If errN <> 0 Then
        MsgBox "Compact ran into a problem and may not have finished." & vbCrLf & vbCrLf & _
               errN & ": " & errD & vbCrLf & vbCrLf & _
               "Used range is now " & ws.UsedRange.Address(False, False) & ".", _
               vbExclamation, "Compact " & OUT_SHEET
        Exit Sub
    End If

    MsgBox "Compacted." & vbCrLf & vbCrLf & _
           "Used range before : " & beforeAddr & vbCrLf & _
           "Used range after  : " & ws.UsedRange.Address(False, False) & vbCrLf & vbCrLf & _
           "Deleted rows " & (floorRow + 1) & ":" & usedLast & "." & vbCrLf & _
           "SAVE AND REOPEN the workbook to see the file size drop -- Excel only " & _
           "writes the smaller sheet on save.", vbInformation, "Compact " & OUT_SHEET
End Sub

' =============================================================================
' SECTION LABELS on the Preferences sheet.
'
' Every block reader below stops HERE rather than at the first blank row. The
' blank rows between sections are structural, and a user adding an entry will
' quite reasonably leave a gap -- under a first-blank rule that silently drops
' everything after the gap, with no error and no visible symptom until someone
' notices the wrong jobs on the tab.
'
' This is the Press lesson, carried across on 2026-08-31: Press's Location
' Filter reader was byte-identical in v32/v34/v38, read column A only, and
' stopped at the first blank. It was one Update Data away from cutting Dicetrax
' from 665 jobs to 119. Perso's three readers had the same shape.
'
' "Work Center Match" is the roster's HEADER row. It sits above every block's
' first data row, so it is never actually reached -- it is listed as a belt-and
' -braces stop in case an anchor is ever moved.
' =============================================================================
Private Function IsPersoSectionLabel(ByVal s As String) As Boolean
    Select Case Trim(s)
        Case "Hub Workbook Path", "Perso Status", "Machine Roster", _
             "Work Center Match", "Location Exclude", _
             "Location Exclude Exceptions", "Work Center Discovery"
            IsPersoSectionLabel = True
    End Select
End Function

' Row of a column-A label, or 0.
Private Function FindLabelRow(ByRef ws As Worksheet, ByVal label As String) As Long
    Dim r As Long
    For r = 1 To 200
        If Trim(CStr(ws.Cells(r, 1).Value)) = label Then
            FindLabelRow = r
            Exit Function
        End If
    Next r
End Function

' Last roster data row, or 0 when the table has no rows yet.
Private Function RosterLastRow(ByRef ws As Worksheet) As Long
    Dim a As Long: a = FindLabelRow(ws, "Machine Roster")
    If a = 0 Then Exit Function
    Dim d As Long: d = a + 3
    Dim last As Long: last = a + 2
    Dim cellText As String
    Dim blankRun As Long: blankRun = 0
    ' Same rule as the readers: step over blanks, stop at the next section
    ' label. DiscoverWorkCenters inserts directly after the row this returns,
    ' so getting it wrong is how the block below gets overwritten.
    Do While d < a + 203
        cellText = Trim(CStr(ws.Cells(d, RC_MATCH).Value))
        If IsPersoSectionLabel(cellText) Then Exit Do
        If cellText = "" Then
            blankRun = blankRun + 1
            If blankRun > MAX_PREF_BLANK_RUN Then Exit Do
        Else
            blankRun = 0
            last = d
        End If
        d = d + 1
    Loop
    If last > a + 2 Then RosterLastRow = last
End Function

' Which work centers may be auto-added to the roster. Seeded Personalization*
' and Thermal*, at the owner's decision (2026-08.20): Hub carries 41 distinct
' work centers and only these are Perso's, so appending everything would bury
' 4 useful rows under 37 that will never be ticked.
Private Sub EnsureDiscoverySection(ByRef ws As Worksheet)
    If FindLabelRow(ws, "Work Center Discovery") > 0 Then Exit Sub

    Dim lastA As Long
    lastA = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 2
    With ws.Cells(lastA, 1)
        .Value = "Work Center Discovery"
        .Font.Bold = True
    End With
    ws.Cells(lastA + 1, 1).Value = "Work Centers in Hub matching one of these patterns, but with no " & _
        "Machine Roster row, are ADDED to the roster automatically -- UNTICKED, so nothing appears on " & _
        "the Tracie tab until you tick it -- and named in the message at the end of the run. " & _
        "Same * and ? wildcards. Empty this list to switch discovery off."
    ws.Cells(lastA + 1, 1).Font.Size = 9
    ws.Cells(lastA + 1, 1).Font.Italic = True
    ws.Cells(lastA + 2, 1).Value = "Personalization*"
    ws.Cells(lastA + 3, 1).Value = "Thermal*"
End Sub

' One-time migration to v2 semantics. Guarded by a marker rather than by the
' shape of the data, so it cannot re-fire and mangle a list the user has since
' edited by hand.
'
' Two rewrites, both of which keep behaviour IDENTICAL under the new matcher:
'   roster "Personalization 12"  ->  "Personalization 12*"
'   exclude "Shipping"           ->  "*Shipping*"
' Without them the switch from contains to Like would silently match nothing
' and empty the Tracie tab -- the same trap as Press's "Ready to Sch".
Private Sub MigratePersoV2(ByRef ws As Worksheet)
    Dim a As Long: a = FindLabelRow(ws, "Machine Roster")
    If a = 0 Then Exit Sub
    If Trim(CStr(ws.Cells(a, PERSO_MARK_COL).Value)) = PERSO_MARK Then Exit Sub

    Dim last As Long: last = RosterLastRow(ws)
    Dim r As Long, t As String
    For r = a + 3 To last
        t = Trim(CStr(ws.Cells(r, RC_MATCH).Value))
        If t <> "" And InStr(t, "*") = 0 And InStr(t, "?") = 0 Then
            ws.Cells(r, RC_MATCH).Value = t & "*"
        End If
        ' every pre-existing machine stays visible
        If VarType(ws.Cells(r, RC_SHOW).Value) <> vbBoolean Then
            ws.Cells(r, RC_SHOW).Value = True
        End If
    Next r

    Dim e As Long: e = FindLabelRow(ws, "Location Exclude")
    If e > 0 Then
        For r = e + 2 To e + 202
            t = Trim(CStr(ws.Cells(r, 1).Value))
            If t = "" Then Exit For
            If InStr(t, "*") = 0 And InStr(t, "?") = 0 Then
                ws.Cells(r, 1).Value = "*" & t & "*"
            End If
        Next r
    End If

    ws.Cells(a, PERSO_MARK_COL).Value = PERSO_MARK
    ws.Cells(a, PERSO_MARK_COL).Font.Color = RGB(217, 217, 217)
End Sub

Private Sub EnsureOnePrefRow(ByRef ws As Worksheet, ByVal lbl As String, ByVal kind As String, _
        ByVal defaultText As String, ByVal fillC As Long, ByVal fontC As Long)
    Dim r As Long
    For r = 1 To 100
        If Trim(CStr(ws.Cells(r, 1).Value)) = lbl Then Exit Sub
    Next r

    Dim lastA As Long
    lastA = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 2

    With ws.Cells(lastA, 1)
        .Value = lbl
        .Font.Bold = True
    End With
    With ws.Cells(lastA, 2)
        Select Case kind
            Case "path"
                .NumberFormat = "@"
                .Value = defaultText
            Case "status"
                .NumberFormat = "@"
                .Value = defaultText
                .Interior.Color = fillC
                .Font.Color = fontC
                .HorizontalAlignment = xlHAlignLeft
        End Select
        .BorderAround Weight:=xlThin
    End With

    If lbl = "Hub Workbook Path" Then
        ws.Cells(lastA + 1, 1).Value = "Full path to PLI Hub.xlsm on the shared drive -- pre-filled with the " & _
            "usual location; correct it if the Hub lives elsewhere."
        ws.Cells(lastA + 1, 1).Font.Size = 9
        ws.Cells(lastA + 1, 1).Font.Italic = True
    ElseIf lbl = "Perso Status" Then
        ws.Cells(lastA + 1, 1).Value = "Ready = last pull succeeded (detail shows the Monarch export's own timestamp). " & _
            "Not Ready = pull in progress or waiting on Hub. Error = pull failed -- see detail after the dash."
        ws.Cells(lastA + 1, 1).Font.Size = 9
        ws.Cells(lastA + 1, 1).Font.Italic = True
    End If
End Sub

' Machine Roster: Work Center Match | Display Name | Row Color. The Row
' Color cells are seeded with each machine's color from the ORIGINAL
' workbook's conditional-formatting rules (theme-based fills reproduced via
' ThemeColor so they track the workbook theme exactly). Created once;
' never rewritten -- edit keywords, names, and colors freely.
Private Sub EnsureMachineRosterTable(ByRef ws As Worksheet)
    Dim r As Long
    For r = 1 To 100
        If Trim(CStr(ws.Cells(r, 1).Value)) = "Machine Roster" Then Exit Sub ' already exists
    Next r

    Dim lastA As Long
    lastA = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 2

    With ws.Cells(lastA, 1)
        .Value = "Machine Roster"
        .Font.Bold = True
    End With
    ws.Cells(lastA + 1, 1).Value = "Add a row below to add a machine -- no VBA change needed. " & _
        "Work Center Match = text to look for anywhere in a job's Work Center (case/spacing don't matter). " & _
        "Display Name = the short name shown in the Machine column. " & _
        "Row Color = paint this CELL's fill any color and that machine's rows follow on the next update " & _
        "(no fill = no row color)."
    ws.Cells(lastA + 1, 1).Font.Size = 9
    ws.Cells(lastA + 1, 1).Font.Italic = True

    Dim headerRow As Long: headerRow = lastA + 2
    ws.Cells(headerRow, 1).Value = "Work Center Match"
    ws.Cells(headerRow, 2).Value = "Display Name"
    ws.Cells(headerRow, 3).Value = "Row Color"
    With ws.Range(ws.Cells(headerRow, 1), ws.Cells(headerRow, 3))
        .Font.Bold = True
        .Interior.Color = RGB(220, 230, 241)
    End With

    ' Seeds: keyword, display name, then the color from the original
    ' workbook's CF rule for that machine (theme index/tint or plain RGB).
    Dim seedRow As Long: seedRow = headerRow + 1
    SeedRosterRow ws, seedRow + 0, "Personalization 12", "RFID 12", True, xlThemeColorAccent2, 0.8, 0
    SeedRosterRow ws, seedRow + 1, "Personalization 13", "RFID 13", True, xlThemeColorDark2, 0.9, 0
    SeedRosterRow ws, seedRow + 2, "Personalization 4", "DOD 4", True, xlThemeColorAccent6, 0.8, 0
    SeedRosterRow ws, seedRow + 3, "Personalization 5", "DOD 5", True, xlThemeColorAccent5, 0.8, 0
    SeedRosterRow ws, seedRow + 4, "Personalization 6", "DOD 6", False, 0, 0, RGB(240, 253, 165)
    SeedRosterRow ws, seedRow + 5, "Personalization 7", "DOD 7", True, xlThemeColorLight2, 0, 0
    SeedRosterRow ws, seedRow + 6, "Thermal Printer 1", "Thermal 1", False, 0, 0, RGB(199, 201, 219)

    ws.Columns(1).ColumnWidth = 28
    ws.Columns(2).ColumnWidth = 14
    ws.Columns(3).ColumnWidth = 12
End Sub

Private Sub SeedRosterRow(ByRef ws As Worksheet, ByVal r As Long, ByVal kw As String, ByVal disp As String, _
        ByVal useTheme As Boolean, ByVal themeIdx As Long, ByVal tint As Double, ByVal rgbColor As Long)
    ws.Cells(r, 1).Value = kw
    ws.Cells(r, 2).Value = disp
    With ws.Cells(r, 3).Interior
        If useTheme Then
            .ThemeColor = themeIdx
            If tint <> 0 Then .TintAndShade = tint
        Else
            .Color = rgbColor
        End If
    End With
    ws.Cells(r, 3).BorderAround Weight:=xlThin
End Sub

Private Sub EnsureLocationExcludeLists(ByRef ws As Worksheet)
    Dim r As Long, found As Boolean

    found = False
    For r = 1 To 200
        If Trim(CStr(ws.Cells(r, 1).Value)) = "Location Exclude" Then found = True
    Next r
    If Not found Then
        Dim lastA As Long
        lastA = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 2
        With ws.Cells(lastA, 1)
            .Value = "Location Exclude"
            .Font.Bold = True
        End With
        ws.Cells(lastA + 1, 1).Value = "A job is HIDDEN if its Location contains any phrase below " & _
            "(case/spacing don't matter) -- unless it exactly matches a row in 'Location Exclude Exceptions'. " & _
            "Delete all rows to disable."
        ws.Cells(lastA + 1, 1).Font.Size = 9
        ws.Cells(lastA + 1, 1).Font.Italic = True
        ws.Cells(lastA + 2, 1).Value = "Shipping"
    End If

    found = False
    For r = 1 To 200
        If Trim(CStr(ws.Cells(r, 1).Value)) = "Location Exclude Exceptions" Then found = True
    Next r
    If Not found Then
        Dim lastB As Long
        lastB = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 2
        With ws.Cells(lastB, 1)
            .Value = "Location Exclude Exceptions"
            .Font.Bold = True
        End With
        ws.Cells(lastB + 1, 1).Value = "Locations kept even though they contain an excluded phrase. " & _
            "Whole-value exact match (case/spacing don't matter)."
        ws.Cells(lastB + 1, 1).Font.Size = 9
        ws.Cells(lastB + 1, 1).Font.Italic = True
        ws.Cells(lastB + 2, 1).Value = "Partial in Shipping"
        ws.Cells(lastB + 3, 1).Value = "Staged Receiving"
    End If
End Sub

' =============================================================================
' UPDATE BUTTON -- user-owned after creation (same 2026-08.12 pattern as
' the Hub's reworked button): if a shape named "PersoUpdateDataButton"
' exists on ANY sheet, only its macro wiring is refreshed -- position,
' size, color, and text are NEVER overwritten on update. Move it or
' restyle it freely. Only when no sheet has the button is it created
' fresh on TRACIE (above the capacity table) with the default style below.
' =============================================================================
' =============================================================================
' BUTTON TEXT CONTRAST
'
' The fill changes -- amber when the Hub data behind the pull is stale, red on
' a failure, and whatever the user picks the rest of the time -- so the caption
' colour cannot be a fixed white. These pick it from the fill actually in place.
'
' WCAG 2.x relative luminance, then the higher of the two contrast ratios:
'   linear(c) = c/12.92                    for c <= 0.03928
'             = ((c + 0.055)/1.055) ^ 2.4  otherwise
'   L = 0.2126*R + 0.7152*G + 0.0722*B
'   vs white = 1.05 / (L + 0.05)      vs black = (L + 0.05) / 0.05
' Deliberately the real formula and not a "brightness > 128" shortcut: green
' carries roughly ten times the perceived weight of blue, so the naive version
' puts white text on amber, which is the one state that most needs reading.
' =============================================================================
Private Function SrgbLinear(ByVal channel As Long) As Double
    Dim c As Double
    c = channel / 255#
    If c <= 0.03928 Then
        SrgbLinear = c / 12.92
    Else
        SrgbLinear = ((c + 0.055) / 1.055) ^ 2.4
    End If
End Function

Private Function PickTextColor(ByVal fillColor As Long) As Long
    Dim r As Long, g As Long, b As Long
    r = fillColor And &HFF&
    g = (fillColor \ &H100&) And &HFF&
    b = (fillColor \ &H10000) And &HFF&

    Dim lum As Double
    lum = 0.2126 * SrgbLinear(r) + 0.7152 * SrgbLinear(g) + 0.0722 * SrgbLinear(b)

    If (1.05 / (lum + 0.05)) >= ((lum + 0.05) / 0.05) Then
        PickTextColor = BTN_TEXT_LIGHT
    Else
        PickTextColor = BTN_TEXT_DARK
    End If
End Function

' Current solid fill of a shape, or -1 when it has none / cannot be read.
Private Function ShapeFillColor(ByRef shp As Shape) As Long
    Dim v As Long: v = -1
    On Error Resume Next
    If shp.Fill.Visible = msoTrue Then v = CLng(shp.Fill.ForeColor.RGB)
    On Error GoTo 0
    ShapeFillColor = v
End Function

' =============================================================================
' UPDATE DATA BUTTON -- line 2 and state colour.
'
' Line 1 is whatever the user has it say; it is read back and re-used verbatim,
' so renaming the button to anything survives. Only line 2 is managed.
'
' Separator: written with vbLf and read back by splitting on BOTH vbCr and
' vbLf, so it does not matter which one Excel normalises shape text to -- a
' detail this build could not exercise, having no Excel to run against.
' =============================================================================
Private Sub RefreshUpdateButton(ByVal state As String, ByVal pulledTs As String, _
        ByVal hubAgeHrs As Double)
    Dim ws As Worksheet, shp As Shape
    Set shp = Nothing
    For Each ws In ThisWorkbook.Worksheets
        On Error Resume Next
        Set shp = ws.Shapes(BTN_NAME)
        On Error GoTo 0
        If Not shp Is Nothing Then Exit For
    Next ws
    If shp Is Nothing Then
        gSetupWarnings = gSetupWarnings & vbCrLf & _
            "- The Update Data button (shape '" & BTN_NAME & "') was not found, " & _
            "so its caption could not be updated."
        Exit Sub
    End If

    ' Failures below are non-fatal -- a caption is not worth killing a good
    ' pull over -- but they are REPORTED. v3.3 wrote this whole body under a
    ' blanket On Error Resume Next, so when the caption did not appear there
    ' was nothing anywhere to say why. This module's own rule is "no silent
    ' failures"; that was a straight violation of it.
    On Error Resume Next
    Err.Clear

    ' ---- Line 1: keep whatever it currently says.
    Dim caption As String, line1 As String
    caption = shp.TextFrame2.TextRange.Text
    caption = Replace(caption, vbCr, vbLf)
    line1 = Trim(Split(caption, vbLf)(0))
    If line1 = "" Then line1 = "Update Data"

    ' ---- Base fill. AlternativeText carries "PersoBtn|base=<n>|last=<n>".
    ' If what is on the shape now is not what this code last set, the user has
    ' recoloured it and that colour becomes the new base.
    Dim tag As String, baseFill As Long, lastSet As Long, nowFill As Long
    tag = shp.AlternativeText
    baseFill = -1: lastSet = -1
    If InStr(tag, BTN_TAG) = 1 Then
        baseFill = ReadTagNum(tag, "base=")
        lastSet = ReadTagNum(tag, "last=")
    End If
    nowFill = ShapeFillColor(shp)
    If baseFill < 0 Then baseFill = IIf(nowFill >= 0, nowFill, BTN_FILL_BASE)
    If nowFill >= 0 And nowFill <> lastSet Then baseFill = nowFill

    Dim useFill As Long, line2 As String
    Select Case state
        Case BTN_STATE_ERROR
            useFill = BTN_FILL_ERROR
            line2 = "! Last pull FAILED " & pulledTs
        Case BTN_STATE_STALE
            useFill = BTN_FILL_STALE
            ' The marker repeats the warning in text, so the state is still
            ' readable to someone who cannot separate amber from blue.
            line2 = "! Pulled " & pulledTs & " - Hub data " & _
                    Format(hubAgeHrs / 24#, "0.0") & "d old"
        Case Else
            useFill = baseFill
            line2 = "Pulled " & pulledTs
    End Select

    ' ---- Height: grow whenever the box is too short for two lines. NEVER
    ' shrink, and never touch width or position.
    '
    ' v3.3 grew the height only while the button was still EXACTLY 160x34, on
    ' the theory that any other size meant the user had chosen it. Wrong test:
    ' the shape uses a twoCellAnchor, so its height is derived from the row
    ' heights it spans and drifts off 34 the moment anything above it changes
    ' -- after which the second line was written into a box too short to show
    ' it, which looks exactly like the caption never updating at all.
    ' Growing on "too short" is what the user actually wants, and a button
    ' someone has deliberately made taller is left alone because it already
    ' passes the test.
    If shp.Height < BTN_TWO_LINE_H - 0.5 Then shp.Height = BTN_TWO_LINE_H

    shp.Fill.Visible = msoTrue
    shp.Fill.ForeColor.RGB = useFill
    shp.AlternativeText = BTN_TAG & "base=" & baseFill & "|last=" & useFill

    Dim txtColor As Long
    txtColor = PickTextColor(useFill)

    ' ---- Write the caption, then PROVE it took.
    '
    ' v3.3 wrote line1 & vbCr & line2 and assumed it worked. It did not, and
    ' there was nothing in the workbook to say so. Which separator an Excel
    ' shape accepts is not settled by the documentation -- Microsoft's
    ' TextRange2 and TextFrame2 pages do not state it, and third-party sources
    ' disagree (vbCrLf shown working in shapes, vbLf documented for cells) --
    ' and this build has no Excel to test against.
    '
    ' So stop asserting and start checking: try each separator and read the
    ' caption back until BOTH lines are actually in the shape. Whichever one
    ' this copy of Excel honours, the button ends up right.
    Dim seps As Variant, si As Long, readBack As String, wrote As Boolean
    seps = Array(vbLf, vbCr, vbCrLf)
    wrote = False
    For si = 0 To UBound(seps)
        Err.Clear
        shp.TextFrame2.TextRange.Text = line1 & CStr(seps(si)) & line2
        readBack = shp.TextFrame2.TextRange.Text
        If Len(readBack) >= Len(line1) + Len(line2) Then
            If InStr(readBack, line2) > 0 Then
                wrote = True
                Exit For
            End If
        End If
    Next si

    shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    shp.TextFrame2.VerticalAnchor = msoAnchorMiddle

    ' Size the two lines by CHARACTER RANGE, not by Paragraphs(). Paragraphs
    ' only exist if the separator created them, which is the very thing that
    ' could not be relied on; character offsets are found in the text that
    ' actually came back, so they are correct whatever Excel did with it.
    Dim p2 As Long
    p2 = InStr(readBack, line2)
    With shp.TextFrame2.TextRange.Characters(1, Len(line1))
        .Font.Size = 14
        .Font.Bold = msoTrue
        .Font.Fill.ForeColor.RGB = txtColor
    End With
    If wrote And p2 > 0 Then
        With shp.TextFrame2.TextRange.Characters(p2, Len(line2))
            .Font.Size = 9
            .Font.Bold = msoFalse
            .Font.Fill.ForeColor.RGB = txtColor
        End With
    Else
        ' Fall back to a single readable line rather than leaving a caption
        ' that is half-written, and say so in the finish message.
        shp.TextFrame2.TextRange.Text = line1
        shp.TextFrame2.TextRange.Font.Size = 14
        shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = txtColor
        gSetupWarnings = gSetupWarnings & vbCrLf & _
            "- The Update Data button would not take a second caption line " & _
            "(tried Lf, Cr and CrLf). Caption left as """ & line1 & """; shape is " & _
            Format(shp.Width, "0") & "x" & Format(shp.Height, "0") & " pt." & _
            IIf(Err.Number <> 0, " Last error " & Err.Number & ": " & Err.Description, "") & _
            " Run DiagnosePersoButton (Alt+F8) and send me what it reports."
    End If

    Err.Clear
    On Error GoTo 0
End Sub

' =============================================================================
' Is this actually the Perso workbook? Two independent marks, either is enough:
' the output tab exists, or Preferences carries Perso's own title in A1. Both
' are written by this module, so a Perso workbook that has never been built
' still passes on the second once EnsurePreferencesSheet has run --  and a
' brand-new empty workbook fails both, which is the correct answer.
' =============================================================================
Private Function SheetExists(ByVal nm As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(nm)
    On Error GoTo 0
    SheetExists = Not ws Is Nothing
End Function

' The guard FAILS OPEN, deliberately.
'
' v3.6 asked "does this look like Perso?" and refused when the answer was no.
' That is the wrong shape for a guard: a Perso workbook whose output tab has
' been renamed, or whose Preferences title has been edited, would have been
' locked out of its own Update Data -- a worse failure than the one being
' guarded against, and caused by the guard itself.
'
' So it now refuses only when the workbook is POSITIVELY something else. Hub
' and the AVL Dashboard both carry a CleanedData sheet and Perso never does,
' which is the one mark that separates them. Anything unrecognised is allowed
' through, exactly as before the guard existed.
Private Function IsPersoWorkbook() As Boolean
    IsPersoWorkbook = True
    If SheetExists(OUT_SHEET) Then Exit Function
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Preferences")
    On Error GoTo 0
    If Not ws Is Nothing Then
        If Trim(CStr(ws.Cells(1, 1).Value)) = "Perso (Hub Satellite)" Then Exit Function
    End If
    ' No Perso mark. Refuse ONLY for a workbook that is recognisably Hub or
    ' the Dashboard -- the case where EnsurePreferencesSheet would wipe a
    ' Preferences sheet that matters.
    If SheetExists("CleanedData") Then IsPersoWorkbook = False
End Function

' =============================================================================
' BUTTON DIAGNOSTIC  --  Alt+F8 > DiagnosePersoButton > Run
'
' Reports what the Update Data button actually IS, so a caption that will not
' render can be diagnosed from facts instead of from guesses about what Excel
' does with shape text. Reads only; changes nothing.
' =============================================================================
Public Sub DiagnosePersoButton()
    Dim ws As Worksheet, shp As Shape, found As String
    Dim rpt As String
    rpt = "UPDATE DATA BUTTON -- Module1 " & MODULE_VERSION & vbCrLf & String(52, "-") & vbCrLf

    For Each ws In ThisWorkbook.Worksheets
        Set shp = Nothing
        On Error Resume Next
        Set shp = ws.Shapes(BTN_NAME)
        On Error GoTo 0
        If Not shp Is Nothing Then
            found = ws.name
            Exit For
        End If
    Next ws

    If shp Is Nothing Then
        rpt = rpt & "NOT FOUND. No shape named '" & BTN_NAME & "' on any sheet." & vbCrLf & vbCrLf

        ' 2026-08-31: v3.5 listed only shapes here, sheet-by-shape. A sheet
        ' with NO shapes printed nothing at all -- so a report with no Tracie
        ' line was read as "no Tracie sheet" and sent the diagnosis into the
        ' wrong workbook entirely. Every sheet is now listed whether it holds
        ' a shape or not, because absence of evidence was mistaken for
        ' evidence of absence.
        rpt = rpt & "Workbook : " & ThisWorkbook.name & vbCrLf
        rpt = rpt & "Path     : " & ThisWorkbook.Path & vbCrLf & vbCrLf
        rpt = rpt & "ALL SHEETS in this workbook:" & vbCrLf
        For Each ws In ThisWorkbook.Worksheets
            rpt = rpt & "   " & ws.name & "   (" & ws.Shapes.Count & " shape(s))" & _
                  IIf(ws.name = OUT_SHEET, "   <-- the output tab", "") & vbCrLf
        Next ws

        rpt = rpt & vbCrLf & "ALL SHAPES, with what each is wired to:" & vbCrLf
        Dim s2 As Shape, act As String
        For Each ws In ThisWorkbook.Worksheets
            For Each s2 In ws.Shapes
                act = ""
                On Error Resume Next
                act = s2.OnAction
                On Error GoTo 0
                rpt = rpt & "   " & ws.name & " : " & s2.name & _
                      "   type " & s2.Type & _
                      "   " & Format(s2.Width, "0") & "x" & Format(s2.Height, "0") & "pt" & _
                      "   macro: " & IIf(act = "", "(none)", act) & vbCrLf
            Next s2
        Next ws

        Dim pws As Worksheet
        On Error Resume Next
        Set pws = ThisWorkbook.Sheets("Preferences")
        If Not pws Is Nothing Then
            rpt = rpt & vbCrLf & "Preferences!A1 : " & _
                  Chr(34) & Trim(CStr(pws.Cells(1, 1).Value)) & Chr(34) & vbCrLf
        End If
        On Error GoTo 0

        rpt = rpt & vbCrLf & "The next Update Data will create the button on '" & _
              OUT_SHEET & "' if that sheet exists." & vbCrLf
        MsgBox rpt, vbExclamation, "Diagnose Button"
        Exit Sub
    End If

    On Error Resume Next
    rpt = rpt & "Sheet          : " & found & vbCrLf
    rpt = rpt & "Shape type     : " & shp.Type & "  (17 = AutoShape, 8 = Form Control)" & vbCrLf
    rpt = rpt & "AutoShape type : " & shp.AutoShapeType & vbCrLf
    rpt = rpt & "Size           : " & Format(shp.Width, "0.0") & " x " & _
                Format(shp.Height, "0.0") & " pt" & vbCrLf
    rpt = rpt & "OnAction       : " & shp.OnAction & vbCrLf
    rpt = rpt & "AlternativeText: " & shp.AlternativeText & vbCrLf
    rpt = rpt & "Fill visible   : " & shp.Fill.Visible & "   RGB: " & shp.Fill.ForeColor.RGB & vbCrLf

    Dim t As String
    Err.Clear
    t = shp.TextFrame2.TextRange.Text
    If Err.Number <> 0 Then
        rpt = rpt & "TextFrame2     : NOT ACCESSIBLE -- " & Err.Number & ": " & Err.Description & vbCrLf
        Err.Clear
    Else
        rpt = rpt & "TextFrame2 len : " & Len(t) & vbCrLf
        rpt = rpt & "Paragraph count: " & shp.TextFrame2.TextRange.Paragraphs.Count & vbCrLf
        Dim i As Long, codes As String
        For i = 1 To Len(t)
            If i > 60 Then
                codes = codes & " ..."
                Exit For
            End If
            If Asc(Mid(t, i, 1)) < 32 Then
                codes = codes & "[" & Asc(Mid(t, i, 1)) & "]"
            Else
                codes = codes & Mid(t, i, 1)
            End If
        Next i
        rpt = rpt & "Caption        : " & codes & vbCrLf
        rpt = rpt & "   ([10] = Lf, [13] = Cr -- a second line shows one of these)" & vbCrLf
    End If

    ' Live probe: does this Excel accept a two-line caption at all, and with
    ' which separator? The original caption is restored either way.
    Dim orig As String: orig = t
    Dim seps As Variant, names As Variant, si As Long, rb As String
    seps = Array(vbLf, vbCr, vbCrLf)
    names = Array("vbLf", "vbCr", "vbCrLf")
    rpt = rpt & vbCrLf & "Separator probe (writes then reads back):" & vbCrLf
    For si = 0 To UBound(seps)
        Err.Clear
        shp.TextFrame2.TextRange.Text = "AAA" & CStr(seps(si)) & "BBB"
        rb = shp.TextFrame2.TextRange.Text
        rpt = rpt & "   " & names(si) & " -> len " & Len(rb) & _
              IIf(InStr(rb, "BBB") > 0, "  BOTH LINES OK", "  SECOND LINE LOST") & _
              IIf(Err.Number <> 0, "  err " & Err.Number, "") & vbCrLf
    Next si
    shp.TextFrame2.TextRange.Text = orig
    Err.Clear
    On Error GoTo 0

    MsgBox rpt, vbInformation, "Diagnose Button"
End Sub

' =============================================================================
' DIAGNOSE THE Rdy CHECKBOXES (2026-09.03) -- Alt+F8, manual, read-only.
'
' Two fixes for "checkboxes in the black spacer rows" each assumed a
' different mechanism and each failed, because there is no way to tell from
' outside Excel which one this workbook actually uses: a cell-native
' checkbox (the modern in-cell control), a leftover Form Control or ActiveX
' checkbox SHAPE from the original hand-built file, or something riding on
' the number format. Nothing here changes anything -- it reports what is
' really on the first black spacer row it finds, so the next fix is aimed
' rather than guessed.
'
' A spacer row is identified the way the rest of this module does: inside
' the data block, with a blank Job ID (column C).
' =============================================================================
Public Sub DiagnoseTracieCheckboxes()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(OUT_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "No '" & OUT_SHEET & "' sheet in this workbook.", vbExclamation, "Diagnose Checkboxes"
        Exit Sub
    End If

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 3).End(xlUp).Row

    ' First spacer row: blank Job ID, black fill, inside the data block.
    Dim r As Long, sepRow As Long
    For r = 2 To lastRow
        If Trim(CStr(ws.Cells(r, 3).Value)) = "" Then
            sepRow = r
            Exit For
        End If
    Next r

    Dim rpt As String
    rpt = "Tracie checkbox diagnosis" & vbCrLf & _
          "Module " & MODULE_VERSION & vbCrLf & vbCrLf & _
          "Last job row (col C): " & lastRow & vbCrLf

    If sepRow = 0 Then
        rpt = rpt & vbCrLf & "No spacer row found (every row 2.." & lastRow & _
              " has a Job ID). If you can see a black row, tell me its row number."
        MsgBox rpt, vbInformation, "Diagnose Checkboxes"
        Exit Sub
    End If

    rpt = rpt & "First spacer row: " & sepRow & vbCrLf & vbCrLf

    ' --- what is in / on the Rdy cell of that row -----------------------
    Dim cel As Range
    Set cel = ws.Cells(sepRow, 9)
    rpt = rpt & "Rdy cell " & cel.Address(False, False) & vbCrLf & _
          "  value      : " & IIf(IsEmpty(cel.Value), "(empty)", CStr(cel.Value)) & vbCrLf & _
          "  IsEmpty    : " & IsEmpty(cel.Value) & vbCrLf & _
          "  NumberFmt  : " & cel.NumberFormat & vbCrLf & _
          "  Interior   : " & cel.Interior.Color & " (black = 0)" & vbCrLf

    Dim hasVal As String: hasVal = "none"
    On Error Resume Next
    Err.Clear
    hasVal = "type " & CStr(cel.Validation.Type)
    If Err.Number <> 0 Then hasVal = "none"
    Err.Clear
    On Error GoTo 0
    rpt = rpt & "  Validation : " & hasVal & vbCrLf

    Dim hasCtl As String: hasCtl = "not reachable"
    Dim cc As Object, ccType As Variant
    On Error Resume Next
    Err.Clear
    Set cc = cel.CellControl
    If Err.Number <> 0 Then
        hasCtl = "not reachable (err " & Err.Number & ")"
    ElseIf cc Is Nothing Then
        hasCtl = "none"
    Else
        hasCtl = "PRESENT"
        ccType = Empty
        Err.Clear
        ccType = cc.Type
        If Err.Number = 0 Then
            hasCtl = hasCtl & ", Type " & CStr(ccType) & " (0 = none)"
        Else
            hasCtl = hasCtl & ", Type unreadable (err " & Err.Number & ")"
        End If
    End If
    Err.Clear
    On Error GoTo 0
    rpt = rpt & "  CellControl: " & hasCtl & vbCrLf

    ' --- any shape sitting on that row ---------------------------------
    Dim shp As Shape, shpRow As Long, nShapes As Long
    rpt = rpt & vbCrLf & "Shapes anchored on row " & sepRow & ":" & vbCrLf
    On Error Resume Next
    For Each shp In ws.Shapes
        shpRow = 0
        shpRow = shp.TopLeftCell.Row
        If shpRow = sepRow Then
            nShapes = nShapes + 1
            rpt = rpt & "  " & shp.name & " | Type " & shp.Type & _
                  IIf(shp.Type = msoFormControl, " (FormControl " & shp.FormControlType & ")", "") & _
                  " | col " & shp.TopLeftCell.Column & vbCrLf
        End If
    Next shp
    Err.Clear
    On Error GoTo 0
    If nShapes = 0 Then rpt = rpt & "  (none)" & vbCrLf
    rpt = rpt & "  Total shapes on sheet: " & ws.Shapes.Count & vbCrLf

    rpt = rpt & vbCrLf & "Send this whole box back and the fix can be aimed at " & _
          "whichever of these is actually holding the checkbox."

    MsgBox rpt, vbInformation, "Diagnose Checkboxes"
End Sub

' Reads "<key><digits>" out of the button's AlternativeText tag, or -1.
Private Function ReadTagNum(ByVal tag As String, ByVal key As String) As Long
    ReadTagNum = -1
    Dim p As Long: p = InStr(tag, key)
    If p = 0 Then Exit Function
    Dim rest As String, i As Long, digits As String
    rest = Mid(tag, p + Len(key))
    For i = 1 To Len(rest)
        If Mid(rest, i, 1) Like "#" Then
            digits = digits & Mid(rest, i, 1)
        Else
            Exit For
        End If
    Next i
    If digits <> "" Then ReadTagNum = CLng(digits)
End Function

Private Sub EnsureUpdateButton()
    Dim ws As Worksheet, shp As Shape
    Dim foundOnPrefs As Boolean: foundOnPrefs = False
    For Each ws In ThisWorkbook.Worksheets
        Set shp = Nothing
        On Error Resume Next
        Set shp = ws.Shapes(BTN_NAME)
        On Error GoTo 0
        If Not shp Is Nothing Then
            If ws.name <> "Preferences" Then
                ' Anywhere but Preferences: user-owned, only re-wire the macro.
                shp.OnAction = "UpdatePersoData"
                Exit Sub
            End If
            foundOnPrefs = True
            Exit For
        End If
    Next ws

    ' Standing relocation (2026-08.12): the button belongs on Tracie, but a
    ' pre-v5 run created it on Preferences (and the find-anywhere rule kept
    ' it there). Move the SHAPE itself -- cut/paste, not delete/recreate --
    ' so any custom styling the user gave it survives the move. If the
    ' move fails for any reason, fall through to fresh default creation.
    If foundOnPrefs Then
        Dim wsOut As Worksheet
        On Error Resume Next
        Set wsOut = ThisWorkbook.Sheets(OUT_SHEET)
        On Error GoTo 0
        If wsOut Is Nothing Then
            shp.OnAction = "UpdatePersoData"
            Exit Sub
        End If

        Dim prevActive As Object
        On Error Resume Next
        Set prevActive = ThisWorkbook.ActiveSheet
        On Error GoTo 0

        On Error GoTo MoveFailed
        shp.Cut
        wsOut.Activate
        wsOut.Paste
        Dim moved As Shape
        Set moved = wsOut.Shapes(wsOut.Shapes.Count)
        moved.name = BTN_NAME
        moved.Left = wsOut.Cells(1, 12).Left + 4
        moved.Top = 4
        moved.OnAction = "UpdatePersoData"
        On Error Resume Next
        If Not prevActive Is Nothing Then prevActive.Activate
        On Error GoTo 0
        Exit Sub

MoveFailed:
        ' Clean up any half-moved leftover holding the reserved name, then
        ' fall through to fresh creation below.
        On Error Resume Next
        ThisWorkbook.Sheets("Preferences").Shapes(BTN_NAME).Delete
        Err.Clear
        On Error GoTo 0
    End If

    ' Created on TRACIE (user request 2026-08.12), parked above the
    ' capacity table (row 1, col L area -- was col M while Cards was still
    ' present alongside To Perso). Still user-owned: drag it anywhere --
    ' including another sheet -- and it stays put.
    Dim wsHome As Worksheet
    On Error Resume Next
    Set wsHome = ThisWorkbook.Sheets(OUT_SHEET)
    On Error GoTo 0
    If wsHome Is Nothing Then Exit Sub

    ' Created at the two-line height (v3). A button made at the old 34pt is
    ' grown once by RefreshUpdateButton, and only while it is still untouched.
    Set shp = wsHome.Shapes.AddShape(msoShapeRoundedRectangle, _
        wsHome.Cells(1, 12).Left + 4, 4, BTN_DEF_W, BTN_TWO_LINE_H)

    shp.name = BTN_NAME
    shp.OnAction = "UpdatePersoData"
    With shp
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = BTN_FILL_BASE
        With .Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(0, 0, 0)
            .Weight = 1.5
        End With
        With .TextFrame2.TextRange
            .Text = "Update Data" & vbCr & "not pulled yet"
            .Font.Fill.ForeColor.RGB = PickTextColor(BTN_FILL_BASE)
            .ParagraphFormat.Alignment = msoAlignCenter
        End With
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
    End With
    On Error Resume Next
    With shp.TextFrame2.TextRange.Paragraphs(1)
        .Font.Size = 14
        .Font.Bold = msoTrue
    End With
    With shp.TextFrame2.TextRange.Paragraphs(2)
        .Font.Size = 9
        .Font.Bold = msoFalse
    End With
    shp.AlternativeText = BTN_TAG & "base=" & BTN_FILL_BASE & "|last=" & BTN_FILL_BASE
    On Error GoTo 0
End Sub

' =============================================================================
' UTILITIES -- shared shapes with the Press module (duplicated here since
' these are separate workbooks; see the Press module for provenance notes).
' =============================================================================
' Collection has no Exists, and probing for a missing key raises -- so the
' probe is wrapped rather than the caller having to guard every lookup.
Private Function CollHasKey(ByRef c As Collection, ByVal k As String) As Boolean
    Dim v As Variant
    On Error GoTo NoKey
    v = c(k)
    CollHasKey = True
    Exit Function
NoKey:
    CollHasKey = False
End Function

Private Function CanonText(ByVal s As String) As String
    CanonText = UCase(Replace(Replace(Replace(Trim(s), " ", ""), vbTab, ""), Chr(160), ""))
End Function

Private Sub ParseJobIdBaseSuffix(ByVal jobId As String, ByRef baseOut As String, ByRef suffixOut As Long)
    Dim s As String: s = Trim(jobId)
    baseOut = s
    suffixOut = -1
    Dim hyphenPos As Long: hyphenPos = InStrRev(s, "-")
    If hyphenPos > 0 And hyphenPos < Len(s) Then
        Dim tail As String: tail = Mid(s, hyphenPos + 1)
        If IsAllDigits(tail) Then
            baseOut = Left(s, hyphenPos - 1)
            suffixOut = CLng(tail)
        End If
    End If
End Sub

Private Function IsAllDigits(ByVal s As String) As Boolean
    Dim i As Long
    If Len(s) = 0 Then Exit Function
    For i = 1 To Len(s)
        If Mid(s, i, 1) < "0" Or Mid(s, i, 1) > "9" Then Exit Function
    Next i
    IsAllDigits = True
End Function

'' True only for an actual True or a "TRUE"/"True" string -- anything else
' (empty, error, text) is False. Used when restoring Rdy ticks.
Private Function SafeBool(ByVal v As Variant) As Boolean
    SafeBool = False
    On Error Resume Next
    If VarType(v) = vbBoolean Then
        SafeBool = v
    ElseIf VarType(v) = vbString Then
        SafeBool = (UCase(Trim(CStr(v))) = "TRUE")
    End If
    On Error GoTo 0
End Function

Private Function GetLabeledValueFromSheet(ByRef ws As Worksheet, ByVal label As String) As Variant
    Dim r As Long
    For r = 1 To 100
        If Trim(CStr(ws.Cells(r, 1).Value)) = label Then
            GetLabeledValueFromSheet = ws.Cells(r, 2).Value
            Exit Function
        End If
    Next r
    GetLabeledValueFromSheet = Empty
End Function

Private Function GetPrefValue(ByVal label As String) As Variant
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Preferences")
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    GetPrefValue = GetLabeledValueFromSheet(ws, label)
End Function

' Same fixed 3-state vocabulary as Press: "Ready" / "Not Ready" / "Error",
' detail joined with " - ".
Private Sub SetPersoStatus(ByVal statusWord As String, ByVal detail As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Preferences")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    Dim r As Long
    For r = 1 To 100
        If Trim(CStr(ws.Cells(r, 1).Value)) = "Perso Status" Then
            Dim sw As Range
            Set sw = ws.Cells(r, 2)
            sw.Value = statusWord & IIf(detail <> "", " - " & detail, "")
            ' Same at-a-glance styling as Press: Error runs get Excel's
            ' built-in "Bad" style, everything else reverts to the normal
            ' light-green look so a stale red never lingers.
            On Error Resume Next
            If statusWord = "Error" Then
                sw.Style = "Bad"
            Else
                sw.Style = "Normal"
                sw.Interior.Color = RGB(226, 239, 218)
                sw.Font.Color = RGB(0, 97, 0)
            End If
            On Error GoTo 0
            Exit Sub
        End If
    Next r
End Sub

' Pulls the Monarch export's own timestamp out of Hub's status detail --
' same parse as the Press module: the detail looks like
' "Full Monarch Pull_8.11.xls (modified 2026-08-11 12:22) | updated ...";
' the "(modified ...)" portion is the export's own timestamp.
Private Function ExtractHubExportTimestamp(ByVal hubDetail As String) As String
    ' Every phrasing the Hub has used, best-quality first. Before v2 this
    ' searched only for "(modified " -- the format this Hub happens to emit,
    ' but not the one a SQL/Power-Query Hub would, so the parse would have
    ' begun failing silently the day the Hub was replaced, and the fallback
    ' text blamed the Hub for being old at the exact moment it was updated.
    Dim markers As Variant
    markers = Array("(modified ", "as saved ", "query refreshed ", _
                    "refreshed ", "pulled ", "updated ")
    Dim i As Long, p As Long, tok As String
    For i = LBound(markers) To UBound(markers)
        p = InStr(1, hubDetail, CStr(markers(i)), vbTextCompare)
        If p > 0 Then
            tok = ExtractStampToken(Mid(hubDetail, p + Len(CStr(markers(i)))))
            If tok <> "" Then
                ExtractHubExportTimestamp = tok
                Exit Function
            End If
        End If
    Next i
End Function

' Reads a leading "yyyy-mm-dd hh:nn" or returns "". Hand-checked rather than
' trusting CDate: the old code took everything up to the next ")", which on a
' different Hub's status string returns "2026-08-19 07:42, 6,620 rows" -- a
' plausible-looking value that is not a timestamp.
Private Function ExtractStampToken(ByVal s As String) As String
    s = Trim(s)
    If Len(s) < 16 Then Exit Function
    Dim t As String: t = Left(s, 16)
    If Not IsAllDigits(Mid(t, 1, 4)) Then Exit Function
    If Mid(t, 5, 1) <> "-" Then Exit Function
    If Not IsAllDigits(Mid(t, 6, 2)) Then Exit Function
    If Mid(t, 8, 1) <> "-" Then Exit Function
    If Not IsAllDigits(Mid(t, 9, 2)) Then Exit Function
    If Mid(t, 11, 1) <> " " Then Exit Function
    If Not IsAllDigits(Mid(t, 12, 2)) Then Exit Function
    If Mid(t, 14, 1) <> ":" Then Exit Function
    If Not IsAllDigits(Mid(t, 15, 2)) Then Exit Function
    ExtractStampToken = t
End Function

' "yyyy-mm-dd hh:nn" -> Date, or 0. DateSerial/TimeSerial rather than CDate:
' CDate reads that string through the machine's locale and on a dd/mm
' workstation would throw or silently return the wrong day.
Private Function ParseHubTimestamp(ByVal ts As String) As Date
    ParseHubTimestamp = 0
    If Len(ts) < 16 Then Exit Function
    Dim y As Long, mo As Long, dy As Long, hh As Long, mi As Long
    On Error GoTo BadStamp
    y = CLng(Mid(ts, 1, 4))
    mo = CLng(Mid(ts, 6, 2))
    dy = CLng(Mid(ts, 9, 2))
    hh = CLng(Mid(ts, 12, 2))
    mi = CLng(Mid(ts, 15, 2))
    If mo < 1 Or mo > 12 Or dy < 1 Or dy > 31 Then Exit Function
    If hh < 0 Or hh > 23 Or mi < 0 Or mi > 59 Then Exit Function
    ParseHubTimestamp = DateSerial(y, mo, dy) + TimeSerial(hh, mi, 0)
    Exit Function
BadStamp:
    ParseHubTimestamp = 0
End Function

' Whatever the setup phase quietly could not do. Empty when all is well, so a
' clean run reads clean.
Private Function SetupWarningBlock() As String
    If Len(gSetupWarnings) = 0 Then Exit Function
    SetupWarningBlock = vbCrLf & vbCrLf & _
        "Some Preferences features could not be built:" & gSetupWarnings
End Function

Private Sub ShowPreferencesWhenDone()
    On Error Resume Next
    ThisWorkbook.Sheets("Preferences").Activate
    On Error GoTo 0
End Sub

Private Function GetOrCreateSheet(ByVal rawName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(rawName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.name = rawName
    End If
    Set GetOrCreateSheet = ws
End Function

' Tab order: Tracie FIRST (user request 2026-08.12), then Preferences.
Private Sub EnforcePersoWorksheetOrder()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(OUT_SHEET)
    If Not ws Is Nothing Then ws.Move Before:=ThisWorkbook.Sheets(1)
    Set ws = Nothing
    Set ws = ThisWorkbook.Sheets("Preferences")
    If Not ws Is Nothing Then
        If ThisWorkbook.Sheets.Count >= 2 Then ws.Move Before:=ThisWorkbook.Sheets(2)
    End If
    On Error GoTo 0
End Sub

-------------------------------------------------------------------------------
