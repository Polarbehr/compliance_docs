Option Explicit

' #############################################################################
'
'                        P L I   P R E S S  --  Module4
'
'   VERSION : v49
'   BUILT   : 2026-09-03
'
'   The version is also held in the MODULE_VERSION constant a few lines down
'   and printed at the end of every Update Data / Refresh Data run, so the
'   build a workbook is actually running can be read without opening the VBA
'   editor. That mattered on 2026-08-24, when machine tabs turned out to have
'   been produced by a build four versions older than the one in the file.
'
'   ---------------------------------------------------------------------
'   RECENT HISTORY
'   ---------------------------------------------------------------------
'   v49  Hub started shipping a 12th CleanedData column, To Perso. It's now
'        column A on every machine tab (before Ship Date), and the
'        scheduler uses it as a job's EFFECTIVE deadline whenever Hub has a
'        date there -- queue order, the same-date backfill, and the
'        late/on-time check all key off it instead of Ship Date for that
'        job. Ship Date keeps its own column and always shows the job's
'        actual ship date -- only the scheduling math substitutes. See
'        EffectiveScheduleDate. Every COL_OUT_* constant shifted by one to
'        make room; nothing that read them by name needed touching.
'   v48  The tab build is instrumented properly -- one gStep covered 334
'        lines and reported "clearing the tab" for a failure anywhere in the
'        write loop. The AutoFilter and the merges now come off BEFORE the
'        contents are cleared.
'   v47  AutoFilter no longer covers the merged Total Qty column -- sorting
'        from a filter dropdown over those merges raised 1004 "We can't do
'        that to a merged cell". EnsureRemovedJobsSheet no longer exits before
'        writing the parts added after the sheet first existed.
'   v46  Rules row 8b split across continuations -- at 1,062 characters the
'        single-line version broke VBA's 1023-char physical line limit. A
'        check_linelen.py gate now runs on every build.
'   v45  The RFID tick is now a STANDING FILTER, not a one-shot removal. It
'        survives Refresh Data, and unticking it brings those jobs back. The
'        ticked types are stored on the Removed Jobs sheet (columns I:J) so
'        they outlive the tab rebuild that wipes the checkbox itself.
'   v44  RFID type list in column L with a bulk-remove tick in column M.
'        Ticking a type marks every job of that technology for removal --
'        live via Workbook_SheetChange, and again at Refresh Data so the
'        result never depends on the event firing.
'   v43  Version banner + MODULE_VERSION, surfaced in the finish messages.
'   v42  Gap filling restored: a day no longer ends short when a job of the
'        SAME ship date would fit. v41 wasted 261 such chances on the live
'        queue. 46 days / 94.1% fill, 0 splits, 0 segments out of order.
'   v41  Split rule rewritten: ONLY a job whose total exceeds a full day is
'        ever cut, and it starts on its own fresh day.
'   v40  EnsureRosterValidation bounded to the roster block -- the first cut
'        ran E12:E35 through nine merged banners.
'   v39  Location Filter reads BOTH columns and honours the tick boxes;
'        exact + wildcard matching; "|" means OR. Production Cutoff wired up.
'        Roster dropdowns built in code. Fill-Ahead retired.
'   v38  "mm/dd/yyyy hh:nn" -> "hh:mm". "nn" is a VBA Format() token and is
'        not valid in an Excel cell number format.
'   v37  Every NumberFormat routed through SetFmt: formatting can no longer
'        abort a build, and a refusal reports itself.
'   v36  Cells.Clear does not unmerge or drop the AutoFilter -- both now
'        explicit. Step tracker (gStep) added.
'   v35  Job removal: Remove checkbox column and the Removed Jobs sheet.
'   ---------------------------------------------------------------------
'
' #############################################################################

' =============================================================================
' PLI PRESS WORKBOOK -- Module4
' Satellite workbook: pulls CleanedData (values only) from the Hub workbook
' on demand via the Update Data button, filters it per machine listed on the
' Preferences "Machine Roster" table, and builds one tab per machine with a
' calculated Production Day column (a scheduling projection, NOT the job's
' own promised date -- see ComputeProductionDays for the exact algorithm).
'
' Extensibility: to add a machine/tab, add a row to Preferences' Machine
' Roster table (Work Center Match keyword | Rule Type | Daily Threshold) --
' no VBA change needed. See BuildAllMachineTabs / ReadRosterTable.
'
' Does NOT pull Parameters from Hub -- Press has no Capacity%/rate math, and
' Hub's CleanedData Effective QTY/Job Hours columns are already resolved to
' plain values by the time Hub sets its status to READY, so Press never
' needs Parameters itself.
'
' First-run bootstrap: same limitation as Hub -- the Update Data button does
' not exist until the first run creates it, so the very first run must be
' started via Alt+F8 -> UpdatePressData -> Run.
' =============================================================================

' -----------------------------------------------------------------------------
' Column positions within Hub's CleanedData sheet: A Ship Date,
' B Promised Date, C Job ID, D Customer, E Description, F QTY,
' G Location Date, H Location, I Work Center, J Week Start, K RFID Type,
' L To Perso.
'
' 2026-08.17: column K is RFID TYPE. It used to be Effective QTY, which the
' Hub stopped shipping on 2026-08.12 when the rate math moved to the AVL
' Production Dashboard. SafeEffectiveQty was still reading K and only
' escaped returning an RFID token as a quantity because IsNumeric happens
' to reject text -- that latent trap is now removed.
'
' 2026-09.03: column L is TO PERSO -- Monarch's own field, Hub just carries
' it through. Blank on a HubCache written by an older Hub build (fewer than
' 12 columns) or on a job that hasn't reached that stage; every read of it
' below is guarded with UBound(rawData, 2) >= COL_TOPERSO for that reason.
' See EffectiveScheduleDate for how it changes scheduling.
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
Private Const COL_WEEKSTART As Long = 10
Private Const COL_RFID As Long = 11
Private Const COL_TOPERSO As Long = 12

' -----------------------------------------------------------------------------
' Column positions on the MACHINE TABS (the sheets Press writes), as of the
' 2026-09.03 To Perso insertion: A To Perso, B Ship Date, C Job ID,
' D Customer Name, E Job Description, F Qty, G RFID, H Last Location,
' I Work Center, J Production Day, K Total Qty (merged per day).
' -----------------------------------------------------------------------------
Private Const COL_OUT_TOPERSO As Long = 1
Private Const COL_OUT_SHIP As Long = 2
Private Const COL_OUT_JOBID As Long = 3
Private Const COL_OUT_CUSTOMER As Long = 4
Private Const COL_OUT_DESC As Long = 5
Private Const COL_OUT_QTY As Long = 6
Private Const COL_OUT_RFID As Long = 7
Private Const COL_OUT_LOCATION As Long = 8
Private Const COL_OUT_WORKCENTER As Long = 9
Private Const COL_OUT_PRODDAY As Long = 10
Private Const COL_OUT_TOTALQTY As Long = 11

' 2026-08.19: carries the Remove checkbox (column L as of the 2026-09.03
' To Perso insertion; was K before it). Verified free -- the machine tabs
' occupy only 1-11 plus the hidden state marker at 15.
Private Const COL_OUT_REMOVE As Long = 12

' 2026-08.25: carry the per-technology bulk-remove list (columns M and N as
' of the 2026-09.03 To Perso insertion; were L and M before it). Verified
' free -- the tabs use 1-14 plus the hidden state marker at 15.
Private Const COL_OUT_RFIDLIST As Long = 13
Private Const COL_OUT_RFIDTICK As Long = 14

' Printed at the end of every run. Keep in step with the banner at the top.
Private Const MODULE_VERSION As String = "v49"

Private Const REMOVED_SHEET As String = "Removed Jobs"
Private Const REMOVED_HEADER_ROW As Long = 4
Private Const REMOVED_FIRST_ROW As Long = 5

' Columns I:J of the Removed Jobs sheet hold the STANDING RFID type filters --
' one row per machine + technology. Kept clear of the job list in A:G so the
' two can be read and cleared independently.
Private Const REMOVED_RFID_MACHINE As Long = 9
Private Const REMOVED_RFID_TYPE As Long = 10

' How far down the Removed Jobs sheet the column number formats are applied.
' A bounded range rather than whole columns on purpose: whole-column formatting
' would intersect the merged A1:G1 title banner, and setting a format on part
' of a merged area is one of the two things that broke v35.
Private Const REMOVED_FORMAT_LAST_ROW As Long = 5000

' =============================================================================
' BUILD DIAGNOSTICS (2026-08.24)
'
' v35 failed twice with errors that named a property but not a place -- "We
' can't do that to a merged cell", then "Unable to set the NumberFormat
' property of the Range class" -- and pinning them down cost a whole session.
' gStep is set as the build walks through its stages so both entry points can
' say WHERE they died, not just what Excel objected to. It costs one string
' assignment per stage and it means the next failure diagnoses itself.
' =============================================================================
Private gStep As String
Private gCheckboxTried As Boolean
Private gCheckboxOK As Boolean

' Accumulates every NumberFormat call Excel refused during this run. See SetFmt.
Private gFmtWarn As String

' Row-format state indices into the ReadRowFormats() table.
Private Const FMT_NORMAL As Long = 1
Private Const FMT_SHIPDAY As Long = 2
Private Const FMT_LATE As Long = 3
Private Const FMT_SPLIT As Long = 4
Private Const FMT_SPLIT_LATE As Long = 5

' Hidden per-row state marker. BuildOneMachineTab stamps each written
' job row with its FMT_* state here and hides the column; the Refresh
' Colours button then re-applies the Row Formatting swatches straight
' from these markers, with no Hub pull and no re-scheduling. Column 15 --
' immediately after RFID Tick (14) as of the 2026-09.03 To Perso
' insertion, but still its own column, so it still can't collide with
' the table.
Private Const COL_STATE_MARK As Long = 15

' Very-hidden sheet holding the LAST dataset pulled from Hub. Update Data
' writes it; Refresh Tabs replays the identical build from it, so changing
' a scheduling setting or a colour rebuilds every tab through exactly the
' same code path as a real pull -- just without touching Hub. Caching the
' source beats reconstructing jobs from the finished tabs: split rows show
' only the qty run that day, so the original order quantities could not be
' recovered reliably from the sheet.
Private Const CACHE_SHEET As String = "HubCache"

' =============================================================================
' EFFECTIVE SCHEDULE DATE (2026-09.03) -- the user's rule: "If there is a
' date in To Perso use that date instead of the Ship Date for filling in
' the days capacity." This is the one place that rule is implemented; every
' scheduling touchpoint (the dated/undated split, the Ship-Date-then-Job-ID
' sort that sets queue order, the QTY scheduler's same-date backfill, and
' the late/on-time check) calls this instead of reading COL_SHIP directly.
'
' Display never calls this -- WritePressRow writes the RAW Ship Date and
' the RAW To Perso to their own columns unchanged. Only the scheduling math
' substitutes.
' =============================================================================
Private Function EffectiveScheduleDate(ByRef rawData As Variant, ByVal rowIdx As Long) As Variant
    If UBound(rawData, 2) >= COL_TOPERSO Then
        Dim tp As Variant
        tp = rawData(rowIdx, COL_TOPERSO)
        If IsDate(tp) And Not IsEmpty(tp) Then
            EffectiveScheduleDate = tp
            Exit Function
        End If
    End If
    EffectiveScheduleDate = rawData(rowIdx, COL_SHIP)
End Function

' =============================================================================
' ENTRY POINT -- wired to the "Update Data" button on Preferences.
' =============================================================================
Public Sub UpdatePressData()
    Dim t0 As Single: t0 = Timer
    gStep = "starting up"
    gCheckboxTried = False
    gCheckboxOK = True
    gFmtWarn = ""
    EnsurePreferencesSheet
    EnsureUpdateButton
    EnsureRefreshButton
    BuildRulesSheet
    BuildOnboardingSheet

    Dim hubPath As String
    hubPath = Trim(CStr(GetPrefValue("Hub Workbook Path")))
    If hubPath = "" Then
        MsgBox "Set 'Hub Workbook Path' on Preferences to the PLI Hub.xlsm file first.", vbExclamation, "Update Data"
        ShowPreferencesWhenDone
        Exit Sub
    End If

    SetPressStatus "Not Ready", "pull in progress"

    On Error GoTo PullFailed

    Dim rawData As Variant
    Dim hubStatusWord As String, hubStatusDetail As String
    If Not PullCleanedDataFromHub(hubPath, rawData, hubStatusWord, hubStatusDetail) Then
        GoTo CleanExit  ' PullCleanedDataFromHub already showed the reason
    End If

    Dim roster As Variant
    roster = ReadRosterTable()
    If IsEmpty(roster) Then
        SetPressStatus "Error", "Machine Roster table on Preferences is empty -- nothing to build"
        MsgBox "The Machine Roster table on Preferences has no rows. Add at least one " & _
               "(Work Center Match | Rule Type | Daily Threshold) and try again.", vbExclamation
        GoTo CleanExit
    End If

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    SaveHubCache rawData

    ' Update Data is the clean slate: yesterday's removals are wiped so
    ' every job comes back with an untouched, unticked box.
    ClearRemovedJobs

    BuildAllMachineTabs rawData, roster

    EnforcePressWorksheetOrder roster

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True

    Dim exportTs As String
    exportTs = ExtractHubExportTimestamp(hubStatusDetail)
    If exportTs = "" Then exportTs = "unknown (older Hub build -- update Hub first)"
    SetPressStatus "Ready", "Updated At: " & exportTs
    ShowPreferencesWhenDone

    MsgBox "Press data updated in " & Format(Timer - t0, "0.00") & " seconds." & vbCrLf & vbCrLf & _
           "Hub status at pull time: " & hubStatusWord & IIf(hubStatusDetail <> "", " (" & hubStatusDetail & ")", "") & vbCrLf & _
           "Machines: " & (UBound(roster, 1)) & " tab(s) built." & vbCrLf & _
           "Press module " & MODULE_VERSION & "." & _
           CheckboxWarning() & FormatWarning(), vbInformation
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
    SetPressStatus "Error", "Build failed at '" & gStep & "' (" & errNum & ": " & errMsg & ") at " & Format(Now, "yyyy-mm-dd hh:nn")
    ShowPreferencesWhenDone
    MsgBox "Update Data failed while " & gStep & "." & vbCrLf & vbCrLf & _
           "Excel reported: " & errNum & " - " & errMsg, vbCritical, "Press Update Data Failed"
End Sub

' =============================================================================
' HUB PULL -- opens Hub read-only, checks Hub Status BEFORE trusting its
' CleanedData (REFRESHING = a build is mid-flight, ERROR = last build failed
' and CleanedData may be incomplete -- both block the pull rather than risk
' reading bad data), reads CleanedData as plain values, closes without saving.
' =============================================================================
Private Function PullCleanedDataFromHub(ByVal hubPath As String, ByRef rawData As Variant, _
        ByRef hubStatusWord As String, ByRef hubStatusDetail As String) As Boolean
    PullCleanedDataFromHub = False
    hubStatusWord = "": hubStatusDetail = ""

    If Dir(hubPath) = "" Then
        SetPressStatus "Error", "Hub file not found at " & hubPath
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
        SetPressStatus "Error", "Could not open Hub workbook (locked or blocked)"
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
        SetPressStatus "Not Ready", "Hub is REFRESHING -- retry in a moment"
        MsgBox "Hub is currently refreshing (someone clicked its Update Data and it hasn't finished). " & _
               "Wait a moment and click Update again.", vbInformation, "Hub Is Refreshing"
        Exit Function
    ElseIf UCase(hubStatusWord) = "ERROR" Then
        wbHub.Close SaveChanges:=False
        Application.DisplayAlerts = True
        Application.ScreenUpdating = prevSU
        SetPressStatus "Error", "Hub's last build failed: " & hubStatusDetail
        MsgBox "The Hub workbook's last Update Data run failed:" & vbCrLf & hubStatusDetail & vbCrLf & vbCrLf & _
               "Fix that in the Hub workbook and re-run its Update Data before pulling here -- " & _
               "CleanedData may be incomplete.", vbCritical, "Hub Reported An Error"
        Exit Function
    ElseIf UCase(hubStatusWord) <> "READY" Then
        ' Unknown/blank status -- older Hub build or first-ever run. Warn but
        ' don't hard-block, since this shouldn't brick an otherwise-working pull.
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
        SetPressStatus "Error", "Hub workbook has no CleanedData sheet"
        MsgBox "The Hub workbook does not have a CleanedData sheet.", vbCritical, "Update Data"
        Exit Function
    End If

    rawData = cdSheet.UsedRange.Value
    wbHub.Close SaveChanges:=False
    Application.DisplayAlerts = True
    Application.ScreenUpdating = prevSU

    If IsEmpty(rawData) Then
        SetPressStatus "Error", "Hub CleanedData was empty"
        MsgBox "Hub's CleanedData sheet has no data.", vbExclamation, "Update Data"
        Exit Function
    End If

    PullCleanedDataFromHub = True
End Function

' =============================================================================
' MACHINE ROSTER -- read from Preferences, anchored by the "Machine Roster"
' label (not a fixed row), so reorganizing Preferences never breaks this.
' =============================================================================
Private Function ReadRosterTable() As Variant
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

    Dim headerRow As Long: headerRow = anchorRow + 2  ' anchor, instructions, THEN header
    Dim dataStartRow As Long: dataStartRow = headerRow + 1

    Dim n As Long: n = 0
    Dim results() As Variant
    ReDim results(1 To 200, 1 To 4)

    r = dataStartRow
    Do While Trim(CStr(ws.Cells(r, 1).Value)) <> "" And r < dataStartRow + 200
        n = n + 1
        results(n, 1) = Trim(CStr(ws.Cells(r, 1).Value))          ' Work Center Match keyword
        results(n, 2) = UCase(Trim(CStr(ws.Cells(r, 2).Value)))   ' "COUNT" or "QTY"
        results(n, 3) = Val(ws.Cells(r, 3).Value)                 ' Daily Threshold
        ' Late Rule (2026-08.13): "AFTER" = a job is late only AFTER its
        ' Ship Date (running ON the ship date is fine, and flags as the
        ' amber "ships today" state). "ON/AFTER" = the legacy rule, late
        ' the moment it runs on its own ship date. Blank defaults to
        ' AFTER; MigrateRosterLateRule seeds the column on older sheets.
        results(n, 4) = UCase(Trim(CStr(ws.Cells(r, 4).Value)))
        If results(n, 4) = "" Then results(n, 4) = "AFTER"
        If InStr(results(n, 4), "ON") > 0 Then
            results(n, 4) = "ON/AFTER"
        Else
            results(n, 4) = "AFTER"
        End If
        r = r + 1
    Loop

    If n = 0 Then Exit Function

    Dim finalArr() As Variant
    ReDim finalArr(1 To n, 1 To 4)
    Dim i As Long, j As Long
    For i = 1 To n
        For j = 1 To 4
            finalArr(i, j) = results(i, j)
        Next j
    Next i
    ReadRosterTable = finalArr
End Function

' =============================================================================
' LOCATION FILTER -- read from Preferences, anchored by the "Location
' Filter" label. A single-column list of keywords; a job's Location must
' CONTAIN at least one of them (case-insensitive/trimmed, same "contains"
' semantics as the Machine Roster's Work Center match) to appear on any
' machine tab. Added 2026-08 on top of the Work Center match, not instead
' of it -- narrows the queue to jobs at a relevant workflow stage. Returns
' Empty if the list doesn't exist or has no rows, which callers treat as
' "no location restriction" rather than "match nothing."
' =============================================================================
' =============================================================================
' LOCATION FILTER (rewritten 2026-08-24)
'
' The sheet holds TWO side-by-side filter columns, each with its own tick box
' in the cell immediately to its right: locations in A ticked from B, and
' locations in C ticked from D.
'
' What this replaced, and why it mattered: the previous reader looked at
' column A only and stopped dead at the first blank cell, so column C was
' invisible and every tick box was decorative. Measured against the live
' HubCache on 2026-08-24, that was the difference between 665 Dicetrax job
' rows and 119 -- 82% of the queue, 1.39M of 1.51M pieces, 42 of 46
' production days -- and it was one Update Data away from happening.
'
' Rules, confirmed with the user 2026-08-24:
'   * A BLANK tick counts as ON. A location typed into a spare slot takes
'     effect straight away, which is what the sheet's own instructions
'     promise. Only an explicit FALSE / 0 / NO switches one off.
'   * "|" inside an entry separates ALTERNATIVES. "TI | To Imposition" means
'     TI *or* To Imposition; each side is trimmed and matched on its own.
'   * Matching is EXACT with * and ? wildcards -- see LocationMatchesFilter.
'
' Scanning stops at the next section label rather than at a blank row,
' because the spare slots in the middle of the block ARE blank and the block
' is only two rows clear of "Row Formatting" underneath it.
' =============================================================================
Private Function ReadLocationFilterList() As Variant
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Preferences")
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    Dim anchorRow As Long: anchorRow = 0
    Dim r As Long
    For r = 1 To 100
        If Trim(CStr(ws.Cells(r, 1).Value)) = "Location Filter" Then
            anchorRow = r
            Exit For
        End If
    Next r
    If anchorRow = 0 Then Exit Function

    Dim dataStartRow As Long: dataStartRow = anchorRow + 2  ' anchor, instructions, THEN data

    Dim n As Long: n = 0
    Dim results() As String
    ReDim results(1 To 400)

    ' Column pairs: text column, and the tick column immediately right of it.
    Dim pair As Long, textCol As Long, tickCol As Long
    Dim raw As String, part As Variant, parts As Variant

    For r = dataStartRow To dataStartRow + 199
        If IsPrefSectionLabel(Trim(CStr(ws.Cells(r, 1).Value))) Then Exit For
        For pair = 0 To 1
            textCol = 1 + pair * 2
            tickCol = textCol + 1
            raw = Trim(CStr(ws.Cells(r, textCol).Value))
            If raw <> "" Then
                If FilterTickIsOn(ws.Cells(r, tickCol).Value) Then
                    parts = Split(raw, "|")
                    For Each part In parts
                        If Trim(CStr(part)) <> "" Then
                            n = n + 1
                            If n > 400 Then Exit For
                            results(n) = EscapeLikePattern(UCase(Trim(CStr(part))))
                        End If
                    Next part
                End If
            End If
        Next pair
    Next r

    If n = 0 Then Exit Function

    Dim finalArr() As String
    ReDim finalArr(1 To n)
    Dim i As Long
    For i = 1 To n
        finalArr(i) = results(i)
    Next i
    ReadLocationFilterList = finalArr
End Function

' A tick is ON unless it is explicitly switched off. Blank means ON so that a
' location typed into a spare slot works with no further action -- the tick
' boxes are Form Controls linked to these cells, and a slot that never had one
' would otherwise be silently dead.
Private Function FilterTickIsOn(ByVal v As Variant) As Boolean
    FilterTickIsOn = True
    If IsEmpty(v) Or IsNull(v) Then Exit Function
    If VarType(v) = vbBoolean Then
        FilterTickIsOn = CBool(v)
        Exit Function
    End If
    If IsNumeric(v) Then
        FilterTickIsOn = (CDbl(v) <> 0)
        Exit Function
    End If
    Dim t As String
    t = UCase(Trim(CStr(v)))
    If t = "" Then Exit Function
    FilterTickIsOn = (t <> "FALSE" And t <> "0" And t <> "NO" And t <> "OFF")
End Function

' The Preferences sheet's own section labels. Used to find the end of the
' Location Filter block -- a blank-row rule cannot work, because the spare
' slots inside the block are blank by design.
Private Function IsPrefSectionLabel(ByVal s As String) As Boolean
    Select Case Trim(s)
        Case "Machine Roster", "Location Filter", "Row Formatting", _
             "Production Cutoff", "Press Status", "Hub Workbook Path"
            IsPrefSectionLabel = True
        Case Else
            IsPrefSectionLabel = False
    End Select
End Function

' Keeps * and ? working as wildcards while making everything else literal.
' VBA's Like also treats "#" as any-digit and "[" as the start of a character
' class, neither of which anyone typing a location name intends. Same
' treatment Hub gives its exclusion rules (EscapeLikeLiterals, 2026-08.17).
' "[" must be escaped first or it would escape its own replacement.
Private Function EscapeLikePattern(ByVal pat As String) As String
    Dim t As String
    t = Replace(pat, "[", "[[]")
    t = Replace(t, "#", "[#]")
    EscapeLikePattern = t
End Function

' EXACT match, with * standing for any run of characters and ? for exactly
' one -- which is what the Location Filter instructions on Preferences have
' always said, and what the code did NOT do until 2026-08-24. It used
' substring CONTAINS, so "Ready to Sch*" was searched for as a literal
' asterisk and matched nothing at all, while "Final Proof" silently swept in
' "Final Proof Returned" and "To PP Final Proof" as well.
'
' Both sides are already upper-cased by the callers, so Like's binary compare
' is the right one and no Option Compare change is needed. The patterns were
' escaped by EscapeLikePattern when they were read.
Private Function LocationMatchesFilter(ByVal ucLocation As String, ByRef locationFilter As Variant) As Boolean
    Dim k As Variant
    For Each k In locationFilter
        If ucLocation Like CStr(k) Then
            LocationMatchesFilter = True
            Exit Function
        End If
    Next k
    LocationMatchesFilter = False
End Function

' =============================================================================
' PER-MACHINE TAB BUILD
' =============================================================================
Private Sub BuildAllMachineTabs(ByRef rawData As Variant, ByRef roster As Variant)
    Dim totalRows As Long, totalCols As Long
    totalRows = UBound(rawData, 1)
    totalCols = UBound(rawData, 2)

    Dim locationFilter As Variant
    locationFilter = ReadLocationFilterList()

    ' One read of the Removed Jobs sheet for the whole build.
    Dim removed As Object
    Set removed = LoadRemovedJobs()
    Dim rfidOff As Object
    Set rfidOff = LoadRemovedRfidTypes()

    Dim m As Long
    For m = 1 To UBound(roster, 1)
        Dim keyword As String, ruleType As String, threshold As Double
        Dim lateAfter As Boolean
        keyword = CStr(roster(m, 1))
        ruleType = CStr(roster(m, 2))
        threshold = CDbl(roster(m, 3))
        lateAfter = (CStr(roster(m, 4)) <> "ON/AFTER")

        BuildOneMachineTab keyword, ruleType, threshold, lateAfter, rawData, totalRows, totalCols, removed, rfidOff, locationFilter
    Next m
End Sub

Private Sub BuildOneMachineTab(ByVal keyword As String, ByVal ruleType As String, ByVal threshold As Double, _
        ByVal lateAfter As Boolean, ByRef rawData As Variant, ByVal totalRows As Long, ByVal totalCols As Long, _
        ByRef removed As Object, ByRef rfidOff As Object, _
        ByRef locationFilter As Variant)

    Dim tabName As String
    tabName = SanitizeSheetName(keyword)

    If UCase(tabName) = "PREFERENCES" Or UCase(tabName) = UCase(CACHE_SHEET) Then
        MsgBox "Machine Roster entry '" & keyword & "' would create a tab named 'Preferences', which " & _
               "collides with a reserved sheet name. Skipped -- rename this roster entry.", vbExclamation, "Update Data"
        Exit Sub
    End If

    Dim ws As Worksheet
    Set ws = GetOrCreateSheet(tabName)
    ' 2026-08.26 -- ORDER MATTERS. The AutoFilter comes off FIRST, then the
    ' merges, and only then the contents. The previous order cleared first and
    ' left both in place until afterwards, so the whole-sheet Clear ran against
    ' a sheet that still carried an active filter -- possibly with rows hidden
    ' -- laid across 45 merged cells. That is the combination Excel refuses,
    ' and "failed at clearing the tab" was the report. Release the filter and
    ' the merges before touching the contents and there is nothing left to
    ' trip over.
    gStep = tabName & " -- removing the AutoFilter"
    If ws.AutoFilterMode Then ws.AutoFilterMode = False

    gStep = tabName & " -- unmerging the tab"
    ws.Cells.UnMerge

    gStep = tabName & " -- clearing the tab"
    ws.Cells.Clear

    ' 2026-08.24 -- THE v35 BUG. .Cells.Clear does NOT give a blank slate:
    ' cell MERGES and the sheet-level AUTOFILTER both survive it. The tab
    ' carries a merged Total Qty span (column K as of the 2026-09.03 To
    ' Perso insertion) for every Production Day, and an AutoFilter left
    ' over from the previous run.
    '
    ' That went unnoticed for as long as the day groupings came out identical
    ' on every rebuild -- each new merge landed exactly on top of the stale
    ' one, which Excel accepts silently. Job removal is the first feature that
    ' changes the row count between runs, so the day boundaries moved and a new
    ' merge started PARTIALLY overlapping a stale one. That is precisely when
    ' Excel raises "We can't do that to a merged cell" on the Merge, and
    ' "Unable to set the NumberFormat property of the Range class" on the next
    ' line that formats the same range. Both reported by the user on v35.
    '
    ' Unmerge and drop the filter explicitly -- the same belt-and-braces
    ' BuildOnboardingSheet has carried since 2026-08.12 for the same reason.
    ' (Both now run ABOVE the Clear; see the ordering note there.)
    ' .Cells.Clear wipes values/formats (incl. Interior.Color) but does NOT
    ' reset RowHeight -- RowHeight is a property of the Row object, not of
    ' cell formatting. WriteDaySeparatorRow sets a 3pt separator row height;
    ' since Production Day placement is recomputed from scratch every run
    ' and isn't sticky, a row that was a 3pt separator last run can stay
    ' squished to 3pt this run even after Clear, now holding real data
    ' instead of a separator. Reset every row back to the sheet's standard
    ' height before rebuilding so stale separator heights can't survive.
    gStep = tabName & " -- resetting row heights"
    ws.Rows.RowHeight = ws.StandardHeight

    ' -------------------------------------------------------------
    ' 1. Filter: Work Center CONTAINS keyword, case-insensitive/trimmed,
    ' AND Location CONTAINS at least one entry from the Location Filter
    ' list on Preferences (also case-insensitive/trimmed). The Location
    ' Filter narrows the queue to jobs actually at a relevant workflow
    ' stage right now (e.g. "Plates to Press", "Ready to Schedule") --
    ' added 2026-08 on top of the original "regardless of location" Work
    ' Center rule, not instead of it. If the Location Filter list is
    ' empty, the location restriction is skipped entirely (Work-Center-only
    ' behavior, same as before this feature existed) rather than silently
    ' matching nothing or everything.
    ' -------------------------------------------------------------
    Dim matchIdx() As Long
    ReDim matchIdx(1 To totalRows)
    Dim matchCount As Long: matchCount = 0

    Dim r As Long, wcVal As String, locVal As String
    Dim ucKeyword As String: ucKeyword = UCase(Trim(keyword))
    Dim hasLocationFilter As Boolean: hasLocationFilter = (Not IsEmpty(locationFilter))
    Dim nRemoved As Long: nRemoved = 0
    Dim nRfidSkipped As Long: nRfidSkipped = 0
    Dim hasRemovals As Boolean
    hasRemovals = False
    If Not removed Is Nothing Then hasRemovals = (removed.Count > 0)
    Dim hasRfidFilter As Boolean
    hasRfidFilter = False
    If Not rfidOff Is Nothing Then hasRfidFilter = (rfidOff.Count > 0)
    For r = 2 To totalRows  ' row 1 = headers from Hub's CleanedData
        If totalCols >= COL_WORKCENTER Then
            wcVal = UCase(Trim(CStr(rawData(r, COL_WORKCENTER))))
            If InStr(1, wcVal, ucKeyword) > 0 Then
                ' Removal is tested HERE -- before the queue is scheduled --
                ' so the capacity a pulled job would have used is genuinely
                ' handed back to the rest of the day rather than left as a
                ' hole in the schedule.
                Dim skipRow As Boolean
                skipRow = False
                If hasRemovals Then
                    If removed.Exists(RemovalKey(CStr(rawData(r, COL_JOBID)), _
                                                 CStr(rawData(r, COL_WORKCENTER)))) Then
                        skipRow = True
                        nRemoved = nRemoved + 1
                    End If
                End If

                ' Standing RFID technology filter. Applied HERE, with the rest
                ' of the filtering and before anything is scheduled, so the
                ' capacity these jobs would have used is genuinely handed back
                ' to the day rather than left as a hole.
                If Not skipRow Then
                    If hasRfidFilter Then
                        If UBound(rawData, 2) >= COL_RFID Then
                            If rfidOff.Exists(RfidFilterKey(tabName, CStr(rawData(r, COL_RFID)))) Then
                                skipRow = True
                                nRfidSkipped = nRfidSkipped + 1
                            End If
                        End If
                    End If
                End If

                If Not skipRow Then
                    If Not hasLocationFilter Then
                        matchCount = matchCount + 1
                        matchIdx(matchCount) = r
                    Else
                        locVal = UCase(Trim(CStr(rawData(r, COL_LOCATION))))
                        If LocationMatchesFilter(locVal, locationFilter) Then
                            matchCount = matchCount + 1
                            matchIdx(matchCount) = r
                        End If
                    End If
                End If
            End If
        End If
    Next r

    If matchCount = 0 Then
        ws.Cells(1, 1).Value = "No jobs currently match Work Center containing """ & keyword & _
            """" & IIf(hasLocationFilter, " with a Location Filter match", "") & "."
        ws.Cells(1, 1).Font.Italic = True
        Exit Sub
    End If

    ' -------------------------------------------------------------
    ' 2. Split into dated (has a valid effective schedule date --
    ' schedulable) and undated (neither To Perso nor Ship Date is a valid
    ' date -- shown but NOT given a Production Day, since the bucketing
    ' needs a date to sort by). The effective schedule date (To Perso when
    ' Hub has one for the job, else Ship Date -- see EffectiveScheduleDate)
    ' drives both scheduling and display order (Ship Date itself replaced
    ' Promised Date here, confirmed with the user 2026-08; To Perso now
    ' overrides Ship Date per the user 2026-09.03).
    ' -------------------------------------------------------------
    Dim datedIdx() As Long, undatedIdx() As Long
    ReDim datedIdx(1 To matchCount)
    ReDim undatedIdx(1 To matchCount)
    Dim datedCount As Long: datedCount = 0
    Dim undatedCount As Long: undatedCount = 0

    Dim i As Long, rowIdx As Long, sd As Variant
    For i = 1 To matchCount
        rowIdx = matchIdx(i)
        sd = EffectiveScheduleDate(rawData, rowIdx)
        If IsDate(sd) And Not IsEmpty(sd) Then
            ' 2026-08.12 evening revision: past-due jobs are back IN the
            ' dated queue -- they map to ship day 1 / deadline 0, so the
            ' due-tier greedy gives them first claim on the current
            ' production day. Their rows get the late flag at write time.
            datedCount = datedCount + 1
            datedIdx(datedCount) = rowIdx
        Else
            undatedCount = undatedCount + 1
            undatedIdx(undatedCount) = rowIdx
        End If
    Next i

    ' -------------------------------------------------------------
    ' 3. Sort dated rows by Ship Date asc, then Job ID asc (tiebreak).
    ' Simple insertion sort -- per-machine row counts are small (tens to
    ' low hundreds), not worth QuickSortMultiKey's complexity here.
    ' -------------------------------------------------------------
    If datedCount > 1 Then SortByShipThenJobID rawData, datedIdx, datedCount
    If undatedCount > 1 Then SortByJobIDOnly rawData, undatedIdx, undatedCount

    ' -------------------------------------------------------------
    ' 4. Schedule the dated queue into SEGMENTS (2026-08.12 carryover
    ' revision). A segment = (job, production day, qty run that day).
    ' Almost every job is exactly one whole segment; a job/family bigger
    ' than a full day's capacity is SPLIT: it fills a dedicated day to
    ' exactly the threshold and its remainder leads the next day as the
    ' very first work run (chaining across as many days as needed) --
    ' rule 7. Split (partial) segments render light-yellow.
    ' COUNT rule (Metronics, Indigo): ComputeCountSegments wraps the
    ' unchanged ComputeProductionDays -- one whole segment per job.
    ' QTY rule (Dicetrax): ComputeQtySegments -- see its header.
    ' -------------------------------------------------------------
    Dim segPos() As Long
    Dim segDate() As Date
    Dim segQty() As Double
    Dim segPartial() As Boolean, segCarry() As Boolean
    Dim segCount As Long: segCount = 0
    If datedCount > 0 Then
        If ruleType = "QTY" Then
            ComputeQtySegments rawData, datedIdx, datedCount, threshold, lateAfter, segPos, segDate, segQty, segPartial, segCarry, segCount
        Else
            ComputeCountSegments rawData, datedIdx, datedCount, threshold, segPos, segDate, segQty, segPartial, segCarry, segCount
        End If
    End If

    ' -------------------------------------------------------------
    ' 4b. Sort segments for DISPLAY: Production Day asc, then within a
    ' day the CARRYOVER segment(s) first ('the very first job will be the
    ' job that did not finish prior' -- user rule, 2026-08.12), then Ship
    ' Date, then numeric-aware Job ID. Day-group blocks below (separators,
    ' merged Total Qty, thick outlines) rely on each day's segments being
    ' contiguous in write order.
    ' -------------------------------------------------------------
    If segCount > 1 Then
        SortSegments rawData, datedIdx, segPos, segDate, segQty, segPartial, segCarry, segCount
    End If

    ' -------------------------------------------------------------
    ' 5. Write header + rows (dated first, then undated).
    ' -------------------------------------------------------------
    ' Current layout (12 headers): To Perso, Ship Date, Job ID, Customer,
    ' Description, Qty, RFID, Last Location, Work Center, Production Day,
    ' Total Qty, Remove -- see the COL_OUT_* constants near the top of this
    ' module for the column each one lands in. "Remove" (COL_OUT_REMOVE) is
    ' deliberately NOT part of the formatted/bordered block below -- the
    ' operator's tick must stand outside the schedule's own styling -- so
    ' the border/fill/AutoFilter ranges stop at COL_OUT_TOTALQTY, one short
    ' of it.
    ' 2026-09.03: "To Perso" inserted as the new first column, ahead of
    ' Ship Date -- see EffectiveScheduleDate for what it changes beyond
    ' display.
    Dim headers As Variant
    headers = Array("To Perso", "Ship Date", "Job ID", "Customer Name", _
                     "Job Description", "Qty", "RFID", "Last Location", "Work Center", "Production Day", _
                     "Total Qty", "Remove")

    Dim c As Long
    For c = 0 To UBound(headers)
        ws.Cells(1, c + 1).Value = headers(c)
    Next c
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, COL_OUT_REMOVE))
        .Font.Bold = True
        .Interior.Color = RGB(31, 78, 120)
        .Font.Color = RGB(255, 255, 255)
    End With
    ApplyThinBorderGrid ws, 1, 1, COL_OUT_TOTALQTY

    ' Total Qty (col K): one merged, centered cell per Production Day group,
    ' showing the SUM of that day's raw Qty (col F) -- confirmed with the
    ' user 2026-08, verified against their own worked example (33,080 across
    ' 11 Dicetrax rows) before building this. Undated rows intentionally get
    ' no total -- there's no Production Day to group them by (confirmed with
    ' the user). groupStartRow/groupSumQty track the CURRENT day group as the
    ' loop walks the Ship-Date-sorted, already-filtered dated queue; the
    ' previous group is closed out (merged + totaled) the moment a new day
    ' is seen, and the final group is closed out once after the loop ends
    ' since no further boundary will trigger it.
    gStep = tabName & " -- writing job rows"
    Dim outRow As Long: outRow = 2


    Dim groupStartRow As Long: groupStartRow = 0
    Dim groupSumQty As Double: groupSumQty = 0
    Dim isNewGroup As Boolean
    Dim isLate As Boolean, isShipDay As Boolean
    Dim deadlineDate As Variant
    Dim stateIdx As Long
    Dim rowFmts As Variant: rowFmts = ReadRowFormats()
    For i = 1 To segCount
        rowIdx = datedIdx(segPos(i))
        If i = 1 Then
            isNewGroup = True
        Else
            isNewGroup = (segDate(i) <> segDate(i - 1))
        End If

        ' Black separator row between one Production Day's segments and
        ' the next (unchanged from the pre-segment design).
        If isNewGroup Then
            If i > 1 Then
                FinalizeDayTotalQty ws, groupStartRow, outRow - 1, groupSumQty, COL_OUT_TOTALQTY
                ApplyThinBorderGrid ws, groupStartRow, outRow - 1, COL_OUT_TOTALQTY
                ApplyThickGroupOutline ws, groupStartRow, outRow - 1, COL_OUT_TOTALQTY
                WriteDaySeparatorRow ws, outRow, COL_OUT_TOTALQTY
                outRow = outRow + 1
            End If
            groupStartRow = outRow
            groupSumQty = 0
        End If

        WritePressRow ws, outRow, rawData, rowIdx
        ws.Cells(outRow, COL_OUT_PRODDAY).Value = segDate(i)
        ' A split job's row shows the qty RUN THAT DAY, not its full
        ' order qty -- the remainder appears on the next day's row(s).
        ' A split row shows the QUANTITY RUN THAT DAY, not the whole order
        ' quantity and not a piece range (that was withdrawn 2026-08.18).
        If segPartial(i) Then ws.Cells(outRow, COL_OUT_QTY).Value = segQty(i)

        ' LATE / SHIP-DAY check (2026-08.13). Under the AFTER rule
        ' (Dicetrax, Metronics) a job is late only when it runs AFTER its
        ' deadline; running exactly ON it is on time and gets its own
        ' "ships today" state. Under the legacy ON/AFTER rule (Indigo)
        ' running on the deadline is still late. Which rule a machine uses
        ' comes from the Machine Roster's Late Rule column.
        ' 2026-09.03: "deadline" is the effective schedule date -- To Perso
        ' when Hub has one for this job, else Ship Date (same date this row
        ' was queued and bucketed against; see EffectiveScheduleDate).
        deadlineDate = EffectiveScheduleDate(rawData, rowIdx)
        isLate = False
        isShipDay = False
        If IsDate(deadlineDate) Then
            If CLng(segDate(i)) > CLng(CDate(deadlineDate)) Then
                isLate = True
            ElseIf CLng(segDate(i)) = CLng(CDate(deadlineDate)) Then
                If lateAfter Then
                    isShipDay = True
                Else
                    isLate = True
                End If
            End If
        End If

        ' Row state -> formatting, all of it user-editable on Preferences
        ' (see ReadRowFormats / EnsureRowFormatTable). Split beats late
        ' beats ships-today, so a split row that is also late shows the
        ' dedicated "Split + Late" style rather than either one alone.
        If segPartial(i) And isLate Then
            stateIdx = FMT_SPLIT_LATE
        ElseIf segPartial(i) Then
            stateIdx = FMT_SPLIT
        ElseIf isLate Then
            stateIdx = FMT_LATE
        ElseIf isShipDay Then
            stateIdx = FMT_SHIPDAY
        Else
            stateIdx = FMT_NORMAL
        End If
        ApplyRowFormat ws, outRow, COL_OUT_PRODDAY, rowFmts, stateIdx
        ws.Cells(outRow, COL_STATE_MARK).Value = stateIdx

        groupSumQty = groupSumQty + segQty(i)
        outRow = outRow + 1
    Next i
    If segCount > 0 Then
        FinalizeDayTotalQty ws, groupStartRow, outRow - 1, groupSumQty, COL_OUT_TOTALQTY
        ApplyThinBorderGrid ws, groupStartRow, outRow - 1, COL_OUT_TOTALQTY
        ApplyThickGroupOutline ws, groupStartRow, outRow - 1, COL_OUT_TOTALQTY
    End If

    ' Undated rows ("No Ship Date") get the same standard thin grid as every
    ' other occupied cell, but NOT a thick outline -- they aren't a "day"
    ' group (no Production Day to group by), consistent with the earlier
    ' decision that they also don't get a Total Qty total.
    gStep = tabName & " -- writing undated jobs"
    Dim undatedStartRow As Long: undatedStartRow = outRow
    For i = 1 To undatedCount
        rowIdx = undatedIdx(i)
        WritePressRow ws, outRow, rawData, rowIdx
        ws.Cells(outRow, COL_OUT_PRODDAY).Value = "No Ship Date"
        ws.Cells(outRow, COL_OUT_PRODDAY).Font.Italic = True
        outRow = outRow + 1
    Next i
    If undatedCount > 0 Then
        ApplyThinBorderGrid ws, undatedStartRow, outRow - 1, COL_OUT_TOTALQTY
    End If

    If outRow > 2 Then
        gStep = tabName & " -- applying number formats"
        SetFmt ws.Range(ws.Cells(2, COL_OUT_TOPERSO), ws.Cells(outRow - 1, COL_OUT_TOPERSO)), "mm/dd/yyyy", tabName & " To Perso"
        SetFmt ws.Range(ws.Cells(2, COL_OUT_SHIP), ws.Cells(outRow - 1, COL_OUT_SHIP)), "mm/dd/yyyy", tabName & " Ship Date"
        SetFmt ws.Range(ws.Cells(2, COL_OUT_QTY), ws.Cells(outRow - 1, COL_OUT_QTY)), "#,##0", tabName & " Qty"
        SetFmt ws.Range(ws.Cells(2, COL_OUT_RFID), ws.Cells(outRow - 1, COL_OUT_RFID)), "@", tabName & " RFID"
        SetFmt ws.Range(ws.Cells(2, COL_OUT_PRODDAY), ws.Cells(outRow - 1, COL_OUT_PRODDAY)), "mm/dd/yyyy (ddd)", tabName & " Production Day"
    End If

    ' Explicit range, not the bare Rows(1).AutoFilter shortcut -- the black
    ' separator rows are blank (no cell VALUES, just a fill color), and
    ' Excel's implicit CurrentRegion detection treats a value-less row as
    ' the end of the table, which would silently cut the filter off at the
    ' first separator. Spelling out the range through the true last row
    ' keeps every job filterable regardless of how many separators are in
    ' between.
    ' Filter range stops at Production Day (J as of the 2026-09.03 To Perso
    ' insertion; was I before it).
    '
    ' Total Qty (K as of that same insertion; was J before it) is EXCLUDED
    ' because it holds the merged per-day cells. A filter range that spans
    ' merged cells is a trap: filtering behaves oddly and SORTING from a
    ' dropdown fails outright with 1004 "We can't do that to a merged cell"
    ' -- reported from the Dicetrax tab 2026-08.26, where all 45 merges sat
    ' inside an A1:J733 filter range (column letters as they were that day).
    ' Nothing is lost by dropping it: a per-day merged total is not a
    ' meaningful thing to filter or sort on, and hiding a row still hides
    ' that cell along with it.
    '
    ' Remove (L, was K) is excluded too, and always was -- it is an INPUT.
    ' Leaving it filterable would let an operator tick rows, filter the
    ' ticks away, and lose sight of what they had just marked for removal.
    gStep = tabName & " -- applying the AutoFilter"
    ws.Range(ws.Cells(1, 1), ws.Cells(IIf(outRow > 2, outRow - 1, 1), COL_OUT_PRODDAY)).AutoFilter
    gStep = tabName & " -- sizing columns"
    ws.Columns.AutoFit

    ' Every job row gets an unticked box. Tick and click Refresh Data.
    gStep = tabName & " -- drawing the Remove checkboxes"
    ApplyRemoveCheckboxes ws, IIf(outRow > 2, outRow - 1, 1)
    ws.Columns(COL_OUT_REMOVE).ColumnWidth = 9

    gStep = tabName & " -- building the RFID type list"
    BuildRfidTypeList ws, IIf(outRow > 2, outRow - 1, 1), tabName, rfidOff

    ws.Columns(COL_STATE_MARK).Hidden = True

    On Error Resume Next
    ws.Activate
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = False
    ActiveWindow.SplitColumn = 0
    ActiveWindow.SplitRow = 1
    ActiveWindow.FreezePanes = True
    On Error GoTo 0
End Sub

Private Sub WritePressRow(ByRef ws As Worksheet, ByVal outRow As Long, ByRef rawData As Variant, ByVal rowIdx As Long)
    ' To Perso comes straight from Hub's CleanedData column L. Blank on a
    ' HubCache written by an older Hub build (fewer than 12 columns) or on
    ' a job that hasn't reached that stage. This is the RAW To Perso date
    ' -- display only. Scheduling reads it through EffectiveScheduleDate.
    If UBound(rawData, 2) >= COL_TOPERSO Then
        ws.Cells(outRow, COL_OUT_TOPERSO).Value = rawData(rowIdx, COL_TOPERSO)
    End If
    ws.Cells(outRow, COL_OUT_SHIP).Value = rawData(rowIdx, COL_SHIP)
    ws.Cells(outRow, COL_OUT_JOBID).Value = rawData(rowIdx, COL_JOBID)
    ws.Cells(outRow, COL_OUT_CUSTOMER).Value = rawData(rowIdx, COL_CUSTOMER)
    ws.Cells(outRow, COL_OUT_DESC).Value = rawData(rowIdx, COL_DESC)
    ws.Cells(outRow, COL_OUT_QTY).Value = rawData(rowIdx, COL_QTY)
    ' RFID comes straight from Hub's CleanedData column K (RFID Type).
    ' Hub owns the classification -- Press only displays it. Blank when the
    ' job description matched none of Hub's RFID Technologies tokens.
    If UBound(rawData, 2) >= COL_RFID Then
        ws.Cells(outRow, COL_OUT_RFID).Value = rawData(rowIdx, COL_RFID)
    End If
    ws.Cells(outRow, COL_OUT_LOCATION).Value = NormalizeLocationDisplay(rawData(rowIdx, COL_LOCATION))
    ws.Cells(outRow, COL_OUT_WORKCENTER).Value = rawData(rowIdx, COL_WORKCENTER)
End Sub

' Display-only cleanup for the "Last Location" column -- does NOT affect the
' Location Filter matching in LocationMatchesFilter, which still runs against
' the raw value (and already has both "TI" and "To Imposition" seeded as
' separate, independent phrases). This only expands a bare "TI" into its
' full name so the tab reads clearly. Exact-match on the whole field (after
' trim/case-fold), not a "contains" check, so it won't touch a Location that
' merely mentions TI as part of a longer phrase -- confirmed with the user
' 2026-08 as an "is TI" match, e.g. "TI", "ti", "Ti", "tI".
Private Function NormalizeLocationDisplay(ByVal rawLoc As Variant) As String
    Dim s As String: s = Trim(CStr(rawLoc))
    If UCase(s) = "TI" Then
        NormalizeLocationDisplay = "To Imposition"
    Else
        NormalizeLocationDisplay = s
    End If
End Function

' Thin black band marking the boundary between one Production Day's jobs and
' the next on a machine tab. 4 pixels tall -- Excel's RowHeight property is
' in points, not pixels, and the standard conversion at 96 DPI (Excel/Windows
' default) is 1 point = 4/3 pixels, so 4px = 3pt. Confirmed with the user
' 2026-08 (down from an initial "half a normal row" / 7.5pt). Filled solid
' black across every data column -- no cell values, purely a visual divider.
' See the AutoFilter comment at its call site for why the filter range has
' to be spelled out explicitly once these exist on a tab.
' 2026-08.7: added a thick black BorderAround the fill (same xlThick weight
' as ApplyThickGroupOutline uses for each day's block) at the user's
' request, so the divider reads as a distinct, deliberately heavy boundary
' rather than just a thin colored strip -- e.g. when the row's own black
' fill blends into an adjacent dark theme/print setting.
Private Sub WriteDaySeparatorRow(ByRef ws As Worksheet, ByVal outRow As Long, ByVal numCols As Long)
    gStep = ws.name & " -- day separator row " & outRow
    With ws.Range(ws.Cells(outRow, 1), ws.Cells(outRow, numCols))
        .Interior.Color = RGB(0, 0, 0)
        .BorderAround LineStyle:=xlContinuous, Weight:=xlThick, Color:=RGB(0, 0, 0)
    End With
    ws.Rows(outRow).RowHeight = 3
End Sub

' Standard Excel "All Borders" grid (thin, continuous, every edge of every
' cell -- top/bottom/left/right AND the lines between cells) across the
' given A:numCols block. Used on the header row and on every occupied block
' of data rows -- never on the black separator rows, which have no cell
' values and are already a solid fill. Confirmed with the user 2026-08.
Private Sub ApplyThinBorderGrid(ByRef ws As Worksheet, ByVal startRow As Long, ByVal endRow As Long, ByVal numCols As Long)
    gStep = ws.name & " -- grid borders rows " & startRow & "-" & endRow
    With ws.Range(ws.Cells(startRow, 1), ws.Cells(endRow, numCols)).Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
End Sub

' Thick border around just the OUTSIDE of one Production Day's A:numCols
' block, leaving the thin internal grid from ApplyThinBorderGrid untouched --
' confirmed with the user 2026-08. Must run AFTER ApplyThinBorderGrid on the
' same range (see call site comment).
Private Sub ApplyThickGroupOutline(ByRef ws As Worksheet, ByVal startRow As Long, ByVal endRow As Long, ByVal numCols As Long)
    gStep = ws.name & " -- day outline rows " & startRow & "-" & endRow
    ws.Range(ws.Cells(startRow, 1), ws.Cells(endRow, numCols)).BorderAround LineStyle:=xlContinuous, Weight:=xlThick
End Sub

' Closes out one Production Day's Total Qty cell: merges col totalQtyCol
' across that day's row-span (startRow..endRow), writes the day's summed
' raw Qty into it, and centers both horizontally and vertically -- confirmed
' with the user 2026-08. A single-row group (startRow = endRow) is left
' unmerged since merging one cell with itself is a no-op; the value/format/
' alignment are still applied so it looks identical to a multi-row group.
' DisplayAlerts is toggled off around the Merge call defensively: every cell
' in this column is guaranteed blank going in (WritePressRow never touches
' col J, and BuildOneMachineTab clears the whole tab before rebuilding), so
' Excel's "discard other cells' data" merge prompt should never fire, but
' there's no reason to risk a hung dialog on an unattended rebuild.
Private Sub FinalizeDayTotalQty(ByRef ws As Worksheet, ByVal startRow As Long, ByVal endRow As Long, _
        ByVal sumQty As Double, ByVal totalQtyCol As Long)
    gStep = ws.name & " -- Total Qty merge rows " & startRow & "-" & endRow
    Dim rng As Range
    Set rng = ws.Range(ws.Cells(startRow, totalQtyCol), ws.Cells(endRow, totalQtyCol))

    ' Second line of defence behind BuildOneMachineTab's sheet-wide UnMerge.
    ' If anything ever leaves a merge straddling this span again, unmerging it
    ' first turns a hard 1004 into a cosmetic non-event. UnMerge on a range
    ' that only intersects a merged area releases the whole area, which is
    ' what we want here.
    rng.UnMerge

    If endRow > startRow Then
        Dim prevAlerts As Boolean: prevAlerts = Application.DisplayAlerts
        Application.DisplayAlerts = False
        rng.Merge
        Application.DisplayAlerts = prevAlerts
    End If

    rng.Cells(1, 1).Value = sumQty
    SetFmt rng, "#,##0", ws.name & " Total Qty"
    rng.HorizontalAlignment = xlHAlignCenter
    rng.VerticalAlignment = xlVAlignCenter
End Sub

' =============================================================================
' SORTING (simple insertion sort -- per-machine sets are small)
' =============================================================================
Private Sub SortByShipThenJobID(ByRef rawData As Variant, ByRef idx() As Long, ByVal n As Long)
    Dim i As Long, j As Long, key As Long
    For i = 2 To n
        key = idx(i)
        j = i - 1
        Do While j >= 1
            If CompareRowsShipThenJobID(rawData, idx(j), key) <= 0 Then Exit Do
            idx(j + 1) = idx(j)
            j = j - 1
        Loop
        idx(j + 1) = key
    Next i
End Sub

' Job ID tiebreak: a hyphen-suffixed Job ID (e.g. "6015514-1", "6015514-2")
' is compared on (base, suffix) with the suffix compared NUMERICALLY, not as
' a raw string -- the old string compare put "-10" before "-2" (fixed
' 2026-08.4, per the Dicetrax scheduling redesign's Rule 1: "Job IDs that
' have hyphen 1 or 2 or more need to stay grouped together in numerical
' order"). This also guarantees a chain's parts always land in the sorted
' queue in ascending-suffix order, which ComputeProductionDaysQty's
' BuildGroupsQty depends on as a hard precondition -- see that function's
' header for the full analysis of why this precondition is safe to rely on
' without additional defensive re-sorting inside the scheduler itself.
' Name kept as-is (widely referenced) though the date compared is now the
' EFFECTIVE schedule date, not necessarily raw Ship Date -- see
' EffectiveScheduleDate (2026-09.03). Every caller already only reaches
' this with rows that passed that same function's IsDate check, so CDate
' here is always safe.
Private Function CompareRowsShipThenJobID(ByRef rawData As Variant, ByVal rowA As Long, ByVal rowB As Long) As Long
    Dim dA As Date, dB As Date
    dA = CDate(EffectiveScheduleDate(rawData, rowA))
    dB = CDate(EffectiveScheduleDate(rawData, rowB))
    If dA < dB Then
        CompareRowsShipThenJobID = -1
    ElseIf dA > dB Then
        CompareRowsShipThenJobID = 1
    Else
        CompareRowsShipThenJobID = CompareJobIdOnly(rawData, rowA, rowB)
    End If
End Function

' Same (base, suffix)-numeric-aware Job ID compare as CompareRowsShipThenJobID,
' without the Ship Date component -- used both for the undated-rows sort below
' and as the tiebreak inside CompareRowsShipThenJobID itself.
Private Function CompareJobIdOnly(ByRef rawData As Variant, ByVal rowA As Long, ByVal rowB As Long) As Long
    Dim baseA As String, baseB As String, sufA As Long, sufB As Long
    ParseJobIdBaseSuffix CStr(rawData(rowA, COL_JOBID)), baseA, sufA
    ParseJobIdBaseSuffix CStr(rawData(rowB, COL_JOBID)), baseB, sufB
    If baseA < baseB Then
        CompareJobIdOnly = -1
    ElseIf baseA > baseB Then
        CompareJobIdOnly = 1
    ElseIf sufA < sufB Then
        CompareJobIdOnly = -1
    ElseIf sufA > sufB Then
        CompareJobIdOnly = 1
    Else
        CompareJobIdOnly = 0
    End If
End Function

' Splits a Job ID into (base, suffix): a trailing "-<digits>" is a hyphen-
' chain suffix (e.g. "6015514-2" -> base "6015514", suffix 2); anything else
' (no trailing "-digits", including a bare numeric Job ID or one with
' non-numeric text after its last hyphen) is an ordinary/solo item with
' suffix 0, i.e. never treated as part of a chain. Hand-parsed rather than
' a Regex object -- avoids adding a VBScript_RegExp library reference for
' one simple pattern (mirrors Python's `^(.*)-(\d+)$` from the validated
' prototype, proto_scheduler2.py).
Private Sub ParseJobIdBaseSuffix(ByVal jobId As String, ByRef baseOut As String, ByRef suffixOut As Long)
    Dim s As String: s = Trim(jobId)
    Dim hyphenPos As Long: hyphenPos = InStrRev(s, "-")
    If hyphenPos > 0 And hyphenPos < Len(s) Then
        Dim tail As String: tail = Mid(s, hyphenPos + 1)
        If IsAllDigits(tail) Then
            baseOut = Left(s, hyphenPos - 1)
            suffixOut = CLng(tail)
            Exit Sub
        End If
    End If
    baseOut = s
    suffixOut = 0
End Sub

Private Function IsAllDigits(ByVal s As String) As Boolean
    If Len(s) = 0 Then
        IsAllDigits = False
        Exit Function
    End If
    Dim k As Long
    For k = 1 To Len(s)
        If Mid(s, k, 1) < "0" Or Mid(s, k, 1) > "9" Then
            IsAllDigits = False
            Exit Function
        End If
    Next k
    IsAllDigits = True
End Function

Private Sub SortByJobIDOnly(ByRef rawData As Variant, ByRef idx() As Long, ByVal n As Long)
    Dim i As Long, j As Long, key As Long
    For i = 2 To n
        key = idx(i)
        j = i - 1
        Do While j >= 1
            If CompareJobIdOnly(rawData, idx(j), key) <= 0 Then Exit Do
            idx(j + 1) = idx(j)
            j = j - 1
        Loop
        idx(j + 1) = key
    Next i
End Sub

' Sorts the display SEGMENTS: Production Day asc, carry-lead first within a
' day, then Ship Date asc, then numeric-aware Job ID. Insertion sort over
' the five parallel segment arrays; segment counts are small.
Private Sub SortSegments(ByRef rawData As Variant, ByRef datedIdx() As Long, _
        ByRef segPos() As Long, ByRef segDate() As Date, ByRef segQty() As Double, _
        ByRef segPartial() As Boolean, ByRef segCarry() As Boolean, ByVal segCount As Long)
    Dim i As Long, j As Long
    Dim kPos As Long, kQty As Double
    Dim kDate As Date
    Dim kPart As Boolean, kCarry As Boolean
    For i = 2 To segCount
        kPos = segPos(i): kDate = segDate(i): kQty = segQty(i)
        kPart = segPartial(i): kCarry = segCarry(i)
        j = i - 1
        Do While j >= 1
            If CompareSegs(rawData, datedIdx, segPos(j), segDate(j), segCarry(j), segPartial(j), _
                           kPos, kDate, kCarry, kPart) > 0 Then
                segPos(j + 1) = segPos(j): segDate(j + 1) = segDate(j): segQty(j + 1) = segQty(j)
                segPartial(j + 1) = segPartial(j): segCarry(j + 1) = segCarry(j)
                j = j - 1
            Else
                Exit Do
            End If
        Loop
        segPos(j + 1) = kPos: segDate(j + 1) = kDate: segQty(j + 1) = kQty
        segPartial(j + 1) = kPart: segCarry(j + 1) = kCarry
    Next i
End Sub

' Ordering WITHIN one production day mirrors the order the machine
' actually runs the work:
'   rank 0  the remainder carried in from yesterday -- it leads the day
'   rank 1  whole jobs, in Ship Date then Job ID order
'   rank 2  the slice cut by today's capacity line -- it closes the day
' Days themselves are always in date order.
Private Function CompareSegs(ByRef rawData As Variant, ByRef datedIdx() As Long, _
        ByVal posA As Long, ByVal dateA As Date, ByVal carryA As Boolean, ByVal partA As Boolean, _
        ByVal posB As Long, ByVal dateB As Date, ByVal carryB As Boolean, ByVal partB As Boolean) As Long
    If dateA < dateB Then CompareSegs = -1: Exit Function
    If dateA > dateB Then CompareSegs = 1: Exit Function

    Dim rA As Long, rB As Long
    rA = SegDayRank(carryA, partA)
    rB = SegDayRank(carryB, partB)
    If rA < rB Then CompareSegs = -1: Exit Function
    If rA > rB Then CompareSegs = 1: Exit Function

    CompareSegs = CompareRowsShipThenJobID(rawData, datedIdx(posA), datedIdx(posB))
End Function

Private Function SegDayRank(ByVal isCarry As Boolean, ByVal isPart As Boolean) As Long
    If isCarry Then
        SegDayRank = 0
    ElseIf isPart Then
        SegDayRank = 2
    Else
        SegDayRank = 1
    End If
End Function


' =============================================================================
' PRODUCTION DAY (COUNT rule only -- Metronics, Indigo) -- a SCHEDULING
' PROJECTION, not the job's own date. Starting from the next business day
' from today, walks the queue (sorted by effective schedule date, then Job
' ID -- see EffectiveScheduleDate) and assigns N jobs per day (threshold =
' N), regardless of qty, in consecutive business-day buckets.
'
' The QTY rule (Dicetrax) is handled by ComputeProductionDaysQty instead
' (2026-08.4 redesign, see its header) -- NOT by a branch in this function
' any more. COUNT does not need that same rolling-window/knapsack treatment:
' with a job-COUNT capacity, every item costs exactly 1 "unit" regardless of
' its qty, so there is nothing to combinatorially optimize -- simple
' sequential bucketing is already optimal. The hyphen-chain grouping/order
' rule (Rule 1 of the redesign, confirmed to apply to "all three machines")
' is satisfied here for free, with no extra code: bucket day number is a
' non-decreasing function of queue position, and CompareRowsShipThenJobID
' guarantees a chain's parts are already queued in ascending-suffix order,
' so a chain's later part can never land on an earlier day than an earlier
' part. Confirmed 2026-08.4.
'
' Recomputed from scratch on every Update Data run -- confirmed with the
' user (2026-08): NOT sticky/persistent across runs. A job's Production Day
' can move if the queue changes between runs.
' =============================================================================
Private Function ComputeProductionDays(ByRef rawData As Variant, ByRef datedIdx() As Long, _
        ByVal n As Long, ByVal threshold As Double) As Date()
    Dim result() As Date
    ReDim result(1 To n)

    Dim currentDay As Date
    currentDay = ScheduleDayOne()

    Dim i As Long
    Dim countInDay As Long: countInDay = 0
    Dim perDay As Long: perDay = CLng(threshold)
    If perDay < 1 Then perDay = 1  ' guard against a blank/zero threshold locking every job to day 1
    For i = 1 To n
        If countInDay >= perDay Then
            currentDay = NextBusinessDayAfter(currentDay)
            countInDay = 0
        End If
        result(i) = currentDay
        countInDay = countInDay + 1
    Next i

    ComputeProductionDays = result
End Function
' =============================================================================
' QTY-RULE SCHEDULER -- v34 STRICT SHIP-DATE ORDER WITH SPLITS (2026-08.18)
'
' 2026-09.03: everywhere below that says "Ship Date" now means a job's
' EFFECTIVE schedule date -- To Perso when Hub has one for that job, else
' Ship Date (EffectiveScheduleDate). The rule itself, and every guarantee
' below, is unchanged; it now just runs against that date instead of
' always the raw Ship Date.
'
' Rule 3 replaced. The owner:
'   "Jobs MUST run in ship date order no matter what. If a job would cause a
'    split and fall over to the next day make sure all smaller jobs are
'    populated for the current ship day as much as possible then setup the
'    split job. In the QTY column of the split job put start and finish."
'
' So:
'   1. Strict Ship Date order is now a HARD constraint. Nothing is ever
'      pulled ahead of a job that ships earlier.
'   2. Each day is filled in queue order. When the next job does not fit
'      the space left, the scheduler first places any smaller job SHARING
'      THAT JOB'S SHIP DATE -- which cannot break ship-date order, since
'      the dates are equal -- and only then cuts the blocked job into
'      whatever space remains.
'   3. The cut job's remainder leads the next day, chaining for as long as
'      it takes.
'   4. A split row shows the QUANTITY RUN THAT DAY. (The piece range
'      shipped briefly on 2026-08.17b was withdrawn the next day.)
'
' OVERDUE JOBS (changed 2026-08.18): they now COUNT against the day's
' capacity, so there is no special case left -- an overdue job has the
' earliest ship date in the queue, so the ordinary fill above places it
' first, on the current production day, charged to that day like any other
' work. The practical effect is that day one can no longer read above its
' threshold, and if there is more overdue work than one day holds, the
' excess splits or rolls forward exactly as rule 3 says.
'
' WHY SAME-SHIP-DATE ONLY. The instruction could also be read as "fill the
' day from anything left that fits". Both were run against the real 566-job
' Dicetrax queue:
'
'                                   days  avg fill  splits  ORDER VIOLATIONS
'   same ship date only (shipped)     43     98.7%      42          0
'   any remaining job that fits       43     98.7%      27      1,613
'
' Identical days and identical fill. Filling from anywhere buys 15 fewer
' splits at the cost of 1,613 places where a job runs ahead of something
' that ships earlier -- precisely what "no matter what" forbids.
'
' The fill scan is a SINGLE left-to-right pass: the space left in a day only
' ever shrinks, so a job that fails to fit can never fit again that day.
'
' Terminates on every input: each turn of the loop either places a whole
' job or cuts at least one piece off the blocked job (spaceLeft is reset to
' a positive threshold at the start of every day), so remaining work
' strictly decreases.
' =============================================================================
' JOB REMOVAL (2026-08.19)
'
' The operator's loop:
'   Update Data  -> fresh pull from Hub. Every job carries an UNTICKED
'                   checkbox in column L. Nothing is removed -- a clean slate.
'   tick boxes   -> mark the jobs that should not run.
'   Refresh Data -> those jobs are removed, then the standard refresh rules
'                   rebuild the schedule around the gap they leave.
'
' Removals accumulate across repeated Refresh presses within a session, and
' are wiped by the next Update Data. That is the user's choice: nothing
' stays hidden without somebody deciding again today.
'
' The tick itself is transient -- machine tabs are wiped and rebuilt on every
' run, so a mark that lived only on the tab would not survive its own
' refresh. The durable record is the "Removed Jobs" sheet, harvested BEFORE
' any tab is touched.
'
' KEY = JOB ID + WORK CENTER. Verified unique against the live queue on
' 2026-08.19: 566/566 Dicetrax, 10/10 Metronics, 51/51 Indigo, zero
' collisions -- so the ship date is not needed to identify a line.
' Two consequences worth knowing:
'   * Ticking ANY slice of a split job removes the whole job from that
'     machine. Correct -- you cannot run half a job you have decided to pull.
'   * If a job's ship date later moves, the removal still holds. A key that
'     included the ship date would quietly stop matching and the job would
'     reappear.
' Removal is per MACHINE: pulling a job off Dicetrax leaves it on Metronics
' and Indigo.
'
' Because the filter runs BEFORE scheduling, the freed capacity is genuinely
' reused -- the rest of the day packs into the gap rather than leaving a hole.
' =============================================================================

' Reads the Removed Jobs sheet into a lookup of "JOBID|WORKCENTER" keys.
Private Function LoadRemovedJobs() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    Set LoadRemovedJobs = d

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(REMOVED_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < REMOVED_FIRST_ROW Then Exit Function

    Dim r As Long, k As String
    For r = REMOVED_FIRST_ROW To lastRow
        k = RemovalKey(CStr(ws.Cells(r, 1).Value), CStr(ws.Cells(r, 2).Value))
        If k <> "|" Then
            If Not d.Exists(k) Then d.Add k, True
        End If
    Next r
End Function

' =============================================================================
' STANDING RFID TYPE FILTERS (2026-08.25)
'
' A ticked technology in column N is NOT a one-shot removal. It is a filter
' that stays on: its jobs keep out of the schedule on every Refresh Data until
' somebody unticks it, and unticking brings them straight back.
'
' That is why it cannot be recorded on the Removed Jobs list in A:G. Entries
' there are permanent by design -- writing 356 ULC jobs into that list would
' make the tick a one-way door. The ticked TYPES are stored instead, in I:J,
' and the whole block is rewritten from the tabs on every Refresh. Untick a
' type and it simply is not written back, so its jobs return.
'
' The checkbox itself cannot be the store: BuildOneMachineTab wipes the tab it
' lives on. Harvested before the wipe, restored after -- the same shape as the
' Cards/Ready ticks in the Perso workbook.
' =============================================================================
' Idempotent: safe to call on every run, and the reason a Removed Jobs sheet
' created by an older build still gains this block.
Private Sub EnsureRemovedRfidBlock(ByRef ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    ws.Cells(REMOVED_HEADER_ROW - 1, REMOVED_RFID_MACHINE).Value = _
        "RFID TYPES CURRENTLY SWITCHED OFF -- untick on the machine tab and Refresh to bring them back"
    ws.Cells(REMOVED_HEADER_ROW - 1, REMOVED_RFID_MACHINE).Font.Italic = True
    ws.Cells(REMOVED_HEADER_ROW - 1, REMOVED_RFID_MACHINE).Font.Size = 9
    ws.Cells(REMOVED_HEADER_ROW, REMOVED_RFID_MACHINE).Value = "Machine"
    ws.Cells(REMOVED_HEADER_ROW, REMOVED_RFID_TYPE).Value = "RFID Type"
    With ws.Range(ws.Cells(REMOVED_HEADER_ROW, REMOVED_RFID_MACHINE), ws.Cells(REMOVED_HEADER_ROW, REMOVED_RFID_TYPE))
        .Font.Bold = True
        .Interior.Color = RGB(84, 130, 53)
        .Font.Color = RGB(255, 255, 255)
    End With
    ws.Columns(REMOVED_RFID_MACHINE).ColumnWidth = 18
    ws.Columns(REMOVED_RFID_TYPE).ColumnWidth = 18
End Sub


Private Function RfidFilterKey(ByVal tabName As String, ByVal rfidType As String) As String
    RfidFilterKey = UCase(Trim(tabName)) & "|" & UCase(Trim(rfidType))
End Function


' Every standing filter currently stored, keyed TAB|TYPE.
Private Function LoadRemovedRfidTypes() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    Set LoadRemovedRfidTypes = d

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(REMOVED_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    Dim r As Long, m As String, t As String, k As String
    For r = REMOVED_FIRST_ROW To REMOVED_FORMAT_LAST_ROW
        m = Trim(CStr(ws.Cells(r, REMOVED_RFID_MACHINE).Value))
        t = Trim(CStr(ws.Cells(r, REMOVED_RFID_TYPE).Value))
        If m = "" And t = "" Then Exit For
        If m <> "" And t <> "" Then
            k = RfidFilterKey(m, t)
            If Not d.Exists(k) Then d.Add k, True
        End If
    Next r
End Function


' Rewrites the whole I:J block from what is ticked on the tabs RIGHT NOW.
' A full overwrite, not an append -- that is what makes unticking work.
' MUST run before any tab is cleared. Returns how many filters are active.
Private Function SaveRemovedRfidTypes(ByRef roster As Variant) As Long
    SaveRemovedRfidTypes = 0
    gStep = "Removed Jobs -- saving RFID type filters"

    Dim wsOut As Worksheet
    Set wsOut = EnsureRemovedJobsSheet()
    If wsOut Is Nothing Then Exit Function

    wsOut.Range(wsOut.Cells(REMOVED_FIRST_ROW, REMOVED_RFID_MACHINE), _
                wsOut.Cells(REMOVED_FORMAT_LAST_ROW, REMOVED_RFID_TYPE)).ClearContents

    Dim writeRow As Long: writeRow = REMOVED_FIRST_ROW
    Dim m As Long, tabName As String, ws As Worksheet
    Dim ticked As Object, k As Variant, n As Long
    n = 0

    For m = 1 To UBound(roster, 1)
        tabName = SanitizeSheetName(CStr(roster(m, 1)))
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(tabName)
        On Error GoTo 0
        If Not ws Is Nothing Then
            Set ticked = TickedRfidTypes(ws)
            For Each k In ticked.Keys
                wsOut.Cells(writeRow, REMOVED_RFID_MACHINE).Value = tabName
                wsOut.Cells(writeRow, REMOVED_RFID_TYPE).Value = ticked(k)
                writeRow = writeRow + 1
                n = n + 1
            Next k
        End If
    Next m

    SaveRemovedRfidTypes = n
End Function


Private Function RemovalKey(ByVal jobId As String, ByVal workCenter As String) As String
    RemovalKey = UCase(Trim(jobId)) & "|" & UCase(Trim(workCenter))
End Function


' Walks every machine tab looking for ticked boxes in column L and appends
' what it finds to the Removed Jobs sheet. MUST run before any tab is
' cleared. Returns how many NEW removals were recorded.
Private Function HarvestTickedRemovals(ByRef roster As Variant) As Long
    HarvestTickedRemovals = 0
    gStep = "Removed Jobs -- recording ticked rows"

    Dim already As Object
    Set already = LoadRemovedJobs()

    Dim wsOut As Worksheet
    Set wsOut = EnsureRemovedJobsSheet()
    If wsOut Is Nothing Then Exit Function

    Dim writeRow As Long
    writeRow = wsOut.Cells(wsOut.Rows.Count, 1).End(xlUp).Row + 1
    If writeRow < REMOVED_FIRST_ROW Then writeRow = REMOVED_FIRST_ROW

    Dim m As Long, tabName As String, ws As Worksheet
    Dim lastRow As Long, r As Long
    Dim jobTxt As String, wcTxt As String, k As String
    Dim n As Long: n = 0

    For m = 1 To UBound(roster, 1)
        tabName = SanitizeSheetName(CStr(roster(m, 1)))
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(tabName)
        On Error GoTo 0
        If Not ws Is Nothing Then
            ' 2026-08.25: a ticked TECHNOLOGY is no longer harvested as a
            ' permanent per-job removal -- it is a standing filter handled by
            ' SaveRemovedRfidTypes, and writing its jobs into the permanent
            ' list here would make unticking impossible.
            '
            ' Rows whose L box was ticked BY that filter (the live toggle sets
            ' them for visual feedback) are skipped for the same reason. Only a
            ' box ticked against a job whose technology is NOT filtered counts
            ' as a genuine one-off removal.
            Dim rfidOn As Object
            Set rfidOn = TickedRfidTypes(ws)

            lastRow = ws.Cells(ws.Rows.Count, COL_OUT_JOBID).End(xlUp).Row
            For r = 2 To lastRow
                If TickIsSet(ws.Cells(r, COL_OUT_REMOVE).Value) _
                   And Not rfidOn.Exists(UCase(Trim(CStr(ws.Cells(r, COL_OUT_RFID).Value)))) Then
                    jobTxt = Trim(CStr(ws.Cells(r, COL_OUT_JOBID).Value))
                    wcTxt = Trim(CStr(ws.Cells(r, COL_OUT_WORKCENTER).Value))
                    If jobTxt <> "" And wcTxt <> "" Then
                        k = RemovalKey(jobTxt, wcTxt)
                        If Not already.Exists(k) Then
                            already.Add k, True
                            ' No per-cell NumberFormat here -- column A is
                            ' already text-formatted by EnsureRemovedJobsSheet.
                            wsOut.Cells(writeRow, 1).Value = jobTxt
                            wsOut.Cells(writeRow, 2).Value = wcTxt
                            wsOut.Cells(writeRow, 3).Value = ws.Cells(r, COL_OUT_CUSTOMER).Value
                            wsOut.Cells(writeRow, 4).Value = ws.Cells(r, COL_OUT_DESC).Value
                            wsOut.Cells(writeRow, 5).Value = ws.Cells(r, COL_OUT_QTY).Value
                            wsOut.Cells(writeRow, 6).Value = ws.Cells(r, COL_OUT_SHIP).Value
                            wsOut.Cells(writeRow, 7).Value = Now
                            writeRow = writeRow + 1
                            n = n + 1
                        End If
                    End If
                End If
            Next r
        End If
    Next m

    ' 2026-08.24: the three NumberFormat lines and the AutoFit that used to
    ' live here are gone. The formats are set once, on the columns, by
    ' EnsureRemovedJobsSheet -- recomputing a range and reformatting it on
    ' every harvest bought nothing and was the statement that threw
    ' "Unable to set the NumberFormat property of the Range class" on the
    ' user's Refresh. AutoFit is gone too: with the intro paragraph in A2 it
    ' stretched column A to the width of the whole sentence.

    HarvestTickedRemovals = n
End Function


' A tick is TRUE from a real checkbox, or any non-blank mark. The fallback
' matters: an "x" typed by hand works even on a build of Excel whose VBA
' does not expose SetCheckbox.
Private Function TickIsSet(ByVal v As Variant) As Boolean
    TickIsSet = False
    If IsEmpty(v) Or IsNull(v) Then Exit Function
    If VarType(v) = vbBoolean Then
        TickIsSet = CBool(v)
    ElseIf IsNumeric(v) Then
        TickIsSet = (CDbl(v) <> 0)
    Else
        Dim t As String
        t = UCase(Trim(CStr(v)))
        If t = "" Then Exit Function
        TickIsSet = (t <> "FALSE" And t <> "0" And t <> "NO")
    End If
End Function


Private Function EnsureRemovedJobsSheet() As Worksheet
    Dim ws As Worksheet
    Set ws = GetOrCreateSheet(REMOVED_SHEET)
    If ws Is Nothing Then Exit Function
    Set EnsureRemovedJobsSheet = ws

    ' 2026-08.26: this used to Exit Function here whenever the sheet already
    ' existed, which quietly froze the layout at whatever the build that FIRST
    ' created it happened to write. Everything added later -- the standing
    ' RFID filter block in I:J and its column widths -- therefore never
    ' appeared on any workbook that already had the sheet. Confirmed on the
    ' live file: I3, I4 and J4 were all empty.
    '
    ' The heavy one-time work is still skipped for an existing sheet; only the
    ' parts that must self-heal are re-applied every run. Same principle as the
    ' rest of Preferences -- idempotent, not once-only.
    Dim isNew As Boolean
    isNew = (Trim(CStr(ws.Cells(1, 1).Value)) <> "REMOVED JOBS")

    EnsureRemovedRfidBlock ws
    If Not isNew Then Exit Function

    ' Widths and number formats go on FIRST, while the sheet is still virgin
    ' and carries no merged areas at all. The formats target rows 5+ and the
    ' banners live on rows 1-2, so the two should never interact -- but v35 and
    ' v36 both died on NumberFormat on this sheet and this ordering costs
    ' nothing. SetFmt means a refusal is recorded, not fatal, either way.
    gStep = "Removed Jobs -- column widths"
    ws.Columns(1).ColumnWidth = 14
    ws.Columns(2).ColumnWidth = 16
    ws.Columns(3).ColumnWidth = 26
    ws.Columns(4).ColumnWidth = 46
    ws.Columns(5).ColumnWidth = 10
    ws.Columns(6).ColumnWidth = 12
    ws.Columns(7).ColumnWidth = 18

    ' Job ID stays TEXT so 6016194 and the hyphenated split 6015823-2 are both
    ' preserved verbatim. Not load-bearing: RemovalKey runs CStr() over whatever
    ' comes back, so a Job ID stored as a number still matches.
    gStep = "Removed Jobs -- column number formats"
    SetFmt ws.Range(ws.Cells(REMOVED_FIRST_ROW, 1), ws.Cells(REMOVED_FORMAT_LAST_ROW, 1)), "@", "Removed Jobs Job ID"
    SetFmt ws.Range(ws.Cells(REMOVED_FIRST_ROW, 5), ws.Cells(REMOVED_FORMAT_LAST_ROW, 5)), "#,##0", "Removed Jobs Qty"
    SetFmt ws.Range(ws.Cells(REMOVED_FIRST_ROW, 6), ws.Cells(REMOVED_FORMAT_LAST_ROW, 6)), "mm/dd/yyyy", "Removed Jobs Ship Date"
    ' "hh:mm", NOT "hh:nn". 2026-08-24: this one character class was the whole
    ' NumberFormat failure. "nn" is a VBA Format() minute token and is perfectly
    ' correct in Format(Now, "yyyy-mm-dd hh:nn") elsewhere in this module -- but
    ' Excel's CELL number-format language has no "n" code at all, so the string
    ' is invalid and Excel answers "Unable to set the NumberFormat property of
    ' the Range class". Verified against Microsoft's "Format numbers as dates or
    ' times" reference: hours are h/hh, minutes m/mm, seconds s/ss, and "m" reads
    ' as MINUTES rather than months precisely because it sits straight after
    ' "hh:" here. The other three formats on this sheet were always valid, which
    ' is why only this one was ever refused.
    SetFmt ws.Range(ws.Cells(REMOVED_FIRST_ROW, 7), ws.Cells(REMOVED_FORMAT_LAST_ROW, 7)), "mm/dd/yyyy hh:mm", "Removed Jobs Removed At"

    gStep = "Removed Jobs -- building the header"
    ws.Cells(1, 1).Value = "REMOVED JOBS"
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, 7))
        .Merge
        .Font.Bold = True
        .Font.Size = 14
        .Interior.Color = RGB(192, 0, 0)
        .Font.Color = RGB(255, 255, 255)
    End With

    ws.Cells(2, 1).Value = "Jobs pulled off the schedule by ticking the Remove box on a machine tab " & _
        "and clicking Refresh Data. They stay off for every later Refresh Data. The next UPDATE DATA " & _
        "clears this list completely and brings every job back. To put one job back before then, " & _
        "delete its row here and click Refresh Data."
    ' Merged and wrapped across A2:G2 so the paragraph reads as one block and,
    ' more to the point, so nothing ever AutoFits column A to its full length.
    With ws.Range(ws.Cells(2, 1), ws.Cells(2, 7))
        .Merge
        .WrapText = True
        .VerticalAlignment = xlVAlignTop
    End With
    ws.Cells(2, 1).Font.Italic = True
    ws.Cells(2, 1).Font.Size = 9
    ws.Rows(2).RowHeight = 42

    Dim hdr As Variant
    hdr = Array("Job ID", "Work Center", "Customer", "Job Description", "Qty", "Ship Date", "Removed At")
    Dim c As Long
    For c = 0 To UBound(hdr)
        ws.Cells(REMOVED_HEADER_ROW, c + 1).Value = hdr(c)
    Next c
    With ws.Range(ws.Cells(REMOVED_HEADER_ROW, 1), ws.Cells(REMOVED_HEADER_ROW, 7))
        .Font.Bold = True
        .Interior.Color = RGB(31, 78, 120)
        .Font.Color = RGB(255, 255, 255)
    End With

    ' (the standing-filter block is written by EnsureRemovedRfidBlock, which
    ' runs on every call so it also reaches sheets created before it existed)

End Function


' Update Data wipes the list -- the clean slate the user asked for.
Private Sub ClearRemovedJobs()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(REMOVED_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow >= REMOVED_FIRST_ROW Then
        ws.Range(ws.Cells(REMOVED_FIRST_ROW, 1), ws.Cells(lastRow, 7)).ClearContents
    End If

    ' Standing RFID filters go too. Update Data is the clean slate for every
    ' kind of removal -- each morning starts with the whole queue visible and
    ' nothing hidden that somebody has not decided again today.
    ws.Range(ws.Cells(REMOVED_FIRST_ROW, REMOVED_RFID_MACHINE), _
             ws.Cells(REMOVED_FORMAT_LAST_ROW, REMOVED_RFID_TYPE)).ClearContents
End Sub


' Native cell checkboxes, via Range.CellControl.SetCheckbox. Verified in
' Microsoft's VBA reference before use -- SetCheckbox applies to a whole
' range in one call, so no per-row shapes and nothing to clone from a
' template. Wrapped in On Error because the API is recent: on a build that
' lacks it the column simply stays a plain cell and the typed-mark fallback
' in TickIsSet still works.
Private Sub ApplyRemoveCheckboxes(ByRef ws As Worksheet, ByVal lastRow As Long)
    ' A tab with no job rows is not an attempt and must not be reported as a
    ' failure -- hence the guard BEFORE gCheckboxTried is raised.
    If lastRow < 2 Then Exit Sub
    gCheckboxTried = True

    Dim span As Range
    Set span = ws.Range(ws.Cells(2, COL_OUT_REMOVE), ws.Cells(lastRow, COL_OUT_REMOVE))

    ' SetCheckbox takes a RANGE, so the whole column goes on in one call
    ' instead of one COM round-trip per row -- 1 call rather than ~600 on
    ' Dicetrax. Whether it succeeded is recorded rather than swallowed: on a
    ' build of Excel without cell controls the column silently stayed blank in
    ' v35 and nothing said so. The typed-mark fallback in TickIsSet still
    ' works in that case, and now the finish message says to use it.
    On Error Resume Next
    Err.Clear
    ws.Range(ws.Cells(2, COL_OUT_REMOVE), ws.Cells(lastRow + 400, COL_OUT_REMOVE)).ClearFormats
    span.CellControl.SetCheckbox
    ' Sticky: one tab that could not draw its boxes is worth reporting even if
    ' the tabs after it succeed.
    If Err.Number <> 0 Then gCheckboxOK = False
    Err.Clear
    On Error GoTo 0

    ' Read the Job ID column once. A single-row span comes back as a scalar
    ' rather than an array, so that case is built by hand.
    Dim ids As Variant
    If lastRow = 2 Then
        ReDim ids(1 To 1, 1 To 1)
        ids(1, 1) = ws.Cells(2, COL_OUT_JOBID).Value
    Else
        ids = ws.Range(ws.Cells(2, COL_OUT_JOBID), ws.Cells(lastRow, COL_OUT_JOBID)).Value
    End If

    ' FALSE on every job row, blank on the black day-separator rows -- written
    ' as one array so this is a single write, not one per row.
    Dim vals() As Variant
    ReDim vals(1 To lastRow - 1, 1 To 1)
    Dim r As Long
    Dim blanks As Range
    For r = 2 To lastRow
        If Trim(CStr(ids(r - 1, 1))) <> "" Then
            vals(r - 1, 1) = False
        Else
            vals(r - 1, 1) = Empty
            If blanks Is Nothing Then
                Set blanks = ws.Cells(r, COL_OUT_REMOVE)
            Else
                Set blanks = Union(blanks, ws.Cells(r, COL_OUT_REMOVE))
            End If
        End If
    Next r
    span.Value = vals

    ' A separator row must not carry a box. ClearFormats is the only way to
    ' take a cell control back off -- checked against Microsoft's VBA
    ' reference on 2026-08.19, there is no RemoveCheckbox counterpart.
    On Error Resume Next
    If Not blanks Is Nothing Then blanks.ClearFormats
    span.HorizontalAlignment = xlHAlignCenter
    On Error GoTo 0
End Sub


' =============================================================================
Private Sub ComputeQtySegments(ByRef rawData As Variant, ByRef datedIdx() As Long, _
        ByVal n As Long, ByVal threshold As Double, ByVal lateAfter As Boolean, _
        ByRef segPos() As Long, ByRef segDate() As Date, ByRef segQty() As Double, _
        ByRef segPartial() As Boolean, ByRef segCarry() As Boolean, ByRef segCount As Long)

    segCount = 0
    If n < 1 Then
        ReDim segPos(1 To 1): ReDim segDate(1 To 1): ReDim segQty(1 To 1)
        ReDim segPartial(1 To 1): ReDim segCarry(1 To 1)
        Exit Sub
    End If

    If threshold <= 0 Then threshold = 1

    ' datedIdx is already in effective-schedule-date, Job ID order
    ' (SortByShipThenJobID -> CompareRowsShipThenJobID -> EffectiveSchedule-
    ' Date), which IS the queue order.
    Dim remQty() As Double, doneQty() As Double
    Dim placed() As Boolean, shipSer() As Long, hasShip() As Boolean
    ReDim remQty(1 To n): ReDim doneQty(1 To n)
    ReDim placed(1 To n): ReDim shipSer(1 To n): ReDim hasShip(1 To n)

    Dim i As Long
    Dim shipVal As Variant
    For i = 1 To n
        remQty(i) = SafeEffectiveQty(rawData, datedIdx(i))
        If remQty(i) < 0 Then remQty(i) = 0
        doneQty(i) = 0
        placed(i) = False
        ' shipSer/hasShip drive the same-date backfill below -- "same
        ' date" means same EFFECTIVE schedule date (2026-09.03), so a
        ' To-Perso-driven job only backfills alongside another job due the
        ' same way, never against raw Ship Date.
        shipVal = EffectiveScheduleDate(rawData, datedIdx(i))
        If IsDate(shipVal) Then
            shipSer(i) = CLng(CDate(shipVal))
            hasShip(i) = True
        End If
    Next i

    Dim segCap As Long
    segCap = n + 32
    Dim locPos() As Long, locQty() As Double
    Dim locDayDate() As Date
    Dim locPart() As Boolean, locCarry() As Boolean
    ReDim locPos(1 To segCap): ReDim locQty(1 To segCap)
    ReDim locDayDate(1 To segCap)
    ReDim locPart(1 To segCap): ReDim locCarry(1 To segCap)

    Dim day1 As Date
    day1 = ScheduleDayOne()

    ' No overdue special case any more (2026-08.18). An overdue job has
    ' the earliest ship date in the queue, so the fill below reaches it
    ' first and puts it on the current production day, charged to that
    ' day's capacity like everything else.
    Dim placedCount As Long: placedCount = 0

    ' --- strict-order fill -------------------------------------------
    Dim curDay As Date: curDay = day1
    Dim spaceLeft As Double: spaceLeft = threshold
    Dim guard As Long: guard = 0
    Dim head As Long, headShip As Long, headHasShip As Boolean
    Dim cut As Double

    Do While placedCount < n
        guard = guard + 1
        If guard > 200000 Then Exit Do

        head = 0
        For i = 1 To n
            If Not placed(i) Then
                head = i
                Exit For
            End If
        Next i
        If head = 0 Then Exit Do

        ' ---------------------------------------------------------------
        ' SPLIT AND FILL RULE (2026-08-24, revised same day)
        '
        ' SPLITTING: only a job whose TOTAL quantity exceeds a full day's
        ' capacity may ever be cut. Nothing else splits, for any reason. An
        ' oversized job also starts on a FRESH day rather than squeezing into
        ' today's remainder -- "that job gets its own day" -- then eats whole
        ' days until what is left of it is smaller than a day. That final
        ' partial day is ordinary space again.
        '
        ' FILLING: when the head job cannot be placed, the day does NOT just
        ' end. Any later job SHARING THE HEAD'S EXACT SHIP DATE that fits the
        ' remaining space is placed first. That cannot disturb Ship Date order
        ' -- the dates are equal -- so nothing runs ahead of anything shipping
        ' earlier, while the day still fills up.
        '
        ' The first cut of this rule had no backfill at all and it was wrong
        ' in practice: a real 8/31 ran 15,000 + 500 and then stopped dead,
        ' wasting 19,500, because the next job in queue order was 30,000 --
        ' too big for the gap, too small to be allowed to split. Six other
        ' jobs shipping THAT SAME 09/01 date were sitting right behind it and
        ' would have fitted. Measured over the live 665-job Dicetrax queue:
        ' no backfill = 50 days at 86.6% fill; this rule = 46 days at 94.1%,
        ' still zero splits and still ZERO segments out of Ship Date order.
        ' Backfilling from any ship date would reach 44 days / 98.4% but put
        ' 34 segments ahead of earlier-shipping work, which is not allowed.
        ' ---------------------------------------------------------------
        Dim isOversized As Boolean
        Dim endDay As Boolean
        isOversized = (SafeEffectiveQty(rawData, datedIdx(head)) > threshold + 0.0000001)
        endDay = False

        If isOversized And doneQty(head) <= 0.0000001 And spaceLeft < threshold - 0.0000001 Then
            ' Its own day: a big job never begins in the tail of a part-used
            ' day. Fill what is left, then roll forward so it starts clean.
            endDay = True

        ElseIf remQty(head) <= spaceLeft + 0.0000001 Then
            ' Fits whole. A slice of an already-cut job is still marked
            ' partial so it keeps the split styling.
            EmitSegDate locPos, locDayDate, locQty, locPart, locCarry, _
                        segCount, segCap, head, curDay, remQty(head), _
                        (doneQty(head) > 0.0000001), (doneQty(head) > 0.0000001)
            spaceLeft = spaceLeft - remQty(head)
            doneQty(head) = doneQty(head) + remQty(head)
            remQty(head) = 0
            placed(head) = True
            placedCount = placedCount + 1
            If spaceLeft <= 0.0000001 Then
                curDay = NextBusinessDayAfter(curDay)
                spaceLeft = threshold
            End If

        ElseIf isOversized Then
            ' Bigger than a whole day: take this day in full and carry the
            ' rest into the next one. This is the ONLY place a job is cut.
            cut = spaceLeft
            EmitSegDate locPos, locDayDate, locQty, locPart, locCarry, _
                        segCount, segCap, head, curDay, cut, True, _
                        (doneQty(head) > 0.0000001)
            remQty(head) = remQty(head) - cut
            doneQty(head) = doneQty(head) + cut
            curDay = NextBusinessDayAfter(curDay)
            spaceLeft = threshold

        Else
            ' Fits in a day but not in what is left of THIS one. It is not
            ' cut; it leads tomorrow. Fill today from its own ship date first.
            endDay = True
        End If

        If endDay Then
            ' Same-ship-date backfill. One pass is enough: spaceLeft only
            ' ever shrinks, so a job rejected now cannot fit later in the
            ' same pass. An oversized job can never qualify -- its quantity
            ' exceeds a whole day, so it cannot fit a partial one.
            headShip = shipSer(head)
            headHasShip = hasShip(head)
            If headHasShip Then
                For i = head + 1 To n
                    If Not placed(i) Then
                        If hasShip(i) Then
                            If shipSer(i) = headShip Then
                                If remQty(i) <= spaceLeft + 0.0000001 Then
                                    EmitSegDate locPos, locDayDate, locQty, locPart, locCarry, _
                                                segCount, segCap, i, curDay, remQty(i), _
                                                (doneQty(i) > 0.0000001), (doneQty(i) > 0.0000001)
                                    spaceLeft = spaceLeft - remQty(i)
                                    doneQty(i) = doneQty(i) + remQty(i)
                                    remQty(i) = 0
                                    placed(i) = True
                                    placedCount = placedCount + 1
                                End If
                            End If
                        End If
                    End If
                Next i
            End If

            curDay = NextBusinessDayAfter(curDay)
            spaceLeft = threshold
        End If
    Loop

    ' --- hand back in the shape the caller expects -------------------
    If segCount < 1 Then
        ReDim segPos(1 To 1): ReDim segDate(1 To 1): ReDim segQty(1 To 1)
        ReDim segPartial(1 To 1): ReDim segCarry(1 To 1)
        Exit Sub
    End If

    ReDim segPos(1 To segCount): ReDim segDate(1 To segCount): ReDim segQty(1 To segCount)
    ReDim segPartial(1 To segCount): ReDim segCarry(1 To segCount)
    For i = 1 To segCount
        segPos(i) = locPos(i)
        segDate(i) = locDayDate(i)
        segQty(i) = locQty(i)
        segPartial(i) = locPart(i)
        segCarry(i) = locCarry(i)
    Next i
End Sub


' Segment appender. The piece-range fields that shipped on 2026-08.17b
' were removed on 2026-08.18 along with the display that used them, rather
' than left threaded through four procedures with no consumer.
Private Sub EmitSegDate(ByRef locPos() As Long, ByRef locDayDate() As Date, ByRef locQty() As Double, _
        ByRef locPart() As Boolean, ByRef locCarry() As Boolean, _
        ByRef segCount As Long, ByRef segCap As Long, _
        ByVal pos As Long, ByVal dayDate As Date, ByVal q As Double, _
        ByVal isPart As Boolean, ByVal isCarry As Boolean)
    If segCount >= segCap Then
        segCap = segCap * 2
        ReDim Preserve locPos(1 To segCap): ReDim Preserve locDayDate(1 To segCap)
        ReDim Preserve locQty(1 To segCap): ReDim Preserve locPart(1 To segCap)
        ReDim Preserve locCarry(1 To segCap)
    End If
    segCount = segCount + 1
    locPos(segCount) = pos
    locDayDate(segCount) = dayDate
    locQty(segCount) = q
    locPart(segCount) = isPart
    locCarry(segCount) = isCarry
End Sub

Private Sub ComputeCountSegments(ByRef rawData As Variant, ByRef datedIdx() As Long, _
        ByVal n As Long, ByVal threshold As Double, _
        ByRef segPos() As Long, ByRef segDate() As Date, ByRef segQty() As Double, _
        ByRef segPartial() As Boolean, ByRef segCarry() As Boolean, ByRef segCount As Long)
    Dim prodDays() As Date
    prodDays = ComputeProductionDays(rawData, datedIdx, n, threshold)
    ReDim segPos(1 To n): ReDim segDate(1 To n): ReDim segQty(1 To n)
    ReDim segPartial(1 To n): ReDim segCarry(1 To n)
    Dim i As Long
    For i = 1 To n
        segPos(i) = i
        segDate(i) = prodDays(i)
        segQty(i) = SafeEffectiveQty(rawData, datedIdx(i))
        segPartial(i) = False
        segCarry(i) = False
    Next i
    segCount = n
End Sub


' Maps a job's own Ship Date to the "actual day number" scale this scheduler
' uses internally (actual day 1 = NextBusinessDay(Date), actual day 2 = the
' business day after that, etc.) so a job's real deadline can be compared
' directly against whichever actual day is currently being solved. A ship
' date that has already passed, or falls on/before actual day 1's calendar
' date, maps to day 1 -- already due, never allowed to drift further.
Private Function ShipDateToActualDayNumber(ByVal shipDate As Date) As Long
    Dim d As Date: d = ScheduleDayOne()
    Dim n As Long: n = 1
    Do While d < shipDate
        d = NextBusinessDayAfter(d)
        n = n + 1
    Loop
    ShipDateToActualDayNumber = n
End Function

' 2026-08.13: RULE 2 (hyphen families run together) WAS REMOVED at the
' user's direction, so this now emits ONE GROUP PER JOB -- every job
' schedules independently. The CSR group structure is kept (rather than
' ripped out) because the scheduler, the oversized-split allocator and
' the pool bookkeeping are all written against it; singleton groups cost
' nothing and leave that machinery untouched and still able to express
' multi-job groups should grouping ever come back.
'
' DEADLINE (2026-08.13): a group's deadline is now its own SHIP DAY when
' the machine's Late Rule is AFTER (running ON the ship date is on time),
' or ship day - 1 under the legacy ON/AFTER rule (must finish strictly
' before). Previously always ship day - 1.
'
' Historic note -- the old grouping was by (PARSED BASE STRING,
' SHIP DAY) -- 2026-08.12 rule change: family members sharing a base AND a
' ship date are atomic (one group, one day); members of the same family
' with DIFFERENT ship dates are now independent groups, each against its
' own deadline (previously the whole family was atomic on the EARLIEST
' member's deadline, which forced e.g. a 30,000-qty part due 8/24 and a
' 20,000-qty part not due until November onto one 50,000 block that could
' never fit under the 35,000 threshold on any day). Grouping is still by
' parsed base (ParseJobIdBaseSuffix), never by "does this job have a
' suffix" (see the header note above ComputeProductionDaysQty for the real
' bare-root bug that distinction fixes). Every job belongs to exactly one
' group, so total membership across all groups is always exactly n.
' CSR-style flat output (no Collection/Dictionary):
'   groupMemberStart(1 To numGroups+1) -- group g's members are job indices
'     memberFlat(groupMemberStart(g) To groupMemberStart(g+1)-1).
'   groupQty(1 To numGroups) -- total qty across the group's members.
'   groupDeadline(1 To numGroups) -- the LAST actual day the group may run
'     on = (the group's shared ship day) - 1 -- STRICTLY before the ship
'     date (2026-08.9 correction: was the ship day itself, i.e. "on or
'     before", until the user clarified "prior to" means strictly earlier).
'     All members of a group share one ship day by construction now, so
'     the old "earliest member's deadline" scan is gone. A group whose
'     ship day is actual day 1 gets deadline 0 -- always mandatory starting
'     day 1, since there is no day before day 1; this is a genuine,
'     unavoidable "already due" case (e.g. a Ship Date that has already
'     passed), not a bug.
' Quantity a job consumes against a machine's daily threshold. Hub stopped
' shipping an Effective QTY column on 2026-08.12 (rate math moved to the
' AVL Production Dashboard) and column K is RFID Type as of 2026-08.17, so
' this is now simply the raw order quantity -- which is what every Press
' machine's threshold has always effectively used, since no Press machine
' has a Convert-QTY-to-Sheets factor.
Private Function SafeEffectiveQty(ByRef rawData As Variant, ByVal rowIdx As Long) As Double
    If IsNumeric(rawData(rowIdx, COL_QTY)) Then
        SafeEffectiveQty = CDbl(rawData(rowIdx, COL_QTY))
    Else
        SafeEffectiveQty = 0
    End If
End Function

' =============================================================================
' Day one of the schedule.
'
' Normally today, or the next weekday if today is a weekend. The Production
' Cutoff setting on Preferences moves it on by one: past that time of day the
' press floor can no longer start today's work, so the schedule rolls forward.
' Confirmed with the user 2026-08-24 -- "the time at which the schedule
' increments current day +1". Until then the setting sat on the control panel
' being read by nothing at all.
'
' The bump is only applied when TODAY is itself a business day. On a Saturday
' there is no production window left to miss, so pushing Monday out to Tuesday
' would be wrong.
'
' A blank, malformed or missing setting simply means no cutoff.
' =============================================================================
Private Function ScheduleDayOne() As Date
    Dim d As Date
    d = NextBusinessDay(Date)
    ScheduleDayOne = d

    If d <> Date Then Exit Function     ' today is a weekend -- nothing to miss

    Dim raw As Variant
    raw = GetPrefValue("Production Cutoff")
    If IsEmpty(raw) Then Exit Function

    ' A time-formatted cell comes back as a Date; a General one as a Double.
    Dim cutoffFrac As Double
    If IsDate(raw) Then
        cutoffFrac = CDbl(CDate(raw))
    ElseIf IsNumeric(raw) Then
        cutoffFrac = CDbl(raw)
    Else
        Exit Function
    End If
    cutoffFrac = cutoffFrac - Int(cutoffFrac)
    If cutoffFrac <= 0 Then Exit Function

    If (CDbl(Now) - Int(CDbl(Now))) >= cutoffFrac Then
        ScheduleDayOne = NextBusinessDayAfter(d)
    End If
End Function

Private Function NextBusinessDay(ByVal d As Date) As Date
    Do While Weekday(d, vbSunday) = 1 Or Weekday(d, vbSunday) = 7
        d = d + 1
    Loop
    NextBusinessDay = d
End Function

Private Function NextBusinessDayAfter(ByVal d As Date) As Date
    d = d + 1
    NextBusinessDayAfter = NextBusinessDay(d)
End Function
Private Function ReadRowFormats() As Variant
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Preferences")
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    Dim anchorRow As Long: anchorRow = 0
    Dim r As Long
    For r = 1 To 200
        If Trim(CStr(ws.Cells(r, 1).Value)) = "Row Formatting" Then
            anchorRow = r
            Exit For
        End If
    Next r
    If anchorRow = 0 Then Exit Function

    ' anchor, instructions, header, then one row per state in FMT_* order.
    Dim firstDataRow As Long: firstDataRow = anchorRow + 3
    Dim fmts(1 To 5, 1 To 6) As Variant
    Dim i As Long, sw As Range
    For i = 1 To 5
        Set sw = ws.Cells(firstDataRow + i - 1, 2)
        If sw.Interior.ColorIndex = xlNone Then
            fmts(i, 1) = -1&
        Else
            fmts(i, 1) = CLng(sw.Interior.Color)
        End If
        fmts(i, 2) = CStr(sw.Font.name)
        fmts(i, 3) = CDbl(sw.Font.Size)
        fmts(i, 4) = (sw.Font.Bold = True)
        fmts(i, 5) = (sw.Font.Italic = True)
        fmts(i, 6) = CLng(sw.Font.Color)
    Next i
    ReadRowFormats = fmts
End Function

' Applies one state's stored style across a written row's data columns.
' Silently no-ops when the Row Formatting table is missing, so a older
' Preferences sheet still produces a readable (if unstyled) tab.
Private Sub ApplyRowFormat(ByRef ws As Worksheet, ByVal rowNum As Long, ByVal lastCol As Long, _
        ByRef fmts As Variant, ByVal stateIdx As Long)
    If IsEmpty(fmts) Then Exit Sub
    Dim rng As Range
    Set rng = ws.Range(ws.Cells(rowNum, 1), ws.Cells(rowNum, lastCol))
    If CLng(fmts(stateIdx, 1)) = -1 Then
        rng.Interior.Pattern = xlNone
    Else
        rng.Interior.Color = CLng(fmts(stateIdx, 1))
    End If
    With rng.Font
        If CStr(fmts(stateIdx, 2)) <> "" Then .name = CStr(fmts(stateIdx, 2))
        If CDbl(fmts(stateIdx, 3)) > 0 Then .Size = CDbl(fmts(stateIdx, 3))
        .Bold = CBool(fmts(stateIdx, 4))
        .Italic = CBool(fmts(stateIdx, 5))
        .Color = CLng(fmts(stateIdx, 6))
    End With
End Sub

' Creates the Row Formatting swatch table once, seeded with the 2026-08.13
' defaults. Never rewritten afterwards -- the whole point is that the user
' restyles the swatches and those choices survive every future run.
Private Sub EnsureRowFormatTable(ByRef ws As Worksheet)
    Dim r As Long
    For r = 1 To 200
        If Trim(CStr(ws.Cells(r, 1).Value)) = "Row Formatting" Then Exit Sub
    Next r

    Dim lastA As Long
    lastA = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 2
    With ws.Cells(lastA, 1)
        .Value = "Row Formatting"
        .Font.Bold = True
    End With
    ws.Cells(lastA + 1, 1).Value = "Format the SWATCH cell in column B however you like -- fill colour, font, size, bold/italic, font colour -- and every matching row on every machine tab picks it up on the next Update Data. No fill on a swatch means those rows stay unfilled. Do not rename the labels in column A; they are how the macro finds each style."
    ws.Cells(lastA + 1, 1).Font.Size = 9
    ws.Cells(lastA + 1, 1).Font.Italic = True

    Dim headerRow As Long: headerRow = lastA + 2
    ws.Cells(headerRow, 1).Value = "Row State"
    ws.Cells(headerRow, 2).Value = "Swatch (format this cell)"
    ws.Cells(headerRow, 3).Value = "Applies to"
    With ws.Range(ws.Cells(headerRow, 1), ws.Cells(headerRow, 3))
        .Font.Bold = True
        .Interior.Color = RGB(220, 230, 241)
    End With

    Dim d As Long: d = headerRow + 1
    SeedFormatSwatch ws, d + 0, "Normal Job", "Runs before its Ship Date, whole, within capacity.", _
        -1, RGB(0, 0, 0), False
    SeedFormatSwatch ws, d + 1, "Ships Today", "Runs ON its Ship Date -- on time, but no slack left.", _
        RGB(255, 235, 156), RGB(0, 0, 0), False
    SeedFormatSwatch ws, d + 2, "Late Job", "Runs AFTER its Ship Date.", _
        RGB(255, 199, 206), RGB(0, 0, 0), True
    SeedFormatSwatch ws, d + 3, "Split Job", "One day's slice of a job too big to run in a single day.", _
        RGB(189, 215, 238), RGB(0, 0, 0), False
    SeedFormatSwatch ws, d + 4, "Split + Late Job", "A split slice that also runs after its Ship Date.", _
        RGB(189, 215, 238), RGB(255, 0, 0), True

    ws.Columns(1).ColumnWidth = 28
End Sub

Private Sub SeedFormatSwatch(ByRef ws As Worksheet, ByVal r As Long, ByVal label As String, _
        ByVal appliesTo As String, ByVal fillColor As Long, ByVal fontColor As Long, ByVal isBold As Boolean)
    ws.Cells(r, 1).Value = label
    ws.Cells(r, 2).Value = "Sample"
    ws.Cells(r, 3).Value = appliesTo
    With ws.Cells(r, 2)
        If fillColor = -1 Then
            .Interior.Pattern = xlNone
        Else
            .Interior.Color = fillColor
        End If
        .Font.Color = fontColor
        .Font.Bold = isBold
        .HorizontalAlignment = xlHAlignCenter
        .BorderAround Weight:=xlThin
    End With
End Sub
' Sizes a row to fit its wrapped text. Excel's AutoFit handles wrapped,
' unmerged cells correctly; the floor guarantees at least the requested
' height (so a short line still gets a double-height row when asked for).
' Used by the Rules and Onboarding builders (2026-08.13 request: 'if you
' need to make rows double sized to accommodate word wrapping have the
' code do that').
Private Sub FitWrappedRow(ByRef ws As Worksheet, ByVal r As Long, ByVal minHeight As Double)
    On Error Resume Next
    ws.Rows(r).AutoFit
    If ws.Rows(r).RowHeight < minHeight Then ws.Rows(r).RowHeight = minHeight
    On Error GoTo 0
End Sub

' =============================================================================
' ONBOARDING SHEET -- rebuilt from scratch on every Update Data run, exactly
' like the Rules tab. It used to be hand-typed static content, which went
' stale the moment the layout and rules changed (it still described the
' Promised Date column and a 3-column Machine Roster after both had moved
' on). Generating it from code is the only way it stays true.
'
' STRUCTURE NOTE (2026-08.13): the content is emitted one OnbRow call per
' line rather than as a single big Array literal. VBA allows at most 25
' line continuations in one logical statement, and the array form blew
' straight past that with a 'Too many line continuations' compile error --
' which LibreOffice's compile check does NOT catch. Keep it one call per
' row; never collect this content back into a continued literal.
' =============================================================================
Private Sub BuildOnboardingSheet()
    Dim ws As Worksheet
    Set ws = GetOrCreateSheet("Onboarding")
    If ws Is Nothing Then Exit Sub
    gStep = "rebuilding the Onboarding sheet"
    ws.Cells.Clear
    ws.Cells.UnMerge
    ws.Rows.RowHeight = ws.StandardHeight

    ws.Columns(1).ColumnWidth = 3
    ws.Columns(2).ColumnWidth = 62
    ws.Columns(3).ColumnWidth = 44
    ws.Columns(4).ColumnWidth = 44

    Dim r As Long: r = 1
    ws.Cells(r, 2).Value = "PLI PRESS - ONBOARDING & QUICK REFERENCE"
    With ws.Range(ws.Cells(r, 2), ws.Cells(r, 4))
        .Merge
        .Font.Bold = True
        .Font.Size = 14
        .Interior.Color = RGB(31, 78, 120)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlHAlignCenter
    End With
    ws.Rows(r).RowHeight = 22
    r = r + 2

    r = OnbRow(ws, r, "H", "1. What This Workbook Is", "", "")
    r = OnbRow(ws, r, "B", "Press is a satellite workbook. It does not read Monarch exports directly -- it pulls already-cleaned production data from PLI Hub.xlsm, filters it down to the machines listed on Preferences, schedules that work, and rebuilds one tab per machine.", "", "")
    r = OnbRow(ws, r, "B", "The pipeline has two hops: the shared Monarch export folder feeds Hub, and Hub feeds Press. If the data here looks stale, refresh Hub first, then click Update Data here.", "", "")
    r = OnbRow(ws, r, "B", "Since 2026-08-13 the Hub no longer calculates machine rates or capacity -- that maths lives in the AVL Production Dashboard now. Press only ever needed Hub's cleaned job columns, so nothing here depends on it.", "", "")
    r = OnbRow(ws, r, "", "", "", "")
    r = OnbRow(ws, r, "H", "2. First-Time Setup (once per new copy of this file)", "", "")
    r = OnbRow(ws, r, "B", "1.  Open the Preferences tab. Press always lands there when Update Data finishes.", "", "")
    r = OnbRow(ws, r, "B", "2.  Set Hub Workbook Path to the full path of PLI Hub.xlsm on the shared drive, for example \\server\share\PLI Hub.xlsm.", "", "")
    r = OnbRow(ws, r, "B", "3.  Click the Update Data button. On a brand-new copy the button does not exist yet -- run it once via Alt+F8, pick UpdatePressData, and click Run. The button is created for you and every later run is one click.", "", "")
    r = OnbRow(ws, r, "B", "4.  Everything else on Preferences is seeded with working defaults. Adjust it whenever you like; your edits are never overwritten.", "", "")
    r = OnbRow(ws, r, "", "", "", "")
    r = OnbRow(ws, r, "H", "3. What The Update Data Button Actually Does", "", "")
    r = OnbRow(ws, r, "B", "-  Opens PLI Hub.xlsm read-only. This never locks Hub for anyone else.", "", "")
    r = OnbRow(ws, r, "B", "-  Checks Hub's own status first and refuses to proceed if Hub is REFRESHING or its last run errored, so Press never shows half-built data.", "", "")
    r = OnbRow(ws, r, "B", "-  Pulls Hub's CleanedData, then for every Machine Roster row keeps the jobs whose Work Center contains that row's match text AND whose Location matches the Location Filter list.", "", "")
    r = OnbRow(ws, r, "B", "-  Splits those into dated and undated jobs, schedules the dated ones by that machine's rule, and rebuilds the tab from scratch.", "", "")
    r = OnbRow(ws, r, "B", "-  Repaints every row from the Row Formatting swatches, then rebuilds the Rules and Onboarding sheets and reorders the tabs to match the roster.", "", "")
    r = OnbRow(ws, r, "B", "-  Recomputes everything fresh on every click. Nothing carries over between runs, so a job's Production Day can move when the data or the settings change.", "", "")
    r = OnbRow(ws, r, "", "", "", "")
    r = OnbRow(ws, r, "H", "4. The Preferences Sheet, Field By Field", "", "")
    r = OnbRow(ws, r, "B", "Hub Workbook Path -- the full path to PLI Hub.xlsm.", "", "")
    r = OnbRow(ws, r, "B", "Press Status -- always one of three words. Ready means the last pull succeeded, and the detail after the dash is the Monarch export's own timestamp, not the time you clicked. Not Ready means a pull is running or Hub is still refreshing, and resolves on its own. Error means the pull failed, with the reason after the dash.", "", "")
    r = OnbRow(ws, r, "B", "Machine Roster -- one row per machine tab. No code change is ever needed to add, edit or remove one.", "", "")
    r = OnbRow(ws, r, "K", "Roster column", "What it means", "Notes")
    r = OnbRow(ws, r, "T", "Work Center Match", "Text to look for anywhere in a job's Work Center.", "Not case- or spacing-sensitive.")
    r = OnbRow(ws, r, "T", "Rule Type", "COUNT = a fixed number of jobs per day. QTY = jobs packed up to a daily quantity.", "Dicetrax is the only QTY machine today.")
    r = OnbRow(ws, r, "T", "Daily Threshold", "The job count for COUNT, or the quantity ceiling for QTY.", "A day never exceeds this.")
    r = OnbRow(ws, r, "T", "Late Rule", "AFTER = late only once a job runs past its Ship Date. ON/AFTER = running on the ship date already counts as late.", "Dicetrax and Metronics use AFTER; Indigo uses ON/AFTER.")
    r = OnbRow(ws, r, "B", "Removing a roster row does NOT delete that tab. Hide or delete the tab by hand if you want it gone.", "", "")
    r = OnbRow(ws, r, "B", "Location Filter -- a job only reaches a machine tab if its Location contains at least one of these phrases, on top of the Work Center match. Delete every row to switch this filter off entirely.", "", "")
    r = OnbRow(ws, r, "B", "Scheduling Settings -- GONE as of 2026-08-17. Pull-Forward Days, Extra Days If Under-Filled and Max Under-Capacity % were deleted along with the policy behind them. Days are now filled sequentially by Ship Date, as full as they will go, with no window and no fill floor.", "", "")
    r = OnbRow(ws, r, "B", "Row Formatting -- five swatch cells, one per row state. Format a swatch with ordinary Excel tools (fill colour, font, size, bold, italic, font colour) and every matching row on every machine tab picks it up on the next Update Data. A swatch left with no fill means those rows stay unfilled.", "", "")
    r = OnbRow(ws, r, "", "", "", "")
    r = OnbRow(ws, r, "H", "5. Reading A Machine Tab", "", "")
    r = OnbRow(ws, r, "B", "Columns, left to right: To Perso, Ship Date, Job ID, Customer Name, Job Description, Qty, RFID, Last Location, Work Center, Production Day, Total Qty, Remove. RFID was added on 2026-08-17 and comes straight from Hub -- Hub reads each job description against its RFID Technologies list and Press only displays the answer, so a blank means the description matched none of Hub's tokens. Promised Date was removed on 2026-08-13. To Perso was added 2026-09-03, also straight from Hub -- blank unless the job has reached that stage.", "", "")
    r = OnbRow(ws, r, "B", "Wherever this tab and the Rules tab say 'Ship Date' below, the scheduler actually means a job's EFFECTIVE date: To Perso when that job has one, otherwise Ship Date. Ship Date and To Perso still each show their own real value in their own column -- only the scheduling math (queue order, same-date backfill, and whether a job reads late or on time) substitutes one for the other.", "", "")
    r = OnbRow(ws, r, "B", "Production Day is a SCHEDULING PROJECTION, not a date that came from the job. It is what this workbook expects the machine to run that work, starting from the next business day. Weekends are skipped.", "", "")
    r = OnbRow(ws, r, "B", "Total Qty is one merged figure per production day: everything that day runs. Black bars separate one day from the next.", "", "")
    r = OnbRow(ws, r, "B", "Row colours tell you the state of each job at a glance. All five are editable in Row Formatting on Preferences, and the Rules tab shows a live legend painted from those same swatches.", "", "")
    r = OnbRow(ws, r, "K", "Row appearance", "Meaning", "Watch for")
    r = OnbRow(ws, r, "T", "Plain", "Runs before its Ship Date, whole, inside capacity.", "Nothing -- this is healthy.")
    r = OnbRow(ws, r, "T", "Ships-today fill", "Runs exactly ON its Ship Date. Still on time.", "No slack left; any slip makes it late.")
    r = OnbRow(ws, r, "T", "Late fill", "Runs AFTER its Ship Date.", "Already overdue, or pushed past the date by capacity.")
    r = OnbRow(ws, r, "T", "Split fill", "One day's slice of a job cut by the capacity line, so the day stays full and Ship Date order holds.", "The Qty column shows only what runs THAT day.")
    r = OnbRow(ws, r, "T", "Split + late fill", "A split slice that also lands after the Ship Date.", "The most urgent thing on the tab.")
    r = OnbRow(ws, r, "B", "Jobs with no valid effective date (no To Perso date AND no valid Ship Date) collect at the bottom of the tab in italics. They are not scheduled and have no Production Day or Total Qty.", "", "")
    r = OnbRow(ws, r, "", "", "", "")
    r = OnbRow(ws, r, "H", "6. The Scheduling Rules", "", "")
    r = OnbRow(ws, r, "B", "The Rules tab holds the full, current rule set with a worked description of each one. It is rebuilt on every Update Data, so it can never drift from what the code actually does. In short: the queue runs in Ship Date order and NEVER out of it, each day is filled as full as it will go, a job that does not fit is cut into the space left rather than jumping the queue, overdue work runs today and counts against today, and no day exceeds its threshold.", "", "")
    r = OnbRow(ws, r, "", "", "", "")
    r = OnbRow(ws, r, "H", "7. Troubleshooting", "", "")
    r = OnbRow(ws, r, "K", "Status shown", "What it means", "What to do")
    r = OnbRow(ws, r, "T", "Not Ready - pull in progress", "A click is actively running.", "Just wait. Do not click again.")
    r = OnbRow(ws, r, "T", "Not Ready - Hub is REFRESHING", "Someone clicked Update on Hub and it has not finished.", "Wait a moment, then click Update Data again.")
    r = OnbRow(ws, r, "T", "Error - Hub file not found", "The Hub Workbook Path is wrong, or the shared drive is unreachable.", "Fix the path and confirm you can browse to that drive.")
    r = OnbRow(ws, r, "T", "Error - Could not open Hub workbook", "Hub is open or mid-save somewhere else.", "Try again shortly.")
    r = OnbRow(ws, r, "T", "Error - Hub's last build failed", "Hub itself has a problem. Press cannot fix it from here.", "Open PLI Hub.xlsm, resolve its error, run its Update Data, then retry here.")
    r = OnbRow(ws, r, "T", "Error - Hub has no CleanedData / it was empty", "Hub has not successfully built data yet.", "Run Hub's Update Data first.")
    r = OnbRow(ws, r, "T", "Error - Machine Roster is empty", "There is nothing to build.", "Add at least one roster row and try again.")
    r = OnbRow(ws, r, "T", "A tab looks emptier than expected", "The Location Filter or Work Center match is excluding more than you meant.", "Check both lists on Preferences; widening either brings more jobs back.")
    r = OnbRow(ws, r, "T", "Work runs weeks before its Ship Date", "Expected since 2026-08-17. Days are packed as full as they will go, so where there is more capacity than queue the whole backlog compresses forward.", "Nothing to change -- it is the rule. Raise a Daily Threshold to compress further, lower it to spread the work out.")
    r = OnbRow(ws, r, "", "", "", "")
    r = OnbRow(ws, r, "H", "8. Removing Jobs From The Schedule", "", "")
    r = OnbRow(ws, r, "B", "Every job row carries an unticked REMOVE box in column L after Update Data.", "", "")
    r = OnbRow(ws, r, "K", "Step", "What you do", "What happens")
    r = OnbRow(ws, r, "T", "1", "Click Update Data.", "Fresh jobs from Hub. Every box is clear and nothing is removed.")
    r = OnbRow(ws, r, "T", "2", "Tick the Remove box next to any jobs that should not run.", "Nothing happens yet -- the tick is only a mark.")
    r = OnbRow(ws, r, "T", "3", "Click Refresh Data.", "Those jobs come off, the schedule rebuilds around the freed capacity, and each one is logged on the Removed Jobs tab.")
    r = OnbRow(ws, r, "T", "4", "Repeat 2 and 3 as often as you like.", "Removals add up through the day.")
    r = OnbRow(ws, r, "T", "Undo one", "Delete its row on the Removed Jobs tab, then Refresh Data.", "That job returns; everything else stays removed.")
    r = OnbRow(ws, r, "T", "Undo all", "Click Update Data.", "The list is cleared and every job comes back.")
    r = OnbRow(ws, r, "B", "Removal is per MACHINE -- pulling a job off Dicetrax leaves it on Metronics and Indigo. Ticking any slice of a split job removes the whole job from that machine.", "", "")
    r = OnbRow(ws, r, "B", "If your Excel is too old to draw the checkbox, column L stays an ordinary cell -- typing anything in it (an x will do) counts as ticked.", "", "")
    r = OnbRow(ws, r, "", "", "", "")
    r = OnbRow(ws, r, "H", "9. What Changed On 2026-08-17", "", "")
    r = OnbRow(ws, r, "B", "-  The pull-forward window, the extra-days reach and the under-capacity floor were ALL removed, and the Scheduling Settings table deleted from Preferences along with them.", "", "")
    r = OnbRow(ws, r, "B", "-  Days are filled sequentially by Ship Date, as full as they will go. Ship Date order is absolute: when a job does not fit, only smaller jobs of the SAME ship date may fill ahead of it, and then the job itself is cut into the space left.", "", "")
    r = OnbRow(ws, r, "B", "-  Job ID is now irrelevant to scheduling. There is no family rule and no grouping of any kind.", "", "")
    r = OnbRow(ws, r, "B", "-  Overdue jobs run on the current production day and COUNT against its threshold (changed 2026-08-18). Day one therefore schedules one full threshold in total, and no day can read above its threshold any more.", "", "")
    r = OnbRow(ws, r, "B", "-  Work now runs as early as capacity allows. On the 2026-08-17 Dicetrax queue that compressed 138 production days into 36, with two thirds of the quantity finishing more than a fortnight before it ships. This is intended.", "", "")
    r = OnbRow(ws, r, "B", "-  Split rows show the quantity run THAT DAY in the Qty column. The remainder carried in from yesterday leads its day; the slice cut by the capacity line closes it.", "", "")
    r = OnbRow(ws, r, "B", "-  A new RFID column sits between Qty and Last Location, filled from Hub's RFID Type column.", "", "")
    r = OnbRow(ws, r, "", "", "", "")
    r = OnbRow(ws, r, "H", "9. What Changed On 2026-08-13", "", "")
    r = OnbRow(ws, r, "B", "-  The old rule that forced hyphen-linked Job IDs onto the same day was REMOVED.", "", "")
    r = OnbRow(ws, r, "B", "-  A job is late only once it runs AFTER its Ship Date. Running on the ship date is on time and gets its own colour. Indigo keeps the older, stricter reading.", "", "")
    r = OnbRow(ws, r, "B", "-  The Promised Date column was removed.", "", "")
    r = OnbRow(ws, r, "B", "-  All row colours, fonts and weights moved into the Row Formatting table on Preferences.", "", "")
    ' Paint the five sample states in section 5 with the user's own
    ' swatches so the colour names above are never guesswork.
    Dim rowFmts As Variant: rowFmts = ReadRowFormats()
    Dim sr As Long
    For sr = 1 To r
        Select Case Trim(CStr(ws.Cells(sr, 2).Value))
        Case "Plain"
            ApplyRowFormat ws, sr, 4, rowFmts, FMT_NORMAL
        Case "Ships-today fill"
            ApplyRowFormat ws, sr, 4, rowFmts, FMT_SHIPDAY
        Case "Late fill"
            ApplyRowFormat ws, sr, 4, rowFmts, FMT_LATE
        Case "Split fill"
            ApplyRowFormat ws, sr, 4, rowFmts, FMT_SPLIT
        Case "Split + late fill"
            ApplyRowFormat ws, sr, 4, rowFmts, FMT_SPLIT_LATE
        End Select
    Next sr

    ' NOTE: do NOT add ws.Cells(1, 1).Select here. Excel cannot Select a
    ' range on a sheet that is not the active sheet, and this sub runs
    ' while another sheet is active -- it raised run-time error 1004
    ' (2026-08.13). The two other Select/Activate calls in this module are
    ' deliberately wrapped in On Error Resume Next for the same reason.
End Sub

' Writes one Onboarding row in the style its kind calls for and returns the
' next free row. Kinds: H = section header, B = body paragraph (merged
' B:D, wrapped, height computed from the text length because merged cells
' defeat AutoFit), K = table header, T = table row, anything else = blank.
Private Function OnbRow(ByRef ws As Worksheet, ByVal r As Long, ByVal kind As String, _
        ByVal bText As String, ByVal cText As String, ByVal dText As String) As Long
    Select Case kind
    Case "H"
        ws.Cells(r, 2).Value = bText
        With ws.Range(ws.Cells(r, 2), ws.Cells(r, 4))
            .Merge
            .Font.Bold = True
            .Font.Size = 12
            .Interior.Color = RGB(220, 230, 241)
        End With
        ws.Rows(r).RowHeight = 20
    Case "B"
        ws.Cells(r, 2).Value = bText
        With ws.Range(ws.Cells(r, 2), ws.Cells(r, 4))
            .Merge
            .WrapText = True
            .VerticalAlignment = xlVAlignTop
        End With
        ws.Rows(r).RowHeight = 15 * (Int(Len(bText) / 150) + 1)
    Case "K"
        ws.Cells(r, 2).Value = bText
        ws.Cells(r, 3).Value = cText
        ws.Cells(r, 4).Value = dText
        With ws.Range(ws.Cells(r, 2), ws.Cells(r, 4))
            .Font.Bold = True
            .Interior.Color = RGB(242, 242, 242)
            .WrapText = True
        End With
        FitWrappedRow ws, r, 15
    Case "T"
        ws.Cells(r, 2).Value = bText
        ws.Cells(r, 3).Value = cText
        ws.Cells(r, 4).Value = dText
        With ws.Range(ws.Cells(r, 2), ws.Cells(r, 4))
            .WrapText = True
            .VerticalAlignment = xlVAlignTop
        End With
        FitWrappedRow ws, r, 15
    End Select
    OnbRow = r + 1
End Function
' =============================================================================
' HUB DATA CACHE -- the last pull, parked on a very-hidden sheet so Refresh
' Tabs can rebuild everything without going back to Hub. Very-hidden rather
' than merely hidden so nobody can unhide and edit it by accident; it is
' overwritten wholesale on every successful Update Data.
' =============================================================================
Private Sub SaveHubCache(ByRef rawData As Variant)
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = GetOrCreateSheet(CACHE_SHEET)
    If ws Is Nothing Then Exit Sub
    ws.Visible = xlSheetVisible
    ws.Cells.Clear
    Dim nRows As Long, nCols As Long
    nRows = UBound(rawData, 1)
    nCols = UBound(rawData, 2)
    ' Job ID and RFID Type are text and must stay text. Without this Excel
    ' coerces anything that merely looks numeric on the way in, and the
    ' Refresh button -- which rebuilds from this sheet rather than from a
    ' fresh Hub pull -- would then produce different tabs than Update Data
    ' did. Same class of bug as WC_Work-Center-ID '687.10' in Hub v8.
    SetFmt ws.Columns(COL_JOBID), "@", "HubCache Job ID"
    If nCols >= COL_RFID Then SetFmt ws.Columns(COL_RFID), "@", "HubCache RFID"

    ws.Range("A1").Resize(nRows, nCols).Value = rawData
    ws.Visible = xlSheetVeryHidden
    On Error GoTo 0
End Sub

' Returns the cached dataset in the same shape PullCleanedDataFromHub
' hands back, or Empty when nothing has been cached yet.
Private Function LoadHubCache() As Variant
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(CACHE_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    Dim lastR As Long, lastC As Long
    lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastC = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If lastR < 2 Or lastC < 9 Then Exit Function
    LoadHubCache = ws.Range(ws.Cells(1, 1), ws.Cells(lastR, lastC)).Value
End Function
' =============================================================================
' REFRESH TABS -- second entry point, wired to the Refresh button beside
' Update Data. Rebuilds every machine tab from the CACHED Hub dataset using
' whatever is currently on Preferences: the Row Formatting swatches AND the
' Scheduling Settings (Pull-Forward Days, Extra Days If Under-Filled, Max
' Under-Capacity %), plus the Machine Roster thresholds and Late Rules.
'
' It runs the SAME BuildAllMachineTabs path a real Update Data runs -- the
' only difference is where the data comes from. That is deliberate: a
' separate 'refresh-only' rendering path would be a second implementation
' of the same rules, free to drift from the real one. Here, whatever Update
' Data would produce from this dataset is exactly what Refresh produces.
'
' Because it re-schedules, Production Days CAN move -- that is the point when
' the window or fill settings change. What it never does is contact Hub, so
' the job list itself is frozen until the next Update Data.
' =============================================================================
Public Sub RefreshPressTabs()
    Dim t0 As Single: t0 = Timer
    gStep = "starting up"
    gCheckboxTried = False
    gCheckboxOK = True
    gFmtWarn = ""
    EnsurePreferencesSheet
    EnsureUpdateButton
    EnsureRefreshButton

    Dim roster As Variant
    roster = ReadRosterTable()
    If IsEmpty(roster) Then
        MsgBox "The Machine Roster table on Preferences has no rows, so there are no tabs to rebuild.", vbExclamation, "Refresh"
        ShowPreferencesWhenDone
        Exit Sub
    End If

    Dim cached As Variant
    cached = LoadHubCache()
    If IsEmpty(cached) Then
        MsgBox "There is no cached Hub data in this workbook yet, so there is nothing to rebuild from." & vbCrLf & vbCrLf & _
               "Click Update Data once. After that, Refresh can rebuild the tabs with new settings or colours without going back to Hub.", _
               vbExclamation, "Refresh"
        ShowPreferencesWhenDone
        Exit Sub
    End If

    On Error GoTo RefreshFailed
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    ' Harvest the ticks BEFORE any tab is cleared -- the marks live on the
    ' tabs, which this rebuild is about to wipe.
    ' Order matters: the standing RFID filters are rewritten from the tabs
    ' FIRST, so the harvest below can see which technologies are switched off
    ' and avoid recording their jobs as permanent one-off removals.
    Dim rfidFilters As Long
    rfidFilters = SaveRemovedRfidTypes(roster)

    Dim newlyRemoved As Long
    newlyRemoved = HarvestTickedRemovals(roster)

    BuildAllMachineTabs cached, roster
    EnforcePressWorksheetOrder roster

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True

    BuildRulesSheet
    BuildOnboardingSheet
    ShowPreferencesWhenDone

    Dim totalRemoved As Long
    totalRemoved = LoadRemovedJobs().Count

    MsgBox "Tabs rebuilt in " & Format(Timer - t0, "0.00") & " seconds -- no Hub pull." & vbCrLf & vbCrLf & _
           "Machines rebuilt: " & UBound(roster, 1) & vbCrLf & _
           "Jobs removed this time: " & newlyRemoved & vbCrLf & _
           "Removed in total: " & totalRemoved & " (see the Removed Jobs tab)" & vbCrLf & _
           "RFID types switched off: " & rfidFilters & vbCrLf & vbCrLf & _
           "Row colours, fonts and the machine roster were re-read from Preferences. " & _
           "Click Update Data when you want fresh jobs from Hub -- that also clears every removal." & vbCrLf & vbCrLf & _
           "Press module " & MODULE_VERSION & "." & _
           CheckboxWarning() & FormatWarning(), vbInformation
    Exit Sub

RefreshFailed:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Dim rErrNum As Long, rErrMsg As String
    rErrNum = Err.Number: rErrMsg = Err.Description
    MsgBox "Refresh failed while " & gStep & "." & vbCrLf & vbCrLf & _
           "Excel reported: " & rErrNum & " - " & rErrMsg, vbCritical, "Refresh"
End Sub

' Creates the Refresh button once, parked immediately to the right of the
' Update Data button wherever the user has dragged that. Same user-owned
' rule as every other button here: once it exists only its macro wiring is
' refreshed. The one exception is the caption -- if it still reads the old
' default 'Refresh Colours' it is corrected to 'Refresh Tabs', because the
' button now re-schedules as well as recolours and the old wording would
' understate what it does. A caption the user has changed is left alone.
Private Sub EnsureRefreshButton()
    Const BTN_NAME As String = "btnRefreshPressColours"

    Dim ws As Worksheet, shp As Shape
    For Each ws In ThisWorkbook.Worksheets
        Set shp = Nothing
        On Error Resume Next
        Set shp = ws.Shapes(BTN_NAME)
        On Error GoTo 0
        If Not shp Is Nothing Then
            shp.OnAction = "RefreshPressTabs"
            On Error Resume Next
            If shp.TextFrame2.TextRange.Text = "Refresh Colours" Then
                shp.TextFrame2.TextRange.Text = "Refresh Tabs"
            End If
            On Error GoTo 0
            Exit Sub
        End If
    Next ws

    Dim wsPrefs As Worksheet
    On Error Resume Next
    Set wsPrefs = ThisWorkbook.Sheets("Preferences")
    On Error GoTo 0
    If wsPrefs Is Nothing Then Exit Sub

    Dim posLeft As Double, posTop As Double
    posLeft = wsPrefs.Cells(4, 5).Left
    posTop = wsPrefs.Cells(4, 5).Top
    Dim upd As Shape
    On Error Resume Next
    Set upd = wsPrefs.Shapes("btnUpdatePressData")
    On Error GoTo 0
    If Not upd Is Nothing Then
        posLeft = upd.Left + upd.Width + 12
        posTop = upd.Top
    End If

    Set shp = wsPrefs.Shapes.AddShape(msoShapeRoundedRectangle, posLeft, posTop, 150, 36)
    shp.name = BTN_NAME
    shp.OnAction = "RefreshPressTabs"
    With shp
        .Fill.ForeColor.RGB = RGB(84, 130, 53)
        With .Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(0, 0, 0)
            .Weight = 1.5
        End With
        With .TextFrame2.TextRange
            .Text = "Refresh Tabs"
            .Font.Size = 12
            .Font.Bold = msoTrue
            .Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
            .ParagraphFormat.Alignment = msoAlignCenter
        End With
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
    End With
End Sub
' =============================================================================
' RULES TAB -- rebuilt from scratch on every Update Data run so it always
' matches the code. Documents every scheduling rule and shows a LIVE
' formatting legend: the example rows are painted from the same Row
' Formatting swatches the machine tabs use, so restyling a swatch on
' Preferences updates this legend too.
' =============================================================================
Private Sub BuildRulesSheet()
    Dim ws As Worksheet
    Set ws = GetOrCreateSheet("Rules")
    If ws Is Nothing Then Exit Sub
    ws.Cells.Clear
    ' Same reason as BuildOneMachineTab -- .Clear leaves merges behind and this
    ' sheet re-merges A1:C1 on every run.
    ws.Cells.UnMerge
    ws.Rows.RowHeight = ws.StandardHeight

    gStep = "rebuilding the Rules sheet"
    ws.Cells(1, 1).Value = "PRESS SCHEDULING RULES"
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, 3))
        .Merge
        .Font.Bold = True
        .Font.Size = 14
        .Interior.Color = RGB(31, 78, 120)
        .Font.Color = RGB(255, 255, 255)
    End With
    ws.Cells(2, 1).Value = "Rebuilt automatically on every Update Data run -- edits here are overwritten. " & _
        "Rules current as of 2026-09-03, Press module " & MODULE_VERSION & "."
    ws.Cells(2, 1).Font.Italic = True

    Dim r As Long: r = 4
    r = RuleRow(ws, r, "1", "Fill each day sequentially by Ship Date", "The queue is sorted by Ship Date and each production day is filled as full as it will go before moving to the next. There is no pull-forward window, no extra reach and no minimum fill -- those were removed on 2026-08.17. Work runs as early as capacity allows, which means a job can run well ahead of its Ship Date if the presses have room.")
    r = RuleRow(ws, r, "2", "Job ID does not matter", "Hyphen families, sub-jobs and matching base numbers are ignored entirely. Every line is scheduled on its own. Two pieces of the same job number may run on different days.")
    r = RuleRow(ws, r, "3", "Ship Date order is absolute; gaps fill from the SAME ship date only", "Nothing ever runs ahead of a job that ships earlier. When the next job does not fit the space left on a day, it is NOT cut and it is NOT skipped -- it leads the next day. The gap it leaves is filled by any job SHARING ITS EXACT SHIP DATE that does fit, which cannot disturb the order because the dates are equal. Jobs shipping later are never pulled forward. Measured on the live 665-job Dicetrax queue 2026-08-24: this runs 46 days at 94.1% average fill with ZERO segments out of Ship Date order. Filling gaps from any ship date would reach 44 days / 98.4% but would put 34 segments ahead of earlier-shipping work, so it is not done. Filling nothing at all was tried first and ran 50 days at 86.6%, wasting 261 separate chances to place a same-date job that fitted -- which is why this rule exists.")
    r = RuleRow(ws, r, "4", "Overdue work runs today and COUNTS against capacity", "A job whose Ship Date has already passed has the earliest date in the queue, so rule 3 reaches it first and puts it on the CURRENT production day. It is charged to that day like any other work, so today schedules a full threshold in total, not a threshold on top of the overdue block. Its row carries the late colour. If there is more overdue work than one day holds, the excess splits or rolls forward exactly as rule 3 says -- and no day, including day one, can read above its threshold.")
    r = RuleRow(ws, r, "5", "Hard capacity cap", "No day's total ever exceeds the Daily Threshold set in the Machine Roster -- there are no exceptions to this since 2026-08-18, when overdue work started counting against capacity too. Work that will not fit is cut or moves to a later day, and is flagged late if that pushes it past its Ship Date.")
    r = RuleRow(ws, r, "6", "Late means AFTER the Ship Date", "A job that runs ON its Ship Date is ON TIME -- it gets the ships-today colour, not the late one. Indigo is the exception and uses the stricter reading where running on the ship date already counts as late. Set per machine in the Machine Roster's Late Rule column.")
    r = RuleRow(ws, r, "7", "ONLY a job bigger than a whole day is ever split", "A job is cut if, and only if, its TOTAL quantity is larger than the machine's Daily Threshold -- it cannot physically run in one day. Every other job runs whole, on one day, however awkwardly it lands in the queue. A job that qualifies starts on its OWN FRESH DAY rather than in the tail of a part-used day, then takes whole days until what is left of it is smaller than a day; that last partial day is ordinary space again and the following jobs fill it in Ship Date order. A split row's Qty column reads only what that slice runs, not the whole order -- add the slices to get the order. The remainder carried in from the day before is listed FIRST on its day.")
    r = RuleRow(ws, r, "8", "Removing a job from the schedule", "Tick the REMOVE box in column L next to any job, then click Refresh Data. That job comes off THIS machine only -- the same job number on another machine is untouched -- and the capacity it would have used is handed straight back, so the rest of the day packs into the gap. Removals build up across repeated Refresh presses and every one is listed on the Removed Jobs tab with its quantity, ship date and the time it was pulled. To put a single job back, delete its row there and click Refresh Data. UPDATE DATA clears the whole list and brings everything back -- so each morning starts from a clean slate and nothing stays hidden without somebody deciding again today. Ticking any slice of a split job removes the whole job; you cannot run half a job you have decided to pull.")
    ' Split across continuations because VBA rejects any PHYSICAL line over
    ' 1023 characters. The single-line version of this row was 1,062 and
    ' failed in Excel with a bare "Compile error" (2026-08.25) -- LibreOffice
    ' does not enforce the limit, so nothing upstream caught it. The joined
    ' statement is ~1.1k, well inside the ~2.0k that is known to compile.
    r = RuleRow(ws, r, "8b", "Switching a whole RFID technology off and on", _
        "Column M on each machine tab lists every RFID technology on that tab; column N puts a " & _
        "tick box beside each. Ticking one and clicking Refresh Data takes EVERY job of that " & _
        "technology off this machine. Unlike a column L removal this is not permanent -- it is a " & _
        "switch that STAYS ON. The tick survives every later Refresh Data, and UNTICKING it and " & _
        "clicking Refresh Data brings all of those jobs straight back. Switched-off types stay in " & _
        "the list in bold italic even after their last job has gone, which is the only way you can " & _
        "still reach the tick box to undo them; they are also listed on the Removed Jobs tab in the " & _
        "green Machine / RFID Type block on the right. Because it is reversible, a technology " & _
        "switched off this way is NOT written into the permanent removed-jobs list on the left. " & _
        "UPDATE DATA clears the switches along with everything else, so every morning starts with " & _
        "the whole queue visible. Blank RFID types are not listed -- remove those jobs individually " & _
        "from column L.")
    r = RuleRow(ws, r, "9", "Count-rule machines", "Metronics and Indigo schedule by JOBS per day (Daily Threshold = a job count) in Ship-Date order -- no quantity maths and no splitting. Their rows still take the late and ships-today colours.")
    r = RuleRow(ws, r, "*", "What is adjustable", "Daily Threshold and Late Rule per machine live in the Machine Roster; which jobs qualify lives in the Location Filter; every colour, font and weight in the legend below lives in the Row Formatting table. All on the Preferences sheet. The old Scheduling Settings table (Pull-Forward Days, Extra Days, Max Under-Capacity %) was deleted on 2026-08.17 because none of it applies any more.")

    r = r + 1
    ws.Cells(r, 1).Value = "FORMATTING LEGEND"
    ws.Cells(r, 3).Value = "These sample rows are painted with the very same styles the machine tabs use. To change any of them, restyle its swatch cell in the Row Formatting table on the Preferences sheet -- colour, font, size and weight all follow."
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, 3))
        .Font.Bold = True
        .Interior.Color = RGB(220, 230, 241)
    End With
    ws.Cells(r, 3).WrapText = True
    FitWrappedRow ws, r, 30
    r = r + 1

    Dim rowFmts As Variant: rowFmts = ReadRowFormats()
    r = LegendRow(ws, r, FMT_NORMAL, "Normal job", "Runs before its Ship Date, whole, inside capacity.", rowFmts)
    r = LegendRow(ws, r, FMT_SHIPDAY, "Ships today", "Runs exactly ON its Ship Date -- still on time, but with no slack left.", rowFmts)
    r = LegendRow(ws, r, FMT_LATE, "Late job", "Runs AFTER its Ship Date: already overdue when pulled, or pushed past the date by the capacity cap.", rowFmts)
    r = LegendRow(ws, r, FMT_SPLIT, "Split job", "One day's slice of a job cut by the capacity line. The Qty column shows only what runs that day.", rowFmts)
    r = LegendRow(ws, r, FMT_SPLIT_LATE, "Split job, also late", "A split slice that also lands after the job's Ship Date.", rowFmts)

    ws.Cells(r, 2).Value = "Day separator"
    ws.Cells(r, 3).Value = "Black bar between one Production Day's rows and the next."
    ws.Range(ws.Cells(r, 2), ws.Cells(r, 3)).Interior.Color = RGB(0, 0, 0)
    ws.Range(ws.Cells(r, 2), ws.Cells(r, 3)).Font.Color = RGB(255, 255, 255)
    r = r + 1
    ws.Cells(r, 2).Value = "No Ship Date"
    ws.Cells(r, 3).Value = "Undated jobs list at the bottom of each tab, italic, unscheduled -- no Production Day and no Total Qty."
    ws.Range(ws.Cells(r, 2), ws.Cells(r, 3)).Font.Italic = True

    ws.Columns(1).ColumnWidth = 4
    ws.Columns(2).ColumnWidth = 30
    ws.Columns(3).ColumnWidth = 112
End Sub

' Writes one rule row and returns the next free row. Emitted as one call
' per rule rather than from a continued Array literal: VBA rejected the
' array form with a plain 'Syntax error' once the joined statement grew
' past roughly 2.7k characters (the 2.0k version in the 2026-08.12 build
' compiled fine), and LibreOffice's compile check does not reproduce
' either failure. Same lesson as the Onboarding builder -- keep long
' content as one statement per row.
Private Function RuleRow(ByRef ws As Worksheet, ByVal r As Long, ByVal num As String, _
        ByVal title As String, ByVal detail As String) As Long
    ws.Cells(r, 1).Value = num
    ws.Cells(r, 2).Value = title
    ws.Cells(r, 3).Value = detail
    ws.Cells(r, 1).Font.Bold = True
    ws.Cells(r, 2).Font.Bold = True
    ws.Cells(r, 3).WrapText = True
    FitWrappedRow ws, r, 30
    RuleRow = r + 1
End Function

' Writes one legend row, painted with the user's own swatch for that state.
Private Function LegendRow(ByRef ws As Worksheet, ByVal r As Long, ByVal stateIdx As Long, _
        ByVal label As String, ByVal detail As String, ByRef rowFmts As Variant) As Long
    ws.Cells(r, 2).Value = label
    ws.Cells(r, 3).Value = detail
    ws.Cells(r, 3).WrapText = True
    ApplyRowFormat ws, r, 3, rowFmts, stateIdx
    FitWrappedRow ws, r, 30
    LegendRow = r + 1
End Function


' =============================================================================
' PREFERENCES SHEET -- control panel. Same label-anchored, self-healing
' pattern as the Hub workbook: every row is found by searching column A for
' exact text, never by fixed row number, so reorganizing this sheet by hand
' (confirmed OK with the user) never breaks the code that reads it.
' =============================================================================
Private Sub EnsurePreferencesSheet()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Preferences")
    On Error GoTo 0

    Dim isFirstBuild As Boolean

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(1))
        ws.name = "Preferences"
    End If

    ' Self-healing wipe: the build template this workbook descends from may
    ' carry inherited cell content (e.g. leftover rate-table data) that has
    ' nothing to do with Press. Safe VBA .Cells.Clear -- not XML surgery --
    ' same pattern as MigratePreferencesToHubLayout in the Hub workbook.
    ' Only fires once, guarded by the title marker, so later runs never
    ' wipe the user's own Hub Workbook Path / roster edits.
    If Trim(CStr(ws.Cells(1, 1).Value)) <> "Press (Hub Satellite)" Then
        isFirstBuild = True
        ws.Cells.Clear
        ws.Cells(1, 1).Value = "Press (Hub Satellite)"
        ws.Cells(1, 1).Font.Bold = True
        ws.Cells(1, 1).Font.Size = 14
        ws.Cells(2, 1).Value = "Point 'Hub Workbook Path' at PLI Hub.xlsm, then click Update Data below."
        ws.Cells(2, 1).Font.Italic = True
    End If

    EnsureOnePressPrefRow ws, "Hub Workbook Path", "path", "", 0, 0
    EnsureOnePressPrefRow ws, "Press Status", "status", "Not Ready - not yet run", RGB(226, 239, 218), RGB(0, 97, 0)
    ' Seeded only where it does not already exist -- the live workbook has
    ' carried this row (unread) since before 2026-08-24.
    EnsureOnePressPrefRow ws, "Production Cutoff", "time", "14:00", 0, 0
    gStep = "Preferences -- Machine Roster"
    EnsureMachineRosterTable ws
    MigrateRosterLateRule ws
    EnsureRosterValidation ws
    EnsureLocationFilterList ws
    MigrateRemoveSchedulingSettings ws
    EnsureRowFormatTable ws

    ' AutoFit was gated to first-build-only 2026-08 (see fix note above --
    ' running it unconditionally used to blow away the user's manually-set
    ' column widths on every Update Data click). Preferences is now fully
    ' populated and the user has set their own column widths by hand, so
    ' this is disabled entirely (not just gated) for now at their request --
    ' uncomment the line below to re-enable first-build AutoFit for a future
    ' from-scratch rebuild of this workbook.
    ' If isFirstBuild Then ws.Columns.AutoFit
End Sub

Private Sub EnsureOnePressPrefRow(ByRef ws As Worksheet, ByVal lbl As String, ByVal kind As String, _
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
            Case "status"
                .Value = defaultText
            Case "time"
                .Value = defaultText
                .Interior.Color = fillC
                .Font.Color = fontC
                .HorizontalAlignment = xlHAlignLeft
        End Select
        .BorderAround Weight:=xlThin
    End With
    ' Text format for both kinds, applied after the value so a path like
    ' 687.10 or a status word is never coerced to a number. Routed through
    ' SetFmt with everything else -- see the note on that procedure.
    Select Case kind
        Case "path", "status"
            SetFmt ws.Cells(lastA, 2), "@", "Preferences '" & lbl & "'"
        Case "time"
            SetFmt ws.Cells(lastA, 2), "hh:mm", "Preferences '" & lbl & "'"
    End Select

    If lbl = "Hub Workbook Path" Then
        ws.Cells(lastA + 1, 1).Value = "Full path to PLI Hub.xlsm on the shared drive (e.g. \\server\share\PLI Hub.xlsm)."
        ws.Cells(lastA + 1, 1).Font.Size = 9
        ws.Cells(lastA + 1, 1).Font.Italic = True
    ElseIf lbl = "Press Status" Then
        ws.Cells(lastA + 1, 1).Value = "Ready = last pull succeeded (detail shows the Monarch export's own timestamp). " & _
            "Not Ready = pull in progress or waiting on Hub. Error = pull failed -- see detail after the dash."
        ws.Cells(lastA + 1, 1).Font.Size = 9
        ws.Cells(lastA + 1, 1).Font.Italic = True
    End If
End Sub

' Machine Roster table: label anchor row "Machine Roster", an instructions
' row, a header row, then data rows (Work Center Match | Rule Type |
' Daily Threshold). Created once with the 3 known Press machines pre-seeded;
' never rewritten after that (only appended-to by the user) -- same
' non-destructive pattern as the rest of this project's self-healing
' Ensure* routines. See ReadRosterTable for the matching row-offset logic.
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
    ws.Cells(lastA + 1, 1).Value = "Add a row below to add a new tab -- no VBA change needed. " & _
        "Work Center Match = text to look for anywhere in a job's Work Center (case/spacing don't matter). " & _
        "Rule Type = COUNT (N jobs/day) or QTY (jobs summed up to a daily total, using Effective QTY). " & _
        "Late Rule = AFTER (a job is late only after its Ship Date; running on the ship date is on time and " & _
        "shows the ships-today colour) or ON/AFTER (running on its own ship date already counts as late). " & _
        "Removing a row here does NOT delete its tab -- remove/hide that manually if wanted."
    ws.Cells(lastA + 1, 1).Font.Size = 9
    ws.Cells(lastA + 1, 1).Font.Italic = True

    Dim headerRow As Long: headerRow = lastA + 2
    ws.Cells(headerRow, 1).Value = "Work Center Match"
    ws.Cells(headerRow, 2).Value = "Rule Type"
    ws.Cells(headerRow, 3).Value = "Daily Threshold"
    ws.Cells(headerRow, 4).Value = "Late Rule"
    With ws.Range(ws.Cells(headerRow, 1), ws.Cells(headerRow, 4))
        .Font.Bold = True
        .Interior.Color = RGB(220, 230, 241)
    End With

    Dim seedRow As Long: seedRow = headerRow + 1
    ws.Cells(seedRow, 1).Value = "Metronics": ws.Cells(seedRow, 2).Value = "Count": ws.Cells(seedRow, 3).Value = 4
    ws.Cells(seedRow, 4).Value = "After"
    ws.Cells(seedRow + 1, 1).Value = "Dicetrax": ws.Cells(seedRow + 1, 2).Value = "Qty": ws.Cells(seedRow + 1, 3).Value = 35000
    ws.Cells(seedRow + 1, 4).Value = "After"
    ws.Cells(seedRow + 2, 1).Value = "Indigo": ws.Cells(seedRow + 2, 2).Value = "Count": ws.Cells(seedRow + 2, 3).Value = 4
    ws.Cells(seedRow + 2, 4).Value = "On/After"

    SetFmt ws.Range(ws.Cells(seedRow, 3), ws.Cells(seedRow + 2, 3)), "#,##0", "Machine Roster thresholds"
    ws.Columns(1).ColumnWidth = 28
End Sub

' One-time, idempotent migration (2026-08.13): older Preferences sheets
' have a 3-column Machine Roster. Adds the Late Rule column and seeds it
' from the user's ruling -- Indigo keeps the legacy ON/AFTER behaviour,
' every other machine moves to AFTER. Runs on every Update Data but does
' nothing once the header exists, so a user's own edits are never
' overwritten.
' =============================================================================
' 2026-08.17 MIGRATION -- delete the retired Scheduling Settings block.
' Pull-Forward Days, Extra Days If Under-Filled and Max Under-Capacity %
' no longer exist: v32 fills every day sequentially by ship date to the
' threshold, with no window and no fill floor. The owner's call was to
' remove the block outright rather than leave dead knobs on the sheet --
' a setting that does nothing is worse than no setting, because someone
' will change it and expect a result.
'
' Finds the block by its column-A label and deletes from that row through
' the last row of its table, so it works no matter where the block sits.
' Idempotent: a no-op forever once the label is gone.
' =============================================================================
Private Sub MigrateRemoveSchedulingSettings(ByRef ws As Worksheet)
    Dim anchorRow As Long: anchorRow = 0
    Dim r As Long
    For r = 1 To 200
        If Trim(CStr(ws.Cells(r, 1).Value)) = "Scheduling Settings" Then
            anchorRow = r
            Exit For
        End If
    Next r
    If anchorRow = 0 Then Exit Sub

    ' anchor, the italic explanation, the Setting/Value/What it does header,
    ' then one row per setting until column A goes blank.
    Dim lastRow As Long: lastRow = anchorRow + 2
    Dim d As Long: d = anchorRow + 3
    Do While Trim(CStr(ws.Cells(d, 1).Value)) <> "" And d < anchorRow + 60
        lastRow = d
        d = d + 1
    Loop

    ' Take the trailing blank separator row too, so removing the block does
    ' not leave a double gap between the sections either side of it.
    If Trim(CStr(ws.Cells(lastRow + 1, 1).Value)) = "" Then lastRow = lastRow + 1

    ws.Rows(anchorRow & ":" & lastRow).Delete Shift:=xlUp
End Sub

' =============================================================================
' Roster dropdowns, rebuilt in code every run (2026-08-24).
'
' They used to exist only as hand-added data validation living in one copy of
' the workbook, so every fresh build silently dropped them. Now they are part
' of the self-healing Ensure* set like everything else on Preferences, and it
' no longer matters which file the VBA is spliced into.
'
' The same pass retires Fill-Ahead (Days). Pull-forward scheduling was deleted
' in v32 on 2026-08.17 and ReadRosterTable has read only columns 1-4 ever
' since, so the column has been inert for a week -- a live-looking control on a
' control panel that changes nothing, which is worse than no control at all.
' The column itself is left in place because the section banners above it are
' merged across A:E; only its content, styling and validation go.
' =============================================================================
Private Sub EnsureRosterValidation(ByRef ws As Worksheet)
    Dim anchorRow As Long: anchorRow = 0
    Dim r As Long
    For r = 1 To 100
        If Trim(CStr(ws.Cells(r, 1).Value)) = "Machine Roster" Then
            anchorRow = r
            Exit For
        End If
    Next r
    If anchorRow = 0 Then Exit Sub

    Dim headerRow As Long: headerRow = anchorRow + 2
    Dim firstRow As Long: firstRow = headerRow + 1

    Dim lastRow As Long: lastRow = firstRow - 1
    For r = firstRow To firstRow + 199
        If Trim(CStr(ws.Cells(r, 1).Value)) = "" Then Exit For
        If IsPrefSectionLabel(Trim(CStr(ws.Cells(r, 1).Value))) Then Exit For
        lastRow = r
    Next r
    If lastRow < firstRow Then Exit Sub

    ' Spare rows for a machine added by hand later -- but the roster block ends
    ' where the NEXT section begins, and on the live sheet that is only one row
    ' below the last machine. 2026-08-24: the first cut of this used a flat
    ' lastRow + 20 and drove E12:E35 straight through nine merged banners
    ' (A17:E17, A18:E18, A27:E27, C28:E28 .. C33:E33), so ClearContents hit
    ' part of a merged area and threw the very 1004 this session existed to
    ' remove. Bound the block properly instead of assuming empty space below it.
    Dim blockEnd As Long: blockEnd = lastRow + 20
    For r = lastRow + 1 To lastRow + 21
        If IsPrefSectionLabel(Trim(CStr(ws.Cells(r, 1).Value))) Then
            blockEnd = r - 1
            Exit For
        End If
    Next r
    If blockEnd < lastRow Then blockEnd = lastRow

    SetListValidation ws.Range(ws.Cells(firstRow, 2), ws.Cells(blockEnd, 2)), "Count,Qty"
    SetListValidation ws.Range(ws.Cells(firstRow, 4), ws.Cells(blockEnd, 4)), "After,On/After"

    ' Fill-Ahead (Days): validation off, content and styling cleared. Guarded --
    ' a Preferences sheet whose banners were laid out differently could still
    ' put a merge in the way, and losing a banner to a stray ClearContents
    ' would be a far worse outcome than leaving a dead column in place.
    Dim eRng As Range
    Set eRng = ws.Range(ws.Cells(headerRow, 5), ws.Cells(blockEnd, 5))
    If RangeIsUnmerged(eRng) Then
        On Error Resume Next
        eRng.Validation.Delete
        On Error GoTo 0
        eRng.ClearContents
        eRng.Interior.Pattern = xlNone
        eRng.Borders.LineStyle = xlNone
    End If
End Sub

' True only when NO cell in the range belongs to a merged area.
' Range.MergeCells returns Null for a mixed range, so a plain "= True" test is
' not enough -- Null is exactly the case that matters, and it is also the case
' that silently breaks a naive comparison.
Private Function RangeIsUnmerged(ByRef rng As Range) As Boolean
    RangeIsUnmerged = False
    If rng Is Nothing Then Exit Function
    Dim v As Variant
    On Error Resume Next
    Err.Clear
    v = rng.MergeCells
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0
    If IsNull(v) Then Exit Function
    RangeIsUnmerged = (v = False)
End Function

' One list dropdown. Delete-then-Add because Validation.Add throws 1004 on a
' range that already carries a rule, and this runs on every single build.
Private Sub SetListValidation(ByRef rng As Range, ByVal csvList As String)
    ' Never hand a dropdown to a merged banner.
    If Not RangeIsUnmerged(rng) Then Exit Sub
    On Error Resume Next
    rng.Validation.Delete
    Err.Clear
    rng.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                       Operator:=xlBetween, Formula1:=csvList
    If Err.Number = 0 Then
        rng.Validation.IgnoreBlank = True
        rng.Validation.InCellDropdown = True
        rng.Validation.ShowError = True
    End If
    Err.Clear
    On Error GoTo 0
End Sub

Private Sub MigrateRosterLateRule(ByRef ws As Worksheet)
    Dim anchorRow As Long: anchorRow = 0
    Dim r As Long
    For r = 1 To 100
        If Trim(CStr(ws.Cells(r, 1).Value)) = "Machine Roster" Then
            anchorRow = r
            Exit For
        End If
    Next r
    If anchorRow = 0 Then Exit Sub

    Dim headerRow As Long: headerRow = anchorRow + 2
    If Trim(CStr(ws.Cells(headerRow, 4).Value)) <> "" Then Exit Sub   ' already migrated

    ws.Cells(headerRow, 4).Value = "Late Rule"
    ws.Cells(headerRow, 4).Font.Bold = True
    ws.Cells(headerRow, 4).Interior.Color = RGB(220, 230, 241)

    Dim d As Long: d = headerRow + 1
    Do While Trim(CStr(ws.Cells(d, 1).Value)) <> "" And d < headerRow + 200
        If InStr(1, CStr(ws.Cells(d, 1).Value), "Indigo", vbTextCompare) > 0 Then
            ws.Cells(d, 4).Value = "On/After"
        Else
            ws.Cells(d, 4).Value = "After"
        End If
        d = d + 1
    Loop
End Sub

' Location Filter list: label anchor row "Location Filter", an instructions
' row, then one keyword per row in column A starting 2 rows below the anchor
' (same shape as Machine Roster's anchor, minus the header/multi-column
' part since this is a single-column list). Created once, seeded with the
' phrases confirmed with the user 2026-08; never rewritten after that (only
' appended-to/edited by the user) -- same non-destructive Ensure* pattern as
' the rest of this project. "   " (bare spaces) from the user's original
' request was deliberately NOT seeded here: as a case-insensitive "contains"
' keyword it would match every row with any Location text at all (an
' include-everything bug), so it was omitted and flagged back to the user
' rather than silently guessed at. See ReadLocationFilterList for the
' matching row-offset logic this must stay in sync with.
Private Sub EnsureLocationFilterList(ByRef ws As Worksheet)
    Dim r As Long
    For r = 1 To 100
        If Trim(CStr(ws.Cells(r, 1).Value)) = "Location Filter" Then Exit Sub ' already exists
    Next r

    Dim lastA As Long
    lastA = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 2

    With ws.Cells(lastA, 1)
        .Value = "Location Filter"
        .Font.Bold = True
    End With
    ws.Cells(lastA + 1, 1).Value = "Add/remove rows below -- no VBA change needed. A job only appears on a " & _
        "machine tab if its Location CONTAINS at least one of these phrases (case/spacing don't matter), " & _
        "on top of the Work Center match above. Delete all rows here to disable this filter entirely."
    ws.Cells(lastA + 1, 1).Font.Size = 9
    ws.Cells(lastA + 1, 1).Font.Italic = True

    Dim seedRow As Long: seedRow = lastA + 2
    Dim seeds As Variant
    seeds = Array("Plates to Press", "Final Proof Returned", "Final Proof to Sales", "Final Proof", _
        "TI", "To Imposition", "Ready to Sch", "To PP Final Proof")

    Dim i As Long
    For i = LBound(seeds) To UBound(seeds)
        ws.Cells(seedRow + i, 1).Value = seeds(i)
    Next i

    ws.Columns(1).ColumnWidth = 28
End Sub

' =============================================================================
' UTILITY -- label-anchored value lookups (shared shape with Hub's
' GetDashSwatch, duplicated here since these are two separate workbooks).
' =============================================================================
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

Private Function GetLabeledRangeOnSheet(ByRef ws As Worksheet, ByVal label As String) As Range
    Dim r As Long
    For r = 1 To 100
        If Trim(CStr(ws.Cells(r, 1).Value)) = label Then
            Set GetLabeledRangeOnSheet = ws.Cells(r, 2)
            Exit Function
        End If
    Next r
End Function

Private Function GetPrefValue(ByVal label As String) As Variant
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Preferences")
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    GetPrefValue = GetLabeledValueFromSheet(ws, label)
End Function

' statusWord is always one of "Ready" / "Not Ready" / "Error" (see call sites
' in UpdatePressData and PullCleanedDataFromHub) -- confirmed with the user
' 2026-08 as the fixed 3-state vocabulary for this cell, replacing the old
' PULLING/READY/WAITING/ERROR wording. Separator changed from " | " to " - "
' to match "Ready - Updated At: <export timestamp>" at the same time.
Private Sub SetPressStatus(ByVal statusWord As String, ByVal detail As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Preferences")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    Dim sw As Range
    Set sw = GetLabeledRangeOnSheet(ws, "Press Status")
    If sw Is Nothing Then Exit Sub
    sw.Value = statusWord & IIf(detail <> "", " - " & detail, "")

    ' Error runs get Excel's built-in "Bad" cell style (red) so a failure is
    ' unmistakable at a glance -- confirmed with the user 2026-08. Any other
    ' state explicitly reverts to "Normal" and reapplies this cell's original
    ' light-green look, so a later successful run doesn't leave a stale red
    ' cell behind from a prior error. On Error Resume Next here is purely
    ' defensive -- if a workbook's built-in "Bad"/"Normal" styles were ever
    ' renamed or deleted, styling is cosmetic and should never block the
    ' actual status text (already written above) from updating.
    On Error Resume Next
    If statusWord = "Error" Then
        sw.Style = "Bad"
    Else
        sw.Style = "Normal"
        sw.Interior.Color = RGB(226, 239, 218)
        sw.Font.Color = RGB(0, 97, 0)
    End If
    On Error GoTo 0
End Sub

' Hub's own Hub Status detail embeds the Monarch export file's actual
' last-modified timestamp, e.g. "ReportExport_2026-08-11.xlsx (modified
' 2026-08-11 08:15) | updated 2026-08-11 08:16" (see srcDesc in Hub's
' PullFromMonarchFolder). Pulls just the "yyyy-mm-dd hh:nn" between
' "(modified " and ")" out of that string so Press can show the export's
' own timestamp instead of the moment Press happened to run. Returns ""
' if the marker isn't found (e.g. an older Hub build before this existed),
' which the caller falls back on rather than showing a made-up value.
Private Function ExtractHubExportTimestamp(ByVal hubDetail As String) As String
    Const marker As String = "(modified "
    Dim p1 As Long, p2 As Long
    p1 = InStr(1, hubDetail, marker)
    If p1 = 0 Then Exit Function
    p1 = p1 + Len(marker)
    p2 = InStr(p1, hubDetail, ")")
    If p2 = 0 Then Exit Function
    ExtractHubExportTimestamp = Trim(Mid(hubDetail, p1, p2 - p1))
End Function

' =============================================================================
' RFID TYPE LIST -- bulk removal by technology (2026-08.25)
'
' Column M lists every DISTINCT RFID technology present in column G on this
' tab; column N puts a checkbox beside each. Ticking one marks every job of
' that technology for removal.
'
' It works two ways on purpose, because the user asked for both:
'   * LIVE -- Workbook_SheetChange in ThisWorkbook calls RfidToggleFromSheet
'     below, which ticks or unticks the column L box on every matching job the
'     moment the N box is clicked. Instant feedback.
'   * AT REFRESH -- HarvestTickedRemovals reads the N ticks directly and
'     removes those jobs regardless of whether the event ever fired. This is
'     the path that must not fail: event code in a document module can be
'     disabled by macro policy or lost if the module is reset, and a checkbox
'     that looks live but does nothing is worse than one that plainly waits
'     for a button press.
'
' Blank RFID types are deliberately NOT listed (confirmed with the user
' 2026-08.25). A blank means Hub's description matched none of its tokens;
' those jobs are still removable one at a time from their own column L box.
'
' The N ticks reset on every rebuild, exactly as the L ticks do. Whatever they
' removed is already recorded on the Removed Jobs sheet, so nothing is lost,
' and no standing filter can quietly keep deleting work nobody re-approved.
' =============================================================================
Private Sub BuildRfidTypeList(ByRef ws As Worksheet, ByVal lastRow As Long, _
        ByVal tabName As String, ByRef rfidOff As Object)
    ' Clear generously -- last run may have listed more technologies than this.
    On Error Resume Next
    ws.Range(ws.Cells(1, COL_OUT_RFIDLIST), ws.Cells(lastRow + 400, COL_OUT_RFIDTICK)).ClearFormats
    On Error GoTo 0
    ws.Range(ws.Cells(1, COL_OUT_RFIDLIST), ws.Cells(lastRow + 400, COL_OUT_RFIDTICK)).ClearContents

    ' Technologies present on the tab...
    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")
    Dim r As Long, t As String
    If lastRow >= 2 Then
        For r = 2 To lastRow
            If Trim(CStr(ws.Cells(r, COL_OUT_JOBID).Value)) <> "" Then
                t = Trim(CStr(ws.Cells(r, COL_OUT_RFID).Value))
                If t <> "" Then
                    If Not seen.Exists(UCase(t)) Then seen.Add UCase(t), t
                End If
            End If
        Next r
    End If

    ' ...PLUS every technology currently switched off for this machine.
    ' Without this the list would be a one-way door: filtering a technology
    ' removes its last job from the tab, the type stops appearing in column G,
    ' and the tick box the operator needs in order to UNDO it disappears with
    ' it. These entries are exactly the ones that must stay visible.
    Dim k As Variant, parts As Variant, storedTab As String, storedType As String
    If Not rfidOff Is Nothing Then
        For Each k In rfidOff.Keys
            parts = Split(CStr(k), "|")
            If UBound(parts) = 1 Then
                storedTab = CStr(parts(0))
                storedType = CStr(parts(1))
                If storedTab = UCase(Trim(tabName)) Then
                    If Not seen.Exists(storedType) Then seen.Add storedType, storedType
                End If
            End If
        Next k
    End If

    If seen.Count = 0 Then Exit Sub

    Dim types() As String
    ReDim types(1 To seen.Count)
    Dim n As Long
    n = 0
    For Each k In seen.Keys
        n = n + 1
        types(n) = CStr(seen(k))
    Next k

    Dim i As Long, j As Long, tmp As String
    For i = 1 To n - 1
        For j = i + 1 To n
            If UCase(types(j)) < UCase(types(i)) Then
                tmp = types(i): types(i) = types(j): types(j) = tmp
            End If
        Next j
    Next i

    ws.Cells(1, COL_OUT_RFIDLIST).Value = "RFID Type"
    ws.Cells(1, COL_OUT_RFIDTICK).Value = "Remove All"
    With ws.Range(ws.Cells(1, COL_OUT_RFIDLIST), ws.Cells(1, COL_OUT_RFIDTICK))
        .Font.Bold = True
        .Interior.Color = RGB(31, 78, 120)
        .Font.Color = RGB(255, 255, 255)
    End With

    For i = 1 To n
        ws.Cells(i + 1, COL_OUT_RFIDLIST).Value = types(i)
    Next i
    ws.Range(ws.Cells(2, COL_OUT_RFIDLIST), ws.Cells(n + 1, COL_OUT_RFIDLIST)).NumberFormat = "@"

    On Error Resume Next
    ws.Range(ws.Cells(2, COL_OUT_RFIDTICK), ws.Cells(n + 1, COL_OUT_RFIDTICK)).CellControl.SetCheckbox
    On Error GoTo 0

    ' Restore the tick state. The checkbox is wiped with the tab, so the store
    ' in Removed Jobs I:J is what actually carries a filter across a rebuild.
    Dim vals() As Variant
    ReDim vals(1 To n, 1 To 1)
    Dim onCount As Long: onCount = 0
    For i = 1 To n
        vals(i, 1) = False
        If Not rfidOff Is Nothing Then
            If rfidOff.Exists(RfidFilterKey(tabName, types(i))) Then
                vals(i, 1) = True
                onCount = onCount + 1
            End If
        End If
    Next i
    ws.Range(ws.Cells(2, COL_OUT_RFIDTICK), ws.Cells(n + 1, COL_OUT_RFIDTICK)).Value = vals
    ws.Range(ws.Cells(2, COL_OUT_RFIDTICK), ws.Cells(n + 1, COL_OUT_RFIDTICK)).HorizontalAlignment = xlHAlignCenter

    ' A switched-off technology is called out, so a tab that looks short is
    ' never a mystery -- the reason is sitting next to the schedule.
    If onCount > 0 Then
        For i = 1 To n
            If vals(i, 1) = True Then
                ws.Cells(i + 1, COL_OUT_RFIDLIST).Font.Bold = True
                ws.Cells(i + 1, COL_OUT_RFIDLIST).Font.Italic = True
            End If
        Next i
        ws.Cells(n + 3, COL_OUT_RFIDLIST).Value = onCount & " type(s) switched OFF -- untick and Refresh Data to bring those jobs back"
        ws.Cells(n + 3, COL_OUT_RFIDLIST).Font.Italic = True
        ws.Cells(n + 3, COL_OUT_RFIDLIST).Font.Size = 9
    End If

    ws.Range(ws.Cells(1, COL_OUT_RFIDLIST), ws.Cells(n + 1, COL_OUT_RFIDTICK)).BorderAround _
        LineStyle:=xlContinuous, Weight:=xlThin
    ws.Columns(COL_OUT_RFIDLIST).ColumnWidth = 18
    ws.Columns(COL_OUT_RFIDTICK).ColumnWidth = 11
End Sub


' Reads the ticked technologies off one machine tab.
Private Function TickedRfidTypes(ByRef ws As Worksheet) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    Set TickedRfidTypes = d
    If Trim(CStr(ws.Cells(1, COL_OUT_RFIDLIST).Value)) <> "RFID Type" Then Exit Function

    Dim r As Long, t As String
    For r = 2 To 400
        t = Trim(CStr(ws.Cells(r, COL_OUT_RFIDLIST).Value))
        If t = "" Then Exit For
        If TickIsSet(ws.Cells(r, COL_OUT_RFIDTICK).Value) Then
            ' Value is the type as WRITTEN, so it can be stored and shown back
            ' verbatim; the key is upper-cased purely for matching.
            If Not d.Exists(UCase(t)) Then d.Add UCase(t), t
        End If
    Next r
End Function


' Called from Workbook_SheetChange in ThisWorkbook. Public because a document
' module has to be able to reach it. Every exit is silent: this runs on every
' single edit anyone makes anywhere in the workbook, so it must be cheap to
' reject and must never raise a dialog.
Public Sub RfidToggleFromSheet(ByVal Sh As Object, ByVal Target As Range)
    On Error GoTo Bail
    If Sh Is Nothing Or Target Is Nothing Then Exit Sub

    Dim ws As Worksheet
    Set ws = Nothing
    On Error Resume Next
    Set ws = Sh
    On Error GoTo Bail
    If ws Is Nothing Then Exit Sub

    ' Machine tabs only -- identified by their own headers, not by name.
    If Trim(CStr(ws.Cells(1, COL_OUT_REMOVE).Value)) <> "Remove" Then Exit Sub
    If Trim(CStr(ws.Cells(1, COL_OUT_RFIDLIST).Value)) <> "RFID Type" Then Exit Sub

    Dim hit As Range
    Set hit = Application.Intersect(Target, ws.Columns(COL_OUT_RFIDTICK))
    If hit Is Nothing Then Exit Sub

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, COL_OUT_JOBID).End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    Dim prevEvents As Boolean
    prevEvents = Application.EnableEvents
    Application.EnableEvents = False

    Dim c As Range, wantOn As Boolean, t As String, r As Long
    For Each c In hit.Cells
        If c.Row >= 2 Then
            t = UCase(Trim(CStr(ws.Cells(c.Row, COL_OUT_RFIDLIST).Value)))
            If t <> "" Then
                wantOn = TickIsSet(c.Value)
                For r = 2 To lastRow
                    If Trim(CStr(ws.Cells(r, COL_OUT_JOBID).Value)) <> "" Then
                        If UCase(Trim(CStr(ws.Cells(r, COL_OUT_RFID).Value))) = t Then
                            ws.Cells(r, COL_OUT_REMOVE).Value = wantOn
                        End If
                    End If
                Next r
            End If
        End If
    Next c

    Application.EnableEvents = prevEvents
    Exit Sub
Bail:
    Application.EnableEvents = True
End Sub


' =============================================================================
' Formatting must never abort a rebuild.
'
' On 2026-08-24 Excel refused NumberFormat on the Removed Jobs sheet with
' "1004 - Unable to set the NumberFormat property of the Range class" and took
' the entire Refresh down with it -- twice, on two different ranges, in two
' different versions. The schedule matters; the thousands separator does not.
'
' So every NumberFormat assignment in this module goes through here. It applies
' the format when Excel allows it and records the range, the format string and
' Excel's own words when it does not, then carries on. What was recorded is
' shown once at the end of the run, which means a failure diagnoses itself in
' place instead of costing a round trip.
'
' Nothing functional rides on these formats. Job IDs round-trip through CStr()
' wherever they are read back, and Excel auto-formats a Date value the moment
' it is assigned, so dates stay readable even when the explicit format is
' rejected.
' =============================================================================
Private Sub SetFmt(ByRef rng As Range, ByVal fmt As String, ByVal where As String)
    If rng Is Nothing Then Exit Sub
    On Error Resume Next
    Err.Clear
    rng.NumberFormat = fmt
    If Err.Number <> 0 Then
        ' Capped so a sheet-wide refusal cannot build a MsgBox nobody can read.
        If Len(gFmtWarn) < 700 Then
            ' Chr(34) rather than doubled quotes so the format string reads
            ' cleanly in the message and this line stays easy to edit.
            gFmtWarn = gFmtWarn & vbCrLf & "  " & where & " (" & rng.Address(False, False) & _
                       ") -> " & Chr(34) & fmt & Chr(34) & ": " & Err.Number & " - " & Err.Description
        End If
        Err.Clear
    End If
    On Error GoTo 0
End Sub

' Empty unless Excel refused at least one number format this run.
Private Function FormatWarning() As String
    FormatWarning = ""
    If Len(gFmtWarn) > 0 Then
        FormatWarning = vbCrLf & vbCrLf & _
            "NOTE: Excel refused some number formats. The schedule itself is correct and " & _
            "complete -- only the display formatting of these ranges was skipped:" & gFmtWarn & _
            vbCrLf & vbCrLf & "Send this text to whoever maintains this workbook."
    End If
End Function

' Empty unless this build tried and failed to draw the cell checkboxes. Says
' so out loud rather than leaving an operator staring at a blank column L --
' the typed-mark fallback in TickIsSet still works and this tells them to use it.
Private Function CheckboxWarning() As String
    CheckboxWarning = ""
    If gCheckboxTried And Not gCheckboxOK Then
        CheckboxWarning = vbCrLf & vbCrLf & _
            "NOTE: this build of Excel would not draw the Remove checkboxes, so column L is " & _
            "an ordinary cell. Type anything in it -- an x will do -- to mark a job for removal."
    End If
End Function

Private Sub ShowPreferencesWhenDone()
    On Error Resume Next
    ThisWorkbook.Sheets("Preferences").Activate
    On Error GoTo 0
End Sub

' =============================================================================
' SHEET / BUTTON MANAGEMENT
' =============================================================================
Private Function SanitizeSheetName(ByVal raw As String) As String
    Dim s As String
    s = Trim(raw)
    Dim badChars As Variant, bc As Variant
    badChars = Array(":", "\", "/", "?", "*", "[", "]")
    For Each bc In badChars
        s = Replace(s, bc, "")
    Next bc
    If Len(s) > 31 Then s = Left(s, 31)
    If s = "" Then s = "Machine"
    SanitizeSheetName = s
End Function

Private Function GetOrCreateSheet(ByVal rawName As String) As Worksheet
    Dim nm As String: nm = SanitizeSheetName(rawName)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(nm)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.name = nm
    End If
    Set GetOrCreateSheet = ws
End Function

' Reorders tabs to: Preferences, then one tab per roster row in roster order.
' Any tab not in the current roster (renamed/removed machine) is left in
' place at the end, untouched -- never auto-deleted (see Preferences notes).
Private Sub EnforcePressWorksheetOrder(ByRef roster As Variant)
    gStep = "reordering the worksheet tabs"
    Dim prefsWs As Worksheet
    On Error Resume Next
    Set prefsWs = ThisWorkbook.Sheets("Preferences")
    On Error GoTo 0
    If Not prefsWs Is Nothing Then
        prefsWs.Move Before:=ThisWorkbook.Sheets(1)
    End If

    Dim m As Long, tabName As String, ws As Worksheet
    Dim afterSheet As Worksheet
    Set afterSheet = prefsWs
    For m = 1 To UBound(roster, 1)
        tabName = SanitizeSheetName(CStr(roster(m, 1)))
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(tabName)
        On Error GoTo 0
        If Not ws Is Nothing And Not afterSheet Is Nothing Then
            ws.Move After:=afterSheet
            Set afterSheet = ws
        End If
    Next m

    ' Removed Jobs sits straight after the machine tabs it belongs to.
    Dim remWs As Worksheet
    On Error Resume Next
    Set remWs = ThisWorkbook.Sheets(REMOVED_SHEET)
    On Error GoTo 0
    If Not remWs Is Nothing And Not afterSheet Is Nothing Then
        remWs.Move After:=afterSheet
    End If
End Sub

Private Sub EnsureUpdateButton()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Preferences")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    Dim btn As Shape
    On Error Resume Next
    Set btn = ws.Shapes("btnUpdatePressData")
    On Error GoTo 0

    If btn Is Nothing Then
        Set btn = ws.Shapes.AddShape(msoShapeRoundedRectangle, ws.Cells(4, 5).Left, ws.Cells(4, 5).Top, 140, 36)
        btn.name = "btnUpdatePressData"
        btn.OnAction = "UpdatePressData"

        ' Styling applied ONLY at creation, never on later runs -- previously
        ' this block ran unconditionally every time Update Data was clicked,
        ' so any manual color change made in Excel got stomped back to this
        ' fixed blue on the very next run. Same bug class as the earlier
        ' Preferences-sheet AutoFit issue. From here on, manual formatting
        ' changes made directly in Excel (e.g. changing the fill color)
        ' persist across Update Data runs.
        With btn
            .TextFrame2.TextRange.Text = "Update Data"
            With .TextFrame2.TextRange.Font
                .Bold = msoTrue
                .Size = 14
                .Fill.ForeColor.RGB = RGB(255, 255, 255)
            End With
            .Fill.ForeColor.RGB = RGB(31, 78, 120)
            .Line.Visible = msoFalse
        End With
    End If
End Sub

-------------------------------------------------------------------------------
