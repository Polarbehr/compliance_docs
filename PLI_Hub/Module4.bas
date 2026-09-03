Option Explicit

' =============================================================================
' PLI HUB WORKBOOK -- Module4 (v52)
' Cleaning engine only. Reads the shared Monarch drop folder, applies the
' business rules (rate math REMOVED 2026-08.12 -- lives in the AVL Dashboard;
' LAMBDAs), maintains Parameters (machine rates), and exposes CleanedData +
' Parameters for satellite workbooks (AVL Production Dashboard, Press,
' Finishing) to pull via Workbooks.Open ReadOnly + copy + Close.
'
' Derived from the AVL Production Dashboard workbook's Module4 (v50) --
' every cleaning/business-rule procedure carried over UNCHANGED. Removed:
' Dashboard/Charts/Capacity building (BuildDashboardAnalytics,
' RefreshDashboardOnly, BuildDashboardFromScratch, FormatCapacitySheet,
' RefreshDashboardDynamicElements, SyncChartsSheet and their helpers) --
' those stay in the AVL Production Dashboard workbook, which now pulls its
' data from here instead of computing it locally. Also removed: dashboard
' color/swatch machinery (this workbook has no dashboard to color) beyond
' the minimum GetDashSwatch/DashFillOrDefault plumbing that BuildFormulasSheet
' still uses for its header styling.
'
' NEW in this workbook: a "Hub Status" flag on Preferences (row added by
' EnsurePreferenceRows) written REFRESHING at the start of Update Data and
' READY <timestamp> at the end. Every satellite pull checks this flag
' first and warns rather than reading a mid-rebuild CleanedData table.
'
' 2026-09.03 (v52): Monarch's export changed its column layout again and
' broke the old hardcoded column numbers -- Customer/Description/QTY/
' Location Date/Location/Work Center all came back blank. Column lookup
' is now header-text-driven (see FindMonarchColumn / FindMonarchDataColumn
' / FindMonarchDenseColumn below ProcessAndDistributeAllData) instead of
' fixed positions, and fails loudly if a header can't be found. Also picks
' up the new "To Perso" field Monarch's export started including, as
' CleanedData column L -- appended at the end so it doesn't shift Week
' Start (J) or RFID Type (K) for anything already reading those by
' position. Blank on exports that don't have the column (older layouts)
' or jobs that haven't reached that stage.
' =============================================================================

Private Const THEME_ACCENT2 As Long = 6
Private Const THEME_ACCENT5 As Long = 9

' =============================================================================
' MAIN ENTRY POINT -- "Update Data" button
' =============================================================================
Public Sub ProcessAndDistributeAllData()
    Dim t0 As Single: t0 = Timer
    MigrateSheetNames
    RemoveParametersArtifacts

    ' Ensure the shared calculation LAMBDAs exist before anything references them

    ' Snapshot the user-entered Parameters values BEFORE anything else
    ' touches any sheet (see BackupSettingsParameters).

    ' One-time migration (idempotent, safe to run every time): strips the
    ' Dashboard-only color-swatch rows this Preferences sheet inherited
    ' from the AVL Production Dashboard workbook it was split from -- a
    ' Hub has no dashboard to color. Same self-healing pattern as
    ' MigrateSheetNames / RepairStrandedMachineColumns.
    MigratePreferencesToHubLayout

    ' Make sure the Preferences sheet (folder path + status flag) exists.
    EnsurePreferencesSheet

    ' Self-healing button, same pattern as the original Dashboard's
    ' "Refresh Only" button -- restyled every run so it can never drift.
    EnsureUpdateDataButton

    ' Regenerate the Formulas documentation sheet -- cheap, static content;
    ' rebuilt every run so it always matches the code.
    BuildFormulasSheet

    ' Tell any satellite mid-pull right now that this data is turning over.
    SetHubStatus "REFRESHING", ""

    On Error GoTo BuildFailed

    ' -------------------------------------------------------------
    ' 1. SETUP WORKSHEETS & SOURCE DATA
    ' The "Monarch Report Folder" path on Preferences is the ONLY data
    ' path in the Hub (the Monarch Import paste tab was retired -- see
    ' HANDOFF v49/v50 "USER'S COPY STATE"). If the share fails, Update
    ' Data stops with an error -- there is no local paste fallback here.
    ' -------------------------------------------------------------
    Dim rawData As Variant
    Dim srcDesc As String
    Dim autoFailed As Boolean

    If Not LoadMonarchFromSharedFolder(rawData, srcDesc, autoFailed) Then
        SetHubStatus "ERROR", "Monarch folder read failed at " & Format(Now, "yyyy-mm-dd hh:nn")
        MsgBox "Could not read the Monarch Report Folder -- see the message above for the reason." & vbCrLf & _
               "Check the path on the Preferences sheet and that the share is reachable.", vbCritical, "Update Data Failed"
        GoTo CleanExit
    End If

    If IsEmpty(rawData) Or UBound(rawData, 1) < 2 Then
        SetHubStatus "ERROR", "No data rows found at " & Format(Now, "yyyy-mm-dd hh:nn")
        MsgBox "No data rows found (source: " & srcDesc & ").", vbExclamation
        GoTo CleanExit
    End If

    Dim totalRows As Long, totalCols As Long
    totalRows = UBound(rawData, 1)
    totalCols = UBound(rawData, 2)

    ' -------------------------------------------------------------
    ' 2. LOCATE THE MONARCH COLUMNS BY HEADER LABEL
    ' -------------------------------------------------------------
    ' 2026-09.03: Monarch's "Open Jobs By Work Center" export keeps
    ' inserting/removing blank spacer columns between fields -- it just did
    ' it again, shifting Customer/Description/QTY/Location Date/Location by
    ' a column and Work Center by two, plus adding a brand-new "To Perso"
    ' field at the far right that nothing here reads. Hardcoded column
    ' NUMBERS broke on every one of those changes (that's what broke this
    ' time), so columns are now found by their header TEXT instead, with a
    ' one-cell search window to absorb the fact Crystal doesn't always put
    ' a field's header directly above its data (Customer Name's data has
    ' been seen both under its header and one cell right of it; Last/Gang
    ' Location's headers sit one cell right of their data in every export
    ' seen so far). If Monarch reshuffles columns again this keeps working
    ' as long as the header text itself doesn't change; if a header goes
    ' missing, Update Data now fails loudly instead of quietly writing an
    ' empty CleanedData column the way this bug did.
    Dim headerRow As Long
    headerRow = FindMonarchHeaderRow(rawData, totalRows, totalCols)
    If headerRow = 0 Then
        SetHubStatus "ERROR", "Monarch header row not found at " & Format(Now, "yyyy-mm-dd hh:nn")
        MsgBox "Could not find the Monarch report's header row (expected a 'Ship' column heading within the first 20 rows). The report layout may have changed -- open the source file and check the Preferences sheet's Monarch Report Folder.", vbCritical, "Update Data Failed"
        GoTo CleanExit
    End If

    Dim claimedCols As Object
    Set claimedCols = CreateObject("Scripting.Dictionary")

    Dim cShip As Long, cProm As Long, cJobID As Long
    cShip = FindMonarchColumn(rawData, headerRow, totalCols, "Ship", claimedCols)
    If cShip > 0 Then claimedCols.Add cShip, True
    cProm = FindMonarchColumn(rawData, headerRow, totalCols, "Prom", claimedCols)
    If cProm > 0 Then claimedCols.Add cProm, True
    cJobID = FindMonarchColumn(rawData, headerRow, totalCols, "Job ID", claimedCols)
    If cJobID > 0 Then claimedCols.Add cJobID, True

    If cShip = 0 Or cProm = 0 Or cJobID = 0 Then
        SetHubStatus "ERROR", "Monarch column headings not recognised at " & Format(Now, "yyyy-mm-dd hh:nn")
        MsgBox "Could not find the 'Ship', 'Prom', or 'Job ID' column headings in the Monarch export. The report layout may have changed -- open the source file and confirm those headings still exist.", vbCritical, "Update Data Failed"
        GoTo CleanExit
    End If

    Dim dataRows As Object
    Set dataRows = SampleMonarchDataRows(rawData, headerRow, totalRows, cShip, cJobID)

    Dim hCust As Long, hDesc As Long, hQty As Long, hLastLoc As Long, hGangLoc As Long
    hCust = FindMonarchColumn(rawData, headerRow, totalCols, "Customer Name", claimedCols)
    hDesc = FindMonarchColumn(rawData, headerRow, totalCols, "Job Description", claimedCols)
    hQty = FindMonarchColumn(rawData, headerRow, totalCols, "Qnty", claimedCols)
    hLastLoc = FindMonarchColumn(rawData, headerRow, totalCols, "Last Location", claimedCols)
    hGangLoc = FindMonarchColumn(rawData, headerRow, totalCols, "Gang Location", claimedCols)

    Dim cCust As Long, cDesc As Long, cQty As Long, cLocDate As Long, cLoc As Long, cWorkCenter As Long
    cCust = FindMonarchDataColumn(rawData, dataRows, hCust, claimedCols, False)
    cDesc = FindMonarchDataColumn(rawData, dataRows, hDesc, claimedCols, False)
    cQty = FindMonarchDataColumn(rawData, dataRows, hQty, claimedCols, False)
    cLocDate = FindMonarchDataColumn(rawData, dataRows, hLastLoc, claimedCols, True)
    cLoc = FindMonarchDataColumn(rawData, dataRows, hGangLoc, claimedCols, False)
    ' Work Center has no header label of its own -- see FindMonarchDenseColumn.
    cWorkCenter = FindMonarchDenseColumn(rawData, dataRows, cLoc, totalCols, claimedCols)

    ' To Perso is OPTIONAL -- Monarch only started including it in the
    ' 2026-09.03 export, older layouts don't have it at all, and it isn't
    ' filled for every job even when the column exists. Missing/unfound
    ' just means CleanedData carries a blank column L, not an error.
    Dim hToPerso As Long, cToPerso As Long
    hToPerso = FindMonarchColumn(rawData, headerRow, totalCols, "To Perso", claimedCols)
    cToPerso = FindMonarchDataColumn(rawData, dataRows, hToPerso, claimedCols, True)

    If cCust = 0 Or cDesc = 0 Or cQty = 0 Or cLocDate = 0 Or cLoc = 0 Or cWorkCenter = 0 Then
        SetHubStatus "ERROR", "Monarch column headings not recognised at " & Format(Now, "yyyy-mm-dd hh:nn")
        MsgBox "Could not confidently locate every Monarch column (Customer Name / Job Description / Qnty / Last Location / Gang Location / Work Center) in this export. The report layout may have changed more than this Hub can self-heal -- open the source file, compare its column headings to what ProcessAndDistributeAllData (Module4) expects, and update the code if needed.", vbCritical, "Update Data Failed"
        GoTo CleanExit
    End If

    ' -------------------------------------------------------------
    ' 2b. DEFINE EXPLICIT INCLUSIONS & HEADERS
    ' -------------------------------------------------------------
    Dim keepCols As Object
    Set keepCols = CreateObject("Scripting.Dictionary")

    keepCols.Add cShip, "Ship Date"
    keepCols.Add cProm, "Promised Date"
    keepCols.Add cJobID, "Job ID"
    keepCols.Add cCust, "Customer"
    keepCols.Add cDesc, "Description"
    keepCols.Add cQty, "QTY"
    keepCols.Add cLocDate, "Location Date"
    keepCols.Add cLoc, "Location"
    keepCols.Add cWorkCenter, "Work Center"

    ' 2026-08.13: CleanedData gains an 11th column, RFID Type -- the
    ' technology token found in the job's Description. Hub's only job with
    ' it is to CREATE the column; every satellite does its own analysis
    ' from there (the user's ruling: 'hub should create the column based on
    ' the description. That's its only job').
    Dim colWeekStart As Long, colRfid As Long
    colWeekStart = keepCols.Count + 1
    colRfid = keepCols.Count + 2

    ' 2026-09.03: CleanedData gains a 12th column, To Perso -- Monarch's
    ' own field, copied through as-is (blank on exports/jobs that don't
    ' have it). Appended after RFID Type so Week Start/RFID Type keep
    ' their existing column letters for anything already reading them by
    ' position.
    Dim colToPerso As Long
    colToPerso = keepCols.Count + 3
    Dim totalOutCols As Long
    totalOutCols = colToPerso

    ' -------------------------------------------------------------
    ' 3. TRANSFORM DATA IN MEMORY
    ' -------------------------------------------------------------
    Dim outputData() As Variant
    ReDim outputData(1 To totalRows, 1 To totalOutCols)

    Dim outC As Long: outC = 1
    Dim colKey As Variant
    For Each colKey In keepCols.Keys
        outputData(1, outC) = keepCols(colKey)
        outC = outC + 1
    Next colKey
    outputData(1, colWeekStart) = "Week Start"
    outputData(1, colRfid) = "RFID Type"
    outputData(1, colToPerso) = "To Perso"

    Dim deleteKeywords As Variant
    deleteKeywords = Array("production", "total", "finished", "open", "jobs", "work", "department", "report", "column", "ship")

    ' Token list read ONCE ahead of the loop -- see ReadRfidTokens for the
    ' longest-first ordering that makes overlapping tokens resolve sanely.
    Dim rfidTokens As Variant
    rfidTokens = ReadRfidTokens()

    Dim outR As Long: outR = 1
    Dim r As Long, cellVal As Variant, cellStr As String
    Dim col1Val As Variant, jobIDVal As Variant
    Dim kw As Variant, hasMatch As Boolean
    Dim valLocDate As Variant, valLoc As Variant
    Dim refDate As Variant, weekStartVal As Variant
    Dim toPersoVal As Variant

    For r = 2 To totalRows
        col1Val = rawData(r, cShip)
        jobIDVal = rawData(r, cJobID)

        If Not IsEmpty(col1Val) And Not IsNull(col1Val) And Not IsEmpty(jobIDVal) And Not IsNull(jobIDVal) Then
            cellStr = LCase(CStr(col1Val))
            hasMatch = False

            For Each kw In deleteKeywords
                If InStr(1, cellStr, kw) > 0 Then
                    hasMatch = True
                    Exit For
                End If
            Next kw

            If Not hasMatch Then
                outR = outR + 1

                valLocDate = rawData(r, cLocDate)
                valLoc = rawData(r, cLoc)

                If Not IsEmpty(valLocDate) And Not IsNull(valLocDate) Then
                    If Not IsDate(valLocDate) Then
                        valLoc = valLocDate
                        valLocDate = Empty
                    End If
                End If
                rawData(r, cLocDate) = valLocDate
                rawData(r, cLoc) = valLoc

                outC = 1
                For Each colKey In keepCols.Keys
                    If CLng(colKey) <= totalCols Then
                        cellVal = rawData(r, CLng(colKey))

                        If colKey = cShip Or colKey = cProm Or colKey = cLocDate Then
                            If IsDate(cellVal) And Not IsEmpty(cellVal) Then
                                cellVal = CDate(cellVal)
                            Else
                                cellVal = Empty
                            End If
                        End If

                        outputData(outR, outC) = cellVal
                    Else
                        outputData(outR, outC) = Empty
                    End If
                    outC = outC + 1
                Next colKey

                refDate = rawData(r, cShip)
                If IsEmpty(refDate) Or Not IsDate(refDate) Then
                    refDate = rawData(r, cProm)
                End If

                If IsDate(refDate) And Not IsEmpty(refDate) Then
                    weekStartVal = CDate(refDate) - Weekday(CDate(refDate), vbMonday) + 1
                    outputData(outR, colWeekStart) = CDate(weekStartVal)
                Else
                    outputData(outR, colWeekStart) = Empty
                End If

                ' Description is out-column 5 (source column found dynamically
                ' above as cDesc). A job with no recognised token gets an
                ' EMPTY cell rather than a placeholder, so satellites can
                ' filter on blank.
                outputData(outR, colRfid) = ClassifyRfidType(CStr(outputData(outR, 5)), rfidTokens)

                If cToPerso > 0 And cToPerso <= totalCols Then
                    toPersoVal = rawData(r, cToPerso)
                    If IsDate(toPersoVal) And Not IsEmpty(toPersoVal) Then
                        outputData(outR, colToPerso) = CDate(toPersoVal)
                    Else
                        outputData(outR, colToPerso) = Empty
                    End If
                Else
                    outputData(outR, colToPerso) = Empty
                End If
            End If
        End If
    Next r

    ' -------------------------------------------------------------
    ' 4. MULTI-LEVEL SORT (Work Center -> Ship Date -> Job ID)
    ' -------------------------------------------------------------
    If outR > 2 Then
        QuickSortMultiKey outputData, 2, outR, 9, 1, 3, totalOutCols
    End If

    ' -------------------------------------------------------------
    ' 5. BULK WRITE CLEANED DATA
    ' -------------------------------------------------------------
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim destSheet As Worksheet
    Set destSheet = GetOrCreateSheet("CleanedData")
    destSheet.Cells.Clear

    If outR > 0 Then
        WriteAndFormatSheet destSheet, outputData, outR, totalOutCols
    End If

    ' Generate Parameters machine-roster maintenance + Week Start Summary

    ' One full Calculate so Effective QTY / Job Hours / Capacity Load are
    ' real values before any satellite pulls them (an uncalculated cell
    ' pulled as a value silently reads blank) -- same lesson as the
    ' original Dashboard pivot cache bug.
    Application.Calculate

    EnforceWorksheetOrder

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True

    SetHubStatus "READY", srcDesc & " | updated " & Format(Now, "yyyy-mm-dd hh:nn")
    ShowPreferencesWhenDone

    MsgBox "Hub data processing complete in " & Format(Timer - t0, "0.00") & " seconds!" & vbCrLf & vbCrLf & _
           "Source: " & srcDesc & vbCrLf & vbCrLf & _
           "Satellite workbooks (AVL Production Dashboard, Press, Finishing) can now pull this data with their own Update buttons.", vbInformation
    Exit Sub

CleanExit:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ShowPreferencesWhenDone
    Exit Sub

BuildFailed:
    Dim errMsg As String, errNum As Long
    errMsg = Err.Description
    errNum = Err.Number
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    SetHubStatus "ERROR", "Build failed (" & errNum & ": " & errMsg & ") at " & Format(Now, "yyyy-mm-dd hh:nn")
    ShowPreferencesWhenDone
    MsgBox "Update Data failed: " & errNum & " - " & errMsg, vbCritical, "Hub Update Data Failed"
End Sub

' =============================================================================
' MONARCH COLUMN DETECTION -- see the "2. LOCATE THE MONARCH COLUMNS BY
' HEADER LABEL" comment in ProcessAndDistributeAllData for why this exists:
' Monarch's export has changed its column layout more than once, and
' hardcoded column numbers silently broke every time. These four helpers
' find each field by its header text and by how real data rows are
' actually populated, instead of by a fixed position.
' =============================================================================

' Scans the first rows of a Monarch export for the header row -- the one
' whose first cell is exactly "Ship". Crystal always puts several title/
' filter-criteria rows above it and that count has varied, so this can't
' be assumed to be any particular fixed row number.
Private Function FindMonarchHeaderRow(ByRef rawData As Variant, ByVal totalRows As Long, ByVal totalCols As Long) As Long
    Dim r As Long, scanLimit As Long
    scanLimit = totalRows
    If scanLimit > 20 Then scanLimit = 20
    If totalCols < 1 Then Exit Function
    For r = 1 To scanLimit
        If Trim(CStr(rawData(r, 1))) = "Ship" Then
            FindMonarchHeaderRow = r
            Exit Function
        End If
    Next r
End Function

' Finds a column by exact header-label match on the header row, skipping
' any column already claimed for another field. Returns 0 if not found.
Private Function FindMonarchColumn(ByRef rawData As Variant, ByVal headerRow As Long, ByVal totalCols As Long, ByVal label As String, ByRef claimedCols As Object) As Long
    Dim c As Long
    For c = 1 To totalCols
        If Not claimedCols.Exists(c) Then
            If Trim(CStr(rawData(headerRow, c))) = label Then
                FindMonarchColumn = c
                Exit Function
            End If
        End If
    Next c
End Function

' Row numbers (within rawData) that look like real job rows: a non-blank
' Ship cell and a non-blank Job ID, and not one of Crystal's title/
' subtotal/group-header rows. This is the same test the main transform
' loop applies to decide what to keep, factored out so column detection
' scores candidates against exactly the rows that will actually be
' written to CleanedData.
Private Function SampleMonarchDataRows(ByRef rawData As Variant, ByVal headerRow As Long, ByVal totalRows As Long, ByVal cShip As Long, ByVal cJobID As Long) As Object
    Dim foundRows As Object
    Set foundRows = CreateObject("Scripting.Dictionary")

    Dim deleteKeywords As Variant
    deleteKeywords = Array("production", "total", "finished", "open", "jobs", "work", "department", "report", "column", "ship")

    Dim r As Long, col1Val As Variant, jobIDVal As Variant, cellStr As String
    Dim kw As Variant, hasMatch As Boolean
    For r = headerRow + 1 To totalRows
        col1Val = rawData(r, cShip)
        jobIDVal = rawData(r, cJobID)
        If Not IsEmpty(col1Val) And Not IsNull(col1Val) And Not IsEmpty(jobIDVal) And Not IsNull(jobIDVal) Then
            cellStr = LCase(CStr(col1Val))
            hasMatch = False
            For Each kw In deleteKeywords
                If InStr(1, cellStr, kw) > 0 Then
                    hasMatch = True
                    Exit For
                End If
            Next kw
            If Not hasMatch Then foundRows.Add r, True
        End If
    Next r
    Set SampleMonarchDataRows = foundRows
End Function

' A field's data isn't always directly under its own header cell -- Crystal
' has shifted it a column either way between exports (Customer Name's data
' has been seen dead-on and one column right of its header; Last/Gang
' Location are consistently one column LEFT of theirs). Rather than
' hardcode a direction, this checks the header column and its immediate
' neighbors and picks whichever is most consistently populated across real
' data rows -- skipping columns already claimed for another field.
'
' wantDate flips the fill test to only count cells Excel actually typed as
' a date (VarType = vbDate, i.e. the source cell was date-formatted) --
' NOT IsDate(), which also returns True for any plain number in a
' plausible date-serial range and would misclassify a QTY column as
' "date-like". This lets a date field (Location Date) find its column
' without accidentally locking onto a neighboring text or numeric field
' that happens to be well populated too, and vice versa.
Private Function FindMonarchDataColumn(ByRef rawData As Variant, ByRef dataRows As Object, ByVal headerCol As Long, ByRef claimedCols As Object, ByVal wantDate As Boolean) As Long
    If headerCol = 0 Then Exit Function
    If dataRows.Count = 0 Then Exit Function

    Dim bestCol As Long, bestScore As Double
    bestCol = 0: bestScore = -1

    Dim c As Long, rKey As Variant, v As Variant, filled As Long
    Dim isDateCell As Boolean, nonBlank As Boolean, score As Double

    For c = headerCol - 1 To headerCol + 1
        If c >= 1 And Not claimedCols.Exists(c) Then
            filled = 0
            For Each rKey In dataRows.Keys
                v = rawData(CLng(rKey), c)
                nonBlank = Not IsEmpty(v) And Not IsNull(v) And Trim(CStr(v)) <> ""
                isDateCell = nonBlank And (VarType(v) = vbDate)
                If wantDate Then
                    If isDateCell Then filled = filled + 1
                Else
                    If nonBlank And Not isDateCell Then filled = filled + 1
                End If
            Next rKey
            score = filled / dataRows.Count
            If score > bestScore Then
                bestScore = score
                bestCol = c
            End If
        End If
    Next c

    If bestScore > 0 Then
        claimedCols.Add bestCol, True
        FindMonarchDataColumn = bestCol
    End If
End Function

' Work Center has no header label at all in Crystal's output -- it's the
' first column right of Location that's populated on essentially every job
' row. (A pure spacer column runs ~0% filled; the newer "To Perso" field
' this report started including only fires for a fraction of jobs and sits
' further right anyway, so a high fill-rate threshold finds Work Center
' first and correctly skips over both.)
Private Function FindMonarchDenseColumn(ByRef rawData As Variant, ByRef dataRows As Object, ByVal afterCol As Long, ByVal totalCols As Long, ByRef claimedCols As Object) As Long
    If afterCol = 0 Then Exit Function
    If dataRows.Count = 0 Then Exit Function

    Const FILL_THRESHOLD As Double = 0.9
    Dim c As Long, rKey As Variant, v As Variant, filled As Long, score As Double

    For c = afterCol + 1 To totalCols
        If Not claimedCols.Exists(c) Then
            filled = 0
            For Each rKey In dataRows.Keys
                v = rawData(CLng(rKey), c)
                If Not IsEmpty(v) And Not IsNull(v) And Trim(CStr(v)) <> "" Then filled = filled + 1
            Next rKey
            score = filled / dataRows.Count
            If score >= FILL_THRESHOLD Then
                claimedCols.Add c, True
                FindMonarchDenseColumn = c
                Exit Function
            End If
        End If
    Next c
End Function

' One-time migration: removes the Dashboard-only color-swatch rows (and
' their instructional header text) that a Hub split off from the AVL
' Production Dashboard workbook inherits on its Preferences sheet. Finds
' rows BY LABEL rather than fixed position so it self-heals regardless of
' exactly what state the sheet is in. No-op forever once already clean.
Private Sub MigratePreferencesToHubLayout()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Preferences")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    Dim dashboardOnlyLabels As Variant
    dashboardOnlyLabels = Array("Main Header", "Sub Headers", "Column Headers", _
        "Alternating Band 1", "Alternating Band 2", "Grand Total Row", _
        "Overbooked Highlight", "Chart Series Color", "Dashboard Title", "Capacity Title")

    Dim lbl As Variant, r As Long, found As Boolean
    found = False
    For Each lbl In dashboardOnlyLabels
        For r = 1 To 100
            If Trim(CStr(ws.Cells(r, 1).Value)) = CStr(lbl) Then
                ws.Rows(r).ClearContents
                ws.Rows(r).Interior.Pattern = xlNone
                found = True
                Exit For
            End If
        Next r
    Next lbl

    ' Retitle row 1 if it still shows the inherited Dashboard-era heading.
    If Trim(CStr(ws.Cells(1, 1).Value)) = "Dashboard Color Settings" Or _
       Trim(CStr(ws.Cells(1, 1).Value)) = "Preferences" Then
        ws.Cells(1, 1).Value = "Preferences (Hub)"
        ws.Cells(2, 1).Value = "Point 'Monarch Report Folder' at the shared drop folder. Everything else here is status-only."
        ws.Cells(3, 1).ClearContents
        ws.Cells(4, 1).ClearContents
        found = True
    End If
End Sub

' "Update Data" button -- 2026-08.12 rework per the user's request:
'   1. The button now lives on PREFERENCES, not Parameters. Any shape named
'      "HubUpdateDataButton" still sitting on Parameters is deleted every
'      run (standing migration -- the user asked for it gone from there).
'   2. The button is USER-OWNED after creation: if a shape with this name
'      exists on ANY sheet (the user may drag it wherever they like), the
'      only thing this Sub touches is .OnAction (so the click stays wired
'      to the macro). Position, size, fill color, text -- never overwritten
'      on update. Same fix class as the Press workbook's v14 button fix,
'      taken one step further (position was still being reset there).
'   3. Only when NO sheet has the button does it get created fresh on
'      Preferences with the default style below.
' Bootstrapping note: on a workbook where this has never run, run
' ProcessAndDistributeAllData once via Alt+F8 to create the button.
Private Sub EnsureUpdateDataButton()
    Const BTN_NAME As String = "HubUpdateDataButton"

    ' 1. Standing removal from Parameters.
    Dim wsParams As Worksheet
    On Error Resume Next
    Set wsParams = ThisWorkbook.Sheets("Parameters")
    On Error GoTo 0
    If Not wsParams Is Nothing Then
        Dim oldShp As Shape
        On Error Resume Next
        Set oldShp = wsParams.Shapes(BTN_NAME)
        On Error GoTo 0
        If Not oldShp Is Nothing Then oldShp.Delete
    End If

    ' 2. If the button exists anywhere else, re-wire the macro and leave
    '    everything else exactly as the user has it.
    Dim ws As Worksheet, shp As Shape
    For Each ws In ThisWorkbook.Worksheets
        If ws.name <> "Parameters" Then
            Set shp = Nothing
            On Error Resume Next
            Set shp = ws.Shapes(BTN_NAME)
            On Error GoTo 0
            If Not shp Is Nothing Then
                shp.OnAction = "ProcessAndDistributeAllData"
                Exit Sub
            End If
        End If
    Next ws

    ' 3. Not found anywhere -- create fresh on Preferences (creation-time
    '    styling only; never re-applied).
    Dim wsPrefs As Worksheet
    On Error Resume Next
    Set wsPrefs = ThisWorkbook.Sheets("Preferences")
    On Error GoTo 0
    If wsPrefs Is Nothing Then Exit Sub

    Set shp = wsPrefs.Shapes.AddShape(msoShapeRoundedRectangle, _
        wsPrefs.Cells(4, 4).Left, wsPrefs.Cells(4, 4).Top, 160, 34)
    shp.name = BTN_NAME
    shp.OnAction = "ProcessAndDistributeAllData"
    With shp
        .Fill.ForeColor.RGB = RGB(31, 78, 120)
        With .Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(0, 0, 0)
            .Weight = 1.5
        End With
        With .TextFrame2.TextRange
            .Text = "Update Data"
            .Font.Size = 14
            .Font.Bold = msoTrue
            .Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
            .ParagraphFormat.Alignment = msoAlignCenter
        End With
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
    End With
End Sub

' Writes the pull-guard flag satellites check before reading CleanedData.
' Row is added by EnsurePreferenceRows if missing.
Private Sub SetHubStatus(ByVal statusWord As String, ByVal detail As String)
    Dim sw As Range
    Set sw = GetDashSwatch("Hub Status")
    If sw Is Nothing Then Exit Sub
    sw.Value = statusWord & IIf(detail <> "", " | " & detail, "")
End Sub

' Lands the user on Preferences (where Hub Status lives) at the end of every
' Update Data run -- success, early "no report found" return, or a hard
' build failure -- so the outcome is always visible without hunting for it.
' Called from all three exit points of ProcessAndDistributeAllData.
Private Sub ShowPreferencesWhenDone()
    On Error Resume Next
    ThisWorkbook.Sheets("Preferences").Activate
    On Error GoTo 0
End Sub

' =============================================================================
' ENFORCE EXACT WORKSHEET ORDERING (Hub sheet set)
' =============================================================================
Private Sub EnforceWorksheetOrder()
    Dim primarySheets As Variant
    primarySheets = Array("Preferences", "Formulas", "CleanedData")

    Dim i As Long, targetIdx As Long: targetIdx = 1
    Dim ws As Worksheet

    For i = LBound(primarySheets) To UBound(primarySheets)
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(CStr(primarySheets(i)))
        On Error GoTo 0

        If Not ws Is Nothing Then
            ws.Move Before:=ThisWorkbook.Worksheets(targetIdx)
            targetIdx = targetIdx + 1
        End If
        Set ws = Nothing
    Next i

    On Error Resume Next
    ThisWorkbook.Worksheets("Preferences").Activate
    On Error GoTo 0
End Sub

' =============================================================================
' FORMULAS DOCUMENTATION SHEET -- unchanged from the AVL Production
' Dashboard; regenerated every Update Data. Slimmed 2026-08.12: the rate
' math documentation (and the live LAMBDA definitions section) moved to
' the AVL Production Dashboard along with the rate math itself.
' =============================================================================
Private Sub BuildFormulasSheet()
    Dim ws As Worksheet
    Set ws = GetOrCreateSheet("Formulas")
    If ws Is Nothing Then Exit Sub

    ws.Cells.Clear
    ws.Cells.UnMerge

    ws.Range("A1").Value = "HUB CLEANING PIPELINE (documentation)"
    With ws.Range("A1:C2")
        .Merge
        .Interior.Color = DashFillOrDefault("Main Header", RGB(31, 78, 120))
        .HorizontalAlignment = xlHAlignCenter
        .VerticalAlignment = xlVAlignCenter
    End With
    ApplySwatchFontTo ws.Range("A1:C2"), "Main Header", 14, True, RGB(255, 255, 255)

    Dim docs As Variant
    docs = Array( _
        Array("WHAT THE HUB DOES", "", ""), _
        Array("Source", "Monarch Report Folder (Preferences)", "Newest export in the shared folder is read on every Update Data click."), _
        Array("Cleaning", "noise rows dropped, columns mapped A:J", "Ship Date, Promised Date, Job ID, Customer, Description, QTY, Location Date, Location, Work Center, Week Start."), _
        Array("Week Start (col J)", "Monday of the Ship Date week", "Computed during the build; blank when a job has no ship date."), _
        Array("Status word", "REFRESHING / READY / ERROR", "Satellites check this on Preferences before pulling CleanedData."), _
        Array("", "", ""), _
        Array("WHERE THE RATE MATH WENT (2026-08.12)", "", ""), _
        Array("Effective QTY / Job Hours / Capacity", "AVL Production Dashboard", "Removed from the Hub at the user's direction. The Dashboard computes all three from ITS OWN Parameters tab after pulling CleanedData. Maintain machine rates THERE, not here -- the Hub no longer has a Parameters sheet."), _
        Array("", "", ""), _
        Array("RFID TYPE (column K, 2026-08.13)", "", ""), _
        Array("RFID Type", "longest matching token in Description", "Hub scans each Description for the tokens listed under RFID Technologies on Preferences and writes the match here; blank when nothing matches. Where tokens nest, the longest wins (Ving 1K beats 1K). Hub only CREATES this column -- charts, slicers and totals are built by the satellite workbooks."), _
        Array("", "", ""), _
        Array("TO PERSO (column L, 2026-09.03)", "", ""), _
        Array("To Perso", "copied as-is from Monarch's own 'To Perso' field", "Blank on exports that don't include the column (older Monarch layouts) or on jobs that haven't reached that stage yet. Hub only carries this column through -- no cleaning or interpretation applied.") _
        )

    Dim i As Long, r As Long
    r = 4
    For i = LBound(docs) To UBound(docs)
        ws.Cells(r, 1).Value = docs(i)(0)
        ws.Cells(r, 2).Value = docs(i)(1)
        ws.Cells(r, 3).Value = docs(i)(2)
        If docs(i)(1) = "" And docs(i)(0) <> "" Then
            With ws.Range(ws.Cells(r, 1), ws.Cells(r, 3))
                .Font.Bold = True
                .Interior.Color = DashFillOrDefault("Column Headers", RGB(221, 235, 247))
            End With
        End If
        r = r + 1
    Next i

    ws.Columns(1).ColumnWidth = 40
    ws.Columns(2).ColumnWidth = 42
    ws.Columns(3).ColumnWidth = 110
End Sub

' =============================================================================
' PREFERENCES SHEET -- Hub's copy holds only the Monarch folder path and
' the Hub Status pull-guard flag (no color swatches -- Hub has no dashboard
' to color; that machinery stays in the AVL Production Dashboard workbook).
' =============================================================================
Private Sub EnsurePreferencesSheet()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Preferences")
    On Error GoTo 0
    If Not ws Is Nothing Then
        EnsurePreferenceRows ws
        Exit Sub
    End If

    Dim prevActive As Object
    On Error Resume Next
    Set prevActive = ThisWorkbook.ActiveSheet
    On Error GoTo 0

    Set ws = GetOrCreateSheet("Preferences")
    If ws Is Nothing Then Exit Sub

    With ws.Cells(1, 1)
        .Value = "Preferences (Hub)"
        .Font.Bold = True
        .Font.Size = 14
    End With
    ws.Cells(2, 1).Value = "Point 'Monarch Report Folder' at the shared drop folder. Everything else here is status-only."

    ws.Columns(1).ColumnWidth = 34
    ws.Columns(2).ColumnWidth = 50

    EnsurePreferenceRows ws

    If Not prevActive Is Nothing Then
        On Error Resume Next
        prevActive.Activate
        On Error GoTo 0
    End If
End Sub

Private Sub EnsurePreferenceRows(ByRef ws As Worksheet)
    EnsureOnePrefRow ws, "Monarch Report Folder", "path", "", 0, 0
    EnsureOnePrefRow ws, "Hub Status", "status", "READY | not yet run", RGB(226, 239, 218), RGB(0, 97, 0)
    EnsureRfidTokenList ws
End Sub

Private Sub EnsureOnePrefRow(ByRef ws As Worksheet, ByVal lbl As String, ByVal kind As String, ByVal defaultText As String, ByVal fillC As Long, ByVal fontC As Long)
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
            Case "status"
                .NumberFormat = "@"
                .Value = defaultText
                .Interior.Color = fillC
                .Font.Color = fontC
                .HorizontalAlignment = xlHAlignLeft
        End Select
        .BorderAround Weight:=xlThin
    End With

    If lbl = "Monarch Report Folder" Then
        With ws.Cells(lastA + 1, 1)
            .Value = "Paste the shared folder path here (e.g. \\server\share\monarch). " & _
                     "Update Data will read the NEWEST .xls/.xlsx/.csv file in it. " & _
                     "There is no local paste-tab fallback in this workbook -- if the share fails, Update Data stops with an error."
            .Font.Size = 9
            .Font.Italic = True
        End With
    ElseIf lbl = "Hub Status" Then
        With ws.Cells(lastA + 1, 1)
            .Value = "Read by every satellite workbook's Update button before it pulls CleanedData/Parameters. " & _
                     "REFRESHING = a build is in progress, wait and retry. READY = safe to pull. ERROR = last build failed, check the detail after the pipe."
            .Font.Size = 9
            .Font.Italic = True
        End With
    End If
End Sub

' =============================================================================
' SHARED MONARCH DROP FOLDER -- unchanged.
' =============================================================================
Private Function LoadMonarchFromSharedFolder(ByRef rawData As Variant, ByRef srcDesc As String, ByRef autoFailed As Boolean) As Boolean
    LoadMonarchFromSharedFolder = False
    autoFailed = False

    Dim sw As Range
    Set sw = GetDashSwatch("Monarch Report Folder")
    If sw Is Nothing Then Exit Function

    Dim folder As String
    folder = Trim(CStr(sw.Value))
    If folder = "" Then Exit Function
    Do While Len(folder) > 1 And Right(folder, 1) = "\"
        folder = Left(folder, Len(folder) - 1)
    Loop

    Dim pats As Variant
    pats = Array("*.xls", "*.xlsx", "*.xlsm", "*.csv")
    Dim newestName As String, newestTime As Double
    Dim p As Variant, f As String, t As Double

    On Error Resume Next
    For Each p In pats
        f = Dir(folder & "\" & CStr(p))
        Do While f <> ""
            t = 0
            t = CDbl(FileDateTime(folder & "\" & f))
            If t > newestTime Then
                newestTime = t
                newestName = f
            End If
            f = Dir()
        Loop
    Next p
    On Error GoTo 0

    If newestName = "" Then
        autoFailed = True
        MsgBox "No Monarch report file found in:" & vbCrLf & folder & vbCrLf & vbCrLf & _
               "Check the path on the Preferences sheet, or that the report has been saved to the share.", _
               vbExclamation, "Monarch Report Folder"
        Exit Function
    End If

    Dim fullPath As String
    fullPath = folder & "\" & newestName

    Dim prevSU As Boolean
    prevSU = Application.ScreenUpdating
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Dim wbSrc As Workbook
    On Error Resume Next
    Set wbSrc = Workbooks.Open(Filename:=fullPath, ReadOnly:=True, UpdateLinks:=0)
    On Error GoTo 0

    If wbSrc Is Nothing Then
        autoFailed = True
        Application.DisplayAlerts = True
        Application.ScreenUpdating = prevSU
        MsgBox "Found '" & newestName & "' but could not open it (it may be mid-save or blocked).", vbExclamation, "Monarch Report Folder"
        Exit Function
    End If

    rawData = wbSrc.Worksheets(1).UsedRange.Value
    wbSrc.Close SaveChanges:=False
    Application.DisplayAlerts = True
    Application.ScreenUpdating = prevSU

    If IsEmpty(rawData) Then
        autoFailed = True
        MsgBox "'" & newestName & "' opened but contained no data.", vbExclamation, "Monarch Report Folder"
        Exit Function
    End If

    srcDesc = newestName & " (modified " & Format(newestTime, "yyyy-mm-dd hh:nn") & ")"
    LoadMonarchFromSharedFolder = True
End Function

' One-time sheet renames, carried over so a Hub built from an older base
' file still migrates correctly. No-op on a fresh Hub.
' 2026-08.12 one-time migration (idempotent, safe every run): the Hub no
' longer owns machine rate parameters -- the AVL Production Dashboard
' computes Effective QTY / Job Hours / Capacity from its OWN Parameters
' tab now, per the user's direction. Deletes the Parameters sheet (and
' its legacy "Settings" name), the SettingsBackup sheet, and the old
' rate-math LAMBDA names so nothing stale lingers in the file.
Private Sub RemoveParametersArtifacts()
    Dim prevAlerts As Boolean: prevAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    Dim nm As Variant, ws As Worksheet
    For Each nm In Array("Parameters", "Settings", "SettingsBackup")
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(CStr(nm))
        On Error GoTo 0
        If Not ws Is Nothing Then
            If ThisWorkbook.Sheets.Count > 1 Then ws.Delete
        End If
    Next nm
    Application.DisplayAlerts = prevAlerts

    On Error Resume Next
    For Each nm In Array("JobHours", "NumberOfDays", "NumberOfWeeks", "NumberOfShifts", _
                         "CapacityPercent", "AvailableHours", "DailyCapacity")
        ThisWorkbook.Names(CStr(nm)).Delete
    Next nm
    On Error GoTo 0
End Sub

' =============================================================================
' RFID TECHNOLOGY TAGGING (2026-08.13)
'
' Scans each job's Description for one of the technology tokens listed on
' Preferences and writes it into CleanedData's RFID Type column. That is the
' whole of Hub's involvement -- charts, slicers and totals are built by the
' satellite workbooks from this column.
'
' MATCHING RULE -- longest token wins. The list contains tokens that nest
' inside each other: '1K' sits inside 'Ving 1K', 'X1K', 'V1K' and '1KC';
' 'VU' sits inside 'VU128'. Matching in list order would tag a 'Ving 1K' job
' as plain '1K', so ReadRfidTokens sorts the list longest-first and the first
' hit wins -- the most specific token always beats the more general one.
' Comparison is case-insensitive and space-normalised on both sides, and it
' is a SUBSTRING test, not whole-word: real descriptions run tokens straight
' into neighbouring text (e.g. 'CHOPS-V2-VAESChoice Privileges').
' =============================================================================
Private Function ReadRfidTokens() As Variant
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Preferences")
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    Dim anchorRow As Long: anchorRow = 0
    Dim r As Long
    For r = 1 To 300
        If Trim(CStr(ws.Cells(r, 1).Value)) = "RFID Technologies" Then
            anchorRow = r
            Exit For
        End If
    Next r
    If anchorRow = 0 Then Exit Function

    Dim dataStart As Long: dataStart = anchorRow + 2   ' anchor, instructions, THEN data
    Dim disp() As String, canon() As String
    ReDim disp(1 To 300): ReDim canon(1 To 300)
    Dim n As Long: n = 0
    Dim raw As String
    r = dataStart
    Do While Trim(CStr(ws.Cells(r, 1).Value)) <> "" And r < dataStart + 300
        raw = Trim(CStr(ws.Cells(r, 1).Value))
        If raw <> "" Then
            n = n + 1
            disp(n) = raw
            canon(n) = CanonRfid(raw)
        End If
        r = r + 1
    Loop
    If n = 0 Then Exit Function

    ' Longest canonical token first -- see the header note on nesting.
    Dim i As Long, j As Long, td As String, tc As String
    For i = 2 To n
        td = disp(i): tc = canon(i)
        j = i - 1
        Do While j >= 1
            If Len(canon(j)) < Len(tc) Then
                disp(j + 1) = disp(j): canon(j + 1) = canon(j)
                j = j - 1
            Else
                Exit Do
            End If
        Loop
        disp(j + 1) = td: canon(j + 1) = tc
    Next i

    Dim outArr() As Variant
    ReDim outArr(1 To n, 1 To 2)
    For i = 1 To n
        outArr(i, 1) = disp(i)
        outArr(i, 2) = canon(i)
    Next i
    ReadRfidTokens = outArr
End Function

' Returns the DISPLAY form of the first (longest) token found in the
' description, or an empty string when nothing matches.
Private Function ClassifyRfidType(ByVal descr As String, ByRef tokens As Variant) As String
    ClassifyRfidType = ""
    If IsEmpty(tokens) Then Exit Function
    Dim d As String: d = CanonRfid(descr)
    If d = "" Then Exit Function
    Dim i As Long
    For i = LBound(tokens, 1) To UBound(tokens, 1)
        If CStr(tokens(i, 2)) <> "" Then
            If InStr(1, d, CStr(tokens(i, 2))) > 0 Then
                ClassifyRfidType = CStr(tokens(i, 1))
                Exit Function
            End If
        End If
    Next i
End Function

' Upper-cases and collapses runs of whitespace to a single space, so
' 'Ving  UL   48b' and 'ving ul 48b' both match the seeded 'Ving UL 48b'.
Private Function CanonRfid(ByVal s As String) As String
    Dim t As String
    t = UCase(Trim(s))
    t = Replace(t, vbTab, " ")
    t = Replace(t, Chr(160), " ")
    Do While InStr(t, "  ") > 0
        t = Replace(t, "  ", " ")
    Loop
    CanonRfid = t
End Function

' Seeds the editable technology list ONCE. Never rewritten afterwards, so
' rows added or removed here survive every future run -- the list is the
' single place these tokens are maintained for the whole workbook family.
Private Sub EnsureRfidTokenList(ByRef ws As Worksheet)
    Dim r As Long
    For r = 1 To 300
        If Trim(CStr(ws.Cells(r, 1).Value)) = "RFID Technologies" Then Exit Sub
    Next r

    Dim lastA As Long
    lastA = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 2
    With ws.Cells(lastA, 1)
        .Value = "RFID Technologies"
        .Font.Bold = True
    End With
    ws.Cells(lastA + 1, 1).Value = "One token per row below. Every job whose Description contains one of these gets it written to CleanedData's RFID Type column, which the satellite workbooks chart and total. Case and extra spaces do not matter, and a token may sit against other text. Where tokens overlap the LONGEST one wins, so a 'Ving 1K' job is tagged Ving 1K rather than 1K -- list order does not matter. Add or remove rows freely; no code change needed."
    ws.Cells(lastA + 1, 1).Font.Size = 9
    ws.Cells(lastA + 1, 1).Font.Italic = True

    Dim seedRow As Long: seedRow = lastA + 2
    Dim seeds As Variant
    seeds = Array( _
        "VAES", _
        "ULC", _
        "Ving 1K", _
        "Ving AES", _
        "X1K", _
        "VU128", _
        "Ving UL 48b", _
        "VU", _
        "1KC", _
        "MP2K", _
        "UL128", _
        "Ving UL 128b", _
        "V1K", _
        "Ving UL AES", _
        "VMG UL VEAS", _
        "1K", _
        "YMG-UL 128b")

    Dim i As Long
    For i = LBound(seeds) To UBound(seeds)
        ws.Cells(seedRow + i, 1).Value = CStr(seeds(i))
    Next i
    ws.Range(ws.Cells(seedRow, 1), ws.Cells(seedRow + UBound(seeds), 1)).NumberFormat = "@"
End Sub
Private Sub MigrateSheetNames()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = Nothing
    Set ws = ThisWorkbook.Sheets("Settings")
    If Not ws Is Nothing Then ws.name = "Parameters"
    Set ws = Nothing
    Set ws = ThisWorkbook.Sheets("Dashboard Colors")
    If Not ws Is Nothing Then ws.name = "Preferences"
    On Error GoTo 0
End Sub

Private Sub ApplySwatchFontTo(ByRef target As Range, ByVal elemName As String, _
                              ByVal defSize As Double, ByVal defBold As Boolean, ByVal defColor As Long)
    Dim sw As Range
    Set sw = GetDashSwatch(elemName)
    If sw Is Nothing Then
        With target.Font
            .Size = defSize
            .Bold = defBold
            .Color = defColor
        End With
        Exit Sub
    End If
    With target.Font
        .name = sw.Font.name
        .Size = sw.Font.Size
        .Bold = sw.Font.Bold
        .Italic = sw.Font.Italic
        .Color = sw.Font.Color
    End With
End Sub

' Returns the value cell (column B) for a named row on Preferences, or
' Nothing if the sheet/row doesn't exist.
Private Function GetDashSwatch(ByVal elemName As String) As Range
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Preferences")
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    Dim r As Long
    For r = 1 To 100
        If Trim(CStr(ws.Cells(r, 1).Value)) = elemName Then
            Set GetDashSwatch = ws.Cells(r, 2)
            Exit Function
        End If
    Next r
End Function

Private Function DashFillOrDefault(ByVal elemName As String, ByVal defaultColor As Long) As Long
    Dim sw As Range
    Set sw = GetDashSwatch(elemName)
    If sw Is Nothing Then
        DashFillOrDefault = defaultColor
    ElseIf sw.Interior.Pattern = xlNone Then
        DashFillOrDefault = defaultColor
    Else
        DashFillOrDefault = sw.Interior.Color
    End If
End Function

Private Function GetOrCreateSheet(ByVal rawName As String) As Worksheet
    Dim cleanName As String
    cleanName = Trim(rawName)

    Dim illegalChars As Variant, ch As Variant
    illegalChars = Array(":", "\", "/", "?", "*", "[", "]")
    For Each ch In illegalChars
        cleanName = Replace(cleanName, ch, "")
    Next ch

    If Len(cleanName) > 31 Then cleanName = Left(cleanName, 31)
    If cleanName = "" Then cleanName = "Unassigned"

    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets(cleanName)
    On Error GoTo 0

    If GetOrCreateSheet Is Nothing Then
        On Error GoTo CreateError
        If LCase(cleanName) = "parameters" Then
            Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
        Else
            Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add(After:=GetLastVisibleSheet())
        End If
        GetOrCreateSheet.name = cleanName
        On Error GoTo 0
    End If
    Exit Function

CreateError:
    MsgBox "Failed to create sheet tab: '" & cleanName & "'." & vbCrLf & _
           "Please check if Workbook Protection is enabled.", vbCritical, "Sheet Creation Error"
    Set GetOrCreateSheet = Nothing
End Function

Private Function GetLastVisibleSheet() As Worksheet
    Dim i As Long
    For i = ThisWorkbook.Worksheets.Count To 1 Step -1
        If ThisWorkbook.Worksheets(i).Visible = xlSheetVisible Then
            Set GetLastVisibleSheet = ThisWorkbook.Worksheets(i)
            Exit Function
        End If
    Next i
    Set GetLastVisibleSheet = ThisWorkbook.ActiveSheet
End Function

Private Sub QuickSortMultiKey(ByRef arr As Variant, ByVal low As Long, ByVal high As Long, ByVal k1 As Long, ByVal k2 As Long, ByVal k3 As Long, ByVal totalCols As Long)
    If low >= high Then Exit Sub

    Dim i As Long: i = low
    Dim j As Long: j = high
    Dim pivotIdx As Long: pivotIdx = (low + high) \ 2

    Dim pivotK1 As Variant: pivotK1 = arr(pivotIdx, k1)
    Dim pivotK2 As Variant: pivotK2 = arr(pivotIdx, k2)
    Dim pivotK3 As Variant: pivotK3 = arr(pivotIdx, k3)

    Do While i <= j
        Do While CompareKeys(arr(i, k1), arr(i, k2), arr(i, k3), pivotK1, pivotK2, pivotK3) < 0
            i = i + 1
        Loop

        Do While CompareKeys(arr(j, k1), arr(j, k2), arr(j, k3), pivotK1, pivotK2, pivotK3) > 0
            j = j - 1
        Loop

        If i <= j Then
            SwapRows arr, i, j, totalCols
            i = i + 1
            j = j - 1
        End If
    Loop

    If low < j Then QuickSortMultiKey arr, low, j, k1, k2, k3, totalCols
    If i < high Then QuickSortMultiKey arr, i, high, k1, k2, k3, totalCols
End Sub

Private Function CompareKeys(ByVal v1_a As Variant, ByVal v2_a As Variant, ByVal v3_a As Variant, ByVal v1_b As Variant, ByVal v2_b As Variant, ByVal v3_b As Variant) As Long
    Dim res As Long
    res = CompareSingleValue(v1_a, v1_b)
    If res <> 0 Then
        CompareKeys = res
        Exit Function
    End If

    res = CompareSingleValue(v2_a, v2_b)
    If res <> 0 Then
        CompareKeys = res
        Exit Function
    End If

    CompareKeys = CompareSingleValue(v3_a, v3_b)
End Function

Private Function CompareSingleValue(ByVal valA As Variant, ByVal valB As Variant) As Long
    If IsEmpty(valA) Or IsNull(valA) Then valA = ""
    If IsEmpty(valB) Or IsNull(valB) Then valB = ""

    If IsDate(valA) And IsDate(valB) Then
        If CDate(valA) < CDate(valB) Then
            CompareSingleValue = -1
        ElseIf CDate(valA) > CDate(valB) Then
            CompareSingleValue = 1
        Else
            CompareSingleValue = 0
        End If
    ElseIf IsNumeric(valA) And IsNumeric(valB) And valA <> "" And valB <> "" Then
        If CDbl(valA) < CDbl(valB) Then
            CompareSingleValue = -1
        ElseIf CDbl(valA) > CDbl(valB) Then
            CompareSingleValue = 1
        Else
            CompareSingleValue = 0
        End If
    Else
        Dim strA As String: strA = LCase(CStr(valA))
        Dim strB As String: strB = LCase(CStr(valB))
        If strA < strB Then
            CompareSingleValue = -1
        ElseIf strA > strB Then
            CompareSingleValue = 1
        Else
            CompareSingleValue = 0
        End If
    End If
End Function

Private Sub SwapRows(ByRef arr As Variant, ByVal r1 As Long, ByVal r2 As Long, ByVal totalCols As Long)
    Dim c As Long, temp As Variant
    For c = 1 To totalCols
        temp = arr(r1, c)
        arr(r1, c) = arr(r2, c)
        arr(r2, c) = temp
    Next c
End Sub

Private Sub WriteAndFormatSheet(ByRef ws As Worksheet, ByRef arr As Variant, ByVal rowsCount As Long, ByVal colsCount As Long)
    ws.Range("A1").Resize(rowsCount, colsCount).Value = arr

    Dim dateFormatStr As String: dateFormatStr = "mm-dd-yyyy"

    ws.Columns(1).NumberFormat = dateFormatStr
    ws.Columns(2).NumberFormat = dateFormatStr
    ws.Columns(6).NumberFormat = "#,##0"
    ws.Columns(7).NumberFormat = dateFormatStr
    ws.Columns(9).NumberFormat = "@"
    ws.Columns(10).NumberFormat = dateFormatStr
    ws.Columns(11).NumberFormat = "@"   ' RFID Type is always text

    AddJobHoursTable ws, rowsCount, colsCount

    ws.Columns.AutoFit
End Sub

Private Sub AddJobHoursTable(ByRef ws As Worksheet, ByVal rowsCount As Long, ByVal colsCount As Long)
    If rowsCount < 1 Then Exit Sub

    On Error Resume Next
    Dim lo As ListObject
    For Each lo In ws.ListObjects
        lo.Unlist
    Next lo
    On Error GoTo 0

    Dim dataRange As Range
    Set dataRange = ws.Range(ws.Cells(1, 1), ws.Cells(rowsCount, colsCount))

    Dim tblName As String
    tblName = "tbl_" & ws.name
    Dim illegalChars As Variant, ic As Variant
    illegalChars = Array(" ", "-", "&", "(", ")", ".", "'")
    For Each ic In illegalChars
        tblName = Replace(tblName, ic, "_")
    Next ic

    On Error Resume Next
    ThisWorkbook.Names(tblName).Delete
    On Error GoTo 0

    Dim tbl As ListObject
    Set tbl = ws.ListObjects.Add(xlSrcRange, dataRange, , xlYes)
    tbl.name = tblName

    ' 2026-08.12: the three computed columns (Effective QTY / Job Hours /
    ' Capacity Load) are GONE -- rate math moved to the AVL Production
    ' Dashboard's own Parameters tab at the user's direction ("Remove
    ' parameters from HUB and have the AVL Dashboard base its calculations
    ' on the parameters tab of itself"). colsCount is passed in from
    ' ProcessAndDistributeAllData (totalOutCols) rather than assumed here,
    ' so this sub always wraps whatever columns CleanedData actually has
    ' (A:L as of the 2026-09.03 To Perso addition) in a filterable table.
End Sub
