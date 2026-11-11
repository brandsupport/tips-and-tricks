Attribute VB_Name = "modVWGExcelTracker"
Option Explicit

'===========================================================
' VWG DISPATCH TRACKER - EXCEL CONNECTOR V5.1
' Late-bound Excel automation. No extra VBA reference required.
' Workbook is the source of truth for assignments and activity.
'===========================================================

Private Const STATE_FOLDER As String = "VWGDispatcher_CLEAN_V5"
Private Const DEFAULT_WORKBOOK_NAME As String = "VWG_Dispatch_Tracker_V5_NoCF_ShiftTimes_Fixed.xlsx"
Private Const CORRECTED_WORKBOOK_PREFIX As String = "VWG_Dispatch_Tracker_V5_NoCF_ShiftTimes_Fixed"
Private Const XL_UP As Long = -4162
Private Const XL_CSV_UTF8 As Long = 62
Private Const HEADER_ROW As Long = 5
Private Const FIRST_DATA_ROW As Long = 6
Private Const VWGX_BUILD_ID As String = "VWG-TRACKER-2026.11.11-R1"
Private Const VWGX_RETRYABLE_EXCEL_ERROR As Long = 50290
Private Const VWGX_RPC_E_CALL_REJECTED As Long = -2147418111

#If VBA7 Then
Private Declare PtrSafe Sub VWGX_Sleep Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)
#Else
Private Declare Sub VWGX_Sleep Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)
#End If

Private gXLApp As Object
Private gXLWorkbook As Object
Private gWorkbookPath As String
Private gConnectionBusy As Boolean
Private gXLBatchDepth As Long
Private gXLWorkbookDirty As Boolean
Private gTrackerWriteBusy As Boolean

'===========================================================
' PUBLIC: CONNECTION TEST - READ ONLY
'===========================================================

Public Sub VWGX_ConnectTest()
    On Error GoTo ConnectFailed

    If Not VWGX_EnsureConnected(True) Then Exit Sub

    MsgBox "Excel tracker connected and structure validated." & vbCrLf & vbCrLf & _
           "Workbook: " & CStr(gXLWorkbook.Name) & vbCrLf & _
           "Path: " & CStr(gXLWorkbook.FullName) & vbCrLf & vbCrLf & _
           "Sheets: Dashboard, Assignments, Activity Log, Team & Shifts, Settings" & vbCrLf & _
           "No assignment or activity rows were changed.", _
           vbInformation, "VWG Tracker Connection Test"
    Exit Sub

ConnectFailed:
    MsgBox "Tracker connection failed: " & Err.Number & " - " & Err.Description, _
           vbExclamation, "VWG Tracker Connection Test"
End Sub

'===========================================================
' PUBLIC: FORCE THE USER TO SELECT THE CORRECTED WORKBOOK
' This intentionally bypasses both the in-memory workbook reference
' and the previously saved tracker-path.txt value.
'===========================================================
Public Sub VWGX_SelectCorrectWorkbook()
    Dim picker As Object
    Dim book As Object
    Dim selectedBook As Object
    Dim selectedPath As String
    Dim selectedFile As String
    Dim openedByThisProcedure As Boolean
    Dim failureNumber As Long
    Dim failureText As String

    If gConnectionBusy Then
        MsgBox "The tracker is currently connecting. Try again once that operation has finished.", _
               vbExclamation, "VWG Tracker"
        Exit Sub
    End If

    gConnectionBusy = True
    On Error GoTo SelectFailed

    If gXLApp Is Nothing Then
        On Error Resume Next
        Set gXLApp = GetObject(, "Excel.Application")
        On Error GoTo SelectFailed
        If gXLApp Is Nothing Then Set gXLApp = CreateObject("Excel.Application")
    End If

    gXLApp.Visible = True
    Set picker = gXLApp.FileDialog(3) ' msoFileDialogFilePicker

    With picker
        .Title = "Select the corrected VWG Dispatch Tracker workbook"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Excel workbook (*.xlsx)", "*.xlsx"
        If .Show <> -1 Then GoTo SelectCancelled
        selectedPath = CStr(.SelectedItems(1))
    End With

    selectedFile = Dir$(selectedPath)
    If StrComp(Left$(selectedFile, Len(CORRECTED_WORKBOOK_PREFIX)), _
               CORRECTED_WORKBOOK_PREFIX, vbTextCompare) <> 0 Then
        MsgBox "Please select the corrected workbook whose filename begins with:" & vbCrLf & _
               CORRECTED_WORKBOOK_PREFIX, vbExclamation, "VWG Tracker"
        GoTo SelectCancelled
    End If

    ' Reuse the selected workbook if it is already open; otherwise open it.
    For Each book In gXLApp.Workbooks
        If StrComp(CStr(book.FullName), selectedPath, vbTextCompare) = 0 Then
            Set selectedBook = book
            Exit For
        End If
    Next book

    If selectedBook Is Nothing Then
        Set selectedBook = gXLApp.Workbooks.Open(selectedPath)
        openedByThisProcedure = True
    End If

    If Not ValidateWorkbook(selectedBook) Then
        If openedByThisProcedure Then selectedBook.Close False
        Set selectedBook = Nothing
        MsgBox "The selected file does not have the expected V5 tracker structure." & vbCrLf & _
               "No tracker path was changed. Select the corrected V5 workbook.", _
               vbExclamation, "VWG Tracker"
        GoTo SelectCancelled
    End If

    ' Persist the new path first. Only then replace the in-memory workbook reference.
    selectedPath = CStr(selectedBook.FullName)
    SaveWorkbookPath selectedPath
    Set gXLWorkbook = selectedBook
    gWorkbookPath = selectedPath
    gXLWorkbookDirty = False
    selectedBook.Activate

    MsgBox "Corrected tracker selected and saved successfully." & vbCrLf & vbCrLf & _
           "Workbook: " & CStr(gXLWorkbook.Name) & vbCrLf & _
           "Path: " & gWorkbookPath & vbCrLf & vbCrLf & _
           "Your existing dispatcher history folder was not changed.", _
           vbInformation, "VWG Tracker"

SelectCancelled:
    gConnectionBusy = False
    Exit Sub

SelectFailed:
    failureNumber = Err.Number
    failureText = Err.Description
    Debug.Print "[VWG TRACKER] Workbook selection failed: " & failureNumber & " - " & failureText
    If openedByThisProcedure Then
        On Error Resume Next
        If Not selectedBook Is Nothing Then selectedBook.Close False
        On Error GoTo 0
    End If
    gConnectionBusy = False
    MsgBox "Could not select the corrected tracker workbook." & vbCrLf & vbCrLf & _
           failureNumber & " - " & failureText, vbExclamation, "VWG Tracker"
End Sub

Public Function VWGX_EnsureConnected(Optional ByVal allowPicker As Boolean = False) As Boolean
    Dim pathText As String
    Dim book As Object
    Dim picker As Object

    If gConnectionBusy Then Exit Function
    gConnectionBusy = True
    On Error GoTo Failed

    If Not gXLWorkbook Is Nothing Then
        If WorkbookIsUsable(gXLWorkbook) Then
            VWGX_EnsureConnected = True
            GoTo CleanExit
        End If
        Set gXLWorkbook = Nothing
    End If

    If gXLApp Is Nothing Then
        On Error Resume Next
        Set gXLApp = GetObject(, "Excel.Application")
        On Error GoTo Failed
        If gXLApp Is Nothing Then Set gXLApp = CreateObject("Excel.Application")
    End If

    gXLApp.Visible = True
    pathText = ReadSavedWorkbookPath()

    If Len(pathText) = 0 Or Len(Dir$(pathText)) = 0 Then
        pathText = Environ$("USERPROFILE") & "\Documents\" & DEFAULT_WORKBOOK_NAME
    End If

    If Len(Dir$(pathText)) = 0 Then
        If Not allowPicker Then GoTo CleanExit
        Set picker = gXLApp.FileDialog(3) ' msoFileDialogFilePicker
        With picker
            .Title = "Select the VWG Dispatch Tracker V5 workbook"
            .AllowMultiSelect = False
            .Filters.Clear
            .Filters.Add "Excel workbook", "*.xlsx"
            If .Show <> -1 Then GoTo CleanExit
            pathText = CStr(.SelectedItems(1))
        End With
    End If

    For Each book In gXLApp.Workbooks
        If StrComp(CStr(book.FullName), pathText, vbTextCompare) = 0 Then
            Set gXLWorkbook = book
            Exit For
        End If
    Next book

    If gXLWorkbook Is Nothing Then
        Set gXLWorkbook = gXLApp.Workbooks.Open(pathText)
    End If

    If Not ValidateWorkbook(gXLWorkbook) Then
        Err.Raise vbObjectError + 1501, "VWGX_EnsureConnected", _
                  "Workbook layout does not match V5. Select VWG_Dispatch_Tracker_V5_NoCF_ShiftTimes_Fixed.xlsx."
    End If

    gWorkbookPath = CStr(gXLWorkbook.FullName)
    SaveWorkbookPath gWorkbookPath
    VWGX_EnsureConnected = True

CleanExit:
    gConnectionBusy = False
    Exit Function

Failed:
    Debug.Print "[VWG TRACKER] Connection error " & Err.Number & ": " & Err.Description
    Set gXLWorkbook = Nothing
    VWGX_EnsureConnected = False
    gConnectionBusy = False
    If allowPicker Then
        MsgBox "Could not connect to the tracker workbook." & vbCrLf & vbCrLf & _
               Err.Description, vbExclamation, "VWG Tracker"
    End If
End Function

Public Sub VWGX_SaveWorkbook()
    On Error GoTo SaveFailed
    If VWGX_EnsureConnected(False) Then SaveWorkbookNow
    Exit Sub
SaveFailed:
    Debug.Print "[VWG TRACKER] Save failed: " & Err.Number & " - " & Err.Description
End Sub

' Batch operations are used during large baseline scans so each existing
' tagged email does not cause several full workbook saves.
Public Function VWGX_BeginBatch() As Boolean
    If Not VWGX_EnsureConnected(False) Then Exit Function
    gXLBatchDepth = gXLBatchDepth + 1
    VWGX_BeginBatch = True
End Function

Public Function VWGX_EndBatch(Optional ByVal forceSave As Boolean = True) As Boolean
    On Error GoTo EndFailed
    If gXLBatchDepth > 0 Then gXLBatchDepth = gXLBatchDepth - 1
    If gXLBatchDepth = 0 Then
        If gXLWorkbookDirty Or forceSave Then SaveWorkbookNow
    End If
    VWGX_EndBatch = True
    Exit Function
EndFailed:
    Debug.Print "[VWG TRACKER] Batch save failed: " & Err.Number & " - " & Err.Description
    VWGX_EndBatch = False
End Function

Private Sub SaveWorkbookIfNeeded()
    If gXLBatchDepth > 0 Then
        gXLWorkbookDirty = True
    Else
        SaveWorkbookNow
    End If
End Sub

Private Sub SaveWorkbookNow()
    Dim attempt As Long
    Dim savedErr As Long
    Dim savedDescription As String

    If gXLWorkbook Is Nothing Then Err.Raise vbObjectError + 1510, "SaveWorkbookNow", "Workbook is not connected."

    For attempt = 1 To 3
        savedErr = 0
        savedDescription = vbNullString
        On Error Resume Next
        gXLWorkbook.Save
        savedErr = Err.Number
        savedDescription = Err.Description
        Err.Clear
        On Error GoTo 0

        If savedErr = 0 Then
            gXLWorkbookDirty = False
            Exit Sub
        End If

        If (savedErr <> VWGX_RETRYABLE_EXCEL_ERROR And _
            savedErr <> VWGX_RPC_E_CALL_REJECTED) Or attempt = 3 Then
            Err.Raise savedErr, "SaveWorkbookNow", savedDescription
        End If

        Debug.Print "[VWG TRACKER] Retrying workbook Save after transient Excel automation error " & _
                    savedErr & " (attempt " & CStr(attempt) & " of 3)."
        VWGX_Sleep 150
    Next attempt
End Sub

'===========================================================
' PUBLIC: WORKBOOK SETTINGS
'===========================================================

Public Function VWGX_GetSetting(ByVal settingName As String) As Variant
    Dim ws As Object
    Dim lastRow As Long
    Dim r As Long

    On Error GoTo ReadFailed
    If Not VWGX_EnsureConnected(False) Then Exit Function
    Set ws = gXLWorkbook.Worksheets("Settings")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(XL_UP).Row

    For r = FIRST_DATA_ROW To lastRow
        If StrComp(Trim$(CStr(ws.Cells(r, 1).Value)), settingName, vbTextCompare) = 0 Then
            VWGX_GetSetting = ws.Cells(r, 2).Value
            Exit Function
        End If
    Next r
    Exit Function

ReadFailed:
    Debug.Print "[VWG TRACKER] Setting read failed: " & settingName & " | " & Err.Description
End Function

Public Function VWGX_SetSetting(ByVal settingName As String, ByVal settingValue As Variant) As Boolean
    Dim ws As Object
    Dim lastRow As Long
    Dim r As Long
    Dim targetRow As Long

    On Error GoTo WriteFailed
    If Not VWGX_EnsureConnected(False) Then Exit Function
    Set ws = gXLWorkbook.Worksheets("Settings")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(XL_UP).Row
    targetRow = 0

    For r = FIRST_DATA_ROW To lastRow
        If StrComp(Trim$(CStr(ws.Cells(r, 1).Value)), settingName, vbTextCompare) = 0 Then
            targetRow = r
            Exit For
        End If
    Next r

    If targetRow = 0 Then targetRow = lastRow + 1
    ws.Cells(targetRow, 1).Value = settingName
    ws.Cells(targetRow, 2).Value = settingValue
    SaveWorkbookIfNeeded
    VWGX_SetSetting = True
    Exit Function

WriteFailed:
    Debug.Print "[VWG TRACKER] Setting write failed: " & settingName & " | " & Err.Description
    VWGX_SetSetting = False
End Function

'===========================================================
' PUBLIC: ASSIGNMENT ROWS
' The passed record is a zero-based Variant array of 21 fields.
' It maps to Assignments columns A:U.
'===========================================================

Public Function VWGX_LoadAssignments() As Variant
    Dim ws As Object
    Dim lastRow As Long
    On Error GoTo ReadFailed

    If Not VWGX_EnsureConnected(False) Then Exit Function
    Set ws = gXLWorkbook.Worksheets("Assignments")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(XL_UP).Row
    If lastRow < FIRST_DATA_ROW Then Exit Function
    VWGX_LoadAssignments = ws.Range("A" & FIRST_DATA_ROW & ":U" & lastRow).Value
    Exit Function

ReadFailed:
    Debug.Print "[VWG TRACKER] Assignment load failed: " & Err.Description
End Function

Public Function VWGX_UpsertAssignment(ByVal record As Variant) As Boolean
    Dim ws As Object
    Dim lastRow As Long
    Dim targetRow As Long
    Dim r As Long
    Dim keyText As String
    Dim i As Long
    Dim col As Long
    Dim headerText As String
    Dim stageText As String
    Dim errNumber As Long
    Dim errText As String
    Dim writeStarted As Boolean
    Dim isNewRow As Boolean
    Dim duplicateCount As Long
    Dim oldValues As Variant
    Dim rollbackOK As Boolean
    Dim cellValue As Variant
    Dim rollbackErr As Long
    Dim rollbackDescription As String

    On Error GoTo WriteFailed

    stageText = "connect to workbook"
    If Not VWGX_EnsureConnected(False) Then Exit Function
    If gTrackerWriteBusy Then
        Debug.Print "[VWG TRACKER] Assignment write skipped: another tracker write is active."
        Exit Function
    End If
    gTrackerWriteBusy = True

    stageText = "validate assignment record"
    If Not IsArray(record) Then Err.Raise vbObjectError + 1601, "VWGX_UpsertAssignment", "Assignment record is not an array."
    If LBound(record) <> 0 Or UBound(record) < 20 Then
        Err.Raise vbObjectError + 1602, "VWGX_UpsertAssignment", "Assignment record must contain zero-based fields 0 to 20."
    End If

    stageText = "read assignment key"
    If IsNull(record(0)) Or IsEmpty(record(0)) Then
        Err.Raise vbObjectError + 1603, "VWGX_UpsertAssignment", "Assignment Email Key is blank."
    End If
    keyText = Trim$(CStr(record(0)))
    If Len(keyText) = 0 Then Err.Raise vbObjectError + 1604, "VWGX_UpsertAssignment", "Assignment Email Key is blank."

    stageText = "get Assignments worksheet"
    Set ws = gXLWorkbook.Worksheets("Assignments")

    stageText = "find last assignment row"
    lastRow = ws.Cells(ws.Rows.Count, 1).End(XL_UP).Row
    If lastRow < HEADER_ROW Then lastRow = HEADER_ROW
    targetRow = 0
    duplicateCount = 0

    stageText = "find existing assignment key"
    For r = FIRST_DATA_ROW To lastRow
        If StrComp(Trim$(CStr(ws.Cells(r, 1).Value2)), keyText, vbTextCompare) = 0 Then
            If targetRow = 0 Then
                targetRow = r
            Else
                duplicateCount = duplicateCount + 1
            End If
        End If
    Next r

    If duplicateCount > 0 Then
        Debug.Print "[VWG TRACKER] WARNING: Duplicate Email Key found. Key=" & keyText & _
                    " | additional matching rows=" & duplicateCount & _
                    " | updating first matching row=" & targetRow & ". No rows deleted."
    End If

    isNewRow = (targetRow = 0)
    If isNewRow Then
        targetRow = lastRow + 1
    Else
        stageText = "snapshot existing assignment row " & CStr(targetRow)
        oldValues = ws.Range("A" & targetRow & ":U" & targetRow).Value2
    End If

    ' The key is established first so a retry can locate the same row.
    ' Status (column J) is written last, preventing a new partial record from
    ' being counted as active before all its other fields have been stored.
    stageText = "write Email Key / ID (column A) row " & CStr(targetRow)
    writeStarted = True
    cellValue = keyText
    If Not VWGX_WriteCellValueSafe(ws.Cells(targetRow, 1), cellValue, stageText, errNumber, errText) Then
        Err.Raise errNumber, "VWGX_UpsertAssignment", errText
    End If
    gXLWorkbookDirty = True

    For i = 1 To 20
        If i <> 9 Then
            col = i + 1
            headerText = Trim$(CStr(ws.Cells(HEADER_ROW, col).Value2))
            If Len(headerText) = 0 Then headerText = "column " & CStr(col)
            stageText = "write " & headerText & " (column " & CStr(col) & ") row " & CStr(targetRow)

            cellValue = record(i)
            If IsNull(cellValue) Or IsEmpty(cellValue) Then cellValue = Empty
            If Not VWGX_WriteCellValueSafe(ws.Cells(targetRow, col), cellValue, stageText, errNumber, errText) Then
                Err.Raise errNumber, "VWGX_UpsertAssignment", errText
            End If
            gXLWorkbookDirty = True
        End If
    Next i

    ' Formatting is cosmetic. Apply it after the required values and never let
    ' a NumberFormat failure cancel a valid assignment reservation or update.
    VWGX_TryNumberFormat ws.Cells(targetRow, 1), "@", "Assignments A" & targetRow
    VWGX_TryNumberFormat ws.Cells(targetRow, 2), "dd/mmm/yyyy hh:mm", "Assignments B" & targetRow
    VWGX_TryNumberFormat ws.Cells(targetRow, 3), "@", "Assignments C" & targetRow
    VWGX_TryNumberFormat ws.Cells(targetRow, 4), "@", "Assignments D" & targetRow
    VWGX_TryNumberFormat ws.Cells(targetRow, 5), "@", "Assignments E" & targetRow
    VWGX_TryNumberFormat ws.Cells(targetRow, 6), "@", "Assignments F" & targetRow
    VWGX_TryNumberFormat ws.Cells(targetRow, 7), "@", "Assignments G" & targetRow
    VWGX_TryNumberFormat ws.Cells(targetRow, 9), "dd/mmm/yyyy hh:mm", "Assignments I" & targetRow
    VWGX_TryNumberFormat ws.Cells(targetRow, 10), "@", "Assignments J" & targetRow
    VWGX_TryNumberFormat ws.Cells(targetRow, 11), "@", "Assignments K" & targetRow
    VWGX_TryNumberFormat ws.Cells(targetRow, 12), "dd/mmm/yyyy hh:mm", "Assignments L" & targetRow
    VWGX_TryNumberFormat ws.Cells(targetRow, 13), "@", "Assignments M" & targetRow
    VWGX_TryNumberFormat ws.Cells(targetRow, 14), "dd/mmm/yyyy hh:mm", "Assignments N" & targetRow
    VWGX_TryNumberFormat ws.Cells(targetRow, 15), "dd/mmm/yyyy hh:mm", "Assignments O" & targetRow
    VWGX_TryNumberFormat ws.Cells(targetRow, 16), "dd/mmm/yyyy", "Assignments P" & targetRow
    VWGX_TryNumberFormat ws.Cells(targetRow, 17), "@", "Assignments Q" & targetRow
    VWGX_TryNumberFormat ws.Cells(targetRow, 18), "@", "Assignments R" & targetRow
    VWGX_TryNumberFormat ws.Cells(targetRow, 19), "dd/mmm/yyyy hh:mm", "Assignments S" & targetRow
    VWGX_TryNumberFormat ws.Cells(targetRow, 20), "@", "Assignments T" & targetRow
    VWGX_TryNumberFormat ws.Cells(targetRow, 21), "0", "Assignments U" & targetRow

    stageText = "write Status (column J) row " & CStr(targetRow)
    cellValue = record(9)
    If IsNull(cellValue) Or IsEmpty(cellValue) Then cellValue = Empty
    If Not VWGX_WriteCellValueSafe(ws.Cells(targetRow, 10), cellValue, stageText, errNumber, errText) Then
        Err.Raise errNumber, "VWGX_UpsertAssignment", errText
    End If
    gXLWorkbookDirty = True

    stageText = "verify assignment key and status row " & CStr(targetRow)
    If StrComp(Trim$(CStr(ws.Cells(targetRow, 1).Value2)), keyText, vbTextCompare) <> 0 Then
        Err.Raise vbObjectError + 1605, "VWGX_UpsertAssignment", "Email Key read-back did not match."
    End If
    If Not VWGX_ValuesMatch(ws.Cells(targetRow, 10).Value2, record(9)) Then
        Err.Raise vbObjectError + 1606, "VWGX_UpsertAssignment", "Status read-back did not match."
    End If

    stageText = "save tracker workbook"
    SaveWorkbookIfNeeded

    VWGX_UpsertAssignment = True
    gTrackerWriteBusy = False
    Exit Function

WriteFailed:
    errNumber = Err.Number
    errText = Err.Description
    Debug.Print "[VWG TRACKER] Assignment write failed at [" & stageText & "]: " & _
                errNumber & " - " & errText

    ' Restore only the affected row if a required write/save failed.
    If writeStarted And targetRow >= FIRST_DATA_ROW Then
        rollbackOK = True
        On Error Resume Next
        If isNewRow Then
            For col = 1 To 21
                rollbackErr = 0
                rollbackDescription = vbNullString
                If Not VWGX_WriteCellValueSafe(ws.Cells(targetRow, col), Empty, _
                    "rollback new Assignments row " & CStr(targetRow) & " column " & CStr(col), _
                    rollbackErr, rollbackDescription) Then rollbackOK = False
            Next col
        ElseIf IsArray(oldValues) Then
            For col = 1 To 21
                rollbackErr = 0
                rollbackDescription = vbNullString
                cellValue = oldValues(1, col)
                If IsNull(cellValue) Or IsEmpty(cellValue) Then cellValue = Empty
                If Not VWGX_WriteCellValueSafe(ws.Cells(targetRow, col), cellValue, _
                    "rollback existing Assignments row " & CStr(targetRow) & " column " & CStr(col), _
                    rollbackErr, rollbackDescription) Then rollbackOK = False
            Next col
        Else
            rollbackOK = False
        End If
        On Error GoTo 0

        If rollbackOK Then
            On Error Resume Next
            gXLWorkbookDirty = True
            SaveWorkbookIfNeeded
            If Err.Number <> 0 Then
                Debug.Print "[VWG TRACKER] Assignment rollback save failed: " & Err.Number & " - " & Err.Description
                Err.Clear
            Else
                Debug.Print "[VWG TRACKER] Assignment row rollback completed for row " & CStr(targetRow) & "."
            End If
            On Error GoTo 0
        Else
            Debug.Print "[VWG TRACKER] CRITICAL: Assignment row rollback was incomplete at row " & _
                        CStr(targetRow) & ". Keep dispatcher stopped and inspect this row before restart."
        End If
    End If

    gTrackerWriteBusy = False
    VWGX_UpsertAssignment = False
End Function

'===========================================================
' PUBLIC: APPEND ACTIVITY
' A zero-based 9-field Variant array maps to Activity Log A:I.
'===========================================================

Public Function VWGX_AppendActivity(ByVal eventRecord As Variant) As Boolean
    Dim ws As Object
    Dim lastRow As Long
    Dim targetRow As Long
    Dim r As Long
    Dim i As Long
    Dim col As Long
    Dim headerText As String
    Dim stageText As String
    Dim errNumber As Long
    Dim errText As String
    Dim writeStarted As Boolean
    Dim candidate As Boolean
    Dim allMatch As Boolean
    Dim existingValue As Variant
    Dim cellValue As Variant
    Dim rollbackErr As Long
    Dim rollbackDescription As String

    On Error GoTo WriteFailed

    stageText = "connect to workbook"
    If Not VWGX_EnsureConnected(False) Then Exit Function
    If gTrackerWriteBusy Then
        Debug.Print "[VWG TRACKER] Activity write skipped: another tracker write is active."
        Exit Function
    End If
    gTrackerWriteBusy = True

    stageText = "validate activity record"
    If Not IsArray(eventRecord) Then Err.Raise vbObjectError + 1611, "VWGX_AppendActivity", "Activity record is not an array."
    If LBound(eventRecord) <> 0 Or UBound(eventRecord) < 8 Then
        Err.Raise vbObjectError + 1612, "VWGX_AppendActivity", "Activity record must contain zero-based fields 0 to 8."
    End If

    stageText = "get Activity Log worksheet"
    Set ws = gXLWorkbook.Worksheets("Activity Log")

    stageText = "find last activity row"
    lastRow = ws.Cells(ws.Rows.Count, 1).End(XL_UP).Row
    If lastRow < HEADER_ROW Then lastRow = HEADER_ROW
    targetRow = 0

    ' Prevent a repeated call from duplicating an already completed event.
    ' Activity Log remains append-only: differing details are never overwritten.
    stageText = "check for an exact duplicate activity event"
    For r = FIRST_DATA_ROW To lastRow
        If VWGX_ValuesMatch(ws.Cells(r, 1).Value2, eventRecord(0)) And _
           VWGX_ValuesMatch(ws.Cells(r, 2).Value2, eventRecord(1)) And _
           VWGX_ValuesMatch(ws.Cells(r, 3).Value2, eventRecord(2)) Then

            allMatch = True
            For i = 3 To 8
                If Not VWGX_ValuesMatch(ws.Cells(r, i + 1).Value2, eventRecord(i)) Then
                    allMatch = False
                    Exit For
                End If
            Next i

            If allMatch Then
                stageText = "save previously recorded activity event"
                SaveWorkbookIfNeeded
                VWGX_AppendActivity = True
                gTrackerWriteBusy = False
                Exit Function
            End If
        End If

        ' An interrupted row has no timestamp in column A, so it will normally
        ' be the next target row. This branch also reuses a partial row if A was
        ' written but the final save was interrupted and its Email Key is blank.
        If targetRow = 0 Then
            If VWGX_ValuesMatch(ws.Cells(r, 1).Value2, eventRecord(0)) And _
               VWGX_ValuesMatch(ws.Cells(r, 2).Value2, eventRecord(1)) Then
                existingValue = ws.Cells(r, 3).Value2
                If IsEmpty(existingValue) Or Len(Trim$(CStr(existingValue))) = 0 Then
                    candidate = True
                    For i = 3 To 8
                        existingValue = ws.Cells(r, i + 1).Value2
                        If Not IsEmpty(existingValue) And Len(Trim$(CStr(existingValue))) > 0 Then
                            If Not VWGX_ValuesMatch(existingValue, eventRecord(i)) Then
                                candidate = False
                                Exit For
                            End If
                        End If
                    Next i
                    If candidate Then targetRow = r
                End If
            End If
        End If
    Next r

    If targetRow = 0 Then targetRow = lastRow + 1

    ' Write B:I first. Timestamp in A is the row's commit marker and is written
    ' last, so an interrupted partial write does not advance the log's last row.
    For i = 1 To 8
        col = i + 1
        headerText = Trim$(CStr(ws.Cells(HEADER_ROW, col).Value2))
        If Len(headerText) = 0 Then headerText = "column " & CStr(col)
        stageText = "write " & headerText & " (column " & CStr(col) & ") row " & CStr(targetRow)

        cellValue = eventRecord(i)
        If IsNull(cellValue) Or IsEmpty(cellValue) Then cellValue = Empty
        If Not VWGX_WriteCellValueSafe(ws.Cells(targetRow, col), cellValue, stageText, errNumber, errText) Then
            Err.Raise errNumber, "VWGX_AppendActivity", errText
        End If
        writeStarted = True
        gXLWorkbookDirty = True
    Next i

    ' Apply display formats after B:I have been written; formatting is non-fatal.
    VWGX_TryNumberFormat ws.Cells(targetRow, 2), "@", "Activity Log B" & targetRow
    VWGX_TryNumberFormat ws.Cells(targetRow, 3), "@", "Activity Log C" & targetRow
    VWGX_TryNumberFormat ws.Cells(targetRow, 4), "@", "Activity Log D" & targetRow
    VWGX_TryNumberFormat ws.Cells(targetRow, 5), "@", "Activity Log E" & targetRow
    VWGX_TryNumberFormat ws.Cells(targetRow, 6), "@", "Activity Log F" & targetRow
    VWGX_TryNumberFormat ws.Cells(targetRow, 7), "@", "Activity Log G" & targetRow
    VWGX_TryNumberFormat ws.Cells(targetRow, 8), "0", "Activity Log H" & targetRow
    VWGX_TryNumberFormat ws.Cells(targetRow, 9), "@", "Activity Log I" & targetRow
    VWGX_TryNumberFormat ws.Cells(targetRow, 1), "dd/mmm/yyyy hh:mm:ss", "Activity Log A" & targetRow

    stageText = "write Event Time completion marker (column A) row " & CStr(targetRow)
    cellValue = eventRecord(0)
    If IsNull(cellValue) Or IsEmpty(cellValue) Then cellValue = Empty
    If Not VWGX_WriteCellValueSafe(ws.Cells(targetRow, 1), cellValue, stageText, errNumber, errText) Then
        Err.Raise errNumber, "VWGX_AppendActivity", errText
    End If
    gXLWorkbookDirty = True

    stageText = "verify Activity Log row " & CStr(targetRow)
    For i = 0 To 8
        If Not VWGX_ValuesMatch(ws.Cells(targetRow, i + 1).Value2, eventRecord(i)) Then
            Err.Raise vbObjectError + 1613, "VWGX_AppendActivity", _
                      "Activity Log read-back mismatch in column " & CStr(i + 1) & "."
        End If
    Next i

    stageText = "save Activity Log"
    SaveWorkbookIfNeeded

    VWGX_AppendActivity = True
    gTrackerWriteBusy = False
    Exit Function

WriteFailed:
    errNumber = Err.Number
    errText = Err.Description
    Debug.Print "[VWG TRACKER] Activity log write failed at [" & stageText & "]: " & _
                errNumber & " - " & errText

    ' Keep the audit row uncommitted if its save failed. The details in B:I may
    ' remain in the next blank row; a retry will overwrite that same row.
    If writeStarted And targetRow >= FIRST_DATA_ROW Then
        On Error Resume Next
        rollbackErr = 0
        rollbackDescription = vbNullString
        If Not VWGX_WriteCellValueSafe(ws.Cells(targetRow, 1), Empty, _
            "clear incomplete Activity Log timestamp row " & CStr(targetRow), _
            rollbackErr, rollbackDescription) Then
            Debug.Print "[VWG TRACKER] Could not clear incomplete Activity Log timestamp: " & _
                        rollbackErr & " - " & rollbackDescription
        End If
        gXLWorkbookDirty = True
        SaveWorkbookIfNeeded
        If Err.Number <> 0 Then
            Debug.Print "[VWG TRACKER] Could not clear/save incomplete Activity Log timestamp: " & _
                        Err.Number & " - " & Err.Description
            Err.Clear
        End If
        On Error GoTo 0
    End If

    gTrackerWriteBusy = False
    VWGX_AppendActivity = False
End Function


'===========================================================
' PUBLIC: ROSTER / FAIRNESS COUNTS
'===========================================================

Public Function VWGX_ReadRoster() As Variant
    Dim ws As Object
    Dim lastRow As Long
    On Error GoTo ReadFailed
    If Not VWGX_EnsureConnected(False) Then Exit Function
    Set ws = gXLWorkbook.Worksheets("Team & Shifts")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(XL_UP).Row
    If lastRow < FIRST_DATA_ROW Then Exit Function
    VWGX_ReadRoster = ws.Range("A" & FIRST_DATA_ROW & ":F" & lastRow).Value
    Exit Function
ReadFailed:
    Debug.Print "[VWG TRACKER] Roster read failed: " & Err.Description
End Function

Public Function VWGX_GetAutoAllocationCounts(ByVal periodStart As Date) As Object
    Dim counts As Object
    Dim ws As Object
    Dim lastRow As Long
    Dim data As Variant
    Dim roster As Variant
    Dim r As Long
    Dim nm As String
    Dim allocatedAt As Variant

    Set counts = CreateObject("Scripting.Dictionary")
    counts.CompareMode = vbTextCompare

    On Error GoTo ReadFailed
    If Not VWGX_EnsureConnected(False) Then
        Set VWGX_GetAutoAllocationCounts = counts
        Exit Function
    End If

    roster = VWGX_ReadRoster()
    If IsArray(roster) Then
        For r = LBound(roster, 1) To UBound(roster, 1)
            nm = Trim$(CStr(roster(r, 1)))
            If Len(nm) > 0 Then counts(nm) = 0
        Next r
    End If

    ' Count automatic allocation reservations from the Assignments register.
    ' Allocated At (Egypt) is blank for pre-existing/manual assignments, so
    ' those do not bias fortnight fairness; pending reservations retain their intended owner and handovers preserve the original allocation.
    Set ws = gXLWorkbook.Worksheets("Assignments")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(XL_UP).Row
    If lastRow < FIRST_DATA_ROW Then GoTo Done

    data = ws.Range("A" & FIRST_DATA_ROW & ":U" & lastRow).Value
    For r = LBound(data, 1) To UBound(data, 1)
        nm = Trim$(CStr(data(r, 5)))            ' Original Assignee, column E
        allocatedAt = data(r, 15)               ' Allocated At (Egypt), column O
        If Len(nm) > 0 And IsDate(allocatedAt) Then
            If CDate(allocatedAt) >= periodStart Then
                If counts.Exists(nm) Then counts(nm) = CLng(counts(nm)) + 1
            End If
        End If
    Next r

Done:
    Set VWGX_GetAutoAllocationCounts = counts
    Exit Function

ReadFailed:
    Debug.Print "[VWG TRACKER] Allocation count read failed: " & Err.Description
    Set VWGX_GetAutoAllocationCounts = counts
End Function

Public Function VWGX_GetCurrentWorkloadCounts() As Object
    Dim counts As Object
    Dim ws As Object
    Dim lastRow As Long
    Dim data As Variant
    Dim roster As Variant
    Dim r As Long
    Dim nm As String
    Dim statusText As String

    Set counts = CreateObject("Scripting.Dictionary")
    counts.CompareMode = vbTextCompare

    On Error GoTo ReadFailed
    If Not VWGX_EnsureConnected(False) Then
        Set VWGX_GetCurrentWorkloadCounts = counts
        Exit Function
    End If

    roster = VWGX_ReadRoster()
    If IsArray(roster) Then
        For r = LBound(roster, 1) To UBound(roster, 1)
            nm = Trim$(CStr(roster(r, 1)))
            If Len(nm) > 0 Then counts(nm) = 0
        Next r
    End If

    Set ws = gXLWorkbook.Worksheets("Assignments")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(XL_UP).Row
    If lastRow < FIRST_DATA_ROW Then GoTo Done

    data = ws.Range("A" & FIRST_DATA_ROW & ":U" & lastRow).Value
    For r = LBound(data, 1) To UBound(data, 1)
        nm = Trim$(CStr(data(r, 6)))
        statusText = UCase$(Trim$(CStr(data(r, 10))))
        If Len(nm) > 0 And IsActiveStatus(statusText) Then
            If counts.Exists(nm) Then counts(nm) = CLng(counts(nm)) + 1
        End If
    Next r

Done:
    Set VWGX_GetCurrentWorkloadCounts = counts
    Exit Function

ReadFailed:
    Debug.Print "[VWG TRACKER] Workload count read failed: " & Err.Description
    Set VWGX_GetCurrentWorkloadCounts = counts
End Function

Private Function IsActiveStatus(ByVal statusText As String) As Boolean
    Select Case UCase$(Trim$(statusText))
        Case "ASSIGNED", "HANDED OVER", "HANDOVER PENDING", "RECOVERED - REVIEW", "NEEDS REVIEW", "RECOVERY FAILED", "ASSIGNING", "ASSIGNING RETRY"
            IsActiveStatus = True
        Case Else
            IsActiveStatus = False
    End Select
End Function

Public Sub VWGX_RefreshSummary()
    Dim wsTeam As Object
    Dim wsDashboard As Object
    Dim ws As Object
    Dim allocCounts As Object
    Dim workloadCounts As Object
    Dim roster As Variant
    Dim data As Variant
    Dim lastRow As Long
    Dim r As Long
    Dim targetRow As Long
    Dim nameText As String
    Dim assignedCount As Long
    Dim handoverCount As Long
    Dim incidentCount As Long
    Dim activeCount As Long
    Dim startValue As Variant
    Dim stageText As String
    Dim errNumber As Long
    Dim errText As String
    Dim cellWriteOK As Boolean

    On Error GoTo RefreshFailed

    stageText = "connect to workbook"
    If Not VWGX_EnsureConnected(False) Then Exit Sub
    If gTrackerWriteBusy Then
        Debug.Print "[VWG TRACKER] Dashboard refresh skipped: another tracker write is active."
        Exit Sub
    End If
    gTrackerWriteBusy = True

    stageText = "read roster period setting"
    startValue = VWGX_GetSetting("Roster Period Start")
    If IsDate(startValue) Then
        Set allocCounts = VWGX_GetAutoAllocationCounts(CDate(startValue))
    Else
        Set allocCounts = VWGX_GetAutoAllocationCounts(Date)
    End If
    Set workloadCounts = VWGX_GetCurrentWorkloadCounts()

    stageText = "get Team & Shifts worksheet"
    Set wsTeam = gXLWorkbook.Worksheets("Team & Shifts")
    stageText = "read roster"
    roster = VWGX_ReadRoster()
    If IsArray(roster) Then
        For r = LBound(roster, 1) To UBound(roster, 1)
            nameText = Trim$(CStr(roster(r, 1)))
            targetRow = r + FIRST_DATA_ROW - 1
            If Len(nameText) > 0 Then
                If Not allocCounts Is Nothing Then
                    If allocCounts.Exists(nameText) Then
                        stageText = "write Team & Shifts allocation count, row " & CStr(targetRow) & _
                                    ", column G, colleague " & nameText
                        cellWriteOK = VWGX_WriteCellValueSafe(wsTeam.Cells(targetRow, 7), _
                                        allocCounts(nameText), stageText, errNumber, errText)
                        If cellWriteOK Then gXLWorkbookDirty = True
                    End If
                End If
                If Not workloadCounts Is Nothing Then
                    If workloadCounts.Exists(nameText) Then
                        stageText = "write Team & Shifts workload count, row " & CStr(targetRow) & _
                                    ", column H, colleague " & nameText
                        cellWriteOK = VWGX_WriteCellValueSafe(wsTeam.Cells(targetRow, 8), _
                                        workloadCounts(nameText), stageText, errNumber, errText)
                        If cellWriteOK Then gXLWorkbookDirty = True
                    End If
                End If
            End If
        Next r
    End If

    stageText = "get Dashboard worksheet"
    Set wsDashboard = gXLWorkbook.Worksheets("Dashboard")
    stageText = "get Assignments worksheet"
    Set ws = gXLWorkbook.Worksheets("Assignments")

    stageText = "read assignment summary"
    lastRow = ws.Cells(ws.Rows.Count, 1).End(XL_UP).Row
    If lastRow >= FIRST_DATA_ROW Then
        data = ws.Range("A" & FIRST_DATA_ROW & ":U" & lastRow).Value2
        For r = LBound(data, 1) To UBound(data, 1)
            If Len(Trim$(CStr(data(r, 1)))) > 0 Then assignedCount = assignedCount + 1
            If IsActiveStatus(UCase$(Trim$(CStr(data(r, 10))))) Then activeCount = activeCount + 1
        Next r
    End If

    stageText = "get Activity Log worksheet"
    Set ws = gXLWorkbook.Worksheets("Activity Log")
    stageText = "read activity summary"
    lastRow = ws.Cells(ws.Rows.Count, 1).End(XL_UP).Row
    If lastRow >= FIRST_DATA_ROW Then
        data = ws.Range("A" & FIRST_DATA_ROW & ":I" & lastRow).Value2
        For r = LBound(data, 1) To UBound(data, 1)
            Select Case UCase$(Trim$(CStr(data(r, 2))))
                Case "HANDOVER": handoverCount = handoverCount + 1
                Case "CATEGORY_REMOVAL_DETECTED", "CATEGORY_RESTORED", "RECOVERY_FAILED"
                    incidentCount = incidentCount + 1
            End Select
        Next r
    End If

    stageText = "write Dashboard KPI A5"
    cellWriteOK = VWGX_WriteCellValueSafe(wsDashboard.Range("A5"), assignedCount, stageText, errNumber, errText)
    If cellWriteOK Then gXLWorkbookDirty = True
    stageText = "write Dashboard KPI D5"
    cellWriteOK = VWGX_WriteCellValueSafe(wsDashboard.Range("D5"), activeCount, stageText, errNumber, errText)
    If cellWriteOK Then gXLWorkbookDirty = True
    stageText = "write Dashboard KPI G5"
    cellWriteOK = VWGX_WriteCellValueSafe(wsDashboard.Range("G5"), handoverCount, stageText, errNumber, errText)
    If cellWriteOK Then gXLWorkbookDirty = True
    stageText = "write Dashboard KPI J5"
    cellWriteOK = VWGX_WriteCellValueSafe(wsDashboard.Range("J5"), incidentCount, stageText, errNumber, errText)
    If cellWriteOK Then gXLWorkbookDirty = True

    ' Full recalculation is deliberately omitted. These KPI cells are hard values.
    stageText = "save dashboard summary"
    SaveWorkbookIfNeeded

    gTrackerWriteBusy = False
    Exit Sub

RefreshFailed:
    errNumber = Err.Number
    errText = Err.Description
    Debug.Print "[VWG TRACKER] Dashboard refresh failed at [" & stageText & "]: " & _
                errNumber & " - " & errText
    gTrackerWriteBusy = False
End Sub


'===========================================================
' PUBLIC: DAILY CSV SNAPSHOT EXPORT
' The output is a snapshot for manual Google Sheets import or a
' later authorised sync. This does not update Google Sheets itself.
'===========================================================

Public Sub VWGX_ExportDailyCSVs()
    Dim exportFolder As String
    Dim sheetsToExport As Variant
    Dim item As Variant
    Dim ws As Object
    Dim exportBook As Object
    Dim outputPath As String
    Dim todayStamp As String

    On Error GoTo ExportFailed
    If Not VWGX_EnsureConnected(True) Then Exit Sub

    exportFolder = Environ$("USERPROFILE") & "\Documents\VWG_Dispatch_Export"
    If Len(Dir$(exportFolder, vbDirectory)) = 0 Then MkDir exportFolder

    todayStamp = Format$(Date, "yyyy-mm-dd")
    sheetsToExport = Array("Dashboard", "Assignments", "Activity Log", "Team & Shifts", "Settings")

    For Each item In sheetsToExport
        Set ws = gXLWorkbook.Worksheets(CStr(item))
        ws.Copy
        Set exportBook = gXLApp.ActiveWorkbook
        outputPath = exportFolder & "\VWG_" & Replace(CStr(item), " ", "_") & "_" & todayStamp & ".csv"

        gXLApp.DisplayAlerts = False
        exportBook.SaveAs outputPath, XL_CSV_UTF8
        exportBook.Close False
        gXLApp.DisplayAlerts = True
    Next item

    MsgBox "Daily CSV snapshots exported to:" & vbCrLf & exportFolder & vbCrLf & vbCrLf & _
           "Import the relevant CSV into the matching Google Sheets tab. This macro does not upload or sync to Google Sheets.", _
           vbInformation, "VWG Daily Export"
    Exit Sub

ExportFailed:
    On Error Resume Next
    gXLApp.DisplayAlerts = True
    If Not exportBook Is Nothing Then exportBook.Close False
    MsgBox "CSV export failed: " & Err.Number & " - " & Err.Description, _
           vbExclamation, "VWG Daily Export"
End Sub

'===========================================================
' INTERNAL HELPERS
'===========================================================

Private Function WorkbookIsUsable(ByVal workbookObject As Object) As Boolean
    Dim nameCheck As String
    On Error GoTo NotUsable
    nameCheck = CStr(workbookObject.Name)
    WorkbookIsUsable = ValidateWorkbook(workbookObject)
    Exit Function
NotUsable:
    WorkbookIsUsable = False
End Function

Private Function ValidateWorkbook(ByVal workbookObject As Object) As Boolean
    Dim names As Variant
    Dim nm As Variant
    Dim ws As Object

    On Error GoTo InvalidWorkbook
    names = Array("Dashboard", "Assignments", "Activity Log", "Team & Shifts", "Settings")
    For Each nm In names
        Set ws = Nothing
        On Error Resume Next
        Set ws = workbookObject.Worksheets(CStr(nm))
        On Error GoTo InvalidWorkbook
        If ws Is Nothing Then Exit Function
    Next nm

    If CStr(workbookObject.Worksheets("Assignments").Cells(HEADER_ROW, 1).Value) <> "Email Key / ID" Then Exit Function
    If CStr(workbookObject.Worksheets("Assignments").Cells(HEADER_ROW, 21).Value) <> "Handover Count" Then Exit Function
    If CStr(workbookObject.Worksheets("Activity Log").Cells(HEADER_ROW, 2).Value) <> "Event Type" Then Exit Function
    If CStr(workbookObject.Worksheets("Team & Shifts").Cells(HEADER_ROW, 1).Value) <> "Colleague" Then Exit Function
    If CStr(workbookObject.Worksheets("Settings").Cells(HEADER_ROW, 1).Value) <> "Setting" Then Exit Function

    ValidateWorkbook = True
    Exit Function
InvalidWorkbook:
    ValidateWorkbook = False
End Function

Private Function StateDirectory() As String
    StateDirectory = Environ$("LOCALAPPDATA") & "\" & STATE_FOLDER
End Function

Private Sub EnsureStateDirectory()
    If Len(Environ$("LOCALAPPDATA")) = 0 Then Err.Raise vbObjectError + 1502, "EnsureStateDirectory", "LOCALAPPDATA is unavailable."
    If Len(Dir$(StateDirectory(), vbDirectory)) = 0 Then MkDir StateDirectory()
End Sub

Private Function WorkbookPathFile() As String
    WorkbookPathFile = StateDirectory() & "\tracker-path.txt"
End Function

Private Function ReadSavedWorkbookPath() As String
    Dim f As Integer
    Dim value As String
    On Error GoTo ReadFailed
    EnsureStateDirectory
    If Len(Dir$(WorkbookPathFile())) = 0 Then Exit Function
    f = FreeFile
    Open WorkbookPathFile() For Input As #f
    If Not EOF(f) Then Line Input #f, value
    Close #f
    ReadSavedWorkbookPath = Trim$(value)
    Exit Function
ReadFailed:
    On Error Resume Next
    Close #f
End Function

Private Sub SaveWorkbookPath(ByVal pathText As String)
    Dim f As Integer
    EnsureStateDirectory
    f = FreeFile
    Open WorkbookPathFile() For Output As #f
    Print #f, pathText
    Close #f
End Sub

Private Function VWGX_WriteCellValueSafe(ByVal targetCell As Object, _
                                          ByVal newValue As Variant, _
                                          ByVal stageText As String, _
                                          ByRef finalError As Long, _
                                          ByRef finalDescription As String) As Boolean
    Dim attempt As Long

    finalError = 0
    finalDescription = vbNullString

    For attempt = 1 To 3
        On Error Resume Next
        targetCell.Value2 = newValue
        finalError = Err.Number
        finalDescription = Err.Description
        Err.Clear
        On Error GoTo 0

        If finalError = 0 Then
            VWGX_WriteCellValueSafe = True
            Exit Function
        End If

        If (finalError <> VWGX_RETRYABLE_EXCEL_ERROR And _
            finalError <> VWGX_RPC_E_CALL_REJECTED) Or attempt = 3 Then Exit For

        Debug.Print "[VWG TRACKER] Transient Excel write error " & finalError & _
                    " at [" & stageText & "]; retry " & CStr(attempt + 1) & " of 3."
        VWGX_Sleep 150
    Next attempt

    Debug.Print "[VWG TRACKER] Excel value write failed at [" & stageText & "] after " & _
                CStr(attempt) & " attempt(s): " & finalError & " - " & finalDescription
    VWGX_WriteCellValueSafe = False
End Function

Private Sub VWGX_TryNumberFormat(ByVal targetCell As Object, _
                                 ByVal formatText As String, _
                                 ByVal stageText As String)
    Dim attempt As Long
    Dim formatError As Long
    Dim formatDescription As String

    For attempt = 1 To 2
        On Error Resume Next
        targetCell.NumberFormat = formatText
        formatError = Err.Number
        formatDescription = Err.Description
        Err.Clear
        On Error GoTo 0

        If formatError = 0 Then Exit Sub
        If (formatError <> VWGX_RETRYABLE_EXCEL_ERROR And _
            formatError <> VWGX_RPC_E_CALL_REJECTED) Or attempt = 2 Then Exit For
        VWGX_Sleep 100
    Next attempt

    ' Display formatting must never cancel a reservation, assignment or audit event.
    Debug.Print "[VWG TRACKER] Non-fatal display-format warning at [" & stageText & "]: " & _
                formatError & " - " & formatDescription & ". Cell value was still written."
End Sub

Public Sub VWGX_ShowBuildInfo()
    Debug.Print "========== VWG TRACKER BUILD =========="
    Debug.Print "Build ID: " & VWGX_BUILD_ID
    Debug.Print "Active module: modVWGExcelTracker"
    Debug.Print "Schema: Assignments A:U / Activity Log A:I / headers row 5"
    Debug.Print "Write policy: 3-attempt transient COM retry; NumberFormat errors are non-fatal"
    If VWGX_EnsureConnected(False) Then
        Debug.Print "Workbook: " & CStr(gXLWorkbook.Name)
        Debug.Print "Workbook path: " & CStr(gXLWorkbook.FullName)
        Debug.Print "Workbook read-only: " & CStr(gXLWorkbook.ReadOnly)
    Else
        Debug.Print "Workbook: NOT CONNECTED"
    End If
    Debug.Print "======================================="
End Sub

Private Function VWGX_ValuesMatch(ByVal valueA As Variant, ByVal valueB As Variant) As Boolean
    On Error GoTo NoMatch

    If IsError(valueA) Or IsError(valueB) Then Exit Function

    If IsNull(valueA) Or IsNull(valueB) Then
        VWGX_ValuesMatch = (IsNull(valueA) And IsNull(valueB))
        Exit Function
    End If

    If IsEmpty(valueA) Or IsEmpty(valueB) Then
        VWGX_ValuesMatch = (IsEmpty(valueA) And IsEmpty(valueB))
        Exit Function
    End If

    ' Preserve exact text for keys/IDs (including leading zeroes).
    If VarType(valueA) = vbString And VarType(valueB) = vbString Then
        VWGX_ValuesMatch = (StrComp(CStr(valueA), CStr(valueB), vbBinaryCompare) = 0)
    ElseIf VarType(valueA) = vbDate Or VarType(valueB) = vbDate Then
        VWGX_ValuesMatch = (Abs(CDbl(CDate(valueA)) - CDbl(CDate(valueB))) < 0.0000001)
    ElseIf IsNumeric(valueA) And IsNumeric(valueB) Then
        VWGX_ValuesMatch = (Abs(CDbl(valueA) - CDbl(valueB)) < 0.0000001)
    Else
        VWGX_ValuesMatch = (StrComp(CStr(valueA), CStr(valueB), vbBinaryCompare) = 0)
    End If
    Exit Function

NoMatch:
    VWGX_ValuesMatch = False
End Function

'===========================================================
' OPTIONAL DIAGNOSTICS - SAFE, NO ASSIGNMENT OR OUTLOOK CHANGES
'===========================================================

Public Sub VWGX_DiagnoseConnection()
    Debug.Print "========== VWG CONNECTION TEST =========="
    Debug.Print "EnsureConnected: "; VWGX_EnsureConnected(False)

    If gXLApp Is Nothing Then
        Debug.Print "Excel Application: NOT CONNECTED"
    Else
        Debug.Print "Excel Application: CONNECTED"
        On Error Resume Next
        Debug.Print "Excel Visible: "; CStr(gXLApp.Visible)
        On Error GoTo 0
    End If

    If gXLWorkbook Is Nothing Then
        Debug.Print "Tracker Workbook: NOT CONNECTED"
    Else
        Debug.Print "Tracker Workbook: CONNECTED"
        On Error Resume Next
        Debug.Print "Workbook Name: "; CStr(gXLWorkbook.Name)
        Debug.Print "Workbook Path: "; CStr(gXLWorkbook.FullName)
        Debug.Print "Workbook ReadOnly: "; CStr(gXLWorkbook.ReadOnly)
        On Error GoTo 0
    End If

    Debug.Print "========================================="
End Sub

Public Sub VWGX_TestTrackerWrite()
    Dim ws As Object
    Dim beforeValue As Variant
    Dim afterValue As Variant
    Dim errNumber As Long
    Dim errText As String

    On Error GoTo TestFailed
    If Not VWGX_EnsureConnected(False) Then
        Debug.Print "Tracker write test: workbook not connected."
        Exit Sub
    End If

    Set ws = gXLWorkbook.Worksheets("Assignments")
    beforeValue = ws.Range("A1").Value2
    Debug.Print "Before write: "; CStr(beforeValue)

    ' Write the exact existing value back to the same title cell. No tracker
    ' records, assignments, categories or audit events are changed.
    If Not VWGX_WriteCellValueSafe(ws.Range("A1"), beforeValue, _
                                   "diagnostic write to Assignments A1", errNumber, errText) Then
        Err.Raise errNumber, "VWGX_TestTrackerWrite", errText
    End If
    Debug.Print "Cell write: SUCCESS"
    afterValue = ws.Range("A1").Value2

    If VWGX_ValuesMatch(beforeValue, afterValue) Then
        Debug.Print "Read-back: SUCCESS"
    Else
        Debug.Print "Read-back: FAILED - value changed during test."
    End If
    Exit Sub

TestFailed:
    Debug.Print "Tracker write test failed: " & Err.Number & " - " & Err.Description
End Sub

