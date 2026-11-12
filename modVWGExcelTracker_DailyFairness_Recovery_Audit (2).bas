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
Private Const VWGX_BUILD_ID As String = "VWG-TRACKER-2026.11.12-ACTIVITY-VERIFY-R3"
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
Private gFlushingActivityFallback As Boolean
Private gFlushingAssignmentFallback As Boolean
Private gDashboardRefreshPending As Boolean
Private gDashboardRefreshDate As Variant

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
    Dim lockAcquired As Boolean
    On Error GoTo SaveFailed
    If gTrackerWriteBusy Then
        gXLWorkbookDirty = True
        Debug.Print "[VWG TRACKER] Workbook save deferred: another tracker write is active."
        Exit Sub
    End If
    If Not VWGX_EnsureConnected(False) Then Exit Sub
    If gTrackerWriteBusy Then
        gXLWorkbookDirty = True
        Debug.Print "[VWG TRACKER] Workbook save deferred after connection: tracker write became active."
        Exit Sub
    End If
    gTrackerWriteBusy = True
    lockAcquired = True
    SaveWorkbookNow
    gTrackerWriteBusy = False
    lockAcquired = False
    Exit Sub
SaveFailed:
    Debug.Print "[VWG TRACKER] Save failed: " & Err.Number & " - " & Err.Description
    gXLWorkbookDirty = True
    If lockAcquired Then gTrackerWriteBusy = False
End Sub

' Batch operations are used during large baseline scans so each existing
' tagged email does not cause several full workbook saves.
Public Function VWGX_BeginBatch() As Boolean
    If Not VWGX_EnsureConnected(False) Then Exit Function
    gXLBatchDepth = gXLBatchDepth + 1
    VWGX_BeginBatch = True
End Function

Public Function VWGX_EndBatch(Optional ByVal forceSave As Boolean = True) As Boolean
    Dim lockAcquired As Boolean
    On Error GoTo EndFailed
    If gXLBatchDepth > 0 Then gXLBatchDepth = gXLBatchDepth - 1
    If gXLBatchDepth = 0 Then
        If gXLWorkbookDirty Or forceSave Then
            If gTrackerWriteBusy Then
                Debug.Print "[VWG TRACKER] Batch save deferred: another tracker write is active. Workbook remains marked dirty."
                VWGX_EndBatch = False
                Exit Function
            End If
            gTrackerWriteBusy = True
            lockAcquired = True
            SaveWorkbookNow
            gTrackerWriteBusy = False
            lockAcquired = False
        End If
    End If
    VWGX_EndBatch = True
    Exit Function
EndFailed:
    Debug.Print "[VWG TRACKER] Batch save failed: " & Err.Number & " - " & Err.Description
    VWGX_EndBatch = False
    If lockAcquired Then gTrackerWriteBusy = False
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

        If (savedErr <> VWGX_RETRYABLE_EXCEL_ERROR And savedErr <> VWGX_RPC_E_CALL_REJECTED) Or attempt = 3 Then
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
    Dim lockAcquired As Boolean

    On Error GoTo WriteFailed
    If gTrackerWriteBusy Then
        Debug.Print "[VWG TRACKER] Setting write deferred/rejected while tracker is busy: " & settingName
        Exit Function
    End If
    If Not VWGX_EnsureConnected(False) Then Exit Function
    If gTrackerWriteBusy Then Exit Function

    gTrackerWriteBusy = True
    lockAcquired = True
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
    gTrackerWriteBusy = False
    lockAcquired = False
    Exit Function

WriteFailed:
    Debug.Print "[VWG TRACKER] Setting write failed: " & settingName & " | " & Err.Number & " - " & Err.Description
    VWGX_SetSetting = False
    If lockAcquired Then gTrackerWriteBusy = False
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
    Dim keyData As Variant
    Dim oldValues As Variant
    Dim leftValues(1 To 1, 1 To 8) As Variant
    Dim rightValues(1 To 1, 1 To 11) As Variant
    Dim verifyValues As Variant
    Dim i As Long
    Dim duplicateCount As Long
    Dim isNewRow As Boolean
    Dim writeStarted As Boolean
    Dim rollbackOK As Boolean
    Dim stageText As String
    Dim errNumber As Long
    Dim errText As String
    Dim cellValue As Variant
    Dim lockAcquired As Boolean
    Dim recordValid As Boolean
    Dim fallbackSaved As Boolean

    On Error GoTo WriteFailed

    stageText = "validate assignment record"
    If Not IsArray(record) Then Err.Raise vbObjectError + 1601, "VWGX_UpsertAssignment", "Assignment record is not an array."
    If LBound(record) <> 0 Or UBound(record) < 20 Then
        Err.Raise vbObjectError + 1602, "VWGX_UpsertAssignment", "Assignment record must contain zero-based fields 0 to 20."
    End If
    If IsNull(record(0)) Or IsEmpty(record(0)) Then
        Err.Raise vbObjectError + 1603, "VWGX_UpsertAssignment", "Assignment Email Key is blank."
    End If
    keyText = Trim$(CStr(record(0)))
    If Len(keyText) = 0 Then Err.Raise vbObjectError + 1604, "VWGX_UpsertAssignment", "Assignment Email Key is blank."
    recordValid = True

    ' Never drop a write just because another Excel operation owns the lock.
    ' Persist the full row locally, but return False so Outlook categories are
    ' NOT changed until the reservation has actually been committed to Excel.
    If gTrackerWriteBusy Then
        fallbackSaved = VWGX_AppendAssignmentToFallback(record)
        If fallbackSaved Then
            Debug.Print "[VWG TRACKER] Assignment write queued for retry while Excel was busy. Key=" & keyText
        Else
            Debug.Print "[VWG TRACKER] CRITICAL: Assignment write was blocked and could not be queued. Key=" & keyText
        End If
        VWGX_UpsertAssignment = False
        Exit Function
    End If

    stageText = "connect to workbook"
    If Not VWGX_EnsureConnected(False) Then
        If Not gFlushingAssignmentFallback Then fallbackSaved = VWGX_AppendAssignmentToFallback(record)
        If fallbackSaved Then Debug.Print "[VWG TRACKER] Assignment write queued because workbook connection was unavailable. Key=" & keyText
        Exit Function
    End If
    If gTrackerWriteBusy Then
        fallbackSaved = VWGX_AppendAssignmentToFallback(record)
        If fallbackSaved Then Debug.Print "[VWG TRACKER] Assignment write queued after connection detected a busy tracker. Key=" & keyText
        Exit Function
    End If

    gTrackerWriteBusy = True
    lockAcquired = True

    stageText = "get Assignments worksheet"
    Set ws = gXLWorkbook.Worksheets("Assignments")
    stageText = "find last assignment row"
    lastRow = ws.Cells(ws.Rows.Count, 1).End(XL_UP).Row
    If lastRow < HEADER_ROW Then lastRow = HEADER_ROW
    targetRow = 0
    duplicateCount = 0

    ' Bulk-read the key column once. Avoid one late-bound Excel call per row.
    stageText = "read assignment keys"
    If lastRow >= FIRST_DATA_ROW Then
        keyData = ws.Range("A" & FIRST_DATA_ROW & ":A" & lastRow).Value2
        For r = LBound(keyData, 1) To UBound(keyData, 1)
            If StrComp(Trim$(CStr(keyData(r, 1))), keyText, vbTextCompare) = 0 Then
                If targetRow = 0 Then
                    targetRow = r + FIRST_DATA_ROW - 1
                Else
                    duplicateCount = duplicateCount + 1
                End If
            End If
        Next r
    End If

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

    If isNewRow Then
        stageText = "write Email Key / ID (column A) row " & CStr(targetRow)
        writeStarted = True
        If Not VWGX_WriteCellValueSafe(ws.Cells(targetRow, 1), keyText, stageText, errNumber, errText) Then
            Err.Raise errNumber, "VWGX_UpsertAssignment", errText
        End If
        gXLWorkbookDirty = True
    End If

    ' B:I and K:U are written as two compact row ranges. Status in J remains
    ' the final commit field so readers never treat a partial row as complete.
    ' Mark the write as started for BOTH new and existing rows so an error in
    ' either payload range restores the snapshot instead of leaving a half-row.
    writeStarted = True
    For i = 1 To 8
        cellValue = record(i)
        If IsNull(cellValue) Or IsEmpty(cellValue) Then cellValue = Empty
        leftValues(1, i) = cellValue
    Next i
    For i = 10 To 20
        cellValue = record(i)
        If IsNull(cellValue) Or IsEmpty(cellValue) Then cellValue = Empty
        rightValues(1, i - 9) = cellValue
    Next i

    stageText = "write assignment payload B:I row " & CStr(targetRow)
    If Not VWGX_WriteRangeSafe(ws.Range("B" & targetRow & ":I" & targetRow), _
                               leftValues, stageText, errNumber, errText) Then
        Err.Raise errNumber, "VWGX_UpsertAssignment", errText
    End If
    gXLWorkbookDirty = True

    stageText = "write assignment payload K:U row " & CStr(targetRow)
    If Not VWGX_WriteRangeSafe(ws.Range("K" & targetRow & ":U" & targetRow), _
                               rightValues, stageText, errNumber, errText) Then
        Err.Raise errNumber, "VWGX_UpsertAssignment", errText
    End If
    gXLWorkbookDirty = True

    ' Apply display formats once when a row is created; existing-row updates
    ' skip these late-bound COM calls to reduce Outlook/Excel lag.
    If isNewRow Then
        VWGX_TryNumberFormat ws.Cells(targetRow, 1), "@", "Assignments A" & targetRow
        VWGX_TryNumberFormat ws.Cells(targetRow, 2), "dd/mmm/yyyy hh:mm:ss", "Assignments B" & targetRow
        VWGX_TryNumberFormat ws.Cells(targetRow, 9), "dd/mmm/yyyy hh:mm:ss", "Assignments I" & targetRow
        VWGX_TryNumberFormat ws.Cells(targetRow, 12), "dd/mmm/yyyy hh:mm:ss", "Assignments L" & targetRow
        VWGX_TryNumberFormat ws.Cells(targetRow, 14), "dd/mmm/yyyy hh:mm:ss", "Assignments N" & targetRow
        VWGX_TryNumberFormat ws.Cells(targetRow, 15), "dd/mmm/yyyy hh:mm:ss", "Assignments O" & targetRow
        VWGX_TryNumberFormat ws.Cells(targetRow, 16), "dd/mmm/yyyy", "Assignments P" & targetRow
        VWGX_TryNumberFormat ws.Cells(targetRow, 17), "@", "Assignments Q" & targetRow
        VWGX_TryNumberFormat ws.Cells(targetRow, 18), "@", "Assignments R" & targetRow
        VWGX_TryNumberFormat ws.Cells(targetRow, 19), "dd/mmm/yyyy hh:mm:ss", "Assignments S" & targetRow
    End If

    stageText = "write Status (column J) row " & CStr(targetRow)
    cellValue = record(9)
    If IsNull(cellValue) Or IsEmpty(cellValue) Then cellValue = Empty
    If Not VWGX_WriteCellValueSafe(ws.Cells(targetRow, 10), cellValue, stageText, errNumber, errText) Then
        Err.Raise errNumber, "VWGX_UpsertAssignment", errText
    End If
    gXLWorkbookDirty = True

    stageText = "verify assignment key/status/allocation timestamp row " & CStr(targetRow)
    verifyValues = ws.Range("A" & targetRow & ":O" & targetRow).Value2
    If StrComp(Trim$(CStr(verifyValues(1, 1))), keyText, vbTextCompare) <> 0 Then
        Err.Raise vbObjectError + 1605, "VWGX_UpsertAssignment", "Email Key read-back did not match."
    End If
    If Not VWGX_ValuesMatch(verifyValues(1, 10), record(9)) Then
        Err.Raise vbObjectError + 1606, "VWGX_UpsertAssignment", "Status read-back did not match."
    End If
    If Not VWGX_ValuesMatch(verifyValues(1, 15), record(14)) Then
        Err.Raise vbObjectError + 1607, "VWGX_UpsertAssignment", _
                  "Allocated At (Egypt) timestamp read-back did not match column O."
    End If

    stageText = "save tracker workbook"
    SaveWorkbookIfNeeded
    VWGX_UpsertAssignment = True
    If lockAcquired Then gTrackerWriteBusy = False
    lockAcquired = False
    Exit Function

WriteFailed:
    errNumber = Err.Number
    errText = Err.Description
    Debug.Print "[VWG TRACKER] Assignment write failed at [" & stageText & "]: " & _
                errNumber & " - " & errText

    ' Roll back only the row being upserted. Never delete or merge another row.
    If writeStarted And targetRow >= FIRST_DATA_ROW Then
        rollbackOK = False
        On Error Resume Next
        If isNewRow Then
            ws.Range("A" & targetRow & ":U" & targetRow).ClearContents
            rollbackOK = (Err.Number = 0)
        ElseIf IsArray(oldValues) Then
            ws.Range("A" & targetRow & ":U" & targetRow).Value2 = oldValues
            rollbackOK = (Err.Number = 0)
        End If
        Err.Clear
        If rollbackOK Then
            gXLWorkbookDirty = True
            SaveWorkbookIfNeeded
            If Err.Number <> 0 Then
                Debug.Print "[VWG TRACKER] Assignment rollback save failed: " & Err.Number & " - " & Err.Description
                Err.Clear
            Else
                Debug.Print "[VWG TRACKER] Assignment row rollback completed for row " & CStr(targetRow) & "."
            End If
        Else
            Debug.Print "[VWG TRACKER] CRITICAL: Assignment row rollback was incomplete at row " & _
                        CStr(targetRow) & ". Inspect this row before restart."
        End If
        On Error GoTo 0
    End If

    If lockAcquired Then gTrackerWriteBusy = False
    lockAcquired = False

    ' A failed write is retryable from the durable local queue. During a queue
    ' flush, the current processing file already retains the failed record.
    If recordValid And Not gFlushingAssignmentFallback Then
        fallbackSaved = VWGX_AppendAssignmentToFallback(record)
        If fallbackSaved Then
            Debug.Print "[VWG TRACKER] Failed assignment row preserved for retry. Key=" & keyText
        Else
            Debug.Print "[VWG TRACKER] CRITICAL: Failed assignment row could not be preserved. Key=" & keyText
        End If
    End If
    VWGX_UpsertAssignment = False
End Function

'===========================================================
' PUBLIC: APPEND ACTIVITY
' A zero-based 9-field Variant array maps to Activity Log A:I.
'===========================================================

Public Function VWGX_AppendActivity(ByVal eventRecord As Variant) As Boolean
    Dim ws As Object
    Dim lastRowA As Long
    Dim lastRowI As Long
    Dim lastRow As Long
    Dim scanStartRow As Long
    Dim targetRow As Long
    Dim existing As Variant
    Dim rowData(1 To 1, 1 To 8) As Variant
    Dim dataRow As Variant
    Dim r As Long
    Dim i As Long
    Dim allMatch As Boolean
    Dim stageText As String
    Dim errNumber As Long
    Dim errText As String
    Dim cellValue As Variant
    Dim lockAcquired As Boolean

    On Error GoTo WriteFailed

    If Not IsArray(eventRecord) Then Err.Raise vbObjectError + 1611, "VWGX_AppendActivity", "Activity record is not an array."
    If LBound(eventRecord) <> 0 Or UBound(eventRecord) < 8 Then
        Err.Raise vbObjectError + 1612, "VWGX_AppendActivity", "Activity record must contain zero-based fields 0 to 8."
    End If

    ' If another tracker operation is active, persist the event locally instead
    ' of dropping it. The maintenance timer flushes this file to Activity Log.
    If gTrackerWriteBusy Then
        VWGX_AppendActivity = VWGX_AppendActivityToFallback(eventRecord)
        If VWGX_AppendActivity Then Debug.Print "[VWG TRACKER] Activity event queued to local fallback while Excel was busy: " & CStr(eventRecord(1))
        Exit Function
    End If

    stageText = "connect to workbook"
    If Not VWGX_EnsureConnected(False) Then Err.Raise vbObjectError + 1610, "VWGX_AppendActivity", "Tracker workbook is not connected."
    If gTrackerWriteBusy Then
        VWGX_AppendActivity = VWGX_AppendActivityToFallback(eventRecord)
        Exit Function
    End If
    gTrackerWriteBusy = True
    lockAcquired = True

    stageText = "get Activity Log worksheet"
    Set ws = gXLWorkbook.Worksheets("Activity Log")

    stageText = "find last used activity row"
    lastRowA = ws.Cells(ws.Rows.Count, 1).End(XL_UP).Row
    lastRowI = ws.Cells(ws.Rows.Count, 9).End(XL_UP).Row
    lastRow = lastRowA
    If lastRowI > lastRow Then lastRow = lastRowI
    If lastRow < HEADER_ROW Then lastRow = HEADER_ROW
    targetRow = 0
    scanStartRow = FIRST_DATA_ROW
    If lastRow - FIRST_DATA_ROW + 1 > 1000 Then scanStartRow = lastRow - 999

    ' One COM read for recent duplicate detection; compare in memory, not cell by cell.
    ' The bounded window avoids rescanning an ever-growing audit history on every event.
    If lastRow >= FIRST_DATA_ROW Then
        existing = ws.Range("A" & scanStartRow & ":I" & lastRow).Value2
        For r = LBound(existing, 1) To UBound(existing, 1)
            allMatch = True
            For i = 0 To 8
                If Not VWGX_ValuesMatch(existing(r, i + 1), eventRecord(i)) Then
                    allMatch = False
                    Exit For
                End If
            Next i
            If allMatch Then
                stageText = "save previously recorded activity event"
                SaveWorkbookIfNeeded
                VWGX_AppendActivity = True
                If lockAcquired Then gTrackerWriteBusy = False
                lockAcquired = False
                Exit Function
            End If

            ' Reuse a partially-written row whose timestamp commit marker is
            ' blank but whose B:I payload matches this event exactly.
            If targetRow = 0 Then
                If IsEmpty(existing(r, 1)) Or Len(Trim$(CStr(existing(r, 1)))) = 0 Then
                    allMatch = True
                    For i = 1 To 8
                        If Not VWGX_ValuesMatch(existing(r, i + 1), eventRecord(i)) Then
                            allMatch = False
                            Exit For
                        End If
                    Next i
                    If allMatch Then targetRow = r + scanStartRow - 1
                End If
            End If
        Next r
    End If

    If targetRow = 0 Then targetRow = lastRow + 1

    ' Keep text identity fields as text before writing. Excel can otherwise
    ' coerce "TRUE" / "FALSE" into Boolean values and blank strings into Empty.
    ' The verifier below also accepts those equivalent Excel read-back forms.
    stageText = "format Activity Log text fields row " & CStr(targetRow)
    VWGX_TryNumberFormat ws.Range("B" & targetRow & ":G" & targetRow), "@", _
                         "Activity Log B:G row " & CStr(targetRow)
    VWGX_TryNumberFormat ws.Cells(targetRow, 9), "@", _
                         "Activity Log I" & targetRow
    VWGX_TryNumberFormat ws.Cells(targetRow, 8), "0", _
                         "Activity Log H" & targetRow

    ' Write B:I in one operation; timestamp A is the final commit marker.
    For i = 1 To 8
        rowData(1, i) = eventRecord(i)
    Next i
    stageText = "write Activity Log payload B:I row " & CStr(targetRow)
    If Not VWGX_WriteRangeSafe(ws.Range("B" & targetRow & ":I" & targetRow), _
                               rowData, stageText, errNumber, errText) Then
        Err.Raise errNumber, "VWGX_AppendActivity", errText
    End If
    gXLWorkbookDirty = True

    VWGX_TryNumberFormat ws.Cells(targetRow, 1), "dd/mmm/yyyy hh:mm:ss", "Activity Log A" & targetRow
    cellValue = eventRecord(0)
    If IsNull(cellValue) Or IsEmpty(cellValue) Then cellValue = Empty
    stageText = "write Activity Log timestamp commit marker A" & CStr(targetRow)
    If Not VWGX_WriteCellValueSafe(ws.Cells(targetRow, 1), cellValue, stageText, errNumber, errText) Then
        Err.Raise errNumber, "VWGX_AppendActivity", errText
    End If
    gXLWorkbookDirty = True

    stageText = "verify Activity Log row " & CStr(targetRow)
    dataRow = ws.Range("A" & targetRow & ":I" & targetRow).Value2
    For i = 0 To 8
        If Not VWGX_ValuesMatch(dataRow(1, i + 1), eventRecord(i)) Then
            Err.Raise vbObjectError + 1613, "VWGX_AppendActivity", _
                      "Activity Log read-back mismatch in column " & CStr(i + 1) & "."
        End If
    Next i

    stageText = "save Activity Log"
    SaveWorkbookIfNeeded
    VWGX_AppendActivity = True
    If lockAcquired Then gTrackerWriteBusy = False
    lockAcquired = False
    Exit Function

WriteFailed:
    errNumber = Err.Number
    errText = Err.Description
    Debug.Print "[VWG TRACKER] Activity log write failed at [" & stageText & "]: " & _
                errNumber & " - " & errText
    If lockAcquired Then gTrackerWriteBusy = False
    lockAcquired = False

    ' Excel automation failures must not discard the audit event. Preserve it
    ' outside Excel; the timer retries it later. During a flush, the caller keeps
    ' the uncommitted line for retry instead of recursively re-queuing it.
    If Not gFlushingActivityFallback Then
        VWGX_AppendActivity = VWGX_AppendActivityToFallback(eventRecord)
        If VWGX_AppendActivity Then
            Debug.Print "[VWG TRACKER] Activity event preserved in local fallback after Excel error " & _
                        errNumber & ": " & CStr(eventRecord(1))
        End If
    Else
        VWGX_AppendActivity = False
    End If
End Function

Private Function VWGX_WriteRangeSafe(ByVal targetRange As Object, _
                                     ByVal newValues As Variant, _
                                     ByVal stageText As String, _
                                     ByRef finalError As Long, _
                                     ByRef finalDescription As String) As Boolean
    Dim attempt As Long
    finalError = 0
    finalDescription = vbNullString

    For attempt = 1 To 3
        On Error Resume Next
        targetRange.Value2 = newValues
        finalError = Err.Number
        finalDescription = Err.Description
        Err.Clear
        On Error GoTo 0
        If finalError = 0 Then
            VWGX_WriteRangeSafe = True
            Exit Function
        End If
        If (finalError <> VWGX_RETRYABLE_EXCEL_ERROR And _
            finalError <> VWGX_RPC_E_CALL_REJECTED) Or attempt = 3 Then Exit For
        Debug.Print "[VWG TRACKER] Transient Excel range-write error " & finalError & _
                    " at [" & stageText & "]; retry " & CStr(attempt + 1) & " of 3."
        VWGX_Sleep 150
    Next attempt

    Debug.Print "[VWG TRACKER] Excel range write failed at [" & stageText & "]: " & _
                finalError & " - " & finalDescription
End Function

Private Function VWGX_AssignmentFallbackPath() As String
    VWGX_AssignmentFallbackPath = StateDirectory() & "\assignment_fallback.tsv"
End Function

Private Function VWGX_AssignmentFallbackProcessingPath() As String
    VWGX_AssignmentFallbackProcessingPath = StateDirectory() & "\assignment_fallback_processing.tsv"
End Function

Private Function VWGX_EncodeAssignmentField(ByVal value As Variant) As String
    Dim valueText As String
    Dim i As Long
    Dim codePoint As Long
    Dim encoded As String

    If IsNull(value) Then
        VWGX_EncodeAssignmentField = "N"
        Exit Function
    ElseIf IsEmpty(value) Then
        VWGX_EncodeAssignmentField = "E"
        Exit Function
    ElseIf VarType(value) = vbDate Then
        valueText = Format$(CDate(value), "yyyymmddhhnnss")
        VWGX_EncodeAssignmentField = "D" & valueText
        Exit Function
    ElseIf VarType(value) = vbBoolean Then
        VWGX_EncodeAssignmentField = IIf(CBool(value), "B1", "B0")
        Exit Function
    ElseIf IsNumeric(value) And VarType(value) <> vbString Then
        VWGX_EncodeAssignmentField = "V" & Trim$(Str$(CDbl(value)))
        Exit Function
    End If

    valueText = CStr(value)
    encoded = "S"
    For i = 1 To Len(valueText)
        codePoint = AscW(Mid$(valueText, i, 1))
        If codePoint < 0 Then codePoint = codePoint + 65536
        encoded = encoded & Right$("0000" & Hex$(codePoint), 4)
    Next i
    VWGX_EncodeAssignmentField = encoded
End Function

Private Function VWGX_DecodeAssignmentField(ByVal encoded As String) As Variant
    Dim kind As String
    Dim valueText As String
    Dim i As Long
    Dim codePoint As Long
    Dim decodedText As String

    If Len(encoded) = 0 Then Err.Raise vbObjectError + 1651, "VWGX_DecodeAssignmentField", "Blank encoded field."
    kind = Left$(encoded, 1)
    valueText = Mid$(encoded, 2)

    Select Case kind
        Case "N"
            VWGX_DecodeAssignmentField = Null
        Case "E"
            VWGX_DecodeAssignmentField = Empty
        Case "B"
            VWGX_DecodeAssignmentField = (valueText = "1")
        Case "V"
            VWGX_DecodeAssignmentField = CDbl(Val(valueText))
        Case "D"
            If Len(valueText) <> 14 Then Err.Raise vbObjectError + 1652, "VWGX_DecodeAssignmentField", "Invalid date token."
            VWGX_DecodeAssignmentField = DateSerial(CInt(Left$(valueText, 4)), _
                CInt(Mid$(valueText, 5, 2)), CInt(Mid$(valueText, 7, 2))) + _
                TimeSerial(CInt(Mid$(valueText, 9, 2)), CInt(Mid$(valueText, 11, 2)), CInt(Mid$(valueText, 13, 2)))
        Case "S"
            If (Len(valueText) Mod 4) <> 0 Then Err.Raise vbObjectError + 1653, "VWGX_DecodeAssignmentField", "Invalid text token length."
            For i = 1 To Len(valueText) Step 4
                codePoint = CLng("&H" & Mid$(valueText, i, 4))
                If codePoint > 32767 Then codePoint = codePoint - 65536
                decodedText = decodedText & ChrW$(codePoint)
            Next i
            VWGX_DecodeAssignmentField = decodedText
        Case Else
            Err.Raise vbObjectError + 1654, "VWGX_DecodeAssignmentField", "Unknown assignment field type."
    End Select
End Function

Private Function VWGX_AppendAssignmentToFallback(ByVal record As Variant) As Boolean
    Dim f As Integer
    Dim i As Long
    Dim lineText As String

    On Error GoTo FallbackFailed
    If Not IsArray(record) Then Exit Function
    If LBound(record) <> 0 Or UBound(record) < 20 Then Exit Function
    EnsureStateDirectory
    For i = 0 To 20
        If i > 0 Then lineText = lineText & vbTab
        lineText = lineText & VWGX_EncodeAssignmentField(record(i))
    Next i
    f = FreeFile
    Open VWGX_AssignmentFallbackPath() For Append As #f
    Print #f, lineText
    Close #f
    VWGX_AppendAssignmentToFallback = True
    Exit Function

FallbackFailed:
    On Error Resume Next
    Close #f
    Debug.Print "[VWG TRACKER] CRITICAL: Could not preserve assignment row in local retry queue: " & _
                Err.Number & " - " & Err.Description
    VWGX_AppendAssignmentToFallback = False
End Function

Private Function VWGX_DecodeAssignmentFallbackLine(ByVal lineText As String, ByRef record As Variant) As Boolean
    Dim fields As Variant
    Dim i As Long
    On Error GoTo DecodeFailed
    fields = Split(lineText, vbTab)
    If UBound(fields) <> 20 Then Exit Function
    ReDim record(0 To 20)
    For i = 0 To 20
        record(i) = VWGX_DecodeAssignmentField(CStr(fields(i)))
    Next i
    If IsNull(record(0)) Or IsEmpty(record(0)) Then Exit Function
    If Len(Trim$(CStr(record(0)))) = 0 Then Exit Function
    VWGX_DecodeAssignmentFallbackLine = True
    Exit Function
DecodeFailed:
    VWGX_DecodeAssignmentFallbackLine = False
End Function

' Returns True only when both the processing queue and newly-arrived queue are clear.
' recordsFlushed counts committed rows so the dispatcher knows when to reload memory.
Public Function VWGX_FlushAssignmentFallback(ByRef recordsFlushed As Long) As Boolean
    Dim mainPath As String
    Dim processingPath As String
    Dim f As Integer
    Dim lineText As String
    Dim pendingLines As Collection
    Dim remainingLines As Collection
    Dim record As Variant
    Dim item As Variant
    Dim parsed As Boolean

    recordsFlushed = 0
    If gTrackerWriteBusy Or gFlushingAssignmentFallback Then Exit Function
    On Error GoTo FlushFailed
    EnsureStateDirectory
    mainPath = VWGX_AssignmentFallbackPath()
    processingPath = VWGX_AssignmentFallbackProcessingPath()

    If Len(Dir$(processingPath)) = 0 Then
        If Len(Dir$(mainPath)) = 0 Then
            VWGX_FlushAssignmentFallback = True
            Exit Function
        End If
        Name mainPath As processingPath
    End If

    Set pendingLines = New Collection
    Set remainingLines = New Collection
    f = FreeFile
    Open processingPath For Input As #f
    Do While Not EOF(f)
        Line Input #f, lineText
        If Len(lineText) > 0 Then pendingLines.Add lineText
    Loop
    Close #f

    gFlushingAssignmentFallback = True
    For Each item In pendingLines
        parsed = VWGX_DecodeAssignmentFallbackLine(CStr(item), record)
        If parsed Then
            If VWGX_UpsertAssignment(record) Then
                recordsFlushed = recordsFlushed + 1
            Else
                remainingLines.Add CStr(item)
            End If
        Else
            Debug.Print "[VWG TRACKER] Invalid assignment retry record retained for inspection."
            remainingLines.Add CStr(item)
        End If
    Next item
    gFlushingAssignmentFallback = False

    If Len(Dir$(processingPath)) > 0 Then Kill processingPath
    If remainingLines.Count > 0 Then
        f = FreeFile
        Open processingPath For Output As #f
        For Each item In remainingLines
            Print #f, CStr(item)
        Next item
        Close #f
        Debug.Print "[VWG TRACKER] Assignment retry queue still has " & CStr(remainingLines.Count) & " row(s)."
    End If

    VWGX_FlushAssignmentFallback = _
        (remainingLines.Count = 0 And Len(Dir$(mainPath)) = 0 And Len(Dir$(processingPath)) = 0)
    If recordsFlushed > 0 Then
        Debug.Print "[VWG TRACKER] Committed " & CStr(recordsFlushed) & " deferred assignment row(s) from the local retry queue."
    End If
    Exit Function

FlushFailed:
    gFlushingAssignmentFallback = False
    On Error Resume Next
    Close #f
    Debug.Print "[VWG TRACKER] Assignment retry queue flush failed: " & Err.Number & " - " & Err.Description
    VWGX_FlushAssignmentFallback = False
End Function

Public Function VWGX_AssignmentFallbackPending() As Boolean
    On Error Resume Next
    VWGX_AssignmentFallbackPending = _
        (Len(Dir$(VWGX_AssignmentFallbackPath())) > 0 Or Len(Dir$(VWGX_AssignmentFallbackProcessingPath())) > 0)
    On Error GoTo 0
End Function

Public Function VWGX_AssignmentFallbackContainsKey(ByVal keyText As String) As Boolean
    Dim paths As Variant
    Dim item As Variant
    Dim f As Integer
    Dim lineText As String
    Dim fields As Variant
    Dim decoded As Variant

    On Error GoTo CheckFailed
    If Len(keyText) = 0 Then Exit Function
    paths = Array(VWGX_AssignmentFallbackProcessingPath(), VWGX_AssignmentFallbackPath())
    For Each item In paths
        If Len(Dir$(CStr(item))) > 0 Then
            f = FreeFile
            Open CStr(item) For Input As #f
            Do While Not EOF(f)
                Line Input #f, lineText
                fields = Split(lineText, vbTab)
                If UBound(fields) = 20 Then
                    decoded = VWGX_DecodeAssignmentField(CStr(fields(0)))
                    If Not IsNull(decoded) And Not IsEmpty(decoded) Then
                        If StrComp(Trim$(CStr(decoded)), keyText, vbTextCompare) = 0 Then
                            Close #f
                            VWGX_AssignmentFallbackContainsKey = True
                            Exit Function
                        End If
                    End If
                End If
            Loop
            Close #f
        End If
    Next item
    Exit Function
CheckFailed:
    On Error Resume Next
    Close #f
End Function

Private Function VWGX_ActivityFallbackPath() As String
    VWGX_ActivityFallbackPath = StateDirectory() & "\activity_fallback.tsv"
End Function

Private Function VWGX_ActivityFallbackProcessingPath() As String
    VWGX_ActivityFallbackProcessingPath = StateDirectory() & "\activity_fallback_processing.tsv"
End Function

Private Function VWGX_FallbackField(ByVal value As Variant) As String
    Dim valueText As String
    If IsNull(value) Or IsEmpty(value) Then
        valueText = vbNullString
    ElseIf IsDate(value) Then
        valueText = Format$(CDate(value), "yyyy-mm-dd hh:nn:ss")
    Else
        valueText = CStr(value)
    End If
    valueText = Replace(valueText, vbTab, " ")
    valueText = Replace(valueText, vbCr, " ")
    valueText = Replace(valueText, vbLf, " ")
    VWGX_FallbackField = valueText
End Function

Private Function VWGX_AppendActivityToFallback(ByVal eventRecord As Variant) As Boolean
    Dim f As Integer
    Dim i As Long
    Dim lineText As String
    On Error GoTo FallbackFailed
    If Not IsArray(eventRecord) Then Exit Function
    If UBound(eventRecord) < 8 Then Exit Function
    EnsureStateDirectory
    For i = 0 To 8
        If i > 0 Then lineText = lineText & vbTab
        lineText = lineText & VWGX_FallbackField(eventRecord(i))
    Next i
    f = FreeFile
    Open VWGX_ActivityFallbackPath() For Append As #f
    Print #f, lineText
    Close #f
    VWGX_AppendActivityToFallback = True
    Exit Function
FallbackFailed:
    On Error Resume Next
    Close #f
    Debug.Print "[VWG TRACKER] CRITICAL: Could not preserve Activity Log event in fallback file: " & _
                Err.Number & " - " & Err.Description
    VWGX_AppendActivityToFallback = False
End Function

Public Sub VWGX_FlushActivityFallback()
    Dim processingPath As String
    Dim mainPath As String
    Dim f As Integer
    Dim lineText As String
    Dim fields As Variant
    Dim eventRecord(0 To 8) As Variant
    Dim pendingLines As Collection
    Dim remainingLines As Collection
    Dim item As Variant
    Dim i As Long
    Dim parseOK As Boolean

    On Error GoTo FlushFailed
    If gTrackerWriteBusy Or gFlushingActivityFallback Then Exit Sub
    EnsureStateDirectory
    mainPath = VWGX_ActivityFallbackPath()
    processingPath = VWGX_ActivityFallbackProcessingPath()

    If Len(Dir$(processingPath)) = 0 Then
        If Len(Dir$(mainPath)) = 0 Then Exit Sub
        Name mainPath As processingPath
    End If

    Set pendingLines = New Collection
    Set remainingLines = New Collection
    f = FreeFile
    Open processingPath For Input As #f
    Do While Not EOF(f)
        Line Input #f, lineText
        If Len(lineText) > 0 Then pendingLines.Add lineText
    Loop
    Close #f

    gFlushingActivityFallback = True
    For Each item In pendingLines
        fields = Split(CStr(item), vbTab)
        parseOK = (UBound(fields) = 8 And IsDate(fields(0)))
        If parseOK Then
            eventRecord(0) = CDate(fields(0))
            For i = 1 To 8
                eventRecord(i) = CStr(fields(i))
            Next i
            If Not VWGX_AppendActivity(eventRecord) Then remainingLines.Add CStr(item)
        Else
            remainingLines.Add CStr(item)
        End If
    Next item
    gFlushingActivityFallback = False

    If Len(Dir$(processingPath)) > 0 Then Kill processingPath
    If remainingLines.Count > 0 Then
        f = FreeFile
        Open processingPath For Output As #f
        For Each item In remainingLines
            Print #f, CStr(item)
        Next item
        Close #f
        Debug.Print "[VWG TRACKER] Activity fallback still has " & CStr(remainingLines.Count) & " event(s) awaiting Excel."
    End If
    Exit Sub

FlushFailed:
    gFlushingActivityFallback = False
    On Error Resume Next
    Close #f
    Debug.Print "[VWG TRACKER] Activity fallback flush failed: " & Err.Number & " - " & Err.Description
End Sub

Public Sub VWGX_BackfillMissingAutoAssignmentEvents_Last7Days()
    Dim wsAssignments As Object
    Dim wsActivity As Object
    Dim lastAssignments As Long
    Dim lastActivity As Long
    Dim assignmentData As Variant
    Dim activityData As Variant
    Dim existing As Object
    Dim eventRecord(0 To 8) As Variant
    Dim r As Long
    Dim keyText As String
    Dim originalName As String
    Dim statusText As String
    Dim categoryState As String
    Dim allocatedAt As Variant
    Dim allocatedDate As Date
    Dim backfilled As Long
    Dim failed As Long

    On Error GoTo BackfillFailed
    VWGX_FlushActivityFallback
    If Not VWGX_EnsureConnected(True) Then Exit Sub

    Set wsAssignments = gXLWorkbook.Worksheets("Assignments")
    Set wsActivity = gXLWorkbook.Worksheets("Activity Log")
    lastAssignments = wsAssignments.Cells(wsAssignments.Rows.Count, 1).End(XL_UP).Row
    lastActivity = wsActivity.Cells(wsActivity.Rows.Count, 3).End(XL_UP).Row
    If lastAssignments < FIRST_DATA_ROW Then
        MsgBox "No assignment rows are available to backfill.", vbInformation, "VWG Activity Log Recovery"
        Exit Sub
    End If

    assignmentData = wsAssignments.Range("A" & FIRST_DATA_ROW & ":T" & lastAssignments).Value2
    Set existing = CreateObject("Scripting.Dictionary")
    existing.CompareMode = vbTextCompare
    If lastActivity >= FIRST_DATA_ROW Then
        activityData = wsActivity.Range("B" & FIRST_DATA_ROW & ":C" & lastActivity).Value2
        For r = LBound(activityData, 1) To UBound(activityData, 1)
            If StrComp(Trim$(CStr(activityData(r, 1))), "AUTO_ASSIGNMENT", vbTextCompare) = 0 Then
                keyText = Trim$(CStr(activityData(r, 2)))
                If Len(keyText) > 0 Then existing(keyText) = True
            End If
        Next r
    End If

    For r = LBound(assignmentData, 1) To UBound(assignmentData, 1)
        keyText = Trim$(CStr(assignmentData(r, 1)))
        originalName = Trim$(CStr(assignmentData(r, 5)))
        allocatedAt = assignmentData(r, 15)
        statusText = UCase$(Trim$(CStr(assignmentData(r, 10))))
        categoryState = CStr(assignmentData(r, 20))
        If Len(keyText) > 0 And Len(originalName) > 0 And VWGX_TryGetDateTime(allocatedAt, allocatedDate) Then
            If DateValue(allocatedDate) >= Date - 6 And DateValue(allocatedDate) <= Date Then
                If statusText <> "ASSIGNING" And statusText <> "QUEUED" And statusText <> "UNASSIGNED" And _
                   Not (statusText = "NEEDS REVIEW" And InStr(1, categoryState, "Assignment category could not be confirmed", vbTextCompare) > 0) Then
                    If Not existing.Exists(keyText) Then
                        eventRecord(0) = allocatedDate
                        eventRecord(1) = "AUTO_ASSIGNMENT"
                        eventRecord(2) = keyText
                        eventRecord(3) = CStr(assignmentData(r, 3))
                        eventRecord(4) = originalName
                        eventRecord(5) = CStr(assignmentData(r, 6))
                        eventRecord(6) = CStr(assignmentData(r, 7))
                        eventRecord(7) = CLng(Val(CStr(assignmentData(r, 8))))
                        eventRecord(8) = "Backfilled from Assignments register; the original Activity Log event was missing. Subject=" & CStr(assignmentData(r, 4))
                        If VWGX_AppendActivity(eventRecord) Then
                            existing(keyText) = True
                            backfilled = backfilled + 1
                        Else
                            failed = failed + 1
                        End If
                    End If
                End If
            End If
        End If
    Next r

    VWGX_RefreshSummary Date
    MsgBox "Seven-day auto-assignment audit recovery finished." & vbCrLf & _
           "Events backfilled or safely queued: " & CStr(backfilled) & vbCrLf & _
           "Events still not preserved: " & CStr(failed) & vbCrLf & vbCrLf & _
           "No Outlook categories or assignment owners were changed.", _
           vbInformation, "VWG Activity Log Recovery"
    Exit Sub

BackfillFailed:
    Debug.Print "[VWG TRACKER] Seven-day Activity Log backfill failed: " & Err.Number & " - " & Err.Description
    MsgBox "Activity Log backfill failed: " & Err.Number & " - " & Err.Description, _
           vbExclamation, "VWG Activity Log Recovery"
End Sub

Public Sub VWGX_DiagnoseDailyAllocationCounts()
    Dim counts As Object
    Dim workloads As Object
    Dim roster As Variant
    Dim r As Long
    Dim nm As String
    Dim shiftText As String
    Dim activeText As String
    Dim countValue As Long
    Dim workloadValue As Long

    On Error GoTo DiagnoseFailed
    If Not VWGX_EnsureConnected(True) Then Exit Sub
    Set counts = VWGX_GetAutoAllocationCountsForDate(Date)
    Set workloads = VWGX_GetCurrentWorkloadCounts()
    roster = VWGX_ReadRoster()
    If counts Is Nothing Or workloads Is Nothing Or Not IsArray(roster) Then
        Err.Raise vbObjectError + 1641, "VWGX_DiagnoseDailyAllocationCounts", _
                  "Could not read today's allocation counts, current workloads, or roster."
    End If

    Debug.Print "========== VWG DAILY FAIRNESS DIAGNOSTIC =========="
    Debug.Print "Date: " & Format$(Date, "dd/mmm/yyyy") & " (Egypt staffed hours use the same calendar date)"
    Debug.Print "Selection: lowest today's auto allocations -> lowest current workload -> rotating tie-break"
    Debug.Print "---------------------------------------------------"
    For r = LBound(roster, 1) To UBound(roster, 1)
        nm = Trim$(CStr(roster(r, 1)))
        shiftText = Trim$(CStr(roster(r, 3)))
        activeText = UCase$(Trim$(CStr(roster(r, 6))))
        If Len(nm) > 0 Then
            countValue = DictCountSafe(counts, nm)
            workloadValue = DictCountSafe(workloads, nm)
            Debug.Print nm & " | Shift=" & shiftText & " | ActiveToday=" & activeText & _
                        " | Today'sAutoAllocations=" & CStr(countValue) & _
                        " | CurrentWorkload=" & CStr(workloadValue)
        End If
    Next r
    Debug.Print "==================================================="
    MsgBox "Daily fairness diagnostic written to the VBA Immediate Window (Ctrl+G)." & vbCrLf & _
           "It shows today's successful automatic allocations, active status, shift group, and current workload for each colleague.", _
           vbInformation, "VWG Daily Fairness Diagnostic"
    Exit Sub
DiagnoseFailed:
    Debug.Print "[VWG TRACKER] Daily fairness diagnostic failed: " & Err.Number & " - " & Err.Description
    MsgBox "Daily fairness diagnostic failed: " & Err.Number & " - " & Err.Description, _
           vbExclamation, "VWG Daily Fairness Diagnostic"
End Sub

Private Function DictCountSafe(ByVal dict As Object, ByVal keyText As String) As Long
    On Error Resume Next
    If Not dict Is Nothing Then
        If dict.Exists(keyText) Then DictCountSafe = CLng(dict(keyText))
    End If
    On Error GoTo 0
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

Private Function VWGX_TryGetDateTime(ByVal value As Variant, ByRef parsedDate As Date) As Boolean
    Dim serialValue As Double
    On Error GoTo ParseFailed
    If IsNull(value) Or IsEmpty(value) Then Exit Function

    ' Excel Range.Value2 returns date-formatted cells as Double serials.
    ' Convert explicitly using Excel's 1900 date system and preserve time.
    If IsNumeric(value) Then
        serialValue = CDbl(value)
        If serialValue <= 0 Then Exit Function
        parsedDate = DateSerial(1899, 12, 30) + serialValue
        VWGX_TryGetDateTime = True
    ElseIf IsDate(value) Then
        parsedDate = CDate(value)
        VWGX_TryGetDateTime = True
    Else
        VWGX_TryGetDateTime = VWGX_ParseFallbackDate(CStr(value), parsedDate)
    End If
    Exit Function
ParseFailed:
    VWGX_TryGetDateTime = False
End Function

Private Function VWGX_ValueIsOnDate(ByVal value As Variant, ByVal targetDate As Date) As Boolean
    Dim parsedDate As Date
    If VWGX_TryGetDateTime(value, parsedDate) Then
        VWGX_ValueIsOnDate = (DateValue(parsedDate) = DateValue(targetDate))
    End If
End Function

Private Function VWGX_ParseFallbackDate(ByVal valueText As String, ByRef parsedDate As Date) As Boolean
    Dim y As Long, m As Long, d As Long
    Dim hh As Long, nn As Long, ss As Long
    On Error GoTo ParseFailed
    valueText = Trim$(valueText)
    If Len(valueText) >= 10 And Mid$(valueText, 5, 1) = "-" And Mid$(valueText, 8, 1) = "-" Then
        y = CLng(Left$(valueText, 4))
        m = CLng(Mid$(valueText, 6, 2))
        d = CLng(Mid$(valueText, 9, 2))
        parsedDate = DateSerial(y, m, d)
        If Len(valueText) >= 19 Then
            hh = CLng(Mid$(valueText, 12, 2))
            nn = CLng(Mid$(valueText, 15, 2))
            ss = CLng(Mid$(valueText, 18, 2))
            parsedDate = parsedDate + TimeSerial(hh, nn, ss)
        End If
        VWGX_ParseFallbackDate = True
    ElseIf IsDate(valueText) Then
        parsedDate = CDate(valueText)
        VWGX_ParseFallbackDate = True
    End If
    Exit Function
ParseFailed:
    VWGX_ParseFallbackDate = False
End Function

Private Function VWGX_CountableAutoAssignmentStatus(ByVal statusText As String) As Boolean
    Select Case UCase$(Trim$(statusText))
        Case "ASSIGNING", "ASSIGNED", "HANDED OVER", "HANDOVER PENDING", "RECOVERED", _
             "RECOVERED - REVIEW", "NEEDS REVIEW", "RECOVERY FAILED", _
             "CLOSED", "COMPLETED", "CANCELLED"
            VWGX_CountableAutoAssignmentStatus = True
    End Select
End Function

Public Function VWGX_GetAutoAllocationCountsForDate(ByVal allocationDate As Date) As Object
    Dim counts As Object
    Dim seenKeys As Object
    Dim ws As Object
    Dim lastRow As Long
    Dim data As Variant
    Dim activityData As Variant
    Dim activityWs As Object
    Dim activityLastRow As Long
    Dim activityStartRow As Long
    Dim r As Long
    Dim nm As String
    Dim itemKey As String
    Dim statusText As String
    Dim allocatedAt As Variant
    Dim categoryState As String
    Dim eventDate As Variant
    Dim eventType As String
    Dim f As Integer
    Dim fallbackLine As String
    Dim fallbackPath As Variant
    Dim fields As Variant
    Dim fallbackDate As Date
    Dim fallbackCount As Long

    Set counts = CreateObject("Scripting.Dictionary")
    counts.CompareMode = vbTextCompare
    Set seenKeys = CreateObject("Scripting.Dictionary")
    seenKeys.CompareMode = vbTextCompare

    On Error GoTo ReadFailed
    If Not VWGX_EnsureConnected(False) Then
        Set VWGX_GetAutoAllocationCountsForDate = Nothing
        Exit Function
    End If

    ' Primary source: one row per email in Assignments. Column O is a date-time
    ' serial in Excel; compare its calendar date, not a locale-sensitive string.
    Set ws = gXLWorkbook.Worksheets("Assignments")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(XL_UP).Row
    If lastRow >= FIRST_DATA_ROW Then
        data = ws.Range("A" & FIRST_DATA_ROW & ":T" & lastRow).Value2
        For r = LBound(data, 1) To UBound(data, 1)
            itemKey = Trim$(CStr(data(r, 1)))
            nm = Trim$(CStr(data(r, 5)))
            statusText = UCase$(Trim$(CStr(data(r, 10))))
            allocatedAt = data(r, 15)
            categoryState = CStr(data(r, 20))
            If Len(itemKey) > 0 And Len(nm) > 0 And VWGX_CountableAutoAssignmentStatus(statusText) And _
               Not (statusText = "NEEDS REVIEW" And InStr(1, categoryState, "Assignment category could not be confirmed", vbTextCompare) > 0) Then
                If VWGX_ValueIsOnDate(allocatedAt, allocationDate) Then
                    If Not counts.Exists(nm) Then counts.Add nm, 0
                    If Not seenKeys.Exists(itemKey) Then
                        seenKeys.Add itemKey, True
                        counts(nm) = CLng(counts(nm)) + 1
                    End If
                End If
            End If
        Next r
    End If

    ' Compatibility source: older successful assignments may have an empty O
    ' timestamp. An AUTO_ASSIGNMENT audit event supplies the original allocation
    ' date and owner; the Email Key prevents double-counting rows already found.
    Set activityWs = gXLWorkbook.Worksheets("Activity Log")
    activityLastRow = activityWs.Cells(activityWs.Rows.Count, 1).End(XL_UP).Row
    activityStartRow = FIRST_DATA_ROW
    If activityLastRow - FIRST_DATA_ROW + 1 > 2000 Then activityStartRow = activityLastRow - 1999
    If activityLastRow >= FIRST_DATA_ROW Then
        activityData = activityWs.Range("A" & activityStartRow & ":I" & activityLastRow).Value2
        For r = LBound(activityData, 1) To UBound(activityData, 1)
            eventType = UCase$(Trim$(CStr(activityData(r, 2))))
            If eventType = "AUTO_ASSIGNMENT" Then
                itemKey = Trim$(CStr(activityData(r, 3)))
                nm = Trim$(CStr(activityData(r, 5)))
                eventDate = activityData(r, 1)
                If Len(itemKey) > 0 And Len(nm) > 0 Then
                    If VWGX_ValueIsOnDate(eventDate, allocationDate) Then
                        If Not counts.Exists(nm) Then counts.Add nm, 0
                        If Not seenKeys.Exists(itemKey) Then
                            seenKeys.Add itemKey, True
                            counts(nm) = CLng(counts(nm)) + 1
                        End If
                    End If
                End If
            End If
        Next r
    End If

    ' Final compatibility source: an audit event may still be waiting in the
    ' local fallback file because Excel was busy when it was first written.
    For Each fallbackPath In Array(VWGX_ActivityFallbackPath(), VWGX_ActivityFallbackProcessingPath())
        If Len(Dir$(CStr(fallbackPath))) > 0 Then
            f = FreeFile
            Open CStr(fallbackPath) For Input As #f
            Do While Not EOF(f)
                Line Input #f, fallbackLine
                fields = Split(fallbackLine, vbTab)
                If UBound(fields) = 8 Then
                    If UCase$(Trim$(CStr(fields(1)))) = "AUTO_ASSIGNMENT" Then
                        itemKey = Trim$(CStr(fields(2)))
                        nm = Trim$(CStr(fields(4)))
                        If Len(itemKey) > 0 And Len(nm) > 0 Then
                            If VWGX_ParseFallbackDate(CStr(fields(0)), fallbackDate) Then
                                If DateValue(fallbackDate) = DateValue(allocationDate) Then
                                    If Not counts.Exists(nm) Then counts.Add nm, 0
                                    If Not seenKeys.Exists(itemKey) Then
                                        seenKeys.Add itemKey, True
                                        counts(nm) = CLng(counts(nm)) + 1
                                        fallbackCount = fallbackCount + 1
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Loop
            Close #f
        End If
    Next fallbackPath

    If fallbackCount > 0 Then
        Debug.Print "[VWG FAIRNESS] Included " & CStr(fallbackCount) & _
                    " successful AUTO_ASSIGNMENT event(s) still in local fallback."
    End If

    Set VWGX_GetAutoAllocationCountsForDate = counts
    Exit Function

ReadFailed:
    On Error Resume Next
    If f <> 0 Then Close #f
    On Error GoTo 0
    Debug.Print "[VWG TRACKER] Daily allocation count read failed: " & Err.Number & " - " & Err.Description
    Set VWGX_GetAutoAllocationCountsForDate = Nothing
End Function

' Compatibility entry point retained for older macros. The dispatcher now uses
' the explicit date-based function above; this routine reports today's count.
Public Function VWGX_GetAutoAllocationCounts(ByVal periodStart As Date) As Object
    Set VWGX_GetAutoAllocationCounts = VWGX_GetAutoAllocationCountsForDate(Date)
End Function

Public Function VWGX_GetCurrentWorkloadCounts() As Object
    Dim counts As Object
    Dim ws As Object
    Dim lastRow As Long
    Dim data As Variant
    Dim r As Long
    Dim nm As String
    Dim statusText As String
    Dim categoryState As String

    Set counts = CreateObject("Scripting.Dictionary")
    counts.CompareMode = vbTextCompare

    On Error GoTo ReadFailed
    If Not VWGX_EnsureConnected(False) Then
        Set VWGX_GetCurrentWorkloadCounts = Nothing
        Exit Function
    End If

    Set ws = gXLWorkbook.Worksheets("Assignments")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(XL_UP).Row
    If lastRow < FIRST_DATA_ROW Then GoTo Done

    data = ws.Range("A" & FIRST_DATA_ROW & ":T" & lastRow).Value2
    For r = LBound(data, 1) To UBound(data, 1)
        nm = Trim$(CStr(data(r, 6)))
        statusText = UCase$(Trim$(CStr(data(r, 10))))
        categoryState = CStr(data(r, 20))
        If Len(nm) > 0 And IsActiveStatus(statusText) And _
           Not (statusText = "NEEDS REVIEW" And InStr(1, categoryState, "Assignment category could not be confirmed", vbTextCompare) > 0) Then
            If counts.Exists(nm) Then
                counts(nm) = CLng(counts(nm)) + 1
            Else
                counts.Add nm, 1
            End If
        End If
    Next r

Done:
    Set VWGX_GetCurrentWorkloadCounts = counts
    Exit Function

ReadFailed:
    Debug.Print "[VWG TRACKER] Workload count read failed: " & Err.Number & " - " & Err.Description
    Set VWGX_GetCurrentWorkloadCounts = Nothing
End Function

Private Function IsActiveStatus(ByVal statusText As String) As Boolean
    Select Case UCase$(Trim$(statusText))
        Case "ASSIGNING", "ASSIGNED", "HANDED OVER", "HANDOVER PENDING", "RECOVERED", "RECOVERED - REVIEW", "NEEDS REVIEW", "RECOVERY FAILED"
            IsActiveStatus = True
        Case Else
            IsActiveStatus = False
    End Select
End Function

Public Sub VWGX_RefreshSummary(Optional ByVal allocationDate As Variant)
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
    Dim categoryState As String
    Dim countDate As Date
    Dim stageText As String
    Dim errNumber As Long
    Dim errText As String
    Dim cellWriteOK As Boolean

    On Error GoTo RefreshFailed

    stageText = "connect to workbook"
    If Not VWGX_EnsureConnected(False) Then Exit Sub
    If gTrackerWriteBusy Then
        gDashboardRefreshPending = True
        If IsDate(allocationDate) Then gDashboardRefreshDate = DateValue(CDate(allocationDate))
        Debug.Print "[VWG TRACKER] Dashboard refresh deferred: another tracker write is active."
        Exit Sub
    End If
    gTrackerWriteBusy = True

    stageText = "read same-day allocation counts"
    If IsDate(allocationDate) Then
        countDate = DateValue(CDate(allocationDate))
    ElseIf gDashboardRefreshPending And IsDate(gDashboardRefreshDate) Then
        countDate = DateValue(CDate(gDashboardRefreshDate))
    Else
        countDate = Date
    End If
    Set allocCounts = VWGX_GetAutoAllocationCountsForDate(countDate)
    Set workloadCounts = VWGX_GetCurrentWorkloadCounts()
    If allocCounts Is Nothing Or workloadCounts Is Nothing Then
        Err.Raise vbObjectError + 1640, "VWGX_RefreshSummary", _
                  "Daily fairness counts could not be read. Existing count cells were left unchanged."
    End If

    stageText = "get Team & Shifts worksheet"
    Set wsTeam = gXLWorkbook.Worksheets("Team & Shifts")
    stageText = "label the daily roster and its operating instructions"
    cellWriteOK = VWGX_WriteCellValueSafe(wsTeam.Cells(1, 1), "TEAM & SHIFTS  |  DAILY ROSTER", stageText, errNumber, errText)
    If cellWriteOK Then gXLWorkbookDirty = True
    cellWriteOK = VWGX_WriteCellValueSafe(wsTeam.Cells(2, 1), _
        "Set Active Today to Yes or No for each colleague every day. Today's allocation counts reset automatically at Egypt-local midnight; historical assignments are retained.", _
        stageText, errNumber, errText)
    If cellWriteOK Then gXLWorkbookDirty = True

    stageText = "label daily availability and fairness columns"
    cellWriteOK = VWGX_WriteCellValueSafe(wsTeam.Cells(HEADER_ROW, 6), "Active Today (Yes/No)", stageText, errNumber, errText)
    If cellWriteOK Then gXLWorkbookDirty = True
    cellWriteOK = VWGX_WriteCellValueSafe(wsTeam.Cells(HEADER_ROW, 7), "Today's Auto Allocations", stageText, errNumber, errText)
    If cellWriteOK Then gXLWorkbookDirty = True
    cellWriteOK = VWGX_WriteCellValueSafe(wsTeam.Cells(HEADER_ROW, 8), "Current Workload", stageText, errNumber, errText)
    If cellWriteOK Then gXLWorkbookDirty = True
    stageText = "read roster"
    roster = VWGX_ReadRoster()

    If IsArray(roster) Then
        For r = LBound(roster, 1) To UBound(roster, 1)
            nameText = Trim$(CStr(roster(r, 1)))
            targetRow = r + FIRST_DATA_ROW - 1
            If Len(nameText) > 0 Then
                stageText = "write Team & Shifts allocation count, row " & CStr(targetRow) & _
                            ", column G, colleague " & nameText
                cellWriteOK = VWGX_WriteCellValueSafe(wsTeam.Cells(targetRow, 7), _
                                DictCountSafe(allocCounts, nameText), stageText, errNumber, errText)
                If cellWriteOK Then gXLWorkbookDirty = True

                stageText = "write Team & Shifts workload count, row " & CStr(targetRow) & _
                            ", column H, colleague " & nameText
                cellWriteOK = VWGX_WriteCellValueSafe(wsTeam.Cells(targetRow, 8), _
                                DictCountSafe(workloadCounts, nameText), stageText, errNumber, errText)
                If cellWriteOK Then gXLWorkbookDirty = True
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
        data = ws.Range("A" & FIRST_DATA_ROW & ":T" & lastRow).Value2
        For r = LBound(data, 1) To UBound(data, 1)
            If Len(Trim$(CStr(data(r, 1)))) > 0 Then assignedCount = assignedCount + 1
            categoryState = CStr(data(r, 20))
            If IsActiveStatus(UCase$(Trim$(CStr(data(r, 10))))) And _
               Not (UCase$(Trim$(CStr(data(r, 10)))) = "NEEDS REVIEW" And _
                    InStr(1, categoryState, "Assignment category could not be confirmed", vbTextCompare) > 0) Then
                activeCount = activeCount + 1
            End If
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
                Case "CATEGORY_REMOVAL_DETECTED", "RECOVERY_FAILED"
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

    stageText = "save dashboard summary"
    SaveWorkbookIfNeeded

    gDashboardRefreshPending = False
    gDashboardRefreshDate = Empty
    gTrackerWriteBusy = False
    Exit Sub

RefreshFailed:
    errNumber = Err.Number
    errText = Err.Description
    Debug.Print "[VWG TRACKER] Dashboard refresh failed at [" & stageText & "]: " & _
                errNumber & " - " & errText
    gDashboardRefreshPending = True
    If IsDate(allocationDate) Then gDashboardRefreshDate = DateValue(CDate(allocationDate))
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
    Debug.Print "Write policy: shared workbook lock; 3-attempt transient COM retry"
    Debug.Print "Assignment retry queue: durable local fallback; Outlook category waits for confirmed reservation save"
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
    Dim blankA As Boolean
    Dim blankB As Boolean
    Dim textValue As String

    On Error GoTo NoMatch

    If IsError(valueA) Or IsError(valueB) Then Exit Function

    ' Excel / fallback serialization can represent a blank as Null, Empty or "".
    ' Treat those blank forms as equivalent; do not trim nonblank IDs.
    blankA = IsNull(valueA) Or IsEmpty(valueA) Or _
             (VarType(valueA) = vbString And Len(CStr(valueA)) = 0)
    blankB = IsNull(valueB) Or IsEmpty(valueB) Or _
             (VarType(valueB) = vbString And Len(CStr(valueB)) = 0)
    If blankA Or blankB Then
        VWGX_ValuesMatch = (blankA And blankB)
        Exit Function
    End If

    ' Excel can coerce the text values "TRUE" / "FALSE" to Boolean values.
    ' Accept that equivalent representation while keeping all other text exact.
    If VarType(valueA) = vbBoolean And VarType(valueB) = vbString Then
        textValue = UCase$(Trim$(CStr(valueB)))
        If textValue = "TRUE" Then
            VWGX_ValuesMatch = CBool(valueA)
        ElseIf textValue = "FALSE" Then
            VWGX_ValuesMatch = Not CBool(valueA)
        End If
        Exit Function
    ElseIf VarType(valueB) = vbBoolean And VarType(valueA) = vbString Then
        textValue = UCase$(Trim$(CStr(valueA)))
        If textValue = "TRUE" Then
            VWGX_ValuesMatch = CBool(valueB)
        ElseIf textValue = "FALSE" Then
            VWGX_ValuesMatch = Not CBool(valueB)
        End If
        Exit Function
    End If

    ' Preserve exact text for keys/IDs (including leading zeroes and case).
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
' OPTIONAL DIAGNOSTICS — SAFE, NO ASSIGNMENT OR OUTLOOK CHANGES
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
    ws.Range("A1").Value2 = beforeValue
    Debug.Print "Cell write: SUCCESS"
    afterValue = ws.Range("A1").Value2

    If VWGX_ValuesMatch(beforeValue, afterValue) Then
        Debug.Print "Read-back: SUCCESS"
    Else
        Debug.Print "Read-back: FAILED — value changed during test."
    End If
    Exit Sub

TestFailed:
    Debug.Print "Tracker write test failed: " & Err.Number & " - " & Err.Description
End Sub
