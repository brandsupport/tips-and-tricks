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

Private gXLApp As Object
Private gXLWorkbook As Object
Private gWorkbookPath As String
Private gConnectionBusy As Boolean
Private gXLBatchDepth As Long
Private gXLWorkbookDirty As Boolean
Private gTrackerWriteBusy As Boolean
Private gLastActivityWriteError As String

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
    If gXLWorkbook Is Nothing Then Err.Raise vbObjectError + 1510, "SaveWorkbookNow", "Workbook is not connected."
    gXLWorkbook.Save
    gXLWorkbookDirty = False
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
    Dim stageText As String
    Dim errNumber As Long
    Dim errText As String
    Dim writeStarted As Boolean
    Dim isNewRow As Boolean
    Dim duplicateCount As Long
    Dim oldValues As Variant
    Dim keyValues As Variant
    Dim coreValues(1 To 1, 1 To 8) As Variant
    Dim tailValues(1 To 1, 1 To 11) As Variant
    Dim payloadIndex As Long
    Dim rollbackOK As Boolean
    Dim rollbackError As String

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

    stageText = "read existing assignment keys"
    If lastRow >= FIRST_DATA_ROW Then
        ' One bulk read replaces one Excel COM round-trip per historical row.
        keyValues = ws.Range("A" & FIRST_DATA_ROW & ":B" & lastRow).Value2
        For r = LBound(keyValues, 1) To UBound(keyValues, 1)
            If StrComp(Trim$(CStr(keyValues(r, 1))), keyText, vbTextCompare) = 0 Then
                If targetRow = 0 Then
                    targetRow = FIRST_DATA_ROW + r - 1
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

    ' The key is established first so a retry can locate the same row.
    ' Status (column J) is written last, preventing a new partial record from
    ' being counted as active before all its other fields have been stored.
    stageText = "format/write Email Key (column A) row " & CStr(targetRow)
    ws.Cells(targetRow, 1).NumberFormat = "@"
    ws.Cells(targetRow, 1).Value2 = keyText
    writeStarted = True
    gXLWorkbookDirty = True

    ' Existing rows already carry these formats. Apply the mixed formats once
    ' when creating a new row, not on every update to an existing email.
    If isNewRow Then
        stageText = "format new assignment row " & CStr(targetRow)
        ws.Cells(targetRow, 2).NumberFormat = "dd/mmm/yyyy hh:mm"
        ws.Range("C" & targetRow & ":G" & targetRow).NumberFormat = "@"
        ws.Cells(targetRow, 8).NumberFormat = "0"
        ws.Cells(targetRow, 9).NumberFormat = "dd/mmm/yyyy hh:mm"
        ws.Range("K" & targetRow & ":L" & targetRow).NumberFormat = "@"
        ws.Range("M" & targetRow & ":O" & targetRow).NumberFormat = "dd/mmm/yyyy hh:mm"
        ws.Cells(targetRow, 16).NumberFormat = "dd/mmm/yyyy"
        ws.Range("Q" & targetRow & ":R" & targetRow).NumberFormat = "@"
        ws.Cells(targetRow, 19).NumberFormat = "dd/mmm/yyyy hh:mm"
        ws.Cells(targetRow, 20).NumberFormat = "@"
        ws.Cells(targetRow, 21).NumberFormat = "0"
    End If

    For i = 1 To 8
        If IsNull(record(i)) Or IsEmpty(record(i)) Then
            coreValues(1, i) = Empty
        Else
            coreValues(1, i) = record(i)
        End If
    Next i

    For i = 10 To 20
        payloadIndex = i - 9
        If IsNull(record(i)) Or IsEmpty(record(i)) Then
            tailValues(1, payloadIndex) = Empty
        Else
            tailValues(1, payloadIndex) = record(i)
        End If
    Next i

    stageText = "bulk write assignment fields B:I row " & CStr(targetRow)
    ws.Range("B" & targetRow & ":I" & targetRow).Value2 = coreValues
    gXLWorkbookDirty = True

    stageText = "bulk write assignment fields K:U row " & CStr(targetRow)
    ws.Range("K" & targetRow & ":U" & targetRow).Value2 = tailValues
    gXLWorkbookDirty = True

    stageText = "write Status (column J) row " & CStr(targetRow)
    ws.Cells(targetRow, 10).NumberFormat = "@"
    If IsNull(record(9)) Or IsEmpty(record(9)) Then
        ws.Cells(targetRow, 10).ClearContents
    Else
        ws.Cells(targetRow, 10).Value2 = record(9)
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

    ' Roll back a failed row write with one range operation. If rollback itself
    ' fails, preserve the exact failure in the Immediate window for recovery.
    If writeStarted And targetRow >= FIRST_DATA_ROW Then
        rollbackOK = True
        On Error Resume Next
        Err.Clear
        If isNewRow Then
            ws.Range("A" & targetRow & ":U" & targetRow).ClearContents
        ElseIf IsArray(oldValues) Then
            ws.Range("A" & targetRow & ":U" & targetRow).Value2 = oldValues
        Else
            rollbackOK = False
        End If
        If Err.Number <> 0 Then
            rollbackOK = False
            rollbackError = "Row restore failed: " & Err.Number & " - " & Err.Description
        End If

        Err.Clear
        gXLWorkbookDirty = True
        SaveWorkbookIfNeeded
        If Err.Number <> 0 Then
            rollbackOK = False
            If Len(rollbackError) > 0 Then rollbackError = rollbackError & " | "
            rollbackError = rollbackError & "Rollback save failed: " & Err.Number & " - " & Err.Description
        End If
        On Error GoTo 0

        If rollbackOK Then
            Debug.Print "[VWG TRACKER] Assignment row rollback completed for row " & CStr(targetRow) & "."
        Else
            Debug.Print "[VWG TRACKER] CRITICAL: Assignment row rollback was incomplete at row " & _
                        CStr(targetRow) & ". " & rollbackError & _
                        " Keep dispatcher stopped and inspect this row before restart."
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
    Dim scanStart As Long
    Dim actualRow As Long
    Dim r As Long
    Dim i As Long
    Dim stageText As String
    Dim errNumber As Long
    Dim errText As String
    Dim rollbackError As String
    Dim writeStarted As Boolean
    Dim candidate As Boolean
    Dim allMatch As Boolean
    Dim existingValue As Variant
    Dim recentValues As Variant
    Dim writeValues(1 To 1, 1 To 8) As Variant
    Dim verifyValues As Variant

    On Error GoTo WriteFailed
    gLastActivityWriteError = vbNullString

    stageText = "connect to workbook"
    If Not VWGX_EnsureConnected(False) Then
        gLastActivityWriteError = "Could not connect to the tracker workbook."
        Exit Function
    End If

    If gTrackerWriteBusy Then
        gLastActivityWriteError = "Tracker write lock is busy; this activity event was not written to the worksheet."
        Debug.Print "[VWG TRACKER] Activity write deferred: " & gLastActivityWriteError
        Exit Function
    End If
    gTrackerWriteBusy = True

    stageText = "validate activity record"
    If Not IsArray(eventRecord) Then
        Err.Raise vbObjectError + 1611, "VWGX_AppendActivity", "Activity record is not an array."
    End If
    If LBound(eventRecord) <> 0 Or UBound(eventRecord) < 8 Then
        Err.Raise vbObjectError + 1612, "VWGX_AppendActivity", "Activity record must contain zero-based fields 0 to 8."
    End If

    stageText = "get Activity Log worksheet"
    Set ws = gXLWorkbook.Worksheets("Activity Log")

    stageText = "find last activity row"
    lastRow = ws.Cells(ws.Rows.Count, 1).End(XL_UP).Row
    If lastRow < HEADER_ROW Then lastRow = HEADER_ROW
    targetRow = 0

    ' Read only the recent audit tail in one COM call. A retry of an event
    ' happens immediately; scanning the entire history made Outlook slower
    ' as the append-only log grew.
    If lastRow >= FIRST_DATA_ROW Then
        scanStart = lastRow - 49
        If scanStart < FIRST_DATA_ROW Then scanStart = FIRST_DATA_ROW

        stageText = "read recent Activity Log rows " & CStr(scanStart) & ":" & CStr(lastRow)
        recentValues = ws.Range("A" & scanStart & ":I" & lastRow).Value2

        For r = 1 To UBound(recentValues, 1)
            actualRow = scanStart + r - 1

            If VWGX_ValuesMatch(recentValues(r, 1), eventRecord(0)) And _
               VWGX_ValuesMatch(recentValues(r, 2), eventRecord(1)) And _
               VWGX_ValuesMatch(recentValues(r, 3), eventRecord(2)) Then

                allMatch = True
                For i = 0 To 8
                    If Not VWGX_ValuesMatch(recentValues(r, i + 1), eventRecord(i)) Then
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

            ' Reuse a matching incomplete row (timestamp + event type, but no key).
            If targetRow = 0 Then
                If VWGX_ValuesMatch(recentValues(r, 1), eventRecord(0)) And _
                   VWGX_ValuesMatch(recentValues(r, 2), eventRecord(1)) Then

                    existingValue = recentValues(r, 3)
                    candidate = IsEmpty(existingValue) Or IsNull(existingValue)
                    If Not candidate Then candidate = (Len(Trim$(CStr(existingValue))) = 0)
                    If candidate Then
                        candidate = True
                        For i = 3 To 8
                            existingValue = recentValues(r, i + 1)
                            If Not IsEmpty(existingValue) And Not IsNull(existingValue) Then
                                If Len(Trim$(CStr(existingValue))) > 0 Then
                                    If Not VWGX_ValuesMatch(existingValue, eventRecord(i)) Then
                                        candidate = False
                                        Exit For
                                    End If
                                End If
                            End If
                        Next i
                        If candidate Then targetRow = actualRow
                    End If
                End If
            End If
        Next r
    End If

    If targetRow = 0 Then targetRow = lastRow + 1

    ' Stage all eight payload fields and write them in one block. Column A is
    ' written last as the commit marker, so an interrupted write is not logged
    ' as a completed event.
    For i = 1 To 8
        If IsNull(eventRecord(i)) Or IsEmpty(eventRecord(i)) Then
            writeValues(1, i) = Empty
        Else
            writeValues(1, i) = eventRecord(i)
        End If
    Next i

    stageText = "format Activity Log row " & CStr(targetRow)
    ws.Range("B" & targetRow & ":G" & targetRow).NumberFormat = "@"
    ws.Cells(targetRow, 8).NumberFormat = "0"
    ws.Cells(targetRow, 9).NumberFormat = "@"

    stageText = "write Activity Log fields B:I row " & CStr(targetRow)
    ws.Range("B" & targetRow & ":I" & targetRow).Value2 = writeValues
    writeStarted = True
    gXLWorkbookDirty = True

    stageText = "write Event Time completion marker in row " & CStr(targetRow)
    ws.Cells(targetRow, 1).NumberFormat = "dd/mmm/yyyy hh:mm:ss"
    ws.Cells(targetRow, 1).Value2 = eventRecord(0)
    gXLWorkbookDirty = True

    stageText = "verify Activity Log row " & CStr(targetRow)
    verifyValues = ws.Range("A" & targetRow & ":I" & targetRow).Value2
    For i = 0 To 8
        If Not VWGX_ValuesMatch(verifyValues(1, i + 1), eventRecord(i)) Then
            Err.Raise vbObjectError + 1613, "VWGX_AppendActivity", _
                      "Activity Log read-back mismatch in column " & CStr(i + 1) & "."
        End If
    Next i

    stageText = "save Activity Log"
    SaveWorkbookIfNeeded

    VWGX_AppendActivity = True
    gLastActivityWriteError = vbNullString
    gTrackerWriteBusy = False
    Exit Function

WriteFailed:
    errNumber = Err.Number
    errText = Err.Description
    gLastActivityWriteError = "[" & stageText & "] " & CStr(errNumber) & " - " & errText
    Debug.Print "[VWG TRACKER] Activity log write failed: " & gLastActivityWriteError

    ' Leave the timestamp blank if the write did not commit. The next attempt
    ' can safely reuse the row without losing earlier completed audit entries.
    If writeStarted And targetRow >= FIRST_DATA_ROW Then
        On Error Resume Next
        Err.Clear
        ws.Cells(targetRow, 1).ClearContents
        If Err.Number <> 0 Then rollbackError = "Clear timestamp failed: " & Err.Number & " - " & Err.Description
        Err.Clear
        gXLWorkbookDirty = True
        SaveWorkbookIfNeeded
        If Err.Number <> 0 Then
            If Len(rollbackError) > 0 Then rollbackError = rollbackError & " | "
            rollbackError = rollbackError & "Rollback save failed: " & Err.Number & " - " & Err.Description
        End If
        On Error GoTo 0
    End If

    If Len(rollbackError) > 0 Then
        gLastActivityWriteError = gLastActivityWriteError & " | " & rollbackError
        Debug.Print "[VWG TRACKER] Activity log rollback warning: " & rollbackError
    End If

    gTrackerWriteBusy = False
    VWGX_AppendActivity = False
End Function

Public Function VWGX_LastActivityWriteError() As String
    VWGX_LastActivityWriteError = gLastActivityWriteError
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

Public Function VWGX_GetAutoAllocationCounts(ByVal allocationDay As Date) As Object
    Dim counts As Object
    Dim ws As Object
    Dim lastRow As Long
    Dim data As Variant
    Dim roster As Variant
    Dim r As Long
    Dim nm As String
    Dim allocatedAt As Variant
    Dim dayStart As Date
    Dim nextDay As Date

    Set counts = CreateObject("Scripting.Dictionary")
    counts.CompareMode = vbTextCompare
    dayStart = DateValue(allocationDay)
    nextDay = DateAdd("d", 1, dayStart)

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

    ' Fairness is a DAILY counter, based on Egypt-local Allocated At (Egypt).
    ' Yesterday's allocations are never carried into today's selection.
    ' The fortnight roster itself and all historical assignment records remain intact.
    Set ws = gXLWorkbook.Worksheets("Assignments")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(XL_UP).Row
    If lastRow < FIRST_DATA_ROW Then GoTo Done

    ' Read only columns E:O; the fairness calculation needs just E and O.
    ' This avoids moving every unrelated assignment field through Excel COM.
    data = ws.Range("E" & FIRST_DATA_ROW & ":O" & lastRow).Value2
    For r = LBound(data, 1) To UBound(data, 1)
        nm = Trim$(CStr(data(r, 1)))       ' Original Assignee, column E
        allocatedAt = data(r, 11)         ' Allocated At (Egypt), column O

        If Len(nm) > 0 And IsDate(allocatedAt) Then
            If CDate(allocatedAt) >= dayStart And CDate(allocatedAt) < nextDay Then
                If counts.Exists(nm) Then counts(nm) = CLng(counts(nm)) + 1
            End If
        End If
    Next r

Done:
    Set VWGX_GetAutoAllocationCounts = counts
    Exit Function

ReadFailed:
    Debug.Print "[VWG TRACKER] Daily allocation count read failed: " & Err.Number & " - " & Err.Description
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

    ' Workload only needs Current Owner (F) and Status (J).
    data = ws.Range("F" & FIRST_DATA_ROW & ":J" & lastRow).Value2
    For r = LBound(data, 1) To UBound(data, 1)
        nm = Trim$(CStr(data(r, 1)))
        statusText = UCase$(Trim$(CStr(data(r, 5))))
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
        Case "ASSIGNED", "HANDED OVER", "HANDOVER PENDING", "RECOVERED - REVIEW", "NEEDS REVIEW", "RECOVERY FAILED"
            IsActiveStatus = True
        Case Else
            IsActiveStatus = False
    End Select
End Function

Public Sub VWGX_RefreshSummary(Optional ByVal allocationDay As Variant)
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
    Dim dayForCounts As Date
    Dim stageText As String
    Dim errNumber As Long
    Dim errText As String

    On Error GoTo RefreshFailed

    stageText = "connect to workbook"
    If Not VWGX_EnsureConnected(False) Then Exit Sub
    If gTrackerWriteBusy Then
        Debug.Print "[VWG TRACKER] Dashboard refresh skipped: another tracker write is active."
        Exit Sub
    End If
    gTrackerWriteBusy = True

    stageText = "resolve Egypt-local daily allocation date"
    If IsDate(allocationDay) Then
        dayForCounts = DateValue(CDate(allocationDay))
    Else
        dayForCounts = Date
    End If
    Set allocCounts = VWGX_GetAutoAllocationCounts(dayForCounts)
    Set workloadCounts = VWGX_GetCurrentWorkloadCounts()

    stageText = "get Team & Shifts worksheet"
    Set wsTeam = gXLWorkbook.Worksheets("Team & Shifts")
    stageText = "label daily allocation counter"
    wsTeam.Cells(HEADER_ROW, 7).Value2 = "Today's Auto Allocations"
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
                        wsTeam.Cells(targetRow, 7).Value2 = allocCounts(nameText)
                    End If
                End If
                If Not workloadCounts Is Nothing Then
                    If workloadCounts.Exists(nameText) Then
                        stageText = "write Team & Shifts workload count, row " & CStr(targetRow) & _
                                    ", column H, colleague " & nameText
                        wsTeam.Cells(targetRow, 8).Value2 = workloadCounts(nameText)
                    End If
                End If
            End If
        Next r
    End If

    stageText = "get Dashboard worksheet"
    Set wsDashboard = gXLWorkbook.Worksheets("Dashboard")
    stageText = "label daily fairness rule on Dashboard"
    wsDashboard.Cells(2, 1).Value2 = "Assignment fairness • legitimate handovers • recovery monitoring • daily allocation counts + fortnightly shift roster"
    wsDashboard.Cells(9, 2).Value2 = "Today's auto allocations"
    wsDashboard.Cells(9, 6).Value2 = "Daily fairness uses today's Egypt-local automatic allocation totals first, then current active workload, then a rotating tie-break. Legitimate handovers do not count as new allocations."
    stageText = "update Settings fairness description"
    gXLWorkbook.Worksheets("Settings").Cells(22, 2).Value2 = "Daily automatic allocation count (Egypt-local date), then current active workload, then rotating tie-break. Previous-day counts are not carried forward."
    stageText = "get Assignments worksheet"
    Set ws = gXLWorkbook.Worksheets("Assignments")

    stageText = "read assignment summary"
    lastRow = ws.Cells(ws.Rows.Count, 1).End(XL_UP).Row
    If lastRow >= FIRST_DATA_ROW Then
        data = ws.Range("A" & FIRST_DATA_ROW & ":J" & lastRow).Value2
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
        data = ws.Range("B" & FIRST_DATA_ROW & ":I" & lastRow).Value2
        For r = LBound(data, 1) To UBound(data, 1)
            Select Case UCase$(Trim$(CStr(data(r, 1))))
                Case "HANDOVER": handoverCount = handoverCount + 1
                Case "CATEGORY_REMOVAL_DETECTED", "CATEGORY_RESTORED", "RECOVERY_FAILED"
                    incidentCount = incidentCount + 1
            End Select
        Next r
    End If

    stageText = "write Dashboard KPI A5"
    wsDashboard.Range("A5").Value2 = assignedCount
    stageText = "write Dashboard KPI D5"
    wsDashboard.Range("D5").Value2 = activeCount
    stageText = "write Dashboard KPI G5"
    wsDashboard.Range("G5").Value2 = handoverCount
    stageText = "write Dashboard KPI J5"
    wsDashboard.Range("J5").Value2 = incidentCount

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

Public Sub VWGX_BackfillMissingAutoAssignmentEvents_Last7Days()
    Dim wsAssignments As Object
    Dim wsActivity As Object
    Dim lastAssignmentRow As Long
    Dim lastActivityRow As Long
    Dim assignmentsData As Variant
    Dim activityData As Variant
    Dim loggedKeys As Object
    Dim eventRecord(0 To 8) As Variant
    Dim r As Long
    Dim keyText As String
    Dim originalAssignee As String
    Dim allocatedAt As Variant
    Dim fromDate As Date
    Dim untilDate As Date
    Dim addedCount As Long
    Dim skippedCount As Long
    Dim failedCount As Long
    Dim lastErrorText As String

    On Error GoTo BackfillFailed
    If Not VWGX_EnsureConnected(True) Then Exit Sub
    If gTrackerWriteBusy Then
        MsgBox "The tracker is currently writing. Stop the dispatcher and try this macro again.", _
               vbExclamation, "VWG Activity Log Backfill"
        Exit Sub
    End If

    fromDate = DateAdd("d", -6, Date)
    untilDate = DateAdd("d", 1, Date)

    Set wsAssignments = gXLWorkbook.Worksheets("Assignments")
    Set wsActivity = gXLWorkbook.Worksheets("Activity Log")
    Set loggedKeys = CreateObject("Scripting.Dictionary")
    loggedKeys.CompareMode = vbTextCompare

    lastActivityRow = wsActivity.Cells(wsActivity.Rows.Count, 1).End(XL_UP).Row
    If lastActivityRow >= FIRST_DATA_ROW Then
        activityData = wsActivity.Range("A" & FIRST_DATA_ROW & ":I" & lastActivityRow).Value2
        For r = LBound(activityData, 1) To UBound(activityData, 1)
            If UCase$(Trim$(CStr(activityData(r, 2)))) = "AUTO_ASSIGNMENT" Then
                keyText = Trim$(CStr(activityData(r, 3)))
                If Len(keyText) > 0 Then loggedKeys(keyText) = True
            End If
        Next r
    End If

    lastAssignmentRow = wsAssignments.Cells(wsAssignments.Rows.Count, 1).End(XL_UP).Row
    If lastAssignmentRow < FIRST_DATA_ROW Then
        MsgBox "There are no assignment rows to reconcile.", vbInformation, "VWG Activity Log Backfill"
        Exit Sub
    End If

    assignmentsData = wsAssignments.Range("A" & FIRST_DATA_ROW & ":U" & lastAssignmentRow).Value2

    For r = LBound(assignmentsData, 1) To UBound(assignmentsData, 1)
        keyText = Trim$(CStr(assignmentsData(r, 1)))
        originalAssignee = Trim$(CStr(assignmentsData(r, 5)))
        allocatedAt = assignmentsData(r, 15)

        If Len(keyText) > 0 And Len(originalAssignee) > 0 And IsDate(allocatedAt) Then
            If CDate(allocatedAt) >= fromDate And CDate(allocatedAt) < untilDate Then
                If loggedKeys.Exists(keyText) Then
                    skippedCount = skippedCount + 1
                Else
                    eventRecord(0) = CDate(allocatedAt)
                    eventRecord(1) = "AUTO_ASSIGNMENT"
                    eventRecord(2) = keyText
                    eventRecord(3) = CStr(assignmentsData(r, 3))
                    eventRecord(4) = originalAssignee
                    eventRecord(5) = CStr(assignmentsData(r, 6))
                    eventRecord(6) = CStr(assignmentsData(r, 7))
                    If IsNumeric(assignmentsData(r, 8)) Then
                        eventRecord(7) = CLng(Val(CStr(assignmentsData(r, 8))))
                    Else
                        eventRecord(7) = 0
                    End If
                    eventRecord(8) = "Backfilled from the saved Assignments register. The original Activity Log event was not committed; no Outlook category or assignment was changed. Subject=" & CStr(assignmentsData(r, 4))

                    If VWGX_AppendActivity(eventRecord) Then
                        loggedKeys(keyText) = True
                        addedCount = addedCount + 1
                    Else
                        failedCount = failedCount + 1
                        lastErrorText = VWGX_LastActivityWriteError()
                    End If
                End If
            End If
        End If
    Next r

    MsgBox "Activity Log backfill finished for " & Format$(fromDate, "dd-mmm-yyyy") & _
           " to " & Format$(Date, "dd-mmm-yyyy") & "." & vbCrLf & vbCrLf & _
           "Rebuilt events: " & addedCount & vbCrLf & _
           "Already logged / skipped: " & skippedCount & vbCrLf & _
           "Write failures: " & failedCount & _
           IIf(Len(lastErrorText) > 0, vbCrLf & vbCrLf & "Last error: " & lastErrorText, ""), _
           IIf(failedCount = 0, vbInformation, vbExclamation), "VWG Activity Log Backfill"
    Exit Sub

BackfillFailed:
    MsgBox "Backfill stopped safely: " & Err.Number & " - " & Err.Description, _
           vbExclamation, "VWG Activity Log Backfill"
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
