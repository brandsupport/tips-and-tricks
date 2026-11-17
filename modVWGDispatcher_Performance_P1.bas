Attribute VB_Name = "modVWGDispatcher"
Option Explicit

'===========================================================
' VWG BRAND SUPPORT - INTEGRATED DISPATCHER V5.1
' Updated 12 Nov 2026: lifecycle/restart guard + stale timer rejection.
' Base: VWG-TRACKER-2026.11.12-WRITE-QUEUE-R2 dispatcher source.
' Updated 13 Nov 2026: recurring backlog discovery + source-Inbox fairness.
' Updated 16 Nov 2026: daily current-workload reset + one-time fairness baseline.
' Updated 17 Nov 2026: event staging, bounded workers, discovery throttling and cached fairness snapshots.
'
' Features:
'   * Four independent Outlook stores, each with an Inbox.
'   * New arrivals are queued; full four-Inbox discovery runs before allocation.
'   * Oldest eligible unassigned email is processed across all four Inboxes.
'   * Fairness is source-Inbox history first, then global daily allocations,
'     active workload, and rotating roster tie-break.
'   * Existing fortnight/shift-group eligibility is retained; no new daily off-duty gate.
'   * Historical assignments are never reset.
'   * Queue outside shift / when no eligible roster is configured.
'   * Legitimate category-to-category handovers are recorded, not flagged.
'   * Missing categories are given a short handover grace period, then
'     restored to the current owner; the incident is tracked in Excel.
'   * No separate "Assignment Restored" Outlook category is left on mail.
'   * Workbook is the operational assignment and activity record.
'   * Periodic timer processes queues and pending recovery.
'
' SAFETY:
'   This code ships with DRY_RUN_MODE = True.
'   For real category changes, BOTH the code constant and the workbook
'   Settings value "Dispatcher Mode" must explicitly be switched to LIVE.
'===========================================================

Private Const DRY_RUN_MODE As Boolean = True ' Performance branch retains DRY RUN
Private Const RESTORE_MARKER_PREFIX As String = "Assignment Restored -"
Private Const RESTORE_MARKER_LEGACY_NAME As String = "Assignment Restored"
Private Const TIMER_INTERVAL_DEFAULT As Long = 30
' Performance controls: keep Outlook event work short and bound each timer pass.
Private Const DEFAULT_BACKLOG_SCAN_INTERVAL_SECONDS As Long = 180
Private Const MIN_BACKLOG_SCAN_GAP_SECONDS As Long = 60
Private Const MAX_ITEM_EVENTS_PER_CYCLE As Long = 6
Private Const MAX_ASSIGNMENTS_PER_CYCLE As Long = 5
Private Const MAX_QUEUE_CANDIDATES_PER_CYCLE As Long = 20
Private Const MAX_RECOVERY_ACTIONS_PER_CYCLE As Long = 5
Private Const MAX_RESERVATION_RETRIES_PER_CYCLE As Long = 5
Private Const DEFAULT_RECOVERY_INTERVAL_SECONDS As Long = 60
Private Const DEFAULT_DASHBOARD_REFRESH_INTERVAL_SECONDS As Long = 180
Private Const DEFAULT_ACTIVITY_FALLBACK_FLUSH_INTERVAL_SECONDS As Long = 120
' Allow reasonable time for a person to replace one category with another.
Private Const DEFAULT_HANDOVER_GRACE_SECONDS As Long = 60
Private Const ASSIGNMENT_STABILITY_SECONDS As Long = 30
' MAPI_E_OBJECT_CHANGED: Outlook item changed after this cached object was read.
Private Const ERR_MAPI_OBJECT_CHANGED As Long = -2147221239

Private Const AF_KEY As Long = 0
Private Const AF_RECEIVED_UK As Long = 1
Private Const AF_FOLDER As Long = 2
Private Const AF_SUBJECT As Long = 3
Private Const AF_ORIGINAL As Long = 4
Private Const AF_OWNER As Long = 5
Private Const AF_FLAG As Long = 6
Private Const AF_INCIDENTS As Long = 7
Private Const AF_LAST_INCIDENT As Long = 8
Private Const AF_STATUS As Long = 9
Private Const AF_HANDOVER_NOTES As Long = 10
Private Const AF_LAST_VERIFIED As Long = 11
Private Const AF_LOGIN_EMAIL As Long = 12
Private Const AF_UPDATED_AT As Long = 13
Private Const AF_ALLOCATED_AT_EGYPT As Long = 14
Private Const AF_ROSTER_PERIOD As Long = 15
Private Const AF_ENTRY_ID As Long = 16
Private Const AF_STORE_ID As Long = 17
Private Const AF_PENDING_SINCE As Long = 18
Private Const AF_CATEGORY_STATE As Long = 19
Private Const AF_HANDOVER_COUNT As Long = 20

#If VBA7 Then
Private Declare PtrSafe Function SetTimer Lib "user32" ( _
    ByVal hWnd As LongPtr, ByVal nIDEvent As LongPtr, _
    ByVal uElapse As Long, ByVal lpTimerFunc As LongPtr) As LongPtr
Private Declare PtrSafe Function KillTimer Lib "user32" ( _
    ByVal hWnd As LongPtr, ByVal uIDEvent As LongPtr) As Long
Private gTimerID As LongPtr
#Else
Private Declare Function SetTimer Lib "user32" ( _
    ByVal hWnd As Long, ByVal nIDEvent As Long, _
    ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" ( _
    ByVal hWnd As Long, ByVal uIDEvent As Long) As Long
Private gTimerID As Long
#End If

'===========================================================
' NATIVE WINDOWS TIME-ZONE API (NO POWERSHELL / WSH PROCESS)
'===========================================================
Private Type VWG_SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Private Type VWG_DYNAMIC_TIME_ZONE_INFORMATION
    Bias As Long
    StandardName(0 To 31) As Integer
    StandardDate As VWG_SYSTEMTIME
    StandardBias As Long
    DaylightName(0 To 31) As Integer
    DaylightDate As VWG_SYSTEMTIME
    DaylightBias As Long
    TimeZoneKeyName(0 To 127) As Integer
    DynamicDaylightTimeDisabled As Long
End Type

#If VBA7 Then
Private Declare PtrSafe Sub GetSystemTime Lib "kernel32" (ByRef lpSystemTime As VWG_SYSTEMTIME)
Private Declare PtrSafe Function GetDynamicTimeZoneInformation Lib "kernel32" (ByRef lpTimeZoneInformation As VWG_DYNAMIC_TIME_ZONE_INFORMATION) As Long
Private Declare PtrSafe Function EnumDynamicTimeZoneInformation Lib "advapi32" (ByVal dwIndex As Long, ByRef lpTimeZoneInformation As VWG_DYNAMIC_TIME_ZONE_INFORMATION) As Long
Private Declare PtrSafe Function SystemTimeToTzSpecificLocalTimeEx Lib "kernel32" (ByRef lpTimeZoneInformation As VWG_DYNAMIC_TIME_ZONE_INFORMATION, ByRef lpUniversalTime As VWG_SYSTEMTIME, ByRef lpLocalTime As VWG_SYSTEMTIME) As Long
Private Declare PtrSafe Function TzSpecificLocalTimeToSystemTimeEx Lib "kernel32" (ByRef lpTimeZoneInformation As VWG_DYNAMIC_TIME_ZONE_INFORMATION, ByRef lpLocalTime As VWG_SYSTEMTIME, ByRef lpUniversalTime As VWG_SYSTEMTIME) As Long
#Else
Private Declare Sub GetSystemTime Lib "kernel32" (ByRef lpSystemTime As VWG_SYSTEMTIME)
Private Declare Function GetDynamicTimeZoneInformation Lib "kernel32" (ByRef lpTimeZoneInformation As VWG_DYNAMIC_TIME_ZONE_INFORMATION) As Long
Private Declare Function EnumDynamicTimeZoneInformation Lib "advapi32" (ByVal dwIndex As Long, ByRef lpTimeZoneInformation As VWG_DYNAMIC_TIME_ZONE_INFORMATION) As Long
Private Declare Function SystemTimeToTzSpecificLocalTimeEx Lib "kernel32" (ByRef lpTimeZoneInformation As VWG_DYNAMIC_TIME_ZONE_INFORMATION, ByRef lpUniversalTime As VWG_SYSTEMTIME, ByRef lpLocalTime As VWG_SYSTEMTIME) As Long
Private Declare Function TzSpecificLocalTimeToSystemTimeEx Lib "kernel32" (ByRef lpTimeZoneInformation As VWG_DYNAMIC_TIME_ZONE_INFORMATION, ByRef lpLocalTime As VWG_SYSTEMTIME, ByRef lpUniversalTime As VWG_SYSTEMTIME) As Long
#End If

Private Const VWG_ERROR_NO_MORE_ITEMS As Long = 259
Private Const VWG_TIME_ZONE_ID_INVALID As Long = -1

Private gMonitors As Collection
Private gRecords As Object              ' key -> zero-based 21-field assignment array
Private gBusy As Object                 ' prevent recursive ItemChange handling
Private gDeferredEvents As Object       ' item key -> EntryID/StoreID/event during baseline
Private gPendingItemEvents As Object    ' lightweight Outlook event queue; drained by timer worker
Private gBacklogScanRequired As Boolean
Private gLastBacklogScanAt As Date
Private gLastRecoveryRunAt As Date
Private gLastSummaryRefreshAt As Date
Private gLastActivityFlushAt As Date
Private gRecordRefreshCursorIndex As Long
Private gRecoveryCursorKey As String
Private gReservationCursorKey As String
Private gTieBreakDirty As Boolean
Private gPreviewSeen As Object          ' dry-run preview dedupe
Private gPreviewAllocationCounts As Object ' simulated allocation totals for fair dry-run previews
Private gPreviewWorkloadCounts As Object  ' simulated current workload for fair dry-run previews
Private gFolderAllocationCounts As Object  ' folder path -> category -> historical auto allocations
Private gPreviewFolderAllocationCounts As Object ' dry-run-only source-folder deltas; never pollute real history
Private gStarted As Boolean
Private gStarting As Boolean
Private gStopping As Boolean
Private gStopRequested As Boolean
Private gInBaseline As Boolean
Private gTimerBusy As Boolean
Private gScanBusy As Boolean
Private gRoster As Variant              ' rows from Team & Shifts A:F
Private gRosterPeriodStart As Date
Private gRosterPeriodSet As Boolean
Private gFairnessDate As Date
Private gFairnessDateSet As Boolean
Private gLastChosenDailyCount As Long
Private gLastChosenFolderAllocationCount As Long
Private gLastChosenWorkload As Long
Private gLastEligiblePoolCount As Long
Private gNextTieBreakIndex As Long
Private gExpectedFolderCount As Long
Private gHandoverGraceSeconds As Long
Private gTimerIntervalSeconds As Long
Private gLastTimerRun As Date
Private gTimeZoneErrorLogged As Boolean
Private gLastFullScanAt As Date
Private gCachedEgyptTime As Date
Private gCachedEgyptTimeAt As Date
Private gHaveCachedEgyptTime As Boolean
Private gLocalTimeZoneId As String
Private gLocalTimeZoneLoaded As Boolean
Private gFairnessBatchActive As Boolean
Private gBatchDailyAllocationCounts As Object
Private gBatchDailyWorkloadCounts As Object
Private gBatchEarlyStart As Date
Private gBatchEarlyEnd As Date
Private gBatchLateStart As Date
Private gBatchLateEnd As Date

'===========================================================
' FOUR OUTLOOK STORE ROOTS
'===========================================================

Private Function GetMailboxNames() As Variant
    GetMailboxNames = Array( _
        "VWGUK R: Milton Keynes, AgencyDetags", _
        "VWGUK R: Milton Keynes, Detags", _
        "VWGUK R: Milton Keynes, Transfers", _
        "VWGUK R: Milton Keynes, Unadopts")
End Function

'===========================================================
' START / STOP
'===========================================================

Public Sub StartDispatcher()
    Dim mailboxNames As Variant
    Dim mailboxName As Variant
    Dim targetFolder As Outlook.Folder
    Dim watcher As clsVWGFolderMonitor
    Dim failedFolders As String
    Dim statusText As String
    Dim i As Long
    Dim baselineStatus As String
    Dim startupEgypt As Date
    Dim timeOk As Boolean
    Dim assignmentFallbackRows As Long
    Dim pendingEventCount As Long
    Dim startErrorNumber As Long
    Dim startErrorDescription As String
    Dim startWasCancelled As Boolean

    On Error GoTo StartFailed

    If gStopping Then
        Debug.Print "[VWG LIFECYCLE] Start ignored: dispatcher is stopping."
        Exit Sub
    End If
    If gStarted Then
        Debug.Print "[VWG LIFECYCLE] Duplicate start ignored: dispatcher is already running."
        MsgBox "VWG Dispatcher is already running in this Outlook session.", vbInformation
        Exit Sub
    End If
    If gStarting Then
        Debug.Print "[VWG LIFECYCLE] Duplicate start ignored: startup is already in progress."
        Exit Sub
    End If

    gStarting = True
    gStopRequested = False
    Debug.Print "[VWG LIFECYCLE] START_REQUEST | manual/approved entry point"

    ' Remove any orphaned timer ID before a fresh initialization.
    If gTimerID <> 0 Then StopTimer

    If Not VWGX_EnsureConnected(True) Then
        gStarting = False
        gStopRequested = False
        Debug.Print "[VWG LIFECYCLE] START_CANCELLED | tracker connection was not established."
        Exit Sub
    End If
    EnsureStartNotCancelled
    VWGX_FlushActivityFallback
    LoadWorkbookConfiguration
    EnsureStartNotCancelled

    ' Hard safety gate: no baseline, recovery, or assignment work may run
    ' until Windows can provide a valid Egypt-local timestamp.
    timeOk = GetEgyptNow(startupEgypt)
    EnsureStartNotCancelled
    If Not timeOk Then
        Err.Raise vbObjectError + 2006, "StartDispatcher", _
                  "Egypt time-zone conversion failed. Dispatcher startup is blocked; no assignments or recoveries were processed."
    End If

    If Not IsArray(gRoster) Then
        Err.Raise vbObjectError + 2001, "StartDispatcher", _
                  "The Team & Shifts roster could not be read from the workbook."
    End If

    Set gRecords = CreateObject("Scripting.Dictionary")
    gRecords.CompareMode = vbTextCompare
    Set gBusy = CreateObject("Scripting.Dictionary")
    gBusy.CompareMode = vbTextCompare
    Set gDeferredEvents = CreateObject("Scripting.Dictionary")
    gDeferredEvents.CompareMode = vbTextCompare
    Set gPendingItemEvents = CreateObject("Scripting.Dictionary")
    gPendingItemEvents.CompareMode = vbTextCompare
    gBacklogScanRequired = False
    gLastBacklogScanAt = 0
    gLastRecoveryRunAt = 0
    gLastSummaryRefreshAt = 0
    gLastActivityFlushAt = startupEgypt
    gRecordRefreshCursorIndex = 0
    gRecoveryCursorKey = vbNullString
    gReservationCursorKey = vbNullString
    gTieBreakDirty = False
    gFairnessBatchActive = False
    Set gBatchDailyAllocationCounts = Nothing
    Set gBatchDailyWorkloadCounts = Nothing
    Set gPreviewSeen = CreateObject("Scripting.Dictionary")
    gPreviewSeen.CompareMode = vbTextCompare
    Set gPreviewAllocationCounts = CreateObject("Scripting.Dictionary")
    gPreviewAllocationCounts.CompareMode = vbTextCompare
    Set gPreviewWorkloadCounts = CreateObject("Scripting.Dictionary")
    gPreviewWorkloadCounts.CompareMode = vbTextCompare
    Set gFolderAllocationCounts = CreateObject("Scripting.Dictionary")
    gFolderAllocationCounts.CompareMode = vbTextCompare
    Set gPreviewFolderAllocationCounts = CreateObject("Scripting.Dictionary")
    gPreviewFolderAllocationCounts.CompareMode = vbTextCompare

    If Not VWGX_FlushAssignmentFallback(assignmentFallbackRows) Then
        Err.Raise vbObjectError + 2010, "StartDispatcher", _
                  "The local assignment retry queue could not be fully committed to the tracker. Dispatcher start was blocked to prevent duplicate or lost allocations."
    End If
    If assignmentFallbackRows > 0 Then
        Debug.Print "[VWG] Restored " & CStr(assignmentFallbackRows) & " deferred assignment row(s) before loading dispatcher state."
    End If
    LoadRecordsFromWorkbook
    EnsureStartNotCancelled

    mailboxNames = GetMailboxNames()
    gExpectedFolderCount = UBound(mailboxNames) - LBound(mailboxNames) + 1
    Set gMonitors = New Collection

    ' Resolve all four Inboxes before registering a partial service.
    For Each mailboxName In mailboxNames
        Set targetFolder = GetFolderByPath(CStr(mailboxName), "Inbox")
        If targetFolder Is Nothing Then
            failedFolders = failedFolders & vbCrLf & "- " & CStr(mailboxName) & "\Inbox"
        Else
            Set watcher = New clsVWGFolderMonitor
            Set watcher.WatchedFolder = targetFolder
            Set watcher.WatchedItems = targetFolder.Items
            gMonitors.Add watcher
            statusText = statusText & vbCrLf & targetFolder.FolderPath
        End If
        Set targetFolder = Nothing
        Set watcher = Nothing
    Next mailboxName

    EnsureStartNotCancelled
    If gMonitors.Count <> gExpectedFolderCount Or Len(failedFolders) > 0 Then
        Err.Raise vbObjectError + 2002, "StartDispatcher", _
                  "Could not resolve all four Inboxes:" & failedFolders
    End If

    gStarted = True
    gInBaseline = True

    baselineStatus = UCase$(Trim$(CStr(VWGX_GetSetting("Baseline Status"))))
    If baselineStatus <> "COMPLETE" Then
        If Not RunInitialBaseline() Then
            Err.Raise vbObjectError + 2003, "StartDispatcher", _
                      "Initial baseline failed. The completion marker was not set."
        End If
    Else
        ' Every start imports any already-categorised message that is not yet tracked.
        If Not ScanExistingCategorisedMessages() Then
            Err.Raise vbObjectError + 2004, "StartDispatcher", _
                      "Existing-assignment reconciliation failed."
        End If
    End If
    EnsureStartNotCancelled

    ' Defer unassigned-backlog enrollment until after baseline ItemAdd events
    ' have been transferred to the lightweight event queue. The final startup
    ' sweep below is the single authoritative complete discovery before release.
    gInBaseline = False
    ProcessDeferredItemEvents
    EnsureStartNotCancelled

    RefreshCurrentRecordStates
    EnsureStartNotCancelled
    ProcessPendingRecoveries
    EnsureStartNotCancelled
    ' The final startup sweep consumes any untracked arrivals captured during
    ' baseline processing. Only events raised while this sweep runs should
    ' request yet another sweep.
    gBacklogScanRequired = False
    If Not ScanAndQueueUnassignedBacklog() Then
        gBacklogScanRequired = True
        Err.Raise vbObjectError + 2031, "StartDispatcher", _
                  "Final startup backlog discovery failed; queued allocations were not released."
    End If
    If Not GetEgyptNow(gLastBacklogScanAt) Then gLastBacklogScanAt = startupEgypt
    If Not gPendingItemEvents Is Nothing Then pendingEventCount = gPendingItemEvents.Count
    ' Do not allocate while captured Outlook events remain. The event handler
    ' independently marks whether an ItemAdd/untracked event requires another
    ' complete sweep; a tracked ItemChange only needs draining.
    EnsureStartNotCancelled
    If gBacklogScanRequired Or pendingEventCount > 0 Then
        Debug.Print "[VWG PERFORMANCE] Startup allocations deferred | events pending=" & _
                    CStr(pendingEventCount) & " | full sweep required=" & CStr(gBacklogScanRequired)
    Else
        ProcessQueuedAssignments
    End If
    EnsureStartNotCancelled

    gTimerIntervalSeconds = GetNumericSetting( _
        "Maintenance Poll Interval (seconds)", TIMER_INTERVAL_DEFAULT)
    If gTimerIntervalSeconds < 30 Then gTimerIntervalSeconds = TIMER_INTERVAL_DEFAULT
    gLastFullScanAt = startupEgypt
    EnsureStartNotCancelled

    ' Windows timer callbacks run inside Outlook's message loop. Stop the
    ' dispatcher before editing/compiling VBA; StopDispatcher kills this timer.
    gTimerID = SetTimer(0, 0, gTimerIntervalSeconds * 1000, AddressOf VWGTimerProc)
    If gTimerID = 0 Then
        Err.Raise vbObjectError + 2005, "StartDispatcher", _
                  "Could not start the maintenance timer. Dispatcher was stopped to avoid stranded queue/recovery work."
    End If

    VWGX_RefreshSummary startupEgypt
    gLastSummaryRefreshAt = startupEgypt
    EnsureStartNotCancelled
    gStarting = False
    gStopRequested = False
    Debug.Print "[VWG LIFECYCLE] STARTED | timer ID=" & CStr(gTimerID) & _
                " | monitors=" & CStr(gMonitors.Count)

    If IsDryRun() Then
        MsgBox "VWG V5 started in DRY-RUN mode." & vbCrLf & _
               "Inboxes watched: " & gMonitors.Count & vbCrLf & _
               "Monitored folders:" & statusText & vbCrLf & _
               "Excel tracker: connected" & vbCrLf & _
               "Maintenance check: every " & gTimerIntervalSeconds & " seconds" & vbCrLf & _
               "Current Workload is recalculated from active automatic assignments made today (Egypt-local date); it resets automatically each day. Today's Auto Allocations remains a separate daily count." & vbCrLf & _
               "No email categories will be changed by normal assignment/recovery logic." & vbCrLf & vbCrLf & _
               "Only ONE computer should run the dispatcher. Stop it before editing VBA.", _
               vbInformation, "VWG Dispatcher V5"
    Else
        MsgBox "VWG V5 is LIVE." & vbCrLf & _
               "Inboxes watched: " & gMonitors.Count & vbCrLf & _
               "Monitored folders:" & statusText & vbCrLf & _
               "Excel tracker: connected" & vbCrLf & _
               "Maintenance check: every " & gTimerIntervalSeconds & " seconds" & vbCrLf & _
               "Only ONE computer should run the dispatcher.", _
               vbExclamation, "VWG Dispatcher V5 - Live"
    End If
    Exit Sub

StartFailed:
    startErrorNumber = Err.Number
    startErrorDescription = Err.Description
    startWasCancelled = (gStopRequested Or startErrorNumber = vbObjectError + 2011)

    ' Tear down the partial start before releasing the lifecycle lock.
    gStopping = True
    gStarted = False
    gInBaseline = False
    StopTimer
    Set gMonitors = Nothing
    Set gRecords = Nothing
    Set gBusy = Nothing
    Set gDeferredEvents = Nothing
    Set gPendingItemEvents = Nothing
    gBacklogScanRequired = False
    gFairnessBatchActive = False
    Set gBatchDailyAllocationCounts = Nothing
    Set gBatchDailyWorkloadCounts = Nothing
    Set gPreviewSeen = Nothing
    Set gPreviewAllocationCounts = Nothing
    Set gPreviewWorkloadCounts = Nothing
    Set gFolderAllocationCounts = Nothing
    Set gPreviewFolderAllocationCounts = Nothing
    gRecordRefreshCursorIndex = 0
    gRecoveryCursorKey = vbNullString
    gReservationCursorKey = vbNullString
    gLastActivityFlushAt = 0
    gTimerBusy = False
    gScanBusy = False
    gStarting = False
    gStopping = False

    If startWasCancelled Then
        Debug.Print "[VWG LIFECYCLE] START_CANCELLED | Startup unwound after Stop request."
    Else
        Debug.Print "[VWG LIFECYCLE] START_FAILED | " & CStr(startErrorNumber) & " - " & startErrorDescription
        On Error Resume Next
        AppendRuntimeEvent "START_FAILED", "Four Inbox roots", _
            CStr(startErrorNumber) & " - " & startErrorDescription
        On Error GoTo 0
        MsgBox "VWG Dispatcher could not start safely:" & vbCrLf & vbCrLf & _
               startErrorDescription & vbCrLf & vbCrLf & _
               "No partial dispatcher remains active.", vbCritical, "VWG Dispatcher V5"
    End If
    gStopRequested = False
End Sub

Public Sub ResetDailyFairnessBaselineNow()
    Dim resetEgypt As Date
    Dim flushedAssignmentRows As Long
    Dim baselineSaved As Boolean

    ' This is an explicit on-demand reset. Every successful invocation
    ' replaces the saved baseline with the current Egypt-local timestamp.
    ' Keep the dispatcher stopped so no allocation can race the reset.
    If gStarted Or gStarting Or gStopping Or gTimerBusy Or gScanBusy Then
        MsgBox "Stop the dispatcher fully before resetting fairness counters.", _
               vbExclamation, "VWG Fairness Baseline"
        Exit Sub
    End If

    On Error GoTo ResetFailed
    If Not VWGX_EnsureConnected(True) Then
        Err.Raise vbObjectError + 2040, "ResetDailyFairnessBaselineNow", _
                  "The tracker workbook could not be connected."
    End If

    ' Commit durable pending reservations first. Never move the fairness
    ' baseline while an assignment write is still waiting to reach the workbook.
    If Not VWGX_FlushAssignmentFallback(flushedAssignmentRows) Then
        Err.Raise vbObjectError + 2041, "ResetDailyFairnessBaselineNow", _
                  "Pending assignment retry rows could not be committed. No baseline was changed."
    End If
    VWGX_FlushActivityFallback

    If Not GetEgyptNow(resetEgypt) Then
        Err.Raise vbObjectError + 2042, "ResetDailyFairnessBaselineNow", _
                  "Egypt-local time could not be determined. No baseline was changed."
    End If

    ' Overwrite any previous baseline. The tracker uses this timestamp to
    ' recalculate G/H and the per-Inbox fairness counts from this exact moment.
    If Not VWGX_SetSetting("Fairness Baseline (Egypt)", resetEgypt) Then
        Err.Raise vbObjectError + 2043, "ResetDailyFairnessBaselineNow", _
                  "The fairness baseline could not be saved to Settings."
    End If
    baselineSaved = True

    ' Discard all in-memory fairness snapshots so the next start reads the
    ' same post-baseline view from the durable tracker.
    Set gFolderAllocationCounts = Nothing
    Set gPreviewFolderAllocationCounts = CreateObject("Scripting.Dictionary")
    gPreviewFolderAllocationCounts.CompareMode = vbTextCompare
    Set gPreviewAllocationCounts = CreateObject("Scripting.Dictionary")
    gPreviewAllocationCounts.CompareMode = vbTextCompare
    Set gPreviewWorkloadCounts = CreateObject("Scripting.Dictionary")
    gPreviewWorkloadCounts.CompareMode = vbTextCompare
    If Not gPreviewSeen Is Nothing Then gPreviewSeen.RemoveAll
    Set gBatchDailyAllocationCounts = Nothing
    Set gBatchDailyWorkloadCounts = Nothing
    gFairnessBatchActive = False
    gFairnessDateSet = False
    gTieBreakDirty = False

    VWGX_RefreshSummary resetEgypt
    If Not VWGX_VerifyDailyFairnessSummary(DateValue(resetEgypt)) Then
        Err.Raise vbObjectError + 2044, "ResetDailyFairnessBaselineNow", _
                  "The baseline was saved, but columns G/H could not be verified after refresh. Leave the dispatcher stopped."
    End If
    Debug.Print "[VWG FAIRNESS] ON-DEMAND BASELINE RESET | Egypt=" & _
                Format$(resetEgypt, "dd/mmm/yyyy hh:nn:ss") & _
                " | prior assignments and activity history preserved."
    MsgBox "Fairness counters reset from: " & Format$(resetEgypt, "dd/mmm/yyyy hh:nn:ss") & _
           " Egypt time." & vbCrLf & vbCrLf & _
           "Today's Auto Allocations and Current Workload now count only assignments made from this point onward." & vbCrLf & _
           "Source-Inbox fairness also restarts from this point." & vbCrLf & _
           "Previous assignment and activity history has NOT been deleted.", _
           vbInformation, "VWG Fairness Baseline"
    Exit Sub

ResetFailed:
    Debug.Print "[VWG FAIRNESS] BASELINE RESET FAILED | " & Err.Number & " - " & Err.Description
    If baselineSaved Then
        MsgBox "The baseline was saved, but the counter refresh/verification did not complete:" & _
               vbCrLf & Err.Description & vbCrLf & vbCrLf & _
               "Leave the dispatcher stopped. Do not start LIVE. Your assignment and activity history is still preserved.", _
               vbCritical, "VWG Fairness Baseline"
    Else
        MsgBox "Fairness baseline was NOT changed:" & vbCrLf & Err.Description, _
               vbCritical, "VWG Fairness Baseline"
    End If
End Sub

Public Sub StopDispatcher()
    If gStopping Then Exit Sub

    ' A Stop requested while StartDispatcher is running must not be overwritten
    ' by the startup routine when it later sets gStarted = True.
    If gStarting Then
        gStopRequested = True
        gStarted = False
        StopTimer
        Debug.Print "[VWG LIFECYCLE] STOP_REQUESTED_DURING_START | Startup will unwind safely."
        Exit Sub
    End If

    gStopping = True
    gStopRequested = True
    gStarted = False                 ' Disable event/timer work first.
    StopTimer                        ' Then kill the native Windows timer.

    Set gMonitors = Nothing
    Set gRecords = Nothing
    Set gBusy = Nothing
    Set gDeferredEvents = Nothing
    Set gPendingItemEvents = Nothing
    gBacklogScanRequired = False
    gFairnessBatchActive = False
    Set gBatchDailyAllocationCounts = Nothing
    Set gBatchDailyWorkloadCounts = Nothing
    Set gPreviewSeen = Nothing
    Set gPreviewAllocationCounts = Nothing
    Set gPreviewWorkloadCounts = Nothing
    Set gFolderAllocationCounts = Nothing
    Set gPreviewFolderAllocationCounts = Nothing
    gInBaseline = False
    gRecordRefreshCursorIndex = 0
    gRecoveryCursorKey = vbNullString
    gReservationCursorKey = vbNullString
    gTimerBusy = False
    gScanBusy = False

    On Error Resume Next
    VWGX_SaveWorkbook
    If Err.Number <> 0 Then
        Debug.Print "[VWG LIFECYCLE] STOP_SAVE_WARNING | " & CStr(Err.Number) & " - " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0

    gStopping = False
    gStopRequested = False
    Debug.Print "[VWG LIFECYCLE] STOPPED | timer removed; folder monitors released."
    MsgBox "VWG Dispatcher monitoring has stopped.", vbInformation, "VWG Dispatcher V5"
End Sub

Private Sub StopTimer()
    Dim killed As Long

    On Error Resume Next
    If gTimerID <> 0 Then
        killed = KillTimer(0, gTimerID)
        If killed = 0 Then
            Debug.Print "[VWG LIFECYCLE] KillTimer reported failure for timer ID " & CStr(gTimerID) & _
                        "; any late callback will be ignored."
        Else
            Debug.Print "[VWG LIFECYCLE] Timer removed | ID=" & CStr(gTimerID)
        End If
    End If
    ' Clear the active ID even on KillTimer failure. VWGTimerProc compares
    ' callback IDs and will reject any queued callback from the old timer.
    gTimerID = 0
    On Error GoTo 0
End Sub

Private Sub EnsureStartNotCancelled()
    If gStopRequested Or gStopping Then
        Err.Raise vbObjectError + 2011, "StartDispatcher", _
                  "Dispatcher startup was cancelled by a Stop request."
    End If
End Sub

'===========================================================
' TIMER CALLBACK AND MAINTENANCE
'===========================================================

#If VBA7 Then
Public Sub VWGTimerProc(ByVal hWnd As LongPtr, ByVal uMsg As Long, _
                        ByVal idEvent As LongPtr, ByVal dwTime As Long)
#Else
Public Sub VWGTimerProc(ByVal hWnd As Long, ByVal uMsg As Long, _
                        ByVal idEvent As Long, ByVal dwTime As Long)
#End If
    Dim callbackStartedAt As Single
    Dim callbackElapsed As Double
    Dim callbackStarted As Boolean

    On Error GoTo TimerProcFailed
    ' Windows can deliver a timer message that was already queued when
    ' KillTimer was called. Only the currently registered timer may run.
    If Not gStarted Or gStarting Or gStopping Then Exit Sub
    If gTimerID = 0 Then Exit Sub
    If idEvent <> gTimerID Then Exit Sub
    If gTimerBusy Then Exit Sub

    gTimerBusy = True
    callbackStartedAt = Timer
    callbackStarted = True
    RunScheduledMaintenance
TimerProcExit:
    If callbackStarted Then
        callbackElapsed = Timer - callbackStartedAt
        If callbackElapsed < 0 Then callbackElapsed = callbackElapsed + 86400#
        If callbackElapsed >= 2# Then
            Debug.Print "[VWG PERFORMANCE] Timer cycle duration=" & Format$(callbackElapsed, "0.00") & _
                        "s | Outlook UI work exceeded 2s; inspect scan/queue/reconciliation phases."
        End If
    End If
    gTimerBusy = False
    Exit Sub

TimerProcFailed:
    Debug.Print "[VWG LIFECYCLE] TIMER_CALLBACK_FAILED | " & CStr(Err.Number) & " - " & Err.Description
    Resume TimerProcExit
End Sub

Private Sub RunScheduledMaintenance()
    Dim nowEgypt As Date
    Dim ok As Boolean
    Dim assignmentFallbackRows As Long
    Dim eventsProcessed As Long
    Dim scanDue As Boolean
    Dim recoveryDue As Boolean

    On Error GoTo MaintenanceFailed
    If Not gStarted Or gStopping Then Exit Sub
    If gScanBusy Then Exit Sub

    ok = GetEgyptNow(nowEgypt)
    If Not ok Then
        If Not gTimeZoneErrorLogged Then
            AppendRuntimeEvent "TIME_ZONE_CONVERSION_FAILED", "Maintenance timer", _
                "Queue and recovery actions were not processed because Egypt time could not be determined."
            gTimeZoneErrorLogged = True
        End If
        Exit Sub
    End If
    gTimeZoneErrorLogged = False
    gLastTimerRun = nowEgypt

    ' R7 safety gate: never enter the retry flusher while an Excel COM write or
    ' workbook connection is active. A later timer tick will retry.
    If VWGX_IsTrackerWriteBusy() Then
        Debug.Print "[VWG PERFORMANCE] Maintenance deferred: tracker connection/write is busy."
        Exit Sub
    End If

    ' Assignment rows are safety-critical. Drain these before any allocation,
    ' recovery action, or fairness snapshot.
    If Not VWGX_FlushAssignmentFallback(assignmentFallbackRows) Then
        Debug.Print "[VWG] Assignment tracker retry queue is still pending; new allocations and recovery actions are paused."
        Exit Sub
    End If
    If assignmentFallbackRows > 0 Then
        LoadRecordsFromWorkbook
        gRecordRefreshCursorIndex = 0
        Set gFolderAllocationCounts = Nothing
        Debug.Print "[VWG] Loaded " & CStr(assignmentFallbackRows) & " recovered assignment row(s) into dispatcher memory."
    End If

    ' Activity audit records are durably staged locally by the tracker. Drain
    ' them less frequently than the 30-second timer to avoid repeated workbook
    ' work during busy periods. Explicit startup/manual scans still flush.
    If gLastActivityFlushAt = 0 Or _
       DateDiff("s", gLastActivityFlushAt, nowEgypt) >= DEFAULT_ACTIVITY_FALLBACK_FLUSH_INTERVAL_SECONDS Then
        VWGX_FlushActivityFallback
        gLastActivityFlushAt = nowEgypt
    End If

    ' Outlook ItemAdd/ItemChange callbacks only enqueue IDs. The timer owns all
    ' subsequent Outlook/Excel work. Bound the number handled per callback.
    DrainPendingItemEvents MAX_ITEM_EVENTS_PER_CYCLE, eventsProcessed
    If VWGX_AssignmentFallbackPending() Then
        Debug.Print "[VWG PERFORMANCE] An Outlook event staged an assignment retry; allocation deferred until the next retry flush."
        Exit Sub
    End If
    If Not gPendingItemEvents Is Nothing Then
        If gPendingItemEvents.Count > 0 Then
            Debug.Print "[VWG PERFORMANCE] Event queue remaining=" & CStr(gPendingItemEvents.Count) & _
                        "; assignment processing deferred to next timer tick."
            Exit Sub
        End If
    End If

    scanDue = gBacklogScanRequired Or gLastBacklogScanAt = 0
    If Not scanDue Then
        scanDue = (DateDiff("s", gLastBacklogScanAt, nowEgypt) >= DEFAULT_BACKLOG_SCAN_INTERVAL_SECONDS)
    End If
    If scanDue Then
        ' Debounce bursts of arrivals: never run a full four-Inbox sweep more
        ' frequently than the minimum gap. Queue allocations remain paused until
        ' the required sweep succeeds.
        If gLastBacklogScanAt <> 0 And _
           DateDiff("s", gLastBacklogScanAt, nowEgypt) < MIN_BACKLOG_SCAN_GAP_SECONDS Then
            Debug.Print "[VWG PERFORMANCE] Backlog sweep requested but minimum scan gap has not elapsed; allocations deferred."
            Exit Sub
        End If
        ' Consume the current scan request. HandleMailItem can set it again if
        ' a new/untracked item arrives while DoEvents yields during the scan.
        gBacklogScanRequired = False
        If Not ScanAndQueueUnassignedBacklog() Then
            gBacklogScanRequired = True
            Debug.Print "[VWG] Backlog discovery failed; queued allocations remain paused until all four Inboxes can be scanned."
            Exit Sub
        End If
        gLastBacklogScanAt = nowEgypt
        If Not gPendingItemEvents Is Nothing Then
            If gPendingItemEvents.Count > 0 Then
                ' Only ItemAdd/untracked events set gBacklogScanRequired. A
                ' tracked ItemChange simply waits for the next event-drain slice.
                Debug.Print "[VWG PERFORMANCE] Outlook event(s) arrived during discovery; event drain continues next cycle."
                Exit Sub
            End If
        End If
        If gBacklogScanRequired Then
            Debug.Print "[VWG PERFORMANCE] An untracked arrival occurred during discovery; another complete sweep is required."
            Exit Sub
        End If
    End If

    ' Do not release any queue unless at least one complete four-Inbox sweep has
    ' succeeded this session. Startup performs the initial complete sweep.
    If gLastBacklogScanAt = 0 Or gBacklogScanRequired Then Exit Sub

    recoveryDue = (gLastRecoveryRunAt = 0 Or _
                   DateDiff("s", gLastRecoveryRunAt, nowEgypt) >= DEFAULT_RECOVERY_INTERVAL_SECONDS)
    If recoveryDue Then
        ProcessPendingRecoveries
        ProcessPendingAssignmentReservations
        gLastRecoveryRunAt = nowEgypt
    End If

    ' Reconcile a small round-robin slice each tick instead of loading and
    ' checking every tracked Outlook item in one long callback.
    RefreshCurrentRecordStates

    ProcessQueuedAssignments
    PersistTieBreakIfNeeded

    If gLastSummaryRefreshAt = 0 Or _
       DateDiff("s", gLastSummaryRefreshAt, nowEgypt) >= DEFAULT_DASHBOARD_REFRESH_INTERVAL_SECONDS Then
        VWGX_RefreshSummary nowEgypt
        gLastSummaryRefreshAt = nowEgypt
    End If
    Exit Sub

MaintenanceFailed:
    Debug.Print "[VWG PERFORMANCE] Maintenance failed | " & CStr(Err.Number) & " - " & Err.Description
    AppendRuntimeEvent "MAINTENANCE_FAILED", "Timer", Err.Number & " - " & Err.Description
End Sub

Public Sub RunQueueAndRecoveryScan()
    Dim summaryEgypt As Date
    Dim assignmentFallbackRows As Long
    Dim pendingQueueBlocked As Boolean
    If Not gStarted Then
        MsgBox "Start the dispatcher first.", vbExclamation, "VWG Dispatcher V5"
        Exit Sub
    End If
    If gTimerBusy Or gScanBusy Then
        MsgBox "A maintenance operation is already running. Try again shortly.", vbInformation
        Exit Sub
    End If
    If VWGX_IsTrackerWriteBusy() Then
        Debug.Print "[VWG] Manual scan deferred: Excel tracker write is still in progress; retry queue is preserved."
        MsgBox "The Excel tracker is finishing another write. No scan or assignment was started. Try again shortly.", _
               vbInformation, "VWG Dispatcher"
        Exit Sub
    End If

    gScanBusy = True
    On Error GoTo ScanFailed

    ' Match scheduled-maintenance safety: assignment reservations must be
    ' committed before manual reconciliation or recovery can touch Outlook.
    If Not VWGX_FlushAssignmentFallback(assignmentFallbackRows) Then
        pendingQueueBlocked = True
        Debug.Print "[VWG] Manual scan stopped: assignment retry queue is still pending. No recovery or new allocation was attempted."
        GoTo ScanComplete
    End If
    If assignmentFallbackRows > 0 Then
        LoadRecordsFromWorkbook
        gRecordRefreshCursorIndex = 0
        Set gFolderAllocationCounts = Nothing
        Debug.Print "[VWG] Manual scan restored " & CStr(assignmentFallbackRows) & _
                    " assignment retry row(s) before reconciliation."
    End If
    VWGX_FlushActivityFallback
    RefreshCurrentRecordStates
    If VWGX_AssignmentFallbackPending() Then
        pendingQueueBlocked = True
        GoTo ScanComplete
    End If
    ProcessPendingRecoveries
    If VWGX_AssignmentFallbackPending() Then
        pendingQueueBlocked = True
        GoTo ScanComplete
    End If
    gBacklogScanRequired = False
    If Not ScanAndQueueUnassignedBacklog() Then
        gBacklogScanRequired = True
        Err.Raise vbObjectError + 2030, "RunQueueAndRecoveryScan", _
                  "Backlog discovery failed. Queued assignments were not released."
    End If
    If Not GetEgyptNow(summaryEgypt) Then
        Err.Raise vbObjectError + 2032, "RunQueueAndRecoveryScan", _
                  "Egypt-local time unavailable; backlog sweep completion could not be timestamped."
    End If
    gLastBacklogScanAt = summaryEgypt
    If VWGX_AssignmentFallbackPending() Then
        Debug.Print "[VWG PERFORMANCE] Manual scan staged pending assignment rows; queue allocation deferred until retry commit."
        GoTo ScanComplete
    End If
    If Not gPendingItemEvents Is Nothing Then
        If gPendingItemEvents.Count > 0 Then
            Debug.Print "[VWG PERFORMANCE] Manual scan complete but pending Outlook events remain; assignments deferred."
            GoTo ScanComplete
        End If
    End If
    If gBacklogScanRequired Then
        Debug.Print "[VWG PERFORMANCE] Manual sweep detected a new/untracked arrival during discovery; allocation deferred."
        GoTo ScanComplete
    End If
    gBacklogScanRequired = False
    ProcessQueuedAssignments
ScanComplete:
    If pendingQueueBlocked Then
        gScanBusy = False
        MsgBox "Manual scan paused because an assignment reservation is still waiting to commit. No new assignment was released. Try again after the retry queue clears.", _
               vbExclamation, "VWG Dispatcher"
        Exit Sub
    End If
    PersistTieBreakIfNeeded
    If GetEgyptNow(summaryEgypt) Then
        VWGX_RefreshSummary summaryEgypt
        gLastSummaryRefreshAt = summaryEgypt
    Else
        Debug.Print "[VWG TRACKER] Summary refresh skipped: Egypt-local date unavailable; no PC-date workload reset was applied."
    End If
    gScanBusy = False
    MsgBox "Queue and recovery scan completed. Check Activity Log for results.", vbInformation
    Exit Sub

ScanFailed:
    gScanBusy = False
    AppendRuntimeEvent "MANUAL_SCAN_FAILED", "RunQueueAndRecoveryScan", Err.Description
    MsgBox "Scan failed: " & Err.Description, vbExclamation
End Sub

' Timer works on a single Windows Outlook session; never run multiple dispatcher PCs.

'===========================================================
' OUTLOOK EVENT ENTRY POINT
'===========================================================

Public Sub HandleMailItem(ByVal Item As Object, _
                          ByVal sourceFolder As Outlook.Folder, _
                          ByVal eventName As String)
    Dim mail As Outlook.MailItem
    Dim itemKey As String
    Dim queuedInfo As Variant

    On Error GoTo HandlerFailed
    If Not gStarted Or gStopping Then Exit Sub
    If Not TypeOf Item Is Outlook.MailItem Then Exit Sub
    Set mail = Item

    itemKey = GetMessageKey(mail, sourceFolder)
    If Len(itemKey) = 0 Then Exit Sub
    If Not gBusy Is Nothing Then
        If gBusy.Exists(itemKey) Then Exit Sub
    End If

    If gInBaseline Then
        If Not gDeferredEvents.Exists(itemKey) Then
            gDeferredEvents.Add itemKey, Array(mail.EntryID, sourceFolder.StoreID, eventName)
        End If
        Exit Sub
    End If

    If gPendingItemEvents Is Nothing Then
        Set gPendingItemEvents = CreateObject("Scripting.Dictionary")
        gPendingItemEvents.CompareMode = vbTextCompare
    End If

    If gPendingItemEvents.Exists(itemKey) Then
        queuedInfo = gPendingItemEvents(itemKey)
        ' Preserve ItemAdd priority if both event types arrive for the same item.
        If StrComp(eventName, "ITEM_ADD", vbTextCompare) = 0 Then queuedInfo(2) = "ITEM_ADD"
        gPendingItemEvents(itemKey) = queuedInfo
    Else
        gPendingItemEvents.Add itemKey, Array(mail.EntryID, sourceFolder.StoreID, eventName)
    End If

    ' A new or not-yet-tracked item requires a complete four-Inbox discovery
    ' before queued allocations are released. Tracked ItemChange events do not.
    If StrComp(eventName, "ITEM_ADD", vbTextCompare) = 0 Or Not gRecords.Exists(itemKey) Then
        gBacklogScanRequired = True
    End If
    Exit Sub

HandlerFailed:
    Debug.Print "[VWG PERFORMANCE] Outlook event capture failed | " & _
                CStr(Err.Number) & " - " & Err.Description & " | Event=" & eventName
    AppendRuntimeEvent "ITEM_HANDLER_FAILED", FolderPathSafe(sourceFolder), _
        Err.Number & " - " & Err.Description & " | Event=" & eventName
End Sub

Private Sub DrainPendingItemEvents(ByVal maxItems As Long, ByRef processedCount As Long)
    Dim keys As Variant
    Dim key As Variant
    Dim info As Variant
    Dim itemObject As Object
    Dim mail As Outlook.MailItem
    Dim sourceFolder As Outlook.Folder
    Dim itemKey As String
    Dim eventName As String

    processedCount = 0
    If gPendingItemEvents Is Nothing Then Exit Sub
    If gPendingItemEvents.Count = 0 Then Exit Sub
    If maxItems < 1 Then Exit Sub

    keys = gPendingItemEvents.Keys
    For Each key In keys
        If processedCount >= maxItems Then Exit For
        If VWGX_AssignmentFallbackPending() Then
            Debug.Print "[VWG PERFORMANCE] Event drain paused because an assignment retry is pending."
            Exit For
        End If
        itemKey = CStr(key)
        If Not gPendingItemEvents.Exists(itemKey) Then GoTo NextPendingEvent
        If gBusy.Exists(itemKey) Then GoTo NextPendingEvent

        info = gPendingItemEvents(itemKey)
        eventName = CStr(info(2))
        ' Remove before processing so a genuinely newer event can enqueue the
        ' same key while this event is being handled; never erase that newer event.
        gPendingItemEvents.Remove itemKey
        Set itemObject = Nothing
        On Error Resume Next
        Set itemObject = Application.Session.GetItemFromID(CStr(info(0)), CStr(info(1)))
        On Error GoTo EventProcessFailed

        If itemObject Is Nothing Then
            ' The item may have been moved/deleted after the event. The key was
            ' already removed before lookup; do not remove it a second time.
            ' The periodic reconciliation scan remains the safety net.
            processedCount = processedCount + 1
            GoTo NextPendingEvent
        End If
        If Not TypeOf itemObject Is Outlook.MailItem Then
            ' The event key was already removed before lookup.
            processedCount = processedCount + 1
            GoTo NextPendingEvent
        End If

        Set mail = itemObject
        Set sourceFolder = mail.Parent
        gBusy.Add itemKey, True
        ProcessMailItemCore mail, sourceFolder, itemKey, eventName
        If gBusy.Exists(itemKey) Then gBusy.Remove itemKey
        processedCount = processedCount + 1
        If VWGX_AssignmentFallbackPending() Then
            Debug.Print "[VWG PERFORMANCE] Event drain paused after preserving an assignment retry."
            Exit Sub
        End If
        GoTo NextPendingEvent

EventProcessFailed:
        Debug.Print "[VWG PERFORMANCE] Deferred item event failed and remains queued | Key=" & _
                    itemKey & " | " & CStr(Err.Number) & " - " & Err.Description
        On Error Resume Next
        If gBusy.Exists(itemKey) Then gBusy.Remove itemKey
        If Not gPendingItemEvents.Exists(itemKey) Then gPendingItemEvents.Add itemKey, info
        On Error GoTo 0
        Exit Sub

NextPendingEvent:
    Next key
End Sub

Private Sub ProcessMailItemCore(ByVal mail As Outlook.MailItem, _
                                ByVal sourceFolder As Outlook.Folder, _
                                ByVal itemKey As String, _
                                ByVal eventName As String)
    Dim record As Variant
    Dim matches As Variant
    Dim matchCount As Long
    Dim matchedCategory As String

    If gRecords.Exists(itemKey) Then
        record = gRecords(itemKey)
        If UCase$(CStr(record(AF_STATUS))) = "QUEUED" Then
            matchCount = GetCategoryMatches(mail.Categories, matches)
            If matchCount = 1 Then
                RegisterManualOwner mail, sourceFolder, itemKey, CStr(matches(0))
            ElseIf matchCount > 1 Then
                MarkAmbiguous record, mail, matches, "Several colleague categories appeared on a queued email."
            End If
            Exit Sub
        End If
        ProcessTrackedMail mail, sourceFolder, itemKey, record
        Exit Sub
    End If

    matchCount = GetCategoryMatches(mail.Categories, matches)
    If matchCount = 1 Then
        RegisterExistingAssignment mail, sourceFolder, itemKey, CStr(matches(0)), "MANUAL_ASSIGNMENT_CAPTURED"
    ElseIf matchCount > 1 Then
        MarkUntrackedAmbiguous mail, sourceFolder, itemKey, matches
    ElseIf StrComp(eventName, "ITEM_ADD", vbTextCompare) = 0 Then
        HandleNewMail mail, sourceFolder, itemKey
    End If
End Sub

Private Sub ProcessDeferredItemEvents()
    Dim keys As Variant
    Dim key As Variant
    Dim info As Variant
    Dim itemObject As Object
    Dim mail As Outlook.MailItem
    Dim sourceFolder As Outlook.Folder

    If gDeferredEvents Is Nothing Then Exit Sub
    If gDeferredEvents.Count = 0 Then Exit Sub

    keys = gDeferredEvents.Keys
    For Each key In keys
        info = gDeferredEvents(CStr(key))
        Set itemObject = Nothing
        On Error Resume Next
        Set itemObject = Application.Session.GetItemFromID(CStr(info(0)), CStr(info(1)))
        On Error GoTo 0
        If Not itemObject Is Nothing Then
            If TypeOf itemObject Is Outlook.MailItem Then
                Set mail = itemObject
                Set sourceFolder = mail.Parent
                HandleMailItem mail, sourceFolder, CStr(info(2))
            End If
        End If
    Next key
    gDeferredEvents.RemoveAll
End Sub

'===========================================================
' BASELINE: TRACK TAGGED EMAILS, NEVER ASSIGN UNTAGGED BACKLOG
'===========================================================

Private Function RunInitialBaseline() As Boolean
    Dim folders As Collection
    Dim folder As Outlook.Folder
    Dim scanned As Long
    Dim nowEgypt As Date
    Dim batchStarted As Boolean

    On Error GoTo BaselineFailed
    Set folders = GetMonitoredFolders()
    If folders.Count <> gExpectedFolderCount Then Exit Function

    If Not VWGX_BeginBatch() Then
        Err.Raise vbObjectError + 2009, "RunInitialBaseline", _
                  "Could not begin a workbook batch for the baseline."
    End If
    batchStarted = True

    For Each folder In folders
        scanned = scanned + BaselineFolder(folder)
    Next folder

    If Not GetEgyptNow(nowEgypt) Then
        Err.Raise vbObjectError + 2016, "RunInitialBaseline", _
                  "Egypt time-zone conversion failed; baseline cannot be marked complete."
    End If

    If Not VWGX_SetSetting("Baseline Status", "COMPLETE") Then
        Err.Raise vbObjectError + 2010, "RunInitialBaseline", _
                  "Could not set baseline status in Settings."
    End If
    If Not VWGX_SetSetting("Last Baseline Timestamp (Egypt)", nowEgypt) Then
        Err.Raise vbObjectError + 2012, "RunInitialBaseline", _
                  "Could not save the baseline timestamp."
    End If

    If Not VWGX_EndBatch(True) Then
        batchStarted = False
        Err.Raise vbObjectError + 2013, "RunInitialBaseline", _
                  "Workbook batch could not be saved. Baseline remains unconfirmed."
    End If
    batchStarted = False

    AppendRuntimeEvent "BASELINE_COMPLETE", "Four Inboxes", _
        "Items examined=" & CStr(scanned) & ". Existing tagged assignments were registered; untagged existing emails were left untouched."
    RunInitialBaseline = True
    Exit Function

BaselineFailed:
    If batchStarted Then
        On Error Resume Next
        Call VWGX_EndBatch(True)
        On Error GoTo 0
    End If
    AppendRuntimeEvent "BASELINE_FAILED", "Four Inboxes", Err.Description
    RunInitialBaseline = False
End Function

Private Function ScanAndQueueUnassignedBacklog() As Boolean
    Dim folders As Collection
    Dim folder As Outlook.Folder
    Dim items As Outlook.Items
    Dim obj As Object
    Dim mail As Outlook.MailItem
    Dim itemKey As String
    Dim matches As Variant
    Dim matchCount As Long
    Dim record As Variant
    Dim i As Long
    Dim queuedCount As Long
    Dim folderItemCount As Long
    Dim folderQueuedCount As Long
    Dim folderQueuedSummary As String
    Dim batchStarted As Boolean
    Dim nowEgypt As Date
    Dim scanStartedAt As Single
    Dim scanElapsed As Double
    Dim scanErrorNumber As Long
    Dim scanErrorDescription As String

    On Error GoTo BacklogFailed
    scanStartedAt = Timer

    If Not GetEgyptNow(nowEgypt) Then
        Err.Raise vbObjectError + 2017, "ScanAndQueueUnassignedBacklog", _
                  "Egypt time-zone conversion failed; backlog discovery is blocked."
    End If

    Set folders = GetMonitoredFolders()
    If folders.Count <> gExpectedFolderCount Then
        Err.Raise vbObjectError + 2018, "ScanAndQueueUnassignedBacklog", _
                  "Not all four monitored Inboxes are available."
    End If

    If Not VWGX_BeginBatch() Then
        Err.Raise vbObjectError + 2019, "ScanAndQueueUnassignedBacklog", _
                  "Could not begin the workbook batch."
    End If
    batchStarted = True

    For Each folder In folders
        folderItemCount = 0
        folderQueuedCount = 0
        Set items = folder.Items
        For i = items.Count To 1 Step -1
            folderItemCount = folderItemCount + 1
            Set obj = Nothing
            On Error Resume Next
            Set obj = items.Item(i)
            On Error GoTo BacklogFailed

            If Not obj Is Nothing Then
                If TypeOf obj Is Outlook.MailItem Then
                    Set mail = obj
                    itemKey = GetMessageKey(mail, folder)
                    If Len(itemKey) > 0 Then
                        If Not gRecords.Exists(itemKey) Then
                            matchCount = GetCategoryMatches(mail.Categories, matches)
                            If matchCount = 0 Then
                                record = NewAssignmentRecord(mail, folder, itemKey)
                                If UCase$(Trim$(CStr(record(AF_STATUS)))) = "NEEDS REVIEW" Then
                                    record(AF_STATUS) = "Needs Review"
                                    If SaveRecord(record) Then
                                        gRecords(itemKey) = record
                                        WriteActivity "BACKLOG_NEEDS_REVIEW", record, _
                                            "Existing unassigned email could not be safely timestamped; it was not queued for automatic assignment."
                                    End If
                                Else
                                    record(AF_STATUS) = "Queued"
                                    record(AF_LAST_VERIFIED) = nowEgypt
                                    record(AF_UPDATED_AT) = nowEgypt
                                    If Not SaveRecord(record) Then
                                        Err.Raise vbObjectError + 2020, "ScanAndQueueUnassignedBacklog", _
                                                  "Could not save a queued backlog record."
                                    End If
                                    gRecords(itemKey) = record
                                    queuedCount = queuedCount + 1
                                    folderQueuedCount = folderQueuedCount + 1
                                End If
                            ElseIf matchCount > 1 Then
                                ' Ambiguous mail is deliberately not auto-assigned.
                                MarkUntrackedAmbiguous mail, folder, itemKey, matches
                            End If
                        End If
                    End If
                End If
            End If

            If (i Mod 50) = 0 Then
                DoEvents
                If Not gStarted Or gStopping Or gStopRequested Then
                    Err.Raise vbObjectError + 2048, "ScanAndQueueUnassignedBacklog", _
                              "Dispatcher stop requested during backlog discovery."
                End If
            End If
        Next i
        folderQueuedSummary = folderQueuedSummary & vbCrLf & _
            folder.FolderPath & " | items examined=" & CStr(folderItemCount) & _
            " | newly queued=" & CStr(folderQueuedCount)
    Next folder

    If Not VWGX_EndBatch(False) Then
        batchStarted = False
        Err.Raise vbObjectError + 2021, "ScanAndQueueUnassignedBacklog", _
                  "Workbook batch could not be saved."
    End If
    batchStarted = False

    If queuedCount > 0 Then
        AppendRuntimeEvent "BACKLOG_DISCOVERY_COMPLETE", "Four Inboxes", _
            "All four monitored Inboxes scanned successfully. New untagged emails enrolled=" & CStr(queuedCount) & _
            ". Existing tracked records were not duplicated." & folderQueuedSummary
    End If
    scanElapsed = Timer - scanStartedAt
    If scanElapsed < 0 Then scanElapsed = scanElapsed + 86400#
    Debug.Print "[VWG BACKLOG] Four-Inbox discovery complete. New queued=" & CStr(queuedCount) & _
                " | duration=" & Format$(scanElapsed, "0.00") & "s" & folderQueuedSummary
    ScanAndQueueUnassignedBacklog = True
    Exit Function

BacklogFailed:
    scanErrorNumber = Err.Number
    scanErrorDescription = Err.Description
    If batchStarted Then
        On Error Resume Next
        Call VWGX_EndBatch(True)
        On Error GoTo 0
    End If
    AppendRuntimeEvent "BACKLOG_DISCOVERY_FAILED", "Four Inboxes", _
        CStr(scanErrorNumber) & " - " & scanErrorDescription
    ScanAndQueueUnassignedBacklog = False
End Function

Private Function ScanExistingCategorisedMessages() As Boolean
    Dim folders As Collection
    Dim folder As Outlook.Folder
    Dim batchStarted As Boolean

    On Error GoTo ScanFailed

    Set folders = GetMonitoredFolders()
    If folders.Count <> gExpectedFolderCount Then Exit Function

    If Not VWGX_BeginBatch() Then
        Err.Raise vbObjectError + 2014, "ScanExistingCategorisedMessages", _
                  "Could not begin workbook batch."
    End If
    batchStarted = True

    For Each folder In folders
        Call BaselineFolder(folder)
    Next folder

    If Not VWGX_EndBatch(True) Then
        batchStarted = False
        Err.Raise vbObjectError + 2015, "ScanExistingCategorisedMessages", _
                  "Workbook batch could not be saved."
    End If
    batchStarted = False

    ScanExistingCategorisedMessages = True
    Exit Function

ScanFailed:
    If batchStarted Then
        On Error Resume Next
        Call VWGX_EndBatch(True)
        On Error GoTo 0
    End If
    AppendRuntimeEvent "ASSIGNED_BASELINE_SCAN_FAILED", "Four Inboxes", Err.Description
End Function

Private Function BaselineFolder(ByVal folder As Outlook.Folder) As Long
    Dim items As Outlook.Items
    Dim obj As Object
    Dim mail As Outlook.MailItem
    Dim i As Long
    Dim itemKey As String
    Dim matches As Variant
    Dim matchCount As Long

    On Error GoTo FolderFailed
    Set items = folder.Items

    For i = items.Count To 1 Step -1
        Set obj = Nothing
        On Error Resume Next
        Set obj = items.Item(i)
        On Error GoTo FolderFailed

        If Not obj Is Nothing Then
            If TypeOf obj Is Outlook.MailItem Then
                Set mail = obj
                itemKey = GetMessageKey(mail, folder)
                If Len(itemKey) > 0 Then
                    matchCount = GetCategoryMatches(mail.Categories, matches)
                    If matchCount = 1 Then
                        If Not gRecords.Exists(itemKey) Then
                            RegisterExistingAssignment mail, folder, itemKey, CStr(matches(0)), "BASELINE_REGISTERED"
                        End If
                    ElseIf matchCount > 1 Then
                        If Not gRecords.Exists(itemKey) Then
                            MarkUntrackedAmbiguous mail, folder, itemKey, matches
                        End If
                    End If
                End If
                BaselineFolder = BaselineFolder + 1
            End If
        End If
        If (i Mod 100) = 0 Then
            DoEvents
            If Not gStarted Or gStopping Or gStopRequested Then
                Err.Raise vbObjectError + 2049, "BaselineFolder", _
                          "Dispatcher stop requested during initial folder baseline."
            End If
        End If
    Next i
    Exit Function

FolderFailed:
    AppendRuntimeEvent "BASELINE_FOLDER_FAILED", folder.FolderPath, Err.Number & " - " & Err.Description
    Err.Raise vbObjectError + 2011, "BaselineFolder", "A folder baseline scan failed: " & folder.FolderPath
End Function

'===========================================================
' REGISTER EXISTING / MANUAL ASSIGNMENTS
'===========================================================

Private Sub RegisterExistingAssignment(ByVal mail As Outlook.MailItem, _
                                       ByVal sourceFolder As Outlook.Folder, _
                                       ByVal itemKey As String, _
                                       ByVal categoryName As String, _
                                       ByVal eventType As String)
    Dim record As Variant
    Dim receivedUK As Date
    Dim receivedOk As Boolean
    Dim nowEgypt As Date

    If gRecords.Exists(itemKey) Then Exit Sub

    receivedOk = ConvertLocalDateToZone(mail.ReceivedTime, GetUKZoneId(), receivedUK)
    If Not receivedOk Then Exit Sub
    If Not GetEgyptNow(nowEgypt) Then Exit Sub

    record = NewAssignmentRecord(mail, sourceFolder, itemKey)
    If UCase$(Trim$(CStr(record(AF_STATUS)))) = "NEEDS REVIEW" Then
        If SaveRecord(record) Then gRecords(itemKey) = record
        WriteActivity "EXISTING_ASSIGNMENT_NEEDS_REVIEW", record, _
            "Timestamp conversion failed; existing category was left unchanged and requires manual review."
        Exit Sub
    End If
    record(AF_ORIGINAL) = categoryName
    record(AF_OWNER) = categoryName
    record(AF_FLAG) = "TRUE"
    record(AF_STATUS) = "Assigned"
    record(AF_LAST_VERIFIED) = nowEgypt
    record(AF_RECEIVED_UK) = receivedUK
    record(AF_ALLOCATED_AT_EGYPT) = Empty ' Existing/manual assignment is not a new automatic allocation.
    record(AF_ROSTER_PERIOD) = IIf(gRosterPeriodSet, gRosterPeriodStart, Empty)

    If Not SaveRecord(record) Then Exit Sub
    gRecords(itemKey) = record
    WriteActivity eventType, record, "Existing category registered without changing the email."
End Sub

Private Sub RegisterManualOwner(ByVal mail As Outlook.MailItem, _
                                ByVal sourceFolder As Outlook.Folder, _
                                ByVal itemKey As String, _
                                ByVal categoryName As String)
    Dim record As Variant
    Dim nowEgypt As Date
    Dim receivedUK As Date

    If Not GetEgyptNow(nowEgypt) Then Exit Sub
    If Not ConvertLocalDateToZone(mail.ReceivedTime, GetUKZoneId(), receivedUK) Then Exit Sub

    record = gRecords(itemKey)
    record(AF_ORIGINAL) = categoryName
    record(AF_OWNER) = categoryName
    record(AF_FLAG) = "TRUE"
    record(AF_STATUS) = "Assigned"
    record(AF_LAST_VERIFIED) = nowEgypt
    record(AF_PENDING_SINCE) = Empty
    record(AF_CATEGORY_STATE) = mail.Categories
    record(AF_RECEIVED_UK) = receivedUK
    record(AF_UPDATED_AT) = nowEgypt

    If Not SaveRecord(record) Then Exit Sub
    gRecords(itemKey) = record
    WriteActivity "MANUAL_ASSIGNMENT_CAPTURED", record, _
        "A colleague category appeared on a previously queued email; no automatic allocation counted."
End Sub

Private Sub MarkUntrackedAmbiguous(ByVal mail As Outlook.MailItem, _
                                   ByVal sourceFolder As Outlook.Folder, _
                                   ByVal itemKey As String, _
                                   ByVal matches As Variant)
    Dim record As Variant
    Dim nowEgypt As Date

    If Not GetEgyptNow(nowEgypt) Then Exit Sub
    record = NewAssignmentRecord(mail, sourceFolder, itemKey)
    record(AF_ORIGINAL) = ""
    record(AF_OWNER) = ""
    record(AF_FLAG) = ""
    record(AF_STATUS) = "Needs Review"
    record(AF_LAST_VERIFIED) = nowEgypt
    record(AF_CATEGORY_STATE) = mail.Categories

    If SaveRecord(record) Then
        gRecords(itemKey) = record
        WriteActivity "AMBIGUOUS_ASSIGNMENT", record, _
            "Multiple colleague categories on an untracked email: " & JoinVariant(matches, ", ") & ". No category changed."
    End If
End Sub

Private Sub MarkAmbiguous(ByVal record As Variant, _
                          ByVal mail As Outlook.MailItem, _
                          ByVal matches As Variant, _
                          ByVal note As String)
    Dim nowEgypt As Date
    If Not GetEgyptNow(nowEgypt) Then Exit Sub
    record(AF_STATUS) = "Needs Review"
    record(AF_PENDING_SINCE) = Empty
    record(AF_CATEGORY_STATE) = mail.Categories
    record(AF_LAST_VERIFIED) = nowEgypt
    record(AF_UPDATED_AT) = nowEgypt
    If SaveRecord(record) Then gRecords(CStr(record(AF_KEY))) = record
    WriteActivity "AMBIGUOUS_ASSIGNMENT", record, note & " Categories=" & JoinVariant(matches, ", ")
End Sub

'===========================================================
' NEW MESSAGE / QUEUE / AUTOMATIC ALLOCATION
'===========================================================

Private Sub HandleNewMail(ByVal mail As Outlook.MailItem, _
                          ByVal sourceFolder As Outlook.Folder, _
                          ByVal itemKey As String)
    On Error GoTo NewMailFailed
    If gRecords.Exists(itemKey) Then Exit Sub

    ' Every new ItemAdd is staged first. The next allocation pass performs a
    ' complete discovery scan of all four Inboxes, then releases the oldest
    ' eligible queued email across the four folders. This prevents a newer
    ' Detags arrival from overtaking an older email whose ItemAdd was missed.
    QueueNewMail mail, sourceFolder, itemKey, _
        "New arrival staged for oldest-first allocation after a complete four-Inbox backlog discovery."
    Exit Sub

NewMailFailed:
    AppendRuntimeEvent "NEW_MAIL_PROCESSING_FAILED", FolderPathSafe(sourceFolder), _
        Err.Number & " - " & Err.Description
End Sub

Private Sub QueueNewMail(ByVal mail As Outlook.MailItem, _
                         ByVal sourceFolder As Outlook.Folder, _
                         ByVal itemKey As String, _
                         ByVal reason As String)
    Dim record As Variant
    Dim nowEgypt As Date
    Dim receivedUK As Date

    If gRecords.Exists(itemKey) Then Exit Sub

    If Not GetEgyptNow(nowEgypt) Then
        AppendRuntimeEvent "QUEUE_BLOCKED_TIME_ZONE", FolderPathSafe(sourceFolder), _
            "New email was not assigned or queued because Egypt time could not be determined. Subject=" & mail.Subject
        Exit Sub
    End If
    If Not ConvertLocalDateToZone(mail.ReceivedTime, GetUKZoneId(), receivedUK) Then
        record = NewAssignmentRecord(mail, sourceFolder, itemKey)
        record(AF_STATUS) = "Needs Review"
        record(AF_CATEGORY_STATE) = "ReceivedTime could not be converted to UK time; not eligible for automatic allocation."
        If SaveRecord(record) Then
            gRecords(itemKey) = record
            WriteActivity "NEW_MAIL_NEEDS_REVIEW", record, reason & " Subject=" & mail.Subject
        End If
        Exit Sub
    End If

    record = NewAssignmentRecord(mail, sourceFolder, itemKey)
    If UCase$(Trim$(CStr(record(AF_STATUS)))) = "NEEDS REVIEW" Then
        If SaveRecord(record) Then
            gRecords(itemKey) = record
            WriteActivity "NEW_MAIL_NEEDS_REVIEW", record, reason & " Subject=" & mail.Subject
        End If
        Exit Sub
    End If

    record(AF_RECEIVED_UK) = receivedUK
    record(AF_STATUS) = "Queued"
    record(AF_LAST_VERIFIED) = nowEgypt
    record(AF_CATEGORY_STATE) = mail.Categories

    ' Queue rows are durable in DRY RUN too. The queue processor will only
    ' print its proposed owner in DRY RUN; it will never change Outlook categories.
    If SaveRecord(record) Then
        gRecords(itemKey) = record
        WriteActivity "QUEUED", record, reason & " Subject=" & mail.Subject
        Debug.Print "[VWG QUEUE] Staged | " & sourceFolder.FolderPath & _
                    " | Received (UK): " & Format$(receivedUK, "yyyy-mm-dd hh:nn:ss") & _
                    " | Subject: " & mail.Subject
    ElseIf VWGX_AssignmentFallbackContainsKey(CStr(record(AF_KEY))) Then
        ' The record is durable locally but is not a committed workbook reservation.
        ' Keep the category unchanged and let the next retry flush reconcile it.
        Debug.Print "[VWG QUEUE] Staged to durable local retry queue while tracker was busy | " & _
                    sourceFolder.FolderPath & " | Subject=" & mail.Subject
        AppendRuntimeEvent "QUEUE_STAGED_TO_RETRY", sourceFolder.FolderPath, _
            "Staged row is safely stored in the local assignment retry queue. No colleague category was changed. Subject=" & mail.Subject
    Else
        AppendRuntimeEvent "QUEUE_SAVE_FAILED", sourceFolder.FolderPath, _
            "The staged email was not confirmed in the workbook or local retry queue. No colleague category was changed. Subject=" & mail.Subject
    End If
End Sub

Private Function ChooseAssignee(ByVal nowEgypt As Date, _
                                ByVal folderPath As String, _
                                ByRef chosenName As String, _
                                ByRef chosenRosterRow As Long, _
                                ByRef reason As String) As Boolean
    Dim eligibleRows As Object
    Dim folderAllocationCounts As Object
    Dim allocationCounts As Object
    Dim workloadCounts As Object
    Dim minFolderAllocation As Long
    Dim minAllocation As Long
    Dim minWorkload As Long
    Dim rowNum As Variant
    Dim rosterCount As Long
    Dim startAt As Long
    Dim stepNum As Long
    Dim candidateRow As Long
    Dim categoryText As String
    Dim folderAllocationCount As Long
    Dim allocationCount As Long
    Dim workloadCount As Long
    Dim refreshedRoster As Variant

    chosenName = ""
    chosenRosterRow = 0
    reason = ""

    If Not gFairnessBatchActive Then
        refreshedRoster = VWGX_ReadRoster()
        If Not IsArray(refreshedRoster) Then
            reason = "Team & Shifts could not be refreshed; email remains queued."
            Exit Function
        End If
        gRoster = refreshedRoster

        EnsureFairnessPreviewDate nowEgypt
    End If

    Set eligibleRows = GetEligibleRosterRows(nowEgypt, reason)
    If eligibleRows Is Nothing Then Exit Function
    If eligibleRows.Count = 0 Then Exit Function

    If Not GetFolderCountsForSelection(folderPath, folderAllocationCounts) Then
        reason = "Historical auto-allocation counts for this Inbox could not be read safely; email remains queued."
        Exit Function
    End If

    If gFairnessBatchActive Then
        Set allocationCounts = gBatchDailyAllocationCounts
        Set workloadCounts = gBatchDailyWorkloadCounts
    Else
        Set allocationCounts = VWGX_GetAutoAllocationCountsForDate(DateValue(nowEgypt))
        If allocationCounts Is Nothing Then
            reason = "Today's all-folder allocation counts could not be read safely; email remains queued."
            Exit Function
        End If
        Set workloadCounts = VWGX_GetCurrentWorkloadCounts(DateValue(nowEgypt))
        If workloadCounts Is Nothing Then
            reason = "Current workload counts could not be read safely; email remains queued."
            Exit Function
        End If
        If IsDryRun() Then
            AddDictionaryCounts allocationCounts, gPreviewAllocationCounts
            AddDictionaryCounts workloadCounts, gPreviewWorkloadCounts
        End If
    End If
    If allocationCounts Is Nothing Or workloadCounts Is Nothing Then
        reason = "Daily fairness snapshot is unavailable; email remains queued."
        Exit Function
    End If

    ' IMPORTANT: assignment and workload counters are keyed by the exact Outlook
    ' category stored in Assignments, not the display-name column in Team & Shifts.
    minFolderAllocation = 2147483647
    For Each rowNum In eligibleRows.Keys
        categoryText = Trim$(CStr(gRoster(CLng(rowNum), 2)))
        folderAllocationCount = DictCount(folderAllocationCounts, categoryText)
        If folderAllocationCount < minFolderAllocation Then minFolderAllocation = folderAllocationCount
    Next rowNum

    minAllocation = 2147483647
    For Each rowNum In eligibleRows.Keys
        categoryText = Trim$(CStr(gRoster(CLng(rowNum), 2)))
        folderAllocationCount = DictCount(folderAllocationCounts, categoryText)
        allocationCount = DictCount(allocationCounts, categoryText)
        If folderAllocationCount = minFolderAllocation Then
            If allocationCount < minAllocation Then minAllocation = allocationCount
        End If
    Next rowNum

    minWorkload = 2147483647
    For Each rowNum In eligibleRows.Keys
        categoryText = Trim$(CStr(gRoster(CLng(rowNum), 2)))
        folderAllocationCount = DictCount(folderAllocationCounts, categoryText)
        allocationCount = DictCount(allocationCounts, categoryText)
        workloadCount = DictCount(workloadCounts, categoryText)
        If folderAllocationCount = minFolderAllocation And allocationCount = minAllocation Then
            If workloadCount < minWorkload Then minWorkload = workloadCount
        End If
    Next rowNum

    rosterCount = UBound(gRoster, 1) - LBound(gRoster, 1) + 1
    startAt = gNextTieBreakIndex
    If startAt < 1 Or startAt > rosterCount Then startAt = 1

    For stepNum = 0 To rosterCount - 1
        candidateRow = ((startAt - 1 + stepNum) Mod rosterCount) + 1
        If eligibleRows.Exists(CStr(candidateRow)) Then
            categoryText = Trim$(CStr(gRoster(candidateRow, 2)))
            folderAllocationCount = DictCount(folderAllocationCounts, categoryText)
            allocationCount = DictCount(allocationCounts, categoryText)
            workloadCount = DictCount(workloadCounts, categoryText)
            If folderAllocationCount = minFolderAllocation And _
               allocationCount = minAllocation And workloadCount = minWorkload Then
                chosenName = categoryText
                chosenRosterRow = candidateRow
                If Len(chosenName) > 0 Then
                    gLastChosenFolderAllocationCount = folderAllocationCount
                    gLastChosenDailyCount = allocationCount
                    gLastChosenWorkload = workloadCount
                    gLastEligiblePoolCount = eligibleRows.Count
                    Debug.Print "[VWG FAIRNESS] " & Format$(nowEgypt, "dd/mmm/yyyy hh:nn:ss") & _
                                " | Inbox=" & folderPath & _
                                " | eligible=" & CStr(eligibleRows.Count) & _
                                " | selected=" & chosenName & _
                                " | this-Inbox auto allocations before pick=" & CStr(folderAllocationCount) & _
                                " | all-Inbox allocations today before pick=" & CStr(allocationCount) & _
                                " | active workload before pick=" & CStr(workloadCount) & _
                                " | order=source-Inbox history, daily global allocations, active workload, rotating tie-break"
                    ChooseAssignee = True
                Else
                    reason = "An eligible colleague has no exact Outlook category in Team & Shifts."
                End If
                Exit Function
            End If
        End If
    Next stepNum

    reason = "No valid eligible agent could be selected."
End Function

Private Sub EnsureFairnessPreviewDate(ByVal nowEgypt As Date)
    If Not gFairnessDateSet Or DateValue(nowEgypt) <> gFairnessDate Then
        Set gPreviewAllocationCounts = CreateObject("Scripting.Dictionary")
        gPreviewAllocationCounts.CompareMode = vbTextCompare
        Set gPreviewWorkloadCounts = CreateObject("Scripting.Dictionary")
        gPreviewWorkloadCounts.CompareMode = vbTextCompare
        Set gPreviewFolderAllocationCounts = CreateObject("Scripting.Dictionary")
        gPreviewFolderAllocationCounts.CompareMode = vbTextCompare
        If gPreviewSeen Is Nothing Then
            Set gPreviewSeen = CreateObject("Scripting.Dictionary")
            gPreviewSeen.CompareMode = vbTextCompare
        Else
            gPreviewSeen.RemoveAll
        End If
        gFairnessDate = DateValue(nowEgypt)
        gFairnessDateSet = True
        Debug.Print "[VWG FAIRNESS] Preview counters rolled to Egypt-local date " & _
                    Format$(gFairnessDate, "dd/mmm/yyyy") & "."
    ElseIf gPreviewSeen Is Nothing Then
        Set gPreviewSeen = CreateObject("Scripting.Dictionary")
        gPreviewSeen.CompareMode = vbTextCompare
    End If
End Sub

Private Function PrepareFairnessBatch(ByVal nowEgypt As Date) As Boolean
    Dim refreshedRoster As Variant
    Dim allocationCounts As Object
    Dim workloadCounts As Object

    On Error GoTo PrepareFailed

    EnsureFairnessPreviewDate nowEgypt

    refreshedRoster = VWGX_ReadRoster()
    If Not IsArray(refreshedRoster) Then GoTo PrepareFailed
    gRoster = refreshedRoster

    Set allocationCounts = VWGX_GetAutoAllocationCountsForDate(DateValue(nowEgypt))
    If allocationCounts Is Nothing Then GoTo PrepareFailed
    Set workloadCounts = VWGX_GetCurrentWorkloadCounts(DateValue(nowEgypt))
    If workloadCounts Is Nothing Then GoTo PrepareFailed

    If IsDryRun() Then
        AddDictionaryCounts allocationCounts, gPreviewAllocationCounts
        AddDictionaryCounts workloadCounts, gPreviewWorkloadCounts
    End If

    ' Shift settings are also read once per allocation batch instead of four
    ' Excel Settings lookups per email.
    gBatchEarlyStart = TimeValue(CStr(VWGX_GetSetting("Early Shift Start (Egypt)")))
    gBatchEarlyEnd = TimeValue(CStr(VWGX_GetSetting("Early Shift End (Egypt)")))
    gBatchLateStart = TimeValue(CStr(VWGX_GetSetting("Late Shift Start (Egypt)")))
    gBatchLateEnd = TimeValue(CStr(VWGX_GetSetting("Late Shift End (Egypt)")))

    Set gBatchDailyAllocationCounts = allocationCounts
    Set gBatchDailyWorkloadCounts = workloadCounts
    gFairnessBatchActive = True
    PrepareFairnessBatch = True
    Exit Function

PrepareFailed:
    gFairnessBatchActive = False
    Set gBatchDailyAllocationCounts = Nothing
    Set gBatchDailyWorkloadCounts = Nothing
    Debug.Print "[VWG FAIRNESS] Could not prepare daily fairness snapshot | " & _
                CStr(Err.Number) & " - " & Err.Description
    PrepareFairnessBatch = False
End Function

Private Function GetCachedFolderAllocationCounts(ByVal folderPath As String, _
                                                   ByRef counts As Object) As Boolean
    On Error GoTo CacheFailed
    If Len(Trim$(folderPath)) = 0 Then Exit Function
    If gFolderAllocationCounts Is Nothing Then
        Set gFolderAllocationCounts = CreateObject("Scripting.Dictionary")
        gFolderAllocationCounts.CompareMode = vbTextCompare
    End If

    If gFolderAllocationCounts.Exists(folderPath) Then
        Set counts = gFolderAllocationCounts(folderPath)
    Else
        Set counts = VWGX_GetAutoAllocationCountsForFolder(folderPath)
        If counts Is Nothing Then Exit Function
        gFolderAllocationCounts.Add folderPath, counts
    End If

    GetCachedFolderAllocationCounts = Not (counts Is Nothing)
    Exit Function

CacheFailed:
    Debug.Print "[VWG FAIRNESS] Folder allocation cache read failed | " & folderPath & _
                " | " & Err.Number & " - " & Err.Description
    GetCachedFolderAllocationCounts = False
End Function

Private Sub IncrementFolderAllocationCache(ByVal folderPath As String, ByVal agentName As String)
    Dim counts As Object
    If Len(Trim$(folderPath)) = 0 Or Len(Trim$(agentName)) = 0 Then Exit Sub
    If Not GetCachedFolderAllocationCounts(folderPath, counts) Then Exit Sub
    IncrementDictionaryCount counts, agentName
End Sub

Private Sub IncrementPreviewFolderAllocationCache(ByVal folderPath As String, ByVal categoryName As String)
    Dim counts As Object
    If Len(Trim$(folderPath)) = 0 Or Len(Trim$(categoryName)) = 0 Then Exit Sub
    If gPreviewFolderAllocationCounts Is Nothing Then
        Set gPreviewFolderAllocationCounts = CreateObject("Scripting.Dictionary")
        gPreviewFolderAllocationCounts.CompareMode = vbTextCompare
    End If
    If gPreviewFolderAllocationCounts.Exists(folderPath) Then
        Set counts = gPreviewFolderAllocationCounts(folderPath)
    Else
        Set counts = CreateObject("Scripting.Dictionary")
        counts.CompareMode = vbTextCompare
        gPreviewFolderAllocationCounts.Add folderPath, counts
    End If
    IncrementDictionaryCount counts, categoryName
End Sub

Private Function GetFolderCountsForSelection(ByVal folderPath As String, _
                                             ByRef counts As Object) As Boolean
    Dim actualCounts As Object
    Dim previewCounts As Object
    Dim combined As Object
    Dim key As Variant

    If Not GetCachedFolderAllocationCounts(folderPath, actualCounts) Then Exit Function
    If Not IsDryRun() Then
        Set counts = actualCounts
        GetFolderCountsForSelection = True
        Exit Function
    End If

    ' Keep dry-run simulation deltas separate from committed historical counts.
    ' This prevents previews from contaminating later live fairness in the same session.
    Set combined = CreateObject("Scripting.Dictionary")
    combined.CompareMode = vbTextCompare
    For Each key In actualCounts.Keys
        combined(CStr(key)) = CLng(actualCounts(CStr(key)))
    Next key
    If Not gPreviewFolderAllocationCounts Is Nothing Then
        If gPreviewFolderAllocationCounts.Exists(folderPath) Then
            Set previewCounts = gPreviewFolderAllocationCounts(folderPath)
            AddDictionaryCounts combined, previewCounts
        End If
    End If
    Set counts = combined
    GetFolderCountsForSelection = True
End Function

Private Function GetEligibleRosterRows(ByVal nowEgypt As Date, ByRef reason As String) As Object
    Dim result As Object
    Dim t As Date
    Dim earlyStart As Date
    Dim earlyEnd As Date
    Dim lateStart As Date
    Dim lateEnd As Date
    Dim allowedGroup As String
    Dim r As Long
    Dim activeText As String
    Dim groupText As String
    Dim nameText As String
    Dim settingVal As Variant

    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare
    reason = ""

    ' Roster Period Start is historical metadata only. Daily fairness and
    ' shift eligibility do not depend on the fortnight start date.
    On Error GoTo InvalidTimes
    If gFairnessBatchActive Then
        earlyStart = gBatchEarlyStart
        earlyEnd = gBatchEarlyEnd
        lateStart = gBatchLateStart
        lateEnd = gBatchLateEnd
    Else
        earlyStart = TimeValue(CStr(VWGX_GetSetting("Early Shift Start (Egypt)")))
        earlyEnd = TimeValue(CStr(VWGX_GetSetting("Early Shift End (Egypt)")))
        lateStart = TimeValue(CStr(VWGX_GetSetting("Late Shift Start (Egypt)")))
        lateEnd = TimeValue(CStr(VWGX_GetSetting("Late Shift End (Egypt)")))
    End If
    t = TimeValue(nowEgypt)

    If t >= earlyStart And t < lateStart Then
        allowedGroup = "EARLY"
    ElseIf t >= lateStart And t < earlyEnd Then
        allowedGroup = "BOTH"
    ElseIf t >= earlyEnd And t < lateEnd Then
        allowedGroup = "LATE"
    Else
        reason = "Outside the staffed window (Egypt time). Email remains queued."
        Set GetEligibleRosterRows = result
        Exit Function
    End If

    For r = LBound(gRoster, 1) To UBound(gRoster, 1)
        nameText = Trim$(CStr(gRoster(r, 1)))
        groupText = UCase$(Trim$(CStr(gRoster(r, 3))))
        activeText = UCase$(Trim$(CStr(gRoster(r, 6))))

        If Len(nameText) > 0 And Len(Trim$(CStr(gRoster(r, 2)))) > 0 Then
            If activeText = "YES" Then
                If groupText = "EARLY" Or groupText = "LATE" Then
                    If allowedGroup = "BOTH" Or groupText = allowedGroup Then
                        result(CStr(r)) = True
                    End If
                End If
            End If
        End If
    Next r

    If result.Count = 0 Then
        reason = "No active eligible agents are configured for this Egypt-time window. Check Team & Shifts."
    End If

    Set GetEligibleRosterRows = result
    Exit Function

InvalidTimes:
    reason = "Shift times in Settings are missing or invalid. No assignment made."
    Set GetEligibleRosterRows = result
End Function

Private Sub AdvanceTieBreak(ByVal selectedRosterRow As Long, ByVal persistIt As Boolean)
    Dim rosterCount As Long
    If Not IsArray(gRoster) Then Exit Sub
    rosterCount = UBound(gRoster, 1) - LBound(gRoster, 1) + 1
    gNextTieBreakIndex = selectedRosterRow + 1
    If gNextTieBreakIndex > rosterCount Then gNextTieBreakIndex = 1
    If persistIt Then gTieBreakDirty = True
End Sub

Private Sub PersistTieBreakIfNeeded()
    If Not gTieBreakDirty Then Exit Sub
    If VWGX_SetSetting("Round Robin Tie-Break Index", gNextTieBreakIndex) Then
        gTieBreakDirty = False
        Debug.Print "[VWG PERFORMANCE] Tie-break position persisted once after the processing batch."
    Else
        Debug.Print "[VWG PERFORMANCE] Tie-break persistence deferred; in-memory position retained for this session."
    End If
End Sub

Private Sub AdvanceTieBreakForCategory(ByVal categoryName As String, ByVal persistIt As Boolean)
    Dim r As Long
    If Not IsArray(gRoster) Then Exit Sub
    For r = LBound(gRoster, 1) To UBound(gRoster, 1)
        If StrComp(Trim$(CStr(gRoster(r, 2))), Trim$(categoryName), vbTextCompare) = 0 Then
            AdvanceTieBreak r, persistIt
            Exit Sub
        End If
    Next r
    Debug.Print "[VWG] Tie-break reconciliation could not find roster category: " & categoryName
End Sub

Private Sub AddDictionaryCounts(ByVal targetCounts As Object, ByVal additionalCounts As Object)
    Dim key As Variant
    If targetCounts Is Nothing Or additionalCounts Is Nothing Then Exit Sub
    For Each key In additionalCounts.Keys
        If targetCounts.Exists(CStr(key)) Then
            targetCounts(CStr(key)) = CLng(targetCounts(CStr(key))) + CLng(additionalCounts(CStr(key)))
        Else
            targetCounts.Add CStr(key), CLng(additionalCounts(CStr(key)))
        End If
    Next key
End Sub

Private Sub IncrementDictionaryCount(ByVal counts As Object, ByVal keyText As String)
    If counts Is Nothing Or Len(keyText) = 0 Then Exit Sub
    If counts.Exists(keyText) Then
        counts(keyText) = CLng(counts(keyText)) + 1
    Else
        counts.Add keyText, 1
    End If
End Sub

Private Function DictCount(ByVal dict As Object, ByVal keyText As String) As Long
    On Error Resume Next
    If Not dict Is Nothing Then
        If dict.Exists(keyText) Then DictCount = CLng(dict(keyText))
    End If
    On Error GoTo 0
End Function

'===========================================================
' TRACKED EMAIL CHANGE: HANDOVER VS REMOVAL
'===========================================================

Private Sub ProcessTrackedMail(ByVal mail As Outlook.MailItem, _
                               ByVal sourceFolder As Outlook.Folder, _
                               ByVal itemKey As String, _
                               ByVal record As Variant)
    Dim matches As Variant
    Dim matchCount As Long
    Dim nowEgypt As Date
    Dim owner As String
    Dim statusText As String
    Dim eventCategories As String
    Dim freshObject As Object
    Dim freshMail As Outlook.MailItem

    owner = Trim$(CStr(record(AF_OWNER)))
    statusText = UCase$(Trim$(CStr(record(AF_STATUS))))

    If statusText = "CLOSED" Or statusText = "COMPLETED" Or statusText = "CANCELLED" Then Exit Sub
    If statusText = "NEEDS REVIEW" And Len(owner) = 0 Then Exit Sub

    ' ItemChange can deliver a stale MailItem snapshot shortly after our own
    ' category Save. Never interpret that event snapshot as a removal. Re-read
    ' the persisted item by the stable EntryID + StoreID before inspecting it.
    eventCategories = CStr(mail.Categories)
    Set freshObject = Nothing
    On Error Resume Next
    Set freshObject = Application.Session.GetItemFromID( _
        CStr(record(AF_ENTRY_ID)), CStr(record(AF_STORE_ID)))
    On Error GoTo FreshReadFailed

    If freshObject Is Nothing Then
        Debug.Print "[VWG] Category reconciliation skipped: fresh Outlook item unavailable. Key=" & itemKey
        Exit Sub
    End If
    If Not TypeOf freshObject Is Outlook.MailItem Then
        Debug.Print "[VWG] Category reconciliation skipped: item is not a MailItem. Key=" & itemKey
        Exit Sub
    End If

    Set freshMail = freshObject
    Set mail = freshMail
    Set sourceFolder = freshMail.Parent

    If StrComp(eventCategories, CStr(freshMail.Categories), vbBinaryCompare) <> 0 Then
        Debug.Print "[VWG] Ignored stale ItemChange category snapshot; fresh categories=" & _
                    CStr(freshMail.Categories) & " | Key=" & itemKey
    End If

    matchCount = GetCategoryMatches(freshMail.Categories, matches)
    If Not GetEgyptNow(nowEgypt) Then Exit Sub

    If matchCount = 1 Then
        If (statusText = "ASSIGNING" Or _
            (statusText = "NEEDS REVIEW" And _
             InStr(1, CStr(record(AF_CATEGORY_STATE)), "Assignment category could not be confirmed", vbTextCompare) > 0)) _
           And StrComp(CStr(matches(0)), owner, vbTextCompare) = 0 Then
            ' Outlook now confirms the intended owner category, even if the first
            ' assignment save/status update was uncertain. Reconcile this same
            ' reservation rather than creating a second allocation.
            record(AF_STATUS) = "Assigned"
            record(AF_CATEGORY_STATE) = freshMail.Categories
            record(AF_LAST_VERIFIED) = nowEgypt
            record(AF_UPDATED_AT) = nowEgypt
            If SaveRecord(record) Then
                gRecords(itemKey) = record
                If WriteActivity("ASSIGNMENT_STATE_RECONCILED", record, _
                    "Fresh Outlook read confirmed the intended owner category after an interrupted assignment update. Existing reservation retained; no second allocation created.") Then
                    AdvanceTieBreakForCategory owner, True
                Else
                    Debug.Print "[VWG] Reconciled assignment but Activity Log write failed. Key=" & itemKey
                End If
                    Else
                Debug.Print "[VWG] Assignment category is present but Excel could not reconcile status. Key=" & itemKey
            End If
            Exit Sub
        ElseIf Len(owner) = 0 Then
            RegisterManualOwner freshMail, sourceFolder, itemKey, CStr(matches(0))
        ElseIf StrComp(CStr(matches(0)), owner, vbTextCompare) = 0 Then
            ' Clean marker categories left by an earlier dispatcher build.
            ' They are not used as Outlook workflow categories in this version;
            ' the recovery incident remains recorded in the workbook.
            If HasRestorationMarker(freshMail.Categories) Then
                If IsDryRun() Then
                    Debug.Print "[VWG DRY RUN] Would remove legacy Assignment Restored marker from " & itemKey
                    WriteActivity "DRY_RUN_WOULD_REMOVE_RESTORATION_MARKER", record, _
                        "The owner category is present. A legacy Assignment Restored marker would be removed; no Outlook category was changed."
                ElseIf RemoveRestorationMarkerCategories(freshMail, owner) Then
                    record(AF_STATUS) = "Recovered"
                    record(AF_PENDING_SINCE) = Empty
                    record(AF_CATEGORY_STATE) = freshMail.Categories
                    record(AF_LAST_VERIFIED) = nowEgypt
                    record(AF_UPDATED_AT) = nowEgypt
                    If Not SaveRecord(record) Then
                        Debug.Print "[VWG] Legacy restoration marker removed but tracker could not be updated. Key=" & itemKey
                        Exit Sub
                    End If
                    gRecords(itemKey) = record
                    statusText = "RECOVERED"
                    WriteActivity "RESTORATION_MARKER_REMOVED", record, _
                        "Legacy Assignment Restored marker removed. The email remains categorized only by its current owner; recovery history stays in the tracker."
                Else
                    record(AF_STATUS) = "Recovery Failed"
                    record(AF_UPDATED_AT) = nowEgypt
                    If SaveRecord(record) Then gRecords(itemKey) = record
                    WriteActivity "RECOVERY_FAILED", record, _
                        "Owner category is present but a legacy Assignment Restored marker could not be removed. Retry required."
                    Exit Sub
                End If
            End If

            If Not HasRestorationMarker(freshMail.Categories) And _
               (statusText = "ASSIGNED" Or statusText = "HANDED OVER" Or statusText = "RECOVERED") And _
               StrComp(CStr(record(AF_CATEGORY_STATE)), CStr(freshMail.Categories), vbBinaryCompare) = 0 Then
                ' Outlook can raise ItemChange for metadata unrelated to dispatcher
                ' categories. If the owner and category snapshot are unchanged,
                ' do not perform a redundant Excel upsert/save.
                Exit Sub
            End If

            If statusText = "DRY RUN - RECOVERY PENDING" Then
                record(AF_STATUS) = IIf(CLng(record(AF_HANDOVER_COUNT)) > 0, "Handed Over", "Assigned")
                record(AF_PENDING_SINCE) = Empty
                record(AF_CATEGORY_STATE) = freshMail.Categories
                record(AF_LAST_VERIFIED) = nowEgypt
                record(AF_UPDATED_AT) = nowEgypt
                If SaveRecord(record) Then gRecords(itemKey) = record
                WriteActivity "DRY_RUN_PENDING_CLEARED", record, _
                    "Current-owner category is present again; simulated missing-category state cleared without a confirmed incident."
            ElseIf statusText = "HANDOVER PENDING" Then
                record(AF_STATUS) = IIf(CLng(record(AF_HANDOVER_COUNT)) > 0, "Handed Over", "Assigned")
                record(AF_PENDING_SINCE) = Empty
                record(AF_CATEGORY_STATE) = freshMail.Categories
                record(AF_LAST_VERIFIED) = nowEgypt
                record(AF_UPDATED_AT) = nowEgypt
                If SaveRecord(record) Then
                    gRecords(itemKey) = record
                    WriteActivity "PENDING_CHANGE_CANCELLED", record, _
                        "Original/current owner's category is present on the freshly read email; no incident counted."
                End If
            Else
                ' Migrate the previous sheet-only "Recovered - Review" state
                ' to the simpler Recovered status. The row and activity history
                ' remain intact; only the status label is normalized.
                If statusText = "RECOVERED - REVIEW" And Not IsDryRun() Then record(AF_STATUS) = "Recovered"
                record(AF_LAST_VERIFIED) = nowEgypt
                record(AF_CATEGORY_STATE) = freshMail.Categories
                record(AF_UPDATED_AT) = nowEgypt
                If SaveRecord(record) Then gRecords(itemKey) = record
            End If
        Else
            RecordLegitimateHandover freshMail, sourceFolder, itemKey, record, CStr(matches(0)), nowEgypt
        End If

    ElseIf matchCount = 0 Then
        If statusText = "DRY RUN - RECOVERY PENDING" And IsDryRun() Then Exit Sub

        If statusText = "ASSIGNING" Then
            ' An Assigning row is an allocation reservation, not a completed
            ' assignment and must not be counted as a category-removal incident.
            ' Allow the initial Outlook store update to settle, then retry the
            ' SAME intended owner once using the fresh-item save routine.
            If IsDate(record(AF_LAST_VERIFIED)) Then
                If DateDiff("s", CDate(record(AF_LAST_VERIFIED)), nowEgypt) < 45 Then Exit Sub
            End If

            If Len(owner) > 0 And ApplyCategoryToMail(freshMail, owner) Then
                record(AF_STATUS) = "Assigned"
                record(AF_CATEGORY_STATE) = freshMail.Categories
                record(AF_LAST_VERIFIED) = nowEgypt
                record(AF_UPDATED_AT) = nowEgypt
                If SaveRecord(record) Then
                    gRecords(itemKey) = record
                    If WriteActivity("ASSIGNMENT_STATE_RECONCILED", record, _
                        "The original assignment reservation was reconciled by applying and verifying the same intended owner category. No different colleague was allocated.") Then
                        AdvanceTieBreakForCategory owner, True
                    Else
                        Debug.Print "[VWG] Reconciled same-owner assignment but Activity Log write failed. Key=" & itemKey
                    End If
                            End If
            Else
                record(AF_STATUS) = "Needs Review"
                record(AF_PENDING_SINCE) = Empty
                record(AF_CATEGORY_STATE) = "Assignment category could not be confirmed for intended owner: " & owner
                record(AF_UPDATED_AT) = nowEgypt
                If SaveRecord(record) Then
                    gRecords(itemKey) = record
                    WriteActivity "ASSIGNMENT_CONFIRMATION_FAILED", record, _
                        "The reserved owner category was still absent after the stabilization window and one controlled same-owner retry. No removal incident was counted and no different colleague was allocated. Manual review required."
                End If
            End If
            Exit Sub
        End If

        ' A Needs Review reservation with this explicit note is an uncertain
        ' initial assignment, not evidence of colleague-category removal.
        If statusText = "NEEDS REVIEW" Then
            If InStr(1, CStr(record(AF_CATEGORY_STATE)), "Assignment category could not be confirmed", vbTextCompare) > 0 Then Exit Sub
        End If

        ' A missing category immediately after a script write may be a delayed
        ' shared-store update. Do not count an incident yet, but do schedule a
        ' fresh read after the normal handover grace period; never silently drop
        ' the ItemChange event.
        If (statusText = "ASSIGNED" Or statusText = "HANDED OVER" Or statusText = "RECOVERED" Or statusText = "RECOVERED - REVIEW") Then
            If IsDate(record(AF_LAST_VERIFIED)) Then
                If DateDiff("s", CDate(record(AF_LAST_VERIFIED)), nowEgypt) < ASSIGNMENT_STABILITY_SECONDS Then
                    record(AF_STATUS) = "Handover Pending"
                    record(AF_PENDING_SINCE) = nowEgypt
                    record(AF_CATEGORY_STATE) = freshMail.Categories
                    record(AF_LAST_VERIFIED) = nowEgypt
                    record(AF_UPDATED_AT) = nowEgypt
                    If SaveRecord(record) Then
                        gRecords(itemKey) = record
                        WriteActivity "CATEGORY_REMOVAL_VERIFY_PENDING", record, _
                            "Fresh read found no colleague category shortly after a dispatcher category write. No incident counted yet; the message will be rechecked before recovery."
                    End If
                    Exit Sub
                End If
            End If
        End If

        If statusText <> "HANDOVER PENDING" And statusText <> "RECOVERY FAILED" Then
            record(AF_STATUS) = "Handover Pending"
            record(AF_PENDING_SINCE) = nowEgypt
            record(AF_LAST_VERIFIED) = nowEgypt
            record(AF_CATEGORY_STATE) = freshMail.Categories
            record(AF_UPDATED_AT) = nowEgypt
            If SaveRecord(record) Then
                gRecords(itemKey) = record
                WriteActivity "CATEGORY_REMOVAL_PENDING", record, _
                    "No colleague category found on a freshly reacquired Outlook item. Waiting " & _
                    CStr(gHandoverGraceSeconds) & " seconds for a legitimate handover before restoration. Subject=" & freshMail.Subject
            End If
        End If

    Else
        ' More than one agent category. This could be a handover in progress.
        If statusText <> "HANDOVER PENDING" Then
            record(AF_STATUS) = "Handover Pending"
            record(AF_PENDING_SINCE) = nowEgypt
            record(AF_CATEGORY_STATE) = freshMail.Categories
            record(AF_LAST_VERIFIED) = nowEgypt
            record(AF_UPDATED_AT) = nowEgypt
            If SaveRecord(record) Then
                gRecords(itemKey) = record
                WriteActivity "MULTIPLE_CATEGORIES_PENDING", record, _
                    "Multiple colleague categories detected on a freshly reacquired email; allow a short handover grace period. Categories=" & JoinVariant(matches, ", ")
            End If
        End If
    End If
    Exit Sub

FreshReadFailed:
    Debug.Print "[VWG] Category reconciliation skipped after fresh-read error: " & _
                Err.Number & " - " & Err.Description & " | Key=" & itemKey
End Sub

Private Sub RecordLegitimateHandover(ByVal mail As Outlook.MailItem, _
                                     ByVal sourceFolder As Outlook.Folder, _
                                     ByVal itemKey As String, _
                                     ByVal record As Variant, _
                                     ByVal newOwner As String, _
                                     ByVal nowEgypt As Date)
    Dim oldOwner As String
    Dim note As String

    oldOwner = CStr(record(AF_OWNER))
    If Len(oldOwner) = 0 Or StrComp(oldOwner, newOwner, vbTextCompare) = 0 Then Exit Sub

    note = Format$(nowEgypt, "yyyy-mm-dd hh:nn:ss") & ": " & oldOwner & " -> " & newOwner
    If Len(CStr(record(AF_HANDOVER_NOTES))) > 0 Then
        record(AF_HANDOVER_NOTES) = CStr(record(AF_HANDOVER_NOTES)) & " | " & note
    Else
        record(AF_HANDOVER_NOTES) = note
    End If

    record(AF_OWNER) = newOwner
    record(AF_STATUS) = "Handed Over"
    record(AF_PENDING_SINCE) = Empty
    record(AF_CATEGORY_STATE) = mail.Categories
    record(AF_LAST_VERIFIED) = nowEgypt
    record(AF_UPDATED_AT) = nowEgypt
    record(AF_HANDOVER_COUNT) = CLng(record(AF_HANDOVER_COUNT)) + 1

    ' Original owner and allocated timestamp never change for a handover.
    If SaveRecord(record) Then
        gRecords(itemKey) = record
        WriteActivity "HANDOVER", record, _
            "Legitimate category handover: " & oldOwner & " -> " & newOwner & _
            ". Not counted as abuse and does not advance the assignment rotation."
    End If
End Sub

Private Function NextRecordStartIndex(ByVal keys As Variant, ByVal cursorKey As String) As Long
    Dim i As Long
    Dim total As Long
    total = UBound(keys) - LBound(keys) + 1
    If total <= 0 Then Exit Function
    If Len(cursorKey) = 0 Then Exit Function
    For i = LBound(keys) To UBound(keys)
        If StrComp(CStr(keys(i)), cursorKey, vbTextCompare) = 0 Then
            NextRecordStartIndex = (i - LBound(keys) + 1) Mod total
            Exit Function
        End If
    Next i
End Function

Private Sub ProcessPendingRecoveries()
    Const MAX_RECOVERY_RECORDS_SCANNED_PER_CYCLE As Long = 250
    Dim keys As Variant
    Dim key As Variant
    Dim record As Variant
    Dim itemObject As Object
    Dim mail As Outlook.MailItem
    Dim pendingSince As Date
    Dim nowEgypt As Date
    Dim matches As Variant
    Dim matchCount As Long
    Dim grace As Long
    Dim markerClean As Boolean
    Dim total As Long
    Dim startIndex As Long
    Dim scanOffset As Long
    Dim idx As Long
    Dim actionsProcessed As Long
    Dim scannedCount As Long
    Dim currentKey As String
    Dim itemBusyOwned As Boolean
    Dim failureNumber As Long
    Dim failureDescription As String

    If gRecords Is Nothing Then Exit Sub
    If gRecords.Count = 0 Then Exit Sub
    If Not GetEgyptNow(nowEgypt) Then Exit Sub

    grace = gHandoverGraceSeconds
    If grace < 5 Then grace = DEFAULT_HANDOVER_GRACE_SECONDS
    keys = gRecords.Keys
    total = UBound(keys) - LBound(keys) + 1
    If total <= 0 Then Exit Sub
    startIndex = NextRecordStartIndex(keys, gRecoveryCursorKey)

    On Error GoTo RecoveryFailed
    For scanOffset = 0 To total - 1
        If actionsProcessed >= MAX_RECOVERY_ACTIONS_PER_CYCLE Then Exit For
        If scannedCount >= MAX_RECOVERY_RECORDS_SCANNED_PER_CYCLE Then Exit For
        If VWGX_AssignmentFallbackPending() Then Exit For
        idx = LBound(keys) + ((startIndex + scanOffset) Mod total)
        currentKey = CStr(keys(idx))
        If Not gRecords.Exists(currentKey) Then GoTo NextRecoveryRecord
        If gBusy.Exists(currentKey) Then GoTo NextRecoveryRecord
        record = gRecords(currentKey)

        If UCase$(Trim$(CStr(record(AF_STATUS)))) = "HANDOVER PENDING" Then
            If IsDate(record(AF_PENDING_SINCE)) Then
                pendingSince = CDate(record(AF_PENDING_SINCE))
                If DateDiff("s", pendingSince, nowEgypt) < grace Then GoTo NextRecoveryRecord
            End If
            actionsProcessed = actionsProcessed + 1
            gBusy.Add currentKey, True
            itemBusyOwned = True

            If Not IsDate(record(AF_PENDING_SINCE)) Then
                record(AF_PENDING_SINCE) = nowEgypt
                If SaveRecord(record) Then gRecords(currentKey) = record
                GoTo NextRecoveryRecord
            End If

            Set itemObject = Nothing
            On Error Resume Next
            Set itemObject = Application.Session.GetItemFromID( _
                CStr(record(AF_ENTRY_ID)), CStr(record(AF_STORE_ID)))
            On Error GoTo RecoveryFailed

            If itemObject Is Nothing Then
                record(AF_STATUS) = "Needs Review"
                record(AF_PENDING_SINCE) = Empty
                record(AF_UPDATED_AT) = nowEgypt
                If SaveRecord(record) Then gRecords(currentKey) = record
                WriteActivity "PENDING_ITEM_NOT_FOUND", record, _
                    "Pending email could not be retrieved. No category was changed."
            ElseIf TypeOf itemObject Is Outlook.MailItem Then
                Set mail = itemObject
                matchCount = GetCategoryMatches(mail.Categories, matches)
                If matchCount = 1 Then
                    If StrComp(CStr(matches(0)), CStr(record(AF_OWNER)), vbTextCompare) = 0 Then
                        markerClean = True
                        If HasRestorationMarker(mail.Categories) Then
                            markerClean = RemoveRestorationMarkerCategories(mail, CStr(record(AF_OWNER)))
                        End If
                        If markerClean Then
                            record(AF_STATUS) = IIf(CLng(record(AF_HANDOVER_COUNT)) > 0, "Handed Over", "Assigned")
                            record(AF_PENDING_SINCE) = Empty
                            record(AF_CATEGORY_STATE) = mail.Categories
                            record(AF_LAST_VERIFIED) = nowEgypt
                            record(AF_UPDATED_AT) = nowEgypt
                            If SaveRecord(record) Then gRecords(currentKey) = record
                            WriteActivity "PENDING_CHANGE_CANCELLED", record, _
                                "Current owner's category is present; no incident counted. Any legacy Assignment Restored marker was removed."
                        Else
                            record(AF_STATUS) = "Recovery Failed"
                            record(AF_UPDATED_AT) = nowEgypt
                            If SaveRecord(record) Then gRecords(currentKey) = record
                            WriteActivity "RECOVERY_FAILED", record, _
                                "Current owner category is present, but a legacy Assignment Restored marker could not be removed."
                        End If
                    Else
                        RecordLegitimateHandover mail, mail.Parent, currentKey, record, CStr(matches(0)), nowEgypt
                    End If
                ElseIf matchCount = 0 Then
                    RecoverMissingCategory mail, currentKey, record, nowEgypt
                Else
                    record(AF_STATUS) = "Needs Review"
                    record(AF_PENDING_SINCE) = Empty
                    record(AF_CATEGORY_STATE) = mail.Categories
                    record(AF_UPDATED_AT) = nowEgypt
                    If SaveRecord(record) Then gRecords(currentKey) = record
                    WriteActivity "AMBIGUOUS_CATEGORIES_REVIEW", record, _
                        "Multiple colleague categories remain after the grace period. No automatic recovery was applied."
                End If
            End If

        ElseIf UCase$(Trim$(CStr(record(AF_STATUS)))) = "RECOVERY FAILED" Then
            actionsProcessed = actionsProcessed + 1
            gBusy.Add currentKey, True
            itemBusyOwned = True
            Set itemObject = Nothing
            On Error Resume Next
            Set itemObject = Application.Session.GetItemFromID( _
                CStr(record(AF_ENTRY_ID)), CStr(record(AF_STORE_ID)))
            On Error GoTo RecoveryFailed
            If Not itemObject Is Nothing Then
                If TypeOf itemObject Is Outlook.MailItem Then
                    Set mail = itemObject
                    matchCount = GetCategoryMatches(mail.Categories, matches)
                    If matchCount = 0 Then
                        AttemptCategoryRestore mail, currentKey, record, nowEgypt, False
                    ElseIf matchCount = 1 And _
                           StrComp(CStr(matches(0)), CStr(record(AF_OWNER)), vbTextCompare) = 0 Then
                        If IsDryRun() Then
                            WriteActivity "DRY_RUN_WOULD_REMOVE_RESTORATION_MARKER", record, _
                                "The original owner category is present. A legacy Assignment Restored marker would be removed; no Outlook category was changed."
                        ElseIf RemoveRestorationMarkerCategories(mail, CStr(record(AF_OWNER))) Then
                            record(AF_STATUS) = "Recovered"
                            record(AF_PENDING_SINCE) = Empty
                            record(AF_CATEGORY_STATE) = mail.Categories
                            record(AF_LAST_VERIFIED) = nowEgypt
                            record(AF_UPDATED_AT) = nowEgypt
                            If SaveRecord(record) Then
                                gRecords(currentKey) = record
                                WriteActivity "RECOVERY_STATE_RECONCILED", record, _
                                    "The original owner category is present; any legacy restoration marker was removed. Recovery remains recorded in the tracker only."
                            End If
                        Else
                            record(AF_STATUS) = "Recovery Failed"
                            record(AF_UPDATED_AT) = nowEgypt
                            If SaveRecord(record) Then gRecords(currentKey) = record
                            WriteActivity "RECOVERY_FAILED", record, _
                                "Owner category is present but a legacy restoration marker could not be removed. Retry required."
                        End If
                    Else
                        record(AF_STATUS) = "Needs Review"
                        record(AF_PENDING_SINCE) = Empty
                        record(AF_CATEGORY_STATE) = mail.Categories
                        record(AF_UPDATED_AT) = nowEgypt
                        If SaveRecord(record) Then gRecords(currentKey) = record
                        WriteActivity "RECOVERY_REQUIRES_REVIEW", record, _
                            "A different or ambiguous colleague category is present after a failed recovery. No category was changed."
                    End If
                End If
            End If
        End If
NextRecoveryRecord:
        If itemBusyOwned Then
            If gBusy.Exists(currentKey) Then gBusy.Remove currentKey
            itemBusyOwned = False
        End If
        scannedCount = scannedCount + 1
        gRecoveryCursorKey = currentKey
    Next scanOffset

    If scannedCount > 0 Then
        Debug.Print "[VWG PERFORMANCE] Recovery slice scanned=" & CStr(scannedCount) & _
                    " | actions=" & CStr(actionsProcessed) & _
                    " | scan cap=" & CStr(MAX_RECOVERY_RECORDS_SCANNED_PER_CYCLE) & _
                    " | action cap=" & CStr(MAX_RECOVERY_ACTIONS_PER_CYCLE)
    End If
    Exit Sub

RecoveryFailed:
    failureNumber = Err.Number
    failureDescription = Err.Description
    On Error Resume Next
    If itemBusyOwned Then
        If gBusy.Exists(currentKey) Then gBusy.Remove currentKey
    End If
    On Error GoTo 0
    Debug.Print "[VWG PERFORMANCE] Recovery processing deferred | Key=" & currentKey & _
                " | " & CStr(failureNumber) & " - " & failureDescription
End Sub

Private Sub RecoverMissingCategory(ByVal mail As Outlook.MailItem, _
                                   ByVal itemKey As String, _
                                   ByVal record As Variant, _
                                   ByVal nowEgypt As Date)
    Dim incidentNo As Long

    If IsDryRun() Then
        record(AF_STATUS) = "DRY RUN - RECOVERY PENDING"
        record(AF_PENDING_SINCE) = Empty
        record(AF_UPDATED_AT) = nowEgypt
        If SaveRecord(record) Then gRecords(itemKey) = record
        WriteActivity "DRY_RUN_WOULD_RESTORE", record, _
            "Missing category would be restored to the current owner after the handover grace period. No email category changed; no incident counted as confirmed."
        Debug.Print "[VWG DRY RUN] Would restore " & CStr(record(AF_OWNER)) & " | " & CStr(record(AF_FOLDER))
        Exit Sub
    End If

    incidentNo = CLng(record(AF_INCIDENTS)) + 1
    record(AF_INCIDENTS) = incidentNo
    record(AF_FLAG) = "FALSE"
    record(AF_LAST_INCIDENT) = nowEgypt
    record(AF_PENDING_SINCE) = Empty
    record(AF_STATUS) = "Recovery Failed" ' safe retry state until verified
    record(AF_UPDATED_AT) = nowEgypt

    If Not SaveRecord(record) Then
        AppendRuntimeEvent "RECOVERY_BLOCKED", CStr(record(AF_FOLDER)), _
            "Could not save incident state. No category was changed."
        Exit Sub
    End If
    gRecords(itemKey) = record

    WriteActivity "CATEGORY_REMOVAL_DETECTED", record, _
        "Assignment category absent after grace period. Restoring current owner. User identity is not inferred."

    AttemptCategoryRestore mail, itemKey, record, nowEgypt, True
End Sub

Private Sub AttemptCategoryRestore(ByVal mail As Outlook.MailItem, _
                                   ByVal itemKey As String, _
                                   ByVal record As Variant, _
                                   ByVal nowEgypt As Date, _
                                   ByVal logAttempt As Boolean)
    Dim ownerName As String
    Dim oldBusy As Boolean
    Dim ownerApplied As Boolean
    Dim markerCleanupApplied As Boolean

    On Error GoTo RestoreFailed

    ownerName = Trim$(CStr(record(AF_OWNER)))
    If Len(ownerName) = 0 Then
        record(AF_STATUS) = "Needs Review"
        If SaveRecord(record) Then gRecords(itemKey) = record
        WriteActivity "RECOVERY_FAILED", record, "Current owner is blank; manual review required."
        Exit Sub
    End If

    If Not gBusy.Exists(itemKey) Then
        gBusy.Add itemKey, True
        oldBusy = True
    End If

    ' Restore only the current owner's colleague category. Do not add a
    ' separate Assignment Restored category. Remove legacy restoration
    ' markers from earlier versions, preserving all unrelated categories.
    ownerApplied = ApplyCategoryToMail(mail, ownerName)
    If ownerApplied Then markerCleanupApplied = RemoveRestorationMarkerCategories(mail, ownerName)

    If ownerApplied And markerCleanupApplied And HasCategory(mail, ownerName) And _
       Not HasRestorationMarker(mail.Categories) Then
        record(AF_STATUS) = "Recovered"
        record(AF_PENDING_SINCE) = Empty
        record(AF_CATEGORY_STATE) = mail.Categories
        record(AF_LAST_VERIFIED) = nowEgypt
        record(AF_UPDATED_AT) = nowEgypt
        If SaveRecord(record) Then
            gRecords(itemKey) = record
            WriteActivity "CATEGORY_RESTORED", record, _
                "Original current-owner category restored and verified. No separate restoration category was added; recovery incident is recorded in the tracker. Rotation unchanged."
            Else
            Debug.Print "[VWG] Owner category is restored but tracker update failed. Key=" & itemKey
        End If
    Else
        record(AF_STATUS) = "Recovery Failed"
        record(AF_UPDATED_AT) = nowEgypt
        If SaveRecord(record) Then gRecords(itemKey) = record
        WriteActivity "RECOVERY_FAILED", record, _
            "Outlook did not confirm the owner category and removal of legacy restoration markers. The timer may retry without double-counting."
    End If

    If oldBusy Then
        If gBusy.Exists(itemKey) Then gBusy.Remove itemKey
    End If
    Exit Sub

RestoreFailed:
    If oldBusy Then
        On Error Resume Next
        If gBusy.Exists(itemKey) Then gBusy.Remove itemKey
        On Error GoTo 0
    End If
    record(AF_STATUS) = "Recovery Failed"
    record(AF_UPDATED_AT) = nowEgypt
    If SaveRecord(record) Then gRecords(itemKey) = record
    WriteActivity "RECOVERY_FAILED", record, Err.Number & " - " & Err.Description
End Sub

'===========================================================
' RETRY PERSISTED ASSIGNING RESERVATIONS BEFORE NEW QUEUE ALLOCATIONS
'===========================================================

Private Sub ProcessPendingAssignmentReservations()
    Const MAX_RESERVATION_RECORDS_SCANNED_PER_CYCLE As Long = 250
    Dim keys As Variant
    Dim key As Variant
    Dim record As Variant
    Dim itemObject As Object
    Dim mail As Outlook.MailItem
    Dim sourceFolder As Outlook.Folder
    Dim nowEgypt As Date
    Dim statusText As String
    Dim total As Long
    Dim startIndex As Long
    Dim scanOffset As Long
    Dim idx As Long
    Dim retriesProcessed As Long
    Dim scannedCount As Long
    Dim currentKey As String
    Dim failureNumber As Long
    Dim failureDescription As String

    If gRecords Is Nothing Then Exit Sub
    If gRecords.Count = 0 Then Exit Sub
    If VWGX_AssignmentFallbackPending() Then Exit Sub
    If Not GetEgyptNow(nowEgypt) Then Exit Sub

    keys = gRecords.Keys
    total = UBound(keys) - LBound(keys) + 1
    If total <= 0 Then Exit Sub
    startIndex = NextRecordStartIndex(keys, gReservationCursorKey)

    On Error GoTo ReservationFailed
    For scanOffset = 0 To total - 1
        If retriesProcessed >= MAX_RESERVATION_RETRIES_PER_CYCLE Then Exit For
        If scannedCount >= MAX_RESERVATION_RECORDS_SCANNED_PER_CYCLE Then Exit For
        If VWGX_AssignmentFallbackPending() Then Exit For
        idx = LBound(keys) + ((startIndex + scanOffset) Mod total)
        currentKey = CStr(keys(idx))
        If Not gRecords.Exists(currentKey) Then GoTo NextReservation
        If gBusy.Exists(currentKey) Then GoTo NextReservation

        record = gRecords(currentKey)
        statusText = UCase$(Trim$(CStr(record(AF_STATUS))))
        If statusText = "ASSIGNING" Then
            If Not IsDate(record(AF_LAST_VERIFIED)) Or _
               DateDiff("s", CDate(record(AF_LAST_VERIFIED)), nowEgypt) >= 45 Then
                retriesProcessed = retriesProcessed + 1
                Set itemObject = Nothing
                On Error Resume Next
                Set itemObject = Application.Session.GetItemFromID( _
                    CStr(record(AF_ENTRY_ID)), CStr(record(AF_STORE_ID)))
                On Error GoTo ReservationFailed
                If Not itemObject Is Nothing Then
                    If TypeOf itemObject Is Outlook.MailItem Then
                        Set mail = itemObject
                        Set sourceFolder = mail.Parent
                        gBusy.Add currentKey, True
                        ProcessTrackedMail mail, sourceFolder, currentKey, record
                        If gBusy.Exists(currentKey) Then gBusy.Remove currentKey
                    End If
                End If
            End If
        End If
NextReservation:
        scannedCount = scannedCount + 1
        gReservationCursorKey = currentKey
    Next scanOffset

    If scannedCount > 0 Then
        Debug.Print "[VWG PERFORMANCE] Assignment reservation slice scanned=" & CStr(scannedCount) & _
                    " | retries=" & CStr(retriesProcessed) & _
                    " | scan cap=" & CStr(MAX_RESERVATION_RECORDS_SCANNED_PER_CYCLE) & _
                    " | retry cap=" & CStr(MAX_RESERVATION_RETRIES_PER_CYCLE)
    End If
    Exit Sub

ReservationFailed:
    failureNumber = Err.Number
    failureDescription = Err.Description
    On Error Resume Next
    If Len(currentKey) > 0 Then
        If gBusy.Exists(currentKey) Then gBusy.Remove currentKey
    End If
    On Error GoTo 0
    Debug.Print "[VWG PERFORMANCE] Assignment reservation retry deferred | Key=" & currentKey & _
                " | " & CStr(failureNumber) & " - " & failureDescription
End Sub

'===========================================================
' QUEUE PROCESSING
'===========================================================

Private Function QueueKeyComesBefore(ByVal keyA As String, ByVal keyB As String) As Boolean
    Dim recordA As Variant
    Dim recordB As Variant
    Dim dateA As Date
    Dim dateB As Date

    recordA = gRecords(keyA)
    recordB = gRecords(keyB)
    If IsDate(recordA(AF_RECEIVED_UK)) Then
        dateA = CDate(recordA(AF_RECEIVED_UK))
    Else
        dateA = DateSerial(9999, 12, 31)
    End If
    If IsDate(recordB(AF_RECEIVED_UK)) Then
        dateB = CDate(recordB(AF_RECEIVED_UK))
    Else
        dateB = DateSerial(9999, 12, 31)
    End If

    If dateA < dateB Then
        QueueKeyComesBefore = True
    ElseIf dateA > dateB Then
        QueueKeyComesBefore = False
    Else
        ' Stable tie-break: identical ReceivedTime values are ordered by key.
        QueueKeyComesBefore = (StrComp(keyA, keyB, vbTextCompare) < 0)
    End If
End Function

Private Sub QuickSortQueuedKeys(ByRef keys() As String, ByVal first As Long, ByVal last As Long)
    Dim i As Long
    Dim j As Long
    Dim pivot As String
    Dim swapText As String

    i = first
    j = last
    pivot = keys((first + last) \ 2)
    Do While i <= j
        Do While QueueKeyComesBefore(keys(i), pivot)
            i = i + 1
        Loop
        Do While QueueKeyComesBefore(pivot, keys(j))
            j = j - 1
        Loop
        If i <= j Then
            swapText = keys(i)
            keys(i) = keys(j)
            keys(j) = swapText
            i = i + 1
            j = j - 1
        End If
    Loop
    If first < j Then QuickSortQueuedKeys keys, first, j
    If i < last Then QuickSortQueuedKeys keys, i, last
End Sub

Private Sub ProcessQueuedAssignments()
    Dim nowEgypt As Date
    Dim allKeys As Variant
    Dim queueKeys() As String
    Dim key As Variant
    Dim candidateKey As String
    Dim record As Variant
    Dim itemObject As Object
    Dim mail As Outlook.MailItem
    Dim sourceFolder As Outlook.Folder
    Dim assignee As String
    Dim rosterRow As Long
    Dim reason As String
    Dim matches As Variant
    Dim matchCount As Long
    Dim queuedCount As Long
    Dim candidateCount As Long
    Dim assignedCount As Long
    Dim i As Long
    Dim itemBusyOwned As Boolean
    Dim categoryApplied As Boolean
    Dim failureNumber As Long
    Dim failureDescription As String

    If Not gStarted Or gRecords Is Nothing Then Exit Sub
    If gRecords.Count = 0 Then Exit Sub
    If VWGX_AssignmentFallbackPending() Then
        Debug.Print "[VWG QUEUE] Assignment retry queue is pending; queue allocation paused."
        Exit Sub
    End If
    If Not GetEgyptNow(nowEgypt) Then Exit Sub
    EnsureFairnessPreviewDate nowEgypt

    On Error GoTo QueueFailed
    allKeys = gRecords.Keys
    ReDim queueKeys(0 To gRecords.Count - 1)
    For Each key In allKeys
        record = gRecords(CStr(key))
        If UCase$(Trim$(CStr(record(AF_STATUS)))) = "QUEUED" Then
            If Not (IsDryRun() And gPreviewSeen.Exists(CStr(key))) Then
                queueKeys(queuedCount) = CStr(key)
                queuedCount = queuedCount + 1
            End If
        End If
    Next key

    If queuedCount = 0 Then Exit Sub
    If Not PrepareFairnessBatch(nowEgypt) Then
        Debug.Print "[VWG QUEUE] Fairness snapshot unavailable; queued emails remain queued."
        Exit Sub
    End If
    ReDim Preserve queueKeys(0 To queuedCount - 1)
    If queuedCount > 1 Then QuickSortQueuedKeys queueKeys, LBound(queueKeys), UBound(queueKeys)

    ' Bound per-callback work to keep Outlook responsive. Remaining work is
    ' continued on the next timer tick in the same global oldest-first order.
    For i = LBound(queueKeys) To UBound(queueKeys)
        If VWGX_AssignmentFallbackPending() Then
            Debug.Print "[VWG QUEUE] Candidate processing paused because an assignment retry is pending."
            Exit For
        End If
        If candidateCount >= MAX_QUEUE_CANDIDATES_PER_CYCLE Then Exit For
        If assignedCount >= MAX_ASSIGNMENTS_PER_CYCLE Then Exit For
        candidateCount = candidateCount + 1
        candidateKey = queueKeys(i)
        If Not gRecords.Exists(candidateKey) Then GoTo ContinueQueue
        record = gRecords(candidateKey)
        If UCase$(Trim$(CStr(record(AF_STATUS)))) <> "QUEUED" Then GoTo ContinueQueue

        Set itemObject = Nothing
        On Error Resume Next
        Set itemObject = Application.Session.GetItemFromID( _
            CStr(record(AF_ENTRY_ID)), CStr(record(AF_STORE_ID)))
        On Error GoTo QueueFailed

        If itemObject Is Nothing Then
            record(AF_STATUS) = "Needs Review"
            record(AF_UPDATED_AT) = nowEgypt
            If SaveRecord(record) Then gRecords(candidateKey) = record
            WriteActivity "QUEUED_ITEM_NOT_FOUND", record, _
                "Queued email could not be retrieved. Manual review required."
            GoTo ContinueQueue
        End If
        If Not TypeOf itemObject Is Outlook.MailItem Then
            record(AF_STATUS) = "Needs Review"
            If SaveRecord(record) Then gRecords(candidateKey) = record
            WriteActivity "QUEUED_ITEM_NOT_MAIL", record, _
                "The queued Outlook item is not a MailItem. Manual review required."
            GoTo ContinueQueue
        End If

        Set mail = itemObject
        Set sourceFolder = mail.Parent
        matchCount = GetCategoryMatches(mail.Categories, matches)
        If matchCount = 1 Then
            RegisterManualOwner mail, sourceFolder, candidateKey, CStr(matches(0))
            GoTo ContinueQueue
        ElseIf matchCount > 1 Then
            MarkAmbiguous record, mail, matches, "Multiple categories appeared while queued."
            GoTo ContinueQueue
        End If

        If Not ChooseAssignee(nowEgypt, CStr(record(AF_FOLDER)), assignee, rosterRow, reason) Then Exit For

        If IsDryRun() Then
            If Not gPreviewSeen.Exists(candidateKey) Then
                gPreviewSeen.Add candidateKey, assignee
                WriteActivity "DRY_RUN_WOULD_ASSIGN_QUEUED", record, _
                    "Queued email would be assigned to " & assignee & ". Source-Inbox historical allocations before pick=" & _
                    CStr(gLastChosenFolderAllocationCount) & ". " & reason
                Debug.Print "[VWG DRY RUN] Queued item would go to " & assignee & " | " & record(AF_FOLDER)
                IncrementPreviewFolderAllocationCache CStr(record(AF_FOLDER)), assignee
                IncrementDictionaryCount gPreviewAllocationCounts, assignee
                IncrementDictionaryCount gPreviewWorkloadCounts, assignee
                IncrementDictionaryCount gBatchDailyAllocationCounts, assignee
                IncrementDictionaryCount gBatchDailyWorkloadCounts, assignee
                AdvanceTieBreak rosterRow, False
                assignedCount = assignedCount + 1
            End If
            GoTo ContinueQueue
        End If

        record(AF_ORIGINAL) = assignee
        record(AF_OWNER) = assignee
        record(AF_FLAG) = "TRUE"
        record(AF_STATUS) = "Assigning"
        record(AF_ALLOCATED_AT_EGYPT) = nowEgypt
        record(AF_ROSTER_PERIOD) = gRosterPeriodStart
        record(AF_LAST_VERIFIED) = nowEgypt
        record(AF_PENDING_SINCE) = Empty

        If Not SaveRecord(record) Then
            Debug.Print "[VWG QUEUE] Reservation write failed; item remains queued and queue processing is paused. Key=" & candidateKey
            Exit For
        End If
        gRecords(candidateKey) = record
        IncrementFolderAllocationCache CStr(record(AF_FOLDER)), assignee
        IncrementDictionaryCount gBatchDailyAllocationCounts, assignee
        IncrementDictionaryCount gBatchDailyWorkloadCounts, assignee

        If Not gBusy.Exists(candidateKey) Then
            gBusy.Add candidateKey, True
            itemBusyOwned = True
        End If
        categoryApplied = ApplyCategoryToMail(mail, assignee)
        If itemBusyOwned Then
            If gBusy.Exists(candidateKey) Then gBusy.Remove candidateKey
            itemBusyOwned = False
        End If

        If categoryApplied Then
            record(AF_STATUS) = "Assigned"
            record(AF_CATEGORY_STATE) = mail.Categories
            record(AF_UPDATED_AT) = nowEgypt
            If Not SaveRecord(record) Then
                Debug.Print "[VWG QUEUE] Outlook category was applied but final tracker commit failed. Reservation retained for reconciliation. Key=" & candidateKey
                Exit For
            End If
            gRecords(candidateKey) = record
            If Not WriteActivity("AUTO_ASSIGNMENT", record, _
                "Queued email assigned using source-Inbox-first fairness. Source-Inbox historical auto allocations before pick=" & _
                    CStr(gLastChosenFolderAllocationCount) & "; all-Inbox auto allocations today before pick=" & _
                    CStr(gLastChosenDailyCount) & "; active workload before pick=" & _
                CStr(gLastChosenWorkload) & "; eligible pool=" & CStr(gLastEligiblePoolCount) & _
                ". Subject=" & mail.Subject) Then
                AppendRuntimeEvent "AUTO_ASSIGNMENT_LOG_FAILED", sourceFolder.FolderPath, _
                    "Queued email remains Assigned; the Activity Log event could not be written or queued. Key=" & candidateKey
                Debug.Print "[VWG QUEUE] AUTO_ASSIGNMENT_LOG_FAILED; assignment remains Assigned. Key=" & candidateKey
            End If
            AdvanceTieBreak rosterRow, True
            assignedCount = assignedCount + 1
        Else
            record(AF_ORIGINAL) = assignee
            record(AF_OWNER) = assignee
            record(AF_FLAG) = "TRUE"
            record(AF_STATUS) = "Assigning"
            record(AF_PENDING_SINCE) = Empty
            record(AF_CATEGORY_STATE) = "Category update not yet confirmed; awaiting same-owner reconciliation. Intended assignee=" & assignee
            record(AF_UPDATED_AT) = nowEgypt
            record(AF_LAST_VERIFIED) = nowEgypt
            If Not SaveRecord(record) Then
                Debug.Print "[VWG QUEUE] Could not save same-owner reconciliation state. Queue paused. Key=" & candidateKey
                Exit For
            End If
            gRecords(candidateKey) = record
            If Not WriteActivity("AUTO_ASSIGNMENT_PENDING_VERIFICATION", record, _
                "Queued email category save was not yet confirmed. Intended owner reservation retained; reconciliation will retry the same owner and will not allocate a different colleague.") Then Exit For
            assignedCount = assignedCount + 1
        End If

ContinueQueue:
    Next i

    If candidateCount > 0 Then
        Debug.Print "[VWG PERFORMANCE] Queue cycle complete | candidates=" & CStr(candidateCount) & _
                    " | assigned/previews=" & CStr(assignedCount) & _
                    " | total queued snapshot=" & CStr(queuedCount) & _
                    " | processed in one sorted pass."
    End If

QueueCleanup:
    PersistTieBreakIfNeeded
    gFairnessBatchActive = False
    Set gBatchDailyAllocationCounts = Nothing
    Set gBatchDailyWorkloadCounts = Nothing
    Exit Sub

QueueFailed:
    failureNumber = Err.Number
    failureDescription = Err.Description
    On Error Resume Next
    If itemBusyOwned Then
        If gBusy.Exists(candidateKey) Then gBusy.Remove candidateKey
    End If
    On Error GoTo 0
    Debug.Print "[VWG QUEUE] Queue worker failed | " & CStr(failureNumber) & " - " & failureDescription
    Resume QueueCleanup
End Sub

'===========================================================
' FULL RECOVERY RECONCILIATION FOR TRACKED EMAILS
'===========================================================

Private Sub RefreshCurrentRecordStates()
    Const MAX_RECORDS_PER_REFRESH_CYCLE As Long = 5
    Dim keys As Variant
    Dim key As Variant
    Dim record As Variant
    Dim itemObject As Object
    Dim mail As Outlook.MailItem
    Dim sourceFolder As Outlook.Folder
    Dim total As Long
    Dim startIndex As Long
    Dim i As Long
    Dim idx As Long
    Dim processed As Long
    Dim currentKey As String

    If gRecords Is Nothing Then Exit Sub
    total = gRecords.Count
    If total = 0 Then Exit Sub

    keys = gRecords.Keys
    startIndex = gRecordRefreshCursorIndex
    If startIndex < 0 Or startIndex >= total Then startIndex = 0

    For i = 0 To total - 1
        If processed >= MAX_RECORDS_PER_REFRESH_CYCLE Then Exit For
        If VWGX_AssignmentFallbackPending() Then Exit For
        idx = LBound(keys) + ((startIndex + i) Mod total)
        currentKey = CStr(keys(idx))
        If Not gRecords.Exists(currentKey) Then GoTo NextRefreshRecord
        record = gRecords(currentKey)
        If IsTrackableStatus(CStr(record(AF_STATUS))) Then
            If Not gBusy.Exists(currentKey) Then
                Set itemObject = Nothing
                On Error Resume Next
                Set itemObject = Application.Session.GetItemFromID( _
                    CStr(record(AF_ENTRY_ID)), CStr(record(AF_STORE_ID)))
                On Error GoTo 0
                If Not itemObject Is Nothing Then
                    If TypeOf itemObject Is Outlook.MailItem Then
                        Set mail = itemObject
                        Set sourceFolder = mail.Parent
                        gBusy.Add currentKey, True
                        On Error Resume Next
                        ProcessTrackedMail mail, sourceFolder, currentKey, record
                        If Err.Number <> 0 Then
                            Debug.Print "[VWG PERFORMANCE] Record reconciliation deferred | Key=" & currentKey & _
                                        " | " & Err.Number & " - " & Err.Description
                            Err.Clear
                        End If
                        If gBusy.Exists(currentKey) Then gBusy.Remove currentKey
                        On Error GoTo 0
                    End If
                End If
            End If
        End If
NextRefreshRecord:
        processed = processed + 1
    Next i

    gRecordRefreshCursorIndex = (startIndex + processed) Mod total
    Debug.Print "[VWG PERFORMANCE] Reconciliation slice processed=" & CStr(processed) & _
                " | tracked records=" & CStr(total)
End Sub

Private Function IsTrackableStatus(ByVal statusText As String) As Boolean
    Select Case UCase$(Trim$(statusText))
        Case "ASSIGNING", "ASSIGNED", "HANDED OVER", "HANDOVER PENDING", "RECOVERED", "RECOVERED - REVIEW", "RECOVERY FAILED", "NEEDS REVIEW", "DRY RUN - RECOVERY PENDING"
            IsTrackableStatus = True
        Case Else
            IsTrackableStatus = False
    End Select
End Function

'===========================================================
' RECORD BUILD / PERSISTENCE
'===========================================================

Private Function NewAssignmentRecord(ByVal mail As Outlook.MailItem, _
                                     ByVal sourceFolder As Outlook.Folder, _
                                     ByVal itemKey As String) As Variant
    Dim rec(0 To 20) As Variant
    Dim receivedUK As Date
    Dim nowEgypt As Date
    Dim emailAddress As String

    ' Populate stable identifiers first so even a timestamp failure produces
    ' a complete Needs Review record rather than an empty/broken row.
    rec(AF_KEY) = itemKey
    rec(AF_FOLDER) = sourceFolder.FolderPath
    rec(AF_SUBJECT) = mail.Subject
    rec(AF_ORIGINAL) = ""
    rec(AF_OWNER) = ""
    rec(AF_FLAG) = ""
    rec(AF_INCIDENTS) = CLng(0)
    rec(AF_LAST_INCIDENT) = Empty
    rec(AF_STATUS) = "Unassigned"
    rec(AF_HANDOVER_NOTES) = ""
    rec(AF_LOGIN_EMAIL) = ""
    rec(AF_UPDATED_AT) = Empty
    rec(AF_ALLOCATED_AT_EGYPT) = Empty
    If gRosterPeriodSet Then rec(AF_ROSTER_PERIOD) = gRosterPeriodStart Else rec(AF_ROSTER_PERIOD) = Empty
    rec(AF_ENTRY_ID) = mail.EntryID
    rec(AF_STORE_ID) = sourceFolder.StoreID
    rec(AF_PENDING_SINCE) = Empty
    rec(AF_CATEGORY_STATE) = mail.Categories
    rec(AF_HANDOVER_COUNT) = CLng(0)

    If Not ConvertLocalDateToZone(mail.ReceivedTime, GetUKZoneId(), receivedUK) Then
        rec(AF_STATUS) = "Needs Review"
        NewAssignmentRecord = rec
        Exit Function
    End If
    If Not GetEgyptNow(nowEgypt) Then
        rec(AF_STATUS) = "Needs Review"
        NewAssignmentRecord = rec
        Exit Function
    End If

    emailAddress = ExtractLoginUserEmail(mail.Body)
    rec(AF_RECEIVED_UK) = receivedUK
    rec(AF_LAST_VERIFIED) = nowEgypt
    rec(AF_LOGIN_EMAIL) = emailAddress
    rec(AF_UPDATED_AT) = nowEgypt

    NewAssignmentRecord = rec
End Function

Private Function SaveRecord(ByVal record As Variant) As Boolean
    Dim nowEgypt As Date
    If GetEgyptNow(nowEgypt) Then
        record(AF_UPDATED_AT) = nowEgypt
    ElseIf Not IsDate(record(AF_UPDATED_AT)) Then
        ' Preserve an empty timestamp rather than writing the PC clock as Egypt time.
        record(AF_UPDATED_AT) = Empty
    End If
    SaveRecord = VWGX_UpsertAssignment(record)
End Function

Private Sub LoadRecordsFromWorkbook()
    Dim data As Variant
    Dim r As Long
    Dim keyText As String
    Dim rec(0 To 20) As Variant
    Dim col As Long

    data = VWGX_LoadAssignments()
    If Not IsArray(data) Then Exit Sub

    For r = LBound(data, 1) To UBound(data, 1)
        keyText = Trim$(CStr(data(r, 1)))
        If Len(keyText) > 0 Then
            For col = 0 To 20
                rec(col) = data(r, col + 1)
            Next col
            If IsNumeric(rec(AF_INCIDENTS)) Then rec(AF_INCIDENTS) = CLng(Val(CStr(rec(AF_INCIDENTS)))) Else rec(AF_INCIDENTS) = 0
            If IsNumeric(rec(AF_HANDOVER_COUNT)) Then rec(AF_HANDOVER_COUNT) = CLng(Val(CStr(rec(AF_HANDOVER_COUNT)))) Else rec(AF_HANDOVER_COUNT) = 0
            If Len(Trim$(CStr(rec(AF_STATUS)))) = 0 Then rec(AF_STATUS) = "Needs Review"
            gRecords(keyText) = rec
        End If
    Next r
End Sub

Private Function GetMessageKey(ByVal mail As Outlook.MailItem, _
                               ByVal sourceFolder As Outlook.Folder) As String
    Dim entryId As String
    On Error GoTo NoKey
    entryId = Trim$(mail.EntryID)
    If Len(entryId) = 0 Then Exit Function
    GetMessageKey = sourceFolder.StoreID & "|" & entryId
    Exit Function
NoKey:
    GetMessageKey = vbNullString
End Function

Private Function GetMailFromRecord(ByVal record As Variant) As Outlook.MailItem
    Dim obj As Object
    On Error GoTo NotFound
    Set obj = Application.Session.GetItemFromID( _
        CStr(record(AF_ENTRY_ID)), CStr(record(AF_STORE_ID)))
    If Not obj Is Nothing Then
        If TypeOf obj Is Outlook.MailItem Then Set GetMailFromRecord = obj
    End If
    Exit Function
NotFound:
    Set GetMailFromRecord = Nothing
End Function

'===========================================================
' CATEGORY HELPERS / ROSTER MATCHING
'===========================================================

Private Function GetCategoryMatches(ByVal categoriesText As String, _
                                    ByRef matchedCategories As Variant) As Long
    Dim result() As String
    Dim seen As Object
    Dim i As Long
    Dim categoryName As String
    Dim rosterCategory As String
    Dim count As Long

    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare

    If IsArray(gRoster) Then
        For i = LBound(gRoster, 1) To UBound(gRoster, 1)
            rosterCategory = Trim$(CStr(gRoster(i, 2)))
            If Len(rosterCategory) > 0 Then
                If HasCategoryString(categoriesText, rosterCategory) Then
                    If Not seen.Exists(rosterCategory) Then seen.Add rosterCategory, True
                End If
            End If
        Next i
    End If

    count = seen.Count
    If count > 0 Then
        ReDim result(0 To count - 1)
        i = 0
        Dim key As Variant
        For Each key In seen.Keys
            result(i) = CStr(key)
            i = i + 1
        Next key
        matchedCategories = result
    Else
        matchedCategories = Array()
    End If
    GetCategoryMatches = count
End Function

Private Function HasCategory(ByVal mail As Outlook.MailItem, ByVal categoryName As String) As Boolean
    HasCategory = HasCategoryString(mail.Categories, categoryName)
End Function

Private Function HasCategoryString(ByVal categoryText As String, ByVal categoryName As String) As Boolean
    Dim parts As Variant
    Dim part As Variant
    Dim separator As String

    If Len(Trim$(categoryText)) = 0 Or Len(Trim$(categoryName)) = 0 Then Exit Function
    separator = GetCategorySeparator()
    parts = Split(categoryText, separator)
    For Each part In parts
        If StrComp(Trim$(CStr(part)), Trim$(categoryName), vbTextCompare) = 0 Then
            HasCategoryString = True
            Exit Function
        End If
    Next part
End Function

Private Function AddCategory(ByVal mail As Outlook.MailItem, ByVal categoryName As String) As Boolean
    Dim separator As String
    On Error GoTo AddFailed

    If HasCategory(mail, categoryName) Then
        AddCategory = True
        Exit Function
    End If

    separator = GetCategorySeparator()
    If Len(Trim$(mail.Categories)) = 0 Then
        mail.Categories = categoryName
    Else
        mail.Categories = mail.Categories & separator & " " & categoryName
    End If
    AddCategory = True
    Exit Function

AddFailed:
    AddCategory = False
End Function

Private Function ApplyCategoryToMail(ByRef mail As Outlook.MailItem, _
                                     ByVal categoryName As String) As Boolean
    Dim entryId As String
    Dim storeId As String
    Dim parentFolder As Outlook.Folder
    Dim candidateObject As Object
    Dim candidateMail As Outlook.MailItem
    Dim verifyObject As Object
    Dim verifyMail As Outlook.MailItem
    Dim attempt As Long
    Dim saveErr As Long
    Dim saveDescription As String
    Dim getErr As Long
    Dim getDescription As String
    Dim stageText As String

    On Error GoTo ApplyFailed

    stageText = "read message EntryID and StoreID"
    If mail Is Nothing Then Err.Raise vbObjectError + 2301, "ApplyCategoryToMail", "MailItem reference is Nothing."
    entryId = CStr(mail.EntryID)
    Set parentFolder = mail.Parent
    storeId = CStr(parentFolder.StoreID)
    If Len(entryId) = 0 Or Len(storeId) = 0 Then
        Err.Raise vbObjectError + 2302, "ApplyCategoryToMail", _
                  "Message EntryID or StoreID is blank; refusing to save a potentially unsaved item."
    End If

    ' Never save the caller's cached MailItem reference. Reacquire the item
    ' from the owning store, merge the requested category into its CURRENT
    ' category list, save, then reacquire again to verify the server/store state.
    For attempt = 1 To 2
        stageText = "reacquire MailItem, attempt " & CStr(attempt)
        Set candidateObject = Nothing
        On Error Resume Next
        Set candidateObject = Application.Session.GetItemFromID(entryId, storeId)
        getErr = Err.Number
        getDescription = Err.Description
        Err.Clear
        On Error GoTo ApplyFailed

        If candidateObject Is Nothing Then
            Debug.Print "[VWG] Category update could not reacquire message at [" & stageText & "]: " & _
                        getErr & " - " & getDescription
            Exit Function
        End If
        If Not TypeOf candidateObject Is Outlook.MailItem Then
            Debug.Print "[VWG] Category update stopped: reacquired item is not a MailItem."
            Exit Function
        End If

        Set candidateMail = candidateObject
        If HasCategory(candidateMail, categoryName) Then
            Set mail = candidateMail
            ApplyCategoryToMail = True
            Exit Function
        End If

        stageText = "append category to fresh MailItem, attempt " & CStr(attempt)
        If Not AddCategory(candidateMail, categoryName) Then
            Debug.Print "[VWG] Category update failed at [" & stageText & "]: AddCategory returned False."
            Exit Function
        End If

        stageText = "save fresh MailItem, attempt " & CStr(attempt)
        saveErr = 0
        saveDescription = vbNullString
        On Error Resume Next
        candidateMail.Save
        saveErr = Err.Number
        saveDescription = Err.Description
        Err.Clear
        On Error GoTo ApplyFailed

        ' Do not report a Save conflict as a failed assignment until the fresh
        ' read below has determined whether Outlook committed the category.

        ' A failed Save can still have committed the category. Always inspect
        ' a separately reacquired item before deciding to retry or fail.
        stageText = "reacquire MailItem to verify, attempt " & CStr(attempt)
        Set verifyObject = Nothing
        On Error Resume Next
        Set verifyObject = Application.Session.GetItemFromID(entryId, storeId)
        getErr = Err.Number
        getDescription = Err.Description
        Err.Clear
        On Error GoTo ApplyFailed

        If Not verifyObject Is Nothing Then
            If TypeOf verifyObject Is Outlook.MailItem Then
                Set verifyMail = verifyObject
                If HasCategory(verifyMail, categoryName) Then
                    Set mail = verifyMail
                    If saveErr = ERR_MAPI_OBJECT_CHANGED Then
                        Debug.Print "[VWG] Outlook reported the message changed during Save, but a fresh read confirms category '" & _
                                    categoryName & "' is present. Treated as a successful commit; no duplicate retry."
                    ElseIf saveErr <> 0 Then
                        Debug.Print "[VWG] Outlook Save returned " & saveErr & ", but a fresh read confirms category '" & _
                                    categoryName & "' is present. Treated as successful."
                    End If
                    ApplyCategoryToMail = True
                    Exit Function
                End If
            End If
        ElseIf getErr <> 0 Then
            Debug.Print "[VWG] Category verification could not reacquire message: " & _
                        getErr & " - " & getDescription
        End If

        If saveErr <> 0 Then
            Debug.Print "[VWG] Outlook Save result was not committed/visible on fresh verification at [" & _
                        stageText & "]: " & saveErr & " - " & saveDescription
        End If

        ' Only the MAPI object-changed conflict is automatically retried after
        ' a Save error. Other Save errors are not assumed to be transient.
        If saveErr <> 0 And saveErr <> ERR_MAPI_OBJECT_CHANGED Then Exit For
        If attempt = 1 Then
            If saveErr = ERR_MAPI_OBJECT_CHANGED Then
                Debug.Print "[VWG] MAPI object-changed conflict detected. Retrying once with a freshly reacquired MailItem."
            Else
                Debug.Print "[VWG] Category not visible after Save; performing one fresh-item verification retry."
            End If
        End If
    Next attempt

    Debug.Print "[VWG] Category update was not confirmed after controlled fresh-item retry. Category=" & categoryName
    ApplyCategoryToMail = False
    Exit Function

ApplyFailed:
    Debug.Print "[VWG] Category application failed at [" & stageText & "]: " & _
                Err.Number & " - " & Err.Description
    ApplyCategoryToMail = False
End Function

Private Function IsRestorationMarkerName(ByVal categoryName As String) As Boolean
    Dim value As String
    value = Trim$(categoryName)
    If StrComp(value, RESTORE_MARKER_LEGACY_NAME, vbTextCompare) = 0 Then
        IsRestorationMarkerName = True
    ElseIf StrComp(Left$(value, Len(RESTORE_MARKER_PREFIX)), _
                   RESTORE_MARKER_PREFIX, vbTextCompare) = 0 Then
        IsRestorationMarkerName = True
    End If
End Function

Private Function HasRestorationMarker(ByVal categoriesText As String) As Boolean
    Dim parts As Variant
    Dim part As Variant
    Dim separator As String

    If Len(Trim$(categoriesText)) = 0 Then Exit Function
    separator = GetCategorySeparator()
    parts = Split(categoriesText, separator)
    For Each part In parts
        If IsRestorationMarkerName(CStr(part)) Then
            HasRestorationMarker = True
            Exit Function
        End If
    Next part
End Function

Private Function RemoveRestorationMarkerCategories(ByRef mail As Outlook.MailItem, _
                                                   ByVal requiredOwner As String) As Boolean
    Dim entryId As String
    Dim storeId As String
    Dim parentFolder As Outlook.Folder
    Dim candidateObject As Object
    Dim candidateMail As Outlook.MailItem
    Dim verifyObject As Object
    Dim verifyMail As Outlook.MailItem
    Dim parts As Variant
    Dim part As Variant
    Dim separator As String
    Dim rebuilt As String
    Dim foundMarker As Boolean
    Dim attempt As Long
    Dim saveErr As Long
    Dim saveDescription As String
    Dim stageText As String

    On Error GoTo RemoveFailed
    If mail Is Nothing Then Exit Function

    stageText = "read message EntryID and StoreID for marker cleanup"
    entryId = CStr(mail.EntryID)
    Set parentFolder = mail.Parent
    storeId = CStr(parentFolder.StoreID)
    If Len(entryId) = 0 Or Len(storeId) = 0 Then Exit Function

    separator = GetCategorySeparator()

    For attempt = 1 To 2
        stageText = "reacquire message for restoration-marker cleanup, attempt " & CStr(attempt)
        Set candidateObject = Nothing
        On Error Resume Next
        Set candidateObject = Application.Session.GetItemFromID(entryId, storeId)
        On Error GoTo RemoveFailed

        If candidateObject Is Nothing Then Exit Function
        If Not TypeOf candidateObject Is Outlook.MailItem Then Exit Function
        Set candidateMail = candidateObject

        If Not HasCategory(candidateMail, requiredOwner) Then
            Debug.Print "[VWG] Marker cleanup refused: owner category is not present on fresh message. Key=" & entryId
            Exit Function
        End If

        parts = Split(CStr(candidateMail.Categories), separator)
        rebuilt = vbNullString
        foundMarker = False

        For Each part In parts
            If IsRestorationMarkerName(CStr(part)) Then
                foundMarker = True
            ElseIf Len(Trim$(CStr(part))) > 0 Then
                If Len(rebuilt) > 0 Then rebuilt = rebuilt & separator & " "
                rebuilt = rebuilt & Trim$(CStr(part))
            End If
        Next part

        If Not foundMarker Then
            Set mail = candidateMail
            RemoveRestorationMarkerCategories = True
            Exit Function
        End If

        stageText = "remove legacy Assignment Restored category markers, attempt " & CStr(attempt)
        candidateMail.Categories = rebuilt

        stageText = "save marker-free MailItem, attempt " & CStr(attempt)
        saveErr = 0
        saveDescription = vbNullString
        On Error Resume Next
        candidateMail.Save
        saveErr = Err.Number
        saveDescription = Err.Description
        Err.Clear
        On Error GoTo RemoveFailed

        stageText = "verify marker-free MailItem, attempt " & CStr(attempt)
        Set verifyObject = Nothing
        On Error Resume Next
        Set verifyObject = Application.Session.GetItemFromID(entryId, storeId)
        On Error GoTo RemoveFailed

        If Not verifyObject Is Nothing Then
            If TypeOf verifyObject Is Outlook.MailItem Then
                Set verifyMail = verifyObject
                If HasCategory(verifyMail, requiredOwner) And _
                   Not HasRestorationMarker(verifyMail.Categories) Then
                    Set mail = verifyMail
                    If saveErr = ERR_MAPI_OBJECT_CHANGED Then
                        Debug.Print "[VWG] Outlook reported a change conflict, but fresh verification confirms legacy restoration marker removal."
                    ElseIf saveErr <> 0 Then
                        Debug.Print "[VWG] Outlook Save returned " & saveErr & ", but fresh verification confirms marker removal."
                    End If
                    RemoveRestorationMarkerCategories = True
                    Exit Function
                End If
            End If
        End If

        If saveErr <> 0 And saveErr <> ERR_MAPI_OBJECT_CHANGED Then Exit For
        If attempt = 1 Then
            Debug.Print "[VWG] Restoration marker still present after Save; retrying with a fresh item."
        End If
    Next attempt

    Debug.Print "[VWG] Could not confirm removal of restoration-marker categories. EntryID=" & entryId
    RemoveRestorationMarkerCategories = False
    Exit Function

RemoveFailed:
    Debug.Print "[VWG] Restoration-marker cleanup failed at [" & stageText & "]: " & _
                Err.Number & " - " & Err.Description
    RemoveRestorationMarkerCategories = False
End Function

Private Function GetCategorySeparator() As String
    Dim shellObject As Object
    On Error GoTo UseComma
    Set shellObject = CreateObject("WScript.Shell")
    GetCategorySeparator = CStr(shellObject.RegRead( _
        "HKEY_CURRENT_USER\Control Panel\International\sList"))
    If Len(GetCategorySeparator) = 0 Then GoTo UseComma
    Exit Function
UseComma:
    GetCategorySeparator = ","
End Function

Private Function JoinVariant(ByVal values As Variant, ByVal separator As String) As String
    Dim i As Long
    Dim output As String
    On Error GoTo Done
    If Not IsArray(values) Then Exit Function
    For i = LBound(values) To UBound(values)
        If Len(output) > 0 Then output = output & separator
        output = output & CStr(values(i))
    Next i
Done:
    JoinVariant = output
End Function

'===========================================================
' MAIL / FOLDER HELPERS
'===========================================================

Private Function GetFolderByPath(ByVal mailboxDisplayName As String, _
                                 ByVal relativePath As String) As Outlook.Folder
    Dim currentFolder As Outlook.Folder
    Dim parts As Variant
    Dim i As Long
    On Error GoTo FolderFailed
    Set currentFolder = Application.Session.Folders.Item(mailboxDisplayName)
    parts = Split(relativePath, "\")
    For i = LBound(parts) To UBound(parts)
        Set currentFolder = currentFolder.Folders.Item(CStr(parts(i)))
    Next i
    Set GetFolderByPath = currentFolder
    Exit Function
FolderFailed:
    Set GetFolderByPath = Nothing
End Function

Private Function GetMonitoredFolders() As Collection
    Dim items As New Collection
    Dim mailboxNames As Variant
    Dim item As Variant
    Dim folder As Outlook.Folder
    mailboxNames = GetMailboxNames()
    For Each item In mailboxNames
        Set folder = GetFolderByPath(CStr(item), "Inbox")
        If Not folder Is Nothing Then items.Add folder
    Next item
    Set GetMonitoredFolders = items
End Function

Private Function FolderPathSafe(ByVal folder As Outlook.Folder) As String
    On Error Resume Next
    If folder Is Nothing Then
        FolderPathSafe = "(unknown folder)"
    Else
        FolderPathSafe = folder.FolderPath
    End If
    On Error GoTo 0
End Function

Private Function ExtractLoginUserEmail(ByVal bodyText As String) As String
    Dim regex As Object
    Dim matches As Object
    On Error GoTo NoMatch
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "LoginUserEmail\s*:\s*([^\r\n<]+)"
    regex.IgnoreCase = True
    regex.Global = False
    If regex.Test(bodyText) Then
        Set matches = regex.Execute(bodyText)
        ExtractLoginUserEmail = Trim$(CStr(matches(0).SubMatches(0)))
    End If
    Exit Function
NoMatch:
    ExtractLoginUserEmail = ""
End Function

'===========================================================
' CONFIGURATION / TIME ZONES
'===========================================================

Private Sub LoadWorkbookConfiguration()
    Dim value As Variant
    Dim periodValue As Variant

    gRoster = VWGX_ReadRoster()
    If Not IsArray(gRoster) Then Exit Sub

    periodValue = VWGX_GetSetting("Roster Period Start")
    gRosterPeriodSet = IsDate(periodValue)
    If gRosterPeriodSet Then gRosterPeriodStart = DateValue(CDate(periodValue))

    value = VWGX_GetSetting("Round Robin Tie-Break Index")
    gNextTieBreakIndex = CLng(Val(CStr(value)))
    If gNextTieBreakIndex < 1 Or gNextTieBreakIndex > UBound(gRoster, 1) Then gNextTieBreakIndex = 1

    gHandoverGraceSeconds = GetNumericSetting( _
        "Handover Grace (seconds)", DEFAULT_HANDOVER_GRACE_SECONDS)
    If gHandoverGraceSeconds < 5 Then gHandoverGraceSeconds = DEFAULT_HANDOVER_GRACE_SECONDS
End Sub

Private Function IsDryRun() As Boolean
    Dim modeText As String
    modeText = UCase$(Trim$(CStr(VWGX_GetSetting("Dispatcher Mode"))))
    IsDryRun = DRY_RUN_MODE Or (modeText <> "LIVE")
End Function

Private Function GetNumericSetting(ByVal settingName As String, ByVal defaultValue As Long) As Long
    Dim value As Variant
    value = VWGX_GetSetting(settingName)
    If IsNumeric(value) Then
        GetNumericSetting = CLng(value)
    Else
        GetNumericSetting = defaultValue
    End If
End Function

Private Function GetEgyptZoneId() As String
    Dim value As Variant
    value = VWGX_GetSetting("Windows Egypt Time Zone ID")
    If Len(Trim$(CStr(value))) = 0 Then GetEgyptZoneId = "Egypt Standard Time" Else GetEgyptZoneId = CStr(value)
End Function

Private Function GetUKZoneId() As String
    Dim value As Variant
    value = VWGX_GetSetting("Windows UK Time Zone ID")
    If Len(Trim$(CStr(value))) = 0 Then GetUKZoneId = "GMT Standard Time" Else GetUKZoneId = CStr(value)
End Function

Private Function GetEgyptNow(ByRef result As Date) As Boolean
    Dim utcTime As VWG_SYSTEMTIME
    Dim egyptTime As VWG_SYSTEMTIME
    Dim targetZone As VWG_DYNAMIC_TIME_ZONE_INFORMATION
    Dim ageSeconds As Long

    ' Cache briefly: the same Egypt timestamp is requested repeatedly while
    ' one message is processed. No external process is launched.
    If gHaveCachedEgyptTime Then
        ageSeconds = DateDiff("s", gCachedEgyptTimeAt, Now)
        If ageSeconds >= 0 And ageSeconds < 2 Then
            result = gCachedEgyptTime
            GetEgyptNow = True
            Exit Function
        End If
    End If

    GetSystemTime utcTime
    If Not FindWindowsTimeZone(GetEgyptZoneId(), targetZone) Then
        Debug.Print "[VWG TIME] Windows time zone not found: " & GetEgyptZoneId()
        Exit Function
    End If
    If SystemTimeToTzSpecificLocalTimeEx(targetZone, utcTime, egyptTime) = 0 Then
        Debug.Print "[VWG TIME] Windows API could not convert UTC to Egypt local time."
        Exit Function
    End If
    If Not VWGSystemTimeToDate(egyptTime, result) Then Exit Function

    gCachedEgyptTime = result
    gCachedEgyptTimeAt = Now
    gHaveCachedEgyptTime = True
    GetEgyptNow = True
End Function

Private Function ConvertLocalDateToZone(ByVal localValue As Date, _
                                        ByVal targetZoneId As String, _
                                        ByRef convertedValue As Date) As Boolean
    Dim localZone As VWG_DYNAMIC_TIME_ZONE_INFORMATION
    Dim targetZone As VWG_DYNAMIC_TIME_ZONE_INFORMATION
    Dim localSystemTime As VWG_SYSTEMTIME
    Dim utcSystemTime As VWG_SYSTEMTIME
    Dim targetSystemTime As VWG_SYSTEMTIME

    On Error GoTo ConversionFailed
    If Not VWGDateToSystemTime(localValue, localSystemTime) Then Exit Function

    If GetDynamicTimeZoneInformation(localZone) = VWG_TIME_ZONE_ID_INVALID Then
        Debug.Print "[VWG TIME] Could not read the PC's Windows time zone."
        Exit Function
    End If
    If Not FindWindowsTimeZone(targetZoneId, targetZone) Then
        Debug.Print "[VWG TIME] Windows time zone not found: " & targetZoneId
        Exit Function
    End If

    ' Outlook ReceivedTime is expressed in the PC's local wall-clock time.
    ' Convert it to UTC using the PC's current Windows dynamic rules, then
    ' convert UTC into the requested destination zone.
    If TzSpecificLocalTimeToSystemTimeEx(localZone, localSystemTime, utcSystemTime) = 0 Then
        Debug.Print "[VWG TIME] Could not convert PC local time to UTC."
        Exit Function
    End If
    If SystemTimeToTzSpecificLocalTimeEx(targetZone, utcSystemTime, targetSystemTime) = 0 Then
        Debug.Print "[VWG TIME] Could not convert UTC to " & targetZoneId & "."
        Exit Function
    End If
    ConvertLocalDateToZone = VWGSystemTimeToDate(targetSystemTime, convertedValue)
    Exit Function

ConversionFailed:
    Debug.Print "[VWG TIME] Native time-zone conversion error: " & Err.Number & " - " & Err.Description
    ConvertLocalDateToZone = False
End Function

Private Function FindWindowsTimeZone(ByVal zoneKey As String, _
                                      ByRef zoneInfo As VWG_DYNAMIC_TIME_ZONE_INFORMATION) As Boolean
    Dim candidate As VWG_DYNAMIC_TIME_ZONE_INFORMATION
    Dim index As Long
    Dim apiResult As Long

    On Error GoTo FindFailed
    index = 0
    Do
        apiResult = EnumDynamicTimeZoneInformation(index, candidate)
        If apiResult = VWG_ERROR_NO_MORE_ITEMS Then Exit Do
        If apiResult <> 0 Then Exit Do
        If StrComp(VWGZoneKeyName(candidate), zoneKey, vbTextCompare) = 0 Then
            LSet zoneInfo = candidate
            FindWindowsTimeZone = True
            Exit Function
        End If
        index = index + 1
        If index > 512 Then Exit Do
    Loop
    Exit Function
FindFailed:
    Debug.Print "[VWG TIME] Time-zone enumeration failed: " & Err.Number & " - " & Err.Description
End Function

Private Function VWGZoneKeyName(ByRef zoneInfo As VWG_DYNAMIC_TIME_ZONE_INFORMATION) As String
    Dim i As Long
    Dim result As String
    For i = LBound(zoneInfo.TimeZoneKeyName) To UBound(zoneInfo.TimeZoneKeyName)
        If zoneInfo.TimeZoneKeyName(i) = 0 Then Exit For
        result = result & ChrW$(zoneInfo.TimeZoneKeyName(i))
    Next i
    VWGZoneKeyName = result
End Function

Private Function VWGDateToSystemTime(ByVal value As Date, _
                                     ByRef result As VWG_SYSTEMTIME) As Boolean
    On Error GoTo ConvertFailed
    result.wYear = Year(value)
    result.wMonth = Month(value)
    result.wDay = Day(value)
    result.wDayOfWeek = Weekday(value, vbSunday) - 1
    result.wHour = Hour(value)
    result.wMinute = Minute(value)
    result.wSecond = Second(value)
    result.wMilliseconds = 0
    VWGDateToSystemTime = True
    Exit Function
ConvertFailed:
    VWGDateToSystemTime = False
End Function

Private Function VWGSystemTimeToDate(ByRef value As VWG_SYSTEMTIME, _
                                      ByRef result As Date) As Boolean
    On Error GoTo ConvertFailed
    result = DateSerial(CLng(value.wYear), CLng(value.wMonth), CLng(value.wDay)) + _
             TimeSerial(CLng(value.wHour), CLng(value.wMinute), CLng(value.wSecond)) + _
             (CDbl(value.wMilliseconds) / 86400000#)
    VWGSystemTimeToDate = True
    Exit Function
ConvertFailed:
    VWGSystemTimeToDate = False
End Function

'===========================================================
' EXCEL TRACKER WRITE HELPERS
'===========================================================

Private Function WriteActivity(ByVal eventType As String, _
                               ByVal record As Variant, _
                               ByVal details As String) As Boolean
    Dim eventTimeEgypt As Date
    Dim activity(0 To 8) As Variant
    Dim timeOk As Boolean

    timeOk = GetEgyptNow(eventTimeEgypt)
    If Not timeOk Then
        eventTimeEgypt = Now
        details = "[Egypt time conversion unavailable; event time is PC local time] " & details
    End If

    activity(0) = eventTimeEgypt
    activity(1) = eventType
    activity(2) = CStr(record(AF_KEY))
    activity(3) = CStr(record(AF_FOLDER))
    activity(4) = CStr(record(AF_ORIGINAL))
    activity(5) = CStr(record(AF_OWNER))
    activity(6) = CStr(record(AF_FLAG))
    activity(7) = CLng(Val(CStr(record(AF_INCIDENTS))))
    activity(8) = details

    WriteActivity = VWGX_AppendActivity(activity)
    If Not WriteActivity Then Debug.Print "[VWG TRACKER] Failed to append event: " & eventType
End Function

Private Sub AppendRuntimeEvent(ByVal eventType As String, _
                               ByVal target As String, _
                               ByVal details As String)
    Dim nowEgypt As Date
    Dim activity(0 To 8) As Variant
    If GetEgyptNow(nowEgypt) Then
        activity(0) = nowEgypt
    Else
        activity(0) = Empty
    End If
    activity(1) = eventType
    activity(2) = ""
    activity(3) = target
    activity(4) = ""
    activity(5) = ""
    activity(6) = ""
    activity(7) = 0
    activity(8) = details
    Call VWGX_AppendActivity(activity)
    Debug.Print "[VWG] " & eventType & " | " & target & " | " & details
End Sub


'===========================================================
' DASHBOARD / SUMMARY MACROS
'===========================================================

Public Sub VWGRefreshDashboard()
    Dim summaryEgypt As Date
    If Not VWGX_EnsureConnected(True) Then Exit Sub
    If GetEgyptNow(summaryEgypt) Then
        VWGX_RefreshSummary summaryEgypt
    Else
        Debug.Print "[VWG TRACKER] Summary refresh skipped: Egypt-local date unavailable; no PC-date workload reset was applied."
    End If
    MsgBox "Dashboard counts refreshed.", vbInformation, "VWG Dispatcher"
End Sub

Public Sub VWGShowCurrentStatus()
    Dim text As String
    text = "Dispatcher running: " & IIf(gStarted, "YES", "NO") & vbCrLf & _
           "Mode: " & IIf(IsDryRun(), "DRY RUN", "LIVE") & vbCrLf & _
           "Monitored Inboxes: " & IIf(gMonitors Is Nothing, "0", CStr(gMonitors.Count)) & vbCrLf & _
           "Workbook connected: " & IIf(VWGX_EnsureConnected(False), "YES", "NO") & vbCrLf & _
           "Roster period set: " & IIf(gRosterPeriodSet, Format$(gRosterPeriodStart, "dd-mmm-yyyy"), "NO - historical metadata only")
    MsgBox text, vbInformation, "VWG Dispatcher Status"
End Sub

'===========================================================
' AUDIT: DOES NOT IDENTIFY WHO CHANGED A CATEGORY
'===========================================================

Private Function SafeText(ByVal value As String) As String
    value = Replace(value, vbTab, " ")
    value = Replace(value, vbCr, " ")
    value = Replace(value, vbLf, " ")
    SafeText = Trim$(value)
End Function
