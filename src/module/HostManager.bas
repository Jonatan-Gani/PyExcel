Attribute VB_Name = "HostManager"
Option Explicit
#Const DEV = True

'======================
' Globals
'======================
Public DebugMode As Boolean
Public Initialized As Boolean

Public Ribbon As IRibbonUI            ' assigned by your future modRibbon: RibbonOnLoad
Public AppEvents As CAppEvents        ' single Application event sink

Private Hosts As Collection           ' key = workbook key, item = Workbook
Private HostSheets As Collection      ' key = workbook key, item = Worksheet
Private CurrentKey As String          ' current host key

'======================
' Init / Shutdown
'======================
Public Sub HostManager_Init(Optional ByVal EnableDebug As Boolean = True)
    On Error GoTo EH
    If Initialized Then Exit Sub

    DebugMode = EnableDebug
    Initialized = True

    Set Hosts = New Collection
    Set HostSheets = New Collection
    Set AppEvents = New CAppEvents
    CurrentKey = vbNullString

    DebugPrint "HostManager", "Init", "ready"

    ' Poll until a non-addin workbook+sheet exists
    Application.OnTime Now + TimeValue("00:00:00.5"), "HostManager_PollStartup"
    Exit Sub
EH:
    DebugPrint "HostManager", "Init", "ERROR: " & Err.Description
End Sub

Public Sub HostManager_Shutdown()
    On Error Resume Next
    Set AppEvents = Nothing
    Set Hosts = Nothing
    Set HostSheets = Nothing
    CurrentKey = vbNullString
    Initialized = False
    DebugPrint "HostManager", "Shutdown", "done"
End Sub

'======================
' Registry
'======================
Public Sub HostManager_RegisterHost(ByVal wb As Workbook, ByVal reason As String)
    On Error GoTo EH
    If wb Is Nothing Then Exit Sub
    If HostManager_IsAddinWorkbook(wb) Then Exit Sub

    Dim key$: key$ = HostManager_WorkbookKey(wb)

    ' Add workbook if missing
    On Error Resume Next
    Dim tmp As Workbook: Set tmp = Hosts(key$)
    On Error GoTo EH
    If tmp Is Nothing Then
        Hosts.Add wb, key$
        DebugPrint "HostManager", "RegisterHost", wb.name & " (" & reason & ")"
    End If

    ' Ensure sheet entry exists (may remain Nothing until first sheet activate)
    On Error Resume Next
    Dim wsTmp As Worksheet: Set wsTmp = HostSheets(key$)
    On Error GoTo EH
    If wsTmp Is Nothing Then
        HostSheets.Add Nothing, key$
        DebugPrint "HostManager", "RegisterHost", "sheet placeholder created"
    End If
    Exit Sub
EH:
    DebugPrint "HostManager", "RegisterHost", "ERROR: " & Err.Description
End Sub

Public Sub HostManager_UnregisterHost(ByVal wb As Workbook, ByVal reason As String)
    On Error Resume Next
    If wb Is Nothing Then Exit Sub
    Dim key$: key$ = HostManager_WorkbookKey(wb)

    Hosts.Remove key$
    HostSheets.Remove key$

    If StrComp(CurrentKey, key$, vbTextCompare) = 0 Then
        CurrentKey = vbNullString
        DebugPrint "HostManager", "UnregisterHost", wb.name & " cleared current (" & reason & ")"
        HostManager_RibbonRefreshAll
    Else
        DebugPrint "HostManager", "UnregisterHost", wb.name & " (" & reason & ")"
    End If
End Sub

Private Function HostManager_FindWorkbook(ByVal wb As Workbook) As Workbook
    On Error Resume Next
    Set HostManager_FindWorkbook = Hosts(HostManager_WorkbookKey(wb))
End Function

Private Function HostManager_FindWorkbookByKey(ByVal key As String) As Workbook
    On Error Resume Next
    Set HostManager_FindWorkbookByKey = Hosts(key)
End Function

Private Function HostManager_GetSheetByKey(ByVal key As String) As Worksheet
    On Error Resume Next
    Set HostManager_GetSheetByKey = HostSheets(key)
End Function

Private Sub HostManager_SetSheetByKey(ByVal key As String, ByVal ws As Worksheet)
    On Error Resume Next
    If HostSheets Is Nothing Then Exit Sub
    ' Replace item: remove then add with same key
    HostSheets.Remove key
    HostSheets.Add ws, key
End Sub

'======================
' Activation
'======================
Public Sub HostManager_ActivateWorkbook(ByVal wb As Workbook, ByVal reason As String)
    On Error GoTo EH
    If wb Is Nothing Then Exit Sub
    If HostManager_IsAddinWorkbook(wb) Then Exit Sub

    ' Clean up any sheet-scoped workbook settings (e.g., from transferred sheets)
    CleanSheetScopedWorkbookSettings wb

    HostManager_RegisterHost wb, reason

    Dim key$: key$ = HostManager_WorkbookKey(wb)
    If StrComp(CurrentKey, key$, vbTextCompare) <> 0 Then
        CurrentKey = key$
        DebugPrint "HostManager", "ActivateWorkbook", wb.name & " (" & reason & ")"
        HostManager_RibbonRefreshAll

        ' >>>> UPDATE TRIGGER <<<<
        ' Only schedule version check when workbook actually changes (not on every window focus)
        If GetWorkbookValue(wb, "PyExcelEnabled") = "1" Then
            Application.OnTime Now + TimeValue("00:00:01"), "VerifyProjectVersion"
        End If
    Else
        DebugPrint "HostManager", "ActivateWorkbook", "no change"
    End If
    Exit Sub
EH:
    DebugPrint "HostManager", "ActivateWorkbook", "ERROR: " & Err.Description
End Sub

' Refresh ribbon without triggering version check (for WindowActivate)
' This avoids duplicate version checks when user returns to Excel from another app
Public Sub HostManager_RefreshRibbonOnly(ByVal wb As Workbook, ByVal reason As String)
    On Error GoTo EH
    If wb Is Nothing Then Exit Sub
    If HostManager_IsAddinWorkbook(wb) Then Exit Sub

    HostManager_RegisterHost wb, reason

    Dim key$: key$ = HostManager_WorkbookKey(wb)
    If StrComp(CurrentKey, key$, vbTextCompare) <> 0 Then
        CurrentKey = key$
        DebugPrint "HostManager", "RefreshRibbonOnly", wb.name & " (" & reason & ")"
    End If

    HostManager_RibbonRefreshAll
    Exit Sub
EH:
    DebugPrint "HostManager", "RefreshRibbonOnly", "ERROR: " & Err.Description
End Sub

Public Sub HostManager_ActivateSheet(ByVal wb As Workbook, ByVal ws As Worksheet, ByVal reason As String)
    On Error GoTo EH
    If wb Is Nothing Or ws Is Nothing Then Exit Sub
    If HostManager_IsAddinWorkbook(wb) Then Exit Sub

    HostManager_RegisterHost wb, reason

    Dim key$: key$ = HostManager_WorkbookKey(wb)
    HostManager_SetSheetByKey key, ws
    If StrComp(CurrentKey, key$, vbTextCompare) <> 0 Then CurrentKey = key$

    DebugPrint "HostManager", "ActivateSheet", wb.name & " / " & ws.name & " (" & reason & ")"
    HostManager_RibbonRefreshAll
    Exit Sub
EH:
    DebugPrint "HostManager", "ActivateSheet", "ERROR: " & Err.Description
End Sub

'======================
' Polling until a live host exists
'======================
Public Sub HostManager_PollStartup()
    On Error Resume Next
    Dim wb As Workbook: Set wb = Application.ActiveWorkbook
    Dim sh As Object:   Set sh = Application.ActiveSheet

    If wb Is Nothing Then
        DebugPrint "HostManager", "Poll", "no workbook; retry"
        Application.OnTime Now + TimeValue("00:00:00.5"), "HostManager_PollStartup"
        Exit Sub
    End If
    If HostManager_IsAddinWorkbook(wb) Then
        DebugPrint "HostManager", "Poll", "addin workbook; retry"
        Application.OnTime Now + TimeValue("00:00:00.5"), "HostManager_PollStartup"
        Exit Sub
    End If

    If TypeOf sh Is Worksheet Then
        HostManager_ActivateSheet wb, sh, "Initial Poll"
    Else
        HostManager_ActivateWorkbook wb, "Initial Poll"
    End If
End Sub

'======================
' Accessors for other code
'======================
Public Function HostManager_GetCurrentWorkbook() As Workbook
    On Error Resume Next
    If LenB(CurrentKey) = 0 Then Exit Function
    Set HostManager_GetCurrentWorkbook = HostManager_FindWorkbookByKey(CurrentKey)
End Function

Public Function HostManager_GetCurrentSheet() As Worksheet
    On Error Resume Next
    If LenB(CurrentKey) = 0 Then Exit Function
    Set HostManager_GetCurrentSheet = HostManager_GetSheetByKey(CurrentKey)
End Function

'======================
' Utility
'======================
Public Function HostManager_IsAddinWorkbook(ByVal wb As Workbook) As Boolean
    On Error GoTo Bad
    If wb Is Nothing Then Exit Function
    If wb Is ThisWorkbook Then HostManager_IsAddinWorkbook = True: Exit Function
    If LCase$(right$(wb.name, 5)) = ".xlam" Then HostManager_IsAddinWorkbook = True
    Exit Function
Bad:
    HostManager_IsAddinWorkbook = True
End Function

Private Function HostManager_WorkbookKey(ByVal wb As Workbook) As String
    On Error Resume Next
    If wb.path = "" Then
        HostManager_WorkbookKey = "UNSAVED|" & wb.name
    Else
        HostManager_WorkbookKey = wb.FullName
    End If
End Function

Public Function HostManager_IsWorkbookAlive(ByVal wb As Workbook) As Boolean
    On Error Resume Next
    Dim t$: t$ = wb.name
    HostManager_IsWorkbookAlive = (Err.Number = 0)
    Err.Clear
End Function

Public Function HostManager_IsWorksheetAlive(ByVal ws As Worksheet) As Boolean
    On Error Resume Next
    Dim t$: t$ = ws.name
    HostManager_IsWorksheetAlive = (Err.Number = 0)
    Err.Clear
End Function

Public Sub HostManager_RibbonRefreshAll()
    On Error Resume Next
    If Ribbon Is Nothing Then
        DebugPrint "HostManager", "RibbonRefreshAll", "no ribbon handle"
        Exit Sub
    End If

    With Ribbon
        ' Python group
        .InvalidateControl "cmbScript"
        .InvalidateControl "txtPyInput"
        .InvalidateControl "txtPyOutput"
        
        .InvalidateControl "cmbActions"
        
        .InvalidateControl "btnRun"
        .InvalidateControl "btnEdit"
        .InvalidateControl "btnAdd"
        .InvalidateControl "btnEditAction"
        .InvalidateControl "btnDelete"

        ' Import group
        .InvalidateControl "txtImportInput"
        .InvalidateControl "txtImportOutput"
        .InvalidateControl "btnImport"
        .InvalidateControl "btnEditImport"

        ' Export group
        .InvalidateControl "txtExportInput"
        .InvalidateControl "txtExportOutput"
        .InvalidateControl "btnExport"
        .InvalidateControl "btnEditExport"

        ' Paste group
        .InvalidateControl "txtPasteOutput"
        .InvalidateControl "btnPaste"
        .InvalidateControl "btnEditPaste"

        ' Main group
        .InvalidateControl "btnEnablePyExcel"
        .InvalidateControl "btnConvertToPyExcel"
        .InvalidateControl "btnOpenExplorer"
        .InvalidateControl "btnReadMe"

        ' Final catch-all
        .Invalidate
    End With

    DebugPrint "HostManager", "RibbonRefreshAll", "done"
End Sub

Private Sub DebugPrint(ByVal comp As String, ByVal func As String, ByVal msg As String)
#If DEV Then
    If Not DebugMode Then Exit Sub
    LogToFile "[" & Format$(Now, "yyyy-mm-dd hh:nn:ss") & "] [" & comp & "] [" & func & "] " & msg
#End If
End Sub

' Write a single line to the PyExcel debug log file.
' Safe to call from any module. Fails silently on I/O errors.
Public Sub LogToFile(ByVal logLine As String)
    On Error Resume Next
    Dim f As Integer: f = FreeFile
    Open Environ$("TEMP") & "\PyExcel_Debug.log" For Append As #f
    Print #f, logLine
    Close #f
End Sub

' Returns the full path of the debug log file.
Public Function GetLogFilePath() As String
    GetLogFilePath = Environ$("TEMP") & "\PyExcel_Debug.log"
End Function

'============================================================================================================================
' RESILIENCE / SELF-HEALING
'============================================================================================================================

Public Function HostManager_IsDead() As Boolean
    On Error Resume Next
    HostManager_IsDead = (AppEvents Is Nothing) Or (Initialized = False)
End Function

' Called periodically to verify event bindings and repair if needed
Public Sub HostManager_Watchdog()
    On Error GoTo EH

    ' 1. Event sink check
    If AppEvents Is Nothing Then
        DebugPrint "HostManager", "Watchdog", "event sink missing ? reinitializing"
        HostManager_Init DebugMode
        GoTo Reschedule
    End If

    ' 2. Ensure registry is initialized
    If Hosts Is Nothing Then Set Hosts = New Collection
    If HostSheets Is Nothing Then Set HostSheets = New Collection

    ' 3. Attempt context recovery
    Dim wb As Workbook, ws As Worksheet
    Set wb = Application.ActiveWorkbook
    If Not wb Is Nothing Then
        If Not HostManager_IsAddinWorkbook(wb) Then
            HostManager_RegisterHost wb, "Watchdog"
            If CurrentKey = vbNullString Or Not HostManager_IsWorkbookAlive(HostManager_FindWorkbookByKey(CurrentKey)) Then
                CurrentKey = HostManager_WorkbookKey(wb)
            End If
        End If
    End If

    Set ws = Nothing
    On Error Resume Next
    Set ws = Application.ActiveSheet
    On Error GoTo EH
    If TypeOf ws Is Worksheet Then
        HostManager_SetSheetByKey CurrentKey, ws
    End If

    ' 4. Ribbon refresh: full refresh only when PyExcel is active.
    '    When disabled, CAppEvents handles all context changes adequately.
    '    The watchdog's role is self-healing (steps 1-3), not routine refresh.
    If RibbonIsEnabled Then
        HostManager_RibbonRefreshAll
    End If

Reschedule:
    ' 5. Re-arm timer for continuous monitoring
    On Error Resume Next
    Application.OnTime Now + TimeSerial(0, 0, 10), "HostManager_Watchdog"
    Exit Sub

EH:
    DebugPrint "HostManager", "Watchdog", "ERROR: " & Err.Description
    Err.Clear
    GoTo Reschedule
End Sub


