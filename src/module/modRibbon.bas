Attribute VB_Name = "modRibbon"
Option Explicit

Public scriptSelected As String
Public currentAction As String
Public actionData As Object

Public rib As IRibbonUI
Public currentSheetName As String

Public RibbonIsEnabled As Boolean


'Public Sub RibbonOnLoad(Optional Ribbon As IRibbonUI)
'    On Error GoTo EH
'
'    ' Capture Ribbon object
'    If Not Ribbon Is Nothing Then Set rib = Ribbon
'
'    ' Register Ribbon with HostManager so it can refresh on sheet/workbook events
'    On Error Resume Next
'    Set HostManager.Ribbon = rib
'    On Error GoTo EH
'
'    ' Initialize HostManager if not yet initialized
'    HostManager_Init True
'
'    Dim wb As Workbook
'    Dim ws As Worksheet
'    Set wb = HostManager_GetCurrentWorkbook()
'    Set ws = HostManager_GetCurrentSheet()
'
'    ' Wait or exit if no valid context yet
'    If wb Is Nothing Or ws Is Nothing Then
'        Debug.Print "[RibbonOnLoad] No active workbook/sheet yet. Initialization deferred."
'        Exit Sub
'    End If
'
'    currentSheetName = ws.name
'    Set actionData = CreateObject("Scripting.Dictionary")
'    InitActionsF
'
'    ' Load stored selections safely
'    scriptSelected = GetSheetValue(wb, currentSheetName, "cmbScript")
'    currentAction = GetSheetValue(wb, currentSheetName, "SelectedAction")
'
'    ' Validate stored selections
'    If Len(currentAction) > 0 Then
'        If actionData Is Nothing Then Set actionData = LoadActionsForSheet(currentSheetName)
'        If Not actionData.Exists(currentAction) Then
'            currentAction = ""
'            SaveSheetValue wb, currentSheetName, "SelectedAction", ""
'        End If
'    Else
'        currentAction = ""
'    End If
'
'    ' Refresh ribbon controls
'    If Not rib Is Nothing Then
'        rib.InvalidateControl "cmbActions"
'        rib.InvalidateControl "cmbScript"
'        rib.InvalidateControl "txtPyInput"
'        rib.InvalidateControl "txtPyOutput"
'        rib.InvalidateControl "txtImportInput"
'        rib.InvalidateControl "txtImportOutput"
'        rib.InvalidateControl "txtExportInput"
'        rib.InvalidateControl "txtExportOutput"
'        rib.InvalidateControl "txtPasteOutput"
'        rib.Invalidate
'    End If
'
'    Debug.Print "[RibbonOnLoad] Initialization complete for sheet: " & currentSheetName
'    Exit Sub
'
'EH:
'    Debug.Print "[RibbonOnLoad][ERROR] " & Err.Description
'    Err.Clear
'End Sub

Public Sub RibbonOnLoad(Optional Ribbon As IRibbonUI)
    On Error GoTo EH

    If Not Ribbon Is Nothing Then Set rib = Ribbon

    On Error Resume Next
    Set HostManager.Ribbon = rib
    On Error GoTo EH

    HostManager_Init True

    Call RefreshRibbonValues
    Exit Sub

EH:
    Debug.Print "[RibbonOnLoad][ERROR] " & Err.Description
    Err.Clear
End Sub

Public Sub RefreshRibbonValues()
On Error GoTo EH

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim tmp As String

    Set wb = HostManager_GetCurrentWorkbook()
    Set ws = HostManager_GetCurrentSheet()
    If wb Is Nothing Or ws Is Nothing Then Exit Sub

    currentSheetName = ws.name

    ' ============================================================
    ' LOAD RIBBON ENABLE STATE — DEFAULT = DISABLED
    ' ============================================================
    tmp = GetWorkbookValue(wb, "PyExcelEnabled")

    If tmp = "1" Then
        RibbonIsEnabled = True
    Else
        RibbonIsEnabled = False       ' empty or "0"
    End If

    ' ============================================================
    ' Trigger refresh of controls
    ' ============================================================
    If Not rib Is Nothing Then
        rib.InvalidateControl "cmbActions"
        rib.InvalidateControl "cmbScript"
        rib.InvalidateControl "txtPyInput"
        rib.InvalidateControl "txtPyOutput"
        rib.InvalidateControl "txtImportInput"
        rib.InvalidateControl "txtImportOutput"
        rib.InvalidateControl "txtExportInput"
        rib.InvalidateControl "txtExportOutput"
        rib.InvalidateControl "txtPasteOutput"
        rib.Invalidate
    End If

    Exit Sub

EH:
    Debug.Print "[RefreshRibbonValues][ERROR] " & Err.Description
    Err.Clear
End Sub




Public Sub OnPyTabSelect(control As IRibbonControl, ByRef label)
    On Error Resume Next

    Dim wb As Workbook: Set wb = Application.ActiveWorkbook
    Dim sh As Object:   Set sh = Application.ActiveSheet

    If Not wb Is Nothing Then
        If Not HostManager_IsAddinWorkbook(wb) Then
            If TypeOf sh Is Worksheet Then
                HostManager_ActivateSheet wb, sh, "TabSelect"
            Else
                HostManager_ActivateWorkbook wb, "TabSelect"
            End If
        End If
    End If

    HostManager_RibbonRefreshAll
    label = vbNullString
End Sub




' ==============================================================
' SHEET VALUE HELPERS — with detailed debugging
' ==============================================================

Public Function GetSheetValue(wb As Workbook, sheetName As String, ctrlId As String) As String
    On Error GoTo EH

    Dim ws As Worksheet
    Dim nm As name
    Dim raw As String

    ' Use resilient workbook and sheet references
    If wb Is Nothing Then Set wb = HostManager_GetCurrentWorkbook()
    If wb Is Nothing Then
        Debug.Print "[GetSheetValue] Workbook reference is Nothing (after HostManager check)."
        Exit Function
    End If
    If LenB(sheetName) = 0 Then
        Dim tmpWS As Worksheet
        Set tmpWS = HostManager_GetCurrentSheet()
        If Not tmpWS Is Nothing Then sheetName = tmpWS.name
    End If

    Debug.Print "[GetSheetValue] wb=" & wb.name & " | sheet=" & sheetName & " | ctrlId=" & ctrlId

    ' Try to get the worksheet safely
    Set ws = Nothing
    On Error Resume Next
    Set ws = wb.Sheets(sheetName)
    On Error GoTo EH

    If ws Is Nothing Then
        Debug.Print "[GetSheetValue] Sheet '" & sheetName & "' not found in workbook '" & wb.name & "'."
        Exit Function
    End If

    ' Try to get the named range within the sheet
    Set nm = Nothing
    On Error Resume Next
    Set nm = ws.Names(ctrlId)
    On Error GoTo EH

    If nm Is Nothing Then
        Debug.Print "[GetSheetValue] Name '" & ctrlId & "' not found on sheet '" & sheetName & "'."
        Exit Function
    End If

    raw = nm.RefersTo
    Debug.Print "[GetSheetValue] Raw value: " & raw

    ' Normalize the string result
    If Left$(raw, 1) = "=" Then raw = Mid$(raw, 2)
    If Left$(raw, 1) = """" And right$(raw, 1) = """" Then
        raw = Mid$(raw, 2, Len(raw) - 2)
    End If

    Debug.Print "[GetSheetValue] Final value: " & raw
    GetSheetValue = raw
    Exit Function

EH:
    Debug.Print "[GetSheetValue][ERROR] sheet='" & sheetName & "' ctrl='" & ctrlId & "' err=" & Err.Description
    Err.Clear
End Function


Public Sub SaveSheetValue(wb As Workbook, sheetName As String, ctrlId As String, val As Variant)
    On Error GoTo EH

    ' Use resilient workbook and sheet references
    If wb Is Nothing Then Set wb = HostManager_GetCurrentWorkbook()
    If wb Is Nothing Then
        Debug.Print "[SaveSheetValue] Workbook reference is Nothing (after HostManager check)."
        Exit Sub
    End If
    If LenB(sheetName) = 0 Then
        Dim tmpWS As Worksheet
        Set tmpWS = HostManager_GetCurrentSheet()
        If Not tmpWS Is Nothing Then sheetName = tmpWS.name
    End If

    Debug.Print "[SaveSheetValue] wb=" & wb.name & " | sheet=" & sheetName & _
                " | ctrlId=" & ctrlId & " | val='" & val & "'"

    Dim ws As Worksheet
    Set ws = Nothing
    On Error Resume Next
    Set ws = wb.Sheets(sheetName)
    On Error GoTo EH

    If ws Is Nothing Then
        Debug.Print "[SaveSheetValue] Sheet '" & sheetName & "' not found in workbook '" & wb.name & "'."
        Exit Sub
    End If

    ' Remove any old name
    On Error Resume Next
    ws.Names(ctrlId).Delete
    On Error GoTo EH

    ' Add the new name if value is not empty
    If Len(val) > 0 Then
        ws.Names.Add name:=ctrlId, RefersTo:="=""" & Replace(val, """", """""") & """"
        Debug.Print "[SaveSheetValue] Saved '" & ctrlId & "' = '" & val & "' to sheet '" & sheetName & "'."
    Else
        Debug.Print "[SaveSheetValue] Value empty — nothing saved for '" & ctrlId & "'."
    End If

    Exit Sub

EH:
    Debug.Print "[SaveSheetValue][ERROR] sheet='" & sheetName & "' ctrl='" & ctrlId & "' err=" & Err.Description
    Err.Clear
End Sub



' ==============================================================
' MAIN BUTTONS
' ==============================================================
'Public Sub OnEnablePyExcel(control As IRibbonControl)
'    On Error GoTo EH
'
'    Dim wb As Workbook
'    Dim ws As Worksheet
'
'    Set wb = HostManager_GetCurrentWorkbook()
'    Set ws = HostManager_GetCurrentSheet()
'
'    If wb Is Nothing Then
'        MsgBox "No active workbook context.", vbExclamation
'        Exit Sub
'    End If
'
'    '------------------------------------------------------------
'    ' If currently disabled → attempt setup before enabling
'    '------------------------------------------------------------
'    If RibbonIsEnabled = False Then
'        ' Run the installation logic
'        If PyExcelSetup() = False Then
'            ' Setup failed, stay disabled
'            RibbonIsEnabled = False
'            ' Save state (will save disabled)
'            SaveRibbonState
'            If Not rib Is Nothing Then rib.Invalidate
'            Exit Sub
'        End If
'
'        ' Setup succeeded → enable
'        RibbonIsEnabled = True
'    Else
'        ' Currently enabled → toggle OFF
'        RibbonIsEnabled = False
'    End If
'
'    '------------------------------------------------------------
'    ' Save this state per workbook
'    '------------------------------------------------------------
'    SaveRibbonState
'
'    '------------------------------------------------------------
'    ' Refresh ALL ribbon controls
'    '------------------------------------------------------------
'    If Not rib Is Nothing Then rib.Invalidate
'
'    Exit Sub
'
'EH:
'    Debug.Print "OnEnablePyExcel error: " & Err.Description
'End Sub


'Public Sub OnEnablePyExcel(control As IRibbonControl)
'    On Error GoTo EH
'
'    Dim wb As Workbook
'    Dim ws As Worksheet
'
'    Set wb = HostManager_GetCurrentWorkbook()
'    Set ws = HostManager_GetCurrentSheet()
'
'    If wb Is Nothing Then
'        MsgBox "No active workbook context.", vbExclamation
'        Exit Sub
'    End If
'
'    '------------------------------------------------------------
'    ' If currently disabled â†’ attempt setup before enabling
'    '------------------------------------------------------------
'    If RibbonIsEnabled = False Then
'        ' Run the installation logic
'        Dim setupResult As Boolean
'        setupResult = PyExcelSetup()
'        If setupResult = False Then
'            Debug.Print "[OnEnablePyExcel] Setup failed: " & PyExcelSetup_LastMessage
'            ' Setup failed, stay disabled
'            RibbonIsEnabled = False
'            ' Save state (will save disabled)
'            SaveRibbonState
'            If Not rib Is Nothing Then rib.Invalidate
'            Exit Sub
'        End If
'
'        ' Setup succeeded â†’ enable
'        RibbonIsEnabled = True
'    Else
'        ' Currently enabled â†’ toggle OFF
'        RibbonIsEnabled = False
'    End If
'
'    '------------------------------------------------------------
'    ' Save this state per workbook
'    '------------------------------------------------------------
'    SaveRibbonState
'
'    '------------------------------------------------------------
'    ' Refresh ALL ribbon controls
'    '------------------------------------------------------------
'    If Not rib Is Nothing Then rib.Invalidate
'
'    Exit Sub
'
'EH:
'    Debug.Print "OnEnablePyExcel error: " & Err.Description
'End Sub



Public Sub OnEnablePyExcel(control As IRibbonControl)
    On Error GoTo EH

    Dim wb As Workbook
    Set wb = HostManager_GetCurrentWorkbook()
    
    ' 1. Validate Context
    If wb Is Nothing Then
        MsgBox "No active workbook context.", vbExclamation
        Exit Sub
    End If
    
    ' 2. Determine Action based on current state
    If RibbonIsEnabled = False Then
        ' Attempting to ENABLE
        Dim setupResult As Boolean
        setupResult = PyExcelSetup()
        
        If setupResult = False Then
            ' Setup failed: Notify user and remain disabled
            MsgBox "Unable to enable PyExcel. Setup failed: " & vbNewLine & _
                   PyExcelSetup_LastMessage, vbCritical, "Setup Error"
            RibbonIsEnabled = False
        Else
            ' Setup succeeded: Enable
            RibbonIsEnabled = True
        End If
    Else
        ' Currently enabled: Toggle OFF
        RibbonIsEnabled = False
    End If

    ' 3. Centralized State Saving and UI Refresh
    SaveRibbonState
    
    If Not rib Is Nothing Then rib.Invalidate

    Exit Sub

EH:
    MsgBox "An error occurred in OnEnablePyExcel: " & Err.Description, vbCritical
    Debug.Print "OnEnablePyExcel error: " & Err.Description
End Sub



Public Sub OnOpenExplorer(control As IRibbonControl)
    On Error GoTo EH

    Dim wb As Workbook
    Dim folderPath As String

    Set wb = HostManager_GetCurrentWorkbook()
    If wb Is Nothing Then
        MsgBox "No active workbook.", vbExclamation
        Exit Sub
    End If

    folderPath = ResolveProjectPath()
    If Len(folderPath) = 0 Then
        MsgBox "Workbook not saved. Save it first to create a folder path.", vbExclamation
        Exit Sub
    End If

    ' Open the workbook's folder in Windows Explorer
    Shell "explorer.exe """ & folderPath & """", vbNormalFocus
    Exit Sub

EH:
    Debug.Print "OnOpenExplorer error: " & Err.Description
End Sub


Public Sub OnReadMe(control As IRibbonControl)
    On Error GoTo EH

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim folderPath As String
    Dim readmePath As String

    ' Use resilient host references
    Set wb = HostManager_GetCurrentWorkbook()
    Set ws = HostManager_GetCurrentSheet()

    If wb Is Nothing Then
        Debug.Print "[OnReadMe] No active workbook (HostManager returned Nothing)."
        MsgBox "No active workbook.", vbExclamation
        Exit Sub
    End If

    folderPath = ResolveProjectPath()
    If Len(folderPath) = 0 Then
        Debug.Print "[OnReadMe] Workbook not saved — cannot resolve path."
        MsgBox "Workbook not saved. Save it first to create a folder path.", vbExclamation
        Exit Sub
    End If

    ' Build path to ReadMe.txt
    If right$(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    readmePath = folderPath & "ReadMe.txt"

    ' Check file existence
    If Dir(readmePath) = "" Then
        Debug.Print "[OnReadMe] Missing file: " & readmePath
        MsgBox "ReadMe.txt not found in: " & folderPath, vbExclamation
        Exit Sub
    End If

    ' Open with system default text editor (Notepad or associated app)
    Debug.Print "[OnReadMe] Opening file: " & readmePath
    Shell "cmd /c start """" """ & readmePath & """", vbHide
    Exit Sub

EH:
    Debug.Print "[OnReadMe][ERROR] " & Err.Description
    Err.Clear
End Sub


' ==============================================================
' IMPORT
' ==============================================================

Public Sub OnImport(control As IRibbonControl)
    On Error GoTo EH

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim sourcePath As String
    Dim destAddress As String

    Set wb = HostManager_GetCurrentWorkbook()
    Set ws = HostManager_GetCurrentSheet()
    If wb Is Nothing Or ws Is Nothing Then
        MsgBox "No active workbook or sheet context.", vbExclamation
        Exit Sub
    End If

    currentSheetName = ws.name
    sourcePath = GetSheetValue(wb, currentSheetName, "txtImportInput")
    destAddress = GetSheetValue(wb, currentSheetName, "txtImportOutput")

    RunImportForSheet currentSheetName, sourcePath, destAddress
    Exit Sub

EH:
    Debug.Print "OnImport error: " & Err.Description
End Sub


'<editBox id="txtImportInput"  label="Source"  getText="GetImportInput"  sizeString="XXXXXXXXXXXXXXXXXXXX" onChange="OnImportInputChange"/>
Public Sub GetImportInput(control As IRibbonControl, ByRef returnedText)
    On Error GoTo EH

    Dim wb As Workbook
    Set wb = HostManager_GetCurrentWorkbook()
    If wb Is Nothing Then Exit Sub

    returnedText = GetSheetValue(wb, currentSheetName, control.id)
    If Len(returnedText) = 0 Then returnedText = ""
    Exit Sub

EH:
    Debug.Print "GetImportInput error: " & Err.Description
End Sub


Public Sub OnImportInputChange(control As IRibbonControl, text As String)
    On Error GoTo EH
    ' Revert user edits to stored value (UI consistency)
    If Not rib Is Nothing Then rib.InvalidateControl control.id
    Exit Sub

EH:
    Debug.Print "OnImportInputChange error: " & Err.Description
End Sub


Public Sub UpdateImportInput(newValue As String)
    On Error GoTo EH

    Dim wb As Workbook
    Set wb = HostManager_GetCurrentWorkbook()
    If wb Is Nothing Then Exit Sub

    SaveSheetValue wb, currentSheetName, "txtImportInput", newValue
    If Not rib Is Nothing Then rib.InvalidateControl "txtImportInput"
    Exit Sub

EH:
    Debug.Print "UpdateImportInput error: " & Err.Description
End Sub


'<editBox id="txtImportOutput" label="Destination" getText="GetImportOutput" sizeString="XXXXXXXXXXXXXXXXXXXX" onChange="OnImportOutputChange"/>
Public Sub GetImportOutput(control As IRibbonControl, ByRef returnedText)
    On Error GoTo EH

    Dim wb As Workbook
    Set wb = HostManager_GetCurrentWorkbook()
    If wb Is Nothing Then Exit Sub

    returnedText = GetSheetValue(wb, currentSheetName, control.id)
    If Len(returnedText) = 0 Then returnedText = ""
    Exit Sub

EH:
    Debug.Print "GetImportOutput error: " & Err.Description
End Sub


Public Sub OnImportOutputChange(control As IRibbonControl, text As String)
    On Error GoTo EH
    ' Revert user edits to stored value
    If Not rib Is Nothing Then rib.InvalidateControl control.id
    Exit Sub

EH:
    Debug.Print "OnImportOutputChange error: " & Err.Description
End Sub


Public Sub UpdateImportOutput(newValue As String)
    On Error GoTo EH

    Dim wb As Workbook
    Set wb = HostManager_GetCurrentWorkbook()
    If wb Is Nothing Then Exit Sub

    SaveSheetValue wb, currentSheetName, "txtImportOutput", newValue
    If Not rib Is Nothing Then rib.InvalidateControl "txtImportOutput"
    Exit Sub

EH:
    Debug.Print "UpdateImportOutput error: " & Err.Description
End Sub


Public Sub OnEditImport(control As IRibbonControl)
    On Error GoTo EH
    frmEditImport.Show
    Exit Sub

EH:
    Debug.Print "OnEditImport error: " & Err.Description
End Sub


'Public Sub GetImportInput(control As IRibbonControl, ByRef returnedText)
'    returnedText = GetSheetValue(ActiveWorkbook, currentSheetName, control.id)
'    If Len(returnedText) = 0 Then returnedText = ""
'End Sub
'
'Public Sub OnImportInputChange(control As IRibbonControl, text As String)
'    rib.InvalidateControl "txtImportInput"
'End Sub
'
'Public Sub GetImportOutput(control As IRibbonControl, ByRef returnedText)
'    returnedText = GetSheetValue(ActiveWorkbook, currentSheetName, control.id)
'    If Len(returnedText) = 0 Then returnedText = ""
'End Sub
'
'Public Sub OnImportOutputChange(control As IRibbonControl, text As String)
'    rib.InvalidateControl "txtImportOutput"
'End Sub

' ==============================================================
' EXPORT
' ==============================================================

Public Sub OnExport(control As IRibbonControl)
    On Error GoTo EH

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim sourcePath As String
    Dim destAddress As String

    Set wb = HostManager_GetCurrentWorkbook()
    Set ws = HostManager_GetCurrentSheet()
    If wb Is Nothing Or ws Is Nothing Then
        MsgBox "No active workbook or sheet context.", vbExclamation
        Exit Sub
    End If

    currentSheetName = ws.name
    sourcePath = GetSheetValue(wb, currentSheetName, "txtExportInput")
    destAddress = GetSheetValue(wb, currentSheetName, "txtExportOutput")

    RunExportForSheet currentSheetName, sourcePath, destAddress
    Exit Sub

EH:
    Debug.Print "OnExport error: " & Err.Description
End Sub


'<editBox id="txtExportInput"  label="Source"  getText="GetExportInput"  sizeString="XXXXXXXXXXXXXXXXXXXX" onChange="OnExportInputChange"/>
Public Sub GetExportInput(control As IRibbonControl, ByRef returnedText)
    On Error GoTo EH

    Dim wb As Workbook
    Set wb = HostManager_GetCurrentWorkbook()
    If wb Is Nothing Then Exit Sub

    returnedText = GetSheetValue(wb, currentSheetName, control.id)
    If Len(returnedText) = 0 Then returnedText = ""
    Exit Sub

EH:
    Debug.Print "GetExportInput error: " & Err.Description
End Sub


Public Sub OnExportInputChange(control As IRibbonControl, text As String)
    On Error GoTo EH
    If Not rib Is Nothing Then rib.InvalidateControl control.id
    Exit Sub

EH:
    Debug.Print "OnExportInputChange error: " & Err.Description
End Sub


Public Sub UpdateExportInput(newValue As String)
    On Error GoTo EH

    Dim wb As Workbook
    Set wb = HostManager_GetCurrentWorkbook()
    If wb Is Nothing Then Exit Sub

    SaveSheetValue wb, currentSheetName, "txtExportInput", newValue
    If Not rib Is Nothing Then rib.InvalidateControl "txtExportInput"
    Exit Sub

EH:
    Debug.Print "UpdateExportInput error: " & Err.Description
End Sub


'<editBox id="txtExportOutput" label="Destination" getText="GetExportOutput" sizeString="XXXXXXXXXXXXXXXXXXXX" onChange="OnExportOutputChange"/>
Public Sub GetExportOutput(control As IRibbonControl, ByRef returnedText)
    On Error GoTo EH

    Dim wb As Workbook
    Set wb = HostManager_GetCurrentWorkbook()
    If wb Is Nothing Then Exit Sub

    returnedText = GetSheetValue(wb, currentSheetName, control.id)
    If Len(returnedText) = 0 Then returnedText = ""
    Exit Sub

EH:
    Debug.Print "GetExportOutput error: " & Err.Description
End Sub


Public Sub OnExportOutputChange(control As IRibbonControl, text As String)
    On Error GoTo EH
    If Not rib Is Nothing Then rib.InvalidateControl control.id
    Exit Sub

EH:
    Debug.Print "OnExportOutputChange error: " & Err.Description
End Sub


Public Sub UpdateExportOutput(newValue As String)
    On Error GoTo EH

    Dim wb As Workbook
    Set wb = HostManager_GetCurrentWorkbook()
    If wb Is Nothing Then Exit Sub

    SaveSheetValue wb, currentSheetName, "txtExportOutput", newValue
    If Not rib Is Nothing Then rib.InvalidateControl "txtExportOutput"
    Exit Sub

EH:
    Debug.Print "UpdateExportOutput error: " & Err.Description
End Sub


Public Sub OnEditExport(control As IRibbonControl)
    On Error GoTo EH
    frmEditExport.Show
    Exit Sub

EH:
    Debug.Print "OnEditExport error: " & Err.Description
End Sub

' ==============================================================
' PASTE
' ==============================================================

Public Sub OnPaste(control As IRibbonControl)
    On Error GoTo EH

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim targetAddress As String

    Set wb = HostManager_GetCurrentWorkbook()
    Set ws = HostManager_GetCurrentSheet()
    If wb Is Nothing Or ws Is Nothing Then
        MsgBox "No active workbook or sheet context.", vbExclamation
        Exit Sub
    End If

    currentSheetName = ws.name
    targetAddress = GetSheetValue(wb, currentSheetName, "txtPasteOutput")

    If Len(targetAddress) = 0 Then
        MsgBox "Please specify a target address in Paste Output.", vbExclamation
        Exit Sub
    End If

    ' Call the centralized paste function
    PasteFromClipboardToSheet currentSheetName, targetAddress
    Exit Sub

EH:
    Debug.Print "OnPaste error: " & Err.Description
End Sub


Public Sub GetPasteOutput(control As IRibbonControl, ByRef returnedText)
    On Error GoTo EH

    Dim wb As Workbook
    Set wb = HostManager_GetCurrentWorkbook()
    If wb Is Nothing Then Exit Sub

    returnedText = GetSheetValue(wb, currentSheetName, control.id)
    If Len(returnedText) = 0 Then returnedText = ""
    Exit Sub

EH:
    Debug.Print "GetPasteOutput error: " & Err.Description
End Sub


Public Sub OnPasteOutputChange(control As IRibbonControl, text As String)
    On Error GoTo EH
    ' Revert user edits to stored value
    If Not rib Is Nothing Then rib.InvalidateControl control.id
    Exit Sub

EH:
    Debug.Print "OnPasteOutputChange error: " & Err.Description
End Sub


Public Sub UpdatePasteOutput(newValue As String)
    On Error GoTo EH

    Dim wb As Workbook
    Set wb = HostManager_GetCurrentWorkbook()
    If wb Is Nothing Then Exit Sub

    SaveSheetValue wb, currentSheetName, "txtPasteOutput", newValue
    If Not rib Is Nothing Then rib.InvalidateControl "txtPasteOutput"
    Exit Sub

EH:
    Debug.Print "UpdatePasteOutput error: " & Err.Description
End Sub


Public Sub OnEditPaste(control As IRibbonControl)
    On Error GoTo EH
    frmEditPaste.Show
    Exit Sub

EH:
    Debug.Print "OnEditPaste error: " & Err.Description
End Sub







' ==============================================================
' PYTHON
' ==============================================================

Public Sub GetPyInput(control As IRibbonControl, ByRef returnedText)
    On Error GoTo EH

    Dim wb As Workbook
    Dim ws As Worksheet

    Set wb = HostManager_GetCurrentWorkbook()
    Set ws = HostManager_GetCurrentSheet()
    If wb Is Nothing Or ws Is Nothing Then Exit Sub

    ' Always sync sheet context before reading
    currentSheetName = ws.name

    returnedText = GetSheetValue(wb, currentSheetName, control.id)
    If Len(returnedText) = 0 Then returnedText = ""
    Exit Sub

EH:
    Debug.Print "GetPyInput error: " & Err.Description
End Sub

Public Sub OnPyInputChange(control As IRibbonControl, text As String)
    On Error GoTo EH
    ' Revert user edits to stored value
    If Not rib Is Nothing Then rib.InvalidateControl control.id
    Exit Sub

EH:
    Debug.Print "OnPyInputChange error: " & Err.Description
End Sub


Public Sub UpdatePyInput(newValue As String)
    On Error GoTo EH

    Dim wb As Workbook
    Set wb = HostManager_GetCurrentWorkbook()
    If wb Is Nothing Then Exit Sub

    SaveSheetValue wb, currentSheetName, "txtPyInput", newValue
    If Not rib Is Nothing Then rib.InvalidateControl "txtPyInput"
    Exit Sub

EH:
    Debug.Print "UpdatePyInput error: " & Err.Description
End Sub


Public Sub GetPyOutput(control As IRibbonControl, ByRef returnedText)
    On Error GoTo EH

    Dim wb As Workbook
    Set wb = HostManager_GetCurrentWorkbook()
    If wb Is Nothing Then Exit Sub

    returnedText = GetSheetValue(wb, currentSheetName, control.id)
    If Len(returnedText) = 0 Then returnedText = ""
    Exit Sub

EH:
    Debug.Print "GetPyOutput error: " & Err.Description
End Sub


Public Sub OnPyOutputChange(control As IRibbonControl, text As String)
    On Error GoTo EH
    ' Revert user edits to stored value
    If Not rib Is Nothing Then rib.InvalidateControl control.id
    Exit Sub

EH:
    Debug.Print "OnPyOutputChange error: " & Err.Description
End Sub


Public Sub UpdatePyOutput(newValue As String)
    On Error GoTo EH

    Dim wb As Workbook
    Set wb = HostManager_GetCurrentWorkbook()
    If wb Is Nothing Then Exit Sub

    SaveSheetValue wb, currentSheetName, "txtPyOutput", newValue
    If Not rib Is Nothing Then rib.InvalidateControl "txtPyOutput"
    Exit Sub

EH:
    Debug.Print "UpdatePyOutput error: " & Err.Description
End Sub

' ==============================================================
' SCRIPT COMBO BOX  (context-safe)
' ==============================================================

Public Function GetScriptFolderPath() As String
    On Error GoTo EH

    Dim wb As Workbook
    Dim pathBase As String

    Set wb = HostManager_GetCurrentWorkbook()
    If wb Is Nothing Then Exit Function

    pathBase = ResolveProjectPath()
    If Len(pathBase) = 0 Then Exit Function

    GetScriptFolderPath = pathBase & "\userScripts"
    Debug.Print "Configured host workbook path: " & GetScriptFolderPath
    Exit Function

EH:
    Debug.Print "GetScriptFolderPath error: " & Err.Description
End Function


Public Function GetScriptFiles() As Collection
    On Error GoTo EH

    Dim col As New Collection
    Dim fso As Object, folder As Object, file As Object
    Dim folderPath As String

    folderPath = GetScriptFolderPath()
    If Len(folderPath) = 0 Then Exit Function
    If Len(Dir(folderPath, vbDirectory)) = 0 Then Exit Function

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(folderPath)

    For Each file In folder.files
        col.Add file.name
    Next file

    Set GetScriptFiles = col
    Exit Function

EH:
    Debug.Print "GetScriptFiles error: " & Err.Description
End Function


Public Sub GetScriptCount(control As IRibbonControl, ByRef count)
    On Error GoTo EH

    Dim files As Collection
    Set files = GetScriptFiles()
    If files Is Nothing Then
        count = 0
    Else
        count = files.count
    End If
    Exit Sub

EH:
    Debug.Print "GetScriptCount error: " & Err.Description
End Sub


Public Sub GetScriptLabel(control As IRibbonControl, index As Integer, ByRef returnedVal)
    On Error GoTo EH

    Dim files As Collection
    Set files = GetScriptFiles()
    If Not files Is Nothing Then
        If index >= 0 And index < files.count Then
            returnedVal = files(index + 1)
        End If
    End If
    Exit Sub

EH:
    Debug.Print "GetScriptLabel error: " & Err.Description
End Sub


Public Sub GetScriptText(control As IRibbonControl, ByRef returnedVal)
    On Error GoTo EH

    Dim wb As Workbook
    Dim ws As Worksheet

    Set wb = HostManager_GetCurrentWorkbook()
    Set ws = HostManager_GetCurrentSheet()
    If wb Is Nothing Or ws Is Nothing Then
        returnedVal = ""
        Exit Sub
    End If

    ' Always read the live value from the sheet, not the cached dictionary
    returnedVal = GetSheetValue(wb, ws.name, "cmbScript")
    scriptSelected = returnedVal
    If Len(returnedVal) = 0 Then returnedVal = ""
    Exit Sub

EH:
    Debug.Print "[GetScriptText][ERROR] " & Err.Description
    returnedVal = ""
End Sub


Public Sub OnScriptChange(control As IRibbonControl, text As String)
    On Error GoTo EH

    Dim wb As Workbook
    Dim ws As Worksheet

    Set wb = HostManager_GetCurrentWorkbook()
    Set ws = HostManager_GetCurrentSheet()
    If wb Is Nothing Or ws Is Nothing Then Exit Sub

    currentSheetName = ws.name

    If Len(text) = 0 Then
        SaveSheetValue wb, currentSheetName, control.id, ""
        UpdateCurrentActionField "script", ""
        scriptSelected = ""
        If Not rib Is Nothing Then rib.InvalidateControl control.id
        Exit Sub
    End If

    SaveSheetValue wb, currentSheetName, control.id, text
    UpdateCurrentActionField "script", text
    scriptSelected = text
    If Not rib Is Nothing Then rib.InvalidateControl control.id
    Exit Sub

EH:
    Debug.Print "OnScriptChange error: " & Err.Description
End Sub

' ==============================================================
' PYTHON ACTION MANAGEMENT
' ==============================================================

Public Function LoadActionsForSheet(sheetName As String) As Object
    On Error GoTo EH

    Dim wb As Workbook
    Dim raw As String, rows As Variant, cols As Variant, i As Long
    Dim dict As Object, act As Object, k As String

    Set wb = HostManager_GetCurrentWorkbook()
    If wb Is Nothing Then Exit Function

    Set dict = CreateObject("Scripting.Dictionary")

    raw = GetSheetValue(wb, sheetName, "Actions")
    If Len(raw) = 0 Then
        Set LoadActionsForSheet = dict
        Exit Function
    End If

    rows = Split(raw, ";")
    For i = LBound(rows) To UBound(rows)
        If Len(rows(i)) > 0 Then
            cols = Split(rows(i), "|")
            If UBound(cols) >= 3 Then
                k = Trim$(cols(0))
                If Len(k) > 0 Then
                    Set act = CreateObject("Scripting.Dictionary")
                    act("script") = Trim$(cols(1))
                    act("input") = Trim$(cols(2))
                    act("output") = Trim$(cols(3))
                    dict.Add k, act
                End If
            End If
        End If
    Next i

    Set LoadActionsForSheet = dict
    Exit Function

EH:
    Debug.Print "LoadActionsForSheet error: " & Err.Description
    Set LoadActionsForSheet = Nothing
End Function


Public Sub SaveActionsForSheet(sheetName As String, dict As Object)
    On Error GoTo EH

    Dim wb As Workbook
    Dim k As Variant
    Dim s As String

    Set wb = HostManager_GetCurrentWorkbook()
    If wb Is Nothing Then Exit Sub
    If dict Is Nothing Then Exit Sub

    Debug.Print "SaveActionsForSheet", "dict.count=" & dict.count, "currentAction=" & currentAction

    For Each k In dict.keys
        s = s & k & "|" & dict(k)("script") & "|" & dict(k)("input") & "|" & dict(k)("output") & ";"
    Next k

    SaveSheetValue wb, sheetName, "Actions", s
    Exit Sub

EH:
    Debug.Print "SaveActionsForSheet error: " & Err.Description
End Sub


Public Sub InitActions()
    On Error GoTo EH

    Set actionData = LoadActionsForSheet(currentSheetName)
    Exit Sub

EH:
    Debug.Print "InitActions error: " & Err.Description
End Sub

' ==============================================================
' ACTION BUTTONS
' ==============================================================

Public Sub OnAddAction(control As IRibbonControl)
    On Error GoTo EH

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim newName As String
    Dim act As Object

    Set wb = HostManager_GetCurrentWorkbook()
    Set ws = HostManager_GetCurrentSheet()
    If wb Is Nothing Or ws Is Nothing Then
        MsgBox "No active workbook or sheet context.", vbExclamation
        Exit Sub
    End If

    currentSheetName = ws.name

    ' Ensure actionData is a valid dictionary
    If actionData Is Nothing Then Set actionData = LoadActionsForSheet(currentSheetName)
    If TypeName(actionData) <> "Dictionary" Then Set actionData = CreateObject("Scripting.Dictionary")

    newName = Trim$(InputBox("Enter new action name:"))
    If Len(newName) = 0 Then Exit Sub
    If actionData.Exists(newName) Then
        MsgBox "Action '" & newName & "' already exists.", vbInformation
        Exit Sub
    End If

    ' Create new action entry
    Set act = CreateObject("Scripting.Dictionary")
    act("script") = ""
    act("input") = ""
    act("output") = ""
    actionData.Add newName, act

    ' Persist data
    SaveActionsForSheet currentSheetName, actionData
    currentAction = newName
    SaveSheetValue wb, currentSheetName, "SelectedAction", newName
    SaveSheetValue wb, currentSheetName, "cmbScript", ""
    SaveSheetValue wb, currentSheetName, "txtPyInput", ""
    SaveSheetValue wb, currentSheetName, "txtPyOutput", ""

    ' Refresh ribbon controls
    If Not rib Is Nothing Then
        rib.InvalidateControl "cmbActions"
        rib.InvalidateControl "cmbScript"
        rib.InvalidateControl "txtPyInput"
        rib.InvalidateControl "txtPyOutput"
    End If

    MsgBox "Action '" & newName & "' added successfully.", vbInformation
    Exit Sub

EH:
    Debug.Print "OnAddAction error: " & Err.Description
End Sub
Public Sub OnEditPython(control As IRibbonControl)
    On Error GoTo EH

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim scriptName As String
    Dim scriptPath As String

    Set wb = HostManager_GetCurrentWorkbook()
    Set ws = HostManager_GetCurrentSheet()
    If wb Is Nothing Or ws Is Nothing Then
        MsgBox "No active workbook or sheet context.", vbExclamation
        Exit Sub
    End If

    ' Get the script filename from the current selection or sheet
    scriptName = GetSheetValue(wb, ws.name, "cmbScript")
    If Len(scriptName) = 0 Then
        MsgBox "No script selected.", vbExclamation
        Exit Sub
    End If

    ' Resolve full path to the script folder
    scriptPath = GetScriptFolderPath()
    If Len(scriptPath) = 0 Then
        MsgBox "Script folder not found.", vbExclamation
        Exit Sub
    End If

    If right$(scriptPath, 1) <> "\" Then scriptPath = scriptPath & "\"
    scriptPath = scriptPath & scriptName

    ' Verify file existence
    If Dir(scriptPath) = "" Then
        MsgBox "Script file not found: " & scriptPath, vbExclamation
        Exit Sub
    End If

    ' Open with system default .py handler
    Shell "cmd /c start """" """ & scriptPath & """", vbHide
    Exit Sub

EH:
    Debug.Print "OnEditPython error: " & Err.Description
End Sub


Public Sub OnEditAction(control As IRibbonControl)
    On Error GoTo EH
    frmEditAction.Show
    Exit Sub

EH:
    Debug.Print "OnEditAction error: " & Err.Description
End Sub






' ==============================================================
' ==============================================================
' RUN Python Action
' ==============================================================
' ==============================================================
Public Sub OnRunPython(control As IRibbonControl)
    On Error GoTo EH

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim scriptName As String
    Dim inputVal As String
    Dim outputVal As String

    Set wb = HostManager_GetCurrentWorkbook()
    Set ws = HostManager_GetCurrentSheet()
    If wb Is Nothing Or ws Is Nothing Then
        MsgBox "No active workbook or sheet context.", vbExclamation
        Exit Sub
    End If

    currentSheetName = ws.name

    scriptName = GetSheetValue(wb, currentSheetName, "cmbScript")
    inputVal = GetSheetValue(wb, currentSheetName, "txtPyInput")
    outputVal = GetSheetValue(wb, currentSheetName, "txtPyOutput")

    ' Display debug information to the user
    Debug.Print "Action: " & currentAction & vbCrLf & _
           "Script: " & scriptName & vbCrLf & _
           "Input: " & inputVal & vbCrLf & _
           "Output: " & outputVal, vbInformation, "Run Python Action"
           
    MsgBox "Action: " & currentAction & vbCrLf & _
           "Script: " & scriptName & vbCrLf & _
           "Input: " & inputVal & vbCrLf & _
           "Output: " & outputVal, vbInformation, "Run Python Action"

    ' Run the linked Python script
    RunGenericPythonScript scriptName, inputVal, outputVal, wb, ws
    Exit Sub

EH:
    Debug.Print "OnRunPython error: " & Err.Description
    MsgBox "Error running Python action: " & Err.Description, vbExclamation, "Run Python"
End Sub

' ==============================================================
' ==============================================================




' ==============================================================
' ACTION COMBO BOX — context-safe
' ==============================================================

Private Function GetActionList() As Collection
    On Error GoTo EH

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim col As New Collection
    Dim dict As Object
    Dim k As Variant

    Set wb = HostManager_GetCurrentWorkbook()
    Set ws = HostManager_GetCurrentSheet()
    If wb Is Nothing Or ws Is Nothing Then Exit Function

    currentSheetName = ws.name
    Set dict = LoadActionsForSheet(currentSheetName)
    If dict Is Nothing Then Exit Function

    For Each k In dict.keys
        col.Add k
    Next k

    Set actionData = dict  ' refresh module-level cache
    Set GetActionList = col
    Exit Function

EH:
    Debug.Print "GetActionList error: " & Err.Description
End Function


Public Sub GetActionCount(control As IRibbonControl, ByRef count)
    On Error GoTo EH

    Dim col As Collection
    Set col = GetActionList()
    If col Is Nothing Then
        count = 0
    Else
        count = col.count
    End If
    Exit Sub

EH:
    Debug.Print "GetActionCount error: " & Err.Description
End Sub


Public Sub GetActionLabel(control As IRibbonControl, index As Integer, ByRef returnedVal)
    On Error GoTo EH

    Dim col As Collection
    Set col = GetActionList()
    If Not col Is Nothing Then
        If index >= 0 And index < col.count Then
            returnedVal = col(index + 1)
        End If
    End If
    Exit Sub

EH:
    Debug.Print "GetActionLabel error: " & Err.Description
End Sub




' ==============================================================
' ACTION COMBO BOX — Get/Change/Delete/Update
' ==============================================================

Public Sub GetActionText(control As IRibbonControl, ByRef returnedVal)
    On Error GoTo EH

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim sheetName As String
    Dim sel As String
    Dim dict As Object
    Dim k As Variant

    Debug.Print "=== GetActionText START ==="

    Set wb = HostManager_GetCurrentWorkbook()
    Set ws = HostManager_GetCurrentSheet()

    If wb Is Nothing Then
        Debug.Print "wb: Nothing"
    Else
        Debug.Print "wb: " & wb.name
    End If
    
    If ws Is Nothing Then
        Debug.Print "ws: Nothing"
    Else
        Debug.Print "ws: " & ws.name
    End If

    Debug.Print "currentSheetName (module var):", currentSheetName

    If wb Is Nothing Or ws Is Nothing Then
        Debug.Print "GetActionText: wb or ws Nothing"
        returnedVal = ""
        Exit Sub
    End If

    sheetName = ws.name
    Debug.Print "Using sheetName:", sheetName

    sel = GetSheetValue(wb, sheetName, "SelectedAction")
    Debug.Print "SelectedAction from sheet:", sel

    Set dict = LoadActionsForSheet(sheetName)

    If dict Is Nothing Then
        Debug.Print "LoadActionsForSheet returned Nothing"
        returnedVal = ""
        currentAction = ""
        Exit Sub
    End If

    Debug.Print "Loaded action count:", dict.count

    For Each k In dict.keys
        Debug.Print "Action key:", k
    Next k

    If Len(sel) = 0 Then
        Debug.Print "SelectedAction empty"
        returnedVal = ""
        currentAction = ""
        Exit Sub
    End If

    If Not dict.Exists(sel) Then
        Debug.Print "SelectedAction not in action dict"
        returnedVal = ""
        currentAction = ""
        Exit Sub
    End If

    currentAction = sel
    returnedVal = sel

    Debug.Print "=== GetActionText END ==="
    Exit Sub

EH:
    Debug.Print "GetActionText ERR:", Err.Description
    returnedVal = ""
End Sub



Public Sub OnActionChange(control As IRibbonControl, text As String)
    On Error GoTo EH

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim act As Object

    Set wb = HostManager_GetCurrentWorkbook()
    Set ws = HostManager_GetCurrentSheet()
    If wb Is Nothing Or ws Is Nothing Then
        MsgBox "No active workbook or sheet context.", vbExclamation
        Exit Sub
    End If

    currentSheetName = ws.name
    currentAction = text

    ' Persist selection
    SaveSheetValue wb, currentSheetName, "SelectedAction", text

    ' Ensure we have a valid actions dictionary
    If actionData Is Nothing Then Set actionData = LoadActionsForSheet(currentSheetName)
    If text <> "" And actionData.Exists(text) Then
        Set act = actionData(text)

        SaveSheetValue wb, currentSheetName, "cmbScript", act("script")
        SaveSheetValue wb, currentSheetName, "txtPyInput", act("input")
        SaveSheetValue wb, currentSheetName, "txtPyOutput", act("output")
        scriptSelected = act("script")
    End If

    ' Refresh dependent UI controls
    If Not rib Is Nothing Then
        rib.InvalidateControl "cmbScript"
        rib.InvalidateControl "txtPyInput"
        rib.InvalidateControl "txtPyOutput"
    End If
    
    Call RefreshRibbonValues
    
    Exit Sub

EH:
    Debug.Print "OnActionChange error: " & Err.Description
End Sub


Public Sub OnDeleteAction(control As IRibbonControl)
    On Error GoTo EH

    Dim wb As Workbook
    Dim ws As Worksheet

    Set wb = HostManager_GetCurrentWorkbook()
    Set ws = HostManager_GetCurrentSheet()
    If wb Is Nothing Or ws Is Nothing Then
        MsgBox "No active workbook or sheet context.", vbExclamation
        Exit Sub
    End If

    If actionData Is Nothing Then Set actionData = LoadActionsForSheet(ws.name)
    If Not actionData.Exists(currentAction) Then
        MsgBox "No action selected.", vbExclamation
        Exit Sub
    End If

    If MsgBox("Delete action '" & currentAction & "'?", vbYesNo + vbQuestion) = vbYes Then
        actionData.Remove currentAction
        SaveActionsForSheet ws.name, actionData

        currentAction = ""
        SaveSheetValue wb, ws.name, "SelectedAction", ""
        SaveSheetValue wb, ws.name, "cmbScript", ""
        SaveSheetValue wb, ws.name, "txtPyInput", ""
        SaveSheetValue wb, ws.name, "txtPyOutput", ""

        If Not rib Is Nothing Then
            rib.InvalidateControl "cmbActions"
            rib.InvalidateControl "cmbScript"
            rib.InvalidateControl "txtPyInput"
            rib.InvalidateControl "txtPyOutput"
        End If

        MsgBox "Action deleted.", vbInformation
    End If

    Exit Sub

EH:
    Debug.Print "OnDeleteAction error: " & Err.Description
End Sub


Private Sub UpdateCurrentActionField(fieldName As String, fieldValue As String)
    On Error GoTo EH

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim act As Object

    Set wb = HostManager_GetCurrentWorkbook()
    Set ws = HostManager_GetCurrentSheet()
    If wb Is Nothing Or ws Is Nothing Then Exit Sub

    If actionData Is Nothing Then Set actionData = LoadActionsForSheet(ws.name)
    If Len(currentAction) = 0 Then Exit Sub
    If Not actionData.Exists(currentAction) Then Exit Sub

    Set act = actionData(currentAction)
    act(fieldName) = fieldValue

    SaveActionsForSheet ws.name, actionData
    Exit Sub

EH:
    Debug.Print "UpdateCurrentActionField error: " & Err.Description
End Sub




Private Sub SaveRibbonState()
    Dim wb As Workbook
    Set wb = HostManager_GetCurrentWorkbook()
    If wb Is Nothing Then Exit Sub

    ' CHANGE: Use SaveWorkbookValue instead of SaveSheetValue
    SaveWorkbookValue wb, "PyExcelEnabled", IIf(RibbonIsEnabled, "1", "0")
End Sub

Public Sub RibbonEnabled(control As IRibbonControl, ByRef returnedVal)
    Dim wb As Workbook
    Dim val As String

    ' 1. Get Context
    Set wb = HostManager_GetCurrentWorkbook()
    If Not wb Is Nothing Then
         ' CHANGE: Read from Workbook name, not Sheet name
         On Error Resume Next
         val = GetWorkbookValue(wb, "PyExcelEnabled")
         On Error GoTo 0

         ' 2. Update the global state
         RibbonIsEnabled = (val = "1")
    End If

    ' 3. Return the fresh value
    returnedVal = RibbonIsEnabled
End Sub


Public Function GetWorkbookValue(wb As Workbook, ctrlId As String) As String
    On Error GoTo EH

    Dim nm As name
    Dim raw As String

    If wb Is Nothing Then Exit Function
    
    ' Try to get the named range in the workbook scope
    On Error Resume Next
    Set nm = wb.Names(ctrlId)
    On Error GoTo EH

    If nm Is Nothing Then Exit Function ' Name not found in workbook scope

    raw = nm.RefersTo
    
    ' Normalize the string result
    If Left$(raw, 1) = "=" Then raw = Mid$(raw, 2)
    If Left$(raw, 1) = """" And right$(raw, 1) = """" Then
        raw = Mid$(raw, 2, Len(raw) - 2)
    End If

    GetWorkbookValue = raw
    Exit Function

EH:
    Debug.Print "[GetWorkbookValue][ERROR] ctrl='" & ctrlId & "' err=" & Err.Description
End Function


Public Sub SaveWorkbookValue(wb As Workbook, ctrlId As String, val As Variant)
    On Error GoTo EH

    If wb Is Nothing Then Exit Sub
    
    ' Remove any old name
    On Error Resume Next
    wb.Names(ctrlId).Delete
    On Error GoTo EH

    ' Add the new name if value is not empty
    If Len(val) > 0 Then
        ' Workbook level name creation
        wb.Names.Add name:=ctrlId, RefersTo:="=""" & Replace(val, """", """""") & """"
    End If

    Exit Sub

EH:
    Debug.Print "[SaveWorkbookValue][ERROR] ctrl='" & ctrlId & "' err=" & Err.Description
End Sub


