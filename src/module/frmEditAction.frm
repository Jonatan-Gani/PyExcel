VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEditAction 
   Caption         =   "Manage Action"
   ClientHeight    =   5528
   ClientLeft      =   96
   ClientTop       =   416
   ClientWidth     =   7088
   OleObjectBlob   =   "frmEditAction.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmEditAction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private hasLoaded As Boolean
Private currentSheetName As String
Private currentAction As String
Private actionData As Object
Private scriptSelected As String

Private Sub UserForm_Initialize()
    On Error GoTo EH

    ' <<< FIX: initialize ONLY once
    If Not hasLoaded Then
        RefreshFromContext
        hasLoaded = True
    End If

    Exit Sub
EH:
    Debug.Print "[frmEditAction.Initialize][ERROR] " & Err.Description
End Sub

Private Sub UserForm_Activate()
    ' intentionally empty — do NOT refresh here
End Sub

Private Sub RefreshFromContext()
    On Error GoTo EH

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim act As String
    Dim files As Collection
    Dim f As Variant
    Dim inputVals As Variant
    Dim outputVals As Variant
    Dim i As Long
    Dim aliasName As String, rangeAddr As String, sheetName As String, itemType As String

    Set wb = HostManager_GetCurrentWorkbook()
    Set ws = HostManager_GetCurrentSheet()

    If wb Is Nothing Or ws Is Nothing Then
        MsgBox "No active workbook or sheet context.", vbExclamation
        Unload Me
        Exit Sub
    End If

    currentSheetName = ws.name

    act = GetSheetValue(wb, currentSheetName, "SelectedAction")
    boxName.value = act
    currentAction = act

    ComboBoxScript.Clear
    Set files = GetScriptFiles()
    If Not files Is Nothing Then
        For Each f In files
            ComboBoxScript.AddItem f
        Next
    End If

    ComboBoxScript.value = GetSheetValue(wb, currentSheetName, "cmbScript")

    ' Configure 4-column listboxes and matching headers: Name | Range | Sheet | Type
    Dim colWidths As String: colWidths = "60;80;65;40"

    ListBoxInputHeader.ColumnCount = 4
    ListBoxInputHeader.ColumnWidths = colWidths
    ListBoxInputHeader.Clear
    ListBoxInputHeader.AddItem "Name"
    ListBoxInputHeader.List(0, 1) = "Range"
    ListBoxInputHeader.List(0, 2) = "Sheet"
    ListBoxInputHeader.List(0, 3) = "Type"

    ListBoxOutputHeader.ColumnCount = 4
    ListBoxOutputHeader.ColumnWidths = colWidths
    ListBoxOutputHeader.Clear
    ListBoxOutputHeader.AddItem "Name"
    ListBoxOutputHeader.List(0, 1) = "Range"
    ListBoxOutputHeader.List(0, 2) = "Sheet"
    ListBoxOutputHeader.List(0, 3) = "Type"

    ListBoxInput.ColumnCount = 4
    ListBoxInput.ColumnWidths = colWidths
    ListBoxOutput.ColumnCount = 4
    ListBoxOutput.ColumnWidths = colWidths

    ListBoxInput.Clear
    inputVals = Split(GetSheetValue(wb, currentSheetName, "txtPyInput"), ";")
    For i = LBound(inputVals) To UBound(inputVals)
        If Trim$(inputVals(i)) <> "" Then
            ParseRefToColumns Trim$(inputVals(i)), aliasName, rangeAddr, sheetName, itemType
            AddListBoxRow ListBoxInput, aliasName, rangeAddr, sheetName, itemType
        End If
    Next i

    ListBoxOutput.Clear
    outputVals = Split(GetSheetValue(wb, currentSheetName, "txtPyOutput"), ";")
    For i = LBound(outputVals) To UBound(outputVals)
        If Trim$(outputVals(i)) <> "" Then
            ParseRefToColumns Trim$(outputVals(i)), aliasName, rangeAddr, sheetName, itemType
            AddListBoxRow ListBoxOutput, aliasName, rangeAddr, sheetName, itemType
        End If
    Next i

    ' Load EntreBox state from action dictionary
    Set actionData = LoadActionsForSheet(currentSheetName)
    Dim eteVal As String: eteVal = "False"
    If Not actionData Is Nothing And Len(currentAction) > 0 Then
        If actionData.Exists(currentAction) Then
            On Error Resume Next
            eteVal = actionData(currentAction)("entreToEnd")
            On Error GoTo 0
        End If
    End If
    EntreBox.value = (LCase$(Trim$(eteVal)) = "true")

    ' Populate import source combo
    cmbImportSource.Clear
    cmbImportSource.AddItem "From File"
    cmbImportSource.AddItem "From Sheet"
    cmbImportSource.AddItem "From Workbook"
    cmbImportSource.ListIndex = 0

    Debug.Print "[frmEditAction] Refreshed for sheet: " & currentSheetName & ", action='" & currentAction & "'"
    Exit Sub

EH:
    Debug.Print "[frmEditAction.RefreshFromContext][ERROR] " & Err.Description
End Sub



Private Sub btnSave_Click()
    On Error GoTo EH

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim oldAction As String
    Dim newAction As String
    Dim scriptVal As String
    Dim inputVal As String
    Dim outputVal As String
    Dim act As Object
    Dim i As Long
    Dim tempInput As String
    Dim tempOutput As String

    Set wb = HostManager_GetCurrentWorkbook()
    Set ws = HostManager_GetCurrentSheet()
    
    currentSheetName = ws.name
    
    oldAction = currentAction
    newAction = Trim$(boxName.value)
    scriptVal = ComboBoxScript.value

    tempInput = ""
    For i = 0 To ListBoxInput.listCount - 1
        If i > 0 Then tempInput = tempInput & ";"
        tempInput = tempInput & BuildRefFromCols(ListBoxInput.List(i, 0), ListBoxInput.List(i, 2), ListBoxInput.List(i, 1))
    Next i

    tempOutput = ""
    For i = 0 To ListBoxOutput.listCount - 1
        If i > 0 Then tempOutput = tempOutput & ";"
        tempOutput = tempOutput & BuildRefFromCols(ListBoxOutput.List(i, 0), ListBoxOutput.List(i, 2), ListBoxOutput.List(i, 1))
    Next i

    inputVal = tempInput
    outputVal = tempOutput

    If Len(newAction) = 0 Then
        MsgBox "Action name cannot be empty.", vbExclamation
        Exit Sub
    End If

    If actionData Is Nothing Then Set actionData = LoadActionsForSheet(currentSheetName)
    If TypeName(actionData) <> "Dictionary" Then Set actionData = CreateObject("Scripting.Dictionary")

    If oldAction <> newAction And actionData.Exists(newAction) Then
        MsgBox "An action named '" & newAction & "' already exists.", vbExclamation
        Exit Sub
    End If

    If Len(oldAction) > 0 And oldAction <> newAction Then
        If actionData.Exists(oldAction) Then actionData.Remove oldAction
    End If

    If actionData.Exists(newAction) Then
        Set act = actionData(newAction)
    Else
        Set act = CreateObject("Scripting.Dictionary")
    End If

    act("script") = scriptVal
    act("input") = inputVal
    act("output") = outputVal
    act("entireRow") = IIf(act.Exists("entireRow"), act("entireRow"), "False")
    act("entreToEnd") = IIf(EntreBox.value, "True", "False")
    Set actionData(newAction) = act

    SaveActionsForSheet currentSheetName, actionData
    SaveSheetValue wb, currentSheetName, "SelectedAction", newAction
    currentAction = newAction



    SaveSheetValue wb, currentSheetName, "cmbScript", act("script")
    SaveSheetValue wb, currentSheetName, "txtPyInput", act("input")
    SaveSheetValue wb, currentSheetName, "txtPyOutput", act("output")
    SaveSheetValue wb, currentSheetName, "chkEntireRow", act("entireRow")
    scriptSelected = act("script")

    If Not rib Is Nothing Then
        rib.InvalidateControl "cmbActions"
        rib.InvalidateControl "cmbScript"
        rib.InvalidateControl "txtPyInput"
        rib.InvalidateControl "txtPyOutput"
    End If
    
    Call RefreshRibbonValues

    Unload Me
    Exit Sub

EH:
    Debug.Print "[frmEditAction.btnSave_Click][ERROR] " & Err.Description
    MsgBox "Error saving action: " & Err.Description, vbCritical
End Sub



Private Sub btnDiscard_Click()
    Unload Me
End Sub



Private Sub btnAddInput_Click()
    Dim rng As Range
    On Error Resume Next
    Set rng = Application.InputBox("Select Input Range", Type:=8)
    On Error GoTo 0
    If rng Is Nothing Then Exit Sub
    AddListBoxRow ListBoxInput, "", rng.Address(False, False), rng.parent.name, "Range"
End Sub

Private Sub btnAddOutput_Click()
    Dim rng As Range
    On Error Resume Next
    Set rng = Application.InputBox("Select Output Range", Type:=8)
    On Error GoTo 0
    If rng Is Nothing Then Exit Sub
    AddListBoxRow ListBoxOutput, "", rng.Address(False, False), rng.parent.name, "Range"
End Sub



Private Sub btnEditInput_Click()
    Dim i As Long
    Dim rawRef As String
    Dim updated As String
    Dim savedLeft As Single
    Dim savedTop As Single
    Dim aliasName As String, rangeAddr As String, sheetName As String, itemType As String

    i = ListBoxInput.ListIndex
    If i < 0 Then
        MsgBox "Select an input item to edit.", vbExclamation
        Exit Sub
    End If

    rawRef = BuildRefFromCols(ListBoxInput.List(i, 0), ListBoxInput.List(i, 2), ListBoxInput.List(i, 1))

    savedLeft = Me.Left
    savedTop = Me.Top
    Me.Left = -20000
    Me.Top = -20000

    updated = frmRangeSet.GetUpdatedValue(rawRef)

    Me.Left = savedLeft
    Me.Top = savedTop

    ParseRefToColumns updated, aliasName, rangeAddr, sheetName, itemType
    ListBoxInput.List(i, 0) = aliasName
    ListBoxInput.List(i, 1) = rangeAddr
    ListBoxInput.List(i, 2) = sheetName
    ListBoxInput.List(i, 3) = itemType
End Sub



Private Sub btnEditOutput_Click()
    Dim i As Long
    Dim rawRef As String
    Dim updated As String
    Dim savedLeft As Single
    Dim savedTop As Single
    Dim aliasName As String, rangeAddr As String, sheetName As String, itemType As String

    i = ListBoxOutput.ListIndex
    If i < 0 Then
        MsgBox "Select an output item to edit.", vbExclamation
        Exit Sub
    End If

    rawRef = BuildRefFromCols(ListBoxOutput.List(i, 0), ListBoxOutput.List(i, 2), ListBoxOutput.List(i, 1))

    savedLeft = Me.Left
    savedTop = Me.Top


    Me.Left = -20000
    Me.Top = -20000

    updated = frmRangeSet.GetUpdatedValue(rawRef)





    Me.Left = savedLeft
    Me.Top = savedTop

    ParseRefToColumns updated, aliasName, rangeAddr, sheetName, itemType
    ListBoxOutput.List(i, 0) = aliasName
    ListBoxOutput.List(i, 1) = rangeAddr
    ListBoxOutput.List(i, 2) = sheetName
    ListBoxOutput.List(i, 3) = itemType



End Sub










Private Sub btnDeleteInput_Click()
    Dim selIndex As Long
    selIndex = ListBoxInput.ListIndex
    If selIndex < 0 Then
        MsgBox "Select an input item to delete.", vbExclamation
        Exit Sub
    End If
    ListBoxInput.RemoveItem selIndex
End Sub


Private Sub btnDeleteOutput_Click()
    Dim selIndex As Long
    selIndex = ListBoxOutput.ListIndex
    If selIndex < 0 Then

        MsgBox "Select an output item to delete.", vbExclamation
        Exit Sub
    End If
    ListBoxOutput.RemoveItem selIndex

End Sub



Private Sub btnMoveUpInput_Click()
    MoveListBoxItem ListBoxInput, -1
End Sub

Private Sub btnMoveDownInput_Click()
    MoveListBoxItem ListBoxInput, 1
End Sub

Private Sub btnMoveUpOutput_Click()
    MoveListBoxItem ListBoxOutput, -1
End Sub

Private Sub btnMoveDownOutput_Click()
    MoveListBoxItem ListBoxOutput, 1
End Sub

Private Sub MoveListBoxItem(lb As MSForms.ListBox, direction As Integer)
    Dim i As Long, swapIdx As Long, c As Long
    Dim temp As String

    i = lb.ListIndex
    If i < 0 Then Exit Sub

    swapIdx = i + direction
    If swapIdx < 0 Or swapIdx >= lb.listCount Then Exit Sub

    For c = 0 To lb.ColumnCount - 1
        temp = lb.List(i, c)
        lb.List(i, c) = lb.List(swapIdx, c)
        lb.List(swapIdx, c) = temp
    Next c
    lb.ListIndex = swapIdx
End Sub



Private Sub ParseRefToColumns(rawRef As String, aliasName As String, rangeAddr As String, sheetName As String, itemType As String)
    Dim eqPos As Long
    Dim refPart As String

    itemType = "Range"
    aliasName = ""
    rangeAddr = ""
    sheetName = ""

    rawRef = Trim$(rawRef)
    If Len(rawRef) = 0 Then Exit Sub

    eqPos = InStr(rawRef, "=")
    If eqPos > 0 Then
        aliasName = Trim$(Left$(rawRef, eqPos - 1))
        refPart = Trim$(Mid$(rawRef, eqPos + 1))
    Else
        refPart = rawRef
    End If

    ParseRangeRef refPart, sheetName, rangeAddr
End Sub


Private Sub AddListBoxRow(lb As MSForms.ListBox, aliasName As String, rangeAddr As String, sheetName As String, itemType As String)
    lb.AddItem aliasName
    Dim r As Long: r = lb.listCount - 1
    lb.List(r, 1) = rangeAddr
    lb.List(r, 2) = sheetName
    lb.List(r, 3) = itemType
End Sub


Private Function BuildRefFromCols(aliasName As String, sheetName As String, rangeAddr As String) As String
    Dim ref As String
    If Len(sheetName) > 0 Then
        ref = BuildRangeRef(sheetName, rangeAddr)
    Else
        ref = rangeAddr
    End If
    If Len(aliasName) > 0 Then
        BuildRefFromCols = aliasName & "=" & ref
    Else
        BuildRefFromCols = ref
    End If
End Function


' ==============================================================
' IMPORT / EXPORT
' ==============================================================

Private Sub btnExport_Click()
    On Error GoTo EH

    Dim dict As Object
    Dim filePath As String
    Dim ext As String
    Dim content As String

    Set dict = LoadActionsForSheet(currentSheetName)
    If dict Is Nothing Or dict.count = 0 Then
        MsgBox "No actions on this sheet to export.", vbExclamation
        Exit Sub
    End If

    filePath = Application.GetSaveAsFilename( _
        InitialFileName:=currentSheetName & "_actions", _
        FileFilter:="JSON Files (*.json),*.json," & _
                     "XML Files (*.xml),*.xml," & _
                     "Text Files (*.txt),*.txt", _
        Title:="Export Actions")

    If filePath = "False" Or Len(filePath) = 0 Then Exit Sub

    ext = LCase(Mid(filePath, InStrRev(filePath, ".") + 1))

    Select Case ext
        Case "json": content = ActionsToJSON(dict)
        Case "xml":  content = ActionsToXML(dict)
        Case "txt":  content = ActionsToTXT(dict)
        Case Else
            MsgBox "Unsupported file type. Use .json, .xml, or .txt", vbExclamation
            Exit Sub
    End Select

    IOWriteTextFile filePath, content
    MsgBox dict.count & " action(s) exported.", vbInformation
    Exit Sub

EH:
    MsgBox "Export failed: " & Err.Description, vbCritical
End Sub


Private Sub btnImport_Click()
    On Error GoTo EH

    Dim imported As Object

    If cmbImportSource.ListIndex < 0 Then
        MsgBox "Select an import source first.", vbExclamation
        Exit Sub
    End If

    Select Case cmbImportSource.ListIndex
        Case 0: Set imported = DoImportFromFile()
        Case 1: Set imported = DoImportFromSheet()
        Case 2: Set imported = DoImportFromWorkbook()
        Case Else: Exit Sub
    End Select

    If imported Is Nothing Then Exit Sub
    If imported.count = 0 Then
        MsgBox "No valid actions found.", vbExclamation
        Exit Sub
    End If

    MergeImportedActions imported
    Set actionData = LoadActionsForSheet(currentSheetName)
    RefreshFromContext
    MsgBox imported.count & " action(s) imported.", vbInformation
    Exit Sub

EH:
    MsgBox "Import failed: " & Err.Description, vbCritical
End Sub


' ---------------------------------------------------------------
' IMPORT SOURCES
' ---------------------------------------------------------------

Private Function DoImportFromFile() As Object
    On Error GoTo EH

    Dim filePath As Variant
    Dim content As String
    Dim ext As String

    Set DoImportFromFile = Nothing

    filePath = Application.GetOpenFilename( _
        FileFilter:="Action Files (*.json;*.xml;*.txt),*.json;*.xml;*.txt," & _
                     "All Files (*.*),*.*", _
        Title:="Import Actions from File")

    If filePath = False Then Exit Function

    content = IOReadTextFile(CStr(filePath))
    If Len(content) = 0 Then
        MsgBox "File is empty or could not be read.", vbExclamation
        Exit Function
    End If

    ext = LCase(Mid(CStr(filePath), InStrRev(CStr(filePath), ".") + 1))

    Select Case ext
        Case "json": Set DoImportFromFile = JSONToActions(content)
        Case "xml":  Set DoImportFromFile = XMLToActions(content)
        Case "txt":  Set DoImportFromFile = TXTToActions(content)
        Case Else:   MsgBox "Unsupported file type: " & ext, vbExclamation
    End Select
    Exit Function

EH:
    Debug.Print "DoImportFromFile error: " & Err.Description
End Function


Private Function DoImportFromSheet() As Object
    On Error GoTo EH

    Dim sourceSheet As String

    Set DoImportFromSheet = Nothing

    SheetPickerForm.Show vbModal
    sourceSheet = SheetPickerForm.SelectedSheet
    Unload SheetPickerForm
    If Len(sourceSheet) = 0 Then Exit Function

    If sourceSheet = currentSheetName Then
        MsgBox "Cannot import from the current sheet.", vbExclamation
        Exit Function
    End If

    Set DoImportFromSheet = LoadActionsForSheet(sourceSheet)
    Exit Function

EH:
    Debug.Print "DoImportFromSheet error: " & Err.Description
End Function


Private Function DoImportFromWorkbook() As Object
    On Error GoTo EH

    Dim filePath As Variant
    Dim srcWb As Workbook
    Dim srcWs As Worksheet
    Dim sheetNames As Collection
    Dim sName As Variant
    Dim raw As String
    Dim chosenSheet As String

    Set DoImportFromWorkbook = Nothing

    filePath = Application.GetOpenFilename( _
        FileFilter:="Excel Files (*.xlsx;*.xlsm;*.xlsb;*.xls),*.xlsx;*.xlsm;*.xlsb;*.xls", _
        Title:="Select Workbook")

    If filePath = False Then Exit Function

    Application.ScreenUpdating = False
    Set srcWb = Workbooks.Open(CStr(filePath), ReadOnly:=True, UpdateLinks:=0)

    Set sheetNames = New Collection
    For Each srcWs In srcWb.Worksheets
        raw = ReadNamedRangeFromWb(srcWb, srcWs.name, "Actions")
        If Len(raw) > 0 Then sheetNames.Add srcWs.name
    Next srcWs

    If sheetNames.count = 0 Then
        srcWb.Close SaveChanges:=False
        Application.ScreenUpdating = True
        MsgBox "No actions found in the selected workbook.", vbExclamation
        Exit Function
    End If

    If sheetNames.count = 1 Then
        chosenSheet = sheetNames(1)
    Else
        Dim msg As String
        msg = "Sheets with actions:" & vbCrLf & vbCrLf
        For Each sName In sheetNames
            msg = msg & "  - " & sName & vbCrLf
        Next
        srcWb.Close SaveChanges:=False
        Application.ScreenUpdating = True

        chosenSheet = InputBox(msg & vbCrLf & "Enter sheet name:", "Pick Sheet")
        If Len(chosenSheet) = 0 Then Exit Function

        Dim found As Boolean: found = False
        For Each sName In sheetNames
            If LCase(CStr(sName)) = LCase(chosenSheet) Then
                chosenSheet = CStr(sName): found = True: Exit For
            End If
        Next
        If Not found Then
            MsgBox "Sheet '" & chosenSheet & "' not found.", vbExclamation
            Exit Function
        End If

        Application.ScreenUpdating = False
        Set srcWb = Workbooks.Open(CStr(filePath), ReadOnly:=True, UpdateLinks:=0)
    End If

    raw = ReadNamedRangeFromWb(srcWb, chosenSheet, "Actions")
    srcWb.Close SaveChanges:=False
    Application.ScreenUpdating = True

    If Len(raw) > 0 Then Set DoImportFromWorkbook = ParseRawActionsString(raw)
    Exit Function

EH:
    On Error Resume Next
    If Not srcWb Is Nothing Then srcWb.Close SaveChanges:=False
    Application.ScreenUpdating = True
    On Error GoTo 0
    Debug.Print "DoImportFromWorkbook error: " & Err.Description
End Function


' ---------------------------------------------------------------
' MERGE IMPORTED ACTIONS
' ---------------------------------------------------------------

Private Sub MergeImportedActions(imported As Object)
    Dim existing As Object
    Dim k As Variant
    Dim conflicts As String
    Dim conflictCount As Long

    Set existing = LoadActionsForSheet(currentSheetName)
    If existing Is Nothing Then Set existing = CreateObject("Scripting.Dictionary")

    conflicts = ""
    conflictCount = 0
    For Each k In imported.keys
        If existing.Exists(k) Then
            conflictCount = conflictCount + 1
            conflicts = conflicts & "  - " & k & vbCrLf
        End If
    Next k

    If conflictCount > 0 Then
        Dim answer As VbMsgBoxResult
        answer = MsgBox(conflictCount & " action(s) already exist:" & vbCrLf & _
                        conflicts & vbCrLf & _
                        "Yes = Overwrite" & vbCrLf & _
                        "No = Skip duplicates" & vbCrLf & _
                        "Cancel = Abort", _
                        vbYesNoCancel + vbQuestion, "Duplicates")
        If answer = vbCancel Then Exit Sub

        For Each k In imported.keys
            If existing.Exists(k) Then
                If answer = vbYes Then Set existing(k) = imported(k)
            Else
                Set existing(k) = imported(k)
            End If
        Next k
    Else
        For Each k In imported.keys
            Set existing(k) = imported(k)
        Next k
    End If

    SaveActionsForSheet currentSheetName, existing

    If Not rib Is Nothing Then
        rib.InvalidateControl "cmbActions"
        rib.InvalidateControl "cmbScript"
        rib.InvalidateControl "txtPyInput"
        rib.InvalidateControl "txtPyOutput"
    End If
End Sub


' ---------------------------------------------------------------
' SERIALIZATION — JSON
' ---------------------------------------------------------------

Private Function ActionsToJSON(dict As Object) As String
    Dim s As String, k As Variant, act As Object, first As Boolean

    s = "{" & vbCrLf
    first = True
    For Each k In dict.keys
        Set act = dict(k)
        If Not first Then s = s & "," & vbCrLf
        first = False
        s = s & "  " & JEsc(CStr(k)) & ": {" & vbCrLf
        s = s & "    ""script"": " & JEsc(act("script")) & "," & vbCrLf
        s = s & "    ""input"": " & JEsc(act("input")) & "," & vbCrLf
        s = s & "    ""output"": " & JEsc(act("output")) & "," & vbCrLf
        s = s & "    ""entireRow"": " & JEsc(act("entireRow")) & "," & vbCrLf
        s = s & "    ""entreToEnd"": " & JEsc(act("entreToEnd")) & vbCrLf
        s = s & "  }"
    Next k
    s = s & vbCrLf & "}"
    ActionsToJSON = s
End Function

Private Function JEsc(val As String) As String
    Dim s As String
    s = Replace(val, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, vbCrLf, "\n")
    s = Replace(s, vbCr, "\n")
    s = Replace(s, vbLf, "\n")
    JEsc = """" & s & """"
End Function

Private Function JSONToActions(content As String) As Object
    On Error GoTo EH

    Dim dict As Object, pos As Long
    Dim actionName As String, act As Object
    Dim fName As String, fVal As String, nc As String

    Set dict = CreateObject("Scripting.Dictionary")
    pos = InStr(1, content, "{") + 1

    Do While pos > 1 And pos <= Len(content)
        actionName = ReadJStr(content, pos)
        If Len(actionName) = 0 Then Exit Do
        pos = InStr(pos, content, ":") + 1: If pos <= 1 Then Exit Do

        Set act = CreateObject("Scripting.Dictionary")
        act("script") = "": act("input") = "": act("output") = ""
        act("entireRow") = "False": act("entreToEnd") = "False"

        pos = InStr(pos, content, "{") + 1: If pos <= 1 Then Exit Do
        Do
            fName = ReadJStr(content, pos): If Len(fName) = 0 Then Exit Do
            pos = InStr(pos, content, ":") + 1: If pos <= 1 Then Exit Do
            fVal = ReadJStr(content, pos)
            act(fName) = fVal
            nc = SkipWS(content, pos): If nc = "," Then pos = pos + 1
        Loop
        pos = InStr(pos, content, "}") + 1
        dict.Add actionName, act
        If pos > 0 And pos <= Len(content) Then
            nc = SkipWS(content, pos): If nc = "," Then pos = pos + 1
        End If
    Loop

    Set JSONToActions = dict
    Exit Function
EH:
    Set JSONToActions = Nothing
End Function

Private Function ReadJStr(content As String, ByRef pos As Long) As String
    Dim sq As Long, eq As Long, raw As String
    sq = InStr(pos, content, """")
    If sq = 0 Then ReadJStr = "": Exit Function
    eq = sq + 1
    Do While eq <= Len(content)
        If Mid(content, eq, 1) = """" And Mid(content, eq - 1, 1) <> "\" Then Exit Do
        eq = eq + 1
    Loop
    raw = Mid(content, sq + 1, eq - sq - 1)
    raw = Replace(raw, "\""", """")
    raw = Replace(raw, "\\", "\")
    raw = Replace(raw, "\n", vbLf)
    ReadJStr = raw
    pos = eq + 1
End Function

Private Function SkipWS(content As String, ByRef pos As Long) As String
    Dim c As String
    Do While pos <= Len(content)
        c = Mid(content, pos, 1)
        If c <> " " And c <> vbCr And c <> vbLf And c <> vbTab Then
            SkipWS = c: Exit Function
        End If
        pos = pos + 1
    Loop
    SkipWS = ""
End Function


' ---------------------------------------------------------------
' SERIALIZATION — XML
' ---------------------------------------------------------------

Private Function ActionsToXML(dict As Object) As String
    Dim s As String, k As Variant, act As Object
    s = "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf & "<Actions>" & vbCrLf
    For Each k In dict.keys
        Set act = dict(k)
        s = s & "  <Action name=""" & XEsc(CStr(k)) & """>" & vbCrLf
        s = s & "    <script>" & XEsc(act("script")) & "</script>" & vbCrLf
        s = s & "    <input>" & XEsc(act("input")) & "</input>" & vbCrLf
        s = s & "    <output>" & XEsc(act("output")) & "</output>" & vbCrLf
        s = s & "    <entireRow>" & XEsc(act("entireRow")) & "</entireRow>" & vbCrLf
        s = s & "    <entreToEnd>" & XEsc(act("entreToEnd")) & "</entreToEnd>" & vbCrLf
        s = s & "  </Action>" & vbCrLf
    Next k
    s = s & "</Actions>"
    ActionsToXML = s
End Function

Private Function XEsc(val As String) As String
    Dim s As String
    s = Replace(val, "&", "&amp;")
    s = Replace(s, "<", "&lt;")
    s = Replace(s, ">", "&gt;")
    s = Replace(s, """", "&quot;")
    XEsc = s
End Function

Private Function XMLToActions(content As String) As Object
    On Error GoTo EH
    Dim dict As Object, doc As Object, nodes As Object, node As Object
    Dim act As Object, aName As String, i As Long

    Set dict = CreateObject("Scripting.Dictionary")
    Set doc = CreateObject("MSXML2.DOMDocument.6.0")
    doc.async = False
    doc.LoadXML content
    If doc.parseError.ErrorCode <> 0 Then Set XMLToActions = Nothing: Exit Function

    Set nodes = doc.SelectNodes("//Action")
    For i = 0 To nodes.Length - 1
        Set node = nodes.Item(i)
        aName = node.getAttribute("name")
        If Len(aName) > 0 Then
            Set act = CreateObject("Scripting.Dictionary")
            act("script") = XChild(node, "script")
            act("input") = XChild(node, "input")
            act("output") = XChild(node, "output")
            act("entireRow") = XChild(node, "entireRow"): If Len(act("entireRow")) = 0 Then act("entireRow") = "False"
            act("entreToEnd") = XChild(node, "entreToEnd"): If Len(act("entreToEnd")) = 0 Then act("entreToEnd") = "False"
            dict.Add aName, act
        End If
    Next i
    Set XMLToActions = dict
    Exit Function
EH:
    Set XMLToActions = Nothing
End Function

Private Function XChild(node As Object, childName As String) As String
    On Error Resume Next
    Dim child As Object
    Set child = node.SelectSingleNode(childName)
    If Not child Is Nothing Then XChild = child.text Else XChild = ""
    On Error GoTo 0
End Function


' ---------------------------------------------------------------
' SERIALIZATION — TXT (pipe-delimited)
' ---------------------------------------------------------------

Private Function ActionsToTXT(dict As Object) As String
    Dim s As String, k As Variant, act As Object
    Dim line As String, inP As Variant, outP As Variant, p As Long

    s = "# PyExcel Action Set Export" & vbCrLf
    s = s & "# Format: ActionName|script=...|input=...|output=...|entireRow=...|entreToEnd=..." & vbCrLf

    For Each k In dict.keys
        Set act = dict(k)
        line = CStr(k) & "|script=" & act("script")
        inP = Split(act("input"), "; ")
        For p = LBound(inP) To UBound(inP)
            If Len(Trim$(inP(p))) > 0 Then line = line & "|input=" & Trim$(inP(p))
        Next p
        outP = Split(act("output"), "; ")
        For p = LBound(outP) To UBound(outP)
            If Len(Trim$(outP(p))) > 0 Then line = line & "|output=" & Trim$(outP(p))
        Next p
        line = line & "|entireRow=" & act("entireRow") & "|entreToEnd=" & act("entreToEnd")
        s = s & line & vbCrLf
    Next k
    ActionsToTXT = s
End Function

Private Function TXTToActions(content As String) As Object
    On Error GoTo EH
    Dim dict As Object, lines As Variant, i As Long
    Dim line As String, cols As Variant, k As String, act As Object
    Dim j As Long, kv() As String, fKey As String, fVal As String

    Set dict = CreateObject("Scripting.Dictionary")
    lines = Split(content, vbCrLf)
    If UBound(lines) < 0 Then lines = Split(content, vbLf)

    For i = LBound(lines) To UBound(lines)
        line = Trim$(lines(i))
        If Len(line) = 0 Or Left$(line, 1) = "#" Then GoTo NextTxtLine
        cols = Split(line, "|")
        If UBound(cols) < 1 Then GoTo NextTxtLine
        k = Trim$(cols(0)): If Len(k) = 0 Then GoTo NextTxtLine

        Set act = CreateObject("Scripting.Dictionary")
        act("script") = "": act("input") = "": act("output") = ""
        act("entireRow") = "False": act("entreToEnd") = "False"
        For j = 1 To UBound(cols)
            kv = Split(cols(j), "=", 2)
            If UBound(kv) >= 1 Then
                fKey = Trim$(kv(0)): fVal = Trim$(kv(1))
                If fKey = "input" Or fKey = "output" Then
                    If Len(act(fKey)) = 0 Then act(fKey) = fVal Else act(fKey) = act(fKey) & "; " & fVal
                Else
                    act(fKey) = fVal
                End If
            End If
        Next j
        dict.Add k, act
NextTxtLine:
    Next i
    Set TXTToActions = dict
    Exit Function
EH:
    Set TXTToActions = Nothing
End Function


' ---------------------------------------------------------------
' PARSE RAW ACTIONS STRING (for workbook import)
' ---------------------------------------------------------------

Private Function ParseRawActionsString(raw As String) As Object
    On Error GoTo EH
    Dim dict As Object, rows As Variant, cols As Variant, i As Long
    Dim act As Object, k As String, rowSep As String
    Dim j As Long, kv() As String, fKey As String, fVal As String

    Set dict = CreateObject("Scripting.Dictionary")
    If Len(raw) = 0 Then Set ParseRawActionsString = dict: Exit Function

    If InStr(raw, Chr(29)) > 0 Then rowSep = Chr(29) _
    ElseIf InStr(raw, Chr(10)) > 0 Then rowSep = Chr(10) _
    Else: rowSep = ";"

    rows = Split(raw, rowSep)
    For i = LBound(rows) To UBound(rows)
        If Len(Trim$(rows(i))) > 0 Then
            cols = Split(rows(i), "|")
            If UBound(cols) >= 1 Then
                k = Trim$(cols(0))
                If Len(k) > 0 Then
                    Set act = CreateObject("Scripting.Dictionary")
                    act("script") = "": act("input") = "": act("output") = ""
                    act("entireRow") = "False": act("entreToEnd") = "False"
                    If InStr(cols(1), "=") > 0 Then
                        For j = 1 To UBound(cols)
                            kv = Split(cols(j), "=", 2)
                            If UBound(kv) >= 1 Then
                                fKey = Trim$(kv(0)): fVal = Trim$(kv(1))
                                If fKey = "input" Or fKey = "output" Then
                                    If Len(act(fKey)) = 0 Then act(fKey) = fVal Else act(fKey) = act(fKey) & "; " & fVal
                                Else: act(fKey) = fVal
                                End If
                            End If
                        Next j
                    ElseIf UBound(cols) >= 3 Then
                        act("script") = Trim$(cols(1)): act("input") = Trim$(cols(2)): act("output") = Trim$(cols(3))
                        act("entireRow") = IIf(UBound(cols) >= 4, Trim$(cols(4)), "False")
                        act("entreToEnd") = IIf(UBound(cols) >= 5, Trim$(cols(5)), "False")
                    End If
                    dict.Add k, act
                End If
            End If
        End If
    Next i
    Set ParseRawActionsString = dict
    Exit Function
EH:
    Set ParseRawActionsString = Nothing
End Function


' ---------------------------------------------------------------
' READ NAMED RANGE FROM EXTERNAL WORKBOOK
' ---------------------------------------------------------------

Private Function ReadNamedRangeFromWb(wb As Workbook, sheetName As String, rangeName As String) As String
    On Error Resume Next
    Dim nm As name, refStr As String

    For Each nm In wb.Names
        If InStr(1, nm.name, rangeName, vbTextCompare) > 0 Then
            refStr = nm.RefersTo
            If InStr(1, refStr, sheetName, vbTextCompare) > 0 Or _
               InStr(1, nm.name, sheetName, vbTextCompare) > 0 Then
                If Left(nm.RefersTo, 2) = "=""" Then
                    ReadNamedRangeFromWb = Mid(nm.RefersTo, 3, Len(nm.RefersTo) - 3)
                Else
                    ReadNamedRangeFromWb = Mid(nm.RefersTo, 2)
                End If
                Exit Function
            End If
        End If
    Next nm

    Dim testNm As name
    Set testNm = Nothing
    Set testNm = wb.Names("'" & sheetName & "'!" & rangeName)
    If Not testNm Is Nothing Then
        refStr = testNm.RefersTo
        If Left(refStr, 2) = "=""" Then
            ReadNamedRangeFromWb = Mid(refStr, 3, Len(refStr) - 3)
        Else
            ReadNamedRangeFromWb = Mid(refStr, 2)
        End If
    End If
    On Error GoTo 0
End Function


' ---------------------------------------------------------------
' FILE I/O
' ---------------------------------------------------------------

Private Sub IOWriteTextFile(filePath As String, content As String)
    Dim fso As Object, ts As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.CreateTextFile(filePath, True, True)
    ts.Write content
    ts.Close
End Sub

Private Function IOReadTextFile(filePath As String) As String
    Dim fso As Object, ts As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.fileExists(filePath) Then IOReadTextFile = "": Exit Function
    Set ts = fso.OpenTextFile(filePath, 1, False, -1)
    If Not ts.AtEndOfStream Then IOReadTextFile = ts.ReadAll
    ts.Close
End Function


Private Sub btnNewScript_Click()
    On Error GoTo EH

    Dim scriptName As String
    Dim folderPath As String
    Dim fullPath As String
    Dim fileNum As Integer

    scriptName = Trim$(InputBox("Enter a name for the new script:", "New Script"))
    If Len(scriptName) = 0 Then Exit Sub

    ' Ensure .py extension
    If LCase(right$(scriptName, 3)) <> ".py" Then scriptName = scriptName & ".py"

    folderPath = GetScriptFolderPath()
    If Len(folderPath) = 0 Then
        MsgBox "Could not resolve the userScripts folder.", vbExclamation
        Exit Sub
    End If

    If right$(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    fullPath = folderPath & scriptName

    If Dir(fullPath) <> "" Then
        MsgBox "A script named '" & scriptName & "' already exists.", vbExclamation
        Exit Sub
    End If

    ' Write template file
    fileNum = FreeFile
    Open fullPath For Output As #fileNum
    Print #fileNum, "from typing import Dict, Any"
    Print #fileNum, "import pandas as pd"
    Print #fileNum, "from tools import run_script_cli"
    Print #fileNum, ""
    Print #fileNum, ""
    Print #fileNum, "def transform(inputs: Dict[str, Any]) -> Dict[str, Any]:"
    Print #fileNum, "    # inputs contains DataFrames, lists, or scalars from Excel"
    Print #fileNum, "    # df = inputs.get(""df1"", pd.DataFrame())"
    Print #fileNum, ""
    Print #fileNum, "    return {}"
    Print #fileNum, ""
    Print #fileNum, ""
    Print #fileNum, "if __name__ == ""__main__"":"
    Print #fileNum, "    run_script_cli(transform)"
    Close #fileNum

    ' Refresh the script combo box and select the new file
    Dim files As Collection
    Dim f As Variant
    ComboBoxScript.Clear
    Set files = GetScriptFiles()
    If Not files Is Nothing Then
        For Each f In files
            ComboBoxScript.AddItem f
        Next f
    End If
    ComboBoxScript.value = scriptName

    MsgBox "Script '" & scriptName & "' created.", vbInformation
    Exit Sub

EH:
    If fileNum > 0 Then Close #fileNum
    MsgBox "Error creating script: " & Err.Description, vbCritical
End Sub

