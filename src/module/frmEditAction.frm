VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEditAction 
   Caption         =   "Manage Action"
   ClientHeight    =   4584
   ClientLeft      =   96
   ClientTop       =   416
   ClientWidth     =   7144
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
