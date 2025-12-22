VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEditAction 
   Caption         =   "Edit Action"
   ClientHeight    =   4656
   ClientLeft      =   90
   ClientTop       =   420
   ClientWidth     =   3570
   OleObjectBlob   =   "frmEditAction.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmEditAction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private hasLoaded As Boolean      ' <<< THE ONLY FIX ADDED
Private cur1rentSheetName As String
Private currentAction As String
Private actionData As Object
Private scriptSelected As String
Private rib As Object
' (Whatever other module-level vars you have stay here)

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

    ListBoxInput.Clear
    inputVals = Split(GetSheetValue(wb, currentSheetName, "txtPyInput"), ";")
    For i = LBound(inputVals) To UBound(inputVals)
        If Trim$(inputVals(i)) <> "" Then ListBoxInput.AddItem Trim$(inputVals(i))
    Next i

    ListBoxOutput.Clear
    outputVals = Split(GetSheetValue(wb, currentSheetName, "txtPyOutput"), ";")
    For i = LBound(outputVals) To UBound(outputVals)
        If Trim$(outputVals(i)) <> "" Then ListBoxOutput.AddItem Trim$(outputVals(i))
    Next i

    Debug.Print "[frmEditAction] Refreshed for sheet: " & currentSheetName & ", action='" & currentAction & "'"
    Exit Sub

EH:
    Debug.Print "[frmEditAction.RefreshFromContext][ERROR] " & Err.Description
End Sub



Private Sub btnSave_Click()
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
    For i = 0 To ListBoxInput.ListCount - 1
        If i > 0 Then tempInput = tempInput & ";"
        tempInput = tempInput & ListBoxInput.List(i)
    Next i

    tempOutput = ""
    For i = 0 To ListBoxOutput.ListCount - 1
        If i > 0 Then tempOutput = tempOutput & ";"
        tempOutput = tempOutput & ListBoxOutput.List(i)
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
    Set actionData(newAction) = act

    SaveActionsForSheet currentSheetName, actionData
    SaveSheetValue wb, currentSheetName, "SelectedAction", newAction
    currentAction = newAction

    SaveSheetValue wb, currentSheetName, "cmbScript", act("script")
    SaveSheetValue wb, currentSheetName, "txtPyInput", act("input")
    SaveSheetValue wb, currentSheetName, "txtPyOutput", act("output")
    scriptSelected = act("script")

    If Not rib Is Nothing Then
        rib.InvalidateControl "cmbActions"
        rib.InvalidateControl "cmbScript"
        rib.InvalidateControl "txtPyInput"
        rib.InvalidateControl "txtPyOutput"
    End If
    
    Call RefreshRibbonValues
    
    Unload Me
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
    ListBoxInput.AddItem rng.parent.name & "!" & rng.Address(False, False)
End Sub

Private Sub btnAddOutput_Click()
    Dim rng As Range
    On Error Resume Next
    Set rng = Application.InputBox("Select Output Range", Type:=8)
    On Error GoTo 0
    If rng Is Nothing Then Exit Sub
    ListBoxOutput.AddItem rng.parent.name & "!" & rng.Address(False, False)
End Sub



Private Sub btnEditInput_Click()
    Dim i As Long
    Dim original As String
    Dim updated As String
    Dim savedLeft As Single
    Dim savedTop As Single

    i = ListBoxInput.ListIndex
    If i < 0 Then
        MsgBox "Select an input item to edit.", vbExclamation
        Exit Sub
    End If

    original = ListBoxInput.List(i)

    ' Save current position
    savedLeft = Me.Left
    savedTop = Me.Top

    ' Off-screen hide
    Me.Left = -20000
    Me.Top = -20000

    ' Call child form
    updated = frmRangeSet.GetUpdatedValue(original)

    ' Restore position
    Me.Left = savedLeft
    Me.Top = savedTop

    ' Write updated value
    ListBoxInput.List(i) = updated
End Sub



Private Sub btnEditOutput_Click()
    Dim i As Long
    Dim original As String
    Dim updated As String
    Dim savedLeft As Single
    Dim savedTop As Single

    i = ListBoxOutput.ListIndex
    If i < 0 Then
        MsgBox "Select an output item to edit.", vbExclamation
        Exit Sub
    End If

    original = ListBoxOutput.List(i)

    savedLeft = Me.Left
    savedTop = Me.Top

    Me.Left = -20000
    Me.Top = -20000

    updated = frmRangeSet.GetUpdatedValue(original)

    Me.Left = savedLeft
    Me.Top = savedTop

    ListBoxOutput.List(i) = updated
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

