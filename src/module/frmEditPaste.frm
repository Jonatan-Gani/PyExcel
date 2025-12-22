VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEditPaste 
   Caption         =   "Edit Paste"
   ClientHeight    =   3592
   ClientLeft      =   90
   ClientTop       =   420
   ClientWidth     =   3570
   OleObjectBlob   =   "frmEditPaste.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmEditPaste"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Private Sub UserForm_Initialize()
'    Dim wsName As String
'
'    wsName = ActiveSheet.name
'
'    Debug.Print wsName
'
'    ' --- Load stored values from sheet ---
'    BoxOutput.value = GetSheetValue(ActiveWorkbook, wsName, "txtPasteOutput")
'End Sub
'
'Private Sub btnSave_Click()
'    Dim wsName As String
'    Dim outputVal As String
'
'    outputVal = BoxOutput.value
'    wsName = ActiveSheet.name
'
'    Debug.Print wsName
'
'    SaveSheetValue ActiveWorkbook, wsName, "txtPasteOutput", outputVal
'
'    If Not rib Is Nothing Then
'        rib.InvalidateControl "txtPasteOutput"
'    End If
'
'    Unload Me
'End Sub
'
'Private Sub btnDiscard_Click()
'    Unload Me
'End Sub
'
'Private Sub btnEditOutput_Click()
'    Dim rng As Range
'    Set rng = Application.InputBox("Select Output Range", Type:=8)
'    If rng Is Nothing Then Exit Sub
'
'    BoxOutput.value = rng.parent.name & "!" & rng.Address(False, False)
'End Sub
'
Private Sub UserForm_Initialize()
    On Error GoTo EH

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim wsName As String
    Dim vals As Variant
    Dim i As Long

    Set wb = HostManager_GetCurrentWorkbook()
    Set ws = HostManager_GetCurrentSheet()

    If wb Is Nothing Or ws Is Nothing Then
        MsgBox "No active workbook or sheet context.", vbExclamation
        Unload Me
        Exit Sub
    End If

    wsName = ws.name

    ListBoxOutput.Clear
    vals = Split(GetSheetValue(wb, wsName, "txtPasteOutput"), ";")
    For i = LBound(vals) To UBound(vals)
        If Trim$(vals(i)) <> "" Then ListBoxOutput.AddItem Trim$(vals(i))
    Next i

    Exit Sub
EH:
    Debug.Print "[frmPaste.Initialize][ERROR] " & Err.Description
End Sub


Private Sub btnSave_Click()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim wsName As String
    Dim i As Long
    Dim tempOutput As String

    Set wb = HostManager_GetCurrentWorkbook()
    Set ws = HostManager_GetCurrentSheet()

    If wb Is Nothing Or ws Is Nothing Then
        MsgBox "No active workbook or sheet context.", vbExclamation
        Exit Sub
    End If

    wsName = ws.name

    tempOutput = ""
    For i = 0 To ListBoxOutput.ListCount - 1
        If i > 0 Then tempOutput = tempOutput & ";"
        tempOutput = tempOutput & ListBoxOutput.List(i)
    Next i

    SaveSheetValue wb, wsName, "txtPasteOutput", tempOutput

    If Not rib Is Nothing Then
        rib.InvalidateControl "txtPasteOutput"
    End If

    Unload Me
End Sub


Private Sub btnDiscard_Click()
    Unload Me
End Sub


' =========================================
' OUTPUT = Excel Ranges only (OFF-SCREEN HIDE FIX APPLIED)
' =========================================
Private Sub btnEditOutput_Click()
    Dim rng As Range
    Dim idx As Long
    Dim savedLeft As Single
    Dim savedTop As Single

    idx = ListBoxOutput.ListIndex
    If idx < 0 Then
        MsgBox "Select an output item to edit.", vbExclamation
        Exit Sub
    End If

    ' Save current form position
    savedLeft = Me.Left
    savedTop = Me.Top

    ' Off-screen hide (no .Visible, no .Show)
    Me.Left = -20000
    Me.Top = -20000

    ' Let the user pick a range
    On Error Resume Next
    Set rng = Application.InputBox("Select Output Range", Type:=8)
    On Error GoTo 0

    ' Return form to original position
    Me.Left = savedLeft
    Me.Top = savedTop

    If rng Is Nothing Then Exit Sub

    ' Update the list value
    ListBoxOutput.List(idx) = rng.parent.name & "!" & rng.Address(False, False)
End Sub


Private Sub btnAddOutput_Click()
    Dim rng As Range

    On Error Resume Next
    Set rng = Application.InputBox("Select Output Range", Type:=8)
    On Error GoTo 0
    If rng Is Nothing Then Exit Sub

    ListBoxOutput.AddItem rng.parent.name & "!" & rng.Address(False, False)
End Sub


Private Sub btnDeleteOutput_Click()
    Dim idx As Long

    idx = ListBoxOutput.ListIndex
    If idx < 0 Then
        MsgBox "Select an output item to delete.", vbExclamation
        Exit Sub
    End If

    ListBoxOutput.RemoveItem idx
End Sub

