VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEditExport 
   Caption         =   "Edit Export"
   ClientHeight    =   3592
   ClientLeft      =   96
   ClientTop       =   416
   ClientWidth     =   3568
   OleObjectBlob   =   "frmEditExport.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmEditExport"
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
'    BoxInput.value = GetSheetValue(ActiveWorkbook, wsName, "txtExportInput")
'    BoxOutput.value = GetSheetValue(ActiveWorkbook, wsName, "txtExportOutput")
'End Sub
'
'Private Sub btnSave_Click()
'    Dim wsName As String
'    Dim inputVal As String
'    Dim outputVal As String
'
'    inputVal = BoxInput.value
'    outputVal = BoxOutput.value
'    wsName = ActiveSheet.name
'
'    Debug.Print wsName
'
'    SaveSheetValue ActiveWorkbook, wsName, "txtExportInput", inputVal
'    SaveSheetValue ActiveWorkbook, wsName, "txtExportOutput", outputVal
'
'    If Not rib Is Nothing Then
'        rib.InvalidateControl "txtExportInput"
'        rib.InvalidateControl "txtExportOutput"
'    End If
'
'    Unload Me
'End Sub
'
'Private Sub btnDiscard_Click()
'    Unload Me
'End Sub
'
'Private Sub btnEditInput_Click()
'    Dim v As Variant
'    Dim rng As Range
'
'    Set v = Nothing
'    On Error Resume Next
'    Set rng = Application.InputBox("Select Input Range", Type:=8)
'    On Error GoTo 0
'
'    If rng Is Nothing Then Exit Sub
'
'    BoxInput.value = rng.parent.name & "!" & rng.Address(False, False)
'End Sub
'
'
'Private Sub btnEditOutput_Click()
'    Dim sh As Object
'    Dim folderObj As Object
'    Dim startPath As String
'    Dim chosenPath As String
'
'    startPath = Trim(BoxOutput.value)
'
'    ' Validate or set fallback
'    If Len(startPath) = 0 Or Dir(startPath, vbDirectory) = "" Then
'        If Len(ActiveWorkbook.path) > 0 Then
'            startPath = ActiveWorkbook.path
'        Else
'            startPath = Environ$("USERPROFILE") & "\Desktop"
'        End If
'    End If
'
'    ' Open folder picker starting from "This PC"
'    Set sh = CreateObject("Shell.Application")
'    Set folderObj = sh.BrowseForFolder(0, "Select Output Folder", &H11, "This PC")
'
'    If Not folderObj Is Nothing Then
'        chosenPath = folderObj.items.Item().path
'        BoxOutput.value = chosenPath
'    End If
'End Sub
'
'
'
' =========================================
' Initialization
' =========================================
Private Sub UserForm_Initialize()
    On Error GoTo EH

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim wsName As String
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

    wsName = ws.name

    ListBoxInput.Clear
    inputVals = Split(GetSheetValue(wb, wsName, "txtExportInput"), ";")
    For i = LBound(inputVals) To UBound(inputVals)
        If Trim$(inputVals(i)) <> "" Then ListBoxInput.AddItem Trim$(inputVals(i))
    Next i

    ListBoxOutput.Clear
    outputVals = Split(GetSheetValue(wb, wsName, "txtExportOutput"), ";")
    For i = LBound(outputVals) To UBound(outputVals)
        If Trim$(outputVals(i)) <> "" Then ListBoxOutput.AddItem Trim$(outputVals(i))
    Next i

    Exit Sub
EH:
    Debug.Print "[frmExport.Initialize][ERROR] " & Err.Description
End Sub


' =========================================
' Save
' =========================================
Private Sub btnSave_Click()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim wsName As String
    Dim inputVal As String
    Dim outputVal As String
    Dim i As Long
    Dim tempInput As String
    Dim tempOutput As String

    Set wb = HostManager_GetCurrentWorkbook()
    Set ws = HostManager_GetCurrentSheet()

    If wb Is Nothing Or ws Is Nothing Then
        MsgBox "No active workbook or sheet context.", vbExclamation
        Exit Sub
    End If

    wsName = ws.name

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

    SaveSheetValue wb, wsName, "txtExportInput", inputVal
    SaveSheetValue wb, wsName, "txtExportOutput", outputVal

    If Not rib Is Nothing Then
        rib.InvalidateControl "txtExportInput"
        rib.InvalidateControl "txtExportOutput"
    End If

    Unload Me
End Sub


' =========================================
' Discard
' =========================================
Private Sub btnDiscard_Click()
    Unload Me
End Sub


' =========================================
' Input add/edit/delete
' =========================================
Private Sub btnEditInput_Click()
    Dim rng As Range
    Dim idx As Long
    Dim savedLeft As Single
    Dim savedTop As Single

    idx = ListBoxInput.ListIndex
    If idx < 0 Then
        MsgBox "Select an input item to edit.", vbExclamation
        Exit Sub
    End If

    ' Save current form position
    savedLeft = Me.Left
    savedTop = Me.Top

    ' Off-screen hide (avoid .Visible / .Show)
    Me.Left = -20000
    Me.Top = -20000

    ' Let user select a range
    On Error Resume Next
    Set rng = Application.InputBox("Select Input Range", Type:=8)
    On Error GoTo 0

    ' Restore form position
    Me.Left = savedLeft
    Me.Top = savedTop

    If rng Is Nothing Then Exit Sub

    ' Update the list item
    ListBoxInput.List(idx) = rng.parent.name & "!" & rng.Address(False, False)
End Sub


Private Sub btnAddInput_Click()
    Dim rng As Range
    On Error Resume Next
    Set rng = Application.InputBox("Select Input Range", Type:=8)
    On Error GoTo 0
    If rng Is Nothing Then Exit Sub
    ListBoxInput.AddItem rng.parent.name & "!" & rng.Address(False, False)
End Sub

Private Sub btnDeleteInput_Click()
    Dim idx As Long
    idx = ListBoxInput.ListIndex
    If idx < 0 Then
        MsgBox "Select an input item to delete.", vbExclamation
        Exit Sub
    End If
    ListBoxInput.RemoveItem idx
End Sub


' =========================================
' Output add/edit/delete
' =========================================
Private Sub btnEditOutput_Click()
    Dim folderDlg As Object
    Dim sh As Object
    Dim chosenPath As String

    Set sh = CreateObject("Shell.Application")
    Set folderDlg = sh.BrowseForFolder(0, "Select Output Folder", &H11, "This PC")
    If folderDlg Is Nothing Then Exit Sub

    chosenPath = folderDlg.items.Item().path
    ListBoxOutput.List(ListBoxOutput.ListIndex) = chosenPath
End Sub

Private Sub btnAddOutput_Click()
    Dim folderDlg As Object
    Dim sh As Object
    Dim chosenPath As String

    Set sh = CreateObject("Shell.Application")
    Set folderDlg = sh.BrowseForFolder(0, "Select Output Folder", &H11, "This PC")
    If folderDlg Is Nothing Then Exit Sub

    chosenPath = folderDlg.items.Item().path
    ListBoxOutput.AddItem chosenPath
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

