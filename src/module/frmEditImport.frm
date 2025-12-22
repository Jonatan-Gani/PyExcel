VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEditImport 
   Caption         =   "Edit Import"
   ClientHeight    =   3592
   ClientLeft      =   90
   ClientTop       =   420
   ClientWidth     =   3570
   OleObjectBlob   =   "frmEditImport.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmEditImport"
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
'    BoxInput.value = GetSheetValue(ActiveWorkbook, wsName, "txtImportInput")
'    BoxOutput.value = GetSheetValue(ActiveWorkbook, wsName, "txtImportOutput")
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
'    SaveSheetValue ActiveWorkbook, wsName, "txtImportInput", inputVal
'    SaveSheetValue ActiveWorkbook, wsName, "txtImportOutput", outputVal
'
'    If Not rib Is Nothing Then
'        rib.InvalidateControl "txtImportInput"
'        rib.InvalidateControl "txtImportOutput"
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
'    Dim sh As Object
'    Dim folderObj As Object
'    Dim startPath As String
'    Dim chosenPath As String
'
'    startPath = Trim(BoxInput.value)
'
'    ' Validate or pick a fallback
'    If Len(startPath) = 0 Or Dir(startPath, vbDirectory) = "" Then
'        If Len(ActiveWorkbook.path) > 0 Then
'            startPath = ActiveWorkbook.path
'        Else
'            startPath = Environ$("USERPROFILE") & "\Desktop"
'        End If
'    End If
'
'    ' Create Shell object
'    Set sh = CreateObject("Shell.Application")
'
'    ' Use "My Computer" as the base so navigation is not restricted
'    Set folderObj = sh.BrowseForFolder(0, "Select file or folder", &H11, "This PC")
'
'    ' Preselect start folder after dialog opens
'    If Not folderObj Is Nothing Then
'        chosenPath = folderObj.items.Item().path
'        BoxInput.value = chosenPath
'    Else
'        ' If cancelled, no change
'    End If
'End Sub
'
'Private Sub btnEditOutput_Click()
'    Dim rng As Range
'
'    On Error Resume Next
'    Set rng = Application.InputBox("Select Output Range", Type:=8)
'    On Error GoTo 0
'
'    If rng Is Nothing Then Exit Sub
'
'    BoxOutput.value = rng.parent.name & "!" & rng.Address(False, False)
'End Sub
'
'
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

    ListBoxInput.Clear
    vals = Split(GetSheetValue(wb, wsName, "txtImportInput"), ";")
    For i = LBound(vals) To UBound(vals)
        If Trim$(vals(i)) <> "" Then ListBoxInput.AddItem Trim$(vals(i))
    Next i

    ListBoxOutput.Clear
    vals = Split(GetSheetValue(wb, wsName, "txtImportOutput"), ";")
    For i = LBound(vals) To UBound(vals)
        If Trim$(vals(i)) <> "" Then ListBoxOutput.AddItem Trim$(vals(i))
    Next i

    Exit Sub
EH:
    Debug.Print "[frmImport.Initialize][ERROR] " & Err.Description
End Sub


Private Sub btnSave_Click()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim wsName As String
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

    SaveSheetValue wb, wsName, "txtImportInput", tempInput
    SaveSheetValue wb, wsName, "txtImportOutput", tempOutput

    If Not rib Is Nothing Then
        rib.InvalidateControl "txtImportInput"
        rib.InvalidateControl "txtImportOutput"
    End If

    Unload Me
End Sub


Private Sub btnDiscard_Click()
    Unload Me
End Sub


' =========================================
' INPUT = folder path
' =========================================
Private Sub btnEditInput_Click()
    Dim sh As Object
    Dim folderObj As Object
    Dim chosenPath As String
    Dim idx As Long

    idx = ListBoxInput.ListIndex
    If idx < 0 Then
        MsgBox "Select an input folder to edit.", vbExclamation
        Exit Sub
    End If

    Set sh = CreateObject("Shell.Application")
    Set folderObj = sh.BrowseForFolder(0, "Select folder", &H11, "This PC")

    If folderObj Is Nothing Then Exit Sub

    chosenPath = folderObj.items.Item().path
    ListBoxInput.List(idx) = chosenPath
End Sub


Private Sub btnAddInput_Click()
    Dim sh As Object
    Dim folderObj As Object
    Dim chosenPath As String

    Set sh = CreateObject("Shell.Application")
    Set folderObj = sh.BrowseForFolder(0, "Select folder", &H11, "This PC")
    If folderObj Is Nothing Then Exit Sub

    chosenPath = folderObj.items.Item().path
    ListBoxInput.AddItem chosenPath
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

