VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmExportWizard 
   Caption         =   "frmExportWizard"
   ClientHeight    =   5376
   ClientLeft      =   184
   ClientTop       =   736
   ClientWidth     =   4968
   OleObjectBlob   =   "frmExportWizard.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmExportWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'== UserForm: frmExportWizard ==
Option Explicit

' State
Private dictSources As Object
Private dictDestinations As Object
Private nextSourceID As Long
Private nextDestinationID As Long

' Layout
Private Const CTRL_HEIGHT As Double = 15.6
Private Const CTRL_WIDTH_TEXT As Double = 140
Private Const CTRL_WIDTH_BTN  As Double = 20
Private Const CTRL_WIDTH_FORMAT As Double = 100
Private Const V_MARGIN As Double = 8
Private Const H_GAP As Double = 10
Private Const LEFT_START As Double = 10


'==============================================================================
' Initialize – only dictionaries and counters
'==============================================================================
Private Sub UserForm_Initialize()
    Set dictSources = CreateObject("Scripting.Dictionary")
    Set dictDestinations = CreateObject("Scripting.Dictionary")
    nextSourceID = 1
    nextDestinationID = 1
    lblStatus.Caption = "Ready."
End Sub


'==============================================================================
' Public seeding method for external inputs
'==============================================================================
Public Sub InitializeFromInputs(ByVal sourceRef As String, ByVal outputPath As String)
    ' Clear any previous dynamic controls
    Dim ctl As control
    On Error Resume Next
    Dim pg As MSForms.Page, c As control
    For Each ctl In Me.Controls
        If TypeName(ctl) = "MultiPage" Then
            For Each pg In ctl.Pages
                For Each c In pg.Controls
                    pg.Controls.Remove c.name
                Next c
            Next pg
        End If
    Next ctl
    On Error GoTo 0

    ' Reset state
    Set dictSources = CreateObject("Scripting.Dictionary")
    Set dictDestinations = CreateObject("Scripting.Dictionary")
    nextSourceID = 1
    nextDestinationID = 1

    ' Build UI with given values (empty if none)
    CreateSourceGroup Trim$(sourceRef)
    CreateDestinationGroup Trim$(outputPath)

    lblStatus.Caption = "Ready."
End Sub


'==============================================================================
' Helpers: baseline (below static controls) + next stacking position
'==============================================================================
Private Function BaseTop_Sources() As Double
    Dim pg As MSForms.Page: Set pg = mpMain.Pages(0)
    Dim bottom As Double: bottom = 10
    On Error Resume Next
    If Not pg.Controls("btnAddRange") Is Nothing Then
        With pg.Controls("btnAddRange")
            bottom = .Top + .Height + V_MARGIN
        End With
    End If
    On Error GoTo 0
    BaseTop_Sources = bottom
End Function

Private Function BaseTop_Destinations() As Double
    Dim pg As MSForms.Page: Set pg = mpMain.Pages(1)
    Dim bottom As Double: bottom = 10
    On Error Resume Next
    If Not pg.Controls("btnAddDestination") Is Nothing Then
        With pg.Controls("btnAddDestination")
            bottom = .Top + .Height + V_MARGIN
        End With
    End If
    On Error GoTo 0
    BaseTop_Destinations = bottom
End Function

Private Function NextTop_Sources() As Double
    Dim pg As MSForms.Page: Set pg = mpMain.Pages(0)
    Dim c As control, bottom As Double
    bottom = BaseTop_Sources()
    For Each c In pg.Controls
        If (Left$(c.name, 8) = "srcText_") Or (Left$(c.name, 8) = "srcEdit_") Or (Left$(c.name, 10) = "srcRemove_") Then
            If c.Top + c.Height > bottom Then bottom = c.Top + c.Height
        End If
    Next c
    NextTop_Sources = bottom + IIf(bottom = BaseTop_Sources(), 0, V_MARGIN)
End Function

Private Function NextTop_Destinations() As Double
    Dim pg As MSForms.Page: Set pg = mpMain.Pages(1)
    Dim c As control, bottom As Double
    bottom = BaseTop_Destinations()
    For Each c In pg.Controls
        If (Left$(c.name, 8) = "dstText_") Or (Left$(c.name, 8) = "dstEdit_") Or _
           (Left$(c.name, 10) = "dstRemove_") Or (Left$(c.name, 9) = "dstLabel_") Or _
           (Left$(c.name, 9) = "dstCombo_") Then
            If c.Top + c.Height > bottom Then bottom = c.Top + c.Height
        End If
    Next c
    NextTop_Destinations = bottom + IIf(bottom = BaseTop_Destinations(), 0, V_MARGIN)
End Function


'==============================================================================
' SOURCES page
'==============================================================================
Private Sub btnAddRange_Click()
    CreateSourceGroup ""
End Sub

Private Sub CreateSourceGroup(ByVal defaultRange As String)
    Dim container As MSForms.Page: Set container = mpMain.Pages(0)
    Dim topOffset As Double: topOffset = NextTop_Sources()

    Dim txt As MSForms.TextBox
    Dim btnEdit As MSForms.CommandButton
    Dim btnRemove As MSForms.CommandButton

    Set txt = container.Controls.Add("Forms.TextBox.1", "srcText_" & nextSourceID, True)
    With txt
        .Top = topOffset
        .Left = LEFT_START
        .Height = CTRL_HEIGHT
        .Width = CTRL_WIDTH_TEXT
        .text = defaultRange
    End With

    Set btnEdit = container.Controls.Add("Forms.CommandButton.1", "srcEdit_" & nextSourceID, True)
    With btnEdit
        .Top = topOffset
        .Left = txt.Left + txt.Width + H_GAP
        .Height = CTRL_HEIGHT
        .Width = CTRL_WIDTH_BTN
        .Caption = "..."
    End With

    Set btnRemove = container.Controls.Add("Forms.CommandButton.1", "srcRemove_" & nextSourceID, True)
    With btnRemove
        .Top = topOffset
        .Left = btnEdit.Left + btnEdit.Width + H_GAP
        .Height = CTRL_HEIGHT
        .Width = CTRL_WIDTH_BTN
        .Caption = "X"
    End With

    Dim gi As Collection: Set gi = New Collection
    gi.Add txt
    gi.Add btnEdit
    gi.Add btnRemove
    dictSources.Add CStr(nextSourceID), gi
    nextSourceID = nextSourceID + 1
End Sub

Private Sub SourceEditClicked(ByVal ctrlName As String)
    Dim parts() As String: parts = Split(ctrlName, "_")
    Dim id As String: id = parts(1)
    Dim gi As Collection: Set gi = dictSources(id)
    Dim txt As MSForms.TextBox: Set txt = gi(1)

    Dim rng As Range
    On Error Resume Next
    Set rng = Application.InputBox("Select range to export (any sheet)", "Pick Range", Type:=8)
    On Error GoTo 0
    If rng Is Nothing Then Exit Sub

    txt.text = rng.Address(External:=True)
End Sub

Private Sub SourceRemoveClicked(ByVal ctrlName As String)
    Dim parts() As String: parts = Split(ctrlName, "_")
    Dim id As String: id = parts(1)
    Dim gi As Collection: Set gi = dictSources(id)
    Dim ctrl As control
    For Each ctrl In gi
        mpMain.Pages(0).Controls.Remove ctrl.name
    Next ctrl
    dictSources.Remove id
    Reflow_Sources
End Sub

Private Sub Reflow_Sources()
    If dictSources.count = 0 Then Exit Sub
    Dim keys() As Variant: keys = dictSources.keys

    Dim topOffset As Double: topOffset = BaseTop_Sources()
    Dim i As Long
    For i = LBound(keys) To UBound(keys)
        Dim gi As Collection: Set gi = dictSources(CStr(keys(i)))
        Dim txt As MSForms.TextBox: Set txt = gi(1)
        Dim bE As MSForms.CommandButton: Set bE = gi(2)
        Dim bR As MSForms.CommandButton: Set bR = gi(3)

        txt.Top = topOffset
        bE.Top = topOffset
        bR.Top = topOffset

        topOffset = txt.Top + txt.Height + V_MARGIN
    Next i
End Sub


'==============================================================================
' DESTINATIONS page
'==============================================================================
Private Sub btnAddDestination_Click()
    CreateDestinationGroup ""
End Sub

Private Sub CreateDestinationGroup(ByVal defaultPath As String)
    Dim container As MSForms.Page: Set container = mpMain.Pages(1)
    Dim topOffset As Double: topOffset = NextTop_Destinations()

    Dim txt As MSForms.TextBox
    Dim btnEdit As MSForms.CommandButton
    Dim btnRemove As MSForms.CommandButton
    Dim lblFormat As MSForms.label
    Dim cmbFormat As MSForms.ComboBox

    Set txt = container.Controls.Add("Forms.TextBox.1", "dstText_" & nextDestinationID, True)
    With txt
        .Top = topOffset
        .Left = LEFT_START
        .Height = CTRL_HEIGHT
        .Width = CTRL_WIDTH_TEXT
        .text = defaultPath
    End With

    Set btnEdit = container.Controls.Add("Forms.CommandButton.1", "dstEdit_" & nextDestinationID, True)
    With btnEdit
        .Top = topOffset
        .Left = txt.Left + txt.Width + H_GAP
        .Height = CTRL_HEIGHT
        .Width = CTRL_WIDTH_BTN
        .Caption = "..."
    End With

    Set btnRemove = container.Controls.Add("Forms.CommandButton.1", "dstRemove_" & nextDestinationID, True)
    With btnRemove
        .Top = topOffset
        .Left = btnEdit.Left + btnEdit.Width + H_GAP
        .Height = CTRL_HEIGHT
        .Width = CTRL_WIDTH_BTN
        .Caption = "X"
    End With

    Set lblFormat = container.Controls.Add("Forms.Label.1", "dstLabel_" & nextDestinationID, True)
    With lblFormat
        .Top = topOffset + CTRL_HEIGHT + 2
        .Left = txt.Left
        .Caption = "Format"
    End With

    Set cmbFormat = container.Controls.Add("Forms.ComboBox.1", "dstCombo_" & nextDestinationID, True)
    With cmbFormat
        .Top = topOffset + CTRL_HEIGHT
        .Left = txt.Left + txt.Width - CTRL_WIDTH_FORMAT
        .Height = CTRL_HEIGHT
        .Width = CTRL_WIDTH_FORMAT
        .AddItem "Excel Workbook (*.xlsx)"
        .AddItem "Macro-Enabled Workbook (*.xlsm)"
        .AddItem "Binary Workbook (*.xlsb)"
        .AddItem "CSV (Comma Delimited) (*.csv)"
        .ListIndex = 0
    End With

    Dim gi As Collection: Set gi = New Collection
    gi.Add txt
    gi.Add btnEdit
    gi.Add btnRemove
    gi.Add lblFormat
    gi.Add cmbFormat
    dictDestinations.Add CStr(nextDestinationID), gi
    nextDestinationID = nextDestinationID + 1
End Sub

Private Sub DestinationEditClicked(ByVal ctrlName As String)
    Dim parts() As String: parts = Split(ctrlName, "_")
    Dim id As String: id = parts(1)
    Dim gi As Collection: Set gi = dictDestinations(id)
    Dim txt As MSForms.TextBox: Set txt = gi(1)

    Dim chosenPath As Variant
    chosenPath = Application.GetSaveAsFilename(InitialFileName:="export.xlsx", _
        FileFilter:="Excel Workbook (*.xlsx), *.xlsx," & _
                    "Macro-Enabled Workbook (*.xlsm), *.xlsm," & _
                    "Binary Workbook (*.xlsb), *.xlsb," & _
                    "CSV (Comma Delimited) (*.csv), *.csv", _
        Title:="Select Export Destination")
    If chosenPath <> False Then txt.text = chosenPath
End Sub

Private Sub DestinationRemoveClicked(ByVal ctrlName As String)
    Dim parts() As String: parts = Split(ctrlName, "_")
    Dim id As String: id = parts(1)
    Dim gi As Collection: Set gi = dictDestinations(id)
    Dim ctrl As control
    For Each ctrl In gi
        mpMain.Pages(1).Controls.Remove ctrl.name
    Next ctrl
    dictDestinations.Remove id
    Reflow_Destinations
End Sub

Private Sub Reflow_Destinations()
    If dictDestinations.count = 0 Then Exit Sub
    Dim keys() As Variant: keys = dictDestinations.keys

    Dim topOffset As Double: topOffset = BaseTop_Destinations()
    Dim i As Long
    For i = LBound(keys) To UBound(keys)
        Dim gi As Collection: Set gi = dictDestinations(CStr(keys(i)))
        Dim txt As MSForms.TextBox: Set txt = gi(1)
        Dim bE As MSForms.CommandButton: Set bE = gi(2)
        Dim bR As MSForms.CommandButton: Set bR = gi(3)
        Dim lbl As MSForms.label: Set lbl = gi(4)
        Dim cmb As MSForms.ComboBox: Set cmb = gi(5)

        txt.Top = topOffset
        bE.Top = topOffset
        bR.Top = topOffset
        lbl.Top = topOffset + CTRL_HEIGHT + 2
        cmb.Top = topOffset + CTRL_HEIGHT

        topOffset = cmb.Top + cmb.Height + V_MARGIN
    Next i
End Sub


'==============================================================================
' Export
'==============================================================================
Private Sub btnExport_Click()
    lblStatus.Caption = "Validating..."
    If dictSources.count = 0 Then lblStatus.Caption = "No sources defined.": Exit Sub
    If dictDestinations.count = 0 Then lblStatus.Caption = "No destinations defined.": Exit Sub

    Dim sKeys() As Variant, dKeys() As Variant
    sKeys = dictSources.keys
    dKeys = dictDestinations.keys

    Dim i As Long, j As Long
    For i = LBound(sKeys) To UBound(sKeys)
        Dim giS As Collection: Set giS = dictSources(CStr(sKeys(i)))
        Dim srcTxt As MSForms.TextBox: Set srcTxt = giS(1)

        Dim rng As Range
        On Error Resume Next
        Set rng = Range(srcTxt.text)
        On Error GoTo 0
        If rng Is Nothing Then lblStatus.Caption = "Invalid range: " & srcTxt.text: Exit Sub

        For j = LBound(dKeys) To UBound(dKeys)
            Dim giD As Collection: Set giD = dictDestinations(CStr(dKeys(j)))
            Dim dstTxt As MSForms.TextBox: Set dstTxt = giD(1)
            Dim cmb As MSForms.ComboBox: Set cmb = giD(5)

            Dim ext As String
            Select Case cmb.text
                Case "Excel Workbook (*.xlsx)": ext = "xlsx"
                Case "Macro-Enabled Workbook (*.xlsm)": ext = "xlsm"
                Case "Binary Workbook (*.xlsb)": ext = "xlsb"
                Case "CSV (Comma Delimited) (*.csv)": ext = "csv"
                Case Else: lblStatus.Caption = "Unknown format for destination: " & dstTxt.text: Exit Sub
            End Select

            If ext = "csv" Then
                ExportRangeToCSV rng, dstTxt.text
            Else
                ExportSingleSheetToExcel rng, dstTxt.text, ext
            End If
        Next j
    Next i

    lblStatus.Caption = "Export complete."
    MsgBox "Export complete.", vbInformation
    Unload Me
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub


'==============================================================================
' Utility
'==============================================================================
Private Sub SortKeysNumeric(ByRef arr As Variant)
    Dim i As Long, j As Long, tmp As Variant
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If CLng(arr(j)) < CLng(arr(i)) Then
                tmp = arr(i): arr(i) = arr(j): arr(j) = tmp
            End If
        Next j
    Next i
End Sub


