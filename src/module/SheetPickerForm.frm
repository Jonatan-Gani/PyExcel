VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SheetPickerForm 
   Caption         =   "Pick sheet"
   ClientHeight    =   1504
   ClientLeft      =   120
   ClientTop       =   464
   ClientWidth     =   3480
   OleObjectBlob   =   "SheetPickerForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SheetPickerForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public SelectedSheet As String

Private Sub UserForm_Initialize()
    Dim ws As Worksheet
    Me.cmbSheets.Clear
    For Each ws In ActiveWorkbook.Worksheets
        Me.cmbSheets.AddItem ws.name
    Next ws
    If Me.cmbSheets.listCount > 0 Then Me.cmbSheets.ListIndex = 0
End Sub

Private Sub btnOK_Click()
    If Me.cmbSheets.ListIndex >= 0 Then
        SelectedSheet = Me.cmbSheets.value
    Else
        SelectedSheet = ""
    End If
    Me.Hide
End Sub

Private Sub btnCancel_Click()
    SelectedSheet = ""
    Me.Hide
End Sub


