VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRangeSet 
   Caption         =   "Set Range"
   ClientHeight    =   1560
   ClientLeft      =   96
   ClientTop       =   416
   ClientWidth     =   3512
   OleObjectBlob   =   "frmRangeSet.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmRangeSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private returnValue As String
Private originalValue As String

'=========================================
' Entry point – called from parent
'=========================================
Public Function GetUpdatedValue(oldValue As String) As String
    originalValue = oldValue
    returnValue = oldValue

    Preload oldValue
    Me.Show vbModal

    GetUpdatedValue = returnValue
'    MsgBox GetUpdatedValue & " --- IN the sub function"
    Unload Me
End Function

'=========================================
' Load name/range from old value
'=========================================
Private Sub Preload(v As String)
    Dim p As Long

    NameBox.text = ""
    RefBox.text = ""

    v = Trim$(v)
    If Len(v) = 0 Then Exit Sub

    p = InStr(v, "=")

    If p > 0 Then
        ' NAME = RANGE
        NameBox.text = Left$(v, p - 1)
        RefBox.text = Mid$(v, p + 1)
    Else
        ' RANGE ONLY
        RefBox.text = v
    End If
End Sub

'=========================================
' Save button – name optional
'=========================================
Private Sub SetButton_Click()
    Dim nm As String
    Dim ref As String
    Dim rng As Range
    Dim ws As Worksheet
    Dim finalRef As String

    nm = Trim$(NameBox.text)
    ref = Trim$(RefBox.text)

    If ref = "" Then
        MsgBox "Select a range.", vbExclamation
        Exit Sub
    End If

    On Error Resume Next
    Set rng = Application.Range(ref)
    On Error GoTo 0

    If rng Is Nothing Then
        MsgBox "Invalid range.", vbExclamation
        Exit Sub
    End If

    Set ws = rng.Worksheet
    finalRef = ws.name & "!" & rng.Address(False, False)

    If nm = "" Then
        returnValue = finalRef
    Else
        returnValue = nm & "=" & finalRef
    End If

'    MsgBox returnValue & " --- IN the form"
    
    Me.Hide
End Sub


'=========================================
' Cancel button – return original
'=========================================
Private Sub CancelButton_Click()
    returnValue = originalValue
    Me.Hide
End Sub

