VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufProgress 
   Caption         =   "PyExcel Installation"
   ClientHeight    =   1410
   ClientLeft      =   180
   ClientTop       =   765
   ClientWidth     =   15960
   OleObjectBlob   =   "ufProgress.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' This code should be placed in the code module of the ufProgress UserForm.

' --- Event Handlers ---

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' This event handler prevents the user from accidentally closing the
    ' installation progress form using the "X" button in the title bar.
    ' If they try, it cancels the close action and shows an informational message.
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        MsgBox "Please wait for the installation to complete.", vbInformation, "Action Prevented"
    End If
End Sub

