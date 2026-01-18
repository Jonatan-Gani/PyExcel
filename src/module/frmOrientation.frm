VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmOrientation 
   Caption         =   "List direction"
   ClientHeight    =   1488
   ClientLeft      =   120
   ClientTop       =   472
   ClientWidth     =   4440
   OleObjectBlob   =   "frmOrientation.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmOrientation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' --- UserForm: frmOrientation (two buttons with explicit labels) ---
' Controls (exact names/captions):
'   CommandButton: Name=cmdHorizontal, Caption="Horizontal"
'   CommandButton: Name=cmdVertical,   Caption="Vertical"
' Form (Name=frmOrientation, Caption="Choose paste direction")

' Code behind frmOrientation:
Option Explicit
Public Choice As String  ' "H" or "V"; empty means closed without choosing

Private Sub UserForm_Initialize()
    Choice = ""
End Sub

Private Sub cmdHorizontal_Click()
    Choice = "H"
    Me.Hide
End Sub

Private Sub cmdVertical_Click()
    Choice = "V"
    Me.Hide
End Sub

