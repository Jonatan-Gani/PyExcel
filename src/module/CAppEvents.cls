VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CAppEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Option Explicit
'Public WithEvents App As Application
'
'Private Sub Class_Initialize()
'    Set App = Application
'    Debug.Print "[" & Format$(Now, "hh:nn:ss") & "] [CAppEvents] [Init] bound"
'End Sub
'
'Private Sub App_WorkbookActivate(ByVal wb As Workbook)
'    On Error Resume Next
'    If wb Is Nothing Then Exit Sub
'    If HostManager_IsAddinWorkbook(wb) Then Exit Sub
'    HostManager_ActivateWorkbook wb, "WorkbookActivate"
'End Sub
'
'Private Sub App_SheetActivate(ByVal sh As Object)
'    On Error Resume Next
'    If Not TypeOf sh Is Worksheet Then Exit Sub
'    Dim wb As Workbook: Set wb = sh.parent
'    If HostManager_IsAddinWorkbook(wb) Then Exit Sub
'    HostManager_ActivateSheet wb, sh, "SheetActivate"
'End Sub
'
'Private Sub App_WindowActivate(ByVal wb As Workbook, ByVal wn As Window)
'    On Error Resume Next
'    If wb Is Nothing Then Exit Sub
'    If HostManager_IsAddinWorkbook(wb) Then Exit Sub
'    HostManager_ActivateWorkbook wb, "WindowActivate"
'End Sub
'
'Private Sub App_WorkbookBeforeClose(ByVal wb As Workbook, Cancel As Boolean)
'    On Error Resume Next
'    HostManager_UnregisterHost wb, "WorkbookBeforeClose"
'End Sub
'
Option Explicit
Public WithEvents App As Application
Attribute App.VB_VarHelpID = -1

Private Sub Class_Initialize()
    Set App = Application
    Debug.Print "[" & Format$(Now, "hh:nn:ss") & "] [CAppEvents] [Init] bound"
End Sub

Private Sub App_WorkbookActivate(ByVal wb As Workbook)
    On Error Resume Next
    If wb Is Nothing Then Exit Sub
    If HostManager_IsAddinWorkbook(wb) Then Exit Sub
    HostManager_ActivateWorkbook wb, "WorkbookActivate"
End Sub

Private Sub App_SheetActivate(ByVal sh As Object)
    On Error Resume Next
    If Not TypeOf sh Is Worksheet Then Exit Sub
    Dim wb As Workbook: Set wb = sh.parent
    If HostManager_IsAddinWorkbook(wb) Then Exit Sub
    HostManager_ActivateSheet wb, sh, "SheetActivate"
End Sub

Private Sub App_WindowActivate(ByVal wb As Workbook, ByVal wn As Window)
    On Error Resume Next
    If wb Is Nothing Then Exit Sub
    If HostManager_IsAddinWorkbook(wb) Then Exit Sub
    ' Only refresh ribbon UI, no version check (avoids duplicate triggers)
    HostManager_RefreshRibbonOnly wb, "WindowActivate"
End Sub

Private Sub App_WorkbookBeforeClose(ByVal wb As Workbook, Cancel As Boolean)
    On Error Resume Next
    HostManager_UnregisterHost wb, "WorkbookBeforeClose"
End Sub

