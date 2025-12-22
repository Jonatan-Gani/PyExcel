Attribute VB_Name = "Paste"
'=== modPasteEngine ===
Option Explicit

Public Sub PasteFromClipboardToSheet(sheetName As String, targetAddress As String)
    On Error GoTo Fail

    Dim wsTarget As Worksheet
    Set wsTarget = ActiveWorkbook.Sheets(sheetName)

    If wsTarget Is Nothing Then
        MsgBox "Target sheet not found: " & sheetName, vbCritical
        Exit Sub
    End If

    If Len(targetAddress) = 0 Then
        MsgBox "No target address provided.", vbExclamation
        Exit Sub
    End If

    Dim targetRange As Range
    On Error Resume Next
    Set targetRange = wsTarget.Range(targetAddress)
    On Error GoTo Fail
    If targetRange Is Nothing Then
        MsgBox "Invalid target address: " & targetAddress, vbCritical
        Exit Sub
    End If

    Dim wsScratch As Worksheet
    Dim pasted As Boolean
    Dim oldCalc As XlCalculation, oldScr As Boolean, oldEvt As Boolean, oldStatus As Boolean

    ' performance guardrails
    oldCalc = Application.Calculation: Application.Calculation = xlCalculationManual
    oldScr = Application.ScreenUpdating: Application.ScreenUpdating = False
    oldEvt = Application.EnableEvents: Application.EnableEvents = False
    oldStatus = Application.DisplayStatusBar: Application.DisplayStatusBar = True
    Application.StatusBar = "Pasting from clipboard..."

    ' scratch sheet
    Set wsScratch = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
    
    wsScratch.Visible = xlSheetVeryHidden
    wsScratch.Cells.Clear

    ' try native paste
    On Error Resume Next
    wsScratch.Range("A1").PasteSpecial xlPasteAll
    pasted = (Err.Number = 0)
    Err.Clear
    If Not pasted Then
        wsScratch.Range("A1").PasteSpecial xlPasteValues
        pasted = (Err.Number = 0)
        Err.Clear
    End If
    On Error GoTo Fail

    ' fallback to text
    If Not pasted Or IsEmpty(wsScratch.Range("A1").Value2) Then
        Dim dobj As Object, clipText As String
        Set dobj = CreateObject("MSForms.DataObject")
        On Error Resume Next
        dobj.GetFromClipboard
        dobj.GetText clipText
        On Error GoTo Fail

        If LenB(clipText) = 0 Then
            Err.Raise vbObjectError + 513, , "Clipboard empty or not readable."
        End If

        wsScratch.Range("A1").Value2 = clipText
        Set dobj = Nothing
    End If

    Dim ur As Range
    Set ur = wsScratch.UsedRange

    ' simple text-to-columns
    If ur.rows.count = 1 And ur.Columns.count = 1 Then
        Dim s As String: s = CStr(ur.Value2)
        If InStr(s, vbTab) > 0 Or InStr(s, ",") > 0 Or InStr(s, ";") > 0 Or InStr(s, "|") > 0 Then
            Dim cComma As Long, cSemi As Long, cPipe As Long
            cComma = CountOccurrences(s, ",")
            cSemi = CountOccurrences(s, ";")
            cPipe = CountOccurrences(s, "|")

            ur.Value2 = Replace$(Replace$(Replace$(s, vbCrLf, vbLf), vbCr, vbLf), vbLf, vbLf)
            ur.TextToColumns Destination:=ur.Cells(1, 1), DataType:=xlDelimited, _
                TextQualifier:=xlTextQualifierDoubleQuote, ConsecutiveDelimiter:=False, _
                Tab:=(InStr(s, vbTab) > 0), Semicolon:=(cSemi > cComma And cSemi > cPipe), _
                Comma:=(cComma >= cSemi And cComma >= cPipe), Space:=False, _
                Other:=(cPipe > cComma And cPipe > cSemi), OtherChar:="|"
        End If
    End If

    ' write back
    Dim outRange As Range
    Set ur = wsScratch.UsedRange
    Set outRange = wsTarget.Range(targetRange.Address).Resize(ur.rows.count, ur.Columns.count)
    outRange.Value2 = ur.Value2
    
    wsTarget.Activate
    
    MsgBox "Pasted " & ur.rows.count & " rows × " & ur.Columns.count & " columns into " & _
           wsTarget.name & "!" & vbCrLf & "Start cell: " & targetRange.Address(False, False), vbInformation
    
CLEANUP:
    On Error Resume Next
    Application.DisplayAlerts = False
    If Not wsScratch Is Nothing Then wsScratch.Delete
    Application.DisplayAlerts = True
    Application.StatusBar = False
    Application.Calculation = oldCalc
    Application.ScreenUpdating = oldScr
    Application.EnableEvents = oldEvt
    Application.DisplayStatusBar = oldStatus
    Exit Sub

Fail:
    MsgBox "Paste failed: " & Err.Description, vbCritical
    Resume CLEANUP
End Sub

Private Function CountOccurrences(ByVal s As String, ByVal token As String) As Long
    If LenB(s) = 0 Or LenB(token) = 0 Then Exit Function
    CountOccurrences = (Len(s) - Len(Replace$(s, token, vbNullString))) \ Len(token)
End Function


