Attribute VB_Name = "modDst"
Option Explicit

Public Function PrepareOutputRange( _
    ByVal dest As Range, _
    Optional ByVal outRows As Long = -1, _
    Optional ByVal outCols As Long = -1, _
    Optional ByVal opName As String = "Write Output", _
    Optional ByVal promptOnSpill As Boolean = True, _
    Optional ByVal ws As Worksheet _
) As Range

    On Error GoTo Fail

    Debug.Print String(80, "-")
    Debug.Print "PrepareOutputRange: START"

    ' Worksheet context
    If ws Is Nothing Then
        Set ws = ActiveSheet
        Debug.Print "  ws: Nothing (defaulted to ActiveSheet)"
    Else
        Debug.Print "  ws: " & ws.name
    End If

    ' Destination
    If dest Is Nothing Then
        Debug.Print "  dest: Nothing -> prompting user"
        Set dest = ResolveDestinationRange(ws, opName)
        If dest Is Nothing Then
            Debug.Print "  User cancelled selection"
            Exit Function
        End If
    Else
        On Error Resume Next
        Debug.Print "  dest: " & dest.Address
        On Error GoTo 0
    End If

    ' Basic validation
    If dest Is Nothing Then
        Debug.Print "  ERROR: Destination unresolved"
        MsgBox "No destination range selected.", vbCritical, opName
        Exit Function
    End If

    If dest.Areas.count > 1 Then
        Debug.Print "  Error: multi-area range"
        MsgBox "Destination must be a single, contiguous range.", vbCritical, opName
        Exit Function
    End If

    If dest.MergeCells Then
        Debug.Print "  Error: merged cells"
        MsgBox "Destination contains merged cells. Unmerge or choose a different range.", vbCritical, opName
        Exit Function
    End If

    ' Defaults
    If outRows = -1 Then outRows = dest.rows.count
    If outCols = -1 Then outCols = dest.Columns.count
    Debug.Print "  outRows: " & outRows & "  outCols: " & outCols

    If outRows < 1 Or outCols < 1 Then
        Debug.Print "  Invalid output size"
        MsgBox "Invalid output size (" & outRows & "x" & outCols & ").", vbCritical, opName
        Exit Function
    End If

    ' Sheet bounds
    Dim wsDest As Worksheet
    Set wsDest = dest.Worksheet
    If wsDest Is Nothing Then
        Debug.Print "  ERROR: dest.Worksheet is Nothing"
        MsgBox "Destination has no parent worksheet.", vbCritical, opName
        Exit Function
    End If
    Debug.Print "  wsDest: " & wsDest.name

    Dim bottom As Long, right As Long
    bottom = dest.Row + outRows - 1
    right = dest.Column + outCols - 1
    Debug.Print "  bottom: " & bottom & "  right: " & right

    If bottom > wsDest.rows.count Or right > wsDest.Columns.count Then
        Debug.Print "  Output exceeds sheet limits"
        MsgBox "Output (" & outRows & "x" & outCols & ") will not fit on sheet """ & wsDest.name & _
               """ starting at " & dest.Cells(1, 1).Address(0, 0) & ".", vbCritical, opName
        Exit Function
    End If

    ' Final output range
    Dim writeRange As Range
    Set writeRange = dest.Resize(outRows, outCols)
    On Error Resume Next
    Debug.Print "  writeRange: " & writeRange.Address(External:=True)
    On Error GoTo 0

    ' Spill confirmation
    Dim spills As Boolean
    spills = (outRows > dest.rows.count) Or (outCols > dest.Columns.count)
    Debug.Print "  spills: " & spills

    If spills And promptOnSpill Then
        Dim msg As String
        msg = "Output size: " & outRows & " rows x " & outCols & " columns." & vbCrLf & vbCrLf & _
              "Destination: " & dest.rows.count & " rows x " & dest.Columns.count & " columns at " & dest.Address(External:=True) & "." & vbCrLf & vbCrLf & _
              "Continuing will write into a larger range: " & writeRange.Address(External:=True) & "." & vbCrLf & _
              "Proceed?"
        If MsgBox(msg, vbQuestion + vbYesNo, opName) <> vbYes Then Exit Function
    End If

    ' Clear target contents and formatting
    Debug.Print "  Clearing contents and formatting"
    With writeRange
        If WorksheetFunction.CountA(.Cells) > 0 Then .ClearContents

        On Error Resume Next
        .Hyperlinks.Delete
        .ClearHyperlinks
        On Error GoTo 0

        .Borders.lineStyle = xlNone
        .Interior.Pattern = xlNone
        .Interior.TintAndShade = 0

        With .Font
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Bold = False
            .Italic = False
            .Underline = xlUnderlineStyleNone
            .Strikethrough = False
            .Subscript = False
            .Superscript = False
            .name = Application.StandardFont
            .Size = Application.StandardFontSize
        End With

        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .Orientation = 0
        .WrapText = False
        .ShrinkToFit = False
        .AddIndent = False
        .IndentLevel = 0
        .ReadingOrder = xlContext
    End With

    Debug.Print "PrepareOutputRange: COMPLETED OK"
    Set PrepareOutputRange = writeRange
    Exit Function

Fail:
    Debug.Print "PrepareOutputRange: FAILED - " & Err.Number & " | " & Err.Description
    MsgBox "PrepareOutputRange failed: " & Err.Description, vbCritical, opName
End Function

Private Function ResolveDestinationRange( _
    ByVal ws As Worksheet, _
    ByVal opName As String _
) As Range
    Debug.Print String(60, "-")
    Debug.Print "ResolveDestinationRange: START (" & ws.name & ")"
    Debug.Print "  Time: " & Format(Now, "hh:nn:ss")

    On Error Resume Next
    Debug.Print "  ScreenUpdating: " & Application.ScreenUpdating
    Debug.Print "  DisplayAlerts: " & Application.DisplayAlerts
    Debug.Print "  Interactive: " & Application.Interactive
    Debug.Print "  EnableEvents: " & Application.EnableEvents
    Debug.Print "  WindowState: " & Application.WindowState
    Debug.Print "  Visible: " & Application.Visible
    Debug.Print "  Hwnd: " & Application.hwnd
    On Error GoTo 0

    On Error Resume Next
    Debug.Print "  ws.Visible: " & ws.Visible
    Debug.Print "  ws.Activate"
    ws.Activate
    Debug.Print "  ActiveSheet after Activate: " & ActiveSheet.name
    On Error GoTo 0

    Debug.Print "  >> Before InputBox"
    Debug.Print "  >> Application.Ready: " & Application.Ready
    Debug.Print "  >> Focus expected on Excel window."

    Dim t0 As Single
    t0 = Timer
    Dim r As Range
    On Error Resume Next

    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.Interactive = True

    ws.Activate
    DoEvents

    '--- FIX: reattach Excel window focus when called from Ribbon ---
    Application.CommandBars.ExecuteMso "ActivateWindow"

    Set r = Application.InputBox( _
        Prompt:="Select destination range", _
        Title:=opName & " - Destination", _
        Type:=8)
    Debug.Print "  >> InputBox elapsed (sec): " & Format(Timer - t0, "0.00")
    On Error GoTo 0

    Debug.Print "  << After InputBox"
    Debug.Print "  ScreenUpdating: " & Application.ScreenUpdating
    Debug.Print "  DisplayAlerts: " & Application.DisplayAlerts
    Debug.Print "  Interactive: " & Application.Interactive
    Debug.Print "  ActiveWindow: " & IIf(Application.ActiveWindow Is Nothing, "Nothing", "OK")
    Debug.Print "  ActiveSheet: " & ActiveSheet.name

    If r Is Nothing Then
        Debug.Print "  User cancelled or invalid selection"
        GoTo clean_exit
    End If

    If Not r.Worksheet Is ws Then
        Debug.Print "  Invalid sheet: user selected on another sheet (" & r.Worksheet.name & ")"
        MsgBox "Please select a range on sheet """ & ws.name & """ only.", vbExclamation, opName
        GoTo clean_exit
    End If

    Debug.Print "  User selected: " & r.Address(External:=True)
    Set ResolveDestinationRange = r
    Debug.Print "ResolveDestinationRange: SUCCESS"

clean_exit:
    Debug.Print "ResolveDestinationRange: END " & Format(Now, "hh:nn:ss")
    Debug.Print String(60, "-")
End Function


