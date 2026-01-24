Attribute VB_Name = "Import"
Option Explicit
Private gExcelApp As Excel.Application

'--- Limits
Private Const MAX_ROWS As Long = 1048570
Private Const MAX_COLS As Long = 16380

'==========================
' PUBLIC ENTRY
'==========================

Public Sub RunImportForSheet(sheetName As String, sourcePath As String, destAddress As String)
    On Error GoTo Fail

    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Sheets(sheetName)
    
    Debug.Print "Started RunImportForSheet"
    
    '---------------------------------------------
    ' 1. Resolve source path (open file dialog if blank or a folder)
    '---------------------------------------------
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Len(Trim$(sourcePath)) = 0 Or (fso.FolderExists(sourcePath) And Not fso.fileExists(sourcePath)) Then
        Dim initDir As String
        If fso.FolderExists(sourcePath) Then
            initDir = sourcePath
        ElseIf fso.fileExists(sourcePath) Then
            initDir = fso.GetParentFolderName(sourcePath)
        Else
            initDir = CurDir$
        End If
    
        With Application.FileDialog(msoFileDialogFilePicker)
            .AllowMultiSelect = False
            .Title = "Select file to import"
            .Filters.Clear
            .Filters.Add "Supported files", "*.csv;*.txt;*.tsv;*.xls;*.xlsx;*.xlsm;*.xlsb;*.ods"
            .Filters.Add "All files", "*.*"
            .InitialFileName = initDir & Application.PathSeparator
            If .Show = -1 Then
                sourcePath = .SelectedItems(1)
            Else
                MsgBox "Import cancelled.", vbInformation
                Exit Sub
            End If
        End With
    End If
    
    Debug.Print "Started 3. Resolve destination properly"

    
    '---------------------------------------------
    ' 2. Load data into memory
    '---------------------------------------------
    Dim data As Variant, formats As Variant
    data = LoadDataArray(sourcePath, formats)
    If IsEmpty(data) Then Exit Sub

    ' Inline Variant2DSize
    Dim nRows As Long, nCols As Long
    On Error Resume Next
    nRows = UBound(data, 1) - LBound(data, 1) + 1
    nCols = UBound(data, 2) - LBound(data, 2) + 1
    On Error GoTo Fail
    If nRows < 1 Or nCols < 1 Then
        MsgBox "Invalid or empty data array.", vbExclamation
        Exit Sub
    End If
    
    Debug.Print "Started 3. Resolve destination properly"
    
    '---------------------------------------------
    ' 3. Resolve destination properly
    '---------------------------------------------
    Dim dest As Range
    On Error Resume Next
    If Len(Trim$(destAddress)) > 0 Then
        Debug.Print "destAddress raw: [" & destAddress & "]"
        Set dest = ws.Range(destAddress)
        Debug.Print "Attempted to set dest to ws.Range(" & destAddress & ")"
    End If
    If dest Is Nothing Then
        Debug.Print "dest is Nothing after assignment."
    Else
        Debug.Print "dest resolved to address: " & dest.Address
    End If
    On Error GoTo Fail

    ' NEW: Capture destination format before clearing
    Dim savedFormatRow As Range
    If Not dest Is Nothing Then
        Set savedFormatRow = CaptureRowFormat(dest)
    End If

    Dim writeRange As Range
    Set writeRange = PrepareOutputRange(dest, nRows, nCols, "Import", True, ws)
    If writeRange Is Nothing Then
        Debug.Print "writeRange is Nothing. Exiting."
        Exit Sub
    Else
        Debug.Print "writeRange resolved to address: " & writeRange.Address
    End If




    '---------------------------------------------
    ' 4. Write data and optional formats
    '---------------------------------------------
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    With writeRange
        .value = data
        If Not IsEmpty(formats) Then
            On Error Resume Next
            .numberFormat = formats
            On Error GoTo 0
        End If
    End With

    ' NEW: Apply saved format and clear excess range
    If Not savedFormatRow Is Nothing Then
        ApplyFormatToRange savedFormatRow, writeRange
    End If
    If Not dest Is Nothing Then
        ClearExcessRange dest, nRows
    End If

CLEANUP:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    MsgBox "Import complete: " & nRows & " rows × " & nCols & " columns.", vbInformation
    Exit Sub

Fail:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    MsgBox "Import failed: " & Err.Description, vbCritical
End Sub


'==========================
' CORE LOADERS
'==========================
Private Function LoadDataArray(filePath As String, Optional ByRef numberFormats As Variant) As Variant
    Dim ext As String
    ext = LCase$(CreateObject("Scripting.FileSystemObject").GetExtensionName(filePath))

    Select Case ext
        Case "csv", "txt", "tsv"
            LoadDataArray = ReadCSVToArray(filePath, ext)
        Case "xls", "xlsx", "xlsm", "xlsb", "ods"
            LoadDataArray = ReadExcelToArray(filePath, numberFormats)
        Case Else
            MsgBox "Unsupported file type: " & ext, vbExclamation
            LoadDataArray = Empty
    End Select
End Function

'==========================
' CSV READER
'==========================
Private Function ReadCSVToArray(filePath As String, Optional ext As String = "csv") As Variant
    On Error GoTo Fail

    Dim fso As Object, ts As Object, content As String
    Dim lines() As String, fields() As String
    Dim i As Long, j As Long, rowCount As Long, maxCols As Long
    Dim data As Variant, delim As String

    Select Case ext
        Case "tsv": delim = vbTab
        Case Else:  delim = ","
    End Select

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(filePath, 1, False, -2)
    content = ts.ReadAll: ts.Close

    lines = Split(content, vbNewLine)
    NormalizeTrailingBlank lines
    delim = DetectDelimiter(lines, delim)

    rowCount = UBound(lines) - LBound(lines) + 1
    maxCols = 0
    For i = LBound(lines) To UBound(lines)
        fields = FastCSVParse(lines(i), delim)
        If UBound(fields) + 1 > maxCols Then maxCols = UBound(fields) + 1
    Next i

    If rowCount = 0 Or maxCols = 0 Then Exit Function
    If rowCount > MAX_ROWS Or maxCols > MAX_COLS Then Exit Function

    ReDim data(1 To rowCount, 1 To maxCols)
    For i = 0 To rowCount - 1
        fields = FastCSVParse(lines(i + LBound(lines)), delim)
        For j = 0 To UBound(fields)
            data(i + 1, j + 1) = fields(j)
        Next j
    Next i

    ReadCSVToArray = data
    Exit Function

Fail:
    MsgBox "CSV import failed: " & Err.Description, vbCritical
    ReadCSVToArray = Empty
End Function

Private Sub NormalizeTrailingBlank(ByRef arr() As String)
    If UBound(arr) >= 0 Then
        If Len(Trim$(arr(UBound(arr)))) = 0 Then ReDim Preserve arr(LBound(arr) To UBound(arr) - 1)
    End If
End Sub

Private Function DetectDelimiter(ByRef lines() As String, Optional defaultDelim As String = ",") As String
    Dim s As String, c1&, c2&
    If UBound(lines) < 0 Then DetectDelimiter = defaultDelim: Exit Function
    s = lines(LBound(lines))
    If LCase$(Left$(s, 4)) = "sep=" Then
        DetectDelimiter = Mid$(s, 5, 1)
    Else
        c1 = CountChar(s, ","): c2 = CountChar(s, ";")
        If c1 >= c2 Then DetectDelimiter = "," Else DetectDelimiter = ";"
    End If
End Function

Private Function CountChar(s As String, ch As String) As Long
    CountChar = Len(s) - Len(Replace$(s, ch, ""))
End Function

Private Function FastCSVParse(line As String, delim As String) As String()
    Dim res() As String, token As String
    Dim i&, n&, inQuotes As Boolean, ch As String * 1

    If Len(line) = 0 Then ReDim res(0 To 0): res(0) = "": FastCSVParse = res: Exit Function
    ReDim res(0 To 15): n = 0

    For i = 1 To Len(line)
        ch = Mid$(line, i, 1)
        Select Case ch
            Case """"
                If inQuotes And i < Len(line) And Mid$(line, i + 1, 1) = """" Then
                    token = token & """": i = i + 1
                Else
                    inQuotes = Not inQuotes
                End If
            Case delim
                If inQuotes Then
                    token = token & ch
                Else
                    res(n) = token: token = "": n = n + 1
                    If n > UBound(res) Then ReDim Preserve res(0 To n * 2)
                End If
            Case Else
                token = token & ch
        End Select
    Next i

    res(n) = token
    ReDim Preserve res(0 To n)
    FastCSVParse = res
End Function

'==========================
' EXCEL READER
'==========================


Private Function ReadExcelToArray(filePath As String, Optional ByRef numberFormats As Variant) As Variant
    On Error GoTo Fail

    Dim ext As String
    ext = LCase$(CreateObject("Scripting.FileSystemObject").GetExtensionName(filePath))

    Select Case ext
'        Case "xls", "xlsx", "xlsm"
'            ReadExcelToArray = ReadExcel_ADO(filePath)
        Case "xlsb", "ods", "xls", "xlsx", "xlsm"
            ReadExcelToArray = ReadExcel_COM(filePath, numberFormats)
        Case Else
            MsgBox "Unsupported file type: " & ext, vbExclamation
            ReadExcelToArray = Empty
    End Select
    Exit Function

Fail:
    MsgBox "Excel import failed: " & Err.Description, vbCritical
    ReadExcelToArray = Empty
End Function


Private Function PickSheetName(wb As Workbook) As String
    On Error GoTo Fail

    Dim picker As New SheetPickerForm
    picker.Show vbModal
    PickSheetName = picker.SelectedSheet
    Unload picker
    Exit Function

Fail:
    Debug.Print "Error in PickSheetName:", Err.Number, Err.Description
    PickSheetName = ""
End Function

Private Function ReadExcel_ADO(filePath As String) As Variant
    On Error GoTo Fail

    Dim cn As Object, rs As Object, sheetNames As Collection
    Dim picker As SheetPickerForm, sName As String
    Dim sql As String, data As Variant
    Dim t0 As Double

    Debug.Print "=== ReadExcel_ADO START ==="
    Debug.Print "File:", filePath
    t0 = Timer

    ' Build connection
    Set cn = CreateObject("ADODB.Connection")
    Debug.Print "Opening ADO connection..."
    cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;" & _
        "Data Source=" & filePath & ";" & _
        "Extended Properties=""Excel 12.0;HDR=NO;IMEX=1"";"
    Debug.Print "Connection open in", Format(Timer - t0, "0.000"), "sec"

    ' Get sheet list
    Set sheetNames = New Collection
    Dim schema As Object, i As Long
    Set schema = cn.OpenSchema(20)
    Do Until schema.EOF
        Dim nm As String
        nm = schema.fields("TABLE_NAME").value
        Debug.Print "Found raw table name:", nm
        If right$(nm, 1) = "$" Or right$(nm, 2) = "$'" Then
            nm = Replace(Replace(nm, "$", ""), "'", "")
            sheetNames.Add nm
        End If
        schema.MoveNext
    Loop
    Debug.Print "Sheet count found:", sheetNames.count

    ' Choose sheet
    If sheetNames.count > 1 Then
        Set picker = New SheetPickerForm
        Dim s As Variant
        For Each s In sheetNames
            picker.cmbSheets.AddItem s
        Next
        picker.cmbSheets.ListIndex = 0
        picker.Show vbModal
        sName = picker.SelectedSheet
        Debug.Print "User picked sheet:", sName
        Unload picker
    Else
        sName = sheetNames(1)
        Debug.Print "Single sheet:", sName
    End If

    If Len(sName) = 0 Then
        Debug.Print "No sheet selected. Exiting CLEANUP."
        GoTo CLEANUP
    End If

    ' Read data
    sql = "SELECT * FROM [" & sName & "$]"
    Debug.Print "SQL:", sql
    Set rs = cn.Execute(sql)
    Debug.Print "Recordset opened. EOF:", rs.EOF

    If Not rs.EOF Then
        Dim raw As Variant, outArr As Variant
        raw = rs.GetRows()
        Debug.Print "GetRows done. LBound(raw,1):", LBound(raw, 1), "UBound(raw,1):", UBound(raw, 1)
        Debug.Print "LBound(raw,2):", LBound(raw, 2), "UBound(raw,2):", UBound(raw, 2)

        Dim fields As Long, recs As Long, r As Long, c As Long
        fields = UBound(raw, 1) - LBound(raw, 1) + 1
        recs = UBound(raw, 2) - LBound(raw, 2) + 1
        Debug.Print "fields:", fields, "records:", recs

        If recs > 0 And fields > 0 Then
            ReDim outArr(1 To recs, 1 To fields)
            For r = 1 To recs
                For c = 1 To fields
                    outArr(r, c) = raw(LBound(raw, 1) + c - 1, LBound(raw, 2) + r - 1)
                Next c
            Next r
            ReadExcel_ADO = outArr
            Debug.Print "Data array built:", recs, "rows x", fields, "cols"
        Else
            Debug.Print "Empty recordset."
            ReadExcel_ADO = Empty
        End If
    Else
        Debug.Print "EOF True – No data."
        ReadExcel_ADO = Empty
    End If

CLEANUP:
    On Error Resume Next
    rs.Close
    cn.Close
    Debug.Print "=== ReadExcel_ADO END ==="
    Exit Function

Fail:
    Debug.Print "Error in ReadExcel_ADO:", Err.Number, Err.Description
    ReadExcel_ADO = Empty
End Function


Private Function ReadExcel_COM(filePath As String, Optional ByRef numberFormats As Variant) As Variant
    On Error GoTo Fail

    Dim srcWB As Workbook
    Dim ws As Worksheet
    Dim sheetName As String
    Dim ur As Range

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    ' open in current instance, not a new Excel.Application
    Set srcWB = Workbooks.Open(filePath, ReadOnly:=True)

    ' pick a valid sheet name from that workbook
    sheetName = PickSheetName(srcWB)
    If Len(sheetName) = 0 Then GoTo CLEANUP

    On Error Resume Next
    Set ws = srcWB.Sheets(sheetName)
    On Error GoTo Fail
    If ws Is Nothing Then
        MsgBox "Sheet not found: " & sheetName, vbExclamation
        GoTo CLEANUP
    End If

    Set ur = ws.UsedRange
    If ur Is Nothing Then GoTo CLEANUP

    ReadExcel_COM = ur.value
    numberFormats = ur.numberFormat

CLEANUP:
    srcWB.Close SaveChanges:=False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Exit Function

Fail:
    On Error Resume Next
    If Not srcWB Is Nothing Then srcWB.Close SaveChanges:=False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    MsgBox "Excel import failed: " & Err.Description, vbCritical
    ReadExcel_COM = Empty
End Function




'==========================
' FILE PICKER + UTILITY
'==========================
Private Function PickFile(initial As String) As String
    With Application.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Supported files", "*.csv;*.txt;*.tsv;*.xls;*.xlsx;*.xlsm;*.xlsb;*.ods"
        .Title = "Select file to import"
        If Len(initial) > 0 Then .InitialFileName = initial
        If .Show = -1 Then PickFile = .SelectedItems(1)
    End With
End Function

Private Function UBound2D(arr As Variant, dimIndex As Long) As Long
    If IsEmpty(arr) Then UBound2D = 0: Exit Function
    UBound2D = UBound(arr, dimIndex) - LBound(arr, dimIndex) + 1
End Function

Public Function Variant2DSize(ByRef v As Variant, ByRef rows As Long, ByRef cols As Long) As Boolean
    On Error GoTo Bad
    rows = UBound(v, 1) - LBound(v, 1) + 1
    cols = UBound(v, 2) - LBound(v, 2) + 1
    Variant2DSize = True
    Exit Function
Bad:
    rows = 0: cols = 0
    Variant2DSize = False
End Function



