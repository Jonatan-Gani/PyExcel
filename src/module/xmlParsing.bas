Attribute VB_Name = "xmlParsing"
Option Explicit

'Public Function SerializeRangeToTypedXML(rngRef As String) As String
'    On Error GoTo Fail
'    Debug.Print "=== BEGIN SerializeRangeToTypedXML ==="
'    Debug.Print "Raw input: [" & rngRef & "]"
'
'    If Trim$(rngRef) = "" Then
'        Debug.Print "No range string provided ? prompting user with InputBox (Type:=8)"
'        Dim userSel As Range
'        On Error Resume Next
'        Set userSel = Application.InputBox("Select one or more input ranges", "Select Input", Type:=8)
'        On Error GoTo Fail
'        If userSel Is Nothing Then
'            Debug.Print "User cancelled range selection or no range selected."
'            GoTo Fail
'        End If
'        rngRef = userSel.Worksheet.name & "!" & userSel.Address
'        Debug.Print "User selected: " & rngRef
'    End If
'
'    Dim rangeParts() As String
'    rangeParts = Split(rngRef, ";")
'    Debug.Print "Split into " & (UBound(rangeParts) - LBound(rangeParts) + 1) & " range part(s)."
'
'    Dim i As Long
'    For i = LBound(rangeParts) To UBound(rangeParts)
'        Dim currentRange As Range
'        Set currentRange = Application.Range(Trim$(rangeParts(i)))
'
'        Dim lastCell As Range
'        Set lastCell = currentRange.Find(What:="*", After:=currentRange.Cells(1, 1), LookIn:=xlValues, LookAt:=xlPart, _
'                                         SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False)
'
'        If Not lastCell Is Nothing Then
'            Dim trimmedRange As Range
'            Set trimmedRange = currentRange.Worksheet.Range( _
'                currentRange.Cells(1, 1), _
'                currentRange.Worksheet.Cells(lastCell.Row, currentRange.Column + currentRange.Columns.count - 1))
'
'            If trimmedRange.rows.count < currentRange.rows.count Then
'                Debug.Print "Trimmed range " & (i + 1) & " from " & currentRange.rows.count & " to " & trimmedRange.rows.count & " rows"
'            End If
'
'            rangeParts(i) = trimmedRange.Worksheet.name & "!" & trimmedRange.Address
'        Else
'            Debug.Print "Range " & (i + 1) & " is completely empty"
'            rangeParts(i) = currentRange.Worksheet.name & "!" & currentRange.Address
'        End If
'    Next i
'
'    Dim folderPath As String
'    folderPath = ResolveProjectPath() & "\Temp"
'    Debug.Print "Temp folder: " & folderPath
'
'    Dim fso As Object
'    Set fso = CreateObject("Scripting.FileSystemObject")
'    If Not fso.FolderExists(folderPath) Then
'        Debug.Print "Creating Temp folder..."
'        fso.CreateFolder folderPath
'    End If
'
'    Dim tempFile As String
'    tempFile = folderPath & "\input.xml"
'    Debug.Print "Creating XML file at: " & tempFile
'
'    Dim stream As Object
'    Set stream = CreateObject("ADODB.Stream")
'    stream.Type = 2
'    stream.Charset = "utf-8"
'    stream.Open
'
'    stream.WriteText "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf
'    stream.WriteText "<data>" & vbCrLf
'
'    For i = LBound(rangeParts) To UBound(rangeParts)
'        Dim part As String: part = Trim$(rangeParts(i))
'        Debug.Print "--- Serializing part " & (i + 1) & ": [" & part & "] ---"
'        If part = "" Then GoTo Fail
'
'        Dim ws As Worksheet, refRange As Range
'        If InStr(1, part, "!") > 0 Then
'            Dim sheetName As String, rangeText As String
'            sheetName = Left$(part, InStr(1, part, "!") - 1)
'            rangeText = Mid$(part, InStr(1, part, "!") + 1)
'            Set ws = Nothing
'            On Error Resume Next
'            Set ws = Worksheets(sheetName)
'            On Error GoTo Fail
'            If ws Is Nothing Then
'                Debug.Print "  Sheet [" & sheetName & "] not found."
'                GoTo Fail
'            End If
'            Set refRange = ws.Range(rangeText)
'        Else
'            Set ws = ActiveSheet
'            Set refRange = ws.Range(part)
'        End If
'
'        If refRange Is Nothing Then
'            Debug.Print "  ERROR: Failed to resolve range: [" & part & "]"
'            GoTo Fail
'        End If
'
'        Debug.Print "  Resolved range: " & refRange.Address(External:=True)
'        Debug.Print "  Areas.Count: " & refRange.Areas.count
'        Debug.Print "  Rows: " & refRange.rows.count & ", Columns: " & refRange.Columns.count
'
'        ' ===== IMPORTANT: use .Value to preserve vbDate =====
'        Dim data As Variant
'        data = refRange.value
'
'        Dim numCols As Long: numCols = refRange.Columns.count
'        Dim numRows As Long: numRows = refRange.rows.count
'        If numRows < 2 Then
'            Debug.Print "  Skipping: range must have at least 2 rows (header + data)"
'            GoTo Fail
'        End If
'
'        Dim headerRow() As String, typeRow() As String
'        ReDim headerRow(1 To numCols)
'        ReDim typeRow(1 To numCols)
'
'        Dim r As Long, c As Long
'        For c = 1 To numCols
'            headerRow(c) = CStr(data(1, c))
'        Next
'
'        ' ==== correct 2-arg call ====
'        For c = 1 To numCols
'            typeRow(c) = InferColumnType(data, c)
'            If LenB(typeRow(c)) = 0 Then typeRow(c) = "string"
'            Debug.Print "    Column " & c & ": " & headerRow(c) & " [" & typeRow(c) & "]"
'        Next
'
'        ' Precompute column type flags
'        Dim isDateCol() As Boolean, isBoolCol() As Boolean, isNumCol() As Boolean, wantsInt() As Boolean
'        ReDim isDateCol(1 To numCols)
'        ReDim isBoolCol(1 To numCols)
'        ReDim isNumCol(1 To numCols)
'        ReDim wantsInt(1 To numCols)
'        For c = 1 To numCols
'            Select Case LCase$(typeRow(c))
'                Case "date":  isDateCol(c) = True
'                Case "bool":  isBoolCol(c) = True
'                Case "int":   isNumCol(c) = True: wantsInt(c) = True
'                Case "float": isNumCol(c) = True
'            End Select
'        Next
'
'        stream.WriteText "<table name=""df" & (i + 1) & """>" & vbCrLf
'        stream.WriteText "<columns>" & vbCrLf
'        For c = 1 To numCols
'            stream.WriteText "<col name=""" & EscapeXml(headerRow(c)) & """ type=""" & typeRow(c) & """ />" & vbCrLf
'        Next
'        stream.WriteText "</columns><rows>" & vbCrLf
'
'        Dim colXml() As String
'        ReDim colXml(1 To numCols)
'
'        ' You can chunk rows if you want even more speed; this writes per-row.
'        For r = 2 To numRows
'            For c = 1 To numCols
'                Dim v As Variant: v = data(r, c)
'                If IsEmpty(v) Then
'                    colXml(c) = "<col/>"
'                ElseIf isDateCol(c) Then
'                    ' v is vbDate here because we loaded with .Value
'                    colXml(c) = "<col>" & Format$(v, "yyyy-mm-dd\THH:nn:ss") & "Z</col>"
'                ElseIf isBoolCol(c) Then
'                    colXml(c) = "<col>" & LCase$(CStr(CBool(v))) & "</col>"
'                ElseIf isNumCol(c) Then
'                    If wantsInt(c) Then
'                        colXml(c) = "<col>" & CStr(v) & "</col>"
'                    Else
'                        colXml(c) = "<col>" & EscapeXml(Format$(CDbl(v), "0.############################")) & "</col>"
'                    End If
'                Else
'                    colXml(c) = "<col>" & EscapeXml(CStr(v)) & "</col>"
'                End If
'            Next
'            stream.WriteText "<row>" & Join(colXml, "") & "</row>" & vbCrLf
'        Next
'
'        stream.WriteText "</rows></table>" & vbCrLf
'    Next
'
'    stream.WriteText "</data>"
'    stream.SaveToFile tempFile, 2
'    stream.Close
'
'    Debug.Print "XML serialization complete."
'    Debug.Print "File written to: " & tempFile
'    Debug.Print "=== END SerializeRangeToTypedXML (Success) ==="
'    SerializeRangeToTypedXML = tempFile
'    Exit Function
'
'Fail:
'    Debug.Print "=== ERROR in SerializeRangeToTypedXML ==="
'    Debug.Print "Error: " & Err.Description
'    On Error Resume Next
'    If Not stream Is Nothing Then If stream.State = 1 Then stream.Close
'    SerializeRangeToTypedXML = ""
'End Function
Public Function SerializeRangeToTypedXML(rngRef As String) As String
    On Error GoTo Fail
    Debug.Print "=== BEGIN SerializeRangeToTypedXML ==="
    Debug.Print "Raw input: [" & rngRef & "]"

    ' ============================================================
    ' If the user passed nothing, prompt with InputBox as before
    ' ============================================================
    If Trim$(rngRef) = "" Then
        Debug.Print "No range string provided ? prompting user with InputBox (Type:=8)"
        Dim userSel As Range
        On Error Resume Next
        Set userSel = Application.InputBox("Select one or more input ranges", "Select Input", Type:=8)
        On Error GoTo Fail
        If userSel Is Nothing Then
            Debug.Print "User cancelled range selection or no range selected."
            GoTo Fail
        End If
        rngRef = userSel.Worksheet.name & "!" & userSel.Address
        Debug.Print "User selected: " & rngRef
    End If

    ' ============================================================
    ' Split parts and detect optional variable names: {name}=range
    ' ============================================================
    Dim rawParts() As String
    rawParts = Split(rngRef, ";")

    Dim partNames() As String
    Dim partRanges() As String
    ReDim partNames(LBound(rawParts) To UBound(rawParts))
    ReDim partRanges(LBound(rawParts) To UBound(rawParts))
    
    Dim varName As String
    
    Dim i As Long
    For i = LBound(rawParts) To UBound(rawParts)
        Dim token As String: token = Trim$(rawParts(i))

        If Left$(token, 1) = "{" Then
            Dim closePos As Long: closePos = InStr(2, token, "}")
            If closePos > 0 Then
'                Dim varName As String: varName = Mid$(token, 2, closePos - 2)
                varName = Mid$(token, 2, closePos - 2)
                Dim eqPos As Long: eqPos = InStr(closePos + 1, token, "=")
                If eqPos > 0 Then
                    partNames(i) = Trim$(varName)
                    partRanges(i) = Trim$(Mid$(token, eqPos + 1))
                Else
                    partNames(i) = ""
                    partRanges(i) = token
                End If
            Else
                partNames(i) = ""
                partRanges(i) = token
            End If
        Else
            partNames(i) = ""
            partRanges(i) = token
        End If
    Next i

    Debug.Print "Detected " & (UBound(partRanges) - LBound(partRanges) + 1) & " input part(s)."

    ' ============================================================
    ' Original trimming logic per range (unchanged)
    ' ============================================================
    For i = LBound(partRanges) To UBound(partRanges)
        Dim currentRange As Range
        Set currentRange = Application.Range(Trim$(partRanges(i)))

        Dim lastCell As Range
        Set lastCell = currentRange.Find(What:="*", After:=currentRange.Cells(1, 1), LookIn:=xlValues, LookAt:=xlPart, _
                                         SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False)

        If Not lastCell Is Nothing Then
            Dim trimmedRange As Range
            Set trimmedRange = currentRange.Worksheet.Range( _
                currentRange.Cells(1, 1), _
                currentRange.Worksheet.Cells(lastCell.Row, currentRange.Column + currentRange.Columns.count - 1))

            If trimmedRange.rows.count < currentRange.rows.count Then
                Debug.Print "Trimmed range " & (i + 1) & " from " & currentRange.rows.count & " to " & trimmedRange.rows.count & " rows"
            End If

            partRanges(i) = trimmedRange.Worksheet.name & "!" & trimmedRange.Address
        Else
            Debug.Print "Range " & (i + 1) & " is completely empty"
            partRanges(i) = currentRange.Worksheet.name & "!" & currentRange.Address
        End If
    Next i

    ' ============================================================
    ' Prepare output XML
    ' ============================================================
    Dim folderPath As String
    folderPath = ResolveProjectPath() & "\Temp"

    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then fso.CreateFolder folderPath

    Dim tempFile As String
    tempFile = folderPath & "\input.xml"

    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open

    stream.WriteText "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf
    stream.WriteText "<data>" & vbCrLf

    ' ============================================================
    ' Process each input unit
    ' ============================================================
    Dim c As Long
    For i = LBound(partRanges) To UBound(partRanges)
        Dim part As String: part = Trim$(partRanges(i))
        If part = "" Then GoTo Fail

        Debug.Print "--- Serializing part " & (i + 1) & ": [" & part & "] ---"

        Dim ws As Worksheet, refRange As Range
        If InStr(1, part, "!") > 0 Then
            Dim sheetName As String, rangeText As String
            sheetName = Left$(part, InStr(1, part, "!") - 1)
            rangeText = Mid$(part, InStr(1, part, "!") + 1)

            Set ws = Worksheets(sheetName)
            Set refRange = ws.Range(rangeText)
        Else
            Set ws = ActiveSheet
            Set refRange = ws.Range(part)
        End If

        If refRange Is Nothing Then GoTo Fail

        Dim data As Variant
        data = refRange.value

        Dim numRows As Long: numRows = refRange.rows.count
        Dim numCols As Long: numCols = refRange.Columns.count

        Dim varBName As String
        varBName = Trim$(partNames(i))

        ' ============================================================
        ' Determine type: scalar / list / dataframe
        ' ============================================================

        ' -------------------------
        ' SCALAR (1x1)
        ' -------------------------
        If numRows = 1 And numCols = 1 Then
            If varBName = "" Then varBName = "value" & (i + 1)

            Dim scalarVal As Variant: scalarVal = data(1, 1)
            Dim dt As String

            If IsDate(scalarVal) Then
                dt = "timestamp"
                scalarVal = Format$(scalarVal, "yyyy-mm-dd\THH:nn:ss") & "Z"
            ElseIf IsNumeric(scalarVal) Then
                If CLng(scalarVal) = CDbl(scalarVal) Then
                    dt = "int"
                Else
                    dt = "decimal"
                End If
            ElseIf VarType(scalarVal) = vbBoolean Then
                dt = "bool"
                scalarVal = LCase$(CStr(CBool(scalarVal)))
            Else
                dt = "string"
                scalarVal = EscapeXml(CStr(scalarVal))
            End If

            stream.WriteText "<value name=""" & varBName & """ datatype=""" & dt & """>" & scalarVal & "</value>" & vbCrLf
            GoTo NextPart
        End If

        ' -------------------------
        ' LIST (1xN or Nx1)
        ' -------------------------
        If (numRows = 1 And numCols > 1) Or (numCols = 1 And numRows > 1) Then
            If varBName = "" Then varBName = "list" & (i + 1)

            stream.WriteText "<list name=""" & varBName & """>" & vbCrLf
            Dim r As Long

            If numRows = 1 Then
                For c = 1 To numCols
                    stream.WriteText "  <item>" & EscapeXml(CStr(data(1, c))) & "</item>" & vbCrLf
                Next
            Else
                For r = 1 To numRows
                    stream.WriteText "  <item>" & EscapeXml(CStr(data(r, 1))) & "</item>" & vbCrLf
                Next
            End If

            stream.WriteText "</list>" & vbCrLf
            GoTo NextPart
        End If

        ' -------------------------
        ' DATAFRAME (existing code)
        ' -------------------------
        If varBName = "" Then varBName = "df" & (i + 1)

        ' Original header/type logic remains unchanged
        Dim headerRow() As String, typeRow() As String
        ReDim headerRow(1 To numCols)
        ReDim typeRow(1 To numCols)

'        Dim c As Long
        For c = 1 To numCols
            headerRow(c) = CStr(data(1, c))
        Next

        For c = 1 To numCols
            typeRow(c) = InferColumnType(data, c)
            If LenB(typeRow(c)) = 0 Then typeRow(c) = "string"
        Next

        Dim isDateCol() As Boolean, isBoolCol() As Boolean, isNumCol() As Boolean, wantsInt() As Boolean
        ReDim isDateCol(1 To numCols), isBoolCol(1 To numCols)
        ReDim isNumCol(1 To numCols), wantsInt(1 To numCols)

        For c = 1 To numCols
            Select Case LCase$(typeRow(c))
                Case "date":  isDateCol(c) = True
                Case "bool":  isBoolCol(c) = True
                Case "int":   isNumCol(c) = True: wantsInt(c) = True
                Case "float": isNumCol(c) = True
            End Select
        Next

        stream.WriteText "<table name=""" & varBName & """>" & vbCrLf
        stream.WriteText "<columns>" & vbCrLf

        For c = 1 To numCols
            stream.WriteText "<col name=""" & EscapeXml(headerRow(c)) & """ type=""" & typeRow(c) & """ />" & vbCrLf
        Next

        stream.WriteText "</columns><rows>" & vbCrLf

        Dim colXml() As String: ReDim colXml(1 To numCols)
        Dim r2 As Long

        For r2 = 2 To numRows
            For c = 1 To numCols
                Dim v As Variant: v = data(r2, c)
                If IsEmpty(v) Then
                    colXml(c) = "<col/>"
                ElseIf isDateCol(c) Then
                    colXml(c) = "<col>" & Format$(v, "yyyy-mm-dd\THH:nn:ss") & "Z</col>"
                ElseIf isBoolCol(c) Then
                    colXml(c) = "<col>" & LCase$(CStr(CBool(v))) & "</col>"
                ElseIf isNumCol(c) Then
                    If wantsInt(c) Then
                        colXml(c) = "<col>" & CStr(v) & "</col>"
                    Else
                        colXml(c) = "<col>" & EscapeXml(Format$(CDbl(v), "0.############################")) & "</col>"
                    End If
                Else
                    colXml(c) = "<col>" & EscapeXml(CStr(v)) & "</col>"
                End If
            Next

            stream.WriteText "<row>" & Join(colXml, "") & "</row>" & vbCrLf
        Next

        stream.WriteText "</rows></table>" & vbCrLf

NextPart:
    Next i

    stream.WriteText "</data>"
    stream.SaveToFile tempFile, 2
    stream.Close

    SerializeRangeToTypedXML = tempFile
    Debug.Print "=== END SerializeRangeToTypedXML (Success) ==="
    Exit Function

Fail:
    Debug.Print "=== ERROR in SerializeRangeToTypedXML ==="
    Debug.Print "Error: " & Err.Description
    On Error Resume Next
    If Not stream Is Nothing Then If stream.State = 1 Then stream.Close
    SerializeRangeToTypedXML = ""
End Function



' data: 2-D Variant from Range.Value (row 1 is header)
' colIndex: 1-based column index within data
Private Function InferColumnType(ByRef data As Variant, ByVal colIndex As Long) As String
    Dim lastRow As Long
    lastRow = UBound(data, 1)
    If lastRow < 2 Then
        InferColumnType = "blank"
        Exit Function
    End If

    Dim r As Long
    Dim v As Variant
    Dim vt As VbVarType
    Dim anyNonBlank As Boolean
    
    Dim allDates As Boolean: allDates = True
    Dim allBools As Boolean: allBools = True
    Dim anyString As Boolean
    Dim anyErrorish As Boolean
    Dim anyNumeric As Boolean
    Dim anyFloat As Boolean  ' among numerics
    Dim d As Double

    For r = 2 To lastRow ' assume row 1 is header
        v = data(r, colIndex)
        If IsEmpty(v) Then GoTo nextRow

        anyNonBlank = True
        vt = VarType(v)

        Select Case vt
            Case vbString
                anyString = True
                Exit For  ' string dominates, result decided

            Case vbBoolean
                ' still possible bool-only column
                allDates = False

            Case vbDate
                ' still possible date-only column
                allBools = False

            Case vbByte, vbInteger, vbLong, vbLongLong, vbSingle, vbDouble, vbCurrency, vbDecimal
                allDates = False
                allBools = False
                anyNumeric = True
                d = CDbl(v)
                If Abs(d - Fix(d)) >= 0.0000001 Then anyFloat = True

            Case vbError, vbNull, vbObject, vbArray, vbDataObject, vbVariant
                anyErrorish = True
                Exit For  ' safest to call it string

            Case Else
                anyErrorish = True
                Exit For
        End Select
nextRow:
    Next

    If Not anyNonBlank Then
        InferColumnType = "blank"
        Exit Function
    End If

    ' Safety-first resolution
    If anyString Or anyErrorish Then
        InferColumnType = "string"
    ElseIf allDates And Not anyNumeric And Not allBools Then
        ' All nonblank are vbDate (and not booleans)
        InferColumnType = "timestamp"
    ElseIf allBools And Not anyNumeric And Not allDates Then
        ' All nonblank are vbBoolean (and not dates)
        InferColumnType = "bool"
    ElseIf anyNumeric And Not anyString And Not anyErrorish And Not allDates And Not allBools Then
        ' Pure numeric column (ints and/or floats)
        If anyFloat Then
            InferColumnType = "float"
        Else
            InferColumnType = "int"
        End If
    Else
        ' Mixed types (e.g., dates + numerics, bools + numerics, etc.) => safest is string
        InferColumnType = "string"
    End If
End Function

Private Sub ShuffleArray(arr() As Long)
    Dim i As Long, j As Long, tmp As Long
    For i = UBound(arr) To LBound(arr) + 1 Step -1
        j = Int((i - LBound(arr) + 1) * Rnd) + LBound(arr)
        tmp = arr(i)
        arr(i) = arr(j)
        arr(j) = tmp
    Next
End Sub

Private Function JoinCollection(col As Collection, Optional delimiter As String = "") As String
    Dim arr() As String, i As Long
    ReDim arr(1 To col.count)
    For i = 1 To col.count
        arr(i) = col(i)
    Next
    JoinCollection = Join(arr, delimiter)
End Function

Private Function attr$(node As Object, ByVal name As String)
    ' Null-safe attribute getter. Returns "" if attribute is missing/null.
    Dim v As Variant
    On Error Resume Next
    v = node.GetAttribute(name)
    On Error GoTo 0
    If IsNull(v) Or IsEmpty(v) Then
        attr = ""
    Else
        attr = CStr(v)
    End If
End Function

Private Function SafeText$(v As Variant)
    If IsNull(v) Or IsEmpty(v) Then
        SafeText = ""
    Else
        SafeText = CStr(v)
    End If
End Function




'
'Public Function PasteTypedXMLToRange(xmlString As String, dstRef As String) As Boolean
'    On Error GoTo fail
'
'    Dim xDoc As Object: Set xDoc = CreateObject("MSXML2.DOMDocument.6.0")
'    xDoc.async = False: xDoc.LoadXML xmlString
'
'    Dim tableNodes As Object
'    Set tableNodes = xDoc.SelectNodes("/data/table")
'
'    Dim hasTables As Boolean
'    hasTables = (Not tableNodes Is Nothing) And (tableNodes.Length > 0)
'
'    ' Parse destinations: allow single or multiple refs separated by ; or ,
'    Dim dstList() As String, rawRefs As String
'    rawRefs = Trim$(dstRef)
'    If LenB(rawRefs) > 0 Then
'        rawRefs = Replace(rawRefs, ",", ";")
'        Dim parts() As String, iPart As Long, tmp As Collection
'        parts = Split(rawRefs, ";")
'        Set tmp = New Collection
'        For iPart = LBound(parts) To UBound(parts)
'            Dim p As String: p = Trim$(parts(iPart))
'            If LenB(p) > 0 Then tmp.Add p
'        Next
'        If tmp.Count > 0 Then
'            ReDim dstList(1 To tmp.Count)
'            For iPart = 1 To tmp.Count
'                dstList(iPart) = CStr(tmp(iPart))
'            Next
'        End If
'    End If
'    Dim numDests As Long
'    If (Not Not dstList) <> 0 Then numDests = UBound(dstList) Else numDests = 0
'
'    Dim numTables As Long
'    If hasTables Then
'        numTables = tableNodes.Length
'    Else
'        ' Legacy single table shape
'        numTables = 1
'    End If
'
'    If numTables = 0 Then
'        Debug.Print "No tables found in XML."
'        Exit Function
'    End If
'
'    ' Helpers to paste one table block and return the pasted target range
'    Dim lastTargetRange As Range  ' track last placed block for stacking
'    Dim lastWS As Worksheet
'    Dim lastLeftCol As Long
'    Dim lastBottomRow As Long
'    Dim placedCount As Long: placedCount = 0
'
'    Dim tIdx As Long
'    For tIdx = 1 To numTables
'        Dim tableName As String
'        Dim colNodes As Object, rowNodes As Object
'
'        If hasTables Then
'            Dim tNode As Object
'            Set tNode = tableNodes.Item(tIdx - 1)
'            tableName = attr(tNode, "name")
'
'            Set colNodes = tNode.SelectNodes("columns/col")
'            Set rowNodes = tNode.SelectNodes("rows/row")
'        Else
'            ' Legacy single-table paths
'            tableName = ""
'            Set colNodes = xDoc.SelectNodes("/data/columns/col")
'            Set rowNodes = xDoc.SelectNodes("/data/rows/row")
'        End If
'
'        If colNodes Is Nothing Or rowNodes Is Nothing Then
'            Debug.Print "Missing columns or rows for table #" & tIdx & " (" & tableName & "). Skipping."
'            GoTo next_table
'        End If
'        If colNodes.Length = 0 Or rowNodes.Length = 0 Then
'            Debug.Print "Empty columns or rows for table #" & tIdx & " (" & tableName & "). Skipping."
'            GoTo next_table
'        End If
'
'        Dim numCols As Long: numCols = colNodes.Length
'        Dim numRows As Long: numRows = rowNodes.Length
'
'        ' Extract headers, types, and A1 formula metadata
'        Dim headers() As String, types() As String
'        Dim isA1() As Boolean, a1Formula() As String, a1Anchor() As String
'
'        ReDim headers(1 To numCols)
'        ReDim types(1 To numCols)
'        ReDim isA1(1 To numCols)
'        ReDim a1Formula(1 To numCols)
'        ReDim a1Anchor(1 To numCols)
'
'        Dim c As Long, modeAttr As String
'        For c = 1 To numCols
'            headers(c) = attr(colNodes.Item(c - 1), "name")
'            types(c) = LCase$(attr(colNodes.Item(c - 1), "type"))
'
'            modeAttr = LCase$(Trim$(attr(colNodes.Item(c - 1), "mode")))
'            If modeAttr = "a1" Then
'                isA1(c) = True
'                a1Formula(c) = attr(colNodes.Item(c - 1), "a1")
'                a1Anchor(c) = attr(colNodes.Item(c - 1), "anchor")
'            End If
'        Next
'
'        ' Build output array (including header row)
'        Dim out() As Variant
'        ReDim out(1 To numRows + 1, 1 To numCols)
'
'        ' Headers
'        For c = 1 To numCols
'            out(1, c) = headers(c)
'        Next
'
'        ' Values
'        Dim r As Long, colNodeList As Object, val As String
'        For r = 1 To numRows
'            Set colNodeList = rowNodes.Item(r - 1).SelectNodes("col")
'            For c = 1 To numCols
'                val = ""
'                If c - 1 < colNodeList.Length Then
'                    val = SafeText(colNodeList.Item(c - 1).text)
'                End If
'
'                Select Case types(c)
'                    Case "blank": out(r + 1, c) = ""
'                    Case "int":   If val <> "" Then out(r + 1, c) = CLng(val) Else out(r + 1, c) = 0
'                    Case "float": If val <> "" Then out(r + 1, c) = CDbl(val) Else out(r + 1, c) = 0#
'                    Case "bool":  out(r + 1, c) = (LCase$(val) = "true")
'                    Case "date":  If val <> "" Then out(r + 1, c) = CDate(val) Else out(r + 1, c) = ""
'                    Case Else:    out(r + 1, c) = val   ' includes "string"
'                End Select
'            Next
'        Next
'
'        ' Determine destination for this table
'        Dim wsCtx As Worksheet, destRange As Range, targetRange As Range
'        Set wsCtx = Nothing
'        Set destRange = Nothing
'        Set targetRange = Nothing
'
'        Dim outRows As Long, outCols As Long
'        outRows = UBound(out, 1)
'        outCols = UBound(out, 2)
'
'        If tIdx <= numDests And numDests > 0 Then
'            ' Use corresponding destination
'            On Error Resume Next
'            Set destRange = Range(dstList(tIdx))
'            On Error GoTo fail
'            If Not destRange Is Nothing Then Set wsCtx = destRange.Worksheet
'            If wsCtx Is Nothing Then Set wsCtx = ActiveSheet
'
'            If Not destRange Is Nothing Then
'                Set targetRange = PrepareOutputRange(destRange, outRows, outCols, _
'                                                     "Paste Typed XML", True)
'            Else
'                Set targetRange = PrepareOutputRange(Nothing, outRows, outCols, _
'                                                     "Paste Typed XML", True, wsCtx, "D3")
'            End If
'            If targetRange Is Nothing Then
'                Debug.Print "Destination not resolved or user cancelled for table #" & tIdx & "."
'                GoTo next_table
'            End If
'
'        ElseIf numDests > 0 Then
'            ' Fewer destinations than tables: stack under the last placed block
'            If lastWS Is Nothing Or lastLeftCol = 0 Or lastBottomRow = 0 Then
'                ' If nothing placed yet, fall back to last provided destination reference
'                On Error Resume Next
'                Set destRange = Range(dstList(numDests))
'                On Error GoTo fail
'                If Not destRange Is Nothing Then Set wsCtx = destRange.Worksheet
'                If wsCtx Is Nothing Then Set wsCtx = ActiveSheet
'
'                If Not destRange Is Nothing Then
'                    Set targetRange = PrepareOutputRange(destRange, outRows, outCols, _
'                                                         "Paste Typed XML", True)
'                Else
'                    Set targetRange = PrepareOutputRange(Nothing, outRows, outCols, _
'                                                         "Paste Typed XML", True, wsCtx, "D3")
'                End If
'                If targetRange Is Nothing Then
'                    Debug.Print "Destination not resolved or user cancelled."
'                    GoTo next_table
'                End If
'
'                ' Initialize stacking anchors
'                Set lastTargetRange = targetRange
'                Set lastWS = targetRange.Worksheet
'                lastLeftCol = targetRange.Column
'                lastBottomRow = targetRange.row + targetRange.rows.Count - 1
'
'            Else
'                ' Insert one blank row, one name row, then the table
'                Dim nameRow As Long
'                nameRow = lastBottomRow + 2   ' +1 blank row, then name row
'                With lastWS
'                    ' Write the name (if available)
'                    If LenB(Trim$(tableName)) > 0 Then
'                        .Cells(nameRow, lastLeftCol).Value = tableName
'                    Else
'                        .Cells(nameRow, lastLeftCol).Value = "Table " & tIdx
'                    End If
'                    ' Next block starts one row below the name row
'                    Dim startCell As Range
'                    Set startCell = .Cells(nameRow + 1, lastLeftCol)
'                    Set targetRange = PrepareOutputRange(startCell, outRows, outCols, _
'                                                         "Paste Typed XML", True)
'                End With
'            End If
'
'        Else
'            ' No destination provided at all: default to ActiveSheet:D3
'            Set wsCtx = ActiveSheet
'            Set targetRange = PrepareOutputRange(Nothing, outRows, outCols, _
'                                                 "Paste Typed XML", True, wsCtx, "D3")
'            If targetRange Is Nothing Then
'                Debug.Print "Destination not resolved or user cancelled."
'                GoTo next_table
'            End If
'        End If
'
'        ' First pass: write data
'        targetRange.Value = out
'
'
'        With targetRange.rows(1)
'            .Font.Bold = True
'            .Interior.Pattern = xlSolid
'            .Interior.Color = RGB(143, 215, 225) ' #8FD7E1
'        End With
'
'
'        ' Apply A1-anchored formulas (preserves $ and complex refs)
'        Dim ws As Worksheet: Set ws = targetRange.Worksheet
'        Debug.Print "Applying A1 formulas where specified... [Table #" & tIdx & " " & tableName & "]"
'        For c = 1 To numCols
'            If isA1(c) Then
'                If LenB(a1Formula(c)) = 0 Then
'                    Debug.Print "  [A1] Empty formula for column #" & c & " (" & headers(c) & "). Skipping."
'                Else
'                    Dim fillRng As Range
'                    Set fillRng = targetRange.Cells(2, c).Resize(numRows, 1) ' first data row of this column
'                    On Error Resume Next
'                    fillRng.Formula = a1Formula(c)  ' e.g., "=D10*100"
'                    If Err.Number <> 0 Then
'                        Debug.Print "  [A1] ERROR writing formula for '" & headers(c) & "': " & Err.Description
'                        Err.Clear
'                    Else
'                        Debug.Print "  [A1] Wrote A1 formula for '" & headers(c) & "' from " & _
'                                    fillRng.Cells(1, 1).Address(0, 0) & " down " & numRows & " rows."
'                    End If
'                    On Error GoTo 0
'                End If
'            End If
'        Next
'
'
'
'
'        ' Second pass: convert string cells starting with "=" to FormulaR1C1 (non-A1 columns)
'        Debug.Print "Begin second pass: applying FormulaR1C1 to string-type formulas (non-A1 columns) [Table #" & tIdx & "]"
'        For r = 1 To numRows
'            Set colNodeList = rowNodes.Item(r - 1).SelectNodes("col")
'            For c = 1 To numCols
'                If Not isA1(c) Then
'                    If LCase$(types(c)) = "string" And c - 1 < colNodeList.Length Then
'                        val = SafeText(colNodeList.Item(c - 1).text)
'                        If LenB(val) > 0 And Left$(val, 1) = "=" Then
'                            On Error Resume Next
'                            targetRange.Cells(r + 1, c).FormulaR1C1 = val
'                            If Err.Number <> 0 Then
'                                Debug.Print "  ERROR setting formula in R" & r + 1 & "C" & c & ": " & Err.Description
'                                Err.Clear
'                            End If
'                            On Error GoTo 0
'                        End If
'                    End If
'                End If
'            Next
'        Next
'        Debug.Print "Second pass complete [Table #" & tIdx & "]"
'
'        ' Update stacking anchors
'        Set lastTargetRange = targetRange
'        Set lastWS = targetRange.Worksheet
'        lastLeftCol = targetRange.Column
'        lastBottomRow = targetRange.row + targetRange.rows.Count - 1
'        placedCount = placedCount + 1
'
'next_table:
'    Next tIdx
'
'    ' Notify if more destinations than tables
'    If numDests > numTables Then
'        Dim extra As Long: extra = numDests - numTables
'        Dim msg As String
'        msg = "Note: " & extra & " destination reference(s) were provided but no matching table existed. " & _
'              "Only the first " & numTables & " destination(s) were used."
'        Debug.Print msg
'        On Error Resume Next
'        MsgBox msg, vbInformation, "Paste Typed XML"
'        On Error GoTo 0
'    End If
'
'    PasteTypedXMLToRange = (placedCount > 0)
'    Exit Function
'
'fail:
'    Debug.Print "Error in PasteTypedXMLToRange: " & Err.Description
'    PasteTypedXMLToRange = False
'End Function
'
'
'
'Public Function PasteTypedXMLToRange(xmlString As String, dstRef As String) As Boolean
'    On Error GoTo fail
'
'    ' Load XML document
'    Dim xDoc As Object: Set xDoc = CreateObject("MSXML2.DOMDocument.6.0")
'    xDoc.async = False: xDoc.LoadXML xmlString
'
'    ' Look for <table> nodes
'    Dim tableNodes As Object
'    Set tableNodes = xDoc.SelectNodes("/data/table")
'
'    Dim hasTables As Boolean
'    hasTables = (Not tableNodes Is Nothing) And (tableNodes.Length > 0)
'
'    ' Parse destinations: allow single or multiple refs separated by ; or ,
'    Dim dstList() As String, rawRefs As String
'    rawRefs = Trim$(dstRef)
'    If LenB(rawRefs) > 0 Then
'        rawRefs = Replace(rawRefs, ",", ";")
'        Dim parts() As String, iPart As Long, tmp As Collection
'        parts = Split(rawRefs, ";")
'        Set tmp = New Collection
'        For iPart = LBound(parts) To UBound(parts)
'            Dim p As String: p = Trim$(parts(iPart))
'            If LenB(p) > 0 Then tmp.Add p
'        Next
'        If tmp.count > 0 Then
'            ReDim dstList(1 To tmp.count)
'            For iPart = 1 To tmp.count
'                dstList(iPart) = CStr(tmp(iPart))
'            Next
'        End If
'    End If
'    Dim numDests As Long
'    If (Not Not dstList) <> 0 Then numDests = UBound(dstList) Else numDests = 0
'
'    ' Count number of tables in the XML
'    Dim numTables As Long
'    If hasTables Then
'        numTables = tableNodes.Length
'    Else
'        ' Legacy: single table without explicit <table> node
'        numTables = 1
'    End If
'
'    ' Flag: only label if more than one table
'    Dim needLabel As Boolean
'    needLabel = (numTables > 1)
'
'    If numTables = 0 Then
'        Debug.Print "No tables found in XML."
'        Exit Function
'    End If
'
'    ' Variables to track placement for stacking
'    Dim lastTargetRange As Range
'    Dim lastWS As Worksheet
'    Dim lastLeftCol As Long
'    Dim lastBottomRow As Long
'    Dim placedCount As Long: placedCount = 0
'
'    Dim clearAll As Range
'    Dim totRows As Long
'    Dim totCols As Long
'
'    Dim tIdx As Long
'
'
'
'
'    For tIdx = 1 To numTables
'        Dim tableName As String
'        Dim colNodes As Object, rowNodes As Object
'
'        ' Extract column/row nodes for this table
'        If hasTables Then
'            Dim tNode As Object
'            Set tNode = tableNodes.item(tIdx - 1)
'            tableName = attr(tNode, "name")
'
'            Set colNodes = tNode.SelectNodes("columns/col")
'            Set rowNodes = tNode.SelectNodes("rows/row")
'        Else
'            tableName = ""
'            Set colNodes = xDoc.SelectNodes("/data/columns/col")
'            Set rowNodes = xDoc.SelectNodes("/data/rows/row")
'        End If
'
'        ' Skip malformed/empty tables
'        If colNodes Is Nothing Or rowNodes Is Nothing Then GoTo next_table
'        If colNodes.Length = 0 Or rowNodes.Length = 0 Then GoTo next_table
'
'        Dim numCols As Long: numCols = colNodes.Length
'        Dim numRows As Long: numRows = rowNodes.Length
'
'        ' Extract column headers and typing information
'        Dim headers() As String, types() As String
'        Dim isA1() As Boolean, a1Formula() As String, a1Anchor() As String
'
'        ReDim headers(1 To numCols)
'        ReDim types(1 To numCols)
'        ReDim isA1(1 To numCols)
'        ReDim a1Formula(1 To numCols)
'        ReDim a1Anchor(1 To numCols)
'
'        Dim c As Long, modeAttr As String
'        For c = 1 To numCols
'            headers(c) = attr(colNodes.item(c - 1), "name")
'            types(c) = LCase$(attr(colNodes.item(c - 1), "type"))
'            modeAttr = LCase$(Trim$(attr(colNodes.item(c - 1), "mode")))
'            If modeAttr = "a1" Then
'                isA1(c) = True
'                a1Formula(c) = attr(colNodes.item(c - 1), "a1")
'                a1Anchor(c) = attr(colNodes.item(c - 1), "anchor")
'            End If
'        Next
'
'        ' Build output array including header row
'        Dim out() As Variant
'        ReDim out(1 To numRows + 1, 1 To numCols)
'
'        ' Headers
'        For c = 1 To numCols
'            out(1, c) = headers(c)
'        Next
'
'        ' Values, typed
'        Dim r As Long, colNodeList As Object, val As String
'        For r = 1 To numRows
'            Set colNodeList = rowNodes.item(r - 1).SelectNodes("col")
'            For c = 1 To numCols
'                val = ""
'                If c - 1 < colNodeList.Length Then val = SafeText(colNodeList.item(c - 1).text)
'                Select Case types(c)
'                    Case "blank": out(r + 1, c) = ""
'                    Case "int":   If val <> "" Then out(r + 1, c) = CLng(val) Else out(r + 1, c) = 0
'                    Case "float": If val <> "" Then out(r + 1, c) = CDbl(val) Else out(r + 1, c) = 0#
'                    Case "bool":  out(r + 1, c) = (LCase$(val) = "true")
'                    Case "date":  If val <> "" Then out(r + 1, c) = CDate(val) Else out(r + 1, c) = ""
'                    Case Else:    out(r + 1, c) = val
'                End Select
'            Next
'        Next
'
'        ' Determine destination for this table
'        Dim wsCtx As Worksheet, destRange As Range, targetRange As Range
'        Set wsCtx = Nothing: Set destRange = Nothing: Set targetRange = Nothing
'        Dim outRows As Long, outCols As Long
'        outRows = UBound(out, 1): outCols = UBound(out, 2)
'
'        If tIdx <= numDests And numDests > 0 Then
'            ' Use explicit destination reference
'            On Error Resume Next
'            Set destRange = Range(dstList(tIdx))
'            On Error GoTo fail
'            If Not destRange Is Nothing Then Set wsCtx = destRange.Worksheet
'            If wsCtx Is Nothing Then Set wsCtx = ActiveSheet
'
'            ' If multiple tables ? write label above data
'            If needLabel Then
'                Dim labelCell As Range, startCell As Range
'                If Not destRange Is Nothing Then
'                    Set labelCell = destRange.Cells(1, 1)
'                Else
'                    Set labelCell = wsCtx.Range("D3")
'                End If
'                If LenB(Trim$(tableName)) > 0 Then
'                    labelCell.Value = tableName
'                Else
'                    labelCell.Value = "Table " & tIdx
'                End If
'                Set startCell = labelCell.Offset(1, 0)
'                Set targetRange = PrepareOutputRange(startCell, outRows, outCols, "Paste Typed XML", True)
'            Else
'                ' Single table: no label
'                If Not destRange Is Nothing Then
'                    Set targetRange = PrepareOutputRange(destRange, outRows, outCols, "Paste Typed XML", True)
'                Else
'                    Set targetRange = PrepareOutputRange(Nothing, outRows, outCols, "Paste Typed XML", True, wsCtx, "D3")
'                End If
'            End If
'
'            ' More tables than destinations: stack below previous
'            If lastWS Is Nothing Or lastLeftCol = 0 Or lastBottomRow = 0 Then
'                ' Fallback: use last provided destination
'                On Error Resume Next
'                Set destRange = Range(dstList(numDests))
'                On Error GoTo fail
'                If Not destRange Is Nothing Then Set wsCtx = destRange.Worksheet
'                If wsCtx Is Nothing Then Set wsCtx = ActiveSheet
'                If needLabel Then
'                    ' Label above first stacked table
'                    Dim firstLabel As Range, firstStart As Range
'                    If Not destRange Is Nothing Then
'                        Set firstLabel = destRange.Cells(1, 1)
'                    Else
'                        Set firstLabel = wsCtx.Range("D3")
'                    End If
'                    If LenB(Trim$(tableName)) > 0 Then
'                        firstLabel.Value = tableName
'                    Else
'                        firstLabel.Value = "Table " & tIdx
'                    End If
'                    Set firstStart = firstLabel.Offset(1, 0)
'                    Set targetRange = PrepareOutputRange(firstStart, outRows, outCols, "Paste Typed XML", True)
'                Else
'                    If Not destRange Is Nothing Then
'                        Set targetRange = PrepareOutputRange(destRange, outRows, outCols, "Paste Typed XML", True)
'                    Else
'                        Set targetRange = PrepareOutputRange(Nothing, outRows, outCols, "Paste Typed XML", True, wsCtx, "D3")
'                    End If
'                End If
'
'                ' Track stacking anchor
'                Set lastTargetRange = targetRange
'                Set lastWS = targetRange.Worksheet
'                lastLeftCol = targetRange.Column
'                lastBottomRow = targetRange.Row + targetRange.rows.count - 1
'
'            Else
'                ' Subsequent stacked tables: always label row
'                Dim nameRow As Long
'                nameRow = lastBottomRow + 2   ' +1 blank row + label row
'                With lastWS
'                    If LenB(Trim$(tableName)) > 0 Then
'                        .Cells(nameRow, lastLeftCol).Value = tableName
'                    Else
'                        .Cells(nameRow, lastLeftCol).Value = "Table " & tIdx
'                    End If
'                    Dim startCell2 As Range
'                    Set startCell2 = .Cells(nameRow + 1, lastLeftCol)
'                    Set targetRange = PrepareOutputRange(startCell2, outRows, outCols, "Paste Typed XML", True)
'                End With
'            End If
'
'        Else
'            ' No destination references at all: default to D3
'            Set wsCtx = ActiveSheet
'            If needLabel Then
'                ' Label in D3, data starts at D4
'                wsCtx.Range("D3").Value = IIf(LenB(Trim$(tableName)) > 0, tableName, "Table " & tIdx)
'                Set targetRange = PrepareOutputRange(wsCtx.Range("D4"), outRows, outCols, "Paste Typed XML", True)
'            Else
'                Set targetRange = PrepareOutputRange(Nothing, outRows, outCols, "Paste Typed XML", True, wsCtx, "D3")
'            End If
'        End If
'
'        If targetRange Is Nothing Then GoTo next_table
'
'        ' Paste output array
'        targetRange.Value = out
'
'        ' Header formatting
'        With targetRange.rows(1)
'            .Font.Bold = True
'            .Interior.pattern = xlSolid
'            .Interior.Color = RGB(143, 215, 225)
'            .Borders(xlEdgeBottom).LineStyle = xlContinuous
'            .Borders(xlEdgeBottom).Weight = xlThin
'            .Borders(xlEdgeBottom).Color = RGB(0, 0, 0)
'        End With
'
'        ' Apply A1 anchored formulas
'        Dim ws As Worksheet: Set ws = targetRange.Worksheet
'        For c = 1 To numCols
'            If isA1(c) Then
'                If LenB(a1Formula(c)) > 0 Then
'                    Dim fillRng As Range
'                    Set fillRng = targetRange.Cells(2, c).Resize(numRows, 1)
'                    On Error Resume Next
'                    fillRng.Formula = a1Formula(c)
'                    On Error GoTo 0
'                End If
'            End If
'        Next
'
'        ' Second pass: convert strings beginning with "=" to R1C1 formulas
'        For r = 1 To numRows
'            Set colNodeList = rowNodes.item(r - 1).SelectNodes("col")
'            For c = 1 To numCols
'                If Not isA1(c) Then
'                    If LCase$(types(c)) = "string" And c - 1 < colNodeList.Length Then
'                        val = SafeText(colNodeList.item(c - 1).text)
'                        If LenB(val) > 0 And Left$(val, 1) = "=" Then
'                            On Error Resume Next
'                            targetRange.Cells(r + 1, c).FormulaR1C1 = val
'                            On Error GoTo 0
'                        End If
'                    End If
'                End If
'            Next
'        Next
'
'        ' Update stacking anchors
'        Set lastTargetRange = targetRange
'        Set lastWS = targetRange.Worksheet
'        lastLeftCol = targetRange.Column
'        lastBottomRow = targetRange.Row + targetRange.rows.count - 1
'        placedCount = placedCount + 1
'
'next_table:
'    Next tIdx
'
'    ' Warn if more destinations were supplied than tables
'    If numDests > numTables Then
'        Dim extra As Long: extra = numDests - numTables
'        Dim msg As String
'        msg = "Note: " & extra & " destination reference(s) were provided but no matching table existed. " & _
'              "Only the first " & numTables & " destination(s) were used."
'        Debug.Print msg
'        On Error Resume Next
'        MsgBox msg, vbInformation, "Paste Typed XML"
'        On Error GoTo 0
'    End If
'
'    PasteTypedXMLToRange = (placedCount > 0)
'    Exit Function
'
'fail:
'    Debug.Print "Error in PasteTypedXMLToRange: " & Err.Description
'    PasteTypedXMLToRange = False
'End Function
'
'

Private Function SafeAttr(n As Object, ByVal name As String) As String
    On Error Resume Next
    Dim v As Variant: v = n.GetAttribute(name)
    If IsError(v) Or IsNull(v) Then
        SafeAttr = vbNullString
    Else
        SafeAttr = CStr(v)
    End If
End Function


Public Function PasteTypedXMLToRange(xmlString As String, dstRef As String) As Boolean
    On Error GoTo Fail

    '=== Disable Excel overhead for speed ===
    Dim oldCalc As XlCalculation
    Dim oldScreen As Boolean, oldEvents As Boolean
    oldCalc = Application.Calculation
    oldScreen = Application.ScreenUpdating
    oldEvents = Application.EnableEvents
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    '=== Load XML ===
    Dim xDoc As Object: Set xDoc = CreateObject("MSXML2.DOMDocument.6.0")
    xDoc.async = False
    xDoc.LoadXML xmlString
    Debug.Print "XML Loaded: "; Left$(xmlString, 200) & "..."

    Dim tableNodes As Object
    Set tableNodes = xDoc.SelectNodes("/data/table")
    Dim hasTables As Boolean: hasTables = (Not tableNodes Is Nothing) And (tableNodes.Length > 0)
    Debug.Print "Tables found: "; IIf(hasTables, tableNodes.Length, 0)

    Dim numTables As Long
    If hasTables Then numTables = tableNodes.Length Else numTables = 1
    If numTables = 0 Then GoTo clean_exit



    
    
    
    Dim wsDst As Worksheet, dstRange As Range
    ' Fail hard if no destination provided
    If LenB(Trim$(dstRef)) = 0 Then
        Err.Raise vbObjectError + 700, "Paste Typed XML", "dstRef is required and cannot be empty."
    End If
    
    ' Resolve destination; invalid references will raise runtime error 1004 naturally
    Set dstRange = Range(dstRef)
    Set wsDst = dstRange.Worksheet
    
    Debug.Print "Destination: "; wsDst.name, dstRange.Address
    
    Dim needLabel As Boolean: needLabel = (numTables > 1)
    Dim nextRow As Long, anchorCol As Long, placedCount As Long
    
    ' Pre-prepare the full destination once (pure-range mode: clears exactly dstRange, no resize)
    Dim preparedAnchor As Range
    Set preparedAnchor = PrepareOutputRange(dstRange, , , "Paste Typed XML", False)
    
    ' Anchor position for first write
    nextRow = preparedAnchor.Row
    anchorCol = preparedAnchor.Column
    placedCount = 0
    
    Dim tIdx As Long
    For tIdx = 1 To numTables
        '=== Extract column/row nodes ===
        Dim tableName As String, tNode As Object
        Dim colNodes As Object, rowNodes As Object
        If hasTables Then
            Set tNode = tableNodes.Item(tIdx - 1)
            tableName = SafeAttr(tNode, "name")
            Set colNodes = tNode.SelectNodes("columns/col")
            Set rowNodes = tNode.SelectNodes("rows/row")
        Else
            tableName = ""
            Set colNodes = xDoc.SelectNodes("/data/columns/col")
            Set rowNodes = xDoc.SelectNodes("/data/rows/row")
        End If
        If colNodes Is Nothing Or rowNodes Is Nothing Then GoTo next_table
        If colNodes.Length = 0 Or rowNodes.Length = 0 Then GoTo next_table
    
        Debug.Print "Table[" & tIdx & "] Name=" & tableName & ", Cols=" & colNodes.Length & ", Rows=" & rowNodes.Length
    
        '=== Headers and types ===
        Dim numCols As Long: numCols = colNodes.Length
        Dim numRows As Long: numRows = rowNodes.Length
        Dim headers() As String, types() As String
        ReDim headers(1 To numCols), types(1 To numCols)
    
        Dim isA1() As Boolean, a1Formula() As String
        ReDim isA1(1 To numCols), a1Formula(1 To numCols)
    
        Dim c As Long
        For c = 1 To numCols
            headers(c) = SafeAttr(colNodes.Item(c - 1), "name")
            types(c) = LCase$(SafeAttr(colNodes.Item(c - 1), "type"))
'            Debug.Print "  Col[" & c & "] Name=" & headers(c) & ", Type=" & types(c)
    
            Dim modeAttr As String: modeAttr = LCase$(SafeAttr(colNodes.Item(c - 1), "mode"))
            Dim a1Text As String: a1Text = SafeAttr(colNodes.Item(c - 1), "a1")
            isA1(c) = (modeAttr = "a1") Or (LenB(a1Text) > 0)
    
            If LenB(a1Text) > 0 Then
                a1Formula(c) = a1Text
            Else
                a1Formula(c) = SafeAttr(colNodes.Item(c - 1), "a1Formula")
            End If
        Next
    
        '=== Build output array (fast) ===
        Dim out() As Variant: ReDim out(1 To numRows + 1, 1 To numCols)
        For c = 1 To numCols: out(1, c) = headers(c): Next
    
        Dim allCols As Object
        If hasTables Then
            Set allCols = tNode.SelectNodes("rows/row/col")
        Else
            Set allCols = xDoc.SelectNodes("/data/rows/row/col")
        End If
    
        Dim r As Long, val As Variant, idx As Long
        idx = 0
    
        For r = 1 To numRows
            For c = 1 To numCols
                val = ""
                If Not allCols Is Nothing Then
                    If idx < allCols.Length Then val = allCols.Item(idx).nodeTypedValue
                End If
    '            Debug.Print "    R=" & r & ", C=" & c & ", RawVal='" & val & "', TargetType=" & types(c)
    
                Select Case types(c)
                    Case "blank"
                        out(r + 1, c) = ""
                    Case "int"
                        If LenB(val) > 0 And IsNumeric(val) Then out(r + 1, c) = CLng(val) Else out(r + 1, c) = 0
                    Case "float"
                        If LenB(val) > 0 And IsNumeric(val) Then out(r + 1, c) = CDbl(val) Else out(r + 1, c) = 0#
                    Case "bool"
                        out(r + 1, c) = (LCase$(CStr(val)) = "true")
                    Case "date"
                        If LenB(val) > 0 And IsDate(val) Then out(r + 1, c) = CDate(val) Else out(r + 1, c) = ""
                    Case Else
                        out(r + 1, c) = val
                End Select
    
                idx = idx + 1
            Next
        Next
    
        '=== Label if stacked ===
        If needLabel Then
            wsDst.Cells(nextRow, anchorCol).value = IIf(LenB(Trim$(tableName)) > 0, tableName, "Table " & tIdx)
            nextRow = nextRow + 1
        End If
    
        '=== Paste ===
        Dim targetRange As Range
        Set targetRange = PrepareOutputRange( _
            wsDst.Cells(nextRow, anchorCol), _
            UBound(out, 1), _
            UBound(out, 2), _
            "Paste Typed XML", _
            False) ' no prompts; raises on spill/bounds
    
        targetRange.value = out
    
        placedCount = placedCount + 1
    
        With targetRange.rows(1)
            .Font.Bold = True
            .Interior.Pattern = xlSolid
            .Interior.Color = RGB(143, 215, 225)
            .Borders(xlEdgeBottom).lineStyle = xlContinuous
            .Borders(xlEdgeBottom).Weight = xlThin
            .Borders(xlEdgeBottom).Color = RGB(0, 0, 0)
        End With
    
        '=== Apply A1 formulas exactly as provided ===
        Dim ws As Worksheet: Set ws = targetRange.Worksheet
        For c = 1 To numCols
            If isA1(c) Then
                If LenB(a1Formula(c)) > 0 Then
                    Dim fillRng As Range
                    Set fillRng = targetRange.Cells(2, c).Resize(numRows, 1)
                    On Error Resume Next
                    fillRng.Formula = a1Formula(c)   ' A1 only
                    If Err.Number <> 0 Then
                        Debug.Print "Formula write failed in col " & c & ": " & a1Formula(c) & " (" & Err.Description & ")"
                        Err.Clear
                    End If
                    On Error GoTo 0
                End If
            End If
        Next
    
        '=== Second pass: inline "=" values become formulas ===
        Dim colNodeList As Object
        For r = 1 To numRows
            Set colNodeList = rowNodes.Item(r - 1).SelectNodes("col")
            For c = 1 To numCols
                If Not isA1(c) Then
                    If LCase$(types(c)) = "string" And c - 1 < colNodeList.Length Then
                        Dim sval As String: sval = SafeText(colNodeList.Item(c - 1).text)
                        If LenB(sval) > 0 And Left$(sval, 1) = "=" Then
                            On Error Resume Next
                            targetRange.Cells(r + 1, c).Formula = sval
                            If Err.Number <> 0 Then
                                Debug.Print "Inline formula write failed R" & r & "C" & c & ": " & sval & " (" & Err.Description & ")"
                                Err.Clear
                            End If
                            On Error GoTo 0
                        End If
                    End If
                End If
            Next
        Next
    
        '=== Advance anchor: exactly once + optional single spacer between tables ===
        nextRow = nextRow + UBound(out, 1)            ' move past the just-written table
        If tIdx < numTables Then nextRow = nextRow + 1 ' add exactly one blank row between tables
    
next_table:
    Next tIdx


clean_exit:
    '=== Restore Excel settings ===
    Application.Calculation = oldCalc
    Application.ScreenUpdating = oldScreen
    Application.EnableEvents = oldEvents

    PasteTypedXMLToRange = (placedCount > 0)
    Exit Function

Fail:
    Debug.Print "Error in PasteTypedXMLToRange: " & Err.Description & " (" & Err.Number & ")"
    PasteTypedXMLToRange = False
    Resume clean_exit
End Function






