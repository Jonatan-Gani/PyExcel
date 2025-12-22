Attribute VB_Name = "Export"
'=== modExportCore ===
Option Explicit

Public Sub RunExportForSheet(sheetName As String, sourceRef As String, outputPath As String)
    On Error GoTo Fail

    Debug.Print "Started Export"
    Debug.Print "Sheet: " & sheetName
    Debug.Print "txtExportInput: " & sourceRef
    Debug.Print "txtExportOutput: " & outputPath

    ' Pass inputs directly into the form (no sheet writes)
    frmExportWizard.InitializeFromInputs sourceRef, outputPath
    frmExportWizard.Show
    Exit Sub

Fail:
    MsgBox "Export failed: " & Err.Description, vbCritical
End Sub



Public Sub ExportRangeToCSV(rng As Range, filePath As String)
    Dim rowValues As Variant, output As String
    Dim r As Long, c As Long, line As String

    rowValues = rng.value
    If Not IsArray(rowValues) Then rowValues = Array(Array(rowValues))

    For r = 1 To UBound(rowValues, 1)
        line = ""
        For c = 1 To UBound(rowValues, 2)
            Dim cellVal As String
            cellVal = CStr(rowValues(r, c))
            If InStr(cellVal, ",") > 0 Or InStr(cellVal, """") > 0 Or InStr(cellVal, vbLf) > 0 Then
                cellVal = """" & Replace(cellVal, """", """""") & """"
            End If
            line = line & IIf(c = 1, "", ",") & cellVal
        Next c
        output = output & line & vbCrLf
    Next r

    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText output
    stream.SaveToFile filePath, 2
    stream.Close
End Sub


Public Sub ExportSingleSheetToExcel(rng As Range, filePath As String, ext As String)
    Dim newWb As Workbook
    Set newWb = Workbooks.Add(1)

    With newWb.Sheets(1)
        .name = "ExportedData"
        .Range("A1").Resize(rng.rows.count, rng.Columns.count).value = rng.value
        .Range("A1").Resize(rng.rows.count, rng.Columns.count).numberFormat = rng.numberFormat
    End With

    newWb.SaveAs fileName:=filePath, FileFormat:=GetExcelFileFormat(ext)
    newWb.Close False
End Sub


Private Function GetExcelFileFormat(ext As String) As Long
    Select Case LCase(ext)
        Case "xlsm": GetExcelFileFormat = 52
        Case "xlsb": GetExcelFileFormat = 50
        Case Else:   GetExcelFileFormat = 51
    End Select
End Function

