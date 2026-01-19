Attribute VB_Name = "PathUtils"
Option Explicit

Public Function ResolveProjectPath() As String
    On Error GoTo Fail

    Dim wbPath As String: wbPath = ActiveWorkbook.path
    Debug.Print "[Resolve] Workbook path:", wbPath

    ' Case 1: Local path
    If wbPath Like "[A-Za-z]:\*" Then
        Debug.Print "[Resolve] Detected local path."
        Call EnsureFolderExists(wbPath)
        ResolveProjectPath = wbPath
        Exit Function
    End If

    ' Case 2: SharePoint/OneDrive path
    If wbPath Like "https://*.sharepoint.com/*" Then
        Dim oneDriveRoot As String
        oneDriveRoot = Environ$("OneDriveCommercial")
        Debug.Print "[Resolve] OneDrive root:", oneDriveRoot

        If oneDriveRoot = "" Then
            Debug.Print "[ERROR] OneDrive not syncing."
            ResolveProjectPath = ""
            Exit Function
        End If

        Dim p As Long: p = InStrRev(wbPath, "/Documents", -1, vbTextCompare)
        If p = 0 Then
            Debug.Print "[ERROR] '/Documents' not found."
            ResolveProjectPath = ""
            Exit Function
        End If

        Dim afterDocs As String
        afterDocs = Mid$(wbPath, p + 10)
        afterDocs = Replace(DecodeUrlComponent(afterDocs), "/", "\")

        If Left(afterDocs, 1) = "\" Then afterDocs = Mid$(afterDocs, 2)
        If Not oneDriveRoot Like "*\Documents" Then oneDriveRoot = oneDriveRoot & "\Documents"

        Dim fullPath As String
        fullPath = oneDriveRoot & "\" & afterDocs
        Debug.Print "[Resolve] Resolved SharePoint path:", fullPath

        Call EnsureFolderExists(fullPath)
        ResolveProjectPath = fullPath
        Exit Function
    End If

    Debug.Print "[ERROR] Path not recognized."
    ResolveProjectPath = ""
    Exit Function

Fail:
    Debug.Print "[ERROR] Failed to resolve path:", Err.Description
    ResolveProjectPath = ""
End Function

Public Function EnsureFolderPath(base As String, subFolder As String) As String
    Dim path As String
    path = base & Application.PathSeparator & subFolder
    Call EnsureFolderExists(path)
    EnsureFolderPath = path
End Function

Public Sub EnsureFolderExists(path As String)
    If Dir(path, vbDirectory) = "" Then
        Debug.Print "[Ensure] Creating folder:", path
        On Error Resume Next
        MkDirRecursive path
        If Err.Number <> 0 Then Debug.Print "[ERROR] MkDir failed:", Err.Description
        On Error GoTo 0
    Else
        Debug.Print "[Ensure] Folder exists:", path
    End If
End Sub

Private Sub MkDirRecursive(ByVal fullPath As String)
    Dim parts() As String: parts = Split(fullPath, "\")
    Dim testPath As String: testPath = parts(0)
    Dim i As Long

    If right(testPath, 1) <> ":" Then testPath = testPath & "\"

    For i = 1 To UBound(parts)
        If Len(parts(i)) > 0 Then
            testPath = testPath & parts(i)
            If Dir(testPath, vbDirectory) = "" Then MkDir testPath
            testPath = testPath & "\"
        End If
    Next i
End Sub

Private Function DecodeUrlComponent(ByVal s As String) As String
    Dim i As Long, code As String, result As String
    i = 1
    Do While i <= Len(s)
        If Mid$(s, i, 1) = "%" And i + 2 <= Len(s) Then
            code = Mid$(s, i + 1, 2)
            If code Like "[0-9A-Fa-f][0-9A-Fa-f]" Then
                result = result & Chr$(CLng("&H" & code))
                i = i + 3
            Else
                result = result & Mid$(s, i, 1)
                i = i + 1
            End If
        Else
            result = result & Mid$(s, i, 1)
            i = i + 1
        End If
    Loop
    DecodeUrlComponent = result
End Function


