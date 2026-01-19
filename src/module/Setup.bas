Attribute VB_Name = "Setup"
'Option Explicit
'
'' Constant for the Workbook-level Named Range used to store the project root path.
'Private Const WORKBOOK_PATH_NAME As String = "ProjectRootPath"
'
'Private Const EMBED_SHEET_NAME As String = "EmbeddedStore"
'Private Const COL_FILENAME As Long = 1
'Private Const COL_CHUNKINDEX As Long = 2
'Private Const COL_BASE64 As Long = 3
'Private Const COL_RELPATH As Long = 4
'
'
'' ==============================================================
'' PROJECT PATH HELPERS
'' ==============================================================
'
'Public Sub SaveProjectPath(wb As Workbook, path As String)
'    On Error GoTo EH
'    Debug.Print "[SaveProjectPath] Saving path '" & path & "' to Workbook name: " & WORKBOOK_PATH_NAME
'    modRibbon.SaveWorkbookValue wb, WORKBOOK_PATH_NAME, path
'    Exit Sub
'EH:
'    Debug.Print "[SaveProjectPath][ERROR] " & Err.Description
'End Sub
'
'
'' ==============================================================
'' CORE INSTALLATION LOGIC (WITH REAL STEP 5 IMPLEMENTATION)
'' ==============================================================
'
'Public Function PyExcelSetup() As Boolean
'    On Error GoTo EH
'
'    Dim wb As Workbook
'    Dim fso As Object
'    Dim hostPath As String
'
'    Debug.Print "============================================================"
'    Debug.Print "[PyExcelSetup] Initialization started."
'
'    Set wb = HostManager_GetCurrentWorkbook()
'    If wb Is Nothing Then GoTo Failed
'
'    Set fso = CreateObject("Scripting.FileSystemObject")
'
'    ' 1. Set main path
'    hostPath = SelectAndSetupRootPath(wb, fso)
'    If Len(hostPath) = 0 Then GoTo Failed
'
'    ' 2. Build project folder tree
'    BuildProjectDirectories fso, hostPath
'
'    ' 3. Save workbook as XLSM
'    SaveHostAsXLSM wb, hostPath
'
'    ' 4. Python VENV
'    If Not CreatePythonVenv(fso, hostPath) Then GoTo Failed
'
'    ' 5. EXTRACT EMBEDDED RESOURCES (IMPLEMENTED)
'    ExtractResources fso, hostPath
'
'    ' 6 + 7 placeholders
'    install_pip_Packages
''    PyExcelSetup_Step7_SaveVersionID wb
'
'    Debug.Print "[PyExcelSetup] Installation finished."
'    Debug.Print "============================================================"
'
'    PyExcelSetup = True
'    Exit Function
'
'Failed:
'    Debug.Print "[PyExcelSetup] Installation terminated due to error/cancellation."
'    PyExcelSetup = False
'    Exit Function
'
'EH:
'    Debug.Print "[PyExcelSetup][FATAL ERROR] " & Err.Description
'    Resume Failed
'End Function
'
'
'' ==============================================================
'' STEP 1 – PATH SELECTOR
'' ==============================================================
'
'Public Function SelectAndSetupRootPath(wb As Workbook, fso As Object) As String
'    On Error GoTo EH
'
'    Dim fldr As Object
'    Dim defaultPath As String
'    Dim pathChosen As String
'
'    If Len(wb.path) > 0 Then
'        defaultPath = wb.path
'    Else
'        defaultPath = Environ("USERPROFILE")
'    End If
'
'    Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
'    With fldr
'        .Title = "Select Project Root Folder"
'        .InitialFileName = defaultPath
'        If .Show <> -1 Then Exit Function
'        pathChosen = .SelectedItems(1)
'    End With
'
'    Dim projectName As String
'    projectName = Left$(wb.name, InStrRev(wb.name, ".") - 1)
'
'    Dim finalPath As String
'    finalPath = pathChosen & Application.PathSeparator & projectName
'
'    If Not fso.FolderExists(finalPath) Then fso.CreateFolder finalPath
'
'    SaveProjectPath wb, finalPath
'    SelectAndSetupRootPath = finalPath
'    Exit Function
'
'EH:
'    Debug.Print "[SelectRootPath][ERROR] " & Err.Description
'End Function
'
'
'' ==============================================================
'' STEP 2 – BUILD FOLDER TREE
'' ==============================================================
'
'Public Sub BuildProjectDirectories(fso As Object, rootPath As String)
'    On Error GoTo EH
'
'    Dim folder As Variant
'    Dim subFolders As Variant: subFolders = Array("AddIn", "Archive", "Python", "Temp", "userScripts")
'
'    For Each folder In subFolders
'        Dim full As String: full = rootPath & Application.PathSeparator & folder
'        If Not fso.FolderExists(full) Then fso.CreateFolder full
'    Next folder
'
'    Dim venvPath As String
'    venvPath = rootPath & "\Python\.venv"
'    If Not fso.FolderExists(venvPath) Then fso.CreateFolder venvPath
'
'    Dim subSub As Variant
'    Dim subSubFolders As Variant: subSubFolders = Array("assets", "lists", "tables", "values")
'
'    For Each subSub In subSubFolders
'        Dim path2 As String
'        path2 = rootPath & "\Temp\" & subSub
'        If Not fso.FolderExists(path2) Then fso.CreateFolder path2
'    Next subSub
'
'    Exit Sub
'
'EH:
'    Debug.Print "[BuildDirs][ERROR] " & Err.Description
'End Sub
'
'
'' ==============================================================
'' STEP 3 – SAVE HOST AS XLSM
'' ==============================================================
'
'Public Sub SaveHostAsXLSM(wb As Workbook, rootPath As String)
'    On Error Resume Next
'
'    Dim targetPath As String
'    targetPath = rootPath & "\" & Left$(wb.name, InStrRev(wb.name, ".") - 1) & ".xlsm"
'
'    If wb.path = rootPath And LCase$(right$(wb.name, 4)) = "xlsm" Then Exit Sub
'
'    Application.DisplayAlerts = False
'    wb.SaveAs fileName:=targetPath, FileFormat:=xlOpenXMLWorkbookMacroEnabled
'    Application.DisplayAlerts = True
'End Sub
'
'
'' ==============================================================
'' STEP 4 – PYTHON VENV
'' ==============================================================
'
'Public Function CreatePythonVenv(fso As Object, rootPath As String) As Boolean
'    On Error GoTo EH
'
'    Dim venvPath As String
'    venvPath = rootPath & "\Python\.venv"
'
'    Dim cmd As String
'    cmd = "python.exe -m venv """ & venvPath & """"
'
'    Dim sh As Object: Set sh = CreateObject("WScript.Shell")
'    sh.Run "cmd /c " & cmd, 0, True
'
'    CreatePythonVenv = fso.FolderExists(venvPath & "\Lib")
'    Exit Function
'
'EH:
'    CreatePythonVenv = False
'End Function
'
'
'' ==============================================================
'' STEP 5 – REAL RESOURCE EXTRACTION
'' ==============================================================
'
'Public Sub ExtractResources(fso As Object, rootPath As String)
'    On Error GoTo EH
'
'    Dim wb As Workbook
'    Set wb = HostManager_GetCurrentWorkbook()
'    If wb Is Nothing Then Exit Sub
'
'    Dim wsStore As Worksheet
'    On Error Resume Next
'    Set wsStore = wb.Worksheets(EMBED_SHEET_NAME)
'    On Error GoTo EH
'
'    If wsStore Is Nothing Then
'        Debug.Print "[Step5] No EmbeddedStore sheet found. Nothing to extract."
'        Exit Sub
'    End If
'
'    Debug.Print "[Step5] Extracting Embedded Resources..."
'
'    Dim outFolder As String
'    outFolder = rootPath & "\AddIn"
'    If Not fso.FolderExists(outFolder) Then fso.CreateFolder outFolder
'
'    ExtractEmbeddedStoreSheet wsStore, outFolder
'
'    DeleteExtractedResources rootPath
'
'    Debug.Print "[Step5] Extraction completed into: " & outFolder
'    Exit Sub
'
'EH:
'    Debug.Print "[Step5][ERROR] " & Err.Description
'End Sub
'
'
'
'' ==============================================================
'' EMBEDDED STORE EXTRACTOR (FULL WORKING VERSION)
'' ==============================================================
'
'Private Sub ExtractEmbeddedStoreSheet(wsStore As Worksheet, outRoot As String)
'    Dim lastRow As Long
'    lastRow = wsStore.Cells(wsStore.rows.count, "A").End(xlUp).Row
'    If lastRow < 2 Then Exit Sub
'
'    Dim fileMap As Object
'    Set fileMap = CreateObject("Scripting.Dictionary")
'
'    Dim r As Long
'    For r = 2 To lastRow
'        Dim fname As String: fname = CStr(wsStore.Cells(r, COL_FILENAME).value)
'        Dim rel As String: rel = CStr(wsStore.Cells(r, COL_RELPATH).value)
'        Dim key As String: key = rel & "|" & fname
'        Dim idx As Long: idx = CLng(wsStore.Cells(r, COL_CHUNKINDEX).value)
'        Dim b64 As String: b64 = CStr(wsStore.Cells(r, COL_BASE64).value)
'
'        If Not fileMap.Exists(key) Then
'            fileMap.Add key, CreateObject("Scripting.Dictionary")
'        End If
'
'        fileMap(key)(idx) = b64
'    Next r
'
'    Dim k As Variant
'    For Each k In fileMap.keys
'        Dim parts() As String: parts = Split(k, "|")
'        Dim relPath As String: relPath = parts(0)
'        Dim fileName As String: fileName = parts(1)
'
'        Dim chunks As Object: Set chunks = fileMap(k)
'        Dim chunkKeys As Variant: chunkKeys = chunks.keys
'
'        SortVariantNumeric chunkKeys
'
'        Dim bigB64 As String: bigB64 = ""
'        Dim i As Long
'        For i = LBound(chunkKeys) To UBound(chunkKeys)
'            bigB64 = bigB64 & chunks(chunkKeys(i))
'        Next i
'
'        Dim bytes() As Byte
'        bytes = Base64ToBinary(bigB64)
'
'        Dim fullOut As String
'        fullOut = outRoot & "\" & relPath
'
'        EnsureFolderExists fullOut
'        WriteBinaryFile fullOut, bytes
'    Next k
'End Sub
'
'
'Private Sub SortVariantNumeric(ByRef a As Variant)
'    Dim i As Long, j As Long, tmp As Variant
'    If IsEmpty(a) Then Exit Sub
'    For i = LBound(a) To UBound(a) - 1
'        For j = i + 1 To UBound(a)
'            If CLng(a(j)) < CLng(a(i)) Then
'                tmp = a(i): a(i) = a(j): a(j) = tmp
'            End If
'        Next j
'    Next i
'End Sub
'
'
'Private Sub EnsureFolderExists(fullPath As String)
'    Dim folder As String
'    folder = Left$(fullPath, InStrRev(fullPath, "\") - 1)
'    If Len(folder) = 0 Then Exit Sub
'
'    If Len(Dir(folder, vbDirectory)) = 0 Then CreateFoldersRecursive folder
'End Sub
'
'Private Sub CreateFoldersRecursive(folderPath As String)
'    Dim parts As Variant: parts = Split(folderPath, "\")
'    Dim build As String: build = parts(0) & "\"
'    Dim i As Long
'    For i = 1 To UBound(parts)
'        build = build & parts(i) & "\"
'        If Len(Dir(build, vbDirectory)) = 0 Then MkDir build
'    Next i
'End Sub
'
'
'Private Sub WriteBinaryFile(path As String, bytes() As Byte)
'    Dim stm As Object
'    Set stm = CreateObject("ADODB.Stream")
'    stm.Type = 1
'    stm.Open
'    stm.Write bytes
'    stm.SaveToFile path, 2
'    stm.Close
'End Sub
'
'Private Function Base64ToBinary(b64 As String) As Byte()
'    Dim xml As Object: Set xml = CreateObject("MSXML2.DOMDocument.6.0")
'    Dim node As Object: Set node = xml.createElement("b64")
'    node.DataType = "bin.base64"
'    node.text = b64
'    Base64ToBinary = node.nodeTypedValue
'End Function
'
'
'Public Sub DeleteExtractedResources(rootPath As String)
'    On Error Resume Next
'
'    Dim fso As Object
'    Set fso = CreateObject("Scripting.FileSystemObject")
'
'    Dim target As String
'    target = rootPath & "\AddIn"
'
'    If Not fso.FolderExists(target) Then Exit Sub
'
'    DeleteRecursiveFilesOnly fso.GetFolder(target)
'End Sub
'
'Private Sub DeleteRecursiveFilesOnly(folder As Object)
'    Dim f As Object
'    For Each f In folder.files
'        On Error Resume Next
'        f.Delete True
'    Next f
'
'    Dim subf As Object
'    For Each subf In folder.subFolders
'        DeleteRecursiveFilesOnly subf
'    Next subf
'End Sub
'
'
'' ==============================================================
'' PLACEHOLDERS 6 + 7
'' ==============================================================
'' --- Run pip install ONLY inside local venv ---
'Private Sub install_pip_Packages()
'    Dim venvPy As String
'
'    Dim wb As Workbook
'    Set wb = HostManager_GetCurrentWorkbook()
'
'    venvPy = JoinPath(wb.path, "Python\.venv\Scripts\python.exe")
'
'    Dim reqFile As String
'    reqFile = JoinPath(wb.path, "Python\Requirements.txt")
'
'    If Len(Dir$(venvPy, vbNormal)) <> 0 And Len(Dir$(reqFile, vbNormal)) <> 0 Then
'        Dim cmd As String
'        cmd = """" & venvPy & """ -m pip install -r """ & reqFile & """"
'
'        Dim sh As Object
'        Set sh = CreateObject("WScript.Shell")
'        sh.Run cmd, 0, True
'    End If
'End Sub
'
'
''Public Sub PyExcelSetup_Step7_SaveVersionID(wb As Workbook)
''    Debug.Print "[Step 7] Saving XLAM Version (Placeholder)."
''End Sub
'
'
'' ==============================================================
'' UTILITY
'' ==============================================================
'
'Public Sub SaveTextFile(filePath As String, fileContent As String)
'    On Error GoTo EH
'    Dim n As Integer: n = FreeFile
'    Open filePath For Output As #n
'    Print #n, fileContent
'    Close #n
'    Exit Sub
'EH:
'    If n <> 0 Then Close #n
'    MsgBox "Error writing file: " & Err.Description, vbCritical
'End Sub
'
'
'Private Function JoinPath(base As String, leaf As String) As String
'    If right$(base, 1) = "\" Then
'        JoinPath = base & leaf
'    Else
'        JoinPath = base & "\" & leaf
'    End If
'End Function
''





























'Option Explicit
'
'' STATUS MESSAGE STORAGE
'Public PyExcelSetup_LastMessage As String
'
'' CONSTANTS
'Private Const WORKBOOK_PATH_NAME As String = "ProjectRootPath"
'Private Const EMBED_SHEET_NAME As String = "EmbeddedStore"
'
'' COLUMNS IN EMBEDDED STORE
'Private Const COL_FILENAME As Long = 1
'Private Const COL_CHUNKINDEX As Long = 2
'Private Const COL_BASE64 As Long = 3
'Private Const COL_RELPATH As Long = 4
'
'' ==============================================================
'' CORE INSTALLATION LOGIC
'' ==============================================================
'
'Public Function PyExcelSetup() As Boolean
'    On Error GoTo EH
'
'    Dim wb As Workbook
'    Dim fso As Object
'    Dim hostPath As String
'
'    Debug.Print "============================================================"
'    Debug.Print "[PyExcelSetup] Initialization started."
'
'    ' Target Workbook (The one being set up)
'    Set wb = HostManager_GetCurrentWorkbook()
'    If wb Is Nothing Then
'        PyExcelSetup_LastMessage = "Failure in Step 0: No active workbook context."
'        GoTo Failed
'    End If
'
'    Set fso = CreateObject("Scripting.FileSystemObject")
'
'    ' 1. Set main path
'    hostPath = SelectAndSetupRootPath(wb, fso)
'    If Len(hostPath) = 0 Then
'        PyExcelSetup_LastMessage = "Failure in Step 1: Path selection cancelled or invalid."
'        GoTo Failed
'    End If
'
'    ' 2. Build project folder tree
'    BuildProjectDirectories fso, hostPath
'
'    ' 3. Save workbook as XLSM
'    If Not SaveHostAsXLSM(wb, hostPath) Then
'        PyExcelSetup_LastMessage = "Failure in Step 3: Could not save workbook as XLSM."
'        GoTo Failed
'    End If
'
'    ' 4. Python VENV
'    If Not CreatePythonVenv(fso, hostPath) Then
'        PyExcelSetup_LastMessage = "Failure in Step 4: Python venv creation did not complete. Ensure 'python' is in your system PATH."
'        GoTo Failed
'    End If
'
'    ' 5. EXTRACT EMBEDDED RESOURCES
'    ' This extracts FROM ThisWorkbook (Addin) INTO the hostPath
'    ExtractResources fso, hostPath
'
'    ' 6. PIP INSTALL
'    install_pip_Packages hostPath
'
'    PyExcelSetup_LastMessage = "Installation completed successfully."
'    Debug.Print "[PyExcelSetup] Installation finished."
'    Debug.Print "============================================================"
'
'    PyExcelSetup = True
'    Exit Function
'
'Failed:
'    If PyExcelSetup_LastMessage = "" Then
'        PyExcelSetup_LastMessage = "Installation terminated due to error/cancellation."
'    End If
'    Debug.Print "[PyExcelSetup] " & PyExcelSetup_LastMessage
'    PyExcelSetup = False
'    Exit Function
'
'EH:
'    PyExcelSetup_LastMessage = "Fatal error in PyExcelSetup: " & Err.Description
'    Debug.Print "[PyExcelSetup][FATAL ERROR] " & Err.Description
'    Resume Failed
'End Function
'
'
'' ==============================================================
'' STEP 1 - PATH SELECTOR
'' ==============================================================
'
'Public Function SelectAndSetupRootPath(wb As Workbook, fso As Object) As String
'    On Error GoTo EH
'
'    Dim fldr As Object
'    Dim defaultPath As String
'    Dim pathChosen As String
'
'    If Len(wb.path) > 0 Then
'        defaultPath = wb.path
'    Else
'        defaultPath = Environ("USERPROFILE")
'    End If
'
'    Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
'    With fldr
'        .Title = "Select Project Root Folder"
'        .InitialFileName = defaultPath
'        If .Show <> -1 Then Exit Function
'        pathChosen = .SelectedItems(1)
'    End With
'
'    ' Handle Unsaved workbooks safely (Book1, etc.)
'    Dim projectName As String
'    Dim dotIndex As Long
'    dotIndex = InStrRev(wb.name, ".")
'
'    If dotIndex > 0 Then
'        projectName = Left$(wb.name, dotIndex - 1)
'    Else
'        projectName = wb.name
'    End If
'
'    Dim finalPath As String
'    finalPath = pathChosen & Application.PathSeparator & projectName
'
'    If Not fso.FolderExists(finalPath) Then fso.CreateFolder finalPath
'
'    SaveProjectPath wb, finalPath
'    SelectAndSetupRootPath = finalPath
'    Exit Function
'
'EH:
'    Debug.Print "[SelectRootPath][ERROR] " & Err.Description
'End Function
'
'Public Sub SaveProjectPath(wb As Workbook, path As String)
'    On Error GoTo EH
'    ' Assumes modRibbon exists as per original code
'    modRibbon.SaveWorkbookValue wb, WORKBOOK_PATH_NAME, path
'    Exit Sub
'EH:
'    Debug.Print "[SaveProjectPath][ERROR] " & Err.Description
'End Sub
'
'
'' ==============================================================
'' STEP 2 - BUILD FOLDER TREE
'' ==============================================================
'
'Public Sub BuildProjectDirectories(fso As Object, rootPath As String)
'    On Error GoTo EH
'
'    Dim folder As Variant
'    Dim subFolders As Variant: subFolders = Array("AddIn", "Archive", "Python", "Temp", "userScripts")
'
'    For Each folder In subFolders
'        Dim full As String: full = rootPath & Application.PathSeparator & folder
'        If Not fso.FolderExists(full) Then fso.CreateFolder full
'    Next folder
'
'    Dim venvPath As String
'    venvPath = rootPath & "\Python\.venv"
'    If Not fso.FolderExists(venvPath) Then fso.CreateFolder venvPath
'
'    Dim subSub As Variant
'    Dim subSubFolders As Variant: subSubFolders = Array("assets", "lists", "tables", "values")
'
'    For Each subSub In subSubFolders
'        Dim path2 As String
'        path2 = rootPath & "\Temp\" & subSub
'        If Not fso.FolderExists(path2) Then fso.CreateFolder path2
'    Next subSub
'
'    Exit Sub
'
'EH:
'    Debug.Print "[BuildDirs][ERROR] " & Err.Description
'End Sub
'
'
'' ==============================================================
'' STEP 3 - SAVE HOST AS XLSM
'' ==============================================================
'
'Public Function SaveHostAsXLSM(wb As Workbook, rootPath As String) As Boolean
'    On Error GoTo EH
'
'    ' Handle unsaved workbook names
'    Dim baseName As String
'    Dim dotIndex As Long
'    dotIndex = InStrRev(wb.name, ".")
'
'    If dotIndex > 0 Then
'        baseName = Left$(wb.name, dotIndex - 1)
'    Else
'        baseName = wb.name
'    End If
'
'    Dim targetPath As String
'    targetPath = rootPath & "\" & baseName & ".xlsm"
'
'    ' Check if already saved in correct location
'    If wb.path = rootPath And LCase$(right$(wb.name, 5)) = ".xlsm" Then
'        SaveHostAsXLSM = True
'        Exit Function
'    End If
'
'    Application.DisplayAlerts = False
'    wb.SaveAs fileName:=targetPath, FileFormat:=xlOpenXMLWorkbookMacroEnabled
'    Application.DisplayAlerts = True
'
'    SaveHostAsXLSM = True
'    Exit Function
'EH:
'    Application.DisplayAlerts = True
'    Debug.Print "[SaveHostAsXLSM] Error: " & Err.Description
'    SaveHostAsXLSM = False
'End Function
'
'
'' ==============================================================
'' STEP 4 - PYTHON VENV
'' ==============================================================
'
'Public Function CreatePythonVenv(fso As Object, rootPath As String) As Boolean
'    On Error GoTo EH
'
'    Dim venvPath As String
'    venvPath = rootPath & "\Python\.venv"
'
'    Dim cmd As String
'    ' Note: Requires 'python' to be in system PATH.
'    cmd = "python.exe -m venv """ & venvPath & """"
'
'    Dim sh As Object: Set sh = CreateObject("WScript.Shell")
'    sh.Run "cmd /c " & cmd, 0, True
'
'    CreatePythonVenv = fso.FolderExists(venvPath & "\Lib")
'    Exit Function
'
'EH:
'    CreatePythonVenv = False
'End Function
'
'
'' ==============================================================
'' STEP 5 - REAL RESOURCE EXTRACTION
'' ==============================================================
'
'Public Sub ExtractResources(fso As Object, rootPath As String)
'    On Error GoTo EH
'
'    ' -------------------------------------------------------------
'    ' Source: ThisWorkbook (The AddIn containing the embedded files)
'    ' -------------------------------------------------------------
'    Dim wbSource As Workbook
'    Set wbSource = ThisWorkbook
'
'    Dim wsStore As Worksheet
'    On Error Resume Next
'    Set wsStore = wbSource.Worksheets(EMBED_SHEET_NAME)
'    On Error GoTo EH
'
'    If wsStore Is Nothing Then
'        Debug.Print "[Step5] No EmbeddedStore sheet found in " & wbSource.name
'        Exit Sub
'    End If
'
'    Debug.Print "[Step5] Extracting Embedded Resources..."
'
'    ' CORRECTION 1: Extract to Root, not AddIn folder
'    Dim outFolder As String
'    outFolder = rootPath
'
'    ' Ensure trailing slash
'    If right$(outFolder, 1) <> "\" Then outFolder = outFolder & "\"
'
'    ' CORRECTION 2: Removed DeleteExtractedResources
'    ' Since we are extracting to Root, we rely on WriteBinaryFile's
'    ' overwrite capability rather than wiping the folder.
'
'    ' 3. Extract using verified logic
'    ExtractEmbeddedStoreSheet wsStore, outFolder
'
'    Debug.Print "[Step5] Extraction completed into: " & outFolder
'    Exit Sub
'
'EH:
'    Debug.Print "[Step5][ERROR] " & Err.Description
'End Sub
'
'Private Sub ExtractEmbeddedStoreSheet(wsStore As Worksheet, outRoot As String)
'    Dim lastRow As Long
'    Dim r As Long
'
'    Dim fileMap As Object
'    Dim fileKey As String
'    Dim fname As String
'    Dim rel As String
'    Dim chunkIndex As Long
'    Dim b64 As String
'
'    lastRow = wsStore.Cells(wsStore.rows.count, "A").End(xlUp).Row
'    If lastRow < 2 Then Exit Sub
'
'    Set fileMap = CreateObject("Scripting.Dictionary")
'
'    ' --- PASS 1: MAP CHUNKS ---
'    For r = 2 To lastRow
'        fname = CStr(wsStore.Cells(r, COL_FILENAME).value)
'        rel = CStr(wsStore.Cells(r, COL_RELPATH).value)
'
'        If Len(fname) > 0 Or Len(rel) > 0 Then
'            fileKey = rel & "|" & fname
'
'            If Not fileMap.Exists(fileKey) Then
'                fileMap.Add fileKey, CreateObject("Scripting.Dictionary")
'            End If
'
'            chunkIndex = CLng(wsStore.Cells(r, COL_CHUNKINDEX).value)
'            b64 = CStr(wsStore.Cells(r, COL_BASE64).value)
'
'            fileMap(fileKey)(chunkIndex) = b64
'        End If
'    Next r
'
'    ' --- PASS 2: REBUILD AND WRITE ---
'    Dim k As Variant
'    For Each k In fileMap.keys
'        Dim parts() As String
'        Dim relPath As String
'        Dim chunksDict As Object
'        Dim idxs As Variant
'        Dim i As Long
'        Dim bigB64 As String
'        Dim bytes() As Byte
'        Dim outPath As String
'
'        parts = Split(CStr(k), "|")
'        relPath = parts(0)
'
'        Set chunksDict = fileMap(k)
'        idxs = chunksDict.keys
'
'        ' Sort chunks numerically
'        SortVariantNumeric idxs
'
'        bigB64 = ""
'        For i = LBound(idxs) To UBound(idxs)
'            bigB64 = bigB64 & chunksDict(idxs(i))
'        Next i
'
'        ' Decode and Write
'        bytes = Base64ToBinary(bigB64)
'        outPath = outRoot & relPath
'
'        EnsureFolderExists outPath
'        WriteBinaryFile outPath, bytes
'    Next k
'End Sub
'
'
'' ==============================================================
'' STEP 6 - PIP INSTALL (DEBUGGING MODE)
'' ==============================================================
'
'Private Sub install_pip_Packages(targetPath As String)
'    Dim venvPy As String
'    Dim reqFile As String
'    Dim cmd As String
'    Dim sh As Object
'
'    ' REMOVED: Dependency on HostManager
'    ' Dim wb As Workbook
'    ' Set wb = HostManager_GetCurrentWorkbook()
'
'    ' LOGIC: Use the explicitly passed path (hostPath)
'    ' This guarantees we install into the folder we just created/saved to.
'
'    ' Construct Paths using targetPath
'    venvPy = JoinPath(targetPath, "Python\.venv\Scripts\python.exe")
'    reqFile = JoinPath(targetPath, "Python\Requirements.txt")
'
'    ' Verify files exist before trying to run
'    If Len(Dir$(venvPy, vbNormal)) = 0 Then
'        Debug.Print "[Step 6] Error: Python executable not found at " & venvPy
'        Exit Sub
'    End If
'    If Len(Dir$(reqFile, vbNormal)) = 0 Then
'        Debug.Print "[Step 6] Error: Requirements file not found at " & reqFile
'        Exit Sub
'    End If
'
'    ' cmd /c closes the window when done (change to /k to keep open for debug)
'    cmd = "cmd /c """"" & venvPy & """ -m pip install -r """ & reqFile & """ --no-input"""
'
'    Debug.Print "[Step 6] Running PIP Install..."
'    Debug.Print "[Step 6] Target VENV: " & venvPy
'
'    Set sh = CreateObject("WScript.Shell")
'    sh.Run cmd, 1, True
'
'    Debug.Print "[Step 6] PIP Install command returned."
'End Sub
'
'' ==============================================================
'' UTILITIES
'' ==============================================================
'
'Private Sub SortVariantNumeric(ByRef a As Variant)
'    Dim i As Long, j As Long, tmp As Variant
'
'    If IsEmpty(a) Then Exit Sub
'    If UBound(a) <= LBound(a) Then Exit Sub
'
'    For i = LBound(a) To UBound(a) - 1
'        For j = i + 1 To UBound(a)
'            If CLng(a(j)) < CLng(a(i)) Then
'                tmp = a(i): a(i) = a(j): a(j) = tmp
'            End If
'        Next j
'    Next i
'End Sub
'
'Private Function Base64ToBinary(b64 As String) As Byte()
'    Dim xml As Object: Set xml = CreateObject("MSXML2.DOMDocument.6.0")
'    Dim node As Object: Set node = xml.createElement("b64")
'    node.DataType = "bin.base64"
'    node.text = b64
'    Base64ToBinary = node.nodeTypedValue
'End Function
'
'Private Sub EnsureFolderExists(fullPath As String)
'    Dim folder As String
'    folder = Left$(fullPath, InStrRev(fullPath, "\") - 1)
'    If Len(folder) = 0 Then Exit Sub
'    If Len(Dir(folder, vbDirectory)) = 0 Then CreateFoldersRecursive folder
'End Sub
'
'Private Sub CreateFoldersRecursive(folderPath As String)
'    Dim parts As Variant: parts = Split(folderPath, "\")
'    Dim build As String: build = parts(0) & "\"
'    Dim i As Long
'    For i = 1 To UBound(parts)
'        build = build & parts(i) & "\"
'        If Len(Dir(build, vbDirectory)) = 0 Then MkDir build
'    Next i
'End Sub
'
'Private Sub WriteBinaryFile(path As String, bytes() As Byte)
'    Dim stm As Object
'    Set stm = CreateObject("ADODB.Stream")
'    stm.Type = 1 ' adTypeBinary
'    stm.Open
'    stm.Write bytes
'    stm.SaveToFile path, 2 ' adSaveCreateOverWrite
'    stm.Close
'End Sub
'
''Private Sub DeleteRecursiveFilesOnly(folder As Object)
''    Dim f As Object
''    For Each f In folder.files
''        On Error Resume Next
''        f.Delete True
''    Next f
''    Dim subf As Object
''    For Each subf In folder.subFolders
''        DeleteRecursiveFilesOnly subf
''    Next subf
''End Sub
'
'Private Function JoinPath(base As String, leaf As String) As String
'    If right$(base, 1) = "\" Then
'        JoinPath = base & leaf
'    Else
'        JoinPath = base & "\" & leaf
'    End If
'End Function
'





























Option Explicit

' STATUS MESSAGE STORAGE
Public PyExcelSetup_LastMessage As String

' PROGRESS BAR GLOBAL
Public CurrentProgressForm As Object

' CONSTANTS
Private Const WORKBOOK_PATH_NAME As String = "ProjectRootPath"
Private Const EMBED_SHEET_NAME As String = "EmbeddedStore"

' COLUMNS IN EMBEDDED STORE
Private Const COL_FILENAME As Long = 1
Private Const COL_CHUNKINDEX As Long = 2
Private Const COL_BASE64 As Long = 3
Private Const COL_RELPATH As Long = 4

' ==============================================================
' PROGRESS BAR UTILITIES
' ==============================================================

Public Sub InitProgressBar()
    Set CurrentProgressForm = New ufProgress
    CurrentProgressForm.lblBar.Width = 0
    CurrentProgressForm.Show vbModeless
    DoEvents
End Sub

Public Sub UpdateProgress(pct As Double, msg As String)
    If CurrentProgressForm Is Nothing Then Exit Sub
    
    ' Update Text
    CurrentProgressForm.lblStatus.Caption = msg
    
    ' Update Bar Width
    Dim maxWidth As Double
    ' Use InsideWidth of the container frame
    maxWidth = CurrentProgressForm.fraBackground.InsideWidth
    
    ' Safety cap at 100%
    If pct > 1 Then pct = 1
    
    CurrentProgressForm.lblBar.Width = maxWidth * pct
    
    ' Force UI Update
    CurrentProgressForm.Repaint
    DoEvents
End Sub

Public Sub CloseProgressBar()
    On Error Resume Next
    If Not CurrentProgressForm Is Nothing Then
        Unload CurrentProgressForm
        Set CurrentProgressForm = Nothing
    End If
End Sub

' ==============================================================
' CORE INSTALLATION LOGIC
' ==============================================================

Public Function PyExcelSetup() As Boolean
    On Error GoTo EH

    Dim wb As Workbook
    Dim fso As Object
    Dim hostPath As String
    Dim userChoice As VbMsgBoxResult
    
    Debug.Print "============================================================"
    Debug.Print "[PyExcelSetup] Initialization started."

    userChoice = MsgBox( _
        "This installation may take several minutes and requires Python to be installed on this machine." & vbCrLf & vbCrLf & _
        "Do you want to continue?", _
        vbYesNo + vbQuestion, _
        "Confirm Installation" _
    )

    If userChoice <> vbYes Then
        PyExcelSetup_LastMessage = "Installation cancelled by user."
        PyExcelSetup = False
        Exit Function
    End If

    Set wb = HostManager_GetCurrentWorkbook()
    If wb Is Nothing Then
        PyExcelSetup_LastMessage = "Failure in Step 0: No active workbook context."
        GoTo Failed
    End If

    Set fso = CreateObject("Scripting.FileSystemObject")

    hostPath = SelectAndSetupRootPath(wb, fso)
    If Len(hostPath) = 0 Then
        PyExcelSetup_LastMessage = "Failure in Step 1: Path selection cancelled or invalid."
        GoTo Failed
    End If

    InitProgressBar
    UpdateProgress 0.1, "Initializing project folders..."

    BuildProjectDirectories fso, hostPath
    UpdateProgress 0.2, "Folders created."

    UpdateProgress 0.25, "Saving workbook as XLSM..."
    If Not SaveHostAsXLSM(wb, hostPath) Then
        PyExcelSetup_LastMessage = "Failure in Step 3: Could not save workbook as XLSM."
        GoTo Failed
    End If
    UpdateProgress 0.3, "Workbook saved."

    UpdateProgress 0.35, "Creating Python Environment (Excel will pause)..."
    If Not CreatePythonVenv(fso, hostPath) Then
        PyExcelSetup_LastMessage = "Failure in Step 4: Python venv creation did not complete."
        GoTo Failed
    End If
    UpdateProgress 0.5, "Python Environment Ready."

    UpdateProgress 0.5, "Starting resource extraction..."
    ExtractResources fso, hostPath
    UpdateProgress 0.8, "Resources Extracted."

    UpdateProgress 0.85, "Installing Python libraries (Excel will pause)..."
    install_pip_Packages hostPath

    ' Stamp the current addin version to the workbook
    UpdateProgress 0.95, "Finalizing setup..."
    Update.SetStoredProjectVersion wb, Update.GetAddinVersion()
    Debug.Print "[PyExcelSetup] Version stamped: " & Update.GetAddinVersion()

    UpdateProgress 1#, "Installation Completed!"
    Application.Wait Now + TimeValue("0:00:01")

    PyExcelSetup_LastMessage = "Installation completed successfully."
    Debug.Print "[PyExcelSetup] Installation finished."
    Debug.Print "============================================================"

    CloseProgressBar
    PyExcelSetup = True
    Exit Function

Failed:
    CloseProgressBar
    If PyExcelSetup_LastMessage = "" Then
        PyExcelSetup_LastMessage = "Installation terminated due to error/cancellation."
    End If
    Debug.Print "[PyExcelSetup] " & PyExcelSetup_LastMessage
    MsgBox "Setup Failed: " & PyExcelSetup_LastMessage, vbCritical
    PyExcelSetup = False
    Exit Function

EH:
    CloseProgressBar
    PyExcelSetup_LastMessage = "Fatal error in PyExcelSetup: " & Err.Description
    Debug.Print "[PyExcelSetup][FATAL ERROR] " & Err.Description
    Resume Failed
End Function


' ==============================================================
' STEP 1 - PATH SELECTOR
' ==============================================================

Public Function SelectAndSetupRootPath(wb As Workbook, fso As Object) As String
    On Error GoTo EH

    Dim fldr As Object
    Dim defaultPath As String
    Dim pathChosen As String
    Dim projectName As String
    Dim dotIndex As Long
    Dim finalPath As String

    ' Set default path for folder picker
    If Len(wb.path) > 0 Then
        defaultPath = wb.path
    Else
        defaultPath = Environ$("USERPROFILE")
    End If

    ' Always show folder picker - user chooses location manually
    Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
    With fldr
        .Title = "Select Project Root Folder"
        .InitialFileName = defaultPath
        If .Show <> -1 Then Exit Function
        pathChosen = .SelectedItems(1)
    End With

    ' Extract Project Name from Workbook
    dotIndex = InStrRev(wb.name, ".")
    If dotIndex > 0 Then
        projectName = Left$(wb.name, dotIndex - 1)
    Else
        projectName = wb.name
    End If

    ' Combine paths
    If right$(pathChosen, 1) <> "\" Then pathChosen = pathChosen & "\"
    finalPath = pathChosen & projectName

    ' Use PathUtils to create the folder (handles recursive creation safely)
    Call EnsureFolderExists(finalPath)

    SelectAndSetupRootPath = finalPath
    Exit Function

EH:
    Debug.Print "[SelectAndSetupRootPath][ERROR] " & Err.Description
    SelectAndSetupRootPath = ""
End Function




' ==============================================================
' STEP 2 - BUILD FOLDER TREE (USING PATHUTILS)
' ==============================================================

Public Sub BuildProjectDirectories(fso As Object, rootPath As String)
    On Error GoTo EH

    ' Main structure - use PathUtils EnsureFolderPath for recursive creation
    Call EnsureFolderPath(rootPath, "AddIn")
    Call EnsureFolderPath(rootPath, "Archive")
    Call EnsureFolderPath(rootPath, "Python")
    Call EnsureFolderPath(rootPath, "userScripts")

    ' Nested structures - venv path
    Dim venvPath As String
    venvPath = rootPath & "\Python\.venv"
    Call EnsureFolderExists(venvPath)

    ' Temp subfolders
    Dim subSub As Variant
    Dim subSubFolders As Variant: subSubFolders = Array("assets", "lists", "tables", "values")

    For Each subSub In subSubFolders
        ' Use PathUtils to handle the subfolder creation
        Call EnsureFolderPath(rootPath & "\Temp", CStr(subSub))
    Next subSub

    Exit Sub

EH:
    Debug.Print "[BuildDirs][ERROR] " & Err.Description
End Sub


' ==============================================================
' STEP 3 - SAVE HOST AS XLSM
' ==============================================================

Public Function SaveHostAsXLSM(wb As Workbook, rootPath As String) As Boolean
    On Error GoTo EH
    
    ' Handle unsaved workbook names
    Dim baseName As String
    Dim dotIndex As Long
    dotIndex = InStrRev(wb.name, ".")
    
    If dotIndex > 0 Then
        baseName = Left$(wb.name, dotIndex - 1)
    Else
        baseName = wb.name
    End If

    Dim targetPath As String
    targetPath = rootPath & "\" & baseName & ".xlsm"

    ' Check if already saved in correct location
    If wb.path = rootPath And LCase$(right$(wb.name, 5)) = ".xlsm" Then
        SaveHostAsXLSM = True
        Exit Function
    End If

    Application.DisplayAlerts = False
    wb.SaveAs fileName:=targetPath, FileFormat:=xlOpenXMLWorkbookMacroEnabled
    Application.DisplayAlerts = True
    
    SaveHostAsXLSM = True
    Exit Function
EH:
    Application.DisplayAlerts = True
    Debug.Print "[SaveHostAsXLSM] Error: " & Err.Description
    SaveHostAsXLSM = False
End Function


' ==============================================================
' STEP 4 - PYTHON VENV
' ==============================================================

Public Function CreatePythonVenv(fso As Object, rootPath As String) As Boolean
    On Error GoTo EH
    
    Dim venvPath As String
    venvPath = rootPath & "\Python\.venv"

    Dim cmd As String
    ' Note: Requires 'python' to be in system PATH.
    cmd = "python.exe -m venv """ & venvPath & """"

    Dim sh As Object: Set sh = CreateObject("WScript.Shell")
    sh.Run "cmd /c " & cmd, 0, True

    CreatePythonVenv = fso.FolderExists(venvPath & "\Lib")
    Exit Function

EH:
    CreatePythonVenv = False
End Function


' ==============================================================
' STEP 5 - REAL RESOURCE EXTRACTION
' ==============================================================

Public Sub ExtractResources(fso As Object, rootPath As String)
    On Error GoTo EH

    ' -------------------------------------------------------------
    ' Source: ThisWorkbook (The AddIn containing the embedded files)
    ' -------------------------------------------------------------
    Dim wbSource As Workbook
    Set wbSource = ThisWorkbook
    
    Dim wsStore As Worksheet
    On Error Resume Next
    Set wsStore = wbSource.Worksheets(EMBED_SHEET_NAME)
    On Error GoTo EH

    If wsStore Is Nothing Then
        Debug.Print "[Step5] No EmbeddedStore sheet found in " & wbSource.name
        Exit Sub
    End If

    Debug.Print "[Step5] Extracting Embedded Resources..."

    ' CORRECTION 1: Extract to Root, not AddIn folder
    Dim outFolder As String
    outFolder = rootPath
    
    ' Ensure trailing slash
    If right$(outFolder, 1) <> "\" Then outFolder = outFolder & "\"

    ' 3. Extract using verified logic
    ExtractEmbeddedStoreSheet wsStore, outFolder

    Debug.Print "[Step5] Extraction completed into: " & outFolder
    Exit Sub

EH:
    Debug.Print "[Step5][ERROR] " & Err.Description
End Sub

Private Sub ExtractEmbeddedStoreSheet(wsStore As Worksheet, outRoot As String)
    Dim lastRow As Long
    Dim r As Long
    
    Dim fileMap As Object
    Dim fileKey As String
    Dim fName As String
    Dim rel As String
    Dim chunkIndex As Long
    Dim b64 As String

    lastRow = wsStore.Cells(wsStore.rows.count, "A").End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    Set fileMap = CreateObject("Scripting.Dictionary")

    ' --- PASS 1: MAP CHUNKS ---
    For r = 2 To lastRow
        fName = CStr(wsStore.Cells(r, COL_FILENAME).value)
        rel = CStr(wsStore.Cells(r, COL_RELPATH).value)
        
        If Len(fName) > 0 Or Len(rel) > 0 Then
            fileKey = rel & "|" & fName
            
            If Not fileMap.Exists(fileKey) Then
                fileMap.Add fileKey, CreateObject("Scripting.Dictionary")
            End If
            
            chunkIndex = CLng(wsStore.Cells(r, COL_CHUNKINDEX).value)
            b64 = CStr(wsStore.Cells(r, COL_BASE64).value)
            
            fileMap(fileKey)(chunkIndex) = b64
        End If
    Next r

    ' --- PASS 2: REBUILD AND WRITE (WITH PROGRESS) ---
    Dim k As Variant
    Dim parts() As String
    Dim relPath As String
    Dim chunksDict As Object
    Dim idxs As Variant
    Dim i As Long
    Dim bigB64 As String
    Dim bytes() As Byte
    Dim outPath As String
    Dim folderPath As String
    
    ' Variables for Progress Calculation
    Dim totalFiles As Long
    Dim currentFile As Long
    Dim startPct As Double: startPct = 0.5 ' Extraction starts at 50%
    Dim endPct As Double: endPct = 0.8     ' Extraction ends at 80%
    Dim rangePct As Double: rangePct = endPct - startPct
    Dim calcPct As Double

    totalFiles = fileMap.count
    currentFile = 0

    For Each k In fileMap.keys
        currentFile = currentFile + 1
        
        parts = Split(CStr(k), "|")
        relPath = parts(0)
        
        ' Update Progress Bar
        ' Logic: Start% + (PercentOfFilesDone * RangeSize)
        calcPct = startPct + ((currentFile / totalFiles) * rangePct)
        UpdateProgress calcPct, "Extracting: " & parts(1)
        
        Set chunksDict = fileMap(k)
        idxs = chunksDict.keys
        
        ' Sort chunks numerically
        SortVariantNumeric idxs
        
        bigB64 = ""
        For i = LBound(idxs) To UBound(idxs)
            bigB64 = bigB64 & chunksDict(idxs(i))
        Next i
        
        ' Decode and Write
        bytes = Base64ToBinary(bigB64)
        outPath = outRoot & relPath

        ' Extract folder from file path and ensure it exists using PathUtils
        folderPath = Left$(outPath, InStrRev(outPath, "\") - 1)
        If Len(folderPath) > 0 Then Call EnsureFolderExists(folderPath)

        WriteBinaryFile outPath, bytes
    Next k
End Sub


' ==============================================================
' STEP 6 - PIP INSTALL (DEBUGGING MODE)
' ==============================================================

Private Sub install_pip_Packages(targetPath As String)
    Dim venvPy As String
    Dim reqFile As String
    Dim cmd As String
    Dim sh As Object

    ' LOGIC: Use the explicitly passed path (hostPath)
    
    ' Construct Paths using targetPath
    venvPy = JoinPath(targetPath, "Python\.venv\Scripts\python.exe")
    reqFile = JoinPath(targetPath, "Python\Requirements.txt")

    ' Verify files exist before trying to run
    If Len(Dir$(venvPy, vbNormal)) = 0 Then
        Debug.Print "[Step 6] Error: Python executable not found at " & venvPy
        Exit Sub
    End If
    If Len(Dir$(reqFile, vbNormal)) = 0 Then
        Debug.Print "[Step 6] Error: Requirements file not found at " & reqFile
        Exit Sub
    End If

    ' cmd /c closes the window when done
    cmd = "cmd /c """"" & venvPy & """ -m pip install -r """ & reqFile & """ --no-input"""

    Debug.Print "[Step 6] Running PIP Install..."
    Debug.Print "[Step 6] Target VENV: " & venvPy
    
    Set sh = CreateObject("WScript.Shell")
    sh.Run cmd, 1, True
    
    Debug.Print "[Step 6] PIP Install command returned."
End Sub

' ==============================================================
' UTILITIES
' ==============================================================

Private Sub SortVariantNumeric(ByRef a As Variant)
    Dim i As Long, j As Long, tmp As Variant
    
    If IsEmpty(a) Then Exit Sub
    If UBound(a) <= LBound(a) Then Exit Sub
    
    For i = LBound(a) To UBound(a) - 1
        For j = i + 1 To UBound(a)
            If CLng(a(j)) < CLng(a(i)) Then
                tmp = a(i): a(i) = a(j): a(j) = tmp
            End If
        Next j
    Next i
End Sub

Private Function Base64ToBinary(b64 As String) As Byte()
    Dim xml As Object: Set xml = CreateObject("MSXML2.DOMDocument.6.0")
    Dim node As Object: Set node = xml.createElement("b64")
    node.DataType = "bin.base64"
    node.text = b64
    Base64ToBinary = node.nodeTypedValue
End Function

' Note: EnsureFolderExists and CreateFoldersRecursive are now in PathUtils module

Private Sub WriteBinaryFile(path As String, bytes() As Byte)
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 1 ' adTypeBinary
    stm.Open
    stm.Write bytes
    stm.SaveToFile path, 2 ' adSaveCreateOverWrite
    stm.Close
End Sub

Public Function JoinPath(base As String, leaf As String) As String
    If right$(base, 1) = "\" Then
        JoinPath = base & leaf
    Else
        JoinPath = base & "\" & leaf
    End If
End Function




