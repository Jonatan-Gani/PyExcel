Attribute VB_Name = "Setup"
Option Explicit

' STATUS MESSAGE STORAGE
Public PyExcelSetup_LastMessage As String

' PROGRESS BAR GLOBAL
Public CurrentProgressForm As Object

' SETUP LOG STORAGE
Private SetupLogEntries As Collection
Private SetupStats As Object ' Dictionary for tracking counts

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
' SETUP LOGGING UTILITIES
' ==============================================================

Private Sub InitSetupLog()
    Set SetupLogEntries = New Collection
    Set SetupStats = CreateObject("Scripting.Dictionary")
    SetupStats("FilesExtracted") = 0
    SetupStats("FilesFailed") = 0
    SetupStats("PackagesInstalled") = 0
    SetupStats("PackagesFailed") = 0
    SetupStats("Warnings") = 0
    LogMessage "INFO", "PyExcel Setup Started", "Version: " & Update.GetAddinVersion()
End Sub

Private Sub LogMessage(logLevel As String, category As String, message As String)
    If SetupLogEntries Is Nothing Then Set SetupLogEntries = New Collection
    Dim entry As String
    entry = Format(Now, "yyyy-mm-dd hh:nn:ss") & " [" & logLevel & "] " & category & ": " & message
    SetupLogEntries.Add entry
    Debug.Print entry
End Sub

Private Sub WriteSetupLog(rootPath As String)
    On Error Resume Next

    Dim logPath As String
    logPath = rootPath & "\Temp\setup_log.txt"

    ' Ensure Temp folder exists
    Call EnsureFolderExists(rootPath & "\Temp")

    Dim fNum As Integer
    fNum = FreeFile
    Open logPath For Output As #fNum

    ' Write header
    Print #fNum, "========================================"
    Print #fNum, "PyExcel Setup Log"
    Print #fNum, "Generated: " & Format(Now, "yyyy-mm-dd hh:nn:ss")
    Print #fNum, "========================================"
    Print #fNum, ""

    ' Write stats summary
    Print #fNum, "--- SUMMARY ---"
    Print #fNum, "Files Extracted: " & SetupStats("FilesExtracted")
    Print #fNum, "Files Failed: " & SetupStats("FilesFailed")
    Print #fNum, "Packages Installed: " & SetupStats("PackagesInstalled")
    Print #fNum, "Packages Failed: " & SetupStats("PackagesFailed")
    Print #fNum, "Warnings: " & SetupStats("Warnings")
    Print #fNum, ""
    Print #fNum, "--- DETAILED LOG ---"

    ' Write all log entries
    Dim entry As Variant
    For Each entry In SetupLogEntries
        Print #fNum, entry
    Next entry

    Close #fNum

    LogMessage "INFO", "Log File", "Written to " & logPath
End Sub

Private Function GetSetupSummary() As String
    Dim summary As String
    summary = "PyExcel Setup Complete!" & vbCrLf & vbCrLf
    summary = summary & "Files Extracted: " & SetupStats("FilesExtracted") & vbCrLf
    summary = summary & "Packages Installed: " & SetupStats("PackagesInstalled") & vbCrLf

    If SetupStats("Warnings") > 0 Then
        summary = summary & vbCrLf & "Warnings: " & SetupStats("Warnings") & vbCrLf
        summary = summary & "(Check setup_log.txt in Temp folder for details)"
    End If

    GetSetupSummary = summary
End Function

' ==============================================================
' CORE INSTALLATION LOGIC
' ==============================================================

Public Function PyExcelSetup() As Boolean
    On Error GoTo EH

    Dim wb As Workbook
    Dim fso As Object
    Dim hostPath As String
    Dim userChoice As VbMsgBoxResult
    Dim extractedCount As Long
    Dim expectedCount As Long

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

    ' Initialize logging
    InitSetupLog

    Set wb = HostManager_GetCurrentWorkbook()
    If wb Is Nothing Then
        PyExcelSetup_LastMessage = "Failure in Step 0: No active workbook context."
        LogMessage "ERROR", "Initialization", PyExcelSetup_LastMessage
        GoTo Failed
    End If
    LogMessage "INFO", "Workbook", "Target workbook: " & wb.name

    Set fso = CreateObject("Scripting.FileSystemObject")

    hostPath = SelectAndSetupRootPath(wb, fso)
    If Len(hostPath) = 0 Then
        PyExcelSetup_LastMessage = "Failure in Step 1: Path selection cancelled or invalid."
        LogMessage "ERROR", "Path Selection", PyExcelSetup_LastMessage
        GoTo Failed
    End If
    LogMessage "INFO", "Path Selection", "Project root: " & hostPath

    InitProgressBar
    UpdateProgress 0.1, "Initializing project folders..."

    BuildProjectDirectories fso, hostPath
    LogMessage "INFO", "Folders", "Project directory structure created"
    UpdateProgress 0.2, "Folders created."

    UpdateProgress 0.25, "Saving workbook as XLSM..."
    If Not SaveHostAsXLSM(wb, hostPath) Then
        PyExcelSetup_LastMessage = "Failure in Step 3: Could not save workbook as XLSM."
        LogMessage "ERROR", "Save Workbook", PyExcelSetup_LastMessage
        GoTo Failed
    End If
    LogMessage "INFO", "Save Workbook", "Workbook saved to " & hostPath
    UpdateProgress 0.3, "Workbook saved."

    UpdateProgress 0.35, "Creating Python Environment (Excel will pause)..."
    If Not CreatePythonVenv(fso, hostPath) Then
        PyExcelSetup_LastMessage = "Failure in Step 4: Python venv creation did not complete. Ensure Python is installed and in PATH."
        LogMessage "ERROR", "Python Venv", PyExcelSetup_LastMessage
        GoTo Failed
    End If
    LogMessage "INFO", "Python Venv", "Virtual environment created at " & hostPath & "\Python\.venv"
    UpdateProgress 0.5, "Python Environment Ready."

    UpdateProgress 0.5, "Starting resource extraction..."
    extractedCount = ExtractResourcesWithVerification(fso, hostPath, expectedCount)
    If extractedCount = 0 Then
        PyExcelSetup_LastMessage = "Failure in Step 5: No files were extracted from EmbeddedStore."
        LogMessage "ERROR", "Extraction", PyExcelSetup_LastMessage
        GoTo Failed
    End If
    LogMessage "INFO", "Extraction", "Extracted " & extractedCount & " of " & expectedCount & " files"
    SetupStats("FilesExtracted") = extractedCount
    UpdateProgress 0.7, "Resources Extracted."

    ' Verify extracted files exist
    UpdateProgress 0.72, "Verifying extracted files..."
    If Not VerifyExtractedFiles(fso, hostPath) Then
        PyExcelSetup_LastMessage = "Failure in Step 5b: Some required files were not extracted correctly."
        LogMessage "ERROR", "Verification", PyExcelSetup_LastMessage
        GoTo Failed
    End If
    LogMessage "INFO", "Verification", "All critical files verified"
    UpdateProgress 0.75, "Files verified."

    UpdateProgress 0.78, "Installing Python libraries (Excel will pause)..."
    If Not InstallPipPackages(hostPath) Then
        PyExcelSetup_LastMessage = "Failure in Step 6: pip install failed. Check setup_log.txt for details."
        LogMessage "ERROR", "Pip Install", PyExcelSetup_LastMessage
        GoTo Failed
    End If
    UpdateProgress 0.88, "Libraries installed."

    ' Verify pip packages
    UpdateProgress 0.9, "Verifying installed packages..."
    Dim installedCount As Long, requiredCount As Long
    If Not VerifyPipPackages(hostPath, installedCount, requiredCount) Then
        PyExcelSetup_LastMessage = "Failure in Step 6b: Some required packages were not installed. Installed " & installedCount & " of " & requiredCount
        LogMessage "ERROR", "Pip Verification", PyExcelSetup_LastMessage
        GoTo Failed
    End If
    SetupStats("PackagesInstalled") = installedCount
    LogMessage "INFO", "Pip Verification", "Verified " & installedCount & " of " & requiredCount & " packages"
    UpdateProgress 0.95, "Packages verified."

    ' Stamp the current addin version to the workbook
    UpdateProgress 0.97, "Finalizing setup..."
    Update.SetStoredProjectVersion wb, Update.GetAddinVersion()
    LogMessage "INFO", "Version", "Version stamped: " & Update.GetAddinVersion()

    ' Write log file
    WriteSetupLog hostPath

    UpdateProgress 1#, "Installation Completed!"
    Application.Wait Now + TimeValue("0:00:02")

    CloseProgressBar

    ' Show summary dialog
    MsgBox GetSetupSummary(), vbInformation, "PyExcel Setup Complete"

    PyExcelSetup_LastMessage = "Installation completed successfully."
    Debug.Print "[PyExcelSetup] Installation finished."
    Debug.Print "============================================================"

    PyExcelSetup = True
    Exit Function

Failed:
    ' Write log even on failure
    If Len(hostPath) > 0 Then WriteSetupLog hostPath

    CloseProgressBar
    If PyExcelSetup_LastMessage = "" Then
        PyExcelSetup_LastMessage = "Installation terminated due to error/cancellation."
    End If
    Debug.Print "[PyExcelSetup] " & PyExcelSetup_LastMessage
    MsgBox "Setup Failed: " & PyExcelSetup_LastMessage & vbCrLf & vbCrLf & _
           "Check setup_log.txt in the Temp folder for details.", vbCritical, "PyExcel Setup Failed"
    PyExcelSetup = False
    Exit Function

EH:
    LogMessage "FATAL", "Exception", Err.Description
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
    
    Debug.Print "[SaveHostAsXLSM] === Starting SaveHostAsXLSM ==="
    Debug.Print "[SaveHostAsXLSM] Workbook name: " & wb.name
    Debug.Print "[SaveHostAsXLSM] Workbook path: " & wb.path
    Debug.Print "[SaveHostAsXLSM] Workbook FullName: " & wb.FullName
    Debug.Print "[SaveHostAsXLSM] Root path: " & rootPath
    Debug.Print "[SaveHostAsXLSM] Workbook.Saved status: " & wb.Saved
    Debug.Print "[SaveHostAsXLSM] Workbook.ReadOnly: " & wb.ReadOnly
    
    ' Handle unsaved workbook names
    Dim baseName As String
    Dim dotIndex As Long
    dotIndex = InStrRev(wb.name, ".")
    
    If dotIndex > 0 Then
        baseName = Left$(wb.name, dotIndex - 1)
    Else
        baseName = wb.name
    End If
    
    Debug.Print "[SaveHostAsXLSM] Base name extracted: " & baseName
    
    ' CREATE THE ROOT PATH IF IT DOESN'T EXIST
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FolderExists(rootPath) Then
        Debug.Print "[SaveHostAsXLSM] Root path does not exist, creating it..."
        On Error Resume Next
        MkDir rootPath
        If Err.Number <> 0 Then
            Debug.Print "[SaveHostAsXLSM] ERROR: Could not create root path: " & Err.Description
            SaveHostAsXLSM = False
            Exit Function
        End If
        On Error GoTo EH
        Debug.Print "[SaveHostAsXLSM] Root path created successfully"
    Else
        Debug.Print "[SaveHostAsXLSM] Root path exists: OK"
    End If
    
    Dim targetPath As String
    targetPath = rootPath & "\" & baseName & ".xlsm"
    Debug.Print "[SaveHostAsXLSM] Target path: " & targetPath
    
    ' Check if already saved in correct location
    If wb.path = rootPath And LCase$(right$(wb.name, 5)) = ".xlsm" Then
        Debug.Print "[SaveHostAsXLSM] Workbook already saved in correct location"
        SaveHostAsXLSM = True
        Exit Function
    End If
    
    ' Check if file already exists at target
    On Error Resume Next
    Dim fileExists As Boolean
    fileExists = (Dir(targetPath) <> "")
    Debug.Print "[SaveHostAsXLSM] Target file exists: " & fileExists
    On Error GoTo EH
    
    ' Check if workbook has VBA project and if it's protected
    On Error Resume Next
    Dim hasVBA As Boolean
    Dim vbaProjectName As String
    vbaProjectName = wb.VBProject.name
    hasVBA = (Err.Number = 0)
    Debug.Print "[SaveHostAsXLSM] Has VBA Project: " & hasVBA
    If hasVBA Then
        Debug.Print "[SaveHostAsXLSM] VBA Project name: " & vbaProjectName
    End If
    On Error GoTo EH
    
    Debug.Print "[SaveHostAsXLSM] About to disable alerts and attempt SaveAs..."
    Application.DisplayAlerts = False
    
    ' Attempt the save with detailed error capture
    On Error Resume Next
    wb.SaveAs fileName:=targetPath, FileFormat:=xlOpenXMLWorkbookMacroEnabled
    
    Dim saveError As Long
    Dim saveErrorDesc As String
    saveError = Err.Number
    saveErrorDesc = Err.Description
    
    Application.DisplayAlerts = True
    
    If saveError <> 0 Then
        Debug.Print "[SaveHostAsXLSM] SaveAs FAILED!"
        Debug.Print "[SaveHostAsXLSM] Error Number: " & saveError
        Debug.Print "[SaveHostAsXLSM] Error Description: " & saveErrorDesc
        Debug.Print "[SaveHostAsXLSM] FileFormat constant value: " & xlOpenXMLWorkbookMacroEnabled
        Err.Raise saveError, "SaveHostAsXLSM", saveErrorDesc
    End If
    
    On Error GoTo EH
    
    Debug.Print "[SaveHostAsXLSM] SaveAs succeeded!"
    Debug.Print "[SaveHostAsXLSM] New workbook FullName: " & wb.FullName
    
    SaveHostAsXLSM = True
    Exit Function
    
EH:
    Application.DisplayAlerts = True
    Debug.Print "[SaveHostAsXLSM] === ERROR HANDLER TRIGGERED ==="
    Debug.Print "[SaveHostAsXLSM] Error Number: " & Err.Number
    Debug.Print "[SaveHostAsXLSM] Error Description: " & Err.Description
    Debug.Print "[SaveHostAsXLSM] Error Source: " & Err.Source
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
' STEP 5 - RESOURCE EXTRACTION (UNIFIED WITH UPDATE.BAS)
' ==============================================================

Public Function ExtractResourcesWithVerification(fso As Object, rootPath As String, ByRef expectedCount As Long) As Long
    On Error GoTo EH

    Dim wbSource As Workbook
    Set wbSource = ThisWorkbook

    Dim wsStore As Worksheet
    On Error Resume Next
    Set wsStore = wbSource.Worksheets(EMBED_SHEET_NAME)
    On Error GoTo EH

    If wsStore Is Nothing Then
        LogMessage "ERROR", "Extraction", "No EmbeddedStore sheet found in " & wbSource.name
        ExtractResourcesWithVerification = 0
        Exit Function
    End If

    LogMessage "INFO", "Extraction", "Starting resource extraction from " & wbSource.name

    ' Ensure trailing slash on root path
    Dim outRoot As String
    outRoot = rootPath
    If right$(outRoot, 1) <> "\" Then outRoot = outRoot & "\"

    ' Extract and return count
    ExtractResourcesWithVerification = ExtractEmbeddedStoreUnified(wsStore, outRoot, expectedCount, fso)

    LogMessage "INFO", "Extraction", "Completed extraction to: " & outRoot
    Exit Function

EH:
    LogMessage "ERROR", "Extraction", "Error: " & Err.Description
    ExtractResourcesWithVerification = 0
End Function

Private Function ExtractEmbeddedStoreUnified(wsStore As Worksheet, outRoot As String, ByRef expectedCount As Long, fso As Object) As Long
    Dim lastRow As Long
    Dim r As Long
    Dim extractedCount As Long: extractedCount = 0

    Dim fileMap As Object
    Dim fileKey As String
    Dim fName As String
    Dim relPath As String
    Dim chunkIndex As Long
    Dim b64 As String

    lastRow = wsStore.Cells(wsStore.rows.count, COL_FILENAME).End(xlUp).Row
    If lastRow < 2 Then
        expectedCount = 0
        ExtractEmbeddedStoreUnified = 0
        Exit Function
    End If

    Set fileMap = CreateObject("Scripting.Dictionary")

    ' --- PASS 1: MAP CHUNKS (Unified with Update.bas logic) ---
    Dim arr As Variant
    arr = wsStore.Range(wsStore.Cells(2, 1), wsStore.Cells(lastRow, 4)).value

    For r = 1 To UBound(arr, 1)
        fName = CStr(arr(r, COL_FILENAME))
        relPath = CStr(arr(r, COL_RELPATH))

        If Len(fName) > 0 Or Len(relPath) > 0 Then
            ' Use relPath|fName as key (matches Update.bas)
            fileKey = relPath & "|" & fName

            If Not fileMap.Exists(fileKey) Then
                fileMap.Add fileKey, CreateObject("Scripting.Dictionary")
            End If

            chunkIndex = CLng(arr(r, COL_CHUNKINDEX))
            b64 = CStr(arr(r, COL_BASE64))

            fileMap(fileKey)(chunkIndex) = b64
        End If
    Next r

    expectedCount = fileMap.count

    ' --- PASS 2: REBUILD AND WRITE (Unified path construction with Update.bas) ---
    Dim k As Variant
    Dim parts() As String
    Dim chunksDict As Object
    Dim i As Long
    Dim bigB64 As String
    Dim bytes() As Byte
    Dim fullPath As String
    Dim folderPath As String

    ' Progress variables
    Dim totalFiles As Long: totalFiles = fileMap.count
    Dim currentFile As Long: currentFile = 0
    Dim startPct As Double: startPct = 0.5
    Dim endPct As Double: endPct = 0.7
    Dim calcPct As Double

    For Each k In fileMap.keys
        currentFile = currentFile + 1

        parts = Split(CStr(k), "|")
        ' parts(0) = relPath (e.g., "Python\Scripts" or "Python\Scripts\tools.py")
        ' parts(1) = fName (e.g., "tools.py")

        ' Update Progress
        calcPct = startPct + ((currentFile / totalFiles) * (endPct - startPct))
        UpdateProgress calcPct, "Extracting: " & parts(1)

        ' Assemble chunks in order (matches Update.bas logic)
        Set chunksDict = fileMap(k)
        bigB64 = ""
        For i = 0 To chunksDict.count - 1
            If chunksDict.Exists(CLng(i)) Then
                bigB64 = bigB64 & chunksDict(i)
            End If
        Next i

        ' Build full path (UNIFIED with Update.bas: rootPath\relPath\fileName)
        ' Handle case where relPath might already contain the filename
        If InStr(parts(0), parts(1)) > 0 Then
            ' RelPath contains filename, use it directly
            fullPath = outRoot & parts(0)
        Else
            ' RelPath is folder only, append filename
            fullPath = outRoot & parts(0) & "\" & parts(1)
        End If
        fullPath = Replace(fullPath, "\\", "\")

        ' Ensure parent folder exists
        folderPath = fso.GetParentFolderName(fullPath)
        If Len(folderPath) > 0 And Not fso.FolderExists(folderPath) Then
            CreateFoldersRecursive folderPath
        End If

        ' Decode and write
        On Error Resume Next
        bytes = Base64ToBinary(bigB64)
        If Err.Number = 0 Then
            WriteBinaryFile fullPath, bytes
            If Err.Number = 0 Then
                extractedCount = extractedCount + 1
                LogMessage "INFO", "Extract", "OK: " & fullPath
            Else
                LogMessage "ERROR", "Extract", "Write failed: " & fullPath & " - " & Err.Description
                SetupStats("FilesFailed") = SetupStats("FilesFailed") + 1
                Err.Clear
            End If
        Else
            LogMessage "ERROR", "Extract", "Decode failed: " & fullPath & " - " & Err.Description
            SetupStats("FilesFailed") = SetupStats("FilesFailed") + 1
            Err.Clear
        End If
        On Error GoTo 0
    Next k

    ExtractEmbeddedStoreUnified = extractedCount
End Function

' Legacy wrapper for compatibility
Public Sub ExtractResources(fso As Object, rootPath As String)
    Dim dummy As Long
    ExtractResourcesWithVerification fso, rootPath, dummy
End Sub


' ==============================================================
' STEP 6 - PIP INSTALL (WITH ENCODING FIX AND VERIFICATION)
' ==============================================================

Private Function InstallPipPackages(targetPath As String) As Boolean
    On Error GoTo EH

    Dim venvPy As String
    Dim reqFile As String
    Dim fixedReqFile As String
    Dim cmd As String
    Dim sh As Object
    Dim exitCode As Long

    ' Construct paths
    venvPy = JoinPath(targetPath, "Python\.venv\Scripts\python.exe")
    reqFile = JoinPath(targetPath, "Python\requirements.txt")
    fixedReqFile = JoinPath(targetPath, "Temp\requirements_fixed.txt")

    ' Verify python executable exists
    If Len(Dir$(venvPy, vbNormal)) = 0 Then
        LogMessage "ERROR", "Pip Install", "Python executable not found at " & venvPy
        InstallPipPackages = False
        Exit Function
    End If
    LogMessage "INFO", "Pip Install", "Python found: " & venvPy

    ' Verify requirements file exists (try multiple case variations)
    If Len(Dir$(reqFile, vbNormal)) = 0 Then
        reqFile = JoinPath(targetPath, "Python\Requirements.txt")
        If Len(Dir$(reqFile, vbNormal)) = 0 Then
            LogMessage "ERROR", "Pip Install", "Requirements file not found at " & reqFile
            InstallPipPackages = False
            Exit Function
        End If
    End If
    LogMessage "INFO", "Pip Install", "Requirements file found: " & reqFile

    ' FIX ENCODING: Convert UTF-16 to ANSI if needed
    If Not FixRequirementsEncoding(reqFile, fixedReqFile) Then
        LogMessage "ERROR", "Pip Install", "Failed to fix requirements file encoding"
        InstallPipPackages = False
        Exit Function
    End If
    LogMessage "INFO", "Pip Install", "Requirements file encoding verified/fixed"

    ' Upgrade pip first
    LogMessage "INFO", "Pip Install", "Upgrading pip..."
    cmd = "cmd /c """ & venvPy & """ -m pip install --upgrade pip --no-input"
    Set sh = CreateObject("WScript.Shell")
    exitCode = sh.Run(cmd, 0, True)
    LogMessage "INFO", "Pip Install", "Pip upgrade exit code: " & exitCode

    ' Install packages from fixed requirements
    LogMessage "INFO", "Pip Install", "Installing packages from requirements..."
    cmd = "cmd /c """ & venvPy & """ -m pip install -r """ & fixedReqFile & """ --no-input"
    exitCode = sh.Run(cmd, 0, True)

    LogMessage "INFO", "Pip Install", "Pip install exit code: " & exitCode

    ' Exit code 0 = success
    If exitCode = 0 Then
        LogMessage "INFO", "Pip Install", "pip install completed successfully"
        InstallPipPackages = True
    Else
        LogMessage "ERROR", "Pip Install", "pip install failed with exit code " & exitCode
        InstallPipPackages = False
    End If
    Exit Function

EH:
    LogMessage "ERROR", "Pip Install", "Exception: " & Err.Description
    InstallPipPackages = False
End Function

' Legacy wrapper for compatibility
Private Sub install_pip_Packages(targetPath As String)
    InstallPipPackages targetPath
End Sub

' ==============================================================
' ENCODING FIX FOR REQUIREMENTS.TXT
' ==============================================================

Private Function FixRequirementsEncoding(sourcePath As String, destPath As String) As Boolean
    On Error GoTo EH

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Ensure Temp folder exists
    Dim tempFolder As String
    tempFolder = fso.GetParentFolderName(destPath)
    If Not fso.FolderExists(tempFolder) Then
        Call EnsureFolderExists(tempFolder)
    End If

    ' Read source file as binary to detect encoding
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")

    stm.Type = 1 ' adTypeBinary
    stm.Open
    stm.LoadFromFile sourcePath

    Dim bytes() As Byte
    bytes = stm.Read
    stm.Close

    ' Detect UTF-16 BOM (FF FE or FE FF) or null bytes pattern
    Dim isUTF16 As Boolean
    isUTF16 = False

    If UBound(bytes) >= 1 Then
        ' Check for BOM
        If (bytes(0) = &HFF And bytes(1) = &HFE) Or (bytes(0) = &HFE And bytes(1) = &HFF) Then
            isUTF16 = True
            LogMessage "INFO", "Encoding", "Detected UTF-16 BOM in requirements.txt"
        ' Check for null byte pattern (UTF-16 without BOM)
        ElseIf UBound(bytes) >= 3 Then
            If bytes(1) = 0 Or bytes(3) = 0 Then
                isUTF16 = True
                LogMessage "INFO", "Encoding", "Detected UTF-16 (no BOM) in requirements.txt"
            End If
        End If
    End If

    If isUTF16 Then
        ' Convert UTF-16 to UTF-8
        LogMessage "INFO", "Encoding", "Converting UTF-16 to UTF-8..."

        Dim stmIn As Object, stmOut As Object
        Set stmIn = CreateObject("ADODB.Stream")
        Set stmOut = CreateObject("ADODB.Stream")

        ' Read as UTF-16
        stmIn.Type = 2 ' adTypeText
        stmIn.Charset = "unicode" ' UTF-16 LE
        stmIn.Open
        stmIn.LoadFromFile sourcePath

        Dim textContent As String
        textContent = stmIn.ReadText
        stmIn.Close

        ' Write as UTF-8 (without BOM for pip compatibility)
        stmOut.Type = 2
        stmOut.Charset = "utf-8"
        stmOut.Open
        stmOut.WriteText textContent

        ' Remove UTF-8 BOM by copying to binary stream
        stmOut.Position = 0
        stmOut.Type = 1 ' switch to binary

        ' Skip BOM if present (first 3 bytes for UTF-8 BOM)
        Dim outBytes() As Byte
        stmOut.Position = 3 ' Skip BOM
        If stmOut.Size > 3 Then
            outBytes = stmOut.Read
        Else
            stmOut.Position = 0
            outBytes = stmOut.Read
        End If
        stmOut.Close

        ' Write clean UTF-8 file
        Dim stmFinal As Object
        Set stmFinal = CreateObject("ADODB.Stream")
        stmFinal.Type = 1
        stmFinal.Open
        stmFinal.Write outBytes
        stmFinal.SaveToFile destPath, 2 ' adSaveCreateOverWrite
        stmFinal.Close

        LogMessage "INFO", "Encoding", "Converted file saved to " & destPath
    Else
        ' File is already in correct encoding, just copy
        LogMessage "INFO", "Encoding", "File encoding OK, copying to " & destPath
        fso.CopyFile sourcePath, destPath, True
    End If

    FixRequirementsEncoding = True
    Exit Function

EH:
    LogMessage "ERROR", "Encoding", "Failed to fix encoding: " & Err.Description
    FixRequirementsEncoding = False
End Function

' ==============================================================
' STEP 7 - VERIFICATION FUNCTIONS
' ==============================================================

Private Function VerifyExtractedFiles(fso As Object, rootPath As String) As Boolean
    On Error GoTo EH

    ' List of critical files that MUST exist after extraction
    Dim criticalFiles As Variant
    criticalFiles = Array( _
        "Python\Scripts\tools.py", _
        "Python\Scripts\xmlParsing.py", _
        "Python\requirements.txt" _
    )

    Dim filePath As String
    Dim missingFiles As String
    Dim i As Long
    Dim allFound As Boolean: allFound = True

    For i = LBound(criticalFiles) To UBound(criticalFiles)
        filePath = JoinPath(rootPath, CStr(criticalFiles(i)))

        ' Try both case variations
        If Len(Dir$(filePath, vbNormal)) = 0 Then
            ' Try with different case for requirements
            If InStr(LCase(filePath), "requirements") > 0 Then
                filePath = Replace(filePath, "requirements", "Requirements")
                If Len(Dir$(filePath, vbNormal)) = 0 Then
                    filePath = Replace(filePath, "Requirements", "requirements")
                End If
            End If
        End If

        If Len(Dir$(filePath, vbNormal)) = 0 Then
            LogMessage "ERROR", "File Verify", "Missing critical file: " & criticalFiles(i)
            missingFiles = missingFiles & vbCrLf & "  - " & criticalFiles(i)
            allFound = False
        Else
            LogMessage "INFO", "File Verify", "OK: " & criticalFiles(i)
        End If
    Next i

    ' Also verify venv structure
    Dim venvPy As String
    venvPy = JoinPath(rootPath, "Python\.venv\Scripts\python.exe")
    If Len(Dir$(venvPy, vbNormal)) = 0 Then
        LogMessage "ERROR", "File Verify", "Missing Python venv executable"
        allFound = False
    Else
        LogMessage "INFO", "File Verify", "OK: Python venv executable"
    End If

    If Not allFound Then
        LogMessage "ERROR", "File Verify", "Missing files:" & missingFiles
        SetupStats("Warnings") = SetupStats("Warnings") + 1
    End If

    VerifyExtractedFiles = allFound
    Exit Function

EH:
    LogMessage "ERROR", "File Verify", "Exception: " & Err.Description
    VerifyExtractedFiles = False
End Function

Private Function VerifyPipPackages(rootPath As String, ByRef installedCount As Long, ByRef requiredCount As Long) As Boolean
    On Error GoTo EH

    Dim venvPy As String
    Dim reqFile As String
    Dim fixedReqFile As String
    Dim fso As Object

    Set fso = CreateObject("Scripting.FileSystemObject")

    venvPy = JoinPath(rootPath, "Python\.venv\Scripts\python.exe")
    fixedReqFile = JoinPath(rootPath, "Temp\requirements_fixed.txt")

    ' Use the fixed requirements file if it exists
    If fso.fileExists(fixedReqFile) Then
        reqFile = fixedReqFile
    Else
        reqFile = JoinPath(rootPath, "Python\requirements.txt")
        If Not fso.fileExists(reqFile) Then
            reqFile = JoinPath(rootPath, "Python\Requirements.txt")
        End If
    End If

    ' Parse requirements to get expected packages
    Dim requiredPackages As Object
    Set requiredPackages = CreateObject("Scripting.Dictionary")

    If fso.fileExists(reqFile) Then
        Dim ts As Object
        Set ts = fso.OpenTextFile(reqFile, 1, False)
        Dim line As String
        Dim pkgName As String
        Dim eqPos As Long

        Do While Not ts.AtEndOfStream
            line = Trim(ts.ReadLine)
            ' Skip empty lines and comments
            If Len(line) > 0 And Left(line, 1) <> "#" Then
                ' Extract package name (before ==, >=, <=, etc.)
                eqPos = InStr(line, "==")
                If eqPos = 0 Then eqPos = InStr(line, ">=")
                If eqPos = 0 Then eqPos = InStr(line, "<=")
                If eqPos = 0 Then eqPos = InStr(line, ">")
                If eqPos = 0 Then eqPos = InStr(line, "<")

                If eqPos > 0 Then
                    pkgName = LCase(Trim(Left(line, eqPos - 1)))
                Else
                    pkgName = LCase(Trim(line))
                End If

                ' Normalize package names (replace _ with -)
                pkgName = Replace(pkgName, "_", "-")

                If Len(pkgName) > 0 Then
                    requiredPackages(pkgName) = True
                End If
            End If
        Loop
        ts.Close
    End If

    requiredCount = requiredPackages.count
    LogMessage "INFO", "Pip Verify", "Required packages: " & requiredCount

    ' Get installed packages via pip list
    Dim tempFile As String
    tempFile = JoinPath(rootPath, "Temp\pip_list.txt")

    Dim cmd As String
    Dim sh As Object
    Set sh = CreateObject("WScript.Shell")

    cmd = "cmd /c """ & venvPy & """ -m pip list --format=freeze > """ & tempFile & """"
    sh.Run cmd, 0, True

    ' Parse pip list output
    Dim installedPackages As Object
    Set installedPackages = CreateObject("Scripting.Dictionary")

    If fso.fileExists(tempFile) Then
        Set ts = fso.OpenTextFile(tempFile, 1, False)

        Do While Not ts.AtEndOfStream
            line = Trim(ts.ReadLine)
            If Len(line) > 0 Then
                eqPos = InStr(line, "==")
                If eqPos > 0 Then
                    pkgName = LCase(Trim(Left(line, eqPos - 1)))
                    pkgName = Replace(pkgName, "_", "-")
                    installedPackages(pkgName) = True
                End If
            End If
        Loop
        ts.Close
    End If

    LogMessage "INFO", "Pip Verify", "Installed packages: " & installedPackages.count

    ' Check each required package
    Dim missingPackages As String
    Dim pkg As Variant
    installedCount = 0

    For Each pkg In requiredPackages.keys
        If installedPackages.Exists(CStr(pkg)) Then
            installedCount = installedCount + 1
        Else
            ' Try with underscores instead of hyphens
            Dim altPkg As String
            altPkg = Replace(CStr(pkg), "-", "_")
            If installedPackages.Exists(altPkg) Then
                installedCount = installedCount + 1
            Else
                LogMessage "WARN", "Pip Verify", "Missing package: " & pkg
                missingPackages = missingPackages & vbCrLf & "  - " & pkg
            End If
        End If
    Next pkg

    LogMessage "INFO", "Pip Verify", "Verified " & installedCount & " of " & requiredCount & " packages"

    If installedCount < requiredCount Then
        LogMessage "WARN", "Pip Verify", "Missing packages:" & missingPackages
        SetupStats("PackagesFailed") = requiredCount - installedCount
        SetupStats("Warnings") = SetupStats("Warnings") + 1
        ' Allow some tolerance - if at least 80% are installed, consider it OK
        If installedCount >= (requiredCount * 0.8) Then
            LogMessage "WARN", "Pip Verify", "Continuing with partial install (80%+ packages present)"
            VerifyPipPackages = True
        Else
            VerifyPipPackages = False
        End If
    Else
        VerifyPipPackages = True
    End If
    Exit Function

EH:
    LogMessage "ERROR", "Pip Verify", "Exception: " & Err.Description
    VerifyPipPackages = False
End Function

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

' CreateFoldersRecursive - Creates full folder path recursively
Private Sub CreateFoldersRecursive(ByVal folderPath As String)
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FolderExists(folderPath) Then Exit Sub

    ' Get parent folder
    Dim parentPath As String
    parentPath = fso.GetParentFolderName(folderPath)

    ' Recurse to create parent first
    If Len(parentPath) > 0 And Not fso.FolderExists(parentPath) Then
        CreateFoldersRecursive parentPath
    End If

    ' Create this folder
    If Not fso.FolderExists(folderPath) Then
        fso.CreateFolder folderPath
    End If
End Sub

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




