Attribute VB_Name = "Update"
Option Explicit

' ==============================================================================
' CONFIGURATION & CONSTANTS
' ==============================================================================

' STORAGE SETTINGS
Private Const EMBED_SHEET_NAME As String = "EmbeddedStore"
Private Const WORKBOOK_VERSION_TAG As String = "PyExcel_ProjectVersion"
Private Const WORKBOOK_DECLINED_TAG As String = "PyExcel_UpdateDeclined"
Private Const PROPERTY_VERSION_TAG As String = "PyExcel_Version"

' UPDATE STATE (Module-level flags for ribbon indicator)
Public UpdateAvailable As Boolean
Public AvailableVersion As String

' SAFE ZONES (Never touch these folders, anywhere in the tree)
Private Const FOLDER_VENV As String = ".venv"
Private Const FOLDER_USER_SCRIPTS As String = "userScripts"

' COLUMN MAPPING (In EmbeddedStore)
Private Const COL_FILENAME As Long = 1
Private Const COL_CHUNKINDEX As Long = 2
Private Const COL_BASE64 As Long = 3
Private Const COL_RELPATH As Long = 4

' PROGRESS BAR GLOBAL
Public CurrentProgressForm As Object

' SESSION FLAG (Prevent repeated version checks)
Private VersionCheckedThisSession As Object

' ==============================================================================
' PART 1: ENTRY POINTS (VERSION CHECK & UPDATE TRIGGER)
' ==============================================================================

' ENTRY 1: AUTOMATIC CHECK (Call this on Workbook Open)
' Now sets UpdateAvailable flag instead of showing blocking MsgBox
Public Sub VerifyProjectVersion()
    On Error GoTo EH

    Dim wb As Workbook
    Dim projectVersion As String
    Dim addinVersion As String
    Dim declinedVersion As String

    Set wb = HostManager_GetCurrentWorkbook()
    If wb Is Nothing Then Exit Sub

    ' 0. Check if already verified this session (prevent repeated checks)
    If WasVersionCheckedThisSession(wb) Then Exit Sub
    MarkVersionCheckedThisSession wb

    ' 1. Get Versions
    projectVersion = GetStoredProjectVersion(wb)
    addinVersion = GetAddinVersion()
    declinedVersion = GetDeclinedVersion(wb)

    ' 2. Up to date - nothing to do
    If projectVersion = addinVersion Then
        UpdateAvailable = False
        AvailableVersion = ""
        RefreshEnableButton
        Exit Sub
    End If

    ' 3. Handle missing version (legacy project from before versioning)
    ' Assume outdated and offer update to ensure files are current
    If projectVersion = "" Then
        UpdateAvailable = True
        AvailableVersion = addinVersion
        Debug.Print "[VerifyProjectVersion] No version found - assuming legacy project needs update to " & addinVersion
        RefreshEnableButton
        Exit Sub
    End If

    ' 4. Check if update available (addin is newer than project)
    If VersionToNumber(addinVersion) > VersionToNumber(projectVersion) Then
        ' Check if user already declined THIS version
        If declinedVersion = addinVersion Then
            Debug.Print "[VerifyProjectVersion] User declined version " & addinVersion
            UpdateAvailable = False
            AvailableVersion = ""
            RefreshEnableButton
            Exit Sub
        End If

        ' Update is available! Set flag for ribbon indicator
        UpdateAvailable = True
        AvailableVersion = addinVersion
        Debug.Print "[VerifyProjectVersion] Update available: " & projectVersion & " -> " & addinVersion
        RefreshEnableButton
    Else
        ' Project version >= addin version (downgrade scenario), do nothing
        UpdateAvailable = False
        AvailableVersion = ""
        RefreshEnableButton
    End If
    Exit Sub

EH:
    Debug.Print "[VerifyProjectVersion] ERROR: " & Err.Description
End Sub

' ENTRY 2: AUTOMATIC UPDATE (Uses currently loaded addin - ThisWorkbook)
Public Sub RunUpdateFromCurrentAddin()
    On Error GoTo EH

    Dim fso As Object
    Dim targetPath As String
    Dim stepName As String

    ' 1. VALIDATE CONTEXT
    stepName = "Validating context"
    Debug.Print "[Update] Step: " & stepName
    Dim wbHost As Workbook
    Set wbHost = HostManager_GetCurrentWorkbook()
    If wbHost Is Nothing Then
        MsgBox "Please open your project workbook first.", vbExclamation
        Exit Sub
    End If

    ' Resolve SharePoint/OneDrive URLs to local paths
    targetPath = ResolveProjectPath()
    If Len(targetPath) = 0 Then
        MsgBox "Could not resolve project path. If using SharePoint/OneDrive, ensure the folder is synced locally.", vbExclamation
        Exit Sub
    End If
    Debug.Print "[Update] Resolved target path: " & targetPath

    ' 2. EXECUTE UPDATE FROM ThisWorkbook (the active addin)
    stepName = "Initializing"
    Debug.Print "[Update] Step: " & stepName & " - Target: " & targetPath
    InitProgressBar
    Set fso = CreateObject("Scripting.FileSystemObject")

    UpdateProgress 0.1, "Analyzing current addin..."
    Application.ScreenUpdating = False

    ' A. RUN SMART CLEANER (Deletes obsolete files safely)
    stepName = "Smart clean"
    Debug.Print "[Update] Step: " & stepName
    UpdateProgress 0.2, "Cleaning obsolete files..."
    SmartCleanFolder fso, targetPath, ThisWorkbook

    ' B. EXTRACT RESOURCES FROM ThisWorkbook
    stepName = "Extract resources"
    Debug.Print "[Update] Step: " & stepName
    UpdateProgress 0.4, "Installing new files..."
    ExtractResources fso, targetPath, ThisWorkbook

    Application.ScreenUpdating = True

    ' C. UPDATE PYTHON (Pip Install + Freeze)
    stepName = "Update Python dependencies"
    Debug.Print "[Update] Step: " & stepName
    UpdateProgress 0.7, "Updating Python libraries..."
    UpdatePythonDependencies targetPath

    ' D. UPDATE PROJECT TAG
    stepName = "Set project version"
    Debug.Print "[Update] Step: " & stepName
    Dim newVer As String
    newVer = GetAddinVersion() ' Get version from ThisWorkbook
    SetStoredProjectVersion wbHost, newVer

    ' Ensure structure integrity
    stepName = "Ensure structure"
    Debug.Print "[Update] Step: " & stepName
    ReEnsureStructure fso, targetPath

    UpdateProgress 1#, "Update Complete!"
    Application.Wait Now + TimeValue("0:00:01")
    CloseProgressBar

    Debug.Print "[Update] SUCCESS - Updated to version " & newVer
    MsgBox "Project successfully updated to version " & newVer, vbInformation
    Exit Sub

EH:
    Dim errMsg As String
    errMsg = Err.Description
    If Len(errMsg) = 0 Then errMsg = "Unknown error (Error " & Err.Number & ")"
    Debug.Print "[Update] FAILED at step: " & stepName & " - " & errMsg
    Application.ScreenUpdating = True
    CloseProgressBar
    MsgBox "Update Failed at '" & stepName & "':" & vbCrLf & vbCrLf & errMsg, vbCritical
End Sub

' ENTRY 3: MANUAL UPDATE (Call this from Ribbon - prompts for external file)
Public Sub RunUpdateFromExternalFile()
    On Error GoTo EH
    
    Dim wbCurrent As Workbook: Set wbCurrent = ThisWorkbook
    Dim wbNew As Workbook
    Dim fso As Object
    Dim targetPath As String
    Dim newXlamPath As String
    Dim fd As FileDialog
    
    ' 1. VALIDATE CONTEXT
    Dim wbHost As Workbook
    Set wbHost = HostManager_GetCurrentWorkbook()
    If wbHost Is Nothing Then
        MsgBox "Please open your project workbook first.", vbExclamation
        Exit Sub
    End If

    ' Resolve SharePoint/OneDrive URLs to local paths
    targetPath = ResolveProjectPath()
    If Len(targetPath) = 0 Then
        MsgBox "Could not resolve project path. If using SharePoint/OneDrive, ensure the folder is synced locally.", vbExclamation
        Exit Sub
    End If

    ' 2. SELECT UPDATE FILE
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Select the NEW version of the Add-in (XLAM)"
        .Filters.Clear
        .Filters.Add "Excel Add-in", "*.xlam"
        .AllowMultiSelect = False
        If .Show <> -1 Then Exit Sub
        newXlamPath = .SelectedItems(1)
    End With
    
    ' Safety: Prevent circular update
    If LCase(newXlamPath) = LCase(wbCurrent.FullName) Then
        MsgBox "You selected the currently installed file. Please select the downloaded update file.", vbExclamation
        Exit Sub
    End If

    ' 3. EXECUTE UPDATE
    InitProgressBar
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    UpdateProgress 0.1, "Analyzing new version..."
    Application.ScreenUpdating = False
    
    ' Open source file Read-Only
    Set wbNew = Workbooks.Open(fileName:=newXlamPath, ReadOnly:=True)
    
    ' A. RUN SMART CLEANER (Deletes obsolete files safely)
    UpdateProgress 0.2, "Cleaning obsolete files..."
    SmartCleanFolder fso, targetPath, wbNew
    
    ' B. EXTRACT RESOURCES
    UpdateProgress 0.4, "Installing new files..."
    ExtractResources fso, targetPath, wbNew
    
    ' Close the source file immediately
    Dim newVer As String
    newVer = GetAddinVersionFromFile(wbNew) ' Grab version before closing
    wbNew.Close SaveChanges:=False
    Application.ScreenUpdating = True
    
    ' C. UPDATE PYTHON (Pip Install + Freeze)
    UpdateProgress 0.7, "Updating Python libraries..."
    UpdatePythonDependencies targetPath
    
    ' D. UPDATE PROJECT TAG
    SetStoredProjectVersion wbHost, newVer
    
    ' Ensure structure integrity
    ReEnsureStructure fso, targetPath
    
    UpdateProgress 1#, "Update Complete!"
    Application.Wait Now + TimeValue("0:00:01")
    CloseProgressBar
    
    MsgBox "Project successfully updated to version " & newVer, vbInformation
    Exit Sub

EH:
    Dim errMsg As String
    errMsg = Err.Description
    If Len(errMsg) = 0 Then errMsg = "Unknown error (Error " & Err.Number & ")"
    Application.ScreenUpdating = True
    If Not wbNew Is Nothing Then wbNew.Close SaveChanges:=False
    CloseProgressBar
    MsgBox "Update Failed: " & errMsg, vbCritical
End Sub


' ==============================================================================
' PART 2: SMART CLEANER (MANIFEST-DRIVEN WITH SAFE ZONES)
' ==============================================================================

Private Sub SmartCleanFolder(fso As Object, rootPath As String, wbSource As Workbook)
    ' Get Manifest from New XLAM
    Dim manifest As Object
    Set manifest = LoadManifest(wbSource)

    ' Derive which top-level folders the manifest "owns"
    Dim ownedFolders As Object
    Set ownedFolders = GetOwnedFolders(manifest)

    ' Clean each owned folder
    Dim folderName As Variant
    Dim folderPath As String

    For Each folderName In ownedFolders.keys
        folderPath = rootPath & "\" & folderName

        If fso.FolderExists(folderPath) Then
            Debug.Print "[Cleaner] Cleaning owned folder: " & folderName
            CleanRecursive fso, fso.GetFolder(folderPath), rootPath, manifest
        End If
    Next folderName
End Sub

' Extract unique top-level folders from manifest RelPaths
Private Function GetOwnedFolders(manifest As Object) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare

    Dim key As Variant
    Dim parts() As String
    Dim topFolder As String

    For Each key In manifest.keys
        ' key is like "PYTHON\SCRIPTS\TOOLS.PY" or "ADDIN\LOGO.PNG"
        parts = Split(CStr(key), "\")

        If UBound(parts) >= 1 Then
            ' Has at least one folder level
            topFolder = parts(0)

            ' Skip if it's a safe zone at top level
            If Not IsSafeZone(topFolder) Then
                If Not d.Exists(topFolder) Then
                    d.Add topFolder, True
                End If
            End If
        End If
    Next key

    Set GetOwnedFolders = d
End Function

' Check if folder name is a protected safe zone
Private Function IsSafeZone(folderName As String) As Boolean
    Dim uName As String
    uName = UCase(folderName)

    Select Case uName
        Case UCase(FOLDER_VENV), UCase(FOLDER_USER_SCRIPTS)
            IsSafeZone = True
        Case Else
            IsSafeZone = False
    End Select
End Function

Private Sub CleanRecursive(fso As Object, fldr As Object, rootPath As String, manifest As Object)
    Dim file As Object
    Dim subFldr As Object
    Dim fName As String
    Dim relPath As String

    ' A. CHECK FILES
    For Each file In fldr.files
        relPath = GetRelativePath(file.path, rootPath)

        ' IF not in Manifest -> DELETE
        If Not manifest.Exists(relPath) Then
            Debug.Print "[Cleaner] Deleting: " & relPath
            On Error Resume Next
            file.Delete True
            On Error GoTo 0
        End If
    Next file

    ' B. RECURSE SUBFOLDERS
    For Each subFldr In fldr.subFolders
        fName = subFldr.name

        ' SAFETY: Skip protected folders anywhere in tree
        If IsSafeZone(fName) Then
            Debug.Print "[Cleaner] Skipping safe zone: " & fName
        ElseIf UCase(fName) = "__PYCACHE__" Then
            ' Nuke Pycache
            On Error Resume Next
            subFldr.Delete True
            On Error GoTo 0
        Else
            ' Recurse
            CleanRecursive fso, subFldr, rootPath, manifest

            ' Remove empty folders
            On Error Resume Next
            If subFldr.files.count = 0 And subFldr.subFolders.count = 0 Then
                subFldr.Delete True
            End If
            On Error GoTo 0
        End If
    Next subFldr
End Sub

Private Function LoadManifest(wb As Workbook) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(EMBED_SHEET_NAME)
    On Error GoTo 0
    
    If ws Is Nothing Then Set LoadManifest = d: Exit Function
    
    Dim arr As Variant
    Dim r As Long, lastRow As Long
    lastRow = ws.Cells(ws.rows.count, COL_FILENAME).End(xlUp).Row
    If lastRow < 2 Then Set LoadManifest = d: Exit Function
    
    arr = ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, 4)).value
    
    Dim key As String
    For r = 1 To UBound(arr, 1)
        ' Trust RelPath (Col 4) as the unique key, ignoring potentially bad FileName (Col 1)
        key = UCase(CStr(arr(r, COL_RELPATH)))
        
        ' Ensure standard path separators
        key = Replace(key, "/", "\")
        
        ' Remove leading backslash if present (RelPath should be clean)
        If Left(key, 1) = "\" Then key = Mid(key, 2)
        
        d(key) = True
    Next r
    Set LoadManifest = d
End Function

' ==============================================================================
' PART 3: EXTRACTION LOGIC
' ==============================================================================

Public Sub ExtractResources(fso As Object, rootPath As String, wbSource As Workbook)
    On Error Resume Next
    Dim wsStore As Worksheet
    Set wsStore = wbSource.Worksheets(EMBED_SHEET_NAME)
    On Error GoTo 0

    If wsStore Is Nothing Then
        Err.Raise vbObjectError + 1001, "ExtractResources", _
            "EmbeddedStore sheet not found in " & wbSource.name & ". The add-in may be corrupted or missing embedded files."
    End If
    
    ' 1. BUILD MAP
    Dim fileMap As Object: Set fileMap = CreateObject("Scripting.Dictionary")
    Dim lastRow As Long: lastRow = wsStore.Cells(wsStore.rows.count, COL_FILENAME).End(xlUp).Row
    Dim arr As Variant: arr = wsStore.Range(wsStore.Cells(2, 1), wsStore.Cells(lastRow, 4)).value
    
    Dim r As Long, key As String, b64 As String, chunkIdx As Long
    
    For r = 1 To UBound(arr, 1)
        key = arr(r, COL_RELPATH) & "|" & arr(r, COL_FILENAME)
        If Not fileMap.Exists(key) Then fileMap.Add key, CreateObject("Scripting.Dictionary")
        chunkIdx = CLng(arr(r, COL_CHUNKINDEX))
        b64 = CStr(arr(r, COL_BASE64))
        fileMap(key)(chunkIdx) = b64
    Next r
    
    ' 2. WRITE FILES
    Dim k As Variant, parts() As String, fullPath As String
    Dim i As Long, bigB64 As String, bytes() As Byte, chunks As Object
    Dim fileCount As Long: fileCount = 0

    For Each k In fileMap.keys
        fileCount = fileCount + 1
        parts = Split(k, "|")
        ' parts(0) = relPath (The correct relative path)
        ' parts(1) = fName (Ignored)

        fullPath = rootPath & "\" & parts(0)
        fullPath = Replace(fullPath, "\\", "\")

        Debug.Print "[Extract] File " & fileCount & ": " & fullPath
        
        EnsureFolderExists fso, fullPath
        
        Set chunks = fileMap(k)
        bigB64 = ""
        ' Assemble chunks (1-based to match Embedder)
        For i = 1 To chunks.count
            If chunks.Exists(CLng(i)) Then bigB64 = bigB64 & chunks(i)
        Next i
        
        bytes = Base64ToBinary(bigB64)
        WriteBinaryFile fullPath, bytes
    Next k
End Sub

' ==============================================================================
' PART 4: PYTHON DEPENDENCY MANAGER
' ==============================================================================

Public Sub UpdatePythonDependencies(rootPath As String)
    On Error GoTo EH
    Dim venvPy As String, reqFile As String, uninstallFile As String
    Dim snapFile As String, cmd As String
    Dim fso As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Define Paths
    venvPy = Quote(rootPath & "\Python\.venv\Scripts\python.exe")
    
    Dim rawReqPath As String
    rawReqPath = rootPath & "\Python\Requirements.txt"
    reqFile = Quote(rawReqPath)
    
    uninstallFile = rootPath & "\Python\Uninstall.txt" ' Keep unquoted for FSO check
    snapFile = Quote(rootPath & "\Python\User_Environment_Snapshot.txt")
    
    Dim fixedReqPath As String
    fixedReqPath = rootPath & "\Temp\requirements_fixed.txt"
    Dim fixedReqFile As String
    fixedReqFile = Quote(fixedReqPath)

    ' ---------------------------------------------------------
    ' STEP 0: EXPLICIT UNINSTALL (The Cleanup Phase)
    ' ---------------------------------------------------------
    ' Only run if file exists AND has content (> 0 bytes)
    If fso.fileExists(uninstallFile) Then
        If fso.GetFile(uninstallFile).Size > 0 Then
            UpdateProgress 0.65, "Removing deprecated libraries..."
            cmd = "cmd /c " & venvPy & " -m pip uninstall -r " & Quote(uninstallFile) & " -y"
            RunShellWait cmd
        End If
        
        ' Optional: Delete it afterwards regardless of size
        ' On Error Resume Next
        ' fso.DeleteFile uninstallFile
        ' On Error GoTo EH
    End If

    ' ---------------------------------------------------------
    ' STEP 1: PIP INSTALL (The Upgrade Phase)
    ' ---------------------------------------------------------
    If fso.fileExists(rawReqPath) Then
        ' Note: We run this AFTER uninstalling to ensure if a package
        ' is re-required by a dependency, it gets pulled back in.
        
        ' FIX ENCODING: Convert UTF-16 to ANSI/UTF-8 if needed
        If FixRequirementsEncoding(rawReqPath, fixedReqPath) Then
            Debug.Print "Pip Install: Using fixed requirements file: " & fixedReqPath
            cmd = "cmd /c " & venvPy & " -m pip install -r " & fixedReqFile & " --upgrade --no-input"
        Else
            Debug.Print "Pip Install: Warning - Encoding fix failed, using original file."
            cmd = "cmd /c " & venvPy & " -m pip install -r " & reqFile & " --upgrade --no-input"
        End If
        
        RunShellWait cmd
    End If
    
    ' ---------------------------------------------------------
    ' STEP 2: PIP FREEZE (Snapshot)
    ' ---------------------------------------------------------
    cmd = "cmd /c " & venvPy & " -m pip freeze > " & snapFile
    RunShellWait cmd
    Exit Sub

EH:
    Debug.Print "Pip Error: " & Err.Description
End Sub

' ==============================================================================
' PART 5: VERSIONING HELPERS (METADATA)
' ==============================================================================

' GETTER: Reads version from Document Properties (No hardcoded constant!)
Public Function GetAddinVersion() As String
    GetAddinVersion = GetAddinVersionFromFile(ThisWorkbook)
End Function

Public Function GetAddinVersionFromFile(wb As Workbook) As String
    On Error Resume Next
    Dim v As String
    v = wb.CustomDocumentProperties(PROPERTY_VERSION_TAG).value
    If v = "" Then v = "0.0.0"
    GetAddinVersionFromFile = v
    On Error GoTo 0
End Function

' SETTER: Run this manually (Immediate Window) to stamp a new version
Public Sub SetAddinVersion(NewVersion As String)
    Dim props As DocumentProperties
    Set props = ThisWorkbook.CustomDocumentProperties
    On Error Resume Next
    props(PROPERTY_VERSION_TAG).value = NewVersion
    If Err.Number <> 0 Then
        props.Add name:=PROPERTY_VERSION_TAG, LinkToContent:=False, _
                  Type:=msoPropertyTypeString, value:=NewVersion
    End If
    If Not ThisWorkbook.ReadOnly Then ThisWorkbook.Save
    Debug.Print "Version Stamped: " & NewVersion
End Sub

' READ/WRITE to User Workbook (Named Range)
Private Function GetStoredProjectVersion(wb As Workbook) As String
    On Error Resume Next
    Dim s As String
    s = wb.Names(WORKBOOK_VERSION_TAG).RefersTo
    s = Replace(Replace(s, "=", ""), """", "")
    GetStoredProjectVersion = s
End Function

Public Sub SetStoredProjectVersion(wb As Workbook, ver As String)
    On Error Resume Next
    wb.Names(WORKBOOK_VERSION_TAG).Delete
    wb.Names.Add name:=WORKBOOK_VERSION_TAG, RefersTo:="=""" & ver & """"

    ' Save the workbook to persist the version tag
    If Not wb.ReadOnly Then
        wb.Save
    End If
End Sub

' ==============================================================================
' PART 5B: DECLINED VERSION HELPERS
' ==============================================================================

' Read the declined version from workbook Named Range
Private Function GetDeclinedVersion(wb As Workbook) As String
    On Error Resume Next
    Dim s As String
    s = wb.Names(WORKBOOK_DECLINED_TAG).RefersTo
    s = Replace(Replace(s, "=", ""), """", "")
    GetDeclinedVersion = s
End Function

' Write the declined version to workbook Named Range
Public Sub SetDeclinedVersion(wb As Workbook, ver As String)
    On Error Resume Next
    wb.Names(WORKBOOK_DECLINED_TAG).Delete
    If Len(ver) > 0 Then
        wb.Names.Add name:=WORKBOOK_DECLINED_TAG, RefersTo:="=""" & ver & """"
    End If
    If Not wb.ReadOnly Then wb.Save
End Sub

' Clear the declined version (called when user manually triggers update)
Public Sub ClearDeclinedVersion(wb As Workbook)
    On Error Resume Next
    wb.Names(WORKBOOK_DECLINED_TAG).Delete
End Sub

' Refresh the Enable button in ribbon to show update state
Private Sub RefreshEnableButton()
    On Error Resume Next
    If Not rib Is Nothing Then rib.InvalidateControl "btnEnablePyExcel"
End Sub

' ==============================================================================
' PART 5C: MANUAL UPDATE ENTRY POINTS (Called from Ribbon)
' ==============================================================================

' ENTRY 4: MANUAL UPDATE FROM RIBBON (User clicks Update Available button)
Public Sub RunManualUpdate()
    On Error GoTo EH

    Dim wb As Workbook
    Set wb = HostManager_GetCurrentWorkbook()
    If wb Is Nothing Then
        MsgBox "Please open your project workbook first.", vbExclamation
        Exit Sub
    End If

    ' Clear declined flag when user manually requests update
    ClearDeclinedVersion wb

    ' Run the actual update
    RunUpdateFromCurrentAddin

    ' Clear update flag after successful update
    UpdateAvailable = False
    AvailableVersion = ""
    RefreshEnableButton
    Exit Sub

EH:
    Dim errMsg As String
    errMsg = Err.Description
    If Len(errMsg) = 0 Then errMsg = "Unknown error (Error " & Err.Number & ")"
    MsgBox "Update failed: " & errMsg, vbCritical
End Sub

' ENTRY 5: DISMISS UPDATE (User clicks No to dismiss update prompt)
Public Sub DismissUpdate()
    On Error GoTo EH

    Dim wb As Workbook
    Set wb = HostManager_GetCurrentWorkbook()
    If wb Is Nothing Then Exit Sub

    ' Store the declined version
    SetDeclinedVersion wb, GetAddinVersion()

    ' Clear the update flag
    UpdateAvailable = False
    AvailableVersion = ""
    RefreshEnableButton

    Debug.Print "[DismissUpdate] User declined version " & GetAddinVersion()
    Exit Sub

EH:
    Debug.Print "[DismissUpdate] ERROR: " & Err.Description
End Sub

' ==============================================================================
' PART 6: SESSION FLAG (PREVENT REPEATED CHECKS)
' ==============================================================================

Private Function GetCheckedSessionDict() As Object
    If VersionCheckedThisSession Is Nothing Then
        Set VersionCheckedThisSession = CreateObject("Scripting.Dictionary")
    End If
    Set GetCheckedSessionDict = VersionCheckedThisSession
End Function

Private Function WasVersionCheckedThisSession(wb As Workbook) As Boolean
    Dim d As Object: Set d = GetCheckedSessionDict()
    Dim key As String: key = wb.FullName
    WasVersionCheckedThisSession = d.Exists(key)
End Function

Private Sub MarkVersionCheckedThisSession(wb As Workbook)
    Dim d As Object: Set d = GetCheckedSessionDict()
    Dim key As String: key = wb.FullName
    If Not d.Exists(key) Then d.Add key, True
End Sub

' ==============================================================================
' PART 7: UTILITIES & UI
' ==============================================================================

Private Function VersionToNumber(v As String) As Double
    On Error Resume Next

    ' Handle timestamp format (yyyymmdd_hhnnss) from PackageAddin
    If InStr(v, "_") > 0 Then
        ' Remove underscore and convert: "20250201_143022" -> 20250201143022
        VersionToNumber = CDbl(Replace(v, "_", ""))
        Exit Function
    End If

    ' Handle semantic version format (1.2.3)
    Dim p() As String: p = Split(v, ".")
    Dim n As Double
    If UBound(p) >= 0 Then n = n + CDbl(p(0)) * 10000
    If UBound(p) >= 1 Then n = n + CDbl(p(1)) * 100
    If UBound(p) >= 2 Then n = n + CDbl(p(2))
    VersionToNumber = n
End Function

Private Sub InitProgressBar()
    On Error Resume Next
    Set CurrentProgressForm = New ufProgress
    CurrentProgressForm.lblBar.Width = 0
    CurrentProgressForm.Show vbModeless
    DoEvents
End Sub

Private Sub UpdateProgress(pct As Double, msg As String)
    If CurrentProgressForm Is Nothing Then Exit Sub
    CurrentProgressForm.lblStatus.Caption = msg
    Dim w As Double: w = CurrentProgressForm.fraBackground.InsideWidth
    If pct > 1 Then pct = 1
    CurrentProgressForm.lblBar.Width = w * pct
    CurrentProgressForm.Repaint
    DoEvents
End Sub

Private Sub CloseProgressBar()
    On Error Resume Next
    Unload CurrentProgressForm
    Set CurrentProgressForm = Nothing
End Sub

Private Sub EnsureFolderExists(fso As Object, filePath As String)
    Dim p As String: p = fso.GetParentFolderName(filePath)
    If Not fso.FolderExists(p) Then CreateFoldersRecursive fso, p
End Sub

Private Sub CreateFoldersRecursive(fso As Object, folderPath As String, Optional depth As Long = 0)
    ' Safety: prevent runaway recursion
    If depth > 50 Then
        Debug.Print "[CreateFolders] ERROR: Max depth exceeded at: " & folderPath
        Exit Sub
    End If

    ' Base case: empty or root path - stop recursion
    If Len(folderPath) = 0 Then
        Debug.Print "[CreateFolders] Base case: empty path"
        Exit Sub
    End If
    If fso.FolderExists(folderPath) Then Exit Sub

    Dim parentPath As String
    parentPath = fso.GetParentFolderName(folderPath)

    Debug.Print "[CreateFolders] Depth " & depth & ": " & folderPath & " (parent: " & parentPath & ")"

    ' Base case: no parent (at root like "C:\") or parent same as current
    If Len(parentPath) = 0 Or parentPath = folderPath Then
        Debug.Print "[CreateFolders] Base case: at root"
        On Error Resume Next
        fso.CreateFolder folderPath
        On Error GoTo 0
        Exit Sub
    End If

    ' Recurse to create parent first
    If Not fso.FolderExists(parentPath) Then
        CreateFoldersRecursive fso, parentPath, depth + 1
    End If

    ' Now create this folder
    On Error Resume Next
    fso.CreateFolder folderPath
    On Error GoTo 0
End Sub

Private Sub WriteBinaryFile(pth As String, dat() As Byte)
    Dim s As Object: Set s = CreateObject("ADODB.Stream")
    s.Type = 1: s.Open: s.Write dat: s.SaveToFile pth, 2: s.Close
End Sub

Private Function Base64ToBinary(s As String) As Byte()
    If Len(s) = 0 Then
        Base64ToBinary = StrConv("", vbFromUnicode)
        Exit Function
    End If

    Dim xml As Object: Set xml = CreateObject("MSXML2.DOMDocument")
    Dim el As Object: Set el = xml.createElement("b64")
    el.DataType = "bin.base64": el.text = s
    Base64ToBinary = el.nodeTypedValue
End Function

Private Function BuildPathKey(folderPart As String, filePart As String) As String
    Dim s As String: s = folderPart
    If right(s, 1) <> "\" And Len(s) > 0 Then s = s & "\"
    s = s & filePart
    BuildPathKey = UCase(s)
End Function

Private Function GetRelativePath(fullPath As String, rootPath As String) As String
    If InStr(1, fullPath, rootPath, vbTextCompare) = 1 Then
        Dim s As String
        s = Mid(fullPath, Len(rootPath) + 1)
        If Left(s, 1) = "\" Then s = Mid(s, 2)
        GetRelativePath = UCase(s)
    Else
        GetRelativePath = ""
    End If
End Function

Private Function Quote(s As String) As String
    Quote = """" & s & """"
End Function

Private Sub RunShellWait(cmd As String)
    CreateObject("WScript.Shell").Run cmd, 0, True
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
        Call EnsureFolderExists(fso, tempFolder)
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
            Debug.Print "INFO: Detected UTF-16 BOM in requirements.txt"
        ' Check for null byte pattern (UTF-16 without BOM)
        ElseIf UBound(bytes) >= 3 Then
            If bytes(1) = 0 Or bytes(3) = 0 Then
                isUTF16 = True
                Debug.Print "INFO: Detected UTF-16 (no BOM) in requirements.txt"
            End If
        End If
    End If

    If isUTF16 Then
        ' Convert UTF-16 to UTF-8
        Debug.Print "INFO: Converting UTF-16 to UTF-8..."

        Dim stmIn As Object, stmOut As Object
        Set stmIn = CreateObject("ADODB.Stream")
        Set stmOut = CreateObject("ADODB.Stream")

        ' Read as UTF-16
        stmIn.Type = 2 ' adTypeText
        stmIn.Charset = "unicode" ' UTF-16 LE
        stmIn.Open
        stmIn.LoadFromFile sourcePath

        Dim textContent As String
        If Not stmIn.EOS Then
            textContent = stmIn.ReadText
        Else
            textContent = ""
        End If
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
        Dim outBytes As Variant
        stmOut.Position = 3 ' Skip BOM
        If stmOut.Size > 3 Then
            outBytes = stmOut.Read
        Else
            stmOut.Position = 0
            outBytes = stmOut.Read
        End If
        stmOut.Close

        ' Force delete destination if exists
        If fso.fileExists(destPath) Then
            On Error Resume Next
            fso.DeleteFile destPath, True
            On Error GoTo EH
        End If

        Debug.Print "INFO: Saving converted file to: " & destPath

        ' Write clean UTF-8 file
        Dim stmFinal As Object
        Set stmFinal = CreateObject("ADODB.Stream")
        stmFinal.Type = 1
        stmFinal.Open
        stmFinal.Write outBytes
        stmFinal.SaveToFile destPath, 2 ' adSaveCreateOverWrite
        stmFinal.Close

        Debug.Print "INFO: Converted file saved to " & destPath
    Else
        ' File is already in correct encoding, just copy
        Debug.Print "INFO: File encoding OK, copying to " & destPath
        fso.CopyFile sourcePath, destPath, True
    End If

    FixRequirementsEncoding = True
    Exit Function

EH:
    Debug.Print "ERROR: Failed to fix encoding: " & Err.Description
    FixRequirementsEncoding = False
End Function

' ==============================================================
' FOLDER STRUCTURE HELPERS
' ==============================================================

Private Sub ReEnsureStructure(fso As Object, rootPath As String)
    On Error Resume Next
    Dim p As String

    ' Archive
    p = rootPath & "\Archive"
    If Not fso.FolderExists(p) Then fso.CreateFolder p
    CreatePlaceholder fso, p, "This folder contains archived versions of your Python scripts."

    ' userScripts
    p = rootPath & "\userScripts"
    If Not fso.FolderExists(p) Then fso.CreateFolder p
    CreatePlaceholder fso, p, "Place your Python scripts here."
End Sub

Private Sub CreatePlaceholder(fso As Object, folderPath As String, content As String)
    On Error Resume Next
    Dim filePath As String
    filePath = folderPath & "\ReadMe.txt"
    If Not fso.fileExists(filePath) Then
        Dim ts As Object
        Set ts = fso.CreateTextFile(filePath, True)
        ts.WriteLine content
        ts.Close
    End If
End Sub

