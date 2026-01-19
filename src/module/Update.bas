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

' FOLDER NAMES (SAFETY ZONES)
Private Const FOLDER_PYTHON As String = "Python"
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

    ' 3. Handle missing version (first time after enable)
    If projectVersion = "" Then
        SetStoredProjectVersion wb, addinVersion
        Debug.Print "[VerifyProjectVersion] Auto-stamped version " & addinVersion & " for " & wb.name
        UpdateAvailable = False
        AvailableVersion = ""
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

    ' 1. VALIDATE CONTEXT
    Dim wbHost As Workbook
    Set wbHost = HostManager_GetCurrentWorkbook()
    If wbHost Is Nothing Then
        MsgBox "Please open your project workbook first.", vbExclamation
        Exit Sub
    End If
    targetPath = wbHost.path

    ' 2. EXECUTE UPDATE FROM ThisWorkbook (the active addin)
    InitProgressBar
    Set fso = CreateObject("Scripting.FileSystemObject")

    UpdateProgress 0.1, "Analyzing current addin..."
    Application.ScreenUpdating = False

    ' A. RUN SMART CLEANER (Deletes obsolete files safely)
    UpdateProgress 0.2, "Cleaning obsolete files..."
    SmartCleanFolder fso, targetPath, ThisWorkbook

    ' B. EXTRACT RESOURCES FROM ThisWorkbook
    UpdateProgress 0.4, "Installing new files..."
    ExtractResources fso, targetPath, ThisWorkbook

    Application.ScreenUpdating = True

    ' C. UPDATE PYTHON (Pip Install + Freeze)
    UpdateProgress 0.7, "Updating Python libraries..."
    UpdatePythonDependencies targetPath

    ' D. UPDATE PROJECT TAG
    Dim newVer As String
    newVer = GetAddinVersion() ' Get version from ThisWorkbook
    SetStoredProjectVersion wbHost, newVer

    UpdateProgress 1#, "Update Complete!"
    Application.Wait Now + TimeValue("0:00:01")
    CloseProgressBar

    MsgBox "Project successfully updated to version " & newVer, vbInformation
    Exit Sub

EH:
    Application.ScreenUpdating = True
    CloseProgressBar
    MsgBox "Update Failed: " & Err.Description, vbCritical
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
    targetPath = wbHost.path
    
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
    
    UpdateProgress 1#, "Update Complete!"
    Application.Wait Now + TimeValue("0:00:01")
    CloseProgressBar
    
    MsgBox "Project successfully updated to version " & newVer, vbInformation
    Exit Sub

EH:
    Application.ScreenUpdating = True
    If Not wbNew Is Nothing Then wbNew.Close SaveChanges:=False
    CloseProgressBar
    MsgBox "Update Failed: " & Err.Description, vbCritical
End Sub


' ==============================================================================
' PART 2: SMART CLEANER (SAFETY LOGIC)
' ==============================================================================

Private Sub SmartCleanFolder(fso As Object, rootPath As String, wbSource As Workbook)
    ' Only target the Python folder to prevent root accidents
    Dim pythonPath As String
    pythonPath = rootPath & "\" & FOLDER_PYTHON
    
    If Not fso.FolderExists(pythonPath) Then Exit Sub
    
    ' Get Manifest from New XLAM
    Dim manifest As Object
    Set manifest = LoadManifest(wbSource)
    
    ' Recursively clean
    CleanRecursive fso, fso.GetFolder(pythonPath), rootPath, manifest
End Sub

Private Sub CleanRecursive(fso As Object, fldr As Object, rootPath As String, manifest As Object)
    Dim file As Object
    Dim subFldr As Object
    Dim fName As String
    Dim relPath As String
    
    ' A. CHECK FILES
    For Each file In fldr.files
        relPath = GetRelativePath(file.path, rootPath)
        
        ' IF not in Manifest AND not in Safe Zone -> DELETE
        If Not manifest.Exists(relPath) Then
             ' Double check we are deep enough (paranoid safety)
             If InStr(relPath, "\") > 0 Then
                Debug.Print "[Cleaner] Deleting Zombie: " & relPath
                On Error Resume Next
                file.Delete True
                On Error GoTo 0
             End If
        End If
    Next file
    
    ' B. RECURSE SUBFOLDERS
    For Each subFldr In fldr.subFolders
        fName = UCase(subFldr.name)
        
        ' SAFETY: DO NOT ENTER THESE FOLDERS
        If fName = UCase(FOLDER_VENV) Then
            ' Skip .venv
        ElseIf fName = UCase(FOLDER_USER_SCRIPTS) Then
            ' Skip userScripts
        ElseIf fName = "__PYCACHE__" Then
            ' Nuke Pycache
            On Error Resume Next
            subFldr.Delete True
            On Error GoTo 0
        Else
            ' Recurse
            CleanRecursive fso, subFldr, rootPath, manifest
            
            ' Optional: Remove empty folders
            If subFldr.files.count = 0 And subFldr.subFolders.count = 0 Then
                On Error Resume Next
                subFldr.Delete True
                On Error GoTo 0
            End If
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
        key = BuildPathKey(CStr(arr(r, COL_RELPATH)), CStr(arr(r, COL_FILENAME)))
        d(key) = True
    Next r
    Set LoadManifest = d
End Function

' ==============================================================================
' PART 3: EXTRACTION LOGIC
' ==============================================================================

Public Sub ExtractResources(fso As Object, rootPath As String, wbSource As Workbook)
    Dim wsStore As Worksheet
    Set wsStore = wbSource.Worksheets(EMBED_SHEET_NAME)
    
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
    
    For Each k In fileMap.keys
        parts = Split(k, "|")
        
        If Len(parts(1)) > 0 Then
            ' File entry
            fullPath = rootPath & "\" & parts(0) & "\" & parts(1)
            fullPath = Replace(fullPath, "\\", "\")
            
            EnsureFolderExists fso, fullPath
            
            Set chunks = fileMap(k)
            bigB64 = ""
            ' Assemble chunks
            For i = 0 To chunks.count - 1
                If chunks.Exists(CLng(i)) Then bigB64 = bigB64 & chunks(i)
            Next i
            
            bytes = Base64ToBinary(bigB64)
            WriteBinaryFile fullPath, bytes
        Else
            ' Folder-only entry
            fullPath = rootPath & "\" & parts(0)
            fullPath = Replace(fullPath, "\\", "\")
            
            ' Ensure the folder exists
            If Not fso.FolderExists(fullPath) Then
                CreateFoldersRecursive fso, fullPath
            End If
        End If
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
    reqFile = Quote(rootPath & "\Python\Requirements.txt")
    uninstallFile = rootPath & "\Python\Uninstall.txt" ' Keep unquoted for FSO check
    snapFile = Quote(rootPath & "\Python\User_Environment_Snapshot.txt")
    
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
    If fso.fileExists(rootPath & "\Python\Requirements.txt") Then
        ' Note: We run this AFTER uninstalling to ensure if a package
        ' is re-required by a dependency, it gets pulled back in.
        cmd = "cmd /c " & venvPy & " -m pip install -r " & reqFile & " --upgrade --no-input"
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
    MsgBox "Update failed: " & Err.Description, vbCritical
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

Private Sub CreateFoldersRecursive(fso As Object, folderPath As String)
    If Not fso.FolderExists(fso.GetParentFolderName(folderPath)) Then
        CreateFoldersRecursive fso, fso.GetParentFolderName(folderPath)
    End If
    If Not fso.FolderExists(folderPath) Then fso.CreateFolder folderPath
End Sub

Private Sub WriteBinaryFile(pth As String, dat() As Byte)
    Dim s As Object: Set s = CreateObject("ADODB.Stream")
    s.Type = 1: s.Open: s.Write dat: s.SaveToFile pth, 2: s.Close
End Sub

Private Function Base64ToBinary(s As String) As Byte()
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



