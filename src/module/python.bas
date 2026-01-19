Attribute VB_Name = "python"
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If



Sub RunGenericPythonScript(scriptName As String, srcRangeRef As String, dstRangeRef As String, wb As Workbook, ws As Worksheet)
    On Error GoTo Fail

    Dim srcRangeList As New Collection
    Dim part As Variant
    Dim addr As String
    Dim r As Range
    Dim rngItem As Range
    Dim tStart As Double
    Dim tStep As Double

    tStart = Timer
    tStep = tStart

    Debug.Print Format(Now, "hh:nn:ss"), "=== RunGenericPythonScript ==="
    Debug.Print "  t+", Format(Timer - tStart, "0.000"), " scriptName: [" & scriptName & "]"
    Debug.Print "  t+", Format(Timer - tStart, "0.000"), " srcRangeRef: [" & srcRangeRef & "]"
    Debug.Print "  t+", Format(Timer - tStart, "0.000"), " dstRangeRef: [" & dstRangeRef & "]"

    If wb Is Nothing Or ws Is Nothing Then
        MsgBox "No active workbook or sheet context.", vbExclamation
        Exit Sub
    End If

    If Len(scriptName) = 0 Then
        MsgBox "No script provided", vbExclamation
        Exit Sub
    End If

    If Len(srcRangeRef) = 0 Then
        Set r = Application.InputBox("Select source range", "Source Range", Type:=8)
        If r Is Nothing Then Exit Sub
        srcRangeRef = r.Address(External:=True)
    End If

    If Len(dstRangeRef) = 0 Then
        Set r = Application.InputBox("Select destination range", "Destination Range", Type:=8)
        If r Is Nothing Then Exit Sub
        dstRangeRef = r.Address(External:=True)
    End If

    Debug.Print "  t+", Format(Timer - tStep, "0.000"), " Starting to build srcRangeList"
    tStep = Timer

    For Each part In Split(srcRangeRef, ";")
        If InStr(part, "!") > 0 Then
            addr = Split(part, "!")(1)
            Set ws = wb.Sheets(Split(part, "!")(0))
        Else
            addr = part
            Set ws = HostManager_GetCurrentSheet()
        End If
        srcRangeList.Add ws.Range(addr)
    Next part

    Debug.Print "  t+", Format(Timer - tStep, "0.000"), " srcRangeList built"
    tStep = Timer

    Debug.Print "  t+", Format(Timer - tStep, "0.000"), " Clearing filters"
    For Each rngItem In srcRangeList
        ClearFiltersAffectingRange rngItem
    Next rngItem
    Debug.Print "  t+", Format(Timer - tStep, "0.000"), " Filters cleared"
    tStep = Timer

    Debug.Print "  t+", Format(Timer - tStep, "0.000"), " Calling Py()"
    Call Py(scriptName, srcRangeRef, dstRangeRef, wb, ws)
    Debug.Print "  t+", Format(Timer - tStep, "0.000"), " Py() completed"
    tStep = Timer

    Debug.Print "  t+", Format(Timer - tStart, "0.000"), "=== End RunGenericPythonScript ==="
    Exit Sub

Fail:
    Debug.Print "ERROR " & Err.Number & ": " & Err.Description, " at t+", Format(Timer - tStart, "0.000")
    MsgBox "Error in RunGenericPythonScript: " & Err.Description, vbCritical
End Sub




Private Sub ClearFiltersAffectingRange(ByVal rng As Range)
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim af As AutoFilter

    On Error GoTo done

    If Not rng.ListObject Is Nothing Then
        Set lo = rng.ListObject
        If Not lo.AutoFilter Is Nothing Then
            If lo.AutoFilter.FilterMode Then lo.AutoFilter.ShowAllData
        End If
        GoTo done
    End If

    Set ws = rng.parent
    If ws.AutoFilterMode Then
        Set af = ws.AutoFilter
        If Not af Is Nothing Then
            If Not Intersect(af.Range, rng) Is Nothing Then
                If ws.FilterMode Then ws.ShowAllData
            End If
        End If
    End If

done:
    On Error GoTo 0
End Sub






'Public Function Py(scriptName As String, srcRangeRef As String, dstRangeRef As String) As Boolean
'    On Error GoTo Fail
'
'    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
'    Dim wbPath As String, folderScripts As String, folderTemp As String, folderArchive As String
'    Dim inFile As String
'    Dim tempFiles As Object: Set tempFiles = CreateObject("Scripting.Dictionary")
'    Dim meta As Object
'
'    Debug.Print "=== Begin Py ==="
'    Debug.Print "Resolving project path..."
'    wbPath = ResolveProjectPath()
'    If wbPath = "" Then
'        Debug.Print "ResolveProjectPath failed."
'        Exit Function
'    End If
'
'    folderScripts = wbPath & "\Python\scripts"
'    folderTemp = wbPath & "\Temp"
'    folderArchive = wbPath & "\Archive"
'
'    If Not fso.FolderExists(folderTemp) Then
'        Debug.Print "Temp folder doesn't exist. Creating: " & folderTemp
'        fso.CreateFolder folderTemp
'        Debug.Print "Temp folder created."
'    Else
'        Debug.Print "Temp folder already exists."
'    End If
'
'    Debug.Print "Serializing range: " & srcRangeRef
'    inFile = SerializeRangeToTypedXML(srcRangeRef)
'    If inFile = "" Then
'        Debug.Print "SerializeRangeToTypedXML returned empty."
'        MsgBox "Failed to serialize input.", vbCritical
'        Exit Function
'    End If
'    Debug.Print "Input written to: " & inFile
'
'    tempFiles.Add "in", inFile
'
'    Debug.Print "Running Python script: " & scriptName
'    Set meta = RunPythonJob(scriptName, tempFiles)
'    If meta Is Nothing Then
'        Debug.Print "RunPythonJob returned nothing."
'        Exit Function
'    End If
'
'    If LCase$(CStr(meta("status"))) <> "done" Then
'        Dim msg As String
'        msg = "Python script failed."
'        If meta.Exists("message") Then msg = msg & vbCrLf & "Message: " & CStr(meta("message"))
'        If meta.Exists("stderr_log") Then msg = msg & vbCrLf & "Error log: " & CStr(meta("stderr_log"))
'        If meta.Exists("stdout_log") Then msg = msg & vbCrLf & "Output log: " & CStr(meta("stdout_log"))
'        MsgBox msg, vbCritical, "Python Error"
'        Exit Function
'    End If
'
'    If Not meta.Exists("artifacts") Then
'        Debug.Print "Meta XML missing <artifacts>."
'        Exit Function
'    End If
'
'
'    ' -------------------- inserting artifacts to destinations --------------------
'    Dim items As Object
'    Set items = meta("artifacts")  ' real artifacts from Python job
'
'    Dim defaultSheet As Worksheet
'    Set defaultSheet = Application.ThisWorkbook.ActiveSheet
Public Function Py(scriptName As String, srcRangeRef As String, dstRangeRef As String, wb As Workbook, ws As Worksheet) As Boolean
    On Error GoTo Fail

    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim wbPath As String, folderScripts As String, folderTemp As String, folderArchive As String
    Dim inFile As String
    Dim tempFiles As Object: Set tempFiles = CreateObject("Scripting.Dictionary")
    Dim meta As Object

    Dim tStart As Double
    Dim tStep As Double
    tStart = Timer
    tStep = tStart

    Debug.Print Format(Now, "hh:nn:ss"), "=== Begin Py ==="
    Debug.Print "t+", Format(Timer - tStart, "0.000"), " Resolving project path..."

    wbPath = ResolveProjectPath()
    If wbPath = "" Then
        Debug.Print "t+", Format(Timer - tStart, "0.000"), " ResolveProjectPath failed."
        Exit Function
    End If

    folderScripts = wbPath & "\Python\scripts"
    folderTemp = wbPath & "\Temp"
    folderArchive = wbPath & "\Archive"

    If Not fso.FolderExists(folderTemp) Then
        Debug.Print "t+", Format(Timer - tStep, "0.000"), " Temp folder doesn't exist. Creating: " & folderTemp
        fso.CreateFolder folderTemp
        Debug.Print "t+", Format(Timer - tStep, "0.000"), " Temp folder created."
    Else
        Debug.Print "t+", Format(Timer - tStep, "0.000"), " Temp folder already exists."
    End If
    tStep = Timer

    Debug.Print "t+", Format(Timer - tStep, "0.000"), " Serializing range: " & srcRangeRef
    inFile = SerializeRangeToTypedXML(srcRangeRef)
    Debug.Print "t+", Format(Timer - tStep, "0.000"), " SerializeRangeToTypedXML complete"

    If inFile = "" Then
        Debug.Print "t+", Format(Timer - tStart, "0.000"), " SerializeRangeToTypedXML returned empty."
        MsgBox "Failed to serialize input.", vbCritical
        Exit Function
    End If
    Debug.Print "t+", Format(Timer - tStart, "0.000"), " Input written to: " & inFile

    tempFiles.Add "in", inFile

    Debug.Print "t+", Format(Timer - tStep, "0.000"), " Running Python script: " & scriptName
    Set meta = RunPythonJob(scriptName, tempFiles)
    Debug.Print "t+", Format(Timer - tStep, "0.000"), " RunPythonJob complete"
    tStep = Timer

    If meta Is Nothing Then
        Debug.Print "t+", Format(Timer - tStart, "0.000"), " RunPythonJob returned nothing."
        Exit Function
    End If

    If LCase$(CStr(meta("status"))) <> "done" Then
        Dim msg As String
        msg = "Python script failed."
        If meta.Exists("message") Then msg = msg & vbCrLf & "Message: " & CStr(meta("message"))
        
        If meta.Exists("stderr") And Len(CStr(meta("stderr"))) > 0 Then
            msg = msg & vbCrLf & "--- Traceback ---" & vbCrLf & Left$(CStr(meta("stderr")), 800)
        End If
        
        If meta.Exists("stderr_log") Then msg = msg & vbCrLf & "Error log: " & CStr(meta("stderr_log"))
        If meta.Exists("stdout_log") Then msg = msg & vbCrLf & "Output log: " & CStr(meta("stdout_log"))
        Debug.Print "t+", Format(Timer - tStart, "0.000"), " Python script reported failure."
        MsgBox msg, vbCritical, "Python Error"
        Exit Function
    End If
    
    If Not meta.Exists("artifacts") Then
        Debug.Print "t+", Format(Timer - tStart, "0.000"), " Meta XML missing <artifacts>."
        Exit Function
    End If

    Debug.Print "t+", Format(Timer - tStep, "0.000"), " Parsing destination map"
    Dim items As Object
    Set items = meta("artifacts")
    Dim errors As New Collection
    Dim idMap As Object
    Set idMap = ParseIdToRangeMap(dstRangeRef, ws, items, errors)
    Debug.Print "t+", Format(Timer - tStep, "0.000"), " ParseIdToRangeMap complete"
    tStep = Timer
    
    If idMap Is Nothing Or idMap.count = 0 Then
        Dim e As Variant
        For Each e In errors
            Debug.Print "t+", Format(Timer - tStart, "0.000"), " dstSpec error: "; CStr(e)
        Next e
        Debug.Print "t+", Format(Timer - tStart, "0.000"), " No valid destination mappings."
        Exit Function
    End If

    Debug.Print "t+", Format(Timer - tStep, "0.000"), " Pasting artifacts to targets"
    If Not PasteArtifactsToTargets(idMap, items) Then
        Debug.Print "t+", Format(Timer - tStart, "0.000"), " No outputs were applied."
        Exit Function
    End If
    Debug.Print "t+", Format(Timer - tStep, "0.000"), " PasteArtifactsToTargets complete"
    tStep = Timer
    
    Debug.Print "t+", Format(Timer - tStep, "0.000"), " Cleaning up temporary files..."
    On Error GoTo CleanFail
    
    If Not fso.FolderExists(folderArchive) Then
        fso.CreateFolder folderArchive
        Debug.Print "t+", Format(Timer - tStep, "0.000"), " Archive folder created"
    End If
    
    Dim safeScriptName As String
    safeScriptName = scriptName
    safeScriptName = Replace(safeScriptName, "\", "_")
    safeScriptName = Replace(safeScriptName, "/", "_")
    safeScriptName = Replace(safeScriptName, ":", "_")
    safeScriptName = Replace(safeScriptName, "*", "_")
    safeScriptName = Replace(safeScriptName, "?", "_")
    safeScriptName = Replace(safeScriptName, """", "_")
    safeScriptName = Replace(safeScriptName, "<", "_")
    safeScriptName = Replace(safeScriptName, ">", "_")
    safeScriptName = Replace(safeScriptName, "|", "_")
    
    Dim runFolder As String, iTry As Long
    runFolder = folderArchive & "\" & Format$(Now, "dd-mm-yyyy-hh-nn-ss") & "_" & safeScriptName
    iTry = 1
    Do While Len(Dir$(runFolder, vbDirectory)) > 0
        runFolder = folderArchive & "\" & Format$(Now, "dd-mm-yyyy-hh-nn-ss") & "_" & safeScriptName & "_" & CStr(iTry)
        iTry = iTry + 1
    Loop
    MkDir runFolder
    Debug.Print "t+", Format(Timer - tStep, "0.000"), " Run folder created: " & runFolder
    tStep = Timer
    
    Dim art As Variant
    For Each art In items
        If art.Exists("abs") Then
            If fso.fileExists(CStr(art("abs"))) Then
                ArchiveFile CStr(art("abs")), runFolder
            End If
        End If
    Next art
    Debug.Print "t+", Format(Timer - tStep, "0.000"), " Artifacts archived"
    tStep = Timer
    
    Dim runId As String, metaFile As String
    If meta.Exists("run_id") Then
        runId = CStr(meta("run_id"))
        metaFile = folderTemp & "\meta_" & scriptName & "_" & runId & ".xml"
        If fso.fileExists(metaFile) Then
            ArchiveFile metaFile, runFolder
            Debug.Print "t+", Format(Timer - tStep, "0.000"), " Meta file archived"
        Else
            Debug.Print "t+", Format(Timer - tStep, "0.000"), " Meta file not found: " & metaFile
        End If
    Else
        Debug.Print "t+", Format(Timer - tStep, "0.000"), " Meta missing run_id. Cannot archive meta file."
    End If
    
    If fso.fileExists(inFile) Then
        ArchiveFile inFile, runFolder
        Debug.Print "t+", Format(Timer - tStep, "0.000"), " Input file archived"
    End If
    
    TrimArchive folderArchive, 10
    Debug.Print "t+", Format(Timer - tStep, "0.000"), " Archive trimmed"
    tStep = Timer

    Debug.Print "t+", Format(Timer - tStart, "0.000"), " Cleanup complete."
    GoTo CleanExit

CleanFail:
    Debug.Print "t+", Format(Timer - tStart, "0.000"), " Cleanup error: "; Err.Number; Err.Description
    Err.Clear

CleanExit:
    On Error GoTo 0
    Py = True
    Debug.Print "t+", Format(Timer - tStart, "0.000"), " === End Py (Success) ==="
    Exit Function

Fail:
    Debug.Print "t+", Format(Timer - tStart, "0.000"), " Error in Py(): " & Err.Description
    Py = False
End Function




Public Function RunPythonJob(script As String, tempFiles As Object, Optional inputText As String = "") As Object
    On Error GoTo Fail

    Dim tStart As Double
    Dim tStep As Double
    tStart = Timer
    tStep = tStart

    Debug.Print Format(Now, "hh:nn:ss"), "=== Begin RunPythonJob (XML meta) ==="

    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim sh As Object: Set sh = CreateObject("WScript.Shell")

    Debug.Print "t+", Format(Timer - tStep, "0.000"), " Resolving project root"
    Dim wbPath As String: wbPath = ResolveProjectPath()
    If wbPath = "" Then
        MsgBox "Could not resolve workbook path.", vbCritical, "RunPythonJob Error"
        Exit Function
    End If
    tStep = Timer

    Debug.Print "t+", Format(Timer - tStep, "0.000"), " Building folder paths"
    Dim folderUserScripts As String: folderUserScripts = wbPath & "\UserScripts"
    Dim folderPythonScripts As String: folderPythonScripts = wbPath & "\Python\scripts"
    Dim exe As String: exe = wbPath & "\Python\.venv\Scripts\python.exe"
    Dim scriptPath As String: scriptPath = folderUserScripts & "\" & script
    Dim tempFolder As String: tempFolder = wbPath & "\Temp"

    If Not fso.fileExists(exe) Then
        MsgBox "Python executable not found at: " & exe, vbCritical, "RunPythonJob Error"
        Exit Function
    End If
    If Not fso.fileExists(scriptPath) Then
        MsgBox "Python script not found: " & scriptPath, vbCritical, "RunPythonJob Error"
        Exit Function
    End If
    If Not fso.FolderExists(folderPythonScripts) Then
        MsgBox "Required module folder not found: " & folderPythonScripts, vbCritical, "RunPythonJob Error"
        Exit Function
    End If
    If Not fso.FolderExists(tempFolder) Then
        On Error Resume Next
        fso.CreateFolder tempFolder
        On Error GoTo Fail
        If Not fso.FolderExists(tempFolder) Then
            MsgBox "Temp folder missing and could not be created: " & tempFolder, vbCritical, "RunPythonJob Error"
            Exit Function
        End If
    End If
    tStep = Timer
    Debug.Print "t+", Format(Timer - tStep, "0.000"), " Environment verified"

    Debug.Print "t+", Format(Timer - tStep, "0.000"), " Preparing meta/log files"
    Randomize
    Dim runId As String
    runId = Format(Now, "yyyymmdd_hhnnss") & "_" & Hex(Int(Rnd * 100000))

    Dim metaFile As String
    metaFile = tempFolder & "\meta_" & script & "_" & runId & ".xml"
    Debug.Print "Meta file path: " & metaFile

    Dim logOut As String, logErr As String
    logOut = tempFolder & "\py_" & script & "_" & runId & ".out.log"
    logErr = tempFolder & "\py_" & script & "_" & runId & ".err.log"
    Debug.Print "Stdout log: " & logOut
    Debug.Print "Stderr log: " & logErr
    tStep = Timer

    If Len(inputText) > 0 Then
        Debug.Print "t+", Format(Timer - tStep, "0.000"), " Writing input text to file: " & tempFiles("in")
        Dim fNum As Integer: fNum = FreeFile
        Open tempFiles("in") For Output As #fNum
        Print #fNum, inputText
        Close #fNum
    End If
    tStep = Timer

    Debug.Print "t+", Format(Timer - tStep, "0.000"), " Building command line"
    Dim baseCmd As String
    baseCmd = _
        "set ""PYTHONPATH=" & folderPythonScripts & ";%PYTHONPATH%"" && " & _
        "set ""PYTHONUNBUFFERED=1"" && " & _
        """" & exe & """ -u " & _
        """" & scriptPath & """ " & _
        "--in " & """" & tempFiles("in") & """ " & _
        "--out " & """" & tempFiles("out") & """ " & _
        "--meta " & """" & metaFile & """ " & _
        "--run-id " & """" & runId & """" & _
        " 1> """ & logOut & """ 2> """ & logErr & """"

    Dim cmd As String
    cmd = "cmd /c " & """" & baseCmd & """"
    Debug.Print "Full command: " & cmd
    tStep = Timer

    Debug.Print "t+", Format(Timer - tStep, "0.000"), " Launching Python process"
    sh.Run cmd, 1, False
    Debug.Print "Python process launched."
    tStep = Timer

    Debug.Print "t+", Format(Timer - tStep, "0.000"), " Waiting for meta file creation"
    Dim meta As Object, status As String
    Dim tStartWait As Double: tStartWait = Timer
    Dim timeout As Double: timeout = 60
    Dim delay As Long: delay = 200
    Dim delta As Double

    Do
        If Dir(metaFile) <> "" Then
            Set meta = ReadMetaStatus(metaFile)
            If Not meta Is Nothing Then
                If Len(CStr(meta("status"))) > 0 Then
                    status = CStr(meta("status"))
                    Debug.Print "t+", Format(Timer - tStartWait, "0.000"), " Meta status first found: " & status
                    Exit Do
                End If
            End If
        End If

        delta = Timer - tStartWait
        If delta < 0 Then delta = delta + 86400
        If delta > timeout Then
            MsgBox "Meta file was not created with a valid 'status' field within " & timeout & " seconds.", _
                   vbCritical, "RunPythonJob Error"
            Set RunPythonJob = Nothing
            Exit Function
        End If

        DoEvents
        Sleep delay
    Loop
    tStep = Timer
    Debug.Print "t+", Format(Timer - tStep, "0.000"), " Meta file handshake complete"

    Debug.Print "t+", Format(Timer - tStep, "0.000"), " Polling until done/error"
    Dim reactWait As Double: reactWait = 120
    Dim maxWait As Double: maxWait = 300
    Dim lastTimestamp As String: lastTimestamp = ""
    Dim timeSinceLastUpdate As Double: timeSinceLastUpdate = 0
    Dim tStartTotal As Double: tStartTotal = Timer
    Dim tStartLastUpdate As Double: tStartLastUpdate = Timer

    Do
        Set meta = ReadMetaStatus(metaFile)
        If meta Is Nothing Then
            MsgBox "Failed to parse meta file (status).", vbCritical, "RunPythonJob Error"
            Set RunPythonJob = Nothing
            Exit Function
        End If

        If CStr(meta("run_id")) <> runId Or Len(CStr(meta("status"))) = 0 Then
            GoTo WaitNext
        End If

        status = CStr(meta("status"))
        If status = "error" Then
            Debug.Print "t+", Format(Timer - tStartTotal, "0.000"), " Python status=error"
            Dim metaFullErr As Object: Set metaFullErr = ParseMetaXml(metaFile)
            If Not metaFullErr Is Nothing Then
                metaFullErr("meta_path") = metaFile
                metaFullErr("stdout_log") = logOut
                metaFullErr("stderr_log") = logErr
                Set RunPythonJob = metaFullErr
            Else
                meta("meta_path") = metaFile
                meta("stdout_log") = logOut
                meta("stderr_log") = logErr
                Set RunPythonJob = meta
            End If
            Exit Function

        ElseIf status = "done" Then
            Debug.Print "t+", Format(Timer - tStartTotal, "0.000"), " Python status=done"
            Dim metaFull As Object: Set metaFull = ParseMetaXml(metaFile)
            If Not metaFull Is Nothing Then
                metaFull("meta_path") = metaFile
                Debug.Print "Final meta: duration=" & CStr(metaFull("duration"))
                On Error Resume Next
                If metaFull.Exists("artifacts") Then Debug.Print "Artifacts count=" & CStr(metaFull("artifacts").count)
                On Error GoTo Fail
                Set RunPythonJob = metaFull
            Else
                meta("meta_path") = metaFile
                Set RunPythonJob = meta
            End If
            Exit Function

        ElseIf status = "in_progress" Then
            Dim currentTimestamp As String: currentTimestamp = CStr(meta("timestamp"))
            If currentTimestamp = lastTimestamp Then
                timeSinceLastUpdate = Timer - tStartLastUpdate
                If timeSinceLastUpdate > reactWait Then
                    Debug.Print "t+", Format(Timer - tStartTotal, "0.000"), " Stall detected (>reactWait)"
                    MsgBox "Python script appears stalled: no update in " & reactWait & " seconds.", _
                           vbCritical, "RunPythonJob Error"
                    Set RunPythonJob = Nothing
                    Exit Function
                End If
            Else
                lastTimestamp = currentTimestamp
                timeSinceLastUpdate = 0
                tStartLastUpdate = Timer
                Debug.Print "t+", Format(Timer - tStartTotal, "0.000"), " Meta updated, status=in_progress"
            End If
        End If

WaitNext:
        DoEvents
        Sleep delay

        If Timer - tStartTotal > maxWait Then
            Debug.Print "t+", Format(Timer - tStartTotal, "0.000"), " Max wait time reached"
            MsgBox "Max wait time reached without completion.", vbCritical, "RunPythonJob Error"
            Set RunPythonJob = Nothing
            Exit Function
        End If
    Loop

Fail:
    Dim failMsg As String
    If Err.Number <> 0 Then
        failMsg = "Unhandled error in RunPythonJob: " & Err.Description
    Else
        failMsg = "RunPythonJob exited unexpectedly with no error object."
    End If
    Debug.Print "t+", Format(Timer - tStart, "0.000"), " " & failMsg
    MsgBox failMsg, vbCritical, "RunPythonJob Error"
    Set RunPythonJob = Nothing
End Function




