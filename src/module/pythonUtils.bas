Attribute VB_Name = "pythonUtils"
Option Explicit

Public Sub ArchiveFile(ByVal srcPath As String, ByVal destFolder As String)
    On Error GoTo ExitSub

    If Len(Dir$(srcPath, vbNormal)) = 0 Then Exit Sub

    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim base As String
    base = Mid$(srcPath, InStrRev(srcPath, "\") + 1)

    fso.MoveFile srcPath, destFolder & "\" & base

ExitSub:
End Sub

Public Sub TrimArchive(ByVal folderArchive As String, ByVal maxFiles As Long)
    On Error GoTo ExitSub
    If Len(Dir$(folderArchive, vbDirectory)) = 0 Then Exit Sub

    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim root As Object: Set root = fso.GetFolder(folderArchive)

    Dim n As Long: n = root.subFolders.count
    If n <= maxFiles Or n <= 1 Then Exit Sub

    Dim paths() As String, dates() As Date
    ReDim paths(1 To n)
    ReDim dates(1 To n)

    Dim idx As Long: idx = 0
    Dim sf As Object
    For Each sf In root.subFolders
        idx = idx + 1
        paths(idx) = sf.path
        dates(idx) = sf.DateCreated ' stable creation time ordering
    Next sf

    QuickSortByDate paths, dates, 1, n ' DESC (newest first)

    Dim i As Long
    For i = maxFiles + 1 To n
        On Error Resume Next
        fso.DeleteFolder paths(i), True
        On Error GoTo 0
    Next i

ExitSub:
End Sub

Public Sub QuickSortByDate(ByRef paths() As String, ByRef dates() As Date, ByVal lo As Long, ByVal hi As Long)
    Dim i As Long, j As Long
    Dim pivot As Date
    Dim tp As String, td As Date

    i = lo: j = hi
    pivot = dates((lo + hi) \ 2)

    Do While i <= j
        Do While dates(i) > pivot: i = i + 1: Loop   ' newer first
        Do While dates(j) < pivot: j = j - 1: Loop
        If i <= j Then
            td = dates(i): dates(i) = dates(j): dates(j) = td
            tp = paths(i): paths(i) = paths(j): paths(j) = tp
            i = i + 1: j = j - 1
        End If
    Loop

    If lo < j Then QuickSortByDate paths, dates, lo, j
    If i < hi Then QuickSortByDate paths, dates, i, hi
End Sub

Public Function EscapeXml(s As String) As String
    s = Replace(s, "&", "&amp;")
    s = Replace(s, "<", "&lt;")
    s = Replace(s, ">", "&gt;")
    s = Replace(s, """", "&quot;")
    s = Replace(s, "'", "&apos;")
    EscapeXml = s
End Function

Public Function ReadTextFromFile(path As String) As String
    Dim stm As Object
    On Error GoTo EH
    
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2                 ' adTypeText
    stm.Charset = "utf-8"        ' force UTF-8 decode (handles BOM or no BOM)
    stm.Open
    stm.LoadFromFile path
    ReadTextFromFile = stm.ReadText(-1)   ' adReadAll

CLEANUP:
    On Error Resume Next
    If Not stm Is Nothing Then stm.Close
    Set stm = Nothing
    Exit Function

EH:
    ReadTextFromFile = vbNullString
    Resume CLEANUP
End Function




Public Function ParseMetaXml(xmlPath As String) As Object
    On Error GoTo Fail

    Debug.Print "ParseMetaXml: path=" & xmlPath

    Dim dom As Object
    Set dom = CreateObject("MSXML2.DOMDocument.6.0")
    dom.async = False
    dom.validateOnParse = False
    dom.resolveExternals = False
    dom.SetProperty "SelectionLanguage", "XPath"
    If dom.Load(xmlPath) = False Then
        Debug.Print "ParseMetaXml: load error " & dom.ParseError.reason
        GoTo Fail
    End If

    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare

    Dim n As Object
    Set n = dom.SelectSingleNode("/meta/run_id"):    d("run_id") = IIf(n Is Nothing, "", CStr(n.text)): Debug.Print "run_id=" & d("run_id")
    Set n = dom.SelectSingleNode("/meta/status"):    d("status") = IIf(n Is Nothing, "", CStr(n.text)): Debug.Print "status=" & d("status")
    Set n = dom.SelectSingleNode("/meta/timestamp"): d("timestamp") = IIf(n Is Nothing, "", CStr(n.text)): Debug.Print "timestamp=" & d("timestamp")
    Set n = dom.SelectSingleNode("/meta/message"):   d("message") = IIf(n Is Nothing, "", CStr(n.text)): Debug.Print "message=" & d("message")
    Set n = dom.SelectSingleNode("/meta/duration"):  d("duration") = IIf(n Is Nothing, "", CStr(n.text)): Debug.Print "duration=" & d("duration")
    Set n = dom.SelectSingleNode("/meta/stderr"):    d("stderr") = IIf(n Is Nothing, "", CStr(n.text)): Debug.Print "stderr=" & d("stderr")

    Dim arts As New Collection
    Dim nodes As Object
    Dim aNode As Object
    Dim at As Object
    Dim a As Object
    Set nodes = dom.SelectNodes("/meta/artifacts/artifact")
    Debug.Print "ParseMetaXml: artifacts count=" & nodes.Length

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    For Each aNode In nodes
        Set a = CreateObject("Scripting.Dictionary")
        a.CompareMode = vbTextCompare

        For Each at In aNode.Attributes
            a(at.nodeName) = CStr(at.text)
        Next at

        Dim href As String
        href = IIf(a.Exists("href"), a("href"), "")
        Dim absPath As String
        absPath = ResolveHref(xmlPath, href)
        a("abs") = absPath
        a("exists") = fso.fileExists(absPath)

        Debug.Print "artifact: type=" & IIf(a.Exists("type"), a("type"), "") & _
                    " id=" & IIf(a.Exists("id"), a("id"), "") & _
                    " href=" & href & " abs=" & absPath & " exists=" & a("exists")

        arts.Add a
    Next aNode

    If d.Exists("artifacts") Then d.Remove "artifacts"
    d.Add "artifacts", arts

    Debug.Print "ParseMetaXml: artifacts collected=" & arts.count
    Set ParseMetaXml = d
    Exit Function

Fail:
    Debug.Print "ParseMetaXml: ERROR " & Err.Number & " - " & Err.Description
    Set ParseMetaXml = Nothing
End Function






Private Function SafeNodeText(dom As Object, xp As String) As String
    On Error GoTo done
    Dim n As Object: Set n = dom.SelectSingleNode(xp)
    If n Is Nothing Then
        SafeNodeText = ""
    Else
        SafeNodeText = CStr(n.text)
    End If
    Exit Function
done:
    SafeNodeText = ""
End Function



Public Function Nz(v As Variant, Optional fallback As String = "") As String
    If IsObject(v) Then
        Nz = fallback
    ElseIf IsNull(v) Or IsEmpty(v) Then
        Nz = fallback
    Else
        Nz = CStr(v)
    End If
End Function

'Public Function Nz(v As Variant, Optional fallback As String = "") As String
'    If IsNull(v) Or IsEmpty(v) Then
'        Nz = fallback
'    Else
'        Nz = CStr(v)
'    End If
'End Function





Public Function ReadMetaStatus(xmlPath As String) As Object
    On Error GoTo Fail
    Debug.Print "ReadMetaStatus: path=" & xmlPath

    Dim dom As Object
    Set dom = CreateObject("MSXML2.DOMDocument.6.0")
    dom.async = False: dom.validateOnParse = False: dom.resolveExternals = False
    dom.SetProperty "SelectionLanguage", "XPath"

    If dom.Load(xmlPath) = False Then
        Debug.Print "ReadMetaStatus: parse error -> " & dom.ParseError.reason & _
                    " line=" & dom.ParseError.line & _
                    " pos=" & dom.ParseError.linepos
        Set ReadMetaStatus = Nothing
        Exit Function
    End If

    If dom.DocumentElement Is Nothing Then
        Debug.Print "ReadMetaStatus: documentElement is Nothing"
    Else
        Debug.Print "ReadMetaStatus: root=" & dom.DocumentElement.nodeName & _
                    " version=" & dom.DocumentElement.GetAttribute("version")
    End If

    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare

    Dim vRun As String:   vRun = GetNode(dom, "/meta/run_id")
    Dim vStatus As String: vStatus = GetNode(dom, "/meta/status")
    Dim vTs As String:    vTs = GetNode(dom, "/meta/timestamp")
    Dim vMsg As String:   vMsg = GetNode(dom, "/meta/message")

    Debug.Print "ReadMetaStatus: run_id=" & "[" & vRun & "]"
    Debug.Print "ReadMetaStatus: status=" & "[" & vStatus & "]"
    Debug.Print "ReadMetaStatus: timestamp=" & "[" & vTs & "]"
    Debug.Print "ReadMetaStatus: message=" & "[" & vMsg & "]"

    d.Add "run_id", vRun
    d.Add "status", vStatus
    d.Add "timestamp", vTs
    d.Add "message", vMsg
    Set ReadMetaStatus = d
    Exit Function

Fail:
    Debug.Print "ReadMetaStatus: VBA error " & Err.Number & " - " & Err.Description
    Set ReadMetaStatus = Nothing
End Function


Private Function GetNode(dom As Object, xp As String) As String
    Dim n As Object
    Set n = dom.SelectSingleNode(xp)
    If Not n Is Nothing Then
        GetNode = CStr(n.text)
    Else
        GetNode = ""
    End If
End Function







Public Function GetXmlText(dom As Object, xpath As String) As String
    On Error GoTo done
    Dim n As Object
    Set n = dom.SelectSingleNode(xpath) ' requires SelectionLanguage="XPath"
    If Not n Is Nothing Then GetXmlText = CStr(n.text) Else GetXmlText = ""
done:
End Function


Public Function GetAttr(node As Object, name As String, Optional defaultValue As String = "") As String
    On Error GoTo miss
    Dim a As Object
    Set a = node.Attributes.getNamedItem(name)
    If Not a Is Nothing Then
        GetAttr = CStr(a.text)
    Else
        GetAttr = defaultValue
    End If
    Exit Function
miss:
    GetAttr = defaultValue
End Function

Public Function NormalizeHref(href As String) As String
    If href = "" Then
        NormalizeHref = ""
    Else
        NormalizeHref = Replace(href, "/", "\")
    End If
End Function

Public Function ResolveHref(metaXmlPath As String, href As String) As String
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If href = "" Then ResolveHref = "": Exit Function
    If InStr(href, ":") > 0 Or Left$(href, 2) = "\\" Then
        ResolveHref = href
        Exit Function
    End If
    Dim baseDir As String
    baseDir = fso.GetParentFolderName(metaXmlPath)
    If right$(baseDir, 1) = "\" Then
        ResolveHref = baseDir & NormalizeHref(href)
    Else
        ResolveHref = baseDir & "\" & NormalizeHref(href)
    End If
End Function






' ===== Misc used by Py() =====
'Public Function ResolveTargetRange(defSheet As Worksheet, targetRef As String) As Range
'    On Error GoTo fail
'
'    Dim acc As Range, part As Variant
'    Dim parts() As String, ref As String
'    Dim ws As Worksheet, wsResult As Worksheet
'    Dim bangPos As Long, sheetName As String, addr As String
'
'    ref = Trim$(targetRef)
'    If Len(ref) = 0 Then GoTo fail
'
'    ' Split on semicolons (your chosen separator). Commas remain valid inside a part.
'    parts = Split(ref, ";")
'
'    For Each part In parts
'        ref = Trim$(CStr(part))
'        If Len(ref) = 0 Then GoTo fail
'
'        bangPos = InStr(1, ref, "!", vbTextCompare)
'        If bangPos > 0 Then
'            sheetName = Left$(ref, bangPos - 1)
'            addr = Mid$(ref, bangPos + 1)
'
'            ' Strip optional surrounding single quotes in sheet names: 'My Sheet'
'            If Left$(sheetName, 1) = "'" And right$(sheetName, 1) = "'" Then
'                sheetName = Mid$(sheetName, 2, Len(sheetName) - 2)
'            End If
'
'            Set ws = Application.ThisWorkbook.Worksheets(sheetName)
'        Else
'            addr = ref
'            Set ws = defSheet
'        End If
'
'        ' Enforce single-worksheet result
'        If wsResult Is Nothing Then
'            Set wsResult = ws
'        ElseIf Not ws Is wsResult Then
'            GoTo fail
'        End If
'
'        If acc Is Nothing Then
'            Set acc = ws.Range(addr)                 ' accepts area lists like "A1,B2:B3" inside a part
'        Else
'            Set acc = Application.Union(acc, ws.Range(addr))
'        End If
'    Next part
'
'    Set ResolveTargetRange = acc
'    Exit Function
'
'fail:
'    Set ResolveTargetRange = Nothing
'End Function
' ===== Resolve parts into iterable areas across sheets =====






























'Public Function ResolveTargetRanges(defSheet As Worksheet, targetRef As String) As Collection
'    On Error GoTo Fail
'
'    Dim parts() As String, part As Variant, ref As String
'    Dim bangPos As Long, sheetName As String, addr As String
'    Dim ws As Worksheet, r As Range, area As Range
'    Dim result As New Collection
'
'    ref = Trim$(targetRef)
'    If Len(ref) = 0 Then GoTo Fail
'
'    ' Semicolons separate parts; commas inside a part remain valid area separators.
'    parts = Split(ref, ";")
'
'    For Each part In parts
'        ref = Trim$(CStr(part))
'        If Len(ref) = 0 Then GoTo Fail
'
'        bangPos = InStr(1, ref, "!", vbTextCompare)
'        If bangPos > 0 Then
'            sheetName = Left$(ref, bangPos - 1)
'            addr = Mid$(ref, bangPos + 1)
'
'            ' Strip optional surrounding single quotes in sheet names.
'            If Left$(sheetName, 1) = "'" And right$(sheetName, 1) = "'" Then
'                sheetName = Mid$(sheetName, 2, Len(sheetName) - 2)
'            End If
'
'            Set ws = Application.ThisWorkbook.Worksheets(sheetName)
'        Else
'            addr = ref
'            Set ws = defSheet
'        End If
'
'        ' This accepts intra-part area lists like "A1,B2:B3"
'        Set r = ws.Range(addr)
'
'        ' Normalize to individual areas so iteration is trivial.
'        For Each area In r.Areas
'            result.Add area
'        Next area
'    Next part
'
'    Set ResolveTargetRanges = result
'    Exit Function
'
'Fail:
'    Set ResolveTargetRanges = Nothing
'End Function






Public Function IsEmfFile(path As String, mime As String) As Boolean
    Dim ext As String
    ext = LCase$(Mid$(path, InStrRev(path, ".")))
    If ext = ".emf" Then
        IsEmfFile = True
    ElseIf LCase$(mime) = "image/x-emf" Then
        IsEmfFile = True
    Else
        IsEmfFile = False
    End If
End Function








' ===== Consumers for independent artifact files =====
Public Function LoadValueXml(path As String) As String
    On Error GoTo Fail
    Dim dom As Object
    Set dom = CreateObject("MSXML2.DOMDocument.6.0")
    dom.async = False: dom.validateOnParse = False: dom.resolveExternals = False
    dom.SetProperty "SelectionLanguage", "XPath"
    If dom.Load(path) = False Then
        Debug.Print "LoadValueXml: parseError ->"; dom.ParseError.reason; " line="; dom.ParseError.line; " pos="; dom.ParseError.linepos
        GoTo Fail
    End If
    Dim n As Object
    Set n = dom.SelectSingleNode("/value")
    If n Is Nothing Then
        LoadValueXml = ""
    Else
        LoadValueXml = CStr(n.text)
    End If
    Exit Function
Fail:
    LoadValueXml = ""
End Function

Public Function LoadListXmlAsColumn(path As String) As Variant
    On Error GoTo Fail
    Dim dom As Object, nodes As Object
    Set dom = CreateObject("MSXML2.DOMDocument.6.0")
    dom.async = False: dom.validateOnParse = False: dom.resolveExternals = False
    dom.SetProperty "SelectionLanguage", "XPath"
    If dom.Load(path) = False Then
        Debug.Print "LoadListXmlAsColumn: parseError ->"; dom.ParseError.reason; " line="; dom.ParseError.line; " pos="; dom.ParseError.linepos
        GoTo Fail
    End If
    Set nodes = dom.SelectNodes("/list/item")
    If nodes Is Nothing Or nodes.Length = 0 Then GoTo Fail

    Dim arr() As Variant
    ReDim arr(1 To nodes.Length, 1 To 1)
    Dim i As Long
    For i = 0 To nodes.Length - 1
        arr(i + 1, 1) = CStr(nodes.Item(i).text)
    Next i
    LoadListXmlAsColumn = arr
    Exit Function
Fail:
    LoadListXmlAsColumn = Empty
End Function







' Returns Scripting.Dictionary: key = artifact id (String), item = Range
' On parse/resolve errors, fills errorsOut (Collection) with diagnostic strings.
Public Function ParseIdToRangeMap( _
    ByVal dstSpec As String, _
    ByVal defaultSheet As Worksheet, _
    ByVal artifactIds As Variant, _
    Optional ByRef errorsOut As Collection _
) As Object
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare ' case-insensitive ids
'    Debug.Print "Initialized dictionary."

    If errorsOut Is Nothing Then Set errorsOut = New Collection

    ' normalize artifacts into ordered list
    Dim artList As New Collection
    Dim v
    If IsArray(artifactIds) Then
'        Debug.Print "artifactIds is Array"
        For Each v In artifactIds
'            Debug.Print "Array item: [" & CStr(v) & "]"
            If LenB(CStr(v)) > 0 Then artList.Add CStr(v)
        Next
    ElseIf TypeName(artifactIds) = "Collection" Then
'        Debug.Print "artifactIds is Collection"
        For Each v In artifactIds
            If IsObject(v) Then
                If TypeName(v) = "Dictionary" Or TypeName(v) = "Scripting.Dictionary" Then
                    If v.Exists("id") Then
'                        Debug.Print "Artifact dict id: [" & v("id") & "]"
                        If LenB(CStr(v("id"))) > 0 Then artList.Add CStr(v("id"))
                    End If
                Else
                    Debug.Print "Collection object item (unsupported): " & TypeName(v)
                End If
            Else
                Debug.Print "Collection item: [" & CStr(v) & "]"
                If LenB(CStr(v)) > 0 Then artList.Add CStr(v)
            End If
        Next
    ElseIf TypeName(artifactIds) = "Dictionary" Or TypeName(artifactIds) = "Scripting.Dictionary" Then
        Debug.Print "artifactIds is Dictionary"
        Dim k: For Each k In artifactIds.keys
            Debug.Print "Dict key: [" & CStr(k) & "]"
            artList.Add CStr(k)
        Next
    ElseIf LenB(CStr(artifactIds)) > 0 Then
        Debug.Print "artifactIds is single value: [" & CStr(artifactIds) & "]"
        artList.Add CStr(artifactIds)
    End If
    Debug.Print "Normalized artList count = " & artList.count

    ' --- parse dstSpec into explicit and orphan parts ---
    Dim explicitPairs As Object: Set explicitPairs = CreateObject("Scripting.Dictionary")
    explicitPairs.CompareMode = vbTextCompare
    Dim orphanAddrs As New Collection
    
    Dim spec As String: spec = Trim$(dstSpec)
    Debug.Print "dstSpec = [" & spec & "]"
    If Len(spec) > 0 Then
        Dim parts() As String: parts = Split(spec, ";")
        Dim i As Long
        For i = LBound(parts) To UBound(parts)
            Dim token As String: token = Trim$(parts(i))
'            Debug.Print "Processing token: [" & token & "]"
            If Len(token) = 0 Then GoTo NextPart

            Dim p As Long: p = InStr(1, token, "=", vbTextCompare)
            If p = 0 Then
'                Debug.Print "  Orphan addr: [" & token & "]"
                orphanAddrs.Add token
                GoTo NextPart
            End If

            Dim id As String, addr As String
            id = Trim$(Left$(token, p - 1))
            addr = Trim$(Mid$(token, p + 1))
            Debug.Print "  Explicit pair: id=[" & id & "], addr=[" & addr & "]"

            If Len(id) = 0 Then
                Debug.Print "  Error: empty id"
                errorsOut.Add "Empty id in token: " & token
                GoTo NextPart
            End If
            If Len(addr) = 0 Then
                Debug.Print "  Error: empty addr for id " & id
                errorsOut.Add "Empty address for id '" & id & "'."
                GoTo NextPart
            End If

            Dim rng As Range
            Set rng = ResolveAddressToRange(addr, defaultSheet, errorsOut)
            If rng Is Nothing Then
                Debug.Print "  Could not resolve addr for id " & id
                errorsOut.Add "Could not resolve address '" & addr & "' for id '" & id & "'."
                GoTo NextPart
            End If

            If explicitPairs.Exists(id) Then
                Debug.Print "  Duplicate explicit id, replacing: " & id
                explicitPairs.Remove id
            End If
            explicitPairs.Add id, rng
            Debug.Print "  Added explicit pair: " & id

NextPart:
        Next i
    End If
    Debug.Print "ExplicitPairs count = " & explicitPairs.count & ", orphanAddrs count = " & orphanAddrs.count

    ' --- stage 1: explicit matches ---
    Dim idKey As Variant
    For Each idKey In explicitPairs.keys
        Set dict(idKey) = explicitPairs(idKey)
        Debug.Print "Stage1 explicit -> dict: " & idKey
    Next idKey

    ' --- stage 2: assign orphan addresses sequentially ---
    Dim orphanIdx As Long: orphanIdx = 1
    For Each v In artList
        If Not dict.Exists(CStr(v)) Then
            If orphanIdx <= orphanAddrs.count Then
'                Debug.Print "Stage2 orphan assign: " & CStr(v) & " -> " & orphanAddrs(orphanIdx)
                Dim orng As Range
                Set orng = ResolveAddressToRange(CStr(orphanAddrs(orphanIdx)), defaultSheet, errorsOut)
                If Not orng Is Nothing Then
                    Set dict(CStr(v)) = orng
'                    Debug.Print "  Assigned orphan addr to " & CStr(v)
                Else
                    Debug.Print "  Failed resolving orphan addr for " & CStr(v)
                End If
                orphanIdx = orphanIdx + 1
            Else
                Debug.Print "  No orphan addr available for " & CStr(v)
            End If
        End If
    Next v

    ' --- stage 3: prompt user for any remaining unassigned artifacts ---
    For Each v In artList
        If Not dict.Exists(CStr(v)) Then
            Debug.Print "Stage3 prompting user for artifact: " & CStr(v)
            On Error Resume Next
            Dim picked As Range
            Set picked = Application.InputBox("Select destination for artifact '" & CStr(v) & "'", _
                                              "Pick Range", Type:=8)
            On Error GoTo 0
            If Not picked Is Nothing Then
                Set dict(CStr(v)) = picked
                Debug.Print "  User selected range for " & CStr(v)
            Else
                Debug.Print "  User canceled for " & CStr(v)
                errorsOut.Add "No destination selected for artifact '" & CStr(v) & "'."
            End If
        End If
    Next v

    ' --- stage 4: warn on extra ids not in artifacts ---
    For Each idKey In explicitPairs.keys
        Dim known As Boolean: known = False
        For Each v In artList
            If StrComp(CStr(v), CStr(idKey), vbTextCompare) = 0 Then
                known = True: Exit For
            End If
        Next v
        If Not known Then
            Debug.Print "Stage4 extra id not in artifacts: " & CStr(idKey)
            MsgBox "Destination specified for id '" & CStr(idKey) & _
                   "' that is not present in artifacts. It will be ignored.", _
                   vbInformation, "Extra destination"
        End If
    Next idKey

'    Debug.Print "Final dict count = " & dict.count
    Set ParseIdToRangeMap = dict
End Function











' Accepts:
'  - Sheet-qualified addresses: Sheet1!A1, 'Other Sheet'!B3
'  - Unqualified A1 references resolved against defaultSheet
'  - Workbook-level named ranges (e.g., MyNamedRange)
Private Function ResolveAddressToRange( _
    ByVal addr As String, _
    ByVal defaultSheet As Worksheet, _
    ByRef errorsOut As Collection _
) As Range
    On Error GoTo Fail

    Dim ws As Worksheet
    Dim a As String: a = addr

    Dim excl As Long: excl = InStrRev(a, "!")
    If excl > 0 Then
        ' Sheet-qualified
        Dim sheetPart As String, refPart As String
        sheetPart = Left$(a, excl - 1)
        refPart = Mid$(a, excl + 1)

        sheetPart = Trim$(sheetPart)
        refPart = Trim$(refPart)

        ' Strip surrounding quotes for sheet names like 'My Sheet'
        If Left$(sheetPart, 1) = "'" And right$(sheetPart, 1) = "'" Then
            sheetPart = Mid$(sheetPart, 2, Len(sheetPart) - 2)
        End If

'        Debug.Print "ResolveAddressToRange: sheetPart=[" & sheetPart & "] refPart=[" & refPart & "]"

        Set ws = SheetByName(defaultSheet.parent, sheetPart)
        If ws Is Nothing Then
            Debug.Print "ResolveAddressToRange: sheet not found: " & sheetPart
            errorsOut.Add "Sheet not found: " & sheetPart
            Set ResolveAddressToRange = Nothing
            Exit Function
        End If

        Set ResolveAddressToRange = ws.Range(refPart)
'        Debug.Print "ResolveAddressToRange: OK [" & addr & "] -> " & ws.name & "!" & ws.Range(refPart).Address(False, False)
        Exit Function

    Else
        ' No sheet qualifier. Try workbook named range first.
        Dim nm As name
        For Each nm In defaultSheet.parent.Names
            If StrComp(nm.name, a, vbTextCompare) = 0 Then
                Debug.Print "ResolveAddressToRange: matched workbook name [" & a & "] -> " & nm.RefersToRange.Address(False, False)
                Set ResolveAddressToRange = nm.RefersToRange
                Exit Function
            End If
        Next nm

        ' Fallback: interpret as A1 on defaultSheet
        Debug.Print "ResolveAddressToRange: fallback on sheet [" & defaultSheet.name & "] ref=[" & a & "]"
        Set ResolveAddressToRange = defaultSheet.Range(a)
        Debug.Print "ResolveAddressToRange: OK [" & addr & "] -> " & defaultSheet.name & "!" & defaultSheet.Range(a).Address(False, False)
        Exit Function
    End If

Fail:
    Debug.Print "ResolveAddressToRange FAIL for [" & addr & "]: " & Err.Description
    errorsOut.Add "Invalid address '" & addr & "': " & Err.Description
    Set ResolveAddressToRange = Nothing
End Function


Private Function SheetByName(ByVal wb As Workbook, ByVal name As String) As Worksheet
    On Error GoTo notfound
    Set SheetByName = wb.Worksheets(name)
    Exit Function
notfound:
    Set SheetByName = Nothing
End Function





Public Function PasteArtifactsToTargets(idMap As Object, artifacts As Object) As Boolean
    On Error GoTo Fail

    Dim didAny As Boolean: didAny = False
    Dim it As Object, id As String, t As String, fpath As String
    Dim dstRng As Range, outXml As String, arr As Variant, sval As String
    Dim shp As shape, mime As String

    Debug.Print "PasteArtifactsToTargets: artifacts count=" & artifacts.count & ", idMap count=" & idMap.count
    Debug.Print String(80, "-")

    For Each it In artifacts
        Debug.Print "Processing new artifact item..."
        
        ' Verify if artifact is a valid object
        If Not IsObject(it) Then
            Debug.Print "  Skipping non-object item: "; TypeName(it)
            GoTo nextItem
        End If

        ' Print all keys of the artifact
        Debug.Print "  Artifact keys: "; Join(it.keys, ", ")

        ' Validate required fields
        If Not it.Exists("id") Or Not it.Exists("type") Or Not it.Exists("abs") Then
            Debug.Print "  Skipping artifact (missing id/type/abs)."
            GoTo nextItem
        End If

        fpath = CStr(it("abs"))
        id = CStr(it("id"))

        ' Check if mapping exists
        If Not idMap.Exists(id) Then
            Debug.Print "  Skipping artifact '" & id & "' (no mapping)."
            GoTo nextItem
        End If

        Debug.Print "  Artifact raw type=" & it("type")
        t = LCase$(Trim(CStr(it("type"))))
        Debug.Print "  Artifact normalized type (t)=" & t

        ' Validate mapping type
        If TypeName(idMap(id)) <> "Range" Then
            Debug.Print "  ERROR: mapping for '" & id & "' is not a Range, got " & TypeName(idMap(id))
            GoTo nextItem
        End If

        Set dstRng = idMap(id)

        ' ---- Select Case for type ----
        Debug.Print "  Entering Select Case with t='" & t & "'"
        
        Select Case t
            Case "table", "range"
                Debug.Print "    [Case table/range] Reading table from " & fpath
                outXml = ReadTextFromFile(fpath)
                If Len(outXml) > 0 Then
                    If PasteTypedXMLToRange(outXml, dstRng.Address(External:=True)) Then
                        Debug.Print "    Pasted table '" & id & "' to " & dstRng.Worksheet.name & "!" & dstRng.Address(False, False)
                        didAny = True
                    Else
                        Debug.Print "    PasteTypedXMLToRange failed for '" & id & "'"
                    End If
                Else
                    Debug.Print "    No XML content read from file: " & fpath
                End If
        
            Case "list"
                Debug.Print "    [Case list] Reading list from " & fpath
                arr = LoadListXmlAsColumn(fpath)
                If IsEmpty(arr) Then
                    Debug.Print "    Loaded list is EMPTY."
                Else
                    Dim items As Long
                    Dim pasteHoriz As Boolean
                    Dim i As Long
                    Dim arrRow() As Variant
        
                    items = UBound(arr, 1)
                    Debug.Print "    List items count=" & items
        
                    ' Orientation detection
                    If dstRng.CountLarge > 1 Then
                        pasteHoriz = (dstRng.Columns.count >= dstRng.rows.count)
                        Debug.Print "    Auto orientation based on dstRng: pasteHoriz=" & pasteHoriz
                    Else
                        Dim dlg As frmOrientation
                        Set dlg = New frmOrientation
                        dlg.Show
                        Debug.Print "    User orientation choice: " & dlg.Choice
                        If dlg.Choice = "H" Then
                            pasteHoriz = True
                        ElseIf dlg.Choice = "V" Then
                            pasteHoriz = False
                        Else
                            Debug.Print "    Paste of list '" & id & "' cancelled (no choice)."
                            Set dlg = Nothing
                            Exit Function
                        End If
                        Unload dlg
                        Set dlg = Nothing
                    End If
        
                    ' Paste operation
                    If pasteHoriz Then
                        If items <= 65000 Then
                            dstRng.Resize(1, items).value = WorksheetFunction.Transpose(arr)
                        Else
                            ReDim arrRow(1 To 1, 1 To items)
                            For i = 1 To items
                                arrRow(1, i) = arr(i, 1)
                            Next i
                            dstRng.Resize(1, items).value = arrRow
                        End If
                    Else
                        dstRng.Resize(items, 1).value = arr
                    End If
        
                    Debug.Print "    Pasted list '" & id & "' (" & items & " items)"
                    didAny = True
                End If
        
            Case "value"
                Debug.Print "    [Case value] Reading value from " & fpath
                sval = LoadValueXml(fpath)
                Debug.Print "    Loaded value=" & sval
                dstRng.value = sval
                Debug.Print "    Pasted value '" & id & "'"
                didAny = True
        
            Case "chart", "plot"
                Debug.Print "    [Case chart/plot] Reading chart from " & fpath
                mime = ""
                If it.Exists("mime") Then mime = LCase$(Trim(CStr(it("mime"))))
                Debug.Print "    Chart mime=" & mime
                Select Case mime
                    Case "image/x-emf", "image/svg+xml"
                        Set shp = dstRng.Worksheet.Shapes.AddPicture( _
                            fileName:=fpath, _
                            LinkToFile:=msoFalse, _
                            SaveWithDocument:=msoTrue, _
                            Left:=dstRng.Left, _
                            Top:=dstRng.Top, _
                            Width:=-1, Height:=-1)
                        Debug.Print "    Inserted chart '" & id & "' at " & dstRng.Worksheet.name & "!" & dstRng.Address(False, False)
                        didAny = True
                    Case Else
                        Debug.Print "    Unsupported chart mime for '" & id & "': " & mime
                End Select
        
            Case "plot2.0"
'                Debug.Print "    [Case plot2.0] Reading plot2.0 from " & fpath
'                Dim xmlText As String
'                xmlText = ReadTextFromFile(fpath)
'                Debug.Print "    xmlText length=" & Len(xmlText)
'                If Len(xmlText) > 0 Then
'                    If BuildChartFromXml(xmlText, dstRng.Worksheet, dstRng.Left, dstRng.Top, 400, 300) Then
'                        Debug.Print "    Inserted plot2.0 chart '" & id & "' at " & dstRng.Worksheet.name & "!" & dstRng.Address(False, False)
'                        didAny = True
'                    Else
'                        Debug.Print "    BuildChartFromXml failed for '" & id & "'"
'                    End If
'                Else
'                    Debug.Print "    No XML content read for plot2.0."
'                End If
                Call BuildChartFromXML(fpath, dstRng.Worksheet, dstRng)
        
            Case Else
                Debug.Print "    [Case Else] Unknown artifact type='" & t & "' for id='" & id & "'"
        End Select

nextItem:
        Debug.Print String(80, "-")
    Next it

    PasteArtifactsToTargets = didAny
    Debug.Print "PasteArtifactsToTargets finished. didAny=" & didAny
    Exit Function

Fail:
    Debug.Print "Error in PasteArtifactsToTargets: " & Err.Description & " (id=" & id & ")"
    PasteArtifactsToTargets = False
End Function






