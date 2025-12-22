Attribute VB_Name = "chartBuilder"
'' ==========================================================
'' Module: modChartSpec
'' Purpose: Parse ChartSpec XML and build Excel charts (Microsoft 365)
'' Behavior:
''   - Uses ONLY names present in XML; if missing/empty ? "add name here"
''   - Applies to: chart title, both axis titles, every series name
''   - Supports multi-series for line/column/bar/area/scatter; histogram = single series
''   - Series style attributes honored when present:
''       color="#RRGGBB", thickness|thikness="2.25", shape="circle|square|diamond|triangle|x|plus|star|dot|none",
''       style="solid|dash|dot|dashdot|dashdotdot", markerSize="5"
''   - Histogram: type "hist" or "histogram"; values vector from <y> or fallback <x>
'' ==========================================================
'Option Explicit
'
'' -------- Public entry point --------------------------------------------------
'Public Function BuildChartFromXml(xmlText As String, _
'                                  ws As Worksheet, _
'                                  ByVal leftPos As Double, _
'                                  ByVal topPos As Double, _
'                                  ByVal width As Double, _
'                                  ByVal height As Double) As Boolean
'    On Error GoTo ErrHandler
'    BuildChartFromXml = False
'
'    ' --- Load XML ---
'    Dim doc As Object
'    Set doc = CreateObject("MSXML2.DOMDocument.6.0")
'    doc.Async = False
'    doc.ValidateOnParse = False
'    If Not doc.LoadXML(xmlText) Then
'        Debug.Print "XML parse error: " & doc.ParseError.reason
'        Exit Function
'    End If
'    doc.SetProperty "SelectionLanguage", "XPath"
'    doc.SetProperty "SelectionNamespaces", "xmlns:cs='urn:example:chartspec:1.0'"
'
'    ' --- Get intended chart title early (used as unique identifier) ---
'    Dim chartTitle As String
'    chartTitle = GetAttr(doc, "//cs:layout/cs:title/@text")
'    If Len(chartTitle) = 0 Then chartTitle = "add name here"
'
'    ' --- Check for duplicate chart titles ---
'    Dim conflictAction As VbMsgBoxResult
'    Dim existingChart As ChartObject
'    Set existingChart = FindChartByTitle(ws, chartTitle)
'
'    If Not existingChart Is Nothing Then
'        conflictAction = MsgBox( _
'            "A chart with the title '" & chartTitle & "' already exists." & vbCrLf & vbCrLf & _
'            "Yes = Delete the old chart and create the new one" & vbCrLf & _
'            "No = Skip creating the new chart" & vbCrLf & _
'            "Cancel = Keep both (new chart will be renamed and highlighted)", _
'            vbYesNoCancel + vbExclamation, "Duplicate Chart Title")
'
'        Select Case conflictAction
'            Case vbYes
'                existingChart.Delete
'            Case vbNo
'                Exit Function
'            Case vbCancel
'                chartTitle = chartTitle & " (Copy)"
'        End Select
'    End If
'
'    ' --- Determine chart type from XML ---
'    Dim chartTypeText As String
'    chartTypeText = GetAttr(doc, "//cs:chart/@type")
'    Dim tLower As String: tLower = LCase$(chartTypeText)
'
'    Dim wantHistogram As Boolean: wantHistogram = IsHistogramType(tLower)
'    Dim wantScatter As Boolean: wantScatter = (tLower = "scatter")
'
'    ' --- Series nodes ---
'    Dim sNodes As Object
'    Set sNodes = doc.SelectNodes("//cs:series/cs:s")
'    If wantHistogram And sNodes.Length > 1 Then
'        Err.Raise vbObjectError + 7010, "BuildChartFromXml", _
'                  "Histogram charts support a single series. XML contains " & sNodes.Length & " series."
'    End If
'
'    ' --- Decide base chart type dynamically ---
'    Dim baseType As XlChartType
'    If wantHistogram Then
'        baseType = xlColumnClustered
'    ElseIf tLower = "mixed" Then
'        baseType = DetermineBaseType(sNodes)
'    Else
'        baseType = MapChartType(chartTypeText)
'    End If
'
'    ' --- Create chart object ---
'    Dim co As ChartObject
'    Set co = ws.ChartObjects.Add(leftPos, topPos, width, height)
'    co.Chart.chartType = baseType
'
'    ' --- Configure axes (now supports log/min/max/ticks/format/secondary X) ---
'    ApplyAxes co.Chart, doc, chartTypeText
'
'    ' --- Insert series ---
'    Dim s As Object, ser As Series
'    Dim addedAny As Boolean: addedAny = False
'
'    For Each s In sNodes
'        Set ser = co.Chart.SeriesCollection.NewSeries
'
'        Dim nm As String
'        nm = GetAttr(s, "@name")
'        If Len(nm) = 0 Then nm = "add name here"
'        ser.name = nm
'
'        Dim xVals As Variant, yVals As Variant
'        xVals = ReadArray(s, "cs:x/cs:n", IsCategoryDate(doc))
'        yVals = ReadArray(s, "cs:y/cs:n", False)
'
'        If wantHistogram Then
'            Dim vec As Variant
'            vec = GetNumericVector(yVals, xVals)
'            If IsEmpty(vec) Then
'                Err.Raise vbObjectError + 7002, "BuildChartFromXml", _
'                          "Histogram requires a numeric vector in <y> or <x>."
'            End If
'            ser.Values = vec
'        Else
'            If Not IsEmpty(xVals) Then ser.XValues = xVals
'            If Not IsEmpty(yVals) Then ser.Values = yVals
'        End If
'
'        ' --- Determine per-series chart type (now considers smooth markers for scatter) ---
'        Dim kind As String
'        kind = LCase$(GetAttr(s, "cs:style/@kind"))
'        Dim smoothReq As Boolean
'        smoothReq = (LCase$(GetAttr(s, "cs:style/@smooth")) = "true")
'
'        Select Case kind
'            Case "bar", "column"
'                ser.chartType = xlColumnClustered
'            Case "hist", "histogram"
'                ser.chartType = xlColumnClustered
'            Case "line"
'                ser.chartType = xlLine
'            Case "scatter"
'                If smoothReq Then
'                    ' If markers explicitly none, use no-markers variant; else smooth with markers
'                    If LCase$(GetAttr(s, "cs:style/@shape")) = "none" Then
'                        ser.chartType = xlXYScatterSmoothNoMarkers
'                    Else
'                        ser.chartType = xlXYScatterSmooth
'                    End If
'                Else
'                    ser.chartType = xlXYScatter
'                End If
'            Case "area"
'                ser.chartType = xlArea
'            Case Else
'                ser.chartType = baseType
'        End Select
'
'        ' --- Assign to secondary axis if requested (NEW: also checks for secondary X) ---
'        Dim yAxisAssign As String: yAxisAssign = LCase$(GetAttr(s, "@yAxis"))
'        Dim xAxisAssign As String: xAxisAssign = LCase$(GetAttr(s, "@xAxis"))
'        If yAxisAssign = "secondary" Or xAxisAssign = "secondary" Then
'            ser.axisGroup = xlSecondary
'        End If
'
'        ' --- Apply per-series visual styles (extended for border/opacity/smooth) ---
'        ApplySeriesStyle ser, s, chartTypeText
'
'        ' --- New: advanced per-series features (error bars, labels, trendlines) ---
'        ApplySeriesExtras ser, s, co.Chart, ws.Parent
'
'        ' --- NEW: Apply per-point styling for markers/borders ---
'        ApplySeriesPointStyles ser, s
'
'        addedAny = True
'    Next
'
'    ' --- Flip to histogram chart AFTER data series exist ---
'    If wantHistogram Then
'        If Not addedAny Then
'            Err.Raise vbObjectError + 7003, "BuildChartFromXml", "No series found for histogram."
'        End If
'        co.Chart.chartType = xlHistogram
'        ApplyHistogramOptions co.Chart, doc
'    End If
'
'    ' --- Layout: titles / legend / connect-gaps / drop-lines (extended) ---
'    ApplyLayout co.Chart, doc
'
'    ' Force the title set earlier (in case ApplyLayout defaulted it)
'    co.Chart.HasTitle = True
'    co.Chart.chartTitle.text = chartTitle
'
'    If conflictAction = vbCancel Then
'        co.Chart.ChartArea.Format.Fill.ForeColor.RGB = RGB(255, 230, 204)
'    End If
'
'    BuildChartFromXml = True
'    Exit Function
'
'ErrHandler:
'    Debug.Print "BuildChartFromXml error: " & Err.Description
'End Function
'
'
'
'' -------- XPath utils (doc already has SelectionNamespaces set) ---------------
'Private Function GetAttr(node As Object, xpath As String) As String
'    Dim x As Object: Set x = node.SelectSingleNode(xpath)
'    If Not x Is Nothing Then GetAttr = x.text Else GetAttr = ""
'End Function
'
'Private Function SelectNode(doc As Object, xpath As String) As Object
'    Dim n As Object: Set n = doc.SelectSingleNode(xpath)
'    Set SelectNode = n
'End Function
'
'' -------- Chart configuration helpers ----------------------------------------
'Private Function MapChartType(chartType As String) As XlChartType
'    Select Case LCase$(chartType)
'        Case "line":        MapChartType = xlLine
'        Case "column":      MapChartType = xlColumnClustered
'        Case "bar":         MapChartType = xlBarClustered
'        Case "area":        MapChartType = xlArea
'        Case "scatter":     MapChartType = xlXYScatter                 ' markers only
'        Case "hist", "histogram": MapChartType = xlHistogram            ' set later for safety
'        Case "pie":         MapChartType = xlPie
'        Case Else:          MapChartType = xlLine
'    End Select
'End Function
'
'Private Function IsHistogramType(chartTypeLower As String) As Boolean
'    IsHistogramType = (chartTypeLower = "hist" Or chartTypeLower = "histogram")
'End Function
'
'Private Sub ApplyAxes(ch As Chart, doc As Object, chartType As String)
'    Dim t As String: t = LCase$(chartType)
'    ' Keep existing early exit for histogram; extend scatter handling
'    If t = "hist" Or t = "histogram" Then Exit Sub
'
'    ' Determine category kind (date vs category) for non-XY
'    Dim catKind As String
'    catKind = GetAttr(doc, "//cs:chart/cs:categoryAxis/@kind")
'    If LCase$(catKind) = "date" And t <> "scatter" Then
'        ch.Axes(xlCategory).CategoryType = xlTimeScale
'    ElseIf t <> "scatter" Then
'        ch.Axes(xlCategory).CategoryType = xlCategoryScale
'    End If
'
'    ' --- Common axis options helper ---
'    ApplyAxisNode ch, doc, "//cs:chart/cs:categoryAxis", xlCategory, xlPrimary
'    ApplyAxisNode ch, doc, "//cs:chart/cs:valueAxis", xlValue, xlPrimary
'
'    ' Secondary Y axis node (valueAxis2)
'    ApplyAxisNode ch, doc, "//cs:chart/cs:valueAxis2", xlValue, xlSecondary
'
'    ' Secondary X axis node (categoryAxis2) ? useful for XY when series on secondary
'    ApplyAxisNode ch, doc, "//cs:chart/cs:categoryAxis2", xlCategory, xlSecondary
'End Sub
'
'' Apply one axis node?s settings (min/max/log/ticks/format/cross/scale)
'Private Sub ApplyAxisNode(ch As Chart, doc As Object, xp As String, axisType As XlAxisType, axisGroup As XlAxisGroup)
'    On Error Resume Next
'    Dim n As Object: Set n = SelectNode(doc, xp)
'    If n Is Nothing Then Exit Sub
'
'    If Not AxisExists(ch, axisType, axisGroup) Then
'        ' Force creation for secondary axes when needed
'        Dim tmp As Axis
'        Set tmp = ch.Axes(axisType, axisGroup)
'    End If
'
'    If Not AxisExists(ch, axisType, axisGroup) Then Exit Sub
'
'    With ch.Axes(axisType, axisGroup)
'        Dim logTxt As String: logTxt = LCase$(GetAttr(n, "@log"))
'        Dim minTxt As String: minTxt = GetAttr(n, "@min")
'        Dim maxTxt As String: maxTxt = GetAttr(n, "@max")
'        Dim majorTxt As String: majorTxt = GetAttr(n, "@majorUnit")
'        Dim minorTxt As String: minorTxt = GetAttr(n, "@minorUnit")
'        Dim fmt As String: fmt = GetAttr(n, "@numberFormat")
'        Dim crossAtTxt As String: crossAtTxt = GetAttr(n, "@crossAt")
'        Dim reverseTxt As String: reverseTxt = LCase$(GetAttr(n, "@reverse"))
'        Dim tickLblSpc As String: tickLblSpc = GetAttr(n, "@tickLabelSpacing")
'
'        If logTxt = "true" Then .ScaleType = xlLogarithmic
'        If IsNumeric(minTxt) Then .MinimumScale = CDbl(minTxt)
'        If IsNumeric(maxTxt) Then .MaximumScale = CDbl(maxTxt)
'        If IsNumeric(majorTxt) Then .MajorUnit = CDbl(majorTxt)
'        If IsNumeric(minorTxt) Then .MinorUnit = CDbl(minorTxt)
'        If Len(fmt) > 0 Then .TickLabels.numberFormat = fmt
'        If IsNumeric(crossAtTxt) Then .CrossesAt = CDbl(crossAtTxt)
'        If reverseTxt = "true" Then .ReversePlotOrder = True
'        If IsNumeric(tickLblSpc) Then .TickLabelSpacing = CLng(tickLblSpc)
'
'        ' Axis title if provided
'        Dim ttl As String: ttl = GetAttr(n, "@title")
'        If Len(ttl) > 0 Then
'            .HasTitle = True
'            .AxisTitle.Characters.text = ttl
'        End If
'    End With
'    On Error GoTo 0
'End Sub
'
'' -------- Series helpers ------------------------------------------------------
'Private Function ReadArray(node As Object, xpath As String, isDate As Boolean) As Variant
'    Dim list As Object: Set list = node.SelectNodes(xpath)
'    Dim n As Long: n = list.Length
'    If n = 0 Then Exit Function
'
'    Dim arr() As Variant: ReDim arr(1 To n)
'    Dim i As Long, s As String, dt As Variant
'    For i = 1 To n
'        s = Trim$(list.Item(i - 1).text)
'        If isDate Then
'            dt = ParseIso8601Date(s)
'            If IsEmpty(dt) Then
'                arr(i) = s
'            Else
'                arr(i) = dt
'            End If
'        ElseIf IsNumeric(s) Then
'            arr(i) = CDbl(s)
'        Else
'            arr(i) = s
'        End If
'    Next
'    ReadArray = arr
'End Function
'
'Private Sub ApplySeriesStyle(ser As Series, sNode As Object, chartType As String)
'    Dim colorHex As String
'    Dim thicknessTxt As String
'    Dim shapeTxt As String
'    Dim styleTxt As String
'    Dim markerSizeTxt As String
'
'    ' New optional attributes
'    Dim strokeHex As String
'    Dim strokeWidthTxt As String
'    Dim markerOpacityTxt As String
'    Dim smoothTxt As String
'
'    ' NEW: invert negative for area fills
'    Dim invertNegTxt As String
'    Dim invertColorHex As String
'
'    colorHex = GetAttr(sNode, "cs:style/@color")
'    thicknessTxt = GetAttr(sNode, "cs:style/@thickness")
'    If Len(thicknessTxt) = 0 Then thicknessTxt = GetAttr(sNode, "cs:style/@thikness")
'    shapeTxt = LCase$(GetAttr(sNode, "cs:style/@shape"))
'    styleTxt = LCase$(GetAttr(sNode, "cs:style/@style"))
'    markerSizeTxt = GetAttr(sNode, "cs:style/@markerSize")
'
'    strokeHex = GetAttr(sNode, "cs:style/@strokeColor")
'    strokeWidthTxt = GetAttr(sNode, "cs:style/@strokeWidth")
'    markerOpacityTxt = GetAttr(sNode, "cs:style/@markerOpacity")
'    smoothTxt = LCase$(GetAttr(sNode, "cs:style/@smooth"))
'
'    invertNegTxt = LCase$(GetAttr(sNode, "cs:style/@invertNegative"))
'    invertColorHex = GetAttr(sNode, "cs:style/@invertColor")
'
'    On Error Resume Next
'
'    ' Fill & line color
'    Dim rgbCol As Long, rgbStroke As Long
'    If Len(colorHex) > 0 Then
'        rgbCol = HtmlColorToRGB(colorHex)
'        ser.Format.Fill.ForeColor.RGB = rgbCol
'        ser.Format.line.ForeColor.RGB = rgbCol
'        ser.MarkerForegroundColor = rgbCol
'        ser.MarkerBackgroundColor = rgbCol
'    End If
'    If Len(strokeHex) > 0 Then
'        rgbStroke = HtmlColorToRGB(strokeHex)
'        ser.MarkerForegroundColor = rgbStroke
'        ser.Format.line.ForeColor.RGB = rgbStroke
'    End If
'
'    ' Thickness
'    If Len(thicknessTxt) > 0 Then
'        Dim w As Single: w = ToSingleNumber(thicknessTxt)
'        If w > 0 Then
'            ser.Format.line.Weight = w
'            ser.Format.line.Visible = msoTrue
'        End If
'    End If
'    If Len(strokeWidthTxt) > 0 Then
'        Dim sw As Single: sw = ToSingleNumber(strokeWidthTxt)
'        If sw > 0 Then ser.Format.line.Weight = sw
'    End If
'
'    ' Line dash
'    If Len(styleTxt) > 0 Then
'        ser.Format.line.DashStyle = MapLineDashStyle(styleTxt)
'        ser.Format.line.Visible = msoTrue
'    End If
'
'    ' Markers
'    If Len(shapeTxt) > 0 Then ser.MarkerStyle = MapMarkerStyle(shapeTxt)
'    If Len(markerSizeTxt) > 0 Then
'        Dim ms As Long: ms = CLng(val(markerSizeTxt))
'        If ms > 0 Then ser.MarkerSize = ms
'    End If
'
'    ' Marker opacity (series fill transparency)
'    If Len(markerOpacityTxt) > 0 Then
'        Dim op As Double: op = CDbl(val(markerOpacityTxt))
'        If op < 0 Then op = 0
'        If op > 1 Then op = 1
'        ser.Format.Fill.Transparency = op
'    End If
'
'    ' Smooth (applies to line/xy)
'    If smoothTxt = "true" Then
'        ser.Smooth = True
'    End If
'
'    ' NEW: Invert fill for negative values (only if the rendered type is area-like)
'    If ser.chartType = xlArea Or ser.chartType = xlAreaStacked Or ser.chartType = xlAreaStacked100 Then
'        If invertNegTxt = "true" Then
'            ser.InvertIfNegative = True
'            If Len(invertColorHex) > 0 Then
'                ser.InvertColor = HtmlColorToRGB(invertColorHex)
'            End If
'        End If
'    End If
'
'    On Error GoTo 0
'End Sub
'
'' =================
'' NEW: ApplySeriesExtras
'' =================
'' Adds error bars (X/Y; symmetric/asymmetric/custom), data labels (x,y,name),
'' and trendlines. Custom error bar arrays are written to a hidden cache sheet.
'Private Sub ApplySeriesExtras(ser As Series, sNode As Object, ch As Chart, wb As Workbook)
'    On Error GoTo SafeExit
'
'    ' -------- Error Bars --------
'    ' XML:
'    ' <err type="both|x|y" mode="sym|asym|custom" cap="true|false">
'    '   <xPlus><n>...</n></xPlus>
'    '   <xMinus><n>...</n></xMinus>
'    '   <yPlus><n>...</n></yPlus>
'    '   <yMinus><n>...</n></yMinus>
'    '   @value="number"
'    ' </err>
'    Dim errNode As Object: Set errNode = sNode.SelectSingleNode("cs:err")
'    If Not errNode Is Nothing Then
'        Dim eType As String: eType = LCase$(GetAttr(errNode, "@type"))
'        Dim eMode As String: eMode = LCase$(GetAttr(errNode, "@mode"))
'        Dim eCap As String:  eCap = LCase$(GetAttr(errNode, "@cap"))
'        Dim eValTxt As String: eValTxt = GetAttr(errNode, "@value")
'
'        ' NEW: If series is rendered as a Line chart, ignore X-direction error bars
'        Dim isLineRendered As Boolean
'        isLineRendered = (ser.chartType = xlLine)
'
'        If isLineRendered And (eType = "x" Or eType = "both") Then
'            eType = "y" ' enforce Y-only
'        End If
'
'        Dim capBool As Boolean: capBool = (eCap = "true")
'        Dim rngAmt As Range, rngMinus As Range
'
'        If eMode = "custom" Then
'            Dim plusY() As Variant, minusY() As Variant
'
'            If eType = "y" Or eType = "both" Then
'                plusY = ReadArray(errNode, "cs:yPlus/cs:n", False)
'                minusY = ReadArray(errNode, "cs:yMinus/cs:n", False)
'
'                Dim cache As Worksheet: Set cache = EnsureCacheSheet(wb)
'                Dim baseKey As String: baseKey = MakeSafeKey(CStr(ser.name) & "_err_" & ser.Parent.name)
'
'                If Not IsEmpty(plusY) Then Set rngAmt = WriteVectorToRange(cache, baseKey & "_yp", plusY)
'                If Not IsEmpty(minusY) Then Set rngMinus = WriteVectorToRange(cache, baseKey & "_ym", minusY)
'
'                If Not rngAmt Is Nothing Then
'                    ser.HasErrorBars = True
'                    ser.ErrorBar Direction:=xlY, Include:=xlBoth, Type:=xlErrorBarTypeCustom, _
'                                  Amount:=rngAmt, MinusValues:=IIf(rngMinus Is Nothing, rngAmt, rngMinus)
'                    ser.ErrorBars.EndStyle = IIf(capBool, xlCap, xlNoCap)
'                End If
'            End If
'        Else
'            Dim valNum As Double: valNum = 0
'            If IsNumeric(eValTxt) Then valNum = CDbl(eValTxt)
'
'            If eType = "y" Or eType = "both" Then
'                ser.HasErrorBars = True
'                ser.ErrorBar Direction:=xlY, Include:=xlBoth, Type:=xlErrorBarTypeFixedValue, Amount:=valNum
'                ser.ErrorBars.EndStyle = IIf(capBool, xlCap, xlNoCap)
'            End If
'        End If
'    End If
'
'    ' -------- Data Labels --------
'    Dim dlNode As Object: Set dlNode = sNode.SelectSingleNode("cs:dataLabels")
'    If Not dlNode Is Nothing Then
'        Dim showLbl As Boolean: showLbl = (LCase$(GetAttr(dlNode, "@show")) = "true")
'        Dim fields As String: fields = LCase$(GetAttr(dlNode, "@fields"))
'
'        ser.HasDataLabels = showLbl
'        If showLbl Then
'            With ser.dataLabels
'                .ShowSeriesName = (InStr(1, fields, "name", vbTextCompare) > 0)
'                .ShowValue = (InStr(1, fields, "value", vbTextCompare) > 0 Or InStr(1, fields, "y", vbTextCompare) > 0)
'                .ShowCategoryName = (InStr(1, fields, "x", vbTextCompare) > 0)
'            End With
'
'            ' "xy" or both x and y explicitly requested
'            If InStr(1, fields, "xy", vbTextCompare) > 0 Or _
'               (InStr(1, fields, "x", vbTextCompare) > 0 And InStr(1, fields, "y", vbTextCompare) > 0) Then
'                Dim i As Long, xv As Variant, yv As Variant
'                xv = ser.XValues: yv = ser.Values
'                For i = LBound(yv) To UBound(yv)
'                    ser.Points(i).HasDataLabel = True
'                    ser.Points(i).DataLabel.text = Format$(xv(i)) & ", " & Format$(yv(i))
'                Next i
'            End If
'        End If
'    End If
'
'    ' -------- Trendlines --------
'    Dim tl As Object: Set tl = sNode.SelectSingleNode("cs:trendline")
'    If Not tl Is Nothing Then
'        Dim tType As String: tType = LCase$(GetAttr(tl, "@type"))
'        Dim orderTxt As String: orderTxt = GetAttr(tl, "@order")
'        Dim periodTxt As String: periodTxt = GetAttr(tl, "@period")
'        Dim interceptTxt As String: interceptTxt = GetAttr(tl, "@intercept")
'        Dim showEq As Boolean: showEq = (LCase$(GetAttr(tl, "@showEq")) = "true")
'        Dim showR2 As Boolean: showR2 = (LCase$(GetAttr(tl, "@showR2")) = "true")
'
'        Dim tr As Trendline
'        Select Case tType
'            Case "linear":      Set tr = ser.Trendlines.Add(Type:=xlLinear)
'            Case "exp", "exponential": Set tr = ser.Trendlines.Add(Type:=xlExponential)
'            Case "log", "logarithmic": Set tr = ser.Trendlines.Add(Type:=xlLogarithmic)
'            Case "power":       Set tr = ser.Trendlines.Add(Type:=xlPower)
'            Case "poly", "polynomial"
'                Set tr = ser.Trendlines.Add(Type:=xlPolynomial)
'                If IsNumeric(orderTxt) Then tr.Order = CLng(orderTxt)
'            Case "movingavg", "movingaverage"
'                Set tr = ser.Trendlines.Add(Type:=xlMovingAvg)
'                If IsNumeric(periodTxt) Then tr.Period = CLng(periodTxt)
'            Case Else
'                Set tr = ser.Trendlines.Add(Type:=xlLinear)
'        End Select
'
'        If Not tr Is Nothing Then
'            If IsNumeric(interceptTxt) Then tr.Intercept = CDbl(interceptTxt)
'            tr.DisplayEquation = showEq
'            tr.DisplayRSquared = showR2
'        End If
'    End If
'
'SafeExit:
'    On Error GoTo 0
'End Sub
'
'
'
'' -------- Histogram options (optional <histogram .../>) -----------------------
'' <histogram bins="auto|count|width" value="N|W" underflow="v" overflow="v"/>
'' -------- Histogram options (optional <histogram .../>) -----------------------
'' Supports: binning, cumulative display, labels, axis formatting
'Private Sub ApplyHistogramOptions(ch As Chart, doc As Object)
'    Dim h As Object: Set h = SelectNode(doc, "//cs:chart/cs:histogram")
'    If h Is Nothing Then Exit Sub
'
'    On Error Resume Next
'
'    Dim binsMode As String: binsMode = LCase$(GetAttr(h, "@bins"))
'    Dim binsValue As String: binsValue = GetAttr(h, "@value")
'    Dim underflow As String: underflow = GetAttr(h, "@underflow")
'    Dim overflow As String: overflow = GetAttr(h, "@overflow")
'    Dim cumulative As String: cumulative = LCase$(GetAttr(h, "@cumulative"))
'    Dim dataLabels As String: dataLabels = LCase$(GetAttr(h, "@dataLabels"))
'
'    Dim axisMin As String: axisMin = GetAttr(h, "@axisMin")
'    Dim axisMax As String: axisMax = GetAttr(h, "@axisMax")
'    Dim tickInterval As String: tickInterval = GetAttr(h, "@tickInterval")
'    Dim numberFormat As String: numberFormat = GetAttr(h, "@numberFormat")
'
'    Dim s As Series
'    For Each s In ch.SeriesCollection
'        ' --- Binning Mode ---
'        Select Case binsMode
'            Case "count"
'                s.BinType = xlBinsTypeBinCount
'                If IsNumeric(binsValue) Then s.BinCount = CLng(binsValue)
'            Case "width"
'                s.BinType = xlBinsTypeBinSize
'                If IsNumeric(binsValue) Then s.BinWidth = CDbl(binsValue)
'            Case Else
'                s.BinType = xlBinsTypeAutomatic
'        End Select
'
'        ' --- Overflow / Underflow ---
'        If IsNumeric(overflow) Then
'            s.HasOverflowBin = True
'            s.OverflowBin = CDbl(overflow)
'        End If
'        If IsNumeric(underflow) Then
'            s.HasUnderflowBin = True
'            s.UnderflowBin = CDbl(underflow)
'        End If
'
'        ' --- Cumulative Display (Pareto Style) ---
'        If cumulative = "true" Then
'            s.cumulative = True
'        Else
'            s.cumulative = False
'        End If
'
'        ' --- Data Labels ---
'        If dataLabels = "true" Then
'            s.HasDataLabels = True
'            s.dataLabels.ShowValue = True
'        Else
'            s.HasDataLabels = False
'        End If
'    Next s
'
'    ' --- Axis Customization ---
'    With ch.Axes(xlCategory)
'        If IsNumeric(axisMin) Then .MinimumScale = CDbl(axisMin)
'        If IsNumeric(axisMax) Then .MaximumScale = CDbl(axisMax)
'        If IsNumeric(tickInterval) Then .MajorUnit = CDbl(tickInterval)
'        If Len(numberFormat) > 0 Then .TickLabels.numberFormat = numberFormat
'    End With
'
'    On Error GoTo 0
'End Sub
'
'
'' -------- Layout helpers ------------------------------------------------------
'' =================
'' EDITED: ApplyLayout (fix drop lines to use ChartGroups.HasDropLines)
'' =================
'Private Sub ApplyLayout(ch As Chart, doc As Object)
'    On Error GoTo CleanFail
'
'    ' --- Chart Title ---
'    Dim titleText As String
'    titleText = GetAttr(doc, "//cs:layout/cs:title/@text")
'    If Len(titleText) = 0 Then titleText = "add name here"
'    ch.HasTitle = True
'    ch.chartTitle.text = titleText
'
'    ' --- Legend ---
'    Dim legendShow As String
'    legendShow = GetAttr(doc, "//cs:layout/cs:legend/@show")
'    If Len(legendShow) > 0 Then
'        ch.HasLegend = (LCase$(legendShow) = "true")
'    Else
'        ch.HasLegend = False
'    End If
'    Dim legendOrientation As String
'    legendOrientation = LCase$(GetAttr(doc, "//cs:layout/cs:opt[@key='legend_orientation']/@value"))
'    If legendOrientation = "h" Then
'        ch.Legend.Position = xlLegendPositionTop
'    ElseIf legendOrientation = "v" Then
'        ch.Legend.Position = xlLegendPositionRight
'    End If
'
'    ' --- Axis Titles ---
'    Dim xOverride As String, yOverride As String, y2Override As String
'    xOverride = GetAttr(doc, "//cs:layout/cs:axisTitles/@x")
'    yOverride = GetAttr(doc, "//cs:layout/cs:axisTitles/@y")
'    y2Override = GetAttr(doc, "//cs:layout/cs:axisTitles/@y2")
'
'    Dim xTitle As String, yTitle As String, y2Title As String
'    xTitle = IIf(Len(xOverride) > 0, xOverride, GetAttr(doc, "//cs:chart/cs:categoryAxis/@title"))
'    yTitle = IIf(Len(yOverride) > 0, yOverride, GetAttr(doc, "//cs:chart/cs:valueAxis/@title"))
'    y2Title = IIf(Len(y2Override) > 0, y2Override, GetAttr(doc, "//cs:chart/cs:valueAxis2/@title"))
'
'    If Len(xTitle) = 0 Then xTitle = "add name here"
'    If Len(yTitle) = 0 Then yTitle = "add name here"
'    If Len(y2Title) = 0 Then y2Title = "add name here"
'
'    On Error Resume Next
'    ch.Axes(xlCategory).HasTitle = True
'    ch.Axes(xlCategory).AxisTitle.Characters.text = xTitle
'    ch.Axes(xlValue, xlPrimary).HasTitle = True
'    ch.Axes(xlValue, xlPrimary).AxisTitle.Characters.text = yTitle
'    If AxisExists(ch, xlValue, xlSecondary) Then
'        ch.Axes(xlValue, xlSecondary).HasTitle = True
'        ch.Axes(xlValue, xlSecondary).AxisTitle.Characters.text = y2Title
'    End If
'    On Error GoTo CleanFail
'
'    ' --- Bar overlay/grouping (unchanged) ---
'    Dim barMode As String
'    barMode = LCase$(GetAttr(doc, "//cs:layout/cs:opt[@key='barmode']/@value"))
'    If barMode = "overlay" Then
'        Dim ser As Series
'        For Each ser In ch.SeriesCollection
'            ser.Format.Fill.Transparency = 0.45
'        Next ser
'        If ch.ChartGroups.count > 0 Then
'            ch.ChartGroups(1).Overlap = 100
'            ch.ChartGroups(1).GapWidth = 10
'        End If
'    ElseIf barMode = "group" Then
'        If ch.ChartGroups.count > 0 Then
'            ch.ChartGroups(1).Overlap = 0
'            ch.ChartGroups(1).GapWidth = 150
'        End If
'    End If
'
'    ' --- connect-gaps/empty-cells behavior ---
'    Dim cg As String: cg = LCase$(GetAttr(doc, "//cs:layout/cs:opt[@key='connect_gaps']/@value"))
'    Select Case cg
'        Case "interpolate": ch.DisplayBlanksAs = xlInterpolated
'        Case "zero":        ch.DisplayBlanksAs = xlZero
'        Case "gap":         ch.DisplayBlanksAs = xlNotPlotted
'    End Select
'
'    ' --- drop lines toggle (use ChartGroups.HasDropLines; Chart.HasDropLines is not valid) ---
'    Dim dl As String: dl = LCase$(GetAttr(doc, "//cs:layout/cs:opt[@key='drop_lines']/@value"))
'    If dl = "true" Or dl = "false" Then
'        Dim wantDL As Boolean: wantDL = (dl = "true")
'        Dim cgx As ChartGroup
'        For Each cgx In ch.ChartGroups
'            On Error Resume Next
'            cgx.HasDropLines = wantDL
'            On Error GoTo 0
'        Next cgx
'    End If
'    Exit Sub
'
'CleanFail:
'    Debug.Print "ApplyLayout error: " & Err.Description
'End Sub
'
'
'
'
'' -------- Utility -------------------------------------------------------------
'Private Function HtmlColorToRGB(hexColor As String) As Long
'    Dim r As Long, g As Long, b As Long
'    If Left$(hexColor, 1) = "#" Then hexColor = Mid$(hexColor, 2)
'    If Len(hexColor) = 3 Then
'        ' Short hex #RGB ? expand to #RRGGBB
'        hexColor = Mid$(hexColor, 1, 1) & Mid$(hexColor, 1, 1) & _
'                   Mid$(hexColor, 2, 1) & Mid$(hexColor, 2, 1) & _
'                   Mid$(hexColor, 3, 1) & Mid$(hexColor, 3, 1)
'    End If
'    r = CLng("&H" & Mid$(hexColor, 1, 2))
'    g = CLng("&H" & Mid$(hexColor, 3, 2))
'    b = CLng("&H" & Mid$(hexColor, 5, 2))
'    HtmlColorToRGB = RGB(r, g, b)
'End Function
'
'Private Function ParseIso8601Date(ByVal s As String) As Variant
'    On Error GoTo Fail
'    s = Replace$(s, "Z", "")
'    s = Replace$(s, "T", " ")
'    ParseIso8601Date = CDate(s)
'    Exit Function
'Fail:
'    ParseIso8601Date = Empty
'End Function
'
'Private Function IsCategoryDate(doc As Object) As Boolean
'    IsCategoryDate = (LCase$(GetAttr(doc, "//cs:chart/cs:categoryAxis/@kind")) = "date")
'End Function
'
'' Returns a purely numeric 1-D array for histogram Values, or Empty if unusable.
'Private Function GetNumericVector(primary As Variant, fallback As Variant) As Variant
'    Dim src As Variant
'    If IsEmpty(primary) Then
'        If IsEmpty(fallback) Then Exit Function
'        src = fallback
'    Else
'        src = primary
'    End If
'
'    Dim i As Long, n As Long
'    On Error GoTo Bad
'    n = UBound(src) - LBound(src) + 1
'    Dim out() As Double
'    ReDim out(1 To n)
'
'    Dim k As Long: k = 0
'    For i = LBound(src) To UBound(src)
'        If IsNumeric(src(i)) Then
'            k = k + 1: out(k) = CDbl(src(i))
'        ElseIf TypeName(src(i)) = "String" Then
'            Dim s As String: s = Trim$(CStr(src(i)))
'            If Len(s) > 0 And IsNumeric(s) Then
'                k = k + 1: out(k) = CDbl(s)
'            End If
'        End If
'        ' non-numeric tokens are dropped
'    Next i
'
'    If k = 0 Then Exit Function
'    If k < n Then ReDim Preserve out(1 To k)
'    GetNumericVector = out
'    Exit Function
'Bad:
'    GetNumericVector = Empty
'End Function
'
'Private Function MapMarkerStyle(shapeTxt As String) As XlMarkerStyle
'    Select Case LCase$(shapeTxt)
'        Case "circle", "o": MapMarkerStyle = xlMarkerStyleCircle
'        Case "square":      MapMarkerStyle = xlMarkerStyleSquare
'        Case "diamond":     MapMarkerStyle = xlMarkerStyleDiamond
'        Case "triangle", "tri": MapMarkerStyle = xlMarkerStyleTriangle
'        Case "x":           MapMarkerStyle = xlMarkerStyleX
'        Case "plus", "+":   MapMarkerStyle = xlMarkerStylePlus
'        Case "star":        MapMarkerStyle = xlMarkerStyleStar
'        Case "dot", "point": MapMarkerStyle = xlMarkerStyleDot
'        Case "none":        MapMarkerStyle = xlMarkerStyleNone
'        Case Else:          MapMarkerStyle = xlMarkerStyleAutomatic
'    End Select
'End Function
'
'Private Function MapLineDashStyle(styleTxt As String) As MsoLineDashStyle
'    Select Case LCase$(styleTxt)
'        Case "solid":       MapLineDashStyle = msoLineSolid
'        Case "dash":        MapLineDashStyle = msoLineDash
'        Case "dot":         MapLineDashStyle = msoLineSysDot
'        Case "dashdot":     MapLineDashStyle = msoLineDashDot
'        Case "dashdotdot":  MapLineDashStyle = msoLineDashDotDot
'        Case Else:          MapLineDashStyle = msoLineSolid
'    End Select
'End Function
'
'Private Function ToSingleNumber(txt As String) As Single
'    Dim v As Double
'    If IsNumeric(txt) Then
'        v = CDbl(txt)
'    Else
'        v = CDbl(val(txt))
'    End If
'    If v < 0 Then v = 0
'    ToSingleNumber = CSng(v)
'End Function
'
'Private Function DetermineBaseType(sNodes As Object) As XlChartType
'    Dim s As Object
'    Dim foundHist As Boolean, foundBar As Boolean, foundScatter As Boolean, foundLine As Boolean, foundArea As Boolean
'
'    For Each s In sNodes
'        Dim k As String
'        k = LCase$(GetAttr(s, "cs:style/@kind"))
'        Select Case k
'            Case "hist", "histogram": foundHist = True
'            Case "bar", "column": foundBar = True
'            Case "scatter": foundScatter = True
'            Case "line": foundLine = True
'            Case "area": foundArea = True
'        End Select
'    Next
'
'    ' Priority rules: histogram ? bar ? scatter ? area ? line
'    If foundHist Or foundBar Then
'        DetermineBaseType = xlColumnClustered
'    ElseIf foundScatter Then
'        DetermineBaseType = xlXYScatter
'    ElseIf foundArea Then
'        DetermineBaseType = xlArea
'    ElseIf foundLine Then
'        DetermineBaseType = xlLine
'    Else
'        DetermineBaseType = xlLine
'    End If
'End Function
'
'Private Function FindChartByTitle(ws As Worksheet, titleText As String) As ChartObject
'    Dim co As ChartObject
'    For Each co In ws.ChartObjects
'        If co.Chart.HasTitle Then
'            If StrComp(co.Chart.chartTitle.text, titleText, vbTextCompare) = 0 Then
'                Set FindChartByTitle = co
'                Exit Function
'            End If
'        End If
'    Next
'    Set FindChartByTitle = Nothing
'End Function
'
'Private Sub ApplyBarmodeOverlay(ch As Chart)
'    Dim ser As Series
'    For Each ser In ch.SeriesCollection
'        ' Semi-transparent fill for overlay
'        ser.Format.Fill.Transparency = 0.45
'    Next ser
'
'    ' Turn on series overlap
'    ch.ChartGroups(1).Overlap = 100
'    ch.ChartGroups(1).GapWidth = 10
'End Sub
'Private Function AxisExists(ch As Chart, axisType As XlAxisType, Optional axisGroup As XlAxisGroup = xlPrimary) As Boolean
'    Dim ax As Axis
'    On Error Resume Next
'    Set ax = ch.Axes(axisType, axisGroup)
'    AxisExists = Not ax Is Nothing
'    On Error GoTo 0
'End Function
'' =================
'' NEW helpers (self-contained, no side effects on existing features)
'' =================
'
'' Hidden cache sheet for custom error bars
'Private Function EnsureCacheSheet(wb As Workbook) As Worksheet
'    Const sh As String = "_ChartSpecCache"
'    On Error Resume Next
'    Set EnsureCacheSheet = wb.Worksheets(sh)
'    On Error GoTo 0
'    If EnsureCacheSheet Is Nothing Then
'        Set EnsureCacheSheet = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.count))
'        EnsureCacheSheet.name = sh
'        EnsureCacheSheet.Visible = xlSheetVeryHidden
'    End If
'End Function
'
'' Write a 1-D variant array to a column range; returns the Range written
'Private Function WriteVectorToRange(ws As Worksheet, key As String, vec As Variant) As Range
'    Dim n As Long, i As Long
'    n = UBound(vec) - LBound(vec) + 1
'    ' Find or create a named column by key
'    Dim tgt As Range
'    Set tgt = ws.Cells(1, ws.Columns.count).End(xlToLeft)
'    If tgt.Column = ws.Columns.count And Len(tgt.Value2) <> 0 Then
'        Set tgt = ws.Cells(1, ws.Columns.count).Offset(0, 0)
'    Else
'        Set tgt = tgt.Offset(0, 1)
'    End If
'    ws.Cells(1, tgt.Column).Value2 = key
'    Dim r As Range
'    Set r = ws.Range(ws.Cells(2, tgt.Column), ws.Cells(n + 1, tgt.Column))
'    For i = 1 To n
'        r.Cells(i, 1).Value2 = vec(LBound(vec) + i - 1)
'    Next i
'    Set WriteVectorToRange = r
'End Function
'
'' Normalize text to a simple key
'Private Function MakeSafeKey(ByVal s As String) As String
'    s = Replace$(s, " ", "_")
'    s = Replace$(s, ":", "_")
'    s = Replace$(s, "/", "_")
'    s = Replace$(s, "\", "_")
'    s = Replace$(s, ".", "_")
'    MakeSafeKey = s
'End Function
'
'' =================
'' NEW: Per-point styling for series
'' =================
'' Optional XML:
''   <pt i="3">
''     <style color="#RRGGBB" strokeColor="#RRGGBB" strokeWidth="2" markerOpacity="0.3"/>
''   </pt>
'Private Sub ApplySeriesPointStyles(ser As Series, sNode As Object)
'    On Error GoTo done
'    Dim pts As Object, p As Object
'    Set pts = sNode.SelectNodes("cs:pt")
'    If pts Is Nothing Then GoTo done
'    If pts.Length = 0 Then GoTo done
'
'    Dim idxTxt As String, i As Long
'    For Each p In pts
'        idxTxt = GetAttr(p, "@i")
'        If IsNumeric(idxTxt) Then
'            i = CLng(idxTxt)
'            If i >= 1 And i <= ser.Points.count Then
'                Dim c As String, sc As String, sw As String, op As String
'                c = GetAttr(p, "cs:style/@color")
'                sc = GetAttr(p, "cs:style/@strokeColor")
'                sw = GetAttr(p, "cs:style/@strokeWidth")
'                op = GetAttr(p, "cs:style/@markerOpacity")
'
'                With ser.Points(i).Format
'                    If Len(c) > 0 Then
'                        .Fill.ForeColor.RGB = HtmlColorToRGB(c)
'                        .line.ForeColor.RGB = HtmlColorToRGB(c)
'                    End If
'                    If Len(sc) > 0 Then
'                        .line.ForeColor.RGB = HtmlColorToRGB(sc)
'                    End If
'                    If Len(sw) > 0 And IsNumeric(sw) Then
'                        .line.Weight = CSng(sw)
'                        .line.Visible = msoTrue
'                    End If
'                    If Len(op) > 0 And IsNumeric(op) Then
'                        Dim f As Double: f = CDbl(op)
'                        If f < 0 Then f = 0
'                        If f > 1 Then f = 1
'                        .Fill.Transparency = f
'                    End If
'                End With
'            End If
'        End If
'    Next
'done:
'    On Error GoTo 0
'End Sub
'
'
'


Option Explicit

'================== MAIN: BUILD CHART ==================
Public Sub BuildChartFromXML(xmlPath As String, targetSheet As Worksheet, _
                             Optional sizeRange As Range = Nothing)
    Dim doc As Object
    Dim root As Object
    Dim chartTypeCat As String
    Dim chtObj As ChartObject
    Dim traceNode As Object, styleNode As Object
    Dim traceIDs As Object
    Dim chartTypeEnum As XlChartType
    
    ' Load XML
    Set doc = CreateObject("MSXML2.DOMDocument.6.0")
    doc.async = False
    doc.validateOnParse = False
    
    If Not doc.Load(xmlPath) Then
        MsgBox "XML error: " & doc.ParseError.reason, vbCritical
        Exit Sub
    End If

    ' Validate root
    Set root = doc.SelectSingleNode("/plotly_excel_chart")
    If root Is Nothing Then
        MsgBox "Invalid XML: Missing root node <plotly_excel_chart>", vbCritical
        Exit Sub
    End If
    
    If root.Attributes.getNamedItem("version") Is Nothing _
       Or root.Attributes.getNamedItem("version").text <> "2.0" Then
        MsgBox "Unsupported XML version. Expected version='2.0'", vbCritical
        Exit Sub
    End If
    
    ' Track trace IDs for uniqueness
    Set traceIDs = CreateObject("Scripting.Dictionary")

    ' Read global chart category type
    chartTypeCat = LCase$(SafeNodeText(doc, "//chart_meta/chart_type"))
    chartTypeEnum = MapChartCategoryToXlType(chartTypeCat)
    ' Dummy sentinel check used below
    If chartTypeEnum = xlAreaStacked Then
        ' No-op. MapChartCategoryToXlType returns valid enums or a default.
    End If
    
    '============ Chart Position and Size ============
    Dim chartLeft As Double, chartTop As Double
    Dim chartWidth As Double, chartHeight As Double

    If Not sizeRange Is Nothing Then
        ' Use provided range for position and sizing
        chartLeft = sizeRange.Left
        chartTop = sizeRange.Top
        chartWidth = sizeRange.Width
        chartHeight = sizeRange.Height
    Else
        ' Default behavior if no range specified
        chartLeft = 50
        chartTop = 50
        chartWidth = 800
        chartHeight = 500
    End If

    ' Create chart object
    Set chtObj = targetSheet.ChartObjects.Add(Left:=chartLeft, Top:=chartTop, Width:=chartWidth, Height:=chartHeight)
    On Error Resume Next
    chtObj.Chart.chartType = chartTypeEnum
    If Err.Number <> 0 Then
        Err.Clear
        chtObj.Chart.chartType = xlXYScatter
    End If
    On Error GoTo 0

    ' Chart title
    Dim chartTitle As String
    chartTitle = SafeNodeText(doc, "//chart_meta/title")
    If Len(chartTitle) > 0 Then
        chtObj.Chart.HasTitle = True
        chtObj.Chart.chartTitle.text = chartTitle
    Else
        chtObj.Chart.HasTitle = False
    End If
    
    ' export_timestamp is informational
    Dim exportTs As String
    exportTs = SafeNodeText(doc, "//chart_meta/export_timestamp")
    If Len(exportTs) > 0 Then Debug.Print "Export timestamp: "; exportTs

    '================= LOOP THROUGH TRACES ===============
    Dim s As Series
    For Each traceNode In doc.SelectNodes("//traces/trace")
        Dim traceID As String
        Dim stype As String
        Dim xCSV As String, yCSV As String, zCSV As String, sizeCSV As String, textCSV As String
        Dim xVals As Variant, yVals As Variant, sizeVals As Variant, textVals As Variant
        Dim asCategories As Boolean
        
        traceID = ""
        If Not traceNode.Attributes Is Nothing Then
            If Not traceNode.Attributes.getNamedItem("id") Is Nothing Then
                traceID = CStr(traceNode.Attributes.getNamedItem("id").text)
            End If
        End If
        
        ' Validate trace ID: required, numeric, unique
        If traceID = "" Or Not IsNumeric(traceID) Or traceIDs.Exists(traceID) Then
            MsgBox "Invalid or duplicate trace id: " & traceID, vbCritical
            Exit Sub
        End If
        traceIDs.Add traceID, True
        
        ' Extract data CSV
        xCSV = SafeNodeText(traceNode, "data/x")
        yCSV = SafeNodeText(traceNode, "data/y")
        zCSV = SafeNodeText(traceNode, "data/z")
        sizeCSV = SafeNodeText(traceNode, "data/size")
        textCSV = SafeNodeText(traceNode, "data/text")
        
        ' Series style node
        Set styleNode = traceNode.SelectSingleNode("style")
        stype = LCase$(SafeNodeText(traceNode, "style/series_type", "scatter"))
        
        ' Build arrays according to expected type
        ' For bar/column/pie/histogram: X may be categorical labels
        asCategories = IsCategoricalSeriesType(stype)
        If asCategories Then
            BuildXYForCategories xCSV, yCSV, xVals, yVals
        Else
            BuildXYNumericFlexible xCSV, yCSV, xVals, yVals
        End If
        
        ' If no usable data, skip trace
        If IsEmpty(yVals) Then GoTo NextTrace
        
        Set s = chtObj.Chart.SeriesCollection.NewSeries
        
        ' Assign data
        On Error Resume Next
        s.Values = yVals
        If Not IsEmpty(xVals) Then s.XValues = xVals
        On Error GoTo 0
        
        '================ HISTOGRAM CATEGORY LABEL FIX ================
        ' For histogram: replace numeric X axis tick labels with <data/text>
        If stype = "histogram" Then
            textVals = CSVToStringArray(textCSV)
            If Not IsEmpty(textVals) Then
                On Error Resume Next
                ' Apply custom human-readable labels
                With chtObj.Chart.Axes(xlCategory)
                    .CategoryNames = textVals
                    ' Hide default midpoints so only custom labels remain
                    .TickLabelSpacing = 1
                End With
                On Error GoTo 0
            End If
        End If
        '=============================================================
        
        ' Optional bubble size
        If stype = "bubble" Then
            sizeVals = CSVToNumericArray(sizeCSV)
            If Not IsEmpty(sizeVals) Then
                On Error Resume Next
                s.BubbleSizes = sizeVals
                On Error GoTo 0
            End If
        End If
        
        ' Optional text labels (skip for histogram because <text> is used as axis labels)
        textVals = CSVToStringArray(textCSV)
        If stype <> "histogram" Then
            If Not IsEmpty(textVals) Then
                On Error Resume Next
                s.ApplyDataLabels
                s.dataLabels.ShowValue = False
                s.dataLabels.ShowSeriesName = False
                s.dataLabels.ShowCategoryName = False
                Dim iPt As Long, nPts As Long
                nPts = Application.Min(s.Points.count, SafeUBound1(textVals) + 1)
                For iPt = 1 To nPts
                    s.Points(iPt).HasDataLabel = True
                    s.Points(iPt).DataLabel.text = CStr(textVals(iPt - 1))
                Next iPt
                On Error GoTo 0
            End If
        End If

        ' Series name
        Dim sName As String
        sName = SafeNodeText(styleNode, "name")
        If Len(sName) > 0 Then
            On Error Resume Next
            s.name = sName
            On Error GoTo 0
        End If
        
        ' Axis group
        Dim axisGroup As String
        axisGroup = LCase$(SafeNodeText(styleNode, "axis_group", "primary"))
        On Error Resume Next
        If axisGroup = "secondary" Then
            s.axisGroup = xlSecondary
        Else
            s.axisGroup = xlPrimary
        End If
        On Error GoTo 0
        
        ' Per-series chart type
        ApplySeriesChartType s, stype
        
        ' Visibility
        Dim visStr As String
        visStr = LCase$(SafeNodeText(styleNode, "visibility", "true"))
        If visStr = "false" Then
            On Error Resume Next
            ' Excel 2013+: hide via IsFiltered
            s.IsFiltered = True
            If Err.Number <> 0 Then
                Err.Clear
                ' Fallback: remove visual cues
                s.MarkerStyle = xlMarkerStyleNone
                s.Format.line.Visible = msoFalse
                s.Format.Fill.Visible = msoFalse
            End If
            On Error GoTo 0
        End If
        
        ' Line properties
        ApplyLineStyling s, styleNode, stype
        
        ' Marker properties
        ApplyMarkerStyling s, styleNode
        
        ' Fill and area properties (for area/columns/bars/markers fill)
        ApplyFillStyling s, styleNode

NextTrace:
    Next traceNode
    
    '================== SHAPES / EXTRAS ==================
    Call RenderExtrasAnnotations(doc, chtObj)



    '================= X AXIS =================
    If chartTypeCat <> "pie" Then
        ApplyAxisSettings chtObj.Chart, doc, True
    End If

    '================= Y AXIS =================
    If chartTypeCat <> "pie" Then
        ApplyAxisSettings chtObj.Chart, doc, False
    End If

    '================== LEGEND ==================
    ApplyLegendSettings chtObj.Chart, doc
End Sub


'==================== AXIS APPLY =======================
Private Sub ApplyAxisSettings(ByVal ch As Chart, ByVal doc As Object, ByVal isX As Boolean)
    Dim axisPath As String
    Dim ax As Axis
    Dim aMin As String, aMax As String, aLog As String, aTitle As String

    If isX Then
        axisPath = "//chart_meta/x_axis"
        Set ax = ch.Axes(xlCategory)
    Else
        axisPath = "//chart_meta/y_axis"
        Set ax = ch.Axes(xlValue)
    End If

    aTitle = SafeNodeText(doc, axisPath & "/title")
    ax.HasTitle = (Len(aTitle) > 0)
    If ax.HasTitle Then ax.AxisTitle.text = aTitle

    aMin = SafeNodeText(doc, axisPath & "/min")
    aMax = SafeNodeText(doc, axisPath & "/max")
    aLog = LCase$(SafeNodeText(doc, axisPath & "/log_scale", "false"))

    '=== AUTO-DETECT DATE FORMAT FOR X-AXIS ===
    If isX Then
        Dim firstXVal As String
        firstXVal = SafeNodeText(doc, "//traces/trace[1]/data/x")
        
        If Len(firstXVal) > 0 Then
            Dim firstToken As String
            firstToken = Trim$(Split(firstXVal, ",")(0)) ' first date in CSV
            If IsDate(firstToken) Then
                Dim inferredFormat As String
                inferredFormat = InferDateFormat(firstToken)
                On Error Resume Next
                ax.TickLabels.numberFormat = inferredFormat
                On Error GoTo 0
            End If
        End If
    End If

    '=== Apply log/linear scale and min/max ===
    On Error Resume Next
    If aLog = "true" Then
        ax.ScaleType = xlScaleLogarithmic
        ax.LogBase = 10
        If Err.Number <> 0 Then
            Err.Clear
            ax.ScaleType = xlLinear
        End If
    Else
        ax.ScaleType = xlLinear
    End If

    If aLog = "true" And ax.ScaleType = xlScaleLogarithmic Then
        If IsPositiveNumeric(aMin) Then
            ax.MinimumScale = CDbl(aMin)
        Else
            ax.MinimumScaleIsAuto = True
        End If
        If IsPositiveNumeric(aMax) Then
            ax.MaximumScale = CDbl(aMax)
        Else
            ax.MaximumScaleIsAuto = True
        End If
    Else
        If IsNumeric(aMin) And Len(aMin) > 0 Then
            ax.MinimumScale = CDbl(aMin)
        Else
            ax.MinimumScaleIsAuto = True
        End If
        If IsNumeric(aMax) And Len(aMax) > 0 Then
            ax.MaximumScale = CDbl(aMax)
        Else
            ax.MaximumScaleIsAuto = True
        End If
    End If
    On Error GoTo 0
End Sub


'================== LEGEND SETTINGS ====================
Private Sub ApplyLegendSettings(ByVal ch As Chart, ByVal doc As Object)
    Dim legVis As String, legPos As String
    
    legVis = LCase$(SafeNodeText(doc, "//chart_meta/legend/visible", "true"))
    legPos = LCase$(SafeNodeText(doc, "//chart_meta/legend/position", "right"))
    
    ch.HasLegend = (legVis = "true")
    If ch.HasLegend Then
        Select Case legPos
            Case "bottom": ch.Legend.Position = xlLegendPositionBottom
            Case "right":  ch.Legend.Position = xlLegendPositionRight
            Case Else:     ch.Legend.Position = xlLegendPositionRight
        End Select
    End If
End Sub

'================== SERIES TYPE MAPPING ================
Private Sub ApplySeriesChartType(ByVal s As Series, ByVal stype As String)
    On Error Resume Next
    Select Case LCase$(stype)
        Case "scatter":                    s.chartType = xlXYScatter
        Case "scatter_lines":              s.chartType = xlXYScatterLines
        Case "scatter_lines_markers":      s.chartType = xlXYScatterLines
                                          s.MarkerStyle = xlMarkerStyleCircle
        Case "line":                       s.chartType = xlLine
        Case "line_stacked":               s.chartType = xlLineStacked
        Case "area":                       s.chartType = xlArea
        Case "area_stacked":               s.chartType = xlAreaStacked
        Case "bar":                        s.chartType = xlBarClustered
        Case "bar_stacked":                s.chartType = xlBarStacked
        Case "column":                     s.chartType = xlColumnClustered
        Case "column_stacked":             s.chartType = xlColumnStacked
        Case "bubble":                     s.chartType = xlBubble
        Case "pie":                        s.chartType = xlPie
        Case "box":                        ' No native Excel chart; fall back
                                          s.chartType = xlXYScatter
        Case "histogram":                  s.chartType = xlColumnClustered
        Case "heatmap":                    ' Not a native Excel chart; notify and fall back
                                          Debug.Print "Heatmap not natively supported; using column clustered."
                                          s.chartType = xlColumnClustered
        Case "waterfall":                  s.chartType = xlWaterfall
        Case Else:                         s.chartType = xlXYScatter
    End Select
    On Error GoTo 0
End Sub

Private Function MapChartCategoryToXlType(ByVal cat As String) As XlChartType
    Select Case LCase$(cat)
        Case "xy", "scatter": MapChartCategoryToXlType = xlXYScatter
        Case "bar":           MapChartCategoryToXlType = xlBarClustered
        Case "line":          MapChartCategoryToXlType = xlLine
        Case "area":          MapChartCategoryToXlType = xlArea
        Case "pie":           MapChartCategoryToXlType = xlPie
        Case Else:            MapChartCategoryToXlType = xlXYScatter
    End Select
End Function

Private Function IsCategoricalSeriesType(ByVal stype As String) As Boolean
    Select Case LCase$(stype)
        Case "bar", "bar_stacked", "column", "column_stacked", "pie", "histogram", "waterfall"
            IsCategoricalSeriesType = True
        Case Else
            IsCategoricalSeriesType = False
    End Select
End Function

'================== STYLING HELPERS ====================
Private Sub ApplyLineStyling(ByVal s As Series, ByVal styleNode As Object, ByVal stype As String)
    Dim hasLine As Boolean
    Dim lineColor As String, lineStyle As String
    Dim lineWidth As Double
    
    hasLine = (stype = "scatter_lines" Or stype = "scatter_lines_markers" Or _
               stype = "line" Or stype = "line_stacked" Or stype = "area" Or _
               stype = "area_stacked" Or stype = "bar" Or stype = "bar_stacked" Or _
               stype = "column" Or stype = "column_stacked" Or stype = "waterfall")
    
    lineColor = SafeNodeText(styleNode, "line_color", "#000000")
    lineStyle = SafeNodeText(styleNode, "line_style", "solid")
    lineWidth = val(SafeNodeText(styleNode, "line_width", "1"))
    If lineWidth <= 0 Then lineWidth = 1
    
    On Error Resume Next
    If hasLine Then
        With s.Format.line
            .Visible = msoTrue
            .ForeColor.RGB = ParseColor(lineColor)
            .Weight = lineWidth
            .DashStyle = DashToMsoLineStyle(lineStyle)
        End With
    Else
        s.Format.line.Visible = msoFalse
    End If
    On Error GoTo 0
End Sub

Private Sub ApplyMarkerStyling(ByVal s As Series, ByVal styleNode As Object)
    Dim mNode As Object
    Dim mSizeStr As String, mColor As String, mShape As String
    Dim mSize As Double
    
    If styleNode Is Nothing Then Exit Sub
    Set mNode = styleNode.SelectSingleNode("marker")
    If mNode Is Nothing Then Exit Sub
    
    mSizeStr = SafeNodeText(mNode, "size", "6")
    mColor = SafeNodeText(mNode, "color", "#000000")
    mShape = SafeNodeText(mNode, "shape", "circle")
    
    ' If comma-separated sizes, take the first as a reasonable default
    If InStr(1, mSizeStr, ",") > 0 Then
        mSize = val(Split(mSizeStr, ",")(0))
    Else
        mSize = val(mSizeStr)
    End If
    If mSize < 0 Then mSize = 0
    
    On Error Resume Next
    If mSize = 0 Then
        s.MarkerStyle = xlMarkerStyleNone
    Else
        s.MarkerStyle = ShapeToMarkerStyle(mShape)
        s.MarkerSize = mSize
        s.Format.Fill.Visible = msoTrue
        s.Format.Fill.ForeColor.RGB = ParseColor(mColor)
    End If
    On Error GoTo 0
End Sub

Private Sub ApplyFillStyling(ByVal s As Series, ByVal styleNode As Object)
    Dim fillColor As String
    Dim fillOpacityStr As String
    Dim hasFill As Boolean
    
    If styleNode Is Nothing Then Exit Sub
    
    fillColor = SafeNodeText(styleNode, "fill_color")
    fillOpacityStr = SafeNodeText(styleNode, "fill_opacity")
    
    hasFill = (Len(fillColor) > 0 Or Len(fillOpacityStr) > 0)
    If Not hasFill Then Exit Sub
    
    On Error Resume Next
    With s.Format.Fill
        .Visible = msoTrue
        If Len(fillColor) > 0 Then .ForeColor.RGB = ParseColor(fillColor)
        If Len(fillOpacityStr) > 0 Then
            Dim t As Double
            t = CDbl(fillOpacityStr) ' spec: 0..1 transparency
            If t < 0 Then t = 0
            If t > 1 Then t = 1
            .Transparency = t
        End If
    End With
    On Error GoTo 0
End Sub

'==================== UTIL: SAFE NODE ==================
Private Function SafeNodeText(ByVal parent As Object, ByVal nodeName As String, Optional ByVal defaultVal As String = "") As String
    Dim nd As Object
    On Error Resume Next
    Set nd = parent.SelectSingleNode(nodeName)
    On Error GoTo 0
    If nd Is Nothing Then
        SafeNodeText = defaultVal
    Else
        SafeNodeText = Trim(nd.text)
        If Len(SafeNodeText) = 0 Then SafeNodeText = defaultVal
    End If
End Function




'==================== UTIL: COLORS =====================
Private Function ParseColor(ByVal colorSpec As String) As Long
    colorSpec = Trim$(colorSpec)
    If Len(colorSpec) = 0 Then
        ParseColor = RGB(0, 0, 0)
        Exit Function
    End If
    
    '=== Handle hex format ===
    If Left$(colorSpec, 1) = "#" And Len(colorSpec) = 7 Then
        ParseColor = HexToRGB(colorSpec)
        Exit Function
    End If
    
    '=== Handle rgb(r,g,b) format ===
    If LCase$(Left$(colorSpec, 4)) = "rgb(" And right$(colorSpec, 1) = ")" Then
        Dim inner As String, parts As Variant
        inner = Mid$(colorSpec, 5, Len(colorSpec) - 5)
        parts = Split(inner, ",")
        If UBound(parts) = 2 Then
            Dim r As Long, g As Long, b As Long
            r = CLng(val(Trim$(parts(0))))
            g = CLng(val(Trim$(parts(1))))
            b = CLng(val(Trim$(parts(2))))
            ParseColor = RGB(BoundByte(r), BoundByte(g), BoundByte(b))
            Exit Function
        End If
    End If
    
    '=== Handle named colors ===
    Select Case LCase$(colorSpec)
        Case "red":    ParseColor = RGB(255, 0, 0)
        Case "green":  ParseColor = RGB(0, 128, 0)
        Case "blue":   ParseColor = RGB(0, 0, 255)
        Case "yellow": ParseColor = RGB(255, 255, 0)
        Case "cyan":   ParseColor = RGB(0, 255, 255)
        Case "magenta": ParseColor = RGB(255, 0, 255)
        Case "white":  ParseColor = RGB(255, 255, 255)
        Case "gray", "grey": ParseColor = RGB(128, 128, 128)
        Case "black":  ParseColor = RGB(0, 0, 0)
        Case Else
            ' Unknown name defaults to black
            ParseColor = RGB(0, 0, 0)
    End Select
End Function


Private Function HexToRGB(hexColor As String) As Long
    Dim r As Long, g As Long, b As Long
    If Len(hexColor) = 7 And Left$(hexColor, 1) = "#" Then
        r = CLng("&H" & Mid$(hexColor, 2, 2))
        g = CLng("&H" & Mid$(hexColor, 4, 2))
        b = CLng("&H" & Mid$(hexColor, 6, 2))
        HexToRGB = RGB(r, g, b)
    Else
        HexToRGB = RGB(0, 0, 0)
    End If
End Function

Private Function BoundByte(ByVal v As Long) As Byte
    If v < 0 Then v = 0
    If v > 255 Then v = 255
    BoundByte = CByte(v)
End Function

'================== UTIL: DASH MAPPING =================
Private Function DashToMsoLineStyle(dash As String) As MsoLineDashStyle
    Select Case LCase$(Trim$(dash))
        Case "solid":       DashToMsoLineStyle = msoLineSolid
        Case "dash":        DashToMsoLineStyle = msoLineDash
        Case "dot":         DashToMsoLineStyle = msoLineRoundDot
        Case "dashdot":     DashToMsoLineStyle = msoLineDashDot
        Case "longdash":    DashToMsoLineStyle = msoLineLongDash
        Case "longdashdot": DashToMsoLineStyle = msoLineLongDashDot
        Case "dashdotdot":  DashToMsoLineStyle = msoLineDashDotDot
        Case Else:          DashToMsoLineStyle = msoLineSolid
    End Select
End Function

'================== UTIL: MARKER MAPPING ===============
Private Function ShapeToMarkerStyle(shape As String) As XlMarkerStyle
    Select Case LCase$(Trim$(shape))
        Case "circle":         ShapeToMarkerStyle = xlMarkerStyleCircle
        Case "square":         ShapeToMarkerStyle = xlMarkerStyleSquare
        Case "diamond":        ShapeToMarkerStyle = xlMarkerStyleDiamond
        Case "cross":          ShapeToMarkerStyle = xlMarkerStyleX
        Case "x":              ShapeToMarkerStyle = xlMarkerStylePlus
        Case "triangle-up":    ShapeToMarkerStyle = xlMarkerStyleTriangle
        Case "triangle-down":  ShapeToMarkerStyle = xlMarkerStyleTriangle  ' Excel has no rotation here
        Case "star":           ShapeToMarkerStyle = xlMarkerStyleCircle    ' default to circle per spec
        Case "triangle":       ShapeToMarkerStyle = xlMarkerStyleTriangle
        Case Else:             ShapeToMarkerStyle = xlMarkerStyleCircle
    End Select
End Function

'========= UTIL: PARSE ARRAYS, HANDLE CATEGORIES =======
Private Sub BuildXYForCategories(ByVal csvX As String, ByVal csvY As String, _
    ByRef outX As Variant, ByRef outY As Variant)
    
    Dim rx As Variant, ry As Variant
    Dim i As Long, n As Long
    Dim tx() As Variant, ty() As Double
    
    rx = CSVToStringArray(csvX)
    ry = CSVToNumericArray(csvY)
    
    If IsEmpty(rx) Or IsEmpty(ry) Then
        outX = Empty: outY = Empty
        Exit Sub
    End If
    
    ReDim tx(0 To Application.Min(UBound(rx), UBound(ry)))
    ReDim ty(0 To UBound(tx))
    
    n = -1
    For i = 0 To UBound(tx)
        If IsNumeric(ry(i)) Then
            n = n + 1
            tx(n) = CStr(rx(i))
            ty(n) = CDbl(ry(i))
        End If
    Next i
    
    If n < 0 Then
        outX = Empty: outY = Empty
        Exit Sub
    End If
    
    ReDim Preserve tx(0 To n)
    ReDim Preserve ty(0 To n)
    outX = tx
    outY = ty
End Sub

Private Sub BuildXYNumericFlexible(ByVal csvX As String, ByVal csvY As String, _
    ByRef outX As Variant, ByRef outY As Variant)

    Dim rx As Variant, ry As Variant
    Dim i As Long, n As Long
    Dim tx() As Variant, ty() As Double
    Dim treatAsCategory As Boolean
    Dim xVal As Variant, yVal As Variant

    rx = CSVToStringArray(csvX)
    ry = CSVToStringArray(csvY)

    If IsEmpty(rx) Or IsEmpty(ry) Then
        outX = Empty: outY = Empty
        Exit Sub
    End If

    '=== Detect if any X value is not ISO date, standard date, or number ===
    treatAsCategory = False
    For i = LBound(rx) To UBound(rx)
        If Not (rx(i) Like "####-##-##" Or IsDate(rx(i)) Or IsNumeric(rx(i))) Then
            treatAsCategory = True
            Exit For
        End If
    Next i

    '=== Prepare arrays ===
    ReDim tx(0 To Application.Min(UBound(rx), UBound(ry)))
    ReDim ty(0 To UBound(tx))
    n = -1

    '=== Build arrays ===
    For i = 0 To UBound(tx)
        ' Validate both X and Y are non-empty and not NaN
        If Len(Trim$(rx(i))) > 0 And LCase$(Trim$(rx(i))) <> "nan" _
           And Len(Trim$(ry(i))) > 0 And LCase$(Trim$(ry(i))) <> "nan" Then

            n = n + 1

            ' ---------- X handling ----------
            If treatAsCategory Then
                tx(n) = CStr(rx(i))
            ElseIf rx(i) Like "####-##-##" Then
                'Convert ISO 8601 date YYYY-MM-DD to Excel serial
                tx(n) = CDbl(DateSerial(Left$(rx(i), 4), Mid$(rx(i), 6, 2), right$(rx(i), 2)))
            ElseIf IsDate(rx(i)) Then
                'Locale-recognized date
                tx(n) = CDbl(CDate(rx(i)))
            Else
                'Pure numeric
                tx(n) = CDbl(rx(i))
            End If

            ' ---------- Y handling ----------
            If ry(i) Like "####-##-##" Then
                'ISO date
                ty(n) = CDbl(DateSerial(Left$(ry(i), 4), Mid$(ry(i), 6, 2), right$(ry(i), 2)))
            ElseIf IsDate(ry(i)) Then
                'Locale date
                ty(n) = CDbl(CDate(ry(i)))
            ElseIf IsNumeric(ry(i)) Then
                'Numeric
                ty(n) = CDbl(ry(i))
            Else
                'Invalid Y: remove point
                n = n - 1
            End If
        End If
    Next i

    If n < 0 Then
        outX = Empty: outY = Empty
        Exit Sub
    End If

    ReDim Preserve tx(0 To n)
    ReDim Preserve ty(0 To n)
    outX = tx
    outY = ty
End Sub




Private Function CSVToStringArray(ByVal csv As String) As Variant
    Dim parts As Variant
    Dim i As Long
    csv = Trim$(csv)
    If Len(csv) = 0 Then
        CSVToStringArray = Empty
        Exit Function
    End If
    parts = Split(csv, ",")
    For i = LBound(parts) To UBound(parts)
        parts(i) = Trim$(CStr(parts(i)))
    Next i
    CSVToStringArray = parts
End Function

Private Function CSVToNumericArray(ByVal csv As String) As Variant
    Dim parts As Variant
    Dim i As Long, n As Long
    Dim outArr() As Double
    csv = Trim$(csv)
    If Len(csv) = 0 Then
        CSVToNumericArray = Empty
        Exit Function
    End If
    parts = Split(csv, ",")
    ReDim outArr(0 To UBound(parts))
    n = -1
    For i = LBound(parts) To UBound(parts)
        If IsNumeric(parts(i)) And Len(Trim$(parts(i))) > 0 And LCase$(Trim$(parts(i))) <> "nan" Then
            n = n + 1
            outArr(n) = CDbl(parts(i))
        End If
    Next i
    If n < 0 Then
        CSVToNumericArray = Empty
    Else
        ReDim Preserve outArr(0 To n)
        CSVToNumericArray = outArr
    End If
End Function

Private Function SafeUBound1(ByVal v As Variant) As Long
    If IsEmpty(v) Then
        SafeUBound1 = -1
    Else
        SafeUBound1 = UBound(v)
    End If
End Function

Private Function IsPositiveNumeric(ByVal s As String) As Boolean
    If Not IsNumeric(s) Then
        IsPositiveNumeric = False
    Else
        IsPositiveNumeric = (CDbl(s) > 0)
    End If
End Function

Private Function InferDateFormat(ByVal sample As String) As String
    sample = Trim$(sample)
    
    ' ISO 8601 like 2025-03-24
    If sample Like "####-##-##" Then
        InferDateFormat = "yyyy-mm-dd"
        Exit Function
    End If

    ' Standard with slashes like 03/24/2025
    If sample Like "##/##/####" Then
        Dim firstPart As Integer, secondPart As Integer
        firstPart = val(Split(sample, "/")(0))
        secondPart = val(Split(sample, "/")(1))

        If firstPart > 12 And secondPart <= 12 Then
            InferDateFormat = "dd/mm/yyyy" ' Day first
        Else
            InferDateFormat = "mm/dd/yyyy" ' Month first
        End If
        Exit Function
    End If

    ' Default fallback
    InferDateFormat = "yyyy-mm-dd"
End Function


'=============================================
' Render all annotation extras
'=============================================
Private Sub RenderExtrasAnnotations(ByVal doc As Object, ByVal chtObj As ChartObject)
    Dim annNode As Object
    Dim annType As String
    
    For Each annNode In doc.SelectNodes("//extras/annotation")
        annType = LCase$(annNode.Attributes.getNamedItem("type").text)
        
        Select Case annType
            Case "event_line"
                ' Vertical line
                Call DrawLineAnnotation(annNode, chtObj, True)
                
            Case "threshold"
                ' Horizontal line
                Call DrawLineAnnotation(annNode, chtObj, False)
                
            Case "arrow"
                ' Directional arrow
                Call DrawArrowAnnotation(annNode, chtObj)
                
            Case "text"
                ' Label or annotation text
                Call DrawTextAnnotation(annNode, chtObj)
                
            Case Else
                ' Ignore unsupported annotation types
                Debug.Print "Unsupported annotation type in extras:", annType
        End Select
    Next annNode
End Sub


'=============================================
' Robust parser for axis values (number, date, or string)
'=============================================
Private Function ParseAxisValue(ByVal rawVal As String) As Variant
    rawVal = Trim(rawVal)
    If Len(rawVal) = 0 Then
        ParseAxisValue = Empty
        Exit Function
    End If
    
    If IsNumeric(rawVal) Then
        ParseAxisValue = CDbl(rawVal)
        Exit Function
    End If
    
    If IsDate(rawVal) Then
        ParseAxisValue = CDbl(CDate(rawVal))  ' Excel serial date
        Exit Function
    End If
    
    ParseAxisValue = rawVal  ' fallback: string
End Function


'=============================================
' Apply style to any element that supports a .Format.Line
'=============================================
Private Sub ApplyElementStyle(ByVal target As Object, ByVal styleNode As Object)
    Dim dashType As String, opacityVal As Double
    
    If styleNode Is Nothing Then Exit Sub
    
    With target.Format.line
        .ForeColor.RGB = ParseColor(SafeNodeText(styleNode, "color", "#000000"))
        .Weight = val(SafeNodeText(styleNode, "width", "2"))
        
        dashType = LCase$(SafeNodeText(styleNode, "dash", "solid"))
        Select Case dashType
            Case "dot": .DashStyle = 3      ' dot
            Case "dash": .DashStyle = 2     ' dash
            Case "dashdot": .DashStyle = 4  ' dash-dot
            Case "dashdotdot": .DashStyle = 5 ' dash-dot-dot
            Case Else: .DashStyle = 1       ' solid
        End Select
        
        opacityVal = val(SafeNodeText(styleNode, "opacity", "1"))
        If opacityVal < 0 Or opacityVal > 1 Then opacityVal = 1
        .Transparency = 1 - opacityVal
    End With
End Sub



'=============================================
' General line annotation (vertical/horizontal)
'=============================================
Private Sub DrawLineAnnotation(ByVal annNode As Object, ByVal chtObj As ChartObject, ByVal vertical As Boolean)
    Dim rawVal As String, value As Variant
    Dim spanNode As Object, spanAxis As String, spanMode As String
    Dim x0 As Variant, x1 As Variant, y0 As Variant, y1 As Variant
    Dim s As Series, styleNode As Object
    
    rawVal = SafeNodeText(annNode, "value")
    value = ParseAxisValue(rawVal)
    If IsEmpty(value) Then Exit Sub
    
    spanAxis = ""
    spanMode = "full"
    Set spanNode = annNode.SelectSingleNode("span")
    If Not spanNode Is Nothing Then
        spanAxis = spanNode.Attributes.getNamedItem("axis").text
        spanMode = spanNode.Attributes.getNamedItem("mode").text
    End If
    
    If vertical Then
        x0 = value: x1 = value
        Call ResolveSpan(spanAxis, spanMode, "y", chtObj, y0, y1, annNode)
    Else
        y0 = value: y1 = value
        Call ResolveSpan(spanAxis, spanMode, "x", chtObj, x0, x1, annNode)
    End If
    
    Set s = chtObj.Chart.SeriesCollection.NewSeries
    s.XValues = Array(x0, x1)
    s.Values = Array(y0, y1)
    s.chartType = xlXYScatterLines
    s.MarkerStyle = xlMarkerStyleNone
    s.name = SafeNodeText(annNode, "label", "Line")
    
    Set styleNode = annNode.SelectSingleNode("style")
    If Not styleNode Is Nothing Then Call ApplyElementStyle(s, styleNode)
End Sub


'=============================================
' Resolve span along orthogonal axis
'=============================================
Private Sub ResolveSpan(ByVal spanAxis As String, ByVal spanMode As String, _
                        ByVal which As String, ByVal chtObj As ChartObject, _
                        ByRef out0 As Variant, ByRef out1 As Variant, ByVal annNode As Object)
    
    Dim minVal As Double, maxVal As Double
    If which = "y" Then
        minVal = chtObj.Chart.Axes(xlValue).MinimumScale
        maxVal = chtObj.Chart.Axes(xlValue).MaximumScale
    Else
        minVal = chtObj.Chart.Axes(xlCategory).MinimumScale
        maxVal = chtObj.Chart.Axes(xlCategory).MaximumScale
    End If
    
    Select Case LCase$(spanMode)
        Case "explicit"
            out0 = ParseAxisValue(SafeNodeText(annNode, which & "0"))
            out1 = ParseAxisValue(SafeNodeText(annNode, which & "1"))
        Case "domain"
            out0 = minVal
            out1 = maxVal
        Case Else
            out0 = minVal
            out1 = maxVal
    End Select
End Sub


'=============================================
' Arrow annotation
'=============================================
Private Sub DrawArrowAnnotation(ByVal annNode As Object, ByVal chtObj As ChartObject)
    Dim x0 As Variant, x1 As Variant, y0 As Variant, y1 As Variant
    Dim shp As shape, styleNode As Object
    
    x0 = ParseAxisValue(SafeNodeText(annNode, "x0"))
    x1 = ParseAxisValue(SafeNodeText(annNode, "x1"))
    y0 = ParseAxisValue(SafeNodeText(annNode, "y0"))
    y1 = ParseAxisValue(SafeNodeText(annNode, "y1"))
    
    Set shp = chtObj.parent.Shapes.AddLine(chtObj.Left + 20, chtObj.Top + 20, chtObj.Left + 120, chtObj.Top + 120)
    shp.line.EndArrowheadStyle = msoArrowheadTriangle
    shp.name = SafeNodeText(annNode, "label", "Arrow")
    
    Set styleNode = annNode.SelectSingleNode("style")
    If Not styleNode Is Nothing Then Call ApplyElementStyle(shp, styleNode)
End Sub


'=============================================
' Text annotation
'=============================================
Private Sub DrawTextAnnotation(ByVal annNode As Object, ByVal chtObj As ChartObject)
    Dim txt As String, shp As shape, styleNode As Object
    
    txt = SafeNodeText(annNode, "label", "")
    If Len(txt) = 0 Then Exit Sub
    
    Set shp = chtObj.parent.Shapes.AddTextbox(msoTextOrientationHorizontal, chtObj.Left + 50, chtObj.Top + 50, 100, 20)
    shp.TextFrame2.TextRange.text = txt
    shp.name = "TextAnn"
    
    Set styleNode = annNode.SelectSingleNode("style")
    If Not styleNode Is Nothing Then Call ApplyElementStyle(shp, styleNode)
End Sub




