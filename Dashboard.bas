Option Explicit

' ============================================================
'  Dashboard Builder - Lead Analytics Dashboard
'  Quelle: Kundenliste (Pipeline-Sheet)
'  Design: Rounded cards, shadow, blue accent
' ============================================================

Private Sub FormatCard(shp As Shape)
    With shp
        .Fill.ForeColor.RGB = RGB(255, 255, 255)
        .Line.Visible = msoFalse
        On Error Resume Next
        .Shadow.Visible = msoTrue
        .Shadow.ForeColor.RGB = RGB(148, 163, 184)
        .Shadow.Transparency = 0.6
        .Shadow.OffsetX = 2
        .Shadow.OffsetY = 3
        .Shadow.Blur = 8
        .Adjustments(1) = 0.06
        On Error GoTo 0
    End With
End Sub

Private Sub AddLabel(ws As Worksheet, x As Double, y As Double, _
    w As Double, h As Double, txt As String, fontSize As Single, _
    isBold As Boolean, clr As Long)
    Dim shp As Shape
    Set shp = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, x, y, w, h)
    With shp
        .TextFrame2.TextRange.Text = txt
        .TextFrame2.TextRange.Font.Size = fontSize
        .TextFrame2.TextRange.Font.Name = "Avenir Next"
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = clr
        If isBold Then .TextFrame2.TextRange.Font.Bold = msoTrue
        .Fill.Visible = msoFalse
        .Line.Visible = msoFalse
    End With
End Sub

Public Sub BuildDashboard()
    ' ===== DECLARATIONS =====
    Dim ws As Worksheet, tbl As ListObject
    Dim shp As Shape, s As Shape, chartObj As ChartObject
    Dim i As Long, j As Long, r As Long, c As Long
    Dim tmpL As Long, tmpS As String
    Dim found As Boolean, idx As Long
    Dim dataArr As Variant, headers As Variant
    Dim nRows As Long, hCols As Long
    Dim cMonat As Long, cAbschluss As Long, cGrund As Long
    Dim cAbgNach As Long, cStatus As Long
    Dim dt As Date
    Dim rowDate As Variant, ym As Long, mLabel As String
    Dim abschluss As String, grund As String, abgNach As String
    Dim abschlussRate As Double
    Dim totalReasons As Long, totalAbg As Long
    Dim peakIdx As Long, worstIdx As Long, worstRate As Double, rate As Double
    Dim xPos As Double, yPos As Double
    Dim tblTop As Double, ry As Double, iy As Double
    Dim maxR As Long, maxA As Long, maxM As Long, mi2 As Long

    ' Data arrays
    Dim mLabels(1 To 60) As String, mKeys(1 To 60) As Long
    Dim mLeads(1 To 60) As Long, mClosed(1 To 60) As Long, mDropped(1 To 60) As Long
    Dim mCount As Long
    Dim rLabels(1 To 50) As String, rCounts(1 To 50) As Long, rCount As Long
    Dim aLabels(1 To 50) As String, aCounts(1 To 50) As Long, aCount As Long
    Dim totalLeads As Long, totalClosed As Long
    Dim totalDropped As Long, totalLaufend As Long
    Dim curYM As Long, curMLeads As Long, curMClosed As Long, curMDropped As Long

    ' Status Pipeline
    Dim stLabels(1 To 6) As String, stCounts(1 To 6) As Long, stTotal As Long
    Dim statusVal As String, colX As Double

    ' Layout (points)
    Dim LM As Double: LM = 20
    Dim CW As Double: CW = 195
    Dim CH As Double: CH = 85
    Dim CG As Double: CG = 15
    Dim CHW As Double: CHW = 410
    Dim CHH As Double: CHH = 230
    Dim SG As Double: SG = 20
    Dim TH As Double: TH = 230
    Dim rowH As Double: rowH = 18
    Dim lineH As Double: lineH = 36
    Dim dataCol As Long: dataCol = 20

    ' ===== SETUP =====
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Dashboard")
    Set tbl = ThisWorkbook.Worksheets("Pipeline").ListObjects("Kundenliste")
    On Error GoTo 0

    If ws Is Nothing Then
        MsgBox "Blatt 'Dashboard' nicht gefunden.", vbExclamation: Exit Sub
    End If
    If tbl Is Nothing Then
        MsgBox "Tabelle 'Kundenliste' nicht gefunden.", vbExclamation: Exit Sub
    End If

    Application.ScreenUpdating = False

    ' ===== CLEAR =====
    ws.Cells.Clear
    Do While ws.Shapes.Count > 0: ws.Shapes(1).Delete: Loop
    Do While ws.ChartObjects.Count > 0: ws.ChartObjects(1).Delete: Loop

    ' Background & fonts
    ws.Cells.Interior.Color = RGB(245, 248, 252)
    ws.Cells.Font.Name = "Avenir Next"
    ws.Cells.Font.Size = 10
    ws.Cells.Font.Color = RGB(15, 23, 42)

    ' ===== DATA PROCESSING =====
    dataArr = tbl.DataBodyRange.Value
    nRows = UBound(dataArr, 1)
    headers = tbl.HeaderRowRange.Value
    hCols = UBound(headers, 2)

    For c = 1 To hCols
        Select Case Trim(CStr(headers(1, c)))
            Case "Monat Lead erhalten": cMonat = c
            Case "Abschluss": cAbschluss = c
            Case "Grund zum Absprung": cGrund = c
            Case "Abgesprungen nach": cAbgNach = c
            Case "Status": cStatus = c
        End Select
    Next c

    curYM = Year(Date) * 100 + Month(Date)

    ' Status Pipeline Labels initialisieren
    stLabels(1) = "Lead erhalten"
    stLabels(2) = "Nicht erreicht"
    stLabels(3) = "Standby nach Anruf"
    stLabels(4) = "Ersttermin"
    stLabels(5) = "Standby nach Ersttermin"
    stLabels(6) = "Geschlossen"

    For r = 1 To nRows
        rowDate = dataArr(r, cMonat)
        If IsDate(rowDate) Then
            dt = CDate(rowDate)
        ElseIf IsNumeric(rowDate) Then
            If CDbl(rowDate) > 30000 Then
                dt = CDate(CDbl(rowDate))
            Else
                GoTo NextRow
            End If
        Else
            GoTo NextRow
        End If

        ym = Year(dt) * 100 + Month(dt)
        mLabel = Format(dt, "MMM YY")
        abschluss = LCase(Trim(CStr(dataArr(r, cAbschluss) & "")))
        grund = Trim(CStr(dataArr(r, cGrund) & ""))
        abgNach = Trim(CStr(dataArr(r, cAbgNach) & ""))

        ' Month lookup
        found = False
        For idx = 1 To mCount
            If mKeys(idx) = ym Then found = True: Exit For
        Next idx
        If Not found Then
            mCount = mCount + 1: idx = mCount
            mKeys(idx) = ym: mLabels(idx) = mLabel
        End If
        mLeads(idx) = mLeads(idx) + 1
        If abschluss = "ja" Then mClosed(idx) = mClosed(idx) + 1
        If abschluss = "nein" Then mDropped(idx) = mDropped(idx) + 1

        totalLeads = totalLeads + 1
        If abschluss = "ja" Then totalClosed = totalClosed + 1
        If abschluss = "nein" Then totalDropped = totalDropped + 1
        If abschluss = "laufend" Or abschluss = "" Then totalLaufend = totalLaufend + 1

        If ym = curYM Then
            curMLeads = curMLeads + 1
            If abschluss = "ja" Then curMClosed = curMClosed + 1
            If abschluss = "nein" Then curMDropped = curMDropped + 1
        End If

        If Len(grund) > 0 Then
            found = False
            For idx = 1 To rCount
                If rLabels(idx) = grund Then
                    found = True: rCounts(idx) = rCounts(idx) + 1: Exit For
                End If
            Next idx
            If Not found Then
                rCount = rCount + 1
                rLabels(rCount) = grund: rCounts(rCount) = 1
            End If
        End If

        If Len(abgNach) > 0 Then
            found = False
            For idx = 1 To aCount
                If aLabels(idx) = abgNach Then
                    found = True: aCounts(idx) = aCounts(idx) + 1: Exit For
                End If
            Next idx
            If Not found Then
                aCount = aCount + 1
                aLabels(aCount) = abgNach: aCounts(aCount) = 1
            End If
        End If

        ' Status Pipeline zaehlen
        If cStatus > 0 Then
            statusVal = Trim(CStr(dataArr(r, cStatus) & ""))
            For idx = 1 To 6
                If statusVal = stLabels(idx) Then
                    stCounts(idx) = stCounts(idx) + 1
                    Exit For
                End If
            Next idx
        End If
NextRow:
    Next r

    ' Sort months ascending
    For i = 1 To mCount - 1
        For j = i + 1 To mCount
            If mKeys(i) > mKeys(j) Then
                tmpL = mKeys(i): mKeys(i) = mKeys(j): mKeys(j) = tmpL
                tmpS = mLabels(i): mLabels(i) = mLabels(j): mLabels(j) = tmpS
                tmpL = mLeads(i): mLeads(i) = mLeads(j): mLeads(j) = tmpL
                tmpL = mClosed(i): mClosed(i) = mClosed(j): mClosed(j) = tmpL
                tmpL = mDropped(i): mDropped(i) = mDropped(j): mDropped(j) = tmpL
            End If
        Next j
    Next i

    ' Sort reasons desc
    For i = 1 To rCount - 1
        For j = i + 1 To rCount
            If rCounts(i) < rCounts(j) Then
                tmpL = rCounts(i): rCounts(i) = rCounts(j): rCounts(j) = tmpL
                tmpS = rLabels(i): rLabels(i) = rLabels(j): rLabels(j) = tmpS
            End If
        Next j
    Next i

    ' Sort abgesprungen nach desc
    For i = 1 To aCount - 1
        For j = i + 1 To aCount
            If aCounts(i) < aCounts(j) Then
                tmpL = aCounts(i): aCounts(i) = aCounts(j): aCounts(j) = tmpL
                tmpS = aLabels(i): aLabels(i) = aLabels(j): aLabels(j) = tmpS
            End If
        Next j
    Next i

    ' Rates
    If totalLeads > 0 Then abschlussRate = totalClosed / totalLeads

    ' ===== TITLE =====
    yPos = 15
    AddLabel ws, LM, yPos, 300, 30, "Dashboard", 20, True, RGB(25, 55, 95)
    yPos = yPos + 42

    ' ===== KPI CARDS =====
    ' Card 1: Gesamt Leads
    xPos = LM
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, xPos, yPos, CW, CH)
    FormatCard shp
    AddLabel ws, xPos + 20, yPos + 12, CW - 30, 16, "Gesamt Leads", 10, False, RGB(100, 120, 150)
    AddLabel ws, xPos + 20, yPos + 32, CW - 30, 40, CStr(totalLeads), 28, True, RGB(25, 55, 95)
    Set shp = ws.Shapes.AddShape(msoShapeRectangle, _
        xPos + 4, yPos + 10, 4, CH - 20)
    shp.Fill.ForeColor.RGB = RGB(50, 110, 165)
    shp.Line.Visible = msoFalse

    ' Card 2: Abschlussrate
    xPos = xPos + CW + CG
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, xPos, yPos, CW, CH)
    FormatCard shp
    AddLabel ws, xPos + 20, yPos + 12, CW - 30, 16, "Abschlussrate", 10, False, RGB(100, 120, 150)
    AddLabel ws, xPos + 20, yPos + 32, CW - 30, 40, Format(abschlussRate, "0.0%"), 28, True, RGB(50, 110, 165)
    Set shp = ws.Shapes.AddShape(msoShapeRectangle, _
        xPos + 4, yPos + 10, 4, CH - 20)
    shp.Fill.ForeColor.RGB = RGB(50, 110, 165)
    shp.Line.Visible = msoFalse

    ' Card 3: Abspruenge
    xPos = xPos + CW + CG
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, xPos, yPos, CW, CH)
    FormatCard shp
    AddLabel ws, xPos + 20, yPos + 12, CW - 30, 16, "Abspruenge", 10, False, RGB(100, 120, 150)
    AddLabel ws, xPos + 20, yPos + 32, CW - 30, 40, CStr(totalDropped), 28, True, RGB(140, 160, 185)
    Set shp = ws.Shapes.AddShape(msoShapeRectangle, _
        xPos + 4, yPos + 10, 4, CH - 20)
    shp.Fill.ForeColor.RGB = RGB(140, 160, 185)
    shp.Line.Visible = msoFalse

    ' Card 4: Laufend
    xPos = xPos + CW + CG
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, xPos, yPos, CW, CH)
    FormatCard shp
    AddLabel ws, xPos + 20, yPos + 12, CW - 30, 16, "Laufend", 10, False, RGB(100, 120, 150)
    AddLabel ws, xPos + 20, yPos + 32, CW - 30, 40, CStr(totalLaufend), 28, True, RGB(95, 145, 190)
    Set shp = ws.Shapes.AddShape(msoShapeRectangle, _
        xPos + 4, yPos + 10, 4, CH - 20)
    shp.Fill.ForeColor.RGB = RGB(95, 145, 190)
    shp.Line.Visible = msoFalse

    ' ===== CHART DATA (hidden columns T-W) =====
    ws.Cells(1, dataCol).Value = "Monat"
    ws.Cells(1, dataCol + 1).Value = "Leads"
    ws.Cells(1, dataCol + 2).Value = "Abgeschlossen"
    ws.Cells(1, dataCol + 3).Value = "Abgesprungen"
    For i = 1 To mCount
        ws.Cells(1 + i, dataCol).Value = mLabels(i)
        ws.Cells(1 + i, dataCol + 1).Value = mLeads(i)
        ws.Cells(1 + i, dataCol + 2).Value = mClosed(i)
        ws.Cells(1 + i, dataCol + 3).Value = mDropped(i)
    Next i

    ' Chart 2 benoetigt eigene Spalten (Monat + Abgesprungen)
    Dim dataCol2 As Long: dataCol2 = dataCol + 4
    ws.Cells(1, dataCol2).Value = "Monat"
    ws.Cells(1, dataCol2 + 1).Value = "Abgesprungen"
    For i = 1 To mCount
        ws.Cells(1 + i, dataCol2).Value = mLabels(i)
        ws.Cells(1 + i, dataCol2 + 1).Value = mDropped(i)
    Next i

    ' ===== CHARTS SECTION =====
    yPos = yPos + CH + SG

    ' -- Chart 1: Leads & Abschluss Trend (SetSourceData) --
    xPos = LM
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, _
        xPos, yPos, CHW, CHH)
    FormatCard shp
    AddLabel ws, xPos + 15, yPos + 10, 260, 20, _
        "Leads & Abschluss Trend", 13, True, RGB(25, 55, 95)

    If mCount > 0 Then
        Dim rngChart1 As Range
        Set rngChart1 = ws.Range(ws.Cells(1, dataCol), ws.Cells(1 + mCount, dataCol + 2))
        Set chartObj = ws.ChartObjects.Add( _
            xPos + 10, yPos + 35, CHW - 20, CHH - 50)
        With chartObj.Chart
            .SetSourceData Source:=rngChart1, PlotBy:=xlColumns
            .ChartType = xlLineMarkers
            .HasTitle = False
            .HasLegend = True
            .Legend.Position = xlLegendPositionBottom
            On Error Resume Next
            .Legend.Font.Size = 8
            .Legend.Font.Name = "Avenir Next"
            If .SeriesCollection.Count >= 1 Then
                With .SeriesCollection(1)
                    .Format.Line.ForeColor.RGB = RGB(50, 110, 165)
                    .Format.Line.Weight = 2.5
                    .MarkerStyle = xlMarkerStyleCircle
                    .MarkerSize = 7
                    .MarkerForegroundColor = RGB(50, 110, 165)
                    .MarkerBackgroundColor = RGB(255, 255, 255)
                End With
            End If
            If .SeriesCollection.Count >= 2 Then
                With .SeriesCollection(2)
                    .Format.Line.ForeColor.RGB = RGB(95, 145, 190)
                    .Format.Line.Weight = 2.5
                    .MarkerStyle = xlMarkerStyleCircle
                    .MarkerSize = 7
                    .MarkerForegroundColor = RGB(95, 145, 190)
                    .MarkerBackgroundColor = RGB(255, 255, 255)
                End With
            End If
            .PlotArea.Format.Fill.Visible = msoFalse
            .ChartArea.Format.Fill.Visible = msoFalse
            .ChartArea.Format.Line.Visible = msoFalse
            On Error GoTo 0
        End With
    End If

    ' -- Chart 2: Absprung Trend (SetSourceData) --
    xPos = LM + CHW + CG
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, _
        xPos, yPos, CHW, CHH)
    FormatCard shp
    AddLabel ws, xPos + 15, yPos + 10, 260, 20, _
        "Absprung Trend", 13, True, RGB(25, 55, 95)

    If mCount > 0 Then
        Dim rngChart2 As Range
        Set rngChart2 = ws.Range(ws.Cells(1, dataCol2), ws.Cells(1 + mCount, dataCol2 + 1))
        Set chartObj = ws.ChartObjects.Add( _
            xPos + 10, yPos + 35, CHW - 20, CHH - 50)
        With chartObj.Chart
            .SetSourceData Source:=rngChart2, PlotBy:=xlColumns
            .ChartType = xlColumnClustered
            .HasTitle = False
            .HasLegend = False
            On Error Resume Next
            If .SeriesCollection.Count >= 1 Then
                .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(140, 160, 185)
            End If
            .PlotArea.Format.Fill.Visible = msoFalse
            .ChartArea.Format.Fill.Visible = msoFalse
            .ChartArea.Format.Line.Visible = msoFalse
            On Error GoTo 0
        End With
    End If

    ' ===== BOTTOM ROW 1: Absprunggruende + Insights =====
    yPos = yPos + CHH + SG

    ' -- Absprunggruende Card --
    xPos = LM
    maxR = rCount: If maxR > 9 Then maxR = 9
    Dim abgCardH As Double
    abgCardH = 50 + maxR * (rowH + 2) + 10
    If abgCardH < TH Then abgCardH = TH
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, _
        xPos, yPos, CHW, abgCardH)
    FormatCard shp
    AddLabel ws, xPos + 15, yPos + 10, 260, 20, _
        "Absprunggruende", 13, True, RGB(25, 55, 95)

    tblTop = yPos + 38
    AddLabel ws, xPos + 15, tblTop, 250, rowH, _
        "Grund", 9, True, RGB(100, 120, 150)
    AddLabel ws, xPos + 280, tblTop, 50, rowH, _
        "Anz.", 9, True, RGB(100, 120, 150)
    AddLabel ws, xPos + 340, tblTop, 60, rowH, _
        "Anteil", 9, True, RGB(100, 120, 150)

    Set shp = ws.Shapes.AddShape(msoShapeRectangle, _
        xPos + 15, tblTop + rowH, CHW - 30, 1)
    shp.Fill.ForeColor.RGB = RGB(215, 225, 235)
    shp.Line.Visible = msoFalse

    For i = 1 To rCount: totalReasons = totalReasons + rCounts(i): Next i
    For i = 1 To maxR
        ry = tblTop + rowH + 4 + (i - 1) * (rowH + 2)
        AddLabel ws, xPos + 15, ry, 260, rowH, _
            rLabels(i), 9, False, RGB(65, 85, 110)
        AddLabel ws, xPos + 280, ry, 50, rowH, _
            CStr(rCounts(i)), 9, False, RGB(65, 85, 110)
        If totalReasons > 0 Then
            AddLabel ws, xPos + 340, ry, 60, rowH, _
                Format(rCounts(i) / totalReasons, "0.0%"), _
                9, False, RGB(65, 85, 110)
        End If
    Next i

    ' -- Insights Card --
    xPos = LM + CHW + CG
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, _
        xPos, yPos, CHW, abgCardH)
    FormatCard shp
    AddLabel ws, xPos + 15, yPos + 10, 300, 20, _
        "Insights & Empfehlungen", 13, True, RGB(25, 55, 95)

    iy = yPos + 42

    ' Insight 1: Top Absprunggrund
    If rCount > 0 Then
        AddLabel ws, xPos + 15, iy, 380, 14, _
            "Top Absprunggrund", 9, False, RGB(100, 120, 150)
        AddLabel ws, xPos + 15, iy + 15, 380, 16, _
            rLabels(1) & " (" & rCounts(1) & "x)", _
            11, True, RGB(140, 160, 185)
        iy = iy + lineH + 6
    End If

    ' Insight 2: Peak Leads
    peakIdx = 1
    For i = 2 To mCount
        If mLeads(i) > mLeads(peakIdx) Then peakIdx = i
    Next i
    If mCount > 0 Then
        AddLabel ws, xPos + 15, iy, 380, 14, _
            "Peak Leads Monat", 9, False, RGB(100, 120, 150)
        AddLabel ws, xPos + 15, iy + 15, 380, 16, _
            mLabels(peakIdx) & " (" & mLeads(peakIdx) & " Leads)", _
            11, True, RGB(50, 110, 165)
        iy = iy + lineH + 6
    End If

    ' Insight 3: Schwaechste Abschlussrate
    worstIdx = 1: worstRate = 999
    For i = 1 To mCount
        If mLeads(i) > 0 Then rate = mClosed(i) / mLeads(i) Else rate = 0
        If rate < worstRate Then worstRate = rate: worstIdx = i
    Next i
    If mCount > 0 Then
        AddLabel ws, xPos + 15, iy, 380, 14, _
            "Schwaechste Abschlussrate", 9, False, RGB(100, 120, 150)
        AddLabel ws, xPos + 15, iy + 15, 380, 16, _
            mLabels(worstIdx) & " (" & Format(worstRate, "0.0%") & ")", _
            11, True, RGB(95, 145, 190)
        iy = iy + lineH + 6
    End If

    ' Insight 4: Haeufigster Absprungzeitpunkt
    If aCount > 0 Then
        AddLabel ws, xPos + 15, iy, 380, 14, _
            "Haeufigster Absprungzeitpunkt", 9, False, RGB(100, 120, 150)
        AddLabel ws, xPos + 15, iy + 15, 380, 16, _
            aLabels(1) & " (" & aCounts(1) & "x)", _
            11, True, RGB(140, 160, 185)
    End If

    ' ===== BOTTOM ROW 2: Abgesprungen nach + Monatsuebersicht =====
    yPos = yPos + abgCardH + SG

    ' -- Abgesprungen nach Zeitpunkt --
    xPos = LM
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, _
        xPos, yPos, CHW, TH)
    FormatCard shp
    AddLabel ws, xPos + 15, yPos + 10, 300, 20, _
        "Abgesprungen nach Zeitpunkt", 13, True, RGB(25, 55, 95)

    tblTop = yPos + 38
    AddLabel ws, xPos + 15, tblTop, 250, rowH, _
        "Zeitpunkt", 9, True, RGB(100, 120, 150)
    AddLabel ws, xPos + 280, tblTop, 50, rowH, _
        "Anz.", 9, True, RGB(100, 120, 150)
    AddLabel ws, xPos + 340, tblTop, 60, rowH, _
        "Anteil", 9, True, RGB(100, 120, 150)

    Set shp = ws.Shapes.AddShape(msoShapeRectangle, _
        xPos + 15, tblTop + rowH, CHW - 30, 1)
    shp.Fill.ForeColor.RGB = RGB(215, 225, 235)
    shp.Line.Visible = msoFalse

    For i = 1 To aCount: totalAbg = totalAbg + aCounts(i): Next i
    maxA = aCount: If maxA > 9 Then maxA = 9
    For i = 1 To maxA
        ry = tblTop + rowH + 4 + (i - 1) * (rowH + 2)
        AddLabel ws, xPos + 15, ry, 260, rowH, _
            aLabels(i), 9, False, RGB(65, 85, 110)
        AddLabel ws, xPos + 280, ry, 50, rowH, _
            CStr(aCounts(i)), 9, False, RGB(65, 85, 110)
        If totalAbg > 0 Then
            AddLabel ws, xPos + 340, ry, 60, rowH, _
                Format(aCounts(i) / totalAbg, "0.0%"), _
                9, False, RGB(65, 85, 110)
        End If
    Next i

    ' -- Monatsuebersicht --
    xPos = LM + CHW + CG
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, _
        xPos, yPos, CHW, TH)
    FormatCard shp
    AddLabel ws, xPos + 15, yPos + 10, 300, 20, _
        "Monatsuebersicht", 13, True, RGB(25, 55, 95)

    tblTop = yPos + 38
    AddLabel ws, xPos + 15, tblTop, 80, rowH, _
        "Monat", 9, True, RGB(100, 120, 150)
    AddLabel ws, xPos + 110, tblTop, 60, rowH, _
        "Leads", 9, True, RGB(100, 120, 150)
    AddLabel ws, xPos + 190, tblTop, 80, rowH, _
        "Abgeschl.", 9, True, RGB(100, 120, 150)
    AddLabel ws, xPos + 280, tblTop, 80, rowH, _
        "Abgespr.", 9, True, RGB(100, 120, 150)
    AddLabel ws, xPos + 350, tblTop, 60, rowH, _
        "Rate", 9, True, RGB(100, 120, 150)

    Set shp = ws.Shapes.AddShape(msoShapeRectangle, _
        xPos + 15, tblTop + rowH, CHW - 30, 1)
    shp.Fill.ForeColor.RGB = RGB(215, 225, 235)
    shp.Line.Visible = msoFalse

    maxM = mCount: If maxM > 9 Then maxM = 9
    For i = 1 To maxM
        mi2 = mCount - maxM + i
        ry = tblTop + rowH + 4 + (i - 1) * (rowH + 2)
        AddLabel ws, xPos + 15, ry, 80, rowH, _
            mLabels(mi2), 9, False, RGB(65, 85, 110)
        AddLabel ws, xPos + 110, ry, 60, rowH, _
            CStr(mLeads(mi2)), 9, True, RGB(25, 55, 95)
        AddLabel ws, xPos + 190, ry, 80, rowH, _
            CStr(mClosed(mi2)), 9, False, RGB(50, 110, 165)
        AddLabel ws, xPos + 280, ry, 80, rowH, _
            CStr(mDropped(mi2)), 9, False, RGB(140, 160, 185)
        If mLeads(mi2) > 0 Then
            AddLabel ws, xPos + 350, ry, 60, rowH, _
                Format(mClosed(mi2) / mLeads(mi2), "0.0%"), _
                9, False, RGB(65, 85, 110)
        End If
    Next i

    ' ===== STATUS PIPELINE WIDGET =====
    yPos = yPos + TH + SG
    xPos = LM
    Dim spW As Double: spW = CHW * 2 + CG
    Dim spH As Double: spH = 105
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, _
        xPos, yPos, spW, spH)
    FormatCard shp
    AddLabel ws, xPos + 15, yPos + 10, 300, 20, _
        "Status Pipeline", 13, True, RGB(25, 55, 95)

    stTotal = 0
    For i = 1 To 6: stTotal = stTotal + stCounts(i): Next i

    Dim spColW As Double: spColW = (spW - 105) / 7
    Dim spStartX As Double: spStartX = xPos + 105

    ' Zeilen-Labels
    AddLabel ws, xPos + 15, yPos + 35, 85, rowH, _
        "Phase", 8, True, RGB(25, 55, 95)
    AddLabel ws, xPos + 15, yPos + 57, 85, rowH, _
        "Anzahl", 9, True, RGB(100, 120, 150)
    AddLabel ws, xPos + 15, yPos + 75, 85, rowH, _
        "%", 9, True, RGB(100, 120, 150)

    ' Separator
    Set shp = ws.Shapes.AddShape(msoShapeRectangle, _
        spStartX - 2, yPos + 54, spColW * 7 + 4, 1)
    shp.Fill.ForeColor.RGB = RGB(215, 225, 235)
    shp.Line.Visible = msoFalse

    ' Status-Spalten
    For i = 1 To 6
        colX = spStartX + (i - 1) * spColW
        AddLabel ws, colX + 3, yPos + 35, spColW - 6, rowH, _
            stLabels(i), 8, True, RGB(50, 110, 165)
        AddLabel ws, colX + 3, yPos + 57, spColW - 6, rowH, _
            CStr(stCounts(i)), 9, False, RGB(65, 85, 110)
        If stTotal > 0 Then
            AddLabel ws, colX + 3, yPos + 75, spColW - 6, rowH, _
                Format(stCounts(i) / stTotal, "0.0%"), 9, False, RGB(65, 85, 110)
        End If
    Next i

    ' Gesamt-Spalte
    colX = spStartX + 6 * spColW
    AddLabel ws, colX + 3, yPos + 35, spColW - 6, rowH, _
        "Gesamt", 8, True, RGB(50, 110, 165)
    AddLabel ws, colX + 3, yPos + 57, spColW - 6, rowH, _
        CStr(stTotal), 9, True, RGB(25, 55, 95)

    Application.ScreenUpdating = True
End Sub
