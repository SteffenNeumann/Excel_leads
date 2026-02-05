Option Explicit

Public Sub BuildDashboard()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim rngAll As Range

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Dashboard")
    Set tbl = ThisWorkbook.Worksheets("Pipeline").ListObjects("Kundenliste")
    On Error GoTo 0

    If ws Is Nothing Then
        MsgBox "Blatt 'Dashboard' nicht gefunden.", vbExclamation
        Exit Sub
    End If

    If tbl Is Nothing Then
        MsgBox "Tabelle 'Kundenliste' nicht gefunden.", vbExclamation
        Exit Sub
    End If

    Set rngAll = ws.Range("A1:I60")
    rngAll.Clear

    With rngAll.Font
        .Name = "Avenir Next"
        .Size = 11
    End With

    ws.Range("A1").Value = "Dashboard"
    With ws.Range("A1")
        .Font.Size = 16
        .Font.Bold = True
    End With

    ws.Range("A3").Value = "Abschlussrate"
    ws.Range("A11").Value = "Status Pipeline"
    ws.Range("A18").Value = "Kosten"
    ws.Range("A25").Value = "Abgesprungen nach"
    ws.Range("A33").Value = "Lead Orte"

    ' Abschlussrate
    ws.Range("B4:I4").Value = Array("Verbund", "Pflegehelfer24", "Empfehlung", "Sonstiges", "Pflegedienst", "Pflegestützpunkt", "Gesamt", "Anzahl")
    ws.Range("A5").Value = "Ja"
    ws.Range("A6").Value = "Laufend"
    ws.Range("A7").Value = "Nein"

    ws.Range("B5").Formula = "=IFERROR(COUNTIFS(Kundenliste[Abschluss],$A5,Kundenliste[Lead-Quelle],B$4)/COUNTIF(Kundenliste[Lead-Quelle],B$4),0)"
    ws.Range("B5:G5").FillRight
    ws.Range("B5:G5").AutoFill Destination:=ws.Range("B5:G7")

    ws.Range("H5").Formula = "=IFERROR(COUNTIF(Kundenliste[Abschluss],$A5)/COUNTA(Kundenliste[Abschluss]),0)"
    ws.Range("H5").AutoFill Destination:=ws.Range("H5:H7")

    ws.Range("I5").Formula = "=COUNTIF(Kundenliste[Abschluss],$A5)"
    ws.Range("I5").AutoFill Destination:=ws.Range("I5:I7")

    ws.Range("B5:H7").NumberFormat = "0.0%"
    ws.Range("I5:I7").NumberFormat = "0"

    ' Status Pipeline
    ws.Range("B12").Formula = "=TRANSPOSE(SORT(UNIQUE(FILTER(Kundenliste[Status],Kundenliste[Status]<>\"\"))))"
    ws.Range("B13").Formula = "=COUNTIF(Kundenliste[Status],B$12)"
    ws.Range("B14").Formula = "=IFERROR(B13/SUM(B13#),0)"
    ws.Range("B14").NumberFormat = "0.0%"

    ' Kosten
    ws.Range("B19").Formula = "=TRANSPOSE(SORT(UNIQUE(Kundenliste[Monat Lead erhalten])))"
    ws.Range("A20").Value = "Verbund"
    ws.Range("A21").Value = "Pflegehelfer24"
    ws.Range("A22").Value = "Empfehlung"
    ws.Range("A23").Value = "Sonstiges"

    ws.Range("B20").Formula = "=SUMIFS(Kundenliste[Spend],Kundenliste[Lead-Quelle],$A20,Kundenliste[Monat Lead erhalten],B$19)"
    ws.Range("B20").AutoFill Destination:=ws.Range("B20:H23")
    ws.Range("B20:H23").NumberFormat = "# ##0 €"

    ' Abgesprungen nach
    ws.Range("A26").Formula = "=SORT(UNIQUE(FILTER(Kundenliste[Abgesprungen nach],Kundenliste[Abgesprungen nach]<>\"\")))"
    ws.Range("B26").Formula = "=COUNTIF(Kundenliste[Abgesprungen nach],A26)"
    ws.Range("C26").Formula = "=IFERROR(B26/SUM(B26#),0)"
    ws.Range("C26").NumberFormat = "0.0%"

    ' Lead Orte
    ws.Range("A34").Formula = "=SORT(UNIQUE(FILTER(Kundenliste[Ort],Kundenliste[Ort]<>\"\")))"
    ws.Range("B34").Formula = "=XLOOKUP(A34,Kundenliste[Ort],Kundenliste[PLZ],\"\")"
    ws.Range("C34").Formula = "=COUNTIF(Kundenliste[Ort],A34)"

    ' Visual accents
    ws.Range("B4:I4").Font.Bold = True
    ws.Range("A3").Font.Bold = True
    ws.Range("A11").Font.Bold = True
    ws.Range("A18").Font.Bold = True
    ws.Range("A25").Font.Bold = True
    ws.Range("A33").Font.Bold = True

    ws.Range("B4:I4").Interior.Color = RGB(31, 41, 55)
    ws.Range("B4:I4").Font.Color = RGB(255, 255, 255)
    ws.Columns("A:I").AutoFit
End Sub
