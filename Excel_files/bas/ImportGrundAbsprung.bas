Attribute VB_Name = "ImportGrundAbsprung"
' ============================================================
' ImportGrundAbsprung
' -------------------
' Liest grund_absprung_mapping.xlsx (im gleichen Ordner wie
' diese Arbeitsmappe) und trägt die Werte in Spalte O
' (Grund zum Absprung) des Pipeline-Sheets ein.
'
' Aufruf: ImportGrundAbsprungMapping
' ============================================================
Option Explicit

Public Sub ImportGrundAbsprungMapping()

    Dim sMapFile    As String
    Dim wbMap       As Workbook
    Dim wsMap       As Worksheet
    Dim wsPipe      As Worksheet
    Dim lastRow     As Long
    Dim i           As Long
    Dim zeilePipe   As Long
    Dim wert        As String
    Dim changed     As Long

    ' --- Pfad zur Mapping-Datei ---
    sMapFile = ThisWorkbook.Path & Application.PathSeparator & "grund_absprung_mapping.xlsx"

    If Dir(sMapFile) = "" Then
        MsgBox "Mapping-Datei nicht gefunden:" & vbCrLf & sMapFile, vbCritical
        Exit Sub
    End If

    ' --- Mapping-Datei öffnen (unsichtbar) ---
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Set wbMap = Workbooks.Open(Filename:=sMapFile, ReadOnly:=True, UpdateLinks:=False)
    Set wsMap = wbMap.Sheets("Mapping")
    Set wsPipe = ThisWorkbook.Sheets("Pipeline")

    ' --- Letzte Zeile im Mapping ---
    lastRow = wsMap.Cells(wsMap.Rows.Count, 1).End(xlUp).Row

    changed = 0

    ' --- Werte eintragen ---
    For i = 2 To lastRow  ' Zeile 1 = Header
        zeilePipe = CLng(wsMap.Cells(i, 1).Value)  ' Spalte A = Zeilennummer
        wert = CStr(wsMap.Cells(i, 2).Value)       ' Spalte B = Wert

        ' Nur wenn Zeilennummer plausibel
        If zeilePipe >= 7 And zeilePipe <= 500 Then
            wsPipe.Cells(zeilePipe, 15).Value = wert  ' Spalte O = 15
            changed = changed + 1
        End If
    Next i

    ' --- Mapping-Datei schliessen ---
    wbMap.Close SaveChanges:=False

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    MsgBox changed & " Werte in Spalte 'Grund zum Absprung' eingetragen.", vbInformation

End Sub
