Option Explicit

' =========================
' Code Description
' =========================
' Dieses Modul liest Apple Mail Nachrichten mit den Schlagworten "Lead" oder
' "Neue Anfrage", parst die Inhalte und schreibt die Daten in die intelligente
' Tabelle "Kundenliste" auf dem Blatt "Pipeline".
'
' Fokus: clean, simpel, skalierbar
' - klare Unterfunktionen
' - saubere Variablendeklaration
' - robuste Parsing-Logik (Label/Value + Abschnittserkennung)

' =========================
' Konfiguration
' =========================
Private Const SHEET_NAME As String = "Pipeline"
Private Const TABLE_NAME As String = "Kundenliste"
Private Const LEAD_SOURCE As String = "Apple Mail"

Private Const KEYWORD_1 As String = "Lead"
Private Const KEYWORD_2 As String = "Neue Anfrage"

Private Const MSG_DELIM As String = "<<<MSG>>>"
Private Const DATE_TAG As String = "DATE:"
Private Const SUBJECT_TAG As String = "SUBJECT:"
Private Const BODY_TAG As String = "BODY:"

Private Const APPLESCRIPT_FILE As String = "MailReader.scpt"
Private Const APPLESCRIPT_HANDLER As String = "FetchMessages"

' =========================
' Public Entry
' =========================
Public Sub ImportLeadsFromAppleMail()
    ' --- Variablen (Objekte) ---
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim messagesText As String
    Dim messages() As String
    Dim msgBlock As Variant
    Dim payload As Object
    Dim parsed As Object

    ' --- Variablen (Primitives) ---
    Dim i As Long
    Dim msgDate As Date
    Dim msgSubject As String
    Dim msgBody As String
    Dim leadType As String

    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)
    Set tbl = ws.ListObjects(TABLE_NAME)

    messagesText = FetchAppleMailMessages(KEYWORD_1, KEYWORD_2)
    If Len(messagesText) = 0 Then Exit Sub

    messages = Split(messagesText, MSG_DELIM)

    For Each msgBlock In messages
        If Trim$(msgBlock) <> vbNullString Then
            Set payload = ParseMessageBlock(CStr(msgBlock))

            msgDate = payload("Date")
            msgSubject = payload("Subject")
            msgBody = payload("Body")

            leadType = ResolveLeadType(msgSubject, msgBody)

            Set parsed = ParseLeadContent(msgBody)

            If Not LeadAlreadyExists(tbl, parsed, msgDate) Then
                AddLeadRow tbl, parsed, msgDate, leadType
            End If
        End If
    Next msgBlock
End Sub

' =========================
' Apple Mail Read
' =========================
Private Function FetchAppleMailMessages(ByVal keywordA As String, ByVal keywordB As String) As String
    ' Holt Nachrichteninhalte aus Apple Mail (Inbox) per AppleScript
    Dim script As String
    Dim result As String

    script = "" & _
    "tell application ""Mail""" & vbLf & _
    "set theMessages to (every message of inbox whose subject contains """ & keywordA & """ or subject contains """ & keywordB & """ or content contains """ & keywordA & """ or content contains """ & keywordB & """ )" & vbLf & _
    "set outText to """"" & vbLf & _
    "repeat with m in theMessages" & vbLf & _
    "set outText to outText & """ & MSG_DELIM & """ & linefeed" & vbLf & _
    "set outText to outText & """ & DATE_TAG & """ & (date sent of m) & linefeed" & vbLf & _
    "set outText to outText & """ & SUBJECT_TAG & """ & (subject of m) & linefeed" & vbLf & _
    "set outText to outText & """ & BODY_TAG & """ & (content of m) & linefeed" & vbLf & _
    "end repeat" & vbLf & _
    "return outText" & vbLf & _
    "end tell"

    On Error GoTo ErrHandler
    result = AppleScriptTask(APPLESCRIPT_FILE, APPLESCRIPT_HANDLER, script)
    FetchAppleMailMessages = result
    Exit Function

ErrHandler:
    ' Häufige Ursachen:
    ' 1) Script nicht installiert: ~/Library/Application Scripts/com.microsoft.Excel/MailReader.scpt
    ' 2) Fehlende Automation-Rechte (Systemeinstellungen > Datenschutz & Sicherheit > Automation)
    ' Excel muss Apple Mail steuern dürfen.
    MsgBox "AppleScriptTask-Fehler. Prüfe Script-Installation und Automation-Rechte.", vbExclamation
    FetchAppleMailMessages = vbNullString
End Function

' =========================
' Message Parsing
' =========================
Private Function ParseMessageBlock(ByVal blockText As String) As Object
    ' Extrahiert Datum, Betreff und Body aus einem Message-Block
    Dim lines() As String
    Dim i As Long
    Dim lineText As String
    Dim payload As Object

    Set payload = CreateObject("Scripting.Dictionary")
    payload.CompareMode = vbTextCompare

    payload("Date") = Date
    payload("Subject") = vbNullString
    payload("Body") = vbNullString

    lines = Split(blockText, vbLf)
    For i = LBound(lines) To UBound(lines)
        lineText = Trim$(lines(i))
        If Len(lineText) = 0 Then GoTo NextLine

        If Left$(lineText, Len(DATE_TAG)) = DATE_TAG Then
            payload("Date") = CDate(Trim$(Mid$(lineText, Len(DATE_TAG) + 1)))
        ElseIf Left$(lineText, Len(SUBJECT_TAG)) = SUBJECT_TAG Then
            payload("Subject") = Trim$(Mid$(lineText, Len(SUBJECT_TAG) + 1))
        ElseIf Left$(lineText, Len(BODY_TAG)) = BODY_TAG Then
            payload("Body") = Trim$(Mid$(lineText, Len(BODY_TAG) + 1)) & vbLf
        Else
            payload("Body") = payload("Body") & lineText & vbLf
        End If
NextLine:
    Next i

    Set ParseMessageBlock = payload
End Function

Private Function ResolveLeadType(ByVal subjectText As String, ByVal bodyText As String) As String
    If InStr(1, subjectText, KEYWORD_2, vbTextCompare) > 0 Or InStr(1, bodyText, KEYWORD_2, vbTextCompare) > 0 Then
        ResolveLeadType = KEYWORD_2
    Else
        ResolveLeadType = KEYWORD_1
    End If
End Function

Private Function ParseLeadContent(ByVal bodyText As String) As Object
    ' Parst den Nachrichtentext in strukturierte Felder
    Dim result As Object
    Dim lines() As String
    Dim i As Long
    Dim lineText As String
    Dim currentSection As String
    Dim pendingKey As String

    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare

    currentSection = "Kontakt"
    pendingKey = vbNullString

    lines = Split(bodyText, vbLf)
    For i = LBound(lines) To UBound(lines)
        lineText = Trim$(lines(i))
        If Len(lineText) = 0 Then GoTo NextLine

        If InStr(1, lineText, "Informationen zum Senior", vbTextCompare) > 0 Then
            currentSection = "Senior"
            pendingKey = vbNullString
            GoTo NextLine
        End If

        If Right$(lineText, 1) = ":" Then
            pendingKey = Left$(lineText, Len(lineText) - 1)
            GoTo NextLine
        End If

        If Len(pendingKey) > 0 Then
            MapLabelValue result, pendingKey, lineText, currentSection
            pendingKey = vbNullString
        ElseIf InStr(lineText, ":") > 0 Then
            MapInlinePair result, lineText, currentSection
        End If
NextLine:
    Next i

    Set ParseLeadContent = result
End Function

Private Sub MapInlinePair(ByRef fields As Object, ByVal lineText As String, ByVal sectionName As String)
    Dim keyPart As String
    Dim valuePart As String

    keyPart = Trim$(Left$(lineText, InStr(lineText, ":") - 1))
    valuePart = Trim$(Mid$(lineText, InStr(lineText, ":") + 1))

    If Len(keyPart) > 0 Then
        MapLabelValue fields, keyPart, valuePart, sectionName
    End If
End Sub

Private Sub MapLabelValue(ByRef fields As Object, ByVal rawKey As String, ByVal rawValue As String, ByVal sectionName As String)
    Dim keyNorm As String
    Dim valueNorm As String

    keyNorm = NormalizeKey(rawKey)
    valueNorm = Trim$(rawValue)

    Select Case keyNorm
        Case "anrede": fields("Kontakt_Anrede") = valueNorm
        Case "vorname": fields("Kontakt_Vorname") = valueNorm
        Case "nachname": fields("Kontakt_Nachname") = valueNorm
        Case "name"
            If LCase$(sectionName) = "senior" Then
                fields("Senior_Name") = valueNorm
            Else
                fields("Kontakt_Name") = valueNorm
            End If
        Case "mobil", "telefonnummer": fields("Kontakt_Mobil") = valueNorm
        Case "e-mail", "e-mail-adresse": fields("Kontakt_Email") = valueNorm
        Case "erreichbarkeit": fields("Kontakt_Erreichbarkeit") = valueNorm
        Case "beziehung": fields("Senior_Beziehung") = valueNorm
        Case "alter": fields("Senior_Alter") = valueNorm
        Case "pflegegrad status": fields("Senior_Pflegegrad_Status") = valueNorm
        Case "pflegegrad": fields("Senior_Pflegegrad") = valueNorm
        Case "lebenssituation": fields("Senior_Lebenssituation") = valueNorm
        Case "mobilität": fields("Senior_Mobilitaet") = valueNorm
        Case "medizinisches": fields("Senior_Medizinisches") = valueNorm
        Case "postleitzahl", "plz": fields("PLZ") = valueNorm
        Case "nutzer": fields("Nutzer") = valueNorm
        Case "alltagshilfe aufgaben": fields("Alltagshilfe_Aufgaben") = valueNorm
        Case "alltagshilfe häufigkeit": fields("Alltagshilfe_Haeufigkeit") = valueNorm
        Case "id": fields("Anfrage_ID") = valueNorm
    End Select
End Sub

Private Function NormalizeKey(ByVal rawKey As String) As String
    Dim k As String
    k = LCase$(Trim$(rawKey))
    k = Replace(k, vbTab, " ")
    k = Replace(k, "  ", " ")
    NormalizeKey = k
End Function

' =========================
' Excel Output
' =========================
Private Sub AddLeadRow(ByVal tbl As ListObject, ByVal fields As Object, ByVal msgDate As Date, ByVal leadType As String)
    Dim newRow As ListRow
    Dim colIndex As Long

    Set newRow = tbl.ListRows.Add

    SetCellByHeader newRow, "Monat Lead erhalten", DateSerial(Year(msgDate), Month(msgDate), 1)
    SetCellByHeader newRow, "Lead-Quelle", LEAD_SOURCE
    SetCellByHeader newRow, "Leadtyp", leadType
    SetCellByHeader newRow, "Name", ResolveKontaktName(fields)
    SetCellByHeader newRow, "Telefonnummer", GetField(fields, "Kontakt_Mobil")
    SetCellByHeader newRow, "PLZ", GetField(fields, "PLZ")
    SetCellByHeader newRow, "PG", GetField(fields, "Senior_Pflegegrad")
    SetCellByHeader newRow, "Notizen", BuildNotes(fields)
End Sub

Private Sub SetCellByHeader(ByVal rowItem As ListRow, ByVal headerName As String, ByVal valueToSet As Variant)
    Dim idx As Long
    idx = GetColumnIndex(rowItem.Parent, headerName)
    If idx > 0 Then
        rowItem.Range.Cells(1, idx).Value = valueToSet
    End If
End Sub

Private Function GetColumnIndex(ByVal tbl As ListObject, ByVal headerName As String) As Long
    Dim i As Long
    For i = 1 To tbl.ListColumns.Count
        If StrComp(Trim$(tbl.ListColumns(i).Name), headerName, vbTextCompare) = 0 Then
            GetColumnIndex = i
            Exit Function
        End If
    Next i
    GetColumnIndex = 0
End Function

Private Function ResolveKontaktName(ByVal fields As Object) As String
    Dim fullName As String
    fullName = GetField(fields, "Kontakt_Name")

    If Len(fullName) = 0 Then
        fullName = Trim$(GetField(fields, "Kontakt_Vorname") & " " & GetField(fields, "Kontakt_Nachname"))
    End If

    ResolveKontaktName = fullName
End Function

Private Function GetField(ByVal fields As Object, ByVal keyName As String) As String
    If fields.Exists(keyName) Then
        GetField = CStr(fields(keyName))
    Else
        GetField = vbNullString
    End If
End Function

Private Function BuildNotes(ByVal fields As Object) As String
    Dim notes As String

    notes = ""
    notes = AppendNote(notes, "E-Mail", GetField(fields, "Kontakt_Email"))
    notes = AppendNote(notes, "Erreichbarkeit", GetField(fields, "Kontakt_Erreichbarkeit"))
    notes = AppendNote(notes, "Senior Name", GetField(fields, "Senior_Name"))
    notes = AppendNote(notes, "Beziehung", GetField(fields, "Senior_Beziehung"))
    notes = AppendNote(notes, "Alter", GetField(fields, "Senior_Alter"))
    notes = AppendNote(notes, "Pflegegrad Status", GetField(fields, "Senior_Pflegegrad_Status"))
    notes = AppendNote(notes, "Lebenssituation", GetField(fields, "Senior_Lebenssituation"))
    notes = AppendNote(notes, "Mobilität", GetField(fields, "Senior_Mobilitaet"))
    notes = AppendNote(notes, "Medizinisches", GetField(fields, "Senior_Medizinisches"))
    notes = AppendNote(notes, "Nutzer", GetField(fields, "Nutzer"))
    notes = AppendNote(notes, "Alltagshilfe Aufgaben", GetField(fields, "Alltagshilfe_Aufgaben"))
    notes = AppendNote(notes, "Alltagshilfe Häufigkeit", GetField(fields, "Alltagshilfe_Haeufigkeit"))
    notes = AppendNote(notes, "ID", GetField(fields, "Anfrage_ID"))

    BuildNotes = notes
End Function

Private Function AppendNote(ByVal currentText As String, ByVal labelText As String, ByVal valueText As String) As String
    If Len(Trim$(valueText)) = 0 Then
        AppendNote = currentText
    ElseIf Len(currentText) = 0 Then
        AppendNote = labelText & ": " & valueText
    Else
        AppendNote = currentText & " | " & labelText & ": " & valueText
    End If
End Function

' =========================
' Duplicate Handling
' =========================
Private Function LeadAlreadyExists(ByVal tbl As ListObject, ByVal fields As Object, ByVal msgDate As Date) As Boolean
    Dim idValue As String
    Dim nameValue As String
    Dim phoneValue As String
    Dim notesColIndex As Long
    Dim nameColIndex As Long
    Dim phoneColIndex As Long
    Dim dateColIndex As Long
    Dim i As Long

    idValue = GetField(fields, "Anfrage_ID")
    nameValue = ResolveKontaktName(fields)
    phoneValue = GetField(fields, "Kontakt_Mobil")

    notesColIndex = GetColumnIndex(tbl, "Notizen")
    nameColIndex = GetColumnIndex(tbl, "Name")
    phoneColIndex = GetColumnIndex(tbl, "Telefonnummer")
    dateColIndex = GetColumnIndex(tbl, "Monat Lead erhalten")

    If tbl.ListRows.Count = 0 Then Exit Function

    For i = 1 To tbl.ListRows.Count
        If Len(idValue) > 0 And notesColIndex > 0 Then
            If InStr(1, CStr(tbl.DataBodyRange.Cells(i, notesColIndex).Value), "ID: " & idValue, vbTextCompare) > 0 Then
                LeadAlreadyExists = True
                Exit Function
            End If
        End If

        If Len(nameValue) > 0 And Len(phoneValue) > 0 And nameColIndex > 0 And phoneColIndex > 0 And dateColIndex > 0 Then
            If StrComp(CStr(tbl.DataBodyRange.Cells(i, nameColIndex).Value), nameValue, vbTextCompare) = 0 And _
               StrComp(CStr(tbl.DataBodyRange.Cells(i, phoneColIndex).Value), phoneValue, vbTextCompare) = 0 And _
               DateSerial(Year(CDate(tbl.DataBodyRange.Cells(i, dateColIndex).Value)), Month(CDate(tbl.DataBodyRange.Cells(i, dateColIndex).Value)), 1) = DateSerial(Year(msgDate), Month(msgDate), 1) Then
                LeadAlreadyExists = True
                Exit Function
            End If
        End If
    Next i
End Function
