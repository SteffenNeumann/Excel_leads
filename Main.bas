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

Private Const MAX_MESSAGES As Long = 50

Private Const APPLESCRIPT_FILE As String = "MailReader.scpt"
Private Const APPLESCRIPT_HANDLER As String = "FetchMessages"
Private Const APPLESCRIPT_SOURCE As String = "MailReader.applescript"
Private Const AUTO_INSTALL_APPLESCRIPT As Boolean = False

' Zielordner in Apple Mail
' LEAD_FOLDER darf Teilstring sein (z. B. "Archiv" für "Archiv — iCloud")
' Optional: LEAD_MAILBOX leer lassen, um global zu suchen.
Private Const LEAD_MAILBOX As String = "steffen.neumann.ic@icloud.com"
Private Const LEAD_FOLDER As String = "Archiv"

' =========================
' Public Entry
' =========================
Public Sub ImportLeadsFromAppleMail()
    If AUTO_INSTALL_APPLESCRIPT Then
        EnsureAppleScriptInstalled
    End If

    ' --- Variablen (Objekte) ---
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim payload As Object
    Dim parsed As Object

    ' --- Variablen (Primitives) ---
    Dim v As Variant
    Dim msgDate As Date
    Dim msgSubject As String
    Dim msgBody As String
    Dim leadType As String

    Dim messagesText As String
    Dim messages() As String
    Dim msgBlock As Variant

    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)
    Set tbl = ws.ListObjects(TABLE_NAME)

    messagesText = FetchAppleMailMessages(KEYWORD_1, KEYWORD_2)
    If Len(messagesText) = 0 Then Exit Sub

    messages = Split(messagesText, MSG_DELIM)

    For Each msgBlock In messages
        If Trim$(msgBlock) <> vbNullString Then
            Set payload = ParseMessageBlock(CStr(msgBlock))

            msgDate = Date
            msgSubject = vbNullString
            msgBody = vbNullString
            If TryGetKV(payload, "Date", v) Then msgDate = CDate(v)
            If TryGetKV(payload, "Subject", v) Then msgSubject = CStr(v)
            If TryGetKV(payload, "Body", v) Then msgBody = CStr(v)

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
    Dim q As String

    q = Chr$(34)

    script = ""
    script = script & "with timeout of 30 seconds" & vbLf
    script = script & "tell application ""Mail""" & vbLf
    script = script & "set targetBox to missing value" & vbLf
    If Len(LEAD_MAILBOX) > 0 Then
        script = script & "set targetAccountName to " & q & LEAD_MAILBOX & q & vbLf
        script = script & "try" & vbLf
        script = script & "repeat with a in accounts" & vbLf
        script = script & "if (name of a) contains targetAccountName then" & vbLf
        script = script & "try" & vbLf
        script = script & "set targetBox to first mailbox of a whose name contains " & q & LEAD_FOLDER & q & vbLf
        script = script & "exit repeat" & vbLf
        script = script & "end try" & vbLf
        script = script & "end if" & vbLf
        script = script & "end repeat" & vbLf
        script = script & "end try" & vbLf
    End If
    script = script & "if targetBox is missing value then" & vbLf
    script = script & "try" & vbLf
    script = script & "set targetBox to first mailbox whose name contains """ & LEAD_FOLDER & """" & vbLf
    script = script & "end try" & vbLf
    script = script & "end if" & vbLf
    script = script & "if targetBox is missing value then error ""Mailbox nicht gefunden: " & LEAD_FOLDER & """" & vbLf
    script = script & "set theMessages to (every message of targetBox whose subject contains """ & keywordA & """ or subject contains """ & keywordB & """ or content contains """ & keywordA & """ or content contains """ & keywordB & """ )" & vbLf
    script = script & "if (count of theMessages) > " & MAX_MESSAGES & " then set theMessages to items 1 thru " & MAX_MESSAGES & " of theMessages" & vbLf
    script = script & "set outText to """"" & vbLf
    script = script & "repeat with m in theMessages" & vbLf
    script = script & "set outText to outText & """ & MSG_DELIM & """ & linefeed" & vbLf
    script = script & "set outText to outText & """ & DATE_TAG & """ & (date sent of m) & linefeed" & vbLf
    script = script & "set outText to outText & """ & SUBJECT_TAG & """ & (subject of m) & linefeed" & vbLf
        script = script & "set bodyText to " & q & q & vbLf
        script = script & "try" & vbLf
        script = script & "set theAtts to (mail attachments of m)" & vbLf
        script = script & "repeat with a in theAtts" & vbLf
        script = script & "set attName to (name of a)" & vbLf
        script = script & "set attLower to attName as string" & vbLf
        script = script & "try" & vbLf
        script = script & "set attLower to do shell script " & q & "python3 -c 'import sys; print(sys.argv[1].lower())' " & q & " & quoted form of attLower" & vbLf
        script = script & "end try" & vbLf
        script = script & "if (attLower ends with "".txt"") or (attLower ends with "".csv"") or (attLower ends with "".log"") or (attLower ends with "".json"") or (attLower ends with "".xml"") or (attLower ends with "".html"") or (attLower ends with "".htm"") then" & vbLf
        script = script & "set tmpDir to POSIX path of (path to temporary items)" & vbLf
        script = script & "set tmpPath to tmpDir & ""mail-"" & (do shell script ""date +%s"") & ""-"" & attName" & vbLf
        script = script & "try" & vbLf
        script = script & "save a in (POSIX file tmpPath)" & vbLf
        script = script & "if (attLower ends with "".html"") or (attLower ends with "".htm"") then" & vbLf
        script = script & "set bodyText to do shell script ""/usr/bin/textutil -convert txt -stdout "" & quoted form of tmpPath" & vbLf
        script = script & "else" & vbLf
        script = script & "set bodyText to do shell script ""/bin/cat "" & quoted form of tmpPath" & vbLf
        script = script & "end if" & vbLf
        script = script & "end try" & vbLf
        script = script & "if bodyText is not " & q & q & " then exit repeat" & vbLf
        script = script & "end if" & vbLf
        script = script & "end repeat" & vbLf
        script = script & "end try" & vbLf
        script = script & "if bodyText is " & q & q & " then set bodyText to (content of m)" & vbLf
        script = script & "set outText to outText & " & q & BODY_TAG & q & " & bodyText & linefeed" & vbLf
    script = script & "end repeat" & vbLf
    script = script & "return outText" & vbLf
    script = script & "end tell" & vbLf
    script = script & "end timeout"

    On Error GoTo ErrHandler
    result = AppleScriptTask(APPLESCRIPT_FILE, APPLESCRIPT_HANDLER, script)
    If Left$(result, 6) = "ERROR:" Then
        MsgBox "AppleScript-Fehler: " & Mid$(result, 7), vbExclamation
        FetchAppleMailMessages = vbNullString
        Exit Function
    End If
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

Public Sub DebugPrintAppleMailFolders()
    Dim folderText As String
    Dim lines() As String
    Dim i As Long

    folderText = FetchAppleMailFolderList()
    If Len(folderText) = 0 Then Exit Sub

    lines = Split(folderText, vbLf)
    For i = LBound(lines) To UBound(lines)
        If Len(Trim$(lines(i))) > 0 Then
            Debug.Print Trim$(lines(i))
        End If
    Next i
End Sub

Private Function FetchAppleMailFolderList() As String
    Dim script As String
    Dim result As String
    Dim q As String

    q = Chr$(34)

    script = ""
    script = script & "with timeout of 30 seconds" & vbLf
    script = script & "tell application ""Mail""" & vbLf
    script = script & "set targetAccountName to " & q & LEAD_MAILBOX & q & vbLf
    script = script & "script Dump" & vbLf
    script = script & "property outText : " & q & q & vbLf
    script = script & "on addLine(t)" & vbLf
    script = script & "set outText to outText & t & linefeed" & vbLf
    script = script & "end addLine" & vbLf
    script = script & "on walk(boxList, prefix)" & vbLf
    script = script & "repeat with mb in boxList" & vbLf
    script = script & "set mbName to (name of mb)" & vbLf
    script = script & "my addLine(prefix & mbName)" & vbLf
    script = script & "try" & vbLf
    script = script & "set kids to (every mailbox of mb)" & vbLf
    script = script & "if (count of kids) > 0 then my walk(kids, prefix & mbName & " & q & " / " & q & " )" & vbLf
    script = script & "end try" & vbLf
    script = script & "end repeat" & vbLf
    script = script & "end walk" & vbLf
    script = script & "end script" & vbLf
    script = script & "set outText of Dump to " & q & q & vbLf
    script = script & "set matchedCount to 0" & vbLf
    script = script & "repeat with a in accounts" & vbLf
    script = script & "set aName to (name of a)" & vbLf
    script = script & "if (targetAccountName is " & q & q & ") or (aName contains targetAccountName) then" & vbLf
    script = script & "set matchedCount to matchedCount + 1" & vbLf
    script = script & "Dump's addLine(" & q & "ACCOUNT: " & q & " & aName)" & vbLf
    script = script & "Dump's walk((mailboxes of a), " & q & q & ")" & vbLf
    script = script & "Dump's addLine(" & q & q & ")" & vbLf
    script = script & "end if" & vbLf
    script = script & "end repeat" & vbLf
    script = script & "if (targetAccountName is not " & q & q & ") and (matchedCount is 0) then" & vbLf
    script = script & "Dump's addLine(" & q & "NO MATCH FOR ACCOUNT FILTER: " & q & " & targetAccountName)" & vbLf
    script = script & "Dump's addLine(" & q & q & ")" & vbLf
    script = script & "repeat with a in accounts" & vbLf
    script = script & "set aName to (name of a)" & vbLf
    script = script & "Dump's addLine(" & q & "ACCOUNT: " & q & " & aName)" & vbLf
    script = script & "Dump's walk((mailboxes of a), " & q & q & ")" & vbLf
    script = script & "Dump's addLine(" & q & q & ")" & vbLf
    script = script & "end repeat" & vbLf
    script = script & "end if" & vbLf
    script = script & "return (outText of Dump)" & vbLf
    script = script & "end tell" & vbLf
    script = script & "end timeout"

    On Error GoTo ErrHandler
    result = AppleScriptTask(APPLESCRIPT_FILE, APPLESCRIPT_HANDLER, script)
    If Left$(result, 6) = "ERROR:" Then
        MsgBox "AppleScript-Fehler: " & Mid$(result, 7), vbExclamation
        FetchAppleMailFolderList = vbNullString
        Exit Function
    End If
    FetchAppleMailFolderList = result
    Exit Function

ErrHandler:
    MsgBox "AppleScriptTask-Fehler. Prüfe Script-Installation und Automation-Rechte.", vbExclamation
    FetchAppleMailFolderList = vbNullString
End Function

' =========================
' AppleScript Setup
' =========================
Private Sub EnsureAppleScriptInstalled()
    Dim targetPath As String
    Dim sourcePath As String

    targetPath = GetAppleScriptTargetPath()
    sourcePath = ThisWorkbook.Path & "/" & APPLESCRIPT_SOURCE

    If Len(Dir$(targetPath)) = 0 Then
        InstallAppleScript sourcePath, targetPath
    End If
End Sub

Private Function GetAppleScriptTargetPath() As String
    Dim homePath As String
    homePath = Environ$("HOME")
    GetAppleScriptTargetPath = homePath & "/Library/Application Scripts/com.microsoft.Excel/" & APPLESCRIPT_FILE
End Function

Private Sub InstallAppleScript(ByVal sourcePath As String, ByVal targetPath As String)
    Dim folderPath As String

    folderPath = Left$(targetPath, InStrRev(targetPath, "/") - 1)
    EnsureFolderExists folderPath

    If Len(Dir$(sourcePath)) = 0 Then
        MsgBox "AppleScript-Quelle fehlt: " & sourcePath, vbExclamation
        Exit Sub
    End If

    On Error GoTo ErrHandler

    If Len(Dir$(targetPath)) > 0 Then Kill targetPath

    FileCopy sourcePath, targetPath
    Exit Sub

ErrHandler:
    If Err.Number = 75 Then
        MsgBox "Zugriff verweigert. Bitte manuell kopieren nach: " & folderPath & " oder AUTO_INSTALL_APPLESCRIPT aktivieren.", vbExclamation
        Exit Sub
    End If

    MsgBox "AppleScript konnte nicht installiert werden. Prüfe Rechte.", vbExclamation
End Sub

Private Sub EnsureFolderExists(ByVal folderPath As String)
    Dim parts() As String
    Dim i As Long
    Dim currentPath As String

    parts = Split(folderPath, "/")
    currentPath = ""

    For i = LBound(parts) To UBound(parts)
        If Len(parts(i)) > 0 Then
            currentPath = currentPath & "/" & parts(i)
            If Len(Dir$(currentPath, vbDirectory)) = 0 Then
                On Error Resume Next
                MkDir currentPath
                On Error GoTo 0
            End If
        End If
    Next i
End Sub


' =========================
' Cross-Platform Key/Value Store (macOS-safe)
' =========================
Private Function NewKeyValueStore() As Object
    ' macOS: kein ActiveX (Scripting.Dictionary) verfügbar
    ' Windows: Dictionary ist ok und schneller
    If InStr(1, Application.OperatingSystem, "Mac", vbTextCompare) > 0 Then
        Set NewKeyValueStore = New Collection
    Else
        Dim d As Object
        Set d = CreateObject("Scripting.Dictionary")
        d.CompareMode = vbTextCompare
        Set NewKeyValueStore = d
    End If

End Function

Private Function KeyNorm(ByVal keyName As String) As String
    KeyNorm = LCase$(Trim$(keyName))
End Function

Private Sub SetKV(ByRef store As Object, ByVal keyName As String, ByVal valueToSet As Variant)
    Dim k As String
    k = KeyNorm(keyName)

    If TypeName(store) = "Dictionary" Then
        store(k) = valueToSet
    Else
        On Error Resume Next
        store.Remove k
        On Error GoTo 0
        store.Add valueToSet, k
    End If
End Sub

Private Function TryGetKV(ByVal store As Object, ByVal keyName As String, ByRef valueOut As Variant) As Boolean
    Dim k As String
    k = KeyNorm(keyName)

    If TypeName(store) = "Dictionary" Then
        If store.Exists(k) Then
            valueOut = store(k)
            TryGetKV = True
        End If
    Else
        On Error GoTo NotFound
        valueOut = store(k)
        TryGetKV = True
        Exit Function
NotFound:
        TryGetKV = False
    End If
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

    Set payload = NewKeyValueStore()
    SetKV payload, "Date", Date
    SetKV payload, "Subject", vbNullString
    SetKV payload, "Body", vbNullString

    lines = Split(blockText, vbLf)
    For i = LBound(lines) To UBound(lines)
        lineText = Trim$(lines(i))
        If Len(lineText) = 0 Then GoTo NextLine

        If Left$(lineText, Len(DATE_TAG)) = DATE_TAG Then
            SetKV payload, "Date", ParseAppleMailDate(Trim$(Mid$(lineText, Len(DATE_TAG) + 1)))
        ElseIf Left$(lineText, Len(SUBJECT_TAG)) = SUBJECT_TAG Then
            SetKV payload, "Subject", Trim$(Mid$(lineText, Len(SUBJECT_TAG) + 1))
        ElseIf Left$(lineText, Len(BODY_TAG)) = BODY_TAG Then
            SetKV payload, "Body", Trim$(Mid$(lineText, Len(BODY_TAG) + 1)) & vbLf
        Else
            Dim curBody As Variant
            If Not TryGetKV(payload, "Body", curBody) Then curBody = vbNullString
            SetKV payload, "Body", CStr(curBody) & lineText & vbLf
        End If
NextLine:
    Next i

    Set ParseMessageBlock = payload
End Function

Private Function ParseAppleMailDate(ByVal dateText As String) As Date
    Dim t As String
    Dim parts() As String
    Dim datePart As String
    Dim timePart As String
    Dim dayNum As Long
    Dim monthNum As Long
    Dim yearNum As Long
    Dim timeParts() As String
    Dim h As Long, m As Long, s As Long

    t = Trim$(dateText)
    If InStr(t, ",") > 0 Then
        t = Trim$(Mid$(t, InStr(t, ",") + 1))
    End If

    t = Replace(t, " um ", " ")

    On Error GoTo Fallback
    ParseAppleMailDate = CDate(t)
    Exit Function

Fallback:
    On Error GoTo ErrHandler
    parts = Split(t, " ")
    If UBound(parts) < 2 Then GoTo ErrHandler

    dayNum = CLng(Replace(parts(0), ".", ""))
    monthNum = GermanMonthToNumber(parts(1))
    yearNum = CLng(parts(2))

    timePart = vbNullString
    If UBound(parts) >= 3 Then timePart = parts(3)

    h = 0: m = 0: s = 0
    If Len(timePart) > 0 Then
        timeParts = Split(timePart, ":")
        If UBound(timeParts) >= 0 Then h = CLng(timeParts(0))
        If UBound(timeParts) >= 1 Then m = CLng(timeParts(1))
        If UBound(timeParts) >= 2 Then s = CLng(timeParts(2))
    End If

    ParseAppleMailDate = DateSerial(yearNum, monthNum, dayNum) + TimeSerial(h, m, s)
    Exit Function

ErrHandler:
    ParseAppleMailDate = Date
End Function

Private Function GermanMonthToNumber(ByVal monthText As String) As Long
    Dim m As String
    m = LCase$(Trim$(monthText))

    Select Case m
        Case "januar": GermanMonthToNumber = 1
        Case "februar": GermanMonthToNumber = 2
        Case "märz", "maerz": GermanMonthToNumber = 3
        Case "april": GermanMonthToNumber = 4
        Case "mai": GermanMonthToNumber = 5
        Case "juni": GermanMonthToNumber = 6
        Case "juli": GermanMonthToNumber = 7
        Case "august": GermanMonthToNumber = 8
        Case "september": GermanMonthToNumber = 9
        Case "oktober": GermanMonthToNumber = 10
        Case "november": GermanMonthToNumber = 11
        Case "dezember": GermanMonthToNumber = 12
        Case Else: GermanMonthToNumber = 1
    End Select
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

    Set result = NewKeyValueStore()

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
        Case "anrede": SetKV fields, "Kontakt_Anrede", valueNorm
        Case "vorname": SetKV fields, "Kontakt_Vorname", valueNorm
        Case "nachname": SetKV fields, "Kontakt_Nachname", valueNorm
        Case "name"
            If LCase$(sectionName) = "senior" Then
                SetKV fields, "Senior_Name", valueNorm
            Else
                SetKV fields, "Kontakt_Name", valueNorm
            End If
        Case "mobil", "telefonnummer": SetKV fields, "Kontakt_Mobil", valueNorm
        Case "e-mail", "e-mail-adresse": SetKV fields, "Kontakt_Email", valueNorm
        Case "erreichbarkeit": SetKV fields, "Kontakt_Erreichbarkeit", valueNorm
        Case "beziehung": SetKV fields, "Senior_Beziehung", valueNorm
        Case "alter": SetKV fields, "Senior_Alter", valueNorm
        Case "pflegegrad status": SetKV fields, "Senior_Pflegegrad_Status", valueNorm
        Case "pflegegrad": SetKV fields, "Senior_Pflegegrad", valueNorm
        Case "lebenssituation": SetKV fields, "Senior_Lebenssituation", valueNorm
        Case "mobilität": SetKV fields, "Senior_Mobilitaet", valueNorm
        Case "medizinisches": SetKV fields, "Senior_Medizinisches", valueNorm
        Case "postleitzahl", "plz": SetKV fields, "PLZ", valueNorm
        Case "nutzer": SetKV fields, "Nutzer", valueNorm
        Case "alltagshilfe aufgaben": SetKV fields, "Alltagshilfe_Aufgaben", valueNorm
        Case "alltagshilfe häufigkeit": SetKV fields, "Alltagshilfe_Haeufigkeit", valueNorm
        Case "id": SetKV fields, "Anfrage_ID", valueNorm
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
    Dim v As Variant
    If TryGetKV(fields, keyName, v) Then
        GetField = CStr(v)
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
