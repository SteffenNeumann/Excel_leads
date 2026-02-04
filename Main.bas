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
Private Const FROM_TAG As String = "FROM:"
Private Const BODY_TAG As String = "BODY:"

Private Const MAX_MESSAGES As Long = 50

Private Const APPLESCRIPT_FILE As String = "MailReader.scpt"
Private Const APPLESCRIPT_HANDLER As String = "FetchMessages"
Private Const APPLESCRIPT_SOURCE As String = "MailReader.applescript"
Private Const AUTO_INSTALL_APPLESCRIPT As Boolean = False

' Zielordner in Apple Mail
' LEAD_FOLDER darf Teilstring sein (z. B. "Archiv" für "Archiv — iCloud")
' Optional: LEAD_MAILBOX leer lassen, um global zu suchen.
Private Const SETTINGS_SHEET As String = "Berechnung"
Private Const NAME_LEAD_MAILBOX As String = "LEAD_MAILBOX"
Private Const NAME_LEAD_FOLDER As String = "LEAD_FOLDER"
Private Const NAME_MAILPATH As String = "mailpath"
Private Const LEAD_MAILBOX_DEFAULT As String = "iCloud"
Private Const LEAD_FOLDER_DEFAULT As String = "Leads"

' Error Log
Private Const ERROR_LOG_SHEET As String = "ErrLog"

' Index für Dublettenprüfung während eines Imports
Private gLeadIndex As Object
Private gLeadIndexInitialized As Boolean

' =========================
' Funktionsübersicht & Abhängigkeiten
' =========================
' ImportLeadsFromAppleMail: Einstiegspunkt; ruft FetchAppleMailMessages, ParseMessageBlock, ParseLeadContent, LeadAlreadyExists, AddLeadRow. Rückgabe: Sub, schreibt Zeilen.
' FetchAppleMailMessages: Baut AppleScript, ruft AppleScriptTask; liefert zusammengefasste Roh-Nachrichten.
' DebugPrintAppleMailFolders: Debug-Ausgabe der Ordner; nutzt FetchAppleMailFolderList.
' FetchAppleMailFolderList: Baut AppleScript, ruft AppleScriptTask; liefert Ordnerliste als Text.
' EnsureAppleScriptInstalled / GetAppleScriptTargetPath / InstallAppleScript / EnsureFolderExists: Helfer zum Installieren des AppleScripts. Rückgabe: Pfade oder Seiteneffekt.
' NewKeyValueStore / keyNorm / SetKV / TryGetKV: Plattform-sicherer Key/Value-Store. Rückgabe: Collection/Dictionary oder Boolean.
' ParseMessageBlock: Zerlegt einen Nachrichtenblock in Date/Subject/From/Body; nutzt ParseAppleMailDate.
' ParseAppleMailDate / GermanMonthToNumber: Robust Datum parsen aus Apple Mail Text.
' ResolveLeadType: Leitet Lead-Typ aus Betreff/Body ab.
' ParseLeadContent: Parst Body in Felder; nutzt MapLabelValue, MapInlinePair, SetBedarfsort.
' MapInlinePair: Teilt Inline "key: value" und delegiert an MapLabelValue.
' MapLabelValue: Normalisiert Labels und schreibt ins Feld-Store; nutzt NormalizeKey, SetBedarfsort, SetKV.
' NormalizeKey: Vereinheitlicht Label-Texte.
' AddLeadRow: Schreibt Felder in Tabelle; nutzt BuildHeaderIndex, SetCellByHeaderMap, ResolveLeadSource, ResolveKontaktName, NormalizePflegegrad, BuildNotes.
' FindTableByName: Sucht ListObject.
' ResolveLeadSource: Fallback-Logik für Lead-Quelle.
' SetCellByHeaderMap / BuildHeaderIndex / GetHeaderIndex: Tabellen-Header-Utilities.
' ResolveKontaktName: Stellt Kontaktname zusammen; nutzt ExtractSenderName, GetField.
' ExtractSenderName: Extrahiert Name aus Absender-String.
' GetField: Sicheres Lesen aus KV-Store.
' BuildNotes: Baut Notizen-Text; nutzt AppendNote.
' SetBedarfsort: Splittet PLZ/Ort; nutzt FilterDigits, SetKV.
' FilterDigits: Filtert Ziffern.
' NormalizePflegegrad: Extrahiert Ziffern aus PG-Text.
' AppendNote: Fügt Notizen zusammen.
' LeadAlreadyExists: Duplikatsprüfung; nutzt ResolveKontaktName, GetField, BuildHeaderIndex, GetHeaderIndex.
'
' Abhängigkeitsgraph (vereinfacht)
' ImportLeadsFromAppleMail
'   -> FetchAppleMailMessages -> AppleScriptTask
'   -> ParseMessageBlock -> ParseAppleMailDate -> GermanMonthToNumber
'   -> ResolveLeadType
'   -> ParseLeadContent -> MapLabelValue / MapInlinePair -> SetBedarfsort -> FilterDigits
'   -> LeadAlreadyExists -> ResolveKontaktName -> ExtractSenderName
'   -> AddLeadRow -> BuildHeaderIndex -> SetCellByHeaderMap
'                -> ResolveLeadSource
'                -> NormalizePflegegrad
'                -> BuildNotes -> AppendNote

' =========================
' Public Entry
' =========================
Public Sub ImportLeadsFromAppleMail()
    ' Zweck: Apple-Mail-Leads abrufen, parsen und in die Tabelle schreiben.
    ' Abhängigkeiten: EnsureAppleScriptInstalled (optional), FetchAppleMailMessages, ParseMessageBlock, ResolveLeadType, ParseLeadContent, LeadAlreadyExists, AddLeadRow.
    ' Rückgabe: keine (fügt Zeilen in Tabelle ein).
    If AUTO_INSTALL_APPLESCRIPT Then
        EnsureAppleScriptInstalled
    End If

    ' --- Variablen (Objekte) ---
    Dim tbl As ListObject
    Dim payload As Object
    Dim parsed As Object

    ' --- Variablen (Primitives) ---
    Dim v As Variant
    Dim msgDate As Date
    Dim msgSubject As String
    Dim msgBody As String
    Dim msgFrom As String
    Dim leadType As String

    Dim messagesText As String
    Dim messages() As String
    Dim msgBlock As Variant
    Dim analyzedCount As Long
    Dim importedCount As Long
    Dim duplicateCount As Long
    Dim errorCount As Long

    Set tbl = FindTableByName(TABLE_NAME)
    If tbl Is Nothing Then
        MsgBox "Tabelle '" & TABLE_NAME & "' nicht gefunden.", vbExclamation
        Exit Sub
    End If

    messagesText = FetchAppleMailMessages(KEYWORD_1, KEYWORD_2)
    If Len(messagesText) = 0 Then Exit Sub

    Set gLeadIndex = BuildExistingLeadIndex(tbl)
    gLeadIndexInitialized = True

    messages = Split(messagesText, MSG_DELIM)

    For Each msgBlock In messages
        ' Schleife: jeden Nachrichtenblock einzeln verarbeiten.
        If Trim$(msgBlock) <> vbNullString Then
            analyzedCount = analyzedCount + 1
            On Error GoTo MsgError
            Set payload = ParseMessageBlock(CStr(msgBlock))

            msgDate = Date
            msgSubject = vbNullString
            msgBody = vbNullString
            msgFrom = vbNullString
            If TryGetKV(payload, "Date", v) Then msgDate = CDate(v)
            If TryGetKV(payload, "Subject", v) Then msgSubject = CStr(v)
            If TryGetKV(payload, "Body", v) Then msgBody = CStr(v)
            If TryGetKV(payload, "From", v) Then msgFrom = CStr(v)

            leadType = ResolveLeadType(msgSubject, msgBody)

            Set parsed = ParseLeadContent(msgBody)
            SetKV parsed, "From", msgFrom

            If Not LeadAlreadyExists(tbl, parsed, msgDate) Then
                AddLeadRow tbl, parsed, msgDate, leadType
                importedCount = importedCount + 1
                AddLeadToIndex parsed, msgDate
            Else
                duplicateCount = duplicateCount + 1
            End If
            On Error GoTo 0
        End If
NextMsg:
    Next msgBlock

    MsgBox "Import abgeschlossen. " & analyzedCount & " Daten analysiert, " & importedCount & " Daten übertragen. Duplikate: " & duplicateCount & ". Fehler: " & errorCount & ".", vbInformation
    Exit Sub

MsgError:
    errorCount = errorCount + 1
    LogImportError "Fehler beim Verarbeiten einer Nachricht", Err.Description
    Err.Clear
    Resume NextMsg
End Sub

' =========================
' Apple Mail Read
' =========================
Private Function GetSettingValue(ByVal namedRange As String, ByVal defaultValue As String) As String
    ' Zweck: benannten Bereich lesen, Fallback auf Default.
    ' Abhängigkeiten: ThisWorkbook.Names, Worksheets, Range.
    ' Rückgabe: String-Wert (trimmed) oder Default.
    Dim v As Variant
    Dim ws As Worksheet
    Dim found As Boolean

    On Error Resume Next
    v = ThisWorkbook.Names(namedRange).RefersToRange.Value
    If Err.Number = 0 Then found = True
    If Err.Number <> 0 Then Err.Clear
    If Not found Then
        Set ws = ThisWorkbook.Worksheets(SETTINGS_SHEET)
        If Err.Number = 0 Then
            v = ws.Range(namedRange).Value
            If Err.Number = 0 Then found = True Else Err.Clear
        Else
            Err.Clear
        End If
    End If
    On Error GoTo 0

    If Not found Then
        GetSettingValue = defaultValue
    Else
        GetSettingValue = Trim$(CStr(v))
    End If
End Function

Private Function GetLeadMailbox() As String
    GetLeadMailbox = GetSettingValue(NAME_LEAD_MAILBOX, LEAD_MAILBOX_DEFAULT)
End Function

Private Function GetLeadFolder() As String
    GetLeadFolder = GetSettingValue(NAME_LEAD_FOLDER, LEAD_FOLDER_DEFAULT)
End Function

Private Function GetMailPath() As String
    GetMailPath = CleanPathValue(GetSettingValue(NAME_MAILPATH, vbNullString))
End Function

Private Function CleanPathValue(ByVal rawValue As String) As String
    ' Zweck: Pfadwert bereinigen (führende/abschließende Hochkommas entfernen).
    Dim s As String

    s = Trim$(rawValue)
    If Len(s) = 0 Then
        CleanPathValue = s
        Exit Function
    End If

    If Left$(s, 1) = """" Or Left$(s, 1) = "'" Then
        s = Mid$(s, 2)
    End If

    If Len(s) > 0 Then
        If Right$(s, 1) = """" Or Right$(s, 1) = "'" Then
            s = Left$(s, Len(s) - 1)
        End If
    End If

    CleanPathValue = Trim$(s)
End Function

Private Function FolderExists(ByVal folderPath As String) As Boolean
    If Len(Trim$(folderPath)) = 0 Then Exit Function
    FolderExists = (Len(Dir$(folderPath, vbDirectory)) > 0)
End Function

Private Function ReadTextFile(ByVal filePath As String) As String
    Dim f As Integer
    Dim txt As String
    Dim bytes As Long

    On Error GoTo ErrHandler
    f = FreeFile
    Open filePath For Binary Access Read As #f
    bytes = LOF(f)
    If bytes > 0 Then
        txt = String$(bytes, vbNullChar)
        Get #f, , txt
    End If
    Close #f
    ReadTextFile = txt
    Exit Function

ErrHandler:
    On Error Resume Next
    Close #f
    ReadTextFile = vbNullString
End Function

Private Function NormalizeLineEndings(ByVal textIn As String) As String
    ' Zweck: Zeilenenden vereinheitlichen (CRLF/CR -> LF).
    Dim s As String
    s = Replace(textIn, vbCrLf, vbLf)
    s = Replace(s, vbCr, vbLf)
    NormalizeLineEndings = s
End Function

Private Function ExtractHeaderValue(ByVal contentText As String, ByVal headerName As String) As String
    Dim lines() As String
    Dim i As Long
    Dim lineText As String

    contentText = NormalizeLineEndings(contentText)

    lines = Split(contentText, vbLf)
    For i = LBound(lines) To UBound(lines)
        lineText = lines(i)
        If Len(lineText) = 0 Then Exit For
        If LCase$(Left$(lineText, Len(headerName))) = LCase$(headerName) Then
            ExtractHeaderValue = Trim$(Mid$(lineText, Len(headerName) + 1))
            Exit Function
        End If
    Next i
End Function

Private Function ExtractBodyFromEmail(ByVal contentText As String) As String
    Dim splitMarker As String
    Dim pos As Long

    contentText = NormalizeLineEndings(contentText)

    splitMarker = vbLf & vbLf
    pos = InStr(1, contentText, splitMarker)
    If pos > 0 Then
        ExtractBodyFromEmail = Mid$(contentText, pos + Len(splitMarker))
    Else
        ExtractBodyFromEmail = vbNullString
    End If
End Function

Private Function FetchMailMessagesFromPath(ByVal folderPath As String) As String
    ' Zweck: .eml-Dateien aus Ordner lesen und als MSG-Blocks liefern.
    Dim fileName As String
    Dim filePath As String
    Dim outText As String
    Dim rawText As String
    Dim subj As String
    Dim sender As String
    Dim dateText As String
    Dim bodyText As String
    Dim count As Long

    If Not FolderExists(folderPath) Then Exit Function

    fileName = Dir$(folderPath & "/" & "*.eml")
    Do While Len(fileName) > 0
        filePath = folderPath & "/" & fileName
        rawText = ReadTextFile(filePath)
        subj = ExtractHeaderValue(rawText, "Subject:")
        sender = ExtractHeaderValue(rawText, "From:")
        dateText = ExtractHeaderValue(rawText, "Date:")
        If Len(dateText) = 0 Then dateText = CStr(FileDateTime(filePath))
        bodyText = ExtractBodyFromEmail(rawText)

        outText = outText & MSG_DELIM & vbLf
        outText = outText & DATE_TAG & dateText & vbLf
        outText = outText & SUBJECT_TAG & subj & vbLf
        outText = outText & FROM_TAG & sender & vbLf
        outText = outText & BODY_TAG & bodyText & vbLf

        count = count + 1
        If count >= MAX_MESSAGES Then Exit Do
        fileName = Dir$()
    Loop

    FetchMailMessagesFromPath = outText
End Function

Private Function BuildExistingLeadIndex(ByVal tbl As ListObject) As Object
    ' Zweck: Index für bestehende Leads aufbauen (ID + Name/Telefon/Monat).
    ' Abhängigkeiten: NewKeyValueStore, BuildHeaderIndex, GetHeaderIndex, AddLeadKey, ExtractIdFromNotes.
    ' Rückgabe: Key/Value-Store mit bestehenden Schlüsseln.
    Dim idx As Object
    Dim headerMap As Object
    Dim notesColIndex As Long
    Dim nameColIndex As Long
    Dim phoneColIndex As Long
    Dim dateColIndex As Long
    Dim i As Long
    Dim idValue As String
    Dim nameValue As String
    Dim phoneValue As String
    Dim monthKey As String
    Dim rowDate As Date

    Set idx = NewKeyValueStore()
    If tbl Is Nothing Then
        Set BuildExistingLeadIndex = idx
        Exit Function
    End If

    If tbl.ListRows.Count = 0 Then
        Set BuildExistingLeadIndex = idx
        Exit Function
    End If

    Set headerMap = BuildHeaderIndex(tbl)
    notesColIndex = GetHeaderIndex(headerMap, "Notizen")
    nameColIndex = GetHeaderIndex(headerMap, "Name")
    phoneColIndex = GetHeaderIndex(headerMap, "Telefonnummer")
    dateColIndex = GetHeaderIndex(headerMap, "Monat Lead erhalten")

    For i = 1 To tbl.ListRows.Count
        If notesColIndex > 0 Then
            idValue = ExtractIdFromNotes(CStr(tbl.DataBodyRange.Cells(i, notesColIndex).Value))
            If Len(idValue) > 0 Then AddLeadKey idx, MakeIdKey(idValue)
        End If

        If nameColIndex > 0 And phoneColIndex > 0 And dateColIndex > 0 Then
            nameValue = CStr(tbl.DataBodyRange.Cells(i, nameColIndex).Value)
            phoneValue = CStr(tbl.DataBodyRange.Cells(i, phoneColIndex).Value)
            On Error Resume Next
            rowDate = CDate(tbl.DataBodyRange.Cells(i, dateColIndex).Value)
            On Error GoTo 0
            monthKey = MakeNamePhoneMonthKey(nameValue, phoneValue, rowDate)
            If Len(monthKey) > 0 Then AddLeadKey idx, monthKey
        End If
    Next i

    Set BuildExistingLeadIndex = idx
End Function

Private Sub AddLeadToIndex(ByVal fields As Object, ByVal msgDate As Date)
    ' Zweck: aktuellen Lead in den Index aufnehmen.
    Dim idValue As String
    Dim nameValue As String
    Dim phoneValue As String
    Dim keyVal As String

    If Not gLeadIndexInitialized Then Exit Sub

    idValue = GetField(fields, "Anfrage_ID")
    keyVal = MakeIdKey(idValue)
    If Len(keyVal) > 0 Then AddLeadKey gLeadIndex, keyVal

    nameValue = ResolveKontaktName(fields)
    phoneValue = GetField(fields, "Kontakt_Mobil")
    keyVal = MakeNamePhoneMonthKey(nameValue, phoneValue, msgDate)
    If Len(keyVal) > 0 Then AddLeadKey gLeadIndex, keyVal
End Sub

Private Sub AddLeadKey(ByRef idx As Object, ByVal keyName As String)
    If Len(keyName) > 0 Then SetKV idx, keyName, True
End Sub

Private Function MakeIdKey(ByVal idValue As String) As String
    If Len(Trim$(idValue)) > 0 Then MakeIdKey = "ID:" & LCase$(Trim$(idValue))
End Function

Private Function MakeNamePhoneMonthKey(ByVal nameValue As String, ByVal phoneValue As String, ByVal msgDate As Date) As String
    If Len(Trim$(nameValue)) = 0 Or Len(Trim$(phoneValue)) = 0 Then Exit Function
    MakeNamePhoneMonthKey = "NPM:" & LCase$(Trim$(nameValue)) & "|" & Trim$(phoneValue) & "|" & Format$(DateSerial(Year(msgDate), Month(msgDate), 1), "yyyy-mm")
End Function

Private Function ExtractIdFromNotes(ByVal noteText As String) As String
    ' Zweck: ID aus Notizen extrahieren ("ID: ...").
    Dim p As Long
    Dim tailText As String
    Dim endPos As Long

    p = InStr(1, noteText, "ID:", vbTextCompare)
    If p = 0 Then Exit Function

    tailText = Mid$(noteText, p + 3)
    If Left$(tailText, 1) = " " Then tailText = Mid$(tailText, 2)
    endPos = InStr(1, tailText, vbLf)
    If endPos > 0 Then tailText = Left$(tailText, endPos - 1)

    ExtractIdFromNotes = Trim$(tailText)
End Function

Private Function FetchAppleMailMessages(ByVal keywordA As String, ByVal keywordB As String) As String
    ' Zweck: Apple-Mail-Nachrichten per AppleScript als Text abrufen.
    ' Abhängigkeiten: AppleScriptTask, Konstanten für Tags/Delim, ParseAppleMailDate (indirekt via ParseMessageBlock später).
    ' Rückgabe: zusammengeführter Nachrichtentext oder Leerstring bei Fehler.
    ' Rückgabe: zusammengeführter Nachrichtentext oder Leerstring bei Fehler.
    Dim script As String
    Dim result As String
    Dim q As String
    Dim mailboxName As String
    Dim folderName As String
    Dim mailPath As String
    Dim pathResult As String

    q = Chr$(34)
    mailboxName = GetLeadMailbox()
    folderName = GetLeadFolder()
    mailPath = GetMailPath()

    If Len(Trim$(mailPath)) > 0 Then
        If FolderExists(mailPath) Then
            pathResult = FetchMailMessagesFromPath(mailPath)
            FetchAppleMailMessages = pathResult
            Exit Function
        Else
            MsgBox "Mailpath ungültig: " & mailPath, vbExclamation
            LogImportError "Mailpath ungültig", mailPath
        End If
    End If

    script = ""
    script = script & "with timeout of 30 seconds" & vbLf
    script = script & "tell application ""Mail""" & vbLf
    script = script & "set targetBox to missing value" & vbLf
    If Len(mailboxName) > 0 Then
        script = script & "set targetAccountName to " & q & mailboxName & q & vbLf
        script = script & "try" & vbLf
        script = script & "repeat with a in accounts" & vbLf
        script = script & "if (name of a) contains targetAccountName then" & vbLf
        script = script & "try" & vbLf
        script = script & "set targetBox to first mailbox of a whose name contains " & q & folderName & q & vbLf
        script = script & "exit repeat" & vbLf
        script = script & "end try" & vbLf
        script = script & "end if" & vbLf
        script = script & "end repeat" & vbLf
        script = script & "end try" & vbLf
    End If
    script = script & "if targetBox is missing value then" & vbLf
    script = script & "try" & vbLf
    script = script & "set targetBox to first mailbox whose name contains " & q & folderName & q & vbLf
    script = script & "end try" & vbLf
    script = script & "end if" & vbLf
    script = script & "if targetBox is missing value then error ""Mailbox nicht gefunden: " & folderName & """" & vbLf
    script = script & "set theMessages to (every message of targetBox whose subject contains """ & keywordA & """ or subject contains """ & keywordB & """ or content contains """ & keywordA & """ or content contains """ & keywordB & """ )" & vbLf
    script = script & "if (count of theMessages) > " & MAX_MESSAGES & " then set theMessages to items 1 thru " & MAX_MESSAGES & " of theMessages" & vbLf
    script = script & "set outText to """"" & vbLf
    script = script & "repeat with m in theMessages" & vbLf
        script = script & "set outText to outText & """ & MSG_DELIM & """ & linefeed" & vbLf
        script = script & "set outText to outText & """ & DATE_TAG & """ & (date sent of m) & linefeed" & vbLf
        script = script & "set outText to outText & """ & SUBJECT_TAG & """ & (subject of m) & linefeed" & vbLf
        script = script & "set outText to outText & """ & FROM_TAG & """ & (sender of m) & linefeed" & vbLf
            script = script & "set bodyText to (content of m)" & vbLf
        script = script & "set outText to outText & " & q & BODY_TAG & q & " & bodyText & linefeed" & vbLf
    script = script & "end repeat" & vbLf
    script = script & "return outText" & vbLf
    script = script & "end tell" & vbLf
    script = script & "end timeout"

    On Error GoTo ErrHandler
    result = AppleScriptTask(APPLESCRIPT_FILE, APPLESCRIPT_HANDLER, script)
    If Left$(result, 6) = "ERROR:" Then
        MsgBox "AppleScript-Fehler: " & Mid$(result, 7), vbExclamation
        LogImportError "AppleScript-Fehler", Mid$(result, 7)
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
    LogImportError "AppleScriptTask-Fehler", Err.Description
    FetchAppleMailMessages = vbNullString
End Function

Public Sub DebugPrintAppleMailFolders()
    ' Zweck: Mailbox-Ordnerstruktur im Direktfenster ausgeben.
    ' Abhängigkeiten: FetchAppleMailFolderList.
    ' Rückgabe: keine (Debug.Print Ausgabe).
    Dim folderText As String
    Dim lines() As String
    Dim i As Long

    folderText = FetchAppleMailFolderList()
    If Len(folderText) = 0 Then Exit Sub

    lines = Split(folderText, vbLf)
    For i = LBound(lines) To UBound(lines)
        ' Schleife: jede Zeile der Ordnerliste ausgeben.
        If Len(Trim$(lines(i))) > 0 Then
            Debug.Print Trim$(lines(i))
        End If
    Next i
End Sub

Private Function FetchAppleMailFolderList() As String
    ' Zweck: Ordnerliste aus Apple Mail via AppleScript abrufen.
    ' Abhängigkeiten: AppleScriptTask.
    ' Rückgabe: Textliste der Ordner oder Leerstring.
    ' Rückgabe: Textliste der Ordner (eine Zeile pro Ordner) oder Leerstring bei Fehler.
    Dim script As String
    Dim result As String
    Dim q As String
    Dim mailboxName As String

    q = Chr$(34)
    mailboxName = GetLeadMailbox()

    script = ""
    script = script & "with timeout of 30 seconds" & vbLf
    script = script & "tell application ""Mail""" & vbLf
    script = script & "set targetAccountName to " & q & mailboxName & q & vbLf
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
        LogImportError "AppleScript-Fehler", Mid$(result, 7)
        FetchAppleMailFolderList = vbNullString
        Exit Function
    End If
    FetchAppleMailFolderList = result
    Exit Function

ErrHandler:
    MsgBox "AppleScriptTask-Fehler. Prüfe Script-Installation und Automation-Rechte.", vbExclamation
    LogImportError "AppleScriptTask-Fehler", Err.Description
    FetchAppleMailFolderList = vbNullString
End Function

' =========================
' AppleScript Setup
' =========================
Private Sub EnsureAppleScriptInstalled()
    ' Zweck: AppleScript ins Zielverzeichnis kopieren, falls es fehlt.
    ' Abhängigkeiten: GetAppleScriptTargetPath, InstallAppleScript.
    ' Rückgabe: keine (Seiteneffekt Datei-Kopie).
    Dim targetPath As String
    Dim sourcePath As String

    targetPath = GetAppleScriptTargetPath()
    sourcePath = ThisWorkbook.Path & "/" & APPLESCRIPT_SOURCE

    If Len(Dir$(targetPath)) = 0 Then
        InstallAppleScript sourcePath, targetPath
    End If
End Sub

Private Function GetAppleScriptTargetPath() As String
    ' Zweck: Zielpfad für das AppleScript ermitteln.
    ' Abhängigkeiten: Environ$ HOME.
    ' Rückgabe: Vollständiger Pfad.
    ' Rückgabe: Vollständiger Pfad zur scpt-Datei.
    Dim homePath As String
    homePath = Environ$("HOME")
    GetAppleScriptTargetPath = homePath & "/Library/Application Scripts/com.microsoft.Excel/" & APPLESCRIPT_FILE
End Function

Private Sub InstallAppleScript(ByVal sourcePath As String, ByVal targetPath As String)
    ' Zweck: AppleScript aus dem Projektverzeichnis installieren.
    ' Abhängigkeiten: EnsureFolderExists, FileCopy, Kill.
    ' Rückgabe: keine (kopiert Datei oder zeigt MsgBox).
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
    ' Zweck: Zielordner rekursiv anlegen, falls nicht vorhanden.
    ' Abhängigkeiten: MkDir, Dir$.
    ' Rückgabe: keine (stellt Ordner sicher bereit).
    Dim parts() As String
    Dim i As Long
    Dim currentPath As String

    parts = Split(folderPath, "/")
    currentPath = ""

    For i = LBound(parts) To UBound(parts)
        ' Schleife: jeden Pfadteil prüfen und ggf. anlegen.
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
    ' Zweck: Schlüssel/Wert-Store passend zum OS erstellen.
    ' Abhängigkeiten: Application.OperatingSystem, Collection/Scripting.Dictionary.
    ' Rückgabe: Collection (Mac) oder Dictionary (Windows).
    ' Rückgabe: Dictionary (Windows) oder Collection (macOS).
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

Private Function keyNorm(ByVal keyName As String) As String
    ' Zweck: Schlüssel vereinheitlichen (trim + lowercase).
    ' Abhängigkeiten: Trim$, LCase$.
    ' Rückgabe: normalisierter Schlüsselstring.
    ' Rückgabe: Normalisierter Schlüssel.
    keyNorm = LCase$(Trim$(keyName))
End Function

Private Sub SetKV(ByRef store As Object, ByVal keyName As String, ByVal valueToSet As Variant)
    ' Zweck: Wert im Store setzen (Dictionary/Collection abstrahiert).
    ' Abhängigkeiten: keyNorm, Collection/Dictionary API.
    ' Rückgabe: keine (mutiert Store).
    Dim k As String
    k = keyNorm(keyName)

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
    ' Zweck: Wert sicher aus dem Store lesen.
    ' Abhängigkeiten: keyNorm, Collection/Dictionary API.
    ' Rückgabe: True bei Treffer, sonst False.
    ' Rückgabe: True bei Treffer, sonst False.
    Dim k As String
    k = keyNorm(keyName)

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
    ' Zweck: Datum/Betreff/Body aus einem Message-Block extrahieren.
    ' Abhängigkeiten: NewKeyValueStore, ParseAppleMailDate, SetKV, TryGetKV.
    ' Rückgabe: Key/Value-Store mit "Date", "Subject", "Body", "From".
    ' Rückgabe: Key/Value-Store mit "Date", "Subject", "Body".
    Dim lines() As String
    Dim i As Long
    Dim lineText As String
    Dim payload As Object

    Set payload = NewKeyValueStore()
    SetKV payload, "Date", Date
    SetKV payload, "Subject", vbNullString
    SetKV payload, "From", vbNullString
    SetKV payload, "Body", vbNullString

    lines = Split(blockText, vbLf)
    For i = LBound(lines) To UBound(lines)
        ' Schleife: jede Zeile des Message-Blocks auswerten.
        lineText = Trim$(lines(i))
        If Len(lineText) > 0 Then
            If Left$(lineText, Len(DATE_TAG)) = DATE_TAG Then
                SetKV payload, "Date", ParseAppleMailDate(Trim$(Mid$(lineText, Len(DATE_TAG) + 1)))
            ElseIf Left$(lineText, Len(SUBJECT_TAG)) = SUBJECT_TAG Then
                SetKV payload, "Subject", Trim$(Mid$(lineText, Len(SUBJECT_TAG) + 1))
            ElseIf Left$(lineText, Len(FROM_TAG)) = FROM_TAG Then
                SetKV payload, "From", Trim$(Mid$(lineText, Len(FROM_TAG) + 1))
            ElseIf Left$(lineText, Len(BODY_TAG)) = BODY_TAG Then
                SetKV payload, "Body", Trim$(Mid$(lineText, Len(BODY_TAG) + 1)) & vbLf
            Else
                Dim curBody As Variant
                If Not TryGetKV(payload, "Body", curBody) Then curBody = vbNullString
                SetKV payload, "Body", CStr(curBody) & lineText & vbLf
            End If
        End If
    Next i

    Set ParseMessageBlock = payload
End Function

Private Function ParseAppleMailDate(ByVal dateText As String) As Date
    ' Zweck: Apple-Mail-Datumstext robust in Date konvertieren.
    ' Abhängigkeiten: GermanMonthToNumber, CDate, DateSerial/TimeSerial.
    ' Rückgabe: Datum (Fallback: Today).
    ' Rückgabe: VBA-Date (Fallback: Heute).
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
    ' Zweck: deutschen Monatsnamen in Monatszahl wandeln.
    ' Abhängigkeiten: keine externen; verwendet Select Case.
    ' Rückgabe: Monatszahl 1-12 (Fallback 1).
    ' Rückgabe: 1-12 (Fallback: 1).
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
    ' Zweck: Lead-Typ anhand Betreff/Inhalt bestimmen.
    ' Abhängigkeiten: String-Suche InStr.
    ' Rückgabe: KEYWORD_1 oder KEYWORD_2.
    ' Rückgabe: KEYWORD_1 oder KEYWORD_2.
    If InStr(1, subjectText, KEYWORD_2, vbTextCompare) > 0 Or InStr(1, bodyText, KEYWORD_2, vbTextCompare) > 0 Then
        ResolveLeadType = KEYWORD_2
    Else
        ResolveLeadType = KEYWORD_1
    End If
End Function

Private Function ParseLeadContent(ByVal bodyText As String) As Object
    ' Zweck: Nachrichtentext in strukturierte Felder parsen.
    ' Abhängigkeiten: NewKeyValueStore, MapLabelValue, MapInlinePair, SetBedarfsort.
    ' Rückgabe: Key/Value-Store mit Feldwerten.
    ' Rückgabe: Key/Value-Store mit den erkannten Feldern.
    Dim result As Object
    Dim lines() As String
    Dim i As Long
    Dim lineText As String
    Dim currentSection As String
    Dim pendingKey As String
    Dim workText As String
    Dim posSenior As Long

    Set result = NewKeyValueStore()

    currentSection = "Kontakt"
    pendingKey = vbNullString

    workText = bodyText
    posSenior = InStr(1, workText, "Informationen zum Senior", vbTextCompare)
    If posSenior > 0 Then workText = Mid$(workText, posSenior)

    lines = Split(workText, vbLf)
    For i = LBound(lines) To UBound(lines)
        ' Schleife: Zeilen iterieren und Abschnitt/Felder erkennen.
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
    ' Zweck: Inline-Label/Value ("key: value") in Felder mappen.
    ' Abhängigkeiten: MapLabelValue.
    ' Rückgabe: keine (schreibt in fields).
    Dim keyPart As String
    Dim valuePart As String

    keyPart = Trim$(Left$(lineText, InStr(lineText, ":") - 1))
    valuePart = Trim$(Mid$(lineText, InStr(lineText, ":") + 1))

    If Len(keyPart) > 0 Then
        MapLabelValue fields, keyPart, valuePart, sectionName
    End If
End Sub

Private Sub MapLabelValue(ByRef fields As Object, ByVal rawKey As String, ByVal rawValue As String, ByVal sectionName As String)
    ' Zweck: Normalisierten Schlüssel auf Zielspalte mappen.
    ' Abhängigkeiten: NormalizeKey, SetBedarfsort, SetKV.
    ' Rückgabe: keine (schreibt in fields).
    Dim keyNorm As String
    Dim valueNorm As String

    keyNorm = NormalizeKey(rawKey)
    valueNorm = Trim$(rawValue)

    Select Case keyNorm
        Case "anrede": SetKV fields, "Kontakt_Anrede", valueNorm
        Case "vorname": SetKV fields, "Kontakt_Vorname", valueNorm
        Case "nachname": SetKV fields, "Kontakt_Nachname", valueNorm
        Case "vor- und nachname", "vor und nachname": SetKV fields, "Kontakt_Name", valueNorm
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
        Case "pflegegrad", "pflegegrad/-stufe": SetKV fields, "Senior_Pflegegrad", valueNorm
        Case "lebenssituation": SetKV fields, "Senior_Lebenssituation", valueNorm
        Case "mobilität": SetKV fields, "Senior_Mobilitaet", valueNorm
        Case "medizinisches": SetKV fields, "Senior_Medizinisches", valueNorm
        Case "behinderung": SetKV fields, "Senior_Behinderung", valueNorm
        Case "postleitzahl", "plz": SetKV fields, "PLZ", valueNorm
        Case "bedarfsort": SetBedarfsort fields, valueNorm
        Case "nutzer": SetKV fields, "Nutzer", valueNorm
        Case "alltagshilfe aufgaben": SetKV fields, "Alltagshilfe_Aufgaben", valueNorm
        Case "alltagshilfe häufigkeit": SetKV fields, "Alltagshilfe_Haeufigkeit", valueNorm
        Case "aufgaben": SetKV fields, "Aufgaben", valueNorm
        Case "wöchentlicher umfang": SetKV fields, "Woechentlicher_Umfang", valueNorm
        Case "umfang am stück", "umfang am stueck": SetKV fields, "Umfang_am_Stueck", valueNorm
        Case "abrechnung über bet.- & entlastungsleistungen", "abrechnung ueber bet.- & entlastungsleistungen": SetKV fields, "Abrechnung_Betreuungsleistungen", valueNorm
        Case "pflegedienst vorhanden": SetKV fields, "Pflegedienst_Vorhanden", valueNorm
        Case "anfragedetails": SetKV fields, "Anfragedetails", valueNorm
        Case "anfragen-nr", "anfragen-nr.", "anfragen nr": SetKV fields, "Anfrage_ID", valueNorm
        Case "weitere details": SetKV fields, "Weitere_Details", valueNorm
        Case "bedarf": SetKV fields, "Bedarf", valueNorm
        Case "anfragedetails": SetKV fields, "Anfragedetails", valueNorm
        Case "anfragen-nr:": SetKV fields, "Anfrage_ID", valueNorm
        Case "id": SetKV fields, "Anfrage_ID", valueNorm
    End Select
End Sub

Private Function NormalizeKey(ByVal rawKey As String) As String
    ' Zweck: Schlüsseltext vereinheitlichen.
    ' Abhängigkeiten: Trim$, Replace, LCase$.
    ' Rückgabe: normalisierter Label-Text.
    ' Rückgabe: normalisierte Zeichenkette.
    Dim k As String
    k = LCase$(Trim$(rawKey))
    k = Replace(k, vbTab, " ")
    k = Replace(k, "  ", " ")
    NormalizeKey = k
End Function

Private Function GetCellByHeaderMap(ByVal rowItem As ListRow, ByVal headerMap As Object, ByVal headerName As String) As Range
    ' Zweck: Zellobjekt anhand Header-Map holen.
    ' Abhängigkeiten: GetHeaderIndex.
    ' Rückgabe: Range oder Nothing.
    Dim idx As Long
    idx = GetHeaderIndex(headerMap, headerName)
    If idx > 0 Then Set GetCellByHeaderMap = rowItem.Range.Cells(1, idx)
End Function

Private Sub SetImportNote(ByVal targetCell As Range)
    ' Zweck: Import-Notiz an Zelle setzen.
    Dim noteText As String

    If targetCell Is Nothing Then Exit Sub
    noteText = "Automatischer Import vom: " & Format$(Now, "dd.mm.yy hh.nn") & " | Quelle: " & LEAD_SOURCE

    On Error Resume Next
    If Not targetCell.Comment Is Nothing Then targetCell.Comment.Delete
    targetCell.AddComment noteText
    If Not targetCell.Comment Is Nothing Then targetCell.Comment.Visible = False
    On Error GoTo 0
End Sub

Private Sub LogImportError(ByVal errMessage As String, ByVal possibleCause As String)
    ' Zweck: Fehler in ErrLog protokollieren.
    Dim ws As Worksheet
    Dim nextRow As Long

    Set ws = GetOrCreateErrorLogSheet()
    If ws Is Nothing Then Exit Sub

    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If nextRow = 1 And Len(Trim$(CStr(ws.Cells(1, 1).Value))) = 0 Then nextRow = 0
    nextRow = nextRow + 1

    ws.Cells(nextRow, 1).Value = Format$(Now, "dd.mm.yy hh.nn")
    ws.Cells(nextRow, 2).Value = errMessage
    ws.Cells(nextRow, 3).Value = possibleCause
End Sub

Private Function GetOrCreateErrorLogSheet() As Worksheet
    ' Zweck: ErrLog-Sheet holen oder anlegen.
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(ERROR_LOG_SHEET)
    On Error GoTo 0

    If ws Is Nothing Then
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        If Not ws Is Nothing Then
            ws.Name = ERROR_LOG_SHEET
            ws.Cells(1, 1).Value = "Zeitstempel"
            ws.Cells(1, 2).Value = "Fehler"
            ws.Cells(1, 3).Value = "Mögliche Ursache"
        End If
        On Error GoTo 0
    End If

    Set GetOrCreateErrorLogSheet = ws
End Function

' =========================
' Excel Output
' =========================
Private Sub AddLeadRow(ByVal tbl As ListObject, ByVal fields As Object, ByVal msgDate As Date, ByVal leadType As String)
    ' Zweck: neue Tabellenzeile mit Lead-Daten anlegen.
    ' Abhängigkeiten: BuildHeaderIndex, SetCellByHeaderMap, ResolveLeadSource, ResolveKontaktName, NormalizePflegegrad, BuildNotes.
    ' Rückgabe: keine (fügt Zeile hinzu).
    Dim newRow As ListRow
    Dim monthCell As Range

    Set newRow = tbl.ListRows.Add

    Dim headerMap As Object
    Set headerMap = BuildHeaderIndex(tbl)

    SetCellByHeaderMap newRow, headerMap, "Monat Lead erhalten", DateSerial(Year(msgDate), Month(msgDate), 1)
    Set monthCell = GetCellByHeaderMap(newRow, headerMap, "Monat Lead erhalten")
    SetImportNote monthCell
    SetCellByHeaderMap newRow, headerMap, "Lead-Quelle", ResolveLeadSource(fields)
    SetCellByHeaderMap newRow, headerMap, "Leadtyp", leadType
    SetCellByHeaderMap newRow, headerMap, "Status", "Lead erhalten"
    SetCellByHeaderMap newRow, headerMap, "Name", ResolveKontaktName(fields)
    SetCellByHeaderMap newRow, headerMap, "Telefonnummer", GetField(fields, "Kontakt_Mobil")
    SetCellByHeaderMap newRow, headerMap, "PLZ", GetField(fields, "PLZ")
    SetCellByHeaderMap newRow, headerMap, "PG", NormalizePflegegrad(GetField(fields, "Senior_Pflegegrad"))
    SetCellByHeaderMap newRow, headerMap, "Notizen", BuildNotes(fields)
End Sub

Private Function FindTableByName(ByVal tableName As String) As ListObject
    ' Zweck: ListObject global nach Name suchen. Rückgabe: gefundene Tabelle oder Nothing.
    ' Abhängigkeiten: ThisWorkbook.Worksheets, ListObjects.
    ' Rückgabe: ListObject oder Nothing.
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        Dim lo As ListObject
        For Each lo In ws.ListObjects
            If StrComp(lo.Name, tableName, vbTextCompare) = 0 Then
                Set FindTableByName = lo
                Exit Function
            End If
        Next lo
    Next ws
End Function

Private Function ResolveLeadSource(ByVal fields As Object) As String
    ' Zweck: Lead-Quelle aus Absender nutzen, Fallback auf Default.
    ' Abhängigkeiten: GetField.
    ' Rückgabe: Absender oder LEAD_SOURCE.
    Dim fromVal As String
    fromVal = GetField(fields, "From")
    If Len(Trim$(fromVal)) = 0 Then
        ResolveLeadSource = LEAD_SOURCE
    Else
        ResolveLeadSource = fromVal
    End If
End Function

Private Sub SetCellByHeaderMap(ByVal rowItem As ListRow, ByVal headerMap As Object, ByVal headerName As String, ByVal valueToSet As Variant)
    ' Zweck: Zellwert anhand vorberechneter Header-Map setzen.
    ' Abhängigkeiten: GetHeaderIndex.
    ' Rückgabe: keine (schreibt in Zeile).
    Dim idx As Long
    idx = GetHeaderIndex(headerMap, headerName)
    If idx > 0 Then rowItem.Range.Cells(1, idx).Value = valueToSet
End Sub

Private Function BuildHeaderIndex(ByVal tbl As ListObject) As Object
    ' Zweck: Map Headername -> Spaltenindex erzeugen. Rückgabe: Store mit Indizes.
    ' Abhängigkeiten: NewKeyValueStore, SetKV.
    ' Rückgabe: Key/Value-Store Header->Index.
    Dim map As Object
    Dim i As Long
    Set map = NewKeyValueStore()
    For i = 1 To tbl.ListColumns.Count
        SetKV map, Trim$(tbl.ListColumns(i).Name), i
    Next i
    Set BuildHeaderIndex = map
End Function

Private Function GetHeaderIndex(ByVal headerMap As Object, ByVal headerName As String) As Long
    ' Zweck: Spaltenindex aus Map lesen. Rückgabe: Index oder 0.
    ' Abhängigkeiten: TryGetKV.
    ' Rückgabe: Spaltenindex oder 0.
    Dim v As Variant
    If TryGetKV(headerMap, headerName, v) Then GetHeaderIndex = CLng(v) Else GetHeaderIndex = 0
End Function

Private Function ResolveKontaktName(ByVal fields As Object) As String
    ' Zweck: Kontaktname aus Feldern zusammenstellen.
    ' Abhängigkeiten: GetField, ExtractSenderName.
    ' Rückgabe: vollqualifizierter Name oder Leerstring.
    ' Rückgabe: Vollständiger Name (ggf. leer).
    Dim fullName As String
    fullName = GetField(fields, "Kontakt_Name")

    If Len(fullName) = 0 Then
        fullName = Trim$(GetField(fields, "Kontakt_Vorname") & " " & GetField(fields, "Kontakt_Nachname"))
    End If

    If Len(fullName) = 0 Then
        fullName = GetField(fields, "Senior_Name")
    End If

    If Len(fullName) = 0 Then
        fullName = ExtractSenderName(GetField(fields, "From"))
    End If

    fullName = Trim$(fullName)
    ResolveKontaktName = fullName
End Function

Private Function ExtractSenderName(ByVal fromVal As String) As String
    ' Zweck: Namensteil aus From-Header extrahieren.
    ' Abhängigkeiten: Stringfunktionen (InStr, Left$, Replace).
    ' Rückgabe: gereinigter Name.
    Dim s As String
    s = Trim$(fromVal)
    If Len(s) = 0 Then Exit Function

    If InStr(s, "<") > 0 Then
        s = Trim$(Left$(s, InStr(s, "<") - 1))
    End If

    s = Replace(s, """", "")
    ExtractSenderName = Trim$(s)
End Function

Private Function GetField(ByVal fields As Object, ByVal keyName As String) As String
    ' Zweck: Feldwert sicher lesen.
    ' Abhängigkeiten: TryGetKV.
    ' Rückgabe: Feldinhalt oder Leerstring.
    ' Rückgabe: Feldinhalt oder Leerstring.
    Dim v As Variant
    If TryGetKV(fields, keyName, v) Then
        GetField = CStr(v)
    Else
        GetField = vbNullString
    End If
End Function

Private Function BuildNotes(ByVal fields As Object) As String
    ' Zweck: Notizentext aus optionalen Feldern aufbauen.
    ' Abhängigkeiten: AppendNote, GetField.
    ' Rückgabe: zusammengesetzter Notiztext.
    ' Rückgabe: zusammengesetzter Notiztext.
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
    notes = AppendNote(notes, "Behinderung", GetField(fields, "Senior_Behinderung"))
    notes = AppendNote(notes, "Nutzer", GetField(fields, "Nutzer"))
    notes = AppendNote(notes, "Alltagshilfe Aufgaben", GetField(fields, "Alltagshilfe_Aufgaben"))
    notes = AppendNote(notes, "Alltagshilfe Häufigkeit", GetField(fields, "Alltagshilfe_Haeufigkeit"))
    notes = AppendNote(notes, "Aufgaben", GetField(fields, "Aufgaben"))
    notes = AppendNote(notes, "Wöchentlicher Umfang", GetField(fields, "Woechentlicher_Umfang"))
    notes = AppendNote(notes, "Umfang am Stück", GetField(fields, "Umfang_am_Stueck"))
    notes = AppendNote(notes, "Abrechnung über Bet.- & Entlastungsleistungen", GetField(fields, "Abrechnung_Betreuungsleistungen"))
    notes = AppendNote(notes, "Pflegedienst vorhanden", GetField(fields, "Pflegedienst_Vorhanden"))
    notes = AppendNote(notes, "Anfragedetails", GetField(fields, "Anfragedetails"))
    notes = AppendNote(notes, "Weitere Details", GetField(fields, "Weitere_Details"))
    notes = AppendNote(notes, "Bedarf", GetField(fields, "Bedarf"))
    notes = AppendNote(notes, "Bedarfsort Ort", GetField(fields, "Bedarfsort_Ort"))
    notes = AppendNote(notes, "ID", GetField(fields, "Anfrage_ID"))

    BuildNotes = notes
End Function

Private Sub SetBedarfsort(ByRef fields As Object, ByVal rawValue As String)
    ' Zweck: Bedarfsort in PLZ/Ort trennen.
    ' Abhängigkeiten: FilterDigits, SetKV.
    ' Rückgabe: keine (schreibt Felder).
    Dim tokens() As String
    Dim i As Long
    Dim plzToken As String
    Dim ortPart As String
    Dim t As String

    tokens = Split(Trim$(rawValue), " ")
    For i = LBound(tokens) To UBound(tokens)
        t = Trim$(tokens(i))
        If Len(t) >= 4 And Len(FilterDigits(t)) >= 4 And Len(FilterDigits(t)) <= 5 Then
            plzToken = FilterDigits(t)
            tokens(i) = ""
            Exit For
        End If
    Next i

    If Len(plzToken) > 0 Then
        SetKV fields, "PLZ", plzToken
    End If

    ortPart = Trim$(Join(tokens, " "))
    If Len(ortPart) > 0 Then
        SetKV fields, "Bedarfsort_Ort", ortPart
    End If
End Sub

Private Function FilterDigits(ByVal textIn As String) As String
    ' Zweck: Nur Ziffern aus Text extrahieren.
    ' Abhängigkeiten: Stringzugriff.
    ' Rückgabe: Ziffernfolge oder Leerstring.
    Dim i As Long
    Dim digits As String
    For i = 1 To Len(textIn)
        If Mid$(textIn, i, 1) >= "0" And Mid$(textIn, i, 1) <= "9" Then
            digits = digits & Mid$(textIn, i, 1)
        End If
    Next i
    FilterDigits = digits
End Function

Private Function NormalizePflegegrad(ByVal rawValue As String) As String
    ' Zweck: Pflegegrad auf reine Ziffern normalisieren.
    ' Abhängigkeiten: Stringzugriff.
    ' Rückgabe: Ziffernfolge oder Leerstring.
    Dim i As Long
    Dim digits As String

    For i = 1 To Len(rawValue)
        If Mid$(rawValue, i, 1) >= "0" And Mid$(rawValue, i, 1) <= "9" Then
            digits = digits & Mid$(rawValue, i, 1)
        End If
    Next i

    If Len(digits) > 0 Then
        NormalizePflegegrad = digits
    Else
        NormalizePflegegrad = vbNullString
    End If
End Function

Private Function AppendNote(ByVal currentText As String, ByVal labelText As String, ByVal valueText As String) As String
    ' Zweck: Notizfeld anhängen, wenn ein Wert vorhanden ist.
    ' Abhängigkeiten: keine externen.
    ' Rückgabe: aktualisierter Notiztext.
    ' Rückgabe: aktualisierter Notiztext.
    If Len(Trim$(valueText)) = 0 Then
        AppendNote = currentText
    ElseIf Len(currentText) = 0 Then
        AppendNote = labelText & ": " & valueText
    Else
        AppendNote = currentText & vbLf & labelText & ": " & valueText
    End If
End Function

' =========================
' Duplicate Handling
' =========================
Private Function LeadAlreadyExists(ByVal tbl As ListObject, ByVal fields As Object, ByVal msgDate As Date) As Boolean
    ' Zweck: Duplikate anhand ID oder Name+Telefon+Monat verhindern.
    ' Abhängigkeiten: ResolveKontaktName, GetField, BuildHeaderIndex, GetHeaderIndex, DateSerial.
    ' Rückgabe: True bei Duplikat sonst False.
    ' Rückgabe: True bei Treffer, sonst False.
    Dim idValue As String
    Dim nameValue As String
    Dim phoneValue As String
    Dim keyVal As String
    Dim v As Variant
    Dim headerMap As Object
    Dim notesColIndex As Long
    Dim nameColIndex As Long
    Dim phoneColIndex As Long
    Dim dateColIndex As Long
    Dim i As Long

    idValue = GetField(fields, "Anfrage_ID")
    nameValue = ResolveKontaktName(fields)
    phoneValue = GetField(fields, "Kontakt_Mobil")

    If gLeadIndexInitialized Then
        keyVal = MakeIdKey(idValue)
        If Len(keyVal) > 0 Then
            If TryGetKV(gLeadIndex, keyVal, v) Then
                LeadAlreadyExists = True
                Exit Function
            End If
        End If

        keyVal = MakeNamePhoneMonthKey(nameValue, phoneValue, msgDate)
        If Len(keyVal) > 0 Then
            If TryGetKV(gLeadIndex, keyVal, v) Then
                LeadAlreadyExists = True
                Exit Function
            End If
        End If
    End If

    If tbl.ListRows.Count = 0 Then Exit Function

    Set headerMap = BuildHeaderIndex(tbl)
    notesColIndex = GetHeaderIndex(headerMap, "Notizen")
    nameColIndex = GetHeaderIndex(headerMap, "Name")
    phoneColIndex = GetHeaderIndex(headerMap, "Telefonnummer")
    dateColIndex = GetHeaderIndex(headerMap, "Monat Lead erhalten")

    For i = 1 To tbl.ListRows.Count
        ' Schleife: bestehende Zeilen auf Duplikate prüfen.
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


