Option Explicit

' =========================
' Code Description
' =========================
' Dieses Modul liest Apple Mail oder Outlook Nachrichten mit den Schlagworten "Lead" oder
' "Neue Anfrage", parst die Inhalte und schreibt die Daten in die intelligente
' Tabelle "Kundenliste" auf dem Blatt "Pipeline".
' Mail-App wird automatisch erkannt: LEAD_MAILBOX mit "@" -> Outlook, sonst Apple Mail.
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

' Zielordner in Apple Mail / Outlook
' LEAD_FOLDER muss exakter Ordnername sein (z. B. "Archiv", "Leads", "Posteingang")
' LEAD_MAILBOX bestimmt automatisch die Mail-App:
'   - Enthält "@" oder "outlook" oder "exchange" -> Microsoft Outlook
'   - Sonst (z.B. "iCloud") -> Apple Mail
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
Private gLeadSourceNote As String

' =========================
' Funktionsübersicht & Abhängigkeiten
' =========================
' ImportLeadsFromAppleMail: Einstiegspunkt; ruft FetchAppleMailMessages, ParseMessageBlock, ParseLeadContent, LeadAlreadyExists, AddLeadRow. Rückgabe: Sub, schreibt Zeilen.
' FetchAppleMailMessages: Baut AppleScript (Apple Mail oder Outlook), ruft AppleScriptTask; liefert zusammengefasste Roh-Nachrichten.
' BuildAppleMailScript: Generiert AppleScript für Apple Mail.
' BuildOutlookScript: Generiert AppleScript für Microsoft Outlook.
' IsOutlookMailbox: Erkennt pro Mailbox-Eintrag ob Outlook oder Apple Mail (@ -> Outlook).
' DebugPrintAppleMailFolders: Debug-Ausgabe der Ordner; nutzt FetchAppleMailFolderList.
' FetchAppleMailFolderList: Baut AppleScript, ruft AppleScriptTask; liefert Ordnerliste als Text.
' EnsureAppleScriptInstalled / GetAppleScriptTargetPath / InstallAppleScript / EnsureFolderExists: Helfer zum Installieren des AppleScripts. Rückgabe: Pfade oder Seiteneffekt.
' NewKeyValueStore / keyNorm / SetKV / TryGetKV: Plattform-sicherer Key/Value-Store. Rückgabe: Collection/Dictionary oder Boolean.
' ParseMessageBlock: Zerlegt einen Nachrichtenblock in Date/Subject/From/Body; nutzt ParseAppleMailDate. Nach BODY:-Tag werden alle Folgezeilen dem Body zugeordnet.
' ParseAppleMailDate / GermanMonthToNumber: Robust Datum parsen aus Apple Mail Text.
' IsLikelyBase64: Erkennt ob Text reines Base64 ist. Rückgabe: Boolean.
' DecodeBodyIfNeeded: Dekodiert Body automatisch bei Base64/MIME-Kodierung; nutzt IsLikelyBase64, DecodeBase64ToString, ExtractBodyFromEmail.
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
'   -> FetchAppleMailMessages -> BuildAppleMailScript / BuildOutlookScript -> AppleScriptTask
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
Private Function ValidateMailSettings() As Boolean
    ' Zweck: Prüft ob mindestens eine Mail-Quelle konfiguriert ist.
    ' Zeigt MsgBox und springt zum Einstellungs-Sheet wenn nicht.
    ' Rückgabe: True wenn OK, False wenn Einstellungen fehlen.
    Dim mailbox As String
    Dim folder As String
    Dim mailPath As String
    Dim missingFields As String

    mailbox = Trim$(GetSettingValue(NAME_LEAD_MAILBOX, vbNullString))
    folder = Trim$(GetSettingValue(NAME_LEAD_FOLDER, vbNullString))
    mailPath = Trim$(GetMailPath())

    ' Mindestens mailpath ODER (mailbox + folder) muss gesetzt sein
    If Len(mailPath) > 0 Then
        ' mailpath ist gesetzt -> OK auch ohne mailbox/folder
        ValidateMailSettings = True
        Exit Function
    End If

    ' Kein mailpath -> mailbox und folder müssen gesetzt sein
    If Len(mailbox) = 0 Then
        missingFields = missingFields & "  - LEAD_MAILBOX (Account-Name oder E-Mail)" & vbLf
    End If
    If Len(folder) = 0 Then
        missingFields = missingFields & "  - LEAD_FOLDER (Ordnername, z.B. Leads oder Posteingang)" & vbLf
    End If

    If Len(missingFields) > 0 Then
        MsgBox "Import nicht möglich – fehlende Einstellungen:" & vbLf & vbLf & _
               missingFields & vbLf & _
               "Bitte auf dem Blatt '" & SETTINGS_SHEET & "' ergänzen." & vbLf & _
               "Alternativ kann 'mailpath' als lokaler Ordnerpfad gesetzt werden.", _
               vbExclamation, "Fehlende Mail-Konfiguration"
        GoToSettingsSheet
        ValidateMailSettings = False
        Exit Function
    End If

    ValidateMailSettings = True
End Function

Private Sub GoToSettingsSheet()
    ' Zweck: Zum Einstellungs-Sheet springen und erste relevante Zelle aktivieren.
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SETTINGS_SHEET)
    On Error GoTo 0

    If ws Is Nothing Then
        MsgBox "Blatt '" & SETTINGS_SHEET & "' nicht gefunden.", vbExclamation
        Exit Sub
    End If

    ws.Activate
    On Error Resume Next
    ws.Range("A1").Select
    On Error GoTo 0
End Sub

Public Sub ImportLeadsFromAppleMail()
    ' Zweck: Apple-Mail-Leads abrufen, parsen und in die Tabelle schreiben.
    ' Abhängigkeiten: EnsureAppleScriptInstalled (optional), FetchAppleMailMessages, ParseMessageBlock, ResolveLeadType, ParseLeadContent, LeadAlreadyExists, AddLeadRow.
    ' Rückgabe: keine (fügt Zeilen in Tabelle ein).

    Debug.Print "[Main] === Version: 2026-02-16-section-umlaut-fix ==="

    ' --- Eingabeprüfung ---
    If Not ValidateMailSettings() Then Exit Sub

    If AUTO_INSTALL_APPLESCRIPT Then
        EnsureAppleScriptInstalled
    End If

    ' --- Variablen (Objekte) ---
    Dim tbl As ListObject

    ' --- Variablen (Primitives) ---
    Dim messagesText As String
    Dim messages() As String
    Dim msgBlock As Variant
    Dim analyzedCount As Long
    Dim importedCount As Long
    Dim duplicateCount As Long
    Dim errorCount As Long
    Dim totalBlocks As Long

    Set tbl = FindTableByName(TABLE_NAME)
    If tbl Is Nothing Then
        Application.StatusBar = False
        MsgBox "Tabelle '" & TABLE_NAME & "' nicht gefunden.", vbExclamation
        Exit Sub
    End If

    Application.StatusBar = "Nachrichten abrufen..."
    messagesText = FetchAppleMailMessages(KEYWORD_1, KEYWORD_2)
    If Len(messagesText) = 0 Then
        Application.StatusBar = False
        Exit Sub
    End If

    Set gLeadIndex = BuildExistingLeadIndex(tbl)
    gLeadIndexInitialized = True

    messages = Split(messagesText, MSG_DELIM)
    totalBlocks = UBound(messages) - LBound(messages) + 1
    Application.StatusBar = "Nachrichten analysieren... 0/" & totalBlocks

    Dim processResult As Long

    For Each msgBlock In messages
        ' Schleife: jeden Nachrichtenblock einzeln verarbeiten.
        If Trim$(msgBlock) <> vbNullString Then
            analyzedCount = analyzedCount + 1
            If analyzedCount Mod 5 = 0 Then
                Application.StatusBar = "Nachrichten analysieren... " & analyzedCount & "/" & totalBlocks
            End If

            Err.Clear
            On Error Resume Next
            processResult = ProcessSingleMessage(tbl, CStr(msgBlock))
            If Err.Number <> 0 Then
                Dim errMsg As String
                errMsg = "Fehler #" & Err.Number & ": " & Err.Description
                Debug.Print "[Import] " & errMsg & " bei Nachricht " & analyzedCount
            End If
            On Error GoTo 0

            Select Case processResult
                Case 1
                    importedCount = importedCount + 1
                Case 2
                    duplicateCount = duplicateCount + 1
                Case Else
                    errorCount = errorCount + 1
                    Debug.Print "[Import] Nachricht " & analyzedCount & " fehlgeschlagen (processResult=" & processResult & ")"
            End Select
        End If
    Next msgBlock

    Application.StatusBar = False
    On Error Resume Next
    ThisWorkbook.Worksheets(SHEET_NAME).Range("B2").Value = Format$(Now, "hh:nn dd.mm.yy")
    On Error GoTo 0
    MsgBox "Import abgeschlossen. " & analyzedCount & " Daten analysiert, " & importedCount & " Daten übertragen. Duplikate: " & duplicateCount & ". Fehler: " & errorCount & ".", vbInformation
End Sub

' =========================
' Apple Mail Read
' =========================

Private Function ProcessSingleMessage(ByVal tbl As ListObject, ByVal blockText As String) As Long
    ' Zweck: Einzelne Nachricht verarbeiten (parsen, dekodieren, importieren).
    ' Rückgabe: 1 = importiert, 2 = Duplikat, 0 = Fehler.
    ' Fehler werden hier NICHT abgefangen -> propagieren zum Aufrufer.
    Dim payload As Object
    Dim parsed As Object
    Dim v As Variant
    Dim msgDate As Date
    Dim msgSubject As String
    Dim msgBody As String
    Dim msgFrom As String
    Dim leadType As String

    Set payload = ParseMessageBlock(blockText)

    msgDate = Date
    msgSubject = vbNullString
    msgBody = vbNullString
    msgFrom = vbNullString
    If TryGetKV(payload, "Date", v) Then msgDate = CDate(v)
    If TryGetKV(payload, "Subject", v) Then msgSubject = CStr(v)
    If TryGetKV(payload, "Body", v) Then msgBody = CStr(v)
    msgBody = DecodeBodyIfNeeded(msgBody)
    If TryGetKV(payload, "From", v) Then msgFrom = CStr(v)

    leadType = ResolveLeadType(msgSubject, msgBody)

    Set parsed = ParseLeadContent(msgBody)
    SetKV parsed, "From", msgFrom
    SetKV parsed, "MailBody", msgBody
    SetKV parsed, "MailSubject", msgSubject

    ' Diagnose: alle Felder ausgeben (Debug-Fenster Ctrl+G)
    DebugDumpFields parsed, msgSubject

    If LeadAlreadyExists(tbl, parsed, msgDate) Then
        ProcessSingleMessage = 2
    Else
        AddLeadRow tbl, parsed, msgDate, leadType
        AddLeadToIndex parsed, msgDate
        ProcessSingleMessage = 1
    End If
End Function

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

Private Function IsOutlookMailbox(ByVal mailboxName As String) As Boolean
    ' Zweck: Erkennt anhand eines einzelnen Mailbox-Namens ob Outlook oder Apple Mail.
    ' E-Mail-Adresse (@) oder "outlook"/"exchange" im Namen -> Outlook.
    ' Sonst -> Apple Mail.
    Dim mb As String
    mb = LCase$(Trim$(mailboxName))
    If InStr(1, mb, "@") > 0 Then
        IsOutlookMailbox = True
    ElseIf InStr(1, mb, "outlook") > 0 Then
        IsOutlookMailbox = True
    ElseIf InStr(1, mb, "exchange") > 0 Then
        IsOutlookMailbox = True
    Else
        IsOutlookMailbox = False
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

    f = FreeFile
    On Error Resume Next
    Open filePath For Binary Access Read As #f
    If Err.Number <> 0 Then
        Err.Clear
        ReadTextFile = vbNullString
        On Error GoTo 0
        Exit Function
    End If

    bytes = LOF(f)
    If bytes > 0 Then
        txt = String$(bytes, vbNullChar)
        Get #f, , txt
    End If
    Close #f
    On Error GoTo 0
    ReadTextFile = txt
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

Private Function IsLikelyBase64(ByVal textIn As String) As Boolean
    ' Zweck: Erkennt ob ein Text reines Base64 ist (z.B. MIME-kodierter E-Mail-Body).
    ' Prüft die ersten 20 nicht-leeren Zeilen. Wenn >=80% gültig und >=2 lange Zeilen -> True.
    ' Rückgabe: True wenn Text Base64-kodiert erscheint.
    Dim clean As String
    Dim i As Long
    Dim ch As String
    Dim longLineCount As Long
    Dim validLineCount As Long
    Dim checkedLineCount As Long
    Dim invalidLineCount As Long
    Dim lines() As String
    Dim lineText As String
    Dim j As Long
    Dim lineOk As Boolean
    Const MAX_CHECK_LINES As Long = 20

    clean = Replace(textIn, vbCrLf, vbLf)
    clean = Replace(clean, vbCr, vbLf)
    clean = Trim$(clean)

    If Len(clean) < 40 Then Exit Function

    lines = Split(clean, vbLf)
    For i = LBound(lines) To UBound(lines)
        lineText = Trim$(lines(i))
        If Len(lineText) > 0 Then
            checkedLineCount = checkedLineCount + 1
            If checkedLineCount > MAX_CHECK_LINES Then Exit For

            lineOk = True
            For j = 1 To Len(lineText)
                ch = Mid$(lineText, j, 1)
                Select Case ch
                    Case "A" To "Z", "a" To "z", "0" To "9", "+", "/", "="
                    Case Else
                        lineOk = False
                        Exit For
                End Select
            Next j

            If lineOk Then
                validLineCount = validLineCount + 1
                If Len(lineText) >= 40 Then longLineCount = longLineCount + 1
            Else
                invalidLineCount = invalidLineCount + 1
            End If
        End If
    Next i

    ' Mindestens 80% gültige Zeilen und mindestens 2 lange Zeilen
    If checkedLineCount > 0 And longLineCount >= 2 Then
        If (validLineCount / checkedLineCount) >= 0.8 Then
            IsLikelyBase64 = True
        End If
    End If
End Function

Private Function DecodeBodyIfNeeded(ByVal bodyText As String) As String
    ' Zweck: Body automatisch dekodieren falls er Base64-kodiert oder MIME-Rohtext ist.
    '         Nach Dekodierung wird HTML automatisch in Klartext konvertiert.
    ' Abhängigkeiten: IsLikelyBase64, DecodeBase64ToString, ExtractBodyFromEmail, StripMimeHeaders, HtmlToText.
    ' Rückgabe: Dekodierter Text oder Original falls keine Kodierung erkannt.
    Dim trimmed As String
    Dim decoded As String
    Dim strippedBody As String

    trimmed = Trim$(bodyText)
    If Len(trimmed) = 0 Then
        Debug.Print "[DecodeBody] Body ist leer -> übersprungen"
        DecodeBodyIfNeeded = bodyText
        Exit Function
    End If

    Debug.Print "[DecodeBody] Body-Länge: " & Len(trimmed) & ", erste 80 Zeichen: " & Left$(trimmed, 80)

    ' .emlx-Datei: Erste Zeile ist Byte-Zahl -> entfernen
    If IsNumeric(Left$(trimmed, InStr(trimmed, vbLf) - 1)) Then
        Dim emlxStart As Long
        emlxStart = InStr(trimmed, vbLf) + 1
        If emlxStart > 1 And emlxStart < Len(trimmed) Then
            Debug.Print "[DecodeBody] .emlx Byte-Header erkannt -> Strip"
            trimmed = Mid$(trimmed, emlxStart)
        End If
    End If

    ' Fall 1: Volle MIME-Struktur erkannt (Content-Type Header)
    If InStr(1, trimmed, "Content-Type:", vbTextCompare) > 0 Then
        Debug.Print "[DecodeBody] MIME-Struktur erkannt -> ExtractBodyFromEmail"
        decoded = ExtractBodyFromEmail(trimmed)
        Debug.Print "[DecodeBody] MIME-Ergebnis Länge: " & Len(decoded)
        If Len(Trim$(decoded)) > 0 Then
            DecodeBodyIfNeeded = ConvertHtmlIfNeeded(decoded)
        Else
            DecodeBodyIfNeeded = bodyText
        End If
        Exit Function
    End If

    ' Fall 2: MIME-Header ohne Content-Type (z.B. nur Content-Transfer-Encoding: base64)
    If InStr(1, trimmed, "Content-Transfer-Encoding:", vbTextCompare) > 0 Then
        Debug.Print "[DecodeBody] Content-Transfer-Encoding Header gefunden -> Header strippen"
        strippedBody = StripMimeHeaders(trimmed)
        Debug.Print "[DecodeBody] Nach Header-Strip Länge: " & Len(strippedBody) & ", erste 80 Zeichen: " & Left$(strippedBody, 80)
        If IsLikelyBase64(strippedBody) Then
            Debug.Print "[DecodeBody] Gestripter Body ist Base64 -> dekodieren"
            decoded = DecodeBase64ToString(strippedBody, "utf-8")
            Debug.Print "[DecodeBody] Base64-Ergebnis Länge: " & Len(decoded)
            If Len(Trim$(decoded)) > 0 Then
                Debug.Print "[DecodeBody] Dekodiert OK, erste 120 Zeichen: " & Left$(decoded, 120)
                DecodeBodyIfNeeded = ConvertHtmlIfNeeded(decoded)
            Else
                Debug.Print "[DecodeBody] WARNUNG: Base64-Dekodierung nach Strip lieferte leeren String!"
                DecodeBodyIfNeeded = bodyText
            End If
        Else
            Debug.Print "[DecodeBody] Gestripter Body ist kein Base64 -> Original beibehalten"
            DecodeBodyIfNeeded = strippedBody
        End If
        Exit Function
    End If

    ' Fall 3: Reines Base64 (ohne jegliche Header)
    If IsLikelyBase64(trimmed) Then
        Debug.Print "[DecodeBody] Base64 erkannt -> DecodeBase64ToString"
        decoded = DecodeBase64ToString(trimmed, "utf-8")
        Debug.Print "[DecodeBody] Base64-Ergebnis Länge: " & Len(decoded)
        If Len(Trim$(decoded)) > 0 Then
            Debug.Print "[DecodeBody] Dekodiert OK, erste 120 Zeichen: " & Left$(decoded, 120)
            DecodeBodyIfNeeded = ConvertHtmlIfNeeded(decoded)
        Else
            Debug.Print "[DecodeBody] WARNUNG: Base64-Dekodierung lieferte leeren String!"
            DecodeBodyIfNeeded = bodyText
        End If
        Exit Function
    End If

    ' Fall 4: Rohes Quoted-Printable (ohne MIME-Header, aber =XX Sequenzen im Text)
    If IsLikelyQuotedPrintable(trimmed) Then
        Debug.Print "[DecodeBody] Quoted-Printable erkannt -> DecodeQuotedPrintable"
        decoded = DecodeQuotedPrintable(trimmed, "utf-8")
        Debug.Print "[DecodeBody] QP-Ergebnis Länge: " & Len(decoded)
        If Len(Trim$(decoded)) > 0 Then
            Debug.Print "[DecodeBody] QP dekodiert OK, erste 120 Zeichen: " & Left$(decoded, 120)
            DecodeBodyIfNeeded = ConvertHtmlIfNeeded(decoded)
        Else
            DecodeBodyIfNeeded = bodyText
        End If
        Exit Function
    End If

    Debug.Print "[DecodeBody] Kein Encoding erkannt -> ConvertHtmlIfNeeded prüfen"
    ' Fall 5: Kein Encoding erkannt -> trotzdem auf HTML prüfen
    DecodeBodyIfNeeded = ConvertHtmlIfNeeded(bodyText)
End Function

Private Function IsLikelyQuotedPrintable(ByVal textIn As String) As Boolean
    ' Zweck: Erkennt ob ein Text Quoted-Printable kodiert ist (ohne MIME-Header).
    ' Prüft auf typische QP-Muster: =XX Hex-Sequenzen und Soft-Linebreaks (= am Zeilenende).
    ' Rückgabe: True wenn Text QP-kodiert erscheint.
    Dim lines() As String
    Dim i As Long
    Dim lineText As String
    Dim qpHitCount As Long
    Dim checkedCount As Long
    Dim j As Long
    Dim ch As String
    Dim next2 As String
    Const MAX_CHECK As Long = 30

    textIn = Replace(textIn, vbCrLf, vbLf)
    textIn = Replace(textIn, vbCr, vbLf)

    lines = Split(textIn, vbLf)
    For i = LBound(lines) To UBound(lines)
        lineText = lines(i)
        If Len(lineText) > 0 Then
            checkedCount = checkedCount + 1
            If checkedCount > MAX_CHECK Then Exit For

            ' Soft-Linebreak: Zeile endet mit "="
            If Right$(lineText, 1) = "=" Then
                qpHitCount = qpHitCount + 1
            End If

            ' =XX Hex-Sequenzen suchen (z.B. =20, =C3, =BC)
            j = 1
            Do While j <= Len(lineText) - 2
                ch = Mid$(lineText, j, 1)
                If ch = "=" Then
                    next2 = Mid$(lineText, j + 1, 2)
                    If IsHexPair(next2) Then
                        qpHitCount = qpHitCount + 1
                        j = j + 3
                    Else
                        j = j + 1
                    End If
                Else
                    j = j + 1
                End If
            Loop
        End If
    Next i

    ' Mindestens 3 QP-Treffer in den ersten 30 Zeilen -> wahrscheinlich QP
    IsLikelyQuotedPrintable = (qpHitCount >= 3)
End Function

Private Function ConvertHtmlIfNeeded(ByVal textIn As String) As String
    ' Zweck: Falls der Text HTML enthält, in Klartext konvertieren.
    ' Erkennung: Prüft den gesamten Text auf HTML-Tags und -Entities.
    ' Rückgabe: Klartext oder unveränderter Text wenn kein HTML erkannt.
    Dim lowerText As String
    lowerText = LCase$(textIn)

    If InStr(1, lowerText, "<html", vbBinaryCompare) > 0 _
       Or InStr(1, lowerText, "<body", vbBinaryCompare) > 0 _
       Or InStr(1, lowerText, "<table", vbBinaryCompare) > 0 _
       Or InStr(1, lowerText, "<!doctype", vbBinaryCompare) > 0 _
       Or InStr(1, lowerText, "<div", vbBinaryCompare) > 0 _
       Or InStr(1, lowerText, "<td", vbBinaryCompare) > 0 _
       Or InStr(1, lowerText, "<span", vbBinaryCompare) > 0 _
       Or InStr(1, lowerText, "&nbsp;", vbBinaryCompare) > 0 _
       Or InStr(1, lowerText, "&#", vbBinaryCompare) > 0 Then
        Debug.Print "[ConvertHtml] HTML erkannt -> HtmlToText"
        Dim converted As String
        converted = HtmlToText(textIn)
        Debug.Print "[ConvertHtml] Ergebnis Länge: " & Len(converted) & ", erste 300 Zeichen:"
        Debug.Print Left$(converted, 300)
        ConvertHtmlIfNeeded = converted
    Else
        ConvertHtmlIfNeeded = textIn
    End If
End Function

Private Function StripMimeHeaders(ByVal textIn As String) As String
    ' Zweck: MIME-Header-Zeilen am Anfang des Textes entfernen.
    '         Alles vor der ersten Leerzeile wird als Header betrachtet.
    ' Rückgabe: Text nach den Headern.
    Dim normalized As String
    Dim pos As Long

    normalized = Replace(textIn, vbCrLf, vbLf)
    normalized = Replace(normalized, vbCr, vbLf)

    ' Erste Leerzeile finden (trennt Header von Body)
    pos = InStr(1, normalized, vbLf & vbLf)
    If pos > 0 Then
        StripMimeHeaders = Mid$(normalized, pos + 2)
    Else
        ' Kein Leerzeilen-Separator gefunden -> alles zurückgeben
        StripMimeHeaders = normalized
    End If
End Function

Private Function ExtractBodyFromEmail(ByVal contentText As String) As String
    ' Bevorzugt text/plain, unterstützt base64 und quoted-printable; fällt auf text/html zurück.
    Dim bodyText As String

    contentText = NormalizeLineEndings(contentText)

    bodyText = ParseMimeBody(contentText, "text/plain")
    If Len(bodyText) = 0 Then
        bodyText = ParseMimeBody(contentText, "text/html")
        If Len(bodyText) > 0 Then bodyText = HtmlToText(bodyText)
    End If

    If Len(bodyText) = 0 Then bodyText = LegacyExtractBody(contentText)

    ExtractBodyFromEmail = bodyText
End Function

Private Function ParseMimeBody(ByVal contentText As String, ByVal desiredType As String) As String
    Dim lines() As String
    Dim i As Long
    Dim lineText As String
    Dim inTarget As Boolean
    Dim collecting As Boolean
    Dim encoding As String
    Dim collected As String
    Dim charset As String

    lines = Split(contentText, vbLf)
    charset = "utf-8"

    For i = LBound(lines) To UBound(lines)
        lineText = lines(i)

        If collecting Then
            If Left$(Trim$(lineText), 2) = "--" Then Exit For
            collected = collected & lineText & vbLf
        Else
            If InStr(1, lineText, "Content-Type:", vbTextCompare) > 0 Then
                inTarget = (InStr(1, lineText, desiredType, vbTextCompare) > 0)
                encoding = vbNullString
                collected = vbNullString
                charset = ExtractCharset(lineText, charset)
            ElseIf inTarget And InStr(1, lineText, "Content-Transfer-Encoding:", vbTextCompare) > 0 Then
                If InStr(1, lineText, "base64", vbTextCompare) > 0 Then encoding = "base64"
                If InStr(1, lineText, "quoted-printable", vbTextCompare) > 0 Then encoding = "qp"
            ElseIf inTarget And Len(Trim$(lineText)) = 0 Then
                collecting = True
            End If
        End If
    Next i

    If Len(collected) = 0 Then Exit Function

    Select Case encoding
        Case "base64": ParseMimeBody = DecodeBase64ToString(collected, charset)
        Case "qp": ParseMimeBody = DecodeQuotedPrintable(collected, charset)
        Case Else: ParseMimeBody = collected
    End Select
End Function

Private Function ExtractCharset(ByVal headerLine As String, ByVal defaultCharset As String) As String
    Dim p As Long
    Dim part As String
    Dim c As String

    ExtractCharset = defaultCharset

    p = InStr(1, headerLine, "charset=", vbTextCompare)
    If p = 0 Then Exit Function

    part = Mid$(headerLine, p + 8)
    part = Trim$(part)
    If Left$(part, 1) = Chr$(34) Or Left$(part, 1) = "'" Then
        c = Mid$(part, 2)
        p = InStr(c, Left$(part, 1))
        If p > 0 Then c = Left$(c, p - 1)
    Else
        c = part
    End If

    If Len(c) > 0 Then ExtractCharset = c
End Function

Private Function HtmlToText(ByVal html As String) As String
    ' Zweck: HTML in Klartext konvertieren.
    ' Unterstützt: br, p, div, td, tr, th, li, table + HTML-Entities.

    ' Zeilenumbrüche für Block-/Tabellen-Elemente einfügen
    html = Replace(html, "<br>", vbLf, , , vbTextCompare)
    html = Replace(html, "<br/>", vbLf, , , vbTextCompare)
    html = Replace(html, "<br />", vbLf, , , vbTextCompare)
    html = Replace(html, "</p>", vbLf, , , vbTextCompare)
    html = Replace(html, "</div>", vbLf, , , vbTextCompare)
    html = Replace(html, "</td>", vbLf, , , vbTextCompare)
    html = Replace(html, "</th>", vbLf, , , vbTextCompare)
    html = Replace(html, "</tr>", vbLf, , , vbTextCompare)
    html = Replace(html, "</li>", vbLf, , , vbTextCompare)
    html = Replace(html, "</h1>", vbLf, , , vbTextCompare)
    html = Replace(html, "</h2>", vbLf, , , vbTextCompare)
    html = Replace(html, "</h3>", vbLf, , , vbTextCompare)

    ' <style>...</style> und <script>...</script> Blöcke entfernen
    html = StripHtmlBlock(html, "style")
    html = StripHtmlBlock(html, "script")

    ' Debug: tel: Vorkommen im HTML vor Extraktion
    Dim telDebugPos As Long
    telDebugPos = InStr(1, html, "tel:", vbTextCompare)
    Do While telDebugPos > 0
        Debug.Print "[HtmlToText] tel: bei Pos " & telDebugPos & ": '" & Mid$(html, IIf(telDebugPos > 20, telDebugPos - 20, 1), 80) & "'"
        telDebugPos = InStr(telDebugPos + 4, html, "tel:", vbTextCompare)
    Loop

    ' Tel/Mailto-Links: href-Werte als Text sichern bevor Tags entfernt werden
    html = ExtractTelMailtoLinks(html)

    ' Text-Pattern ( tel:NUMBER ) durch Nummer ersetzen (SendGrid-Tracking)
    html = ExtractTelTextPattern(html)

    ' HTML-Tags entfernen
    Dim i As Long, ch As String, inTag As Boolean, outText As String
    For i = 1 To Len(html)
        ch = Mid$(html, i, 1)
        If ch = "<" Then
            inTag = True
        ElseIf ch = ">" Then
            inTag = False
        ElseIf Not inTag Then
            outText = outText & ch
        End If
    Next i

    ' HTML-Entities dekodieren
    outText = DecodeHtmlEntities(outText)

    ' Non-Breaking Spaces (NBSP) in normale Spaces umwandeln
    outText = Replace(outText, ChrW$(160), " ")

    ' Mehrfache Spaces auf einer Zeile komprimieren
    Do While InStr(1, outText, "  ") > 0
        outText = Replace(outText, "  ", " ")
    Loop

    ' Mehrfache Leerzeilen zusammenfassen (max 2 aufeinander)
    outText = Replace(outText, vbCrLf, vbLf)
    outText = Replace(outText, vbCr, vbLf)
    Do While InStr(1, outText, vbLf & vbLf & vbLf) > 0
        outText = Replace(outText, vbLf & vbLf & vbLf, vbLf & vbLf)
    Loop

    ' Zeilen trimmen (je Zeile führende/abschließende Spaces entfernen)
    Dim trimLines() As String
    Dim tl As Long
    trimLines = Split(outText, vbLf)
    outText = vbNullString
    For tl = LBound(trimLines) To UBound(trimLines)
        If tl = LBound(trimLines) Then
            outText = Trim$(trimLines(tl))
        Else
            outText = outText & vbLf & Trim$(trimLines(tl))
        End If
    Next tl

    HtmlToText = outText
End Function

Private Function StripHtmlBlock(ByVal html As String, ByVal tagName As String) As String
    ' Zweck: Kompletten HTML-Block entfernen (z.B. <style>...</style>).
    Dim openTag As String
    Dim closeTag As String
    Dim posOpen As Long
    Dim posClose As Long

    openTag = "<" & tagName
    closeTag = "</" & tagName & ">"

    Do
        posOpen = InStr(1, html, openTag, vbTextCompare)
        If posOpen = 0 Then Exit Do
        posClose = InStr(posOpen, html, closeTag, vbTextCompare)
        If posClose = 0 Then Exit Do
        html = Left$(html, posOpen - 1) & Mid$(html, posClose + Len(closeTag))
    Loop

    StripHtmlBlock = html
End Function

Private Function ExtractTelMailtoLinks(ByVal html As String) As String
    ' Zweck: Tel/Mailto-Werte aus <a href="tel:..."> und <a href="mailto:..."> extrahieren.
    '         Ersetzt das gesamte <a>...</a> durch den href-Wert (Telefonnummer/E-Mail).
    '         Muss VOR dem Tag-Stripping aufgerufen werden.
    Dim pos As Long
    Dim quoteChar As String
    Dim hrefStart As Long
    Dim hrefEnd As Long
    Dim hrefVal As String
    Dim tagStart As Long
    Dim closeAnchor As Long
    Dim telExtracted As String
    Dim mailExtracted As String
    Dim matchCount As Long

    matchCount = 0
    pos = 1
    Debug.Print "[ExtractTelMailto] Start, HTML-Laenge: " & Len(html)
    Do
        pos = InStr(pos, html, "href=", vbTextCompare)
        If pos = 0 Then Exit Do

        quoteChar = Mid$(html, pos + 5, 1)
        If quoteChar <> Chr$(34) And quoteChar <> "'" Then
            pos = pos + 5
        Else
            hrefStart = pos + 6
            hrefEnd = InStr(hrefStart, html, quoteChar)
            If hrefEnd = 0 Then
                pos = pos + 5
            Else
                hrefVal = Mid$(html, hrefStart, hrefEnd - hrefStart)

                If LCase$(Left$(hrefVal, 4)) = "tel:" Then
                    telExtracted = Mid$(hrefVal, 5)
                    telExtracted = Replace(telExtracted, "%20", " ")
                    matchCount = matchCount + 1
                    Debug.Print "[ExtractTelMailto] tel: gefunden -> '" & telExtracted & "' (Pos " & pos & ")"

                    tagStart = InStrRev(html, "<", pos)
                    closeAnchor = InStr(hrefEnd, html, "</a>", vbTextCompare)

                    If tagStart > 0 And closeAnchor > 0 Then
                        html = Left$(html, tagStart - 1) & " " & telExtracted & " " & Mid$(html, closeAnchor + 4)
                        pos = tagStart + Len(telExtracted) + 2
                    Else
                        pos = hrefEnd + 1
                    End If

                ElseIf LCase$(Left$(hrefVal, 7)) = "mailto:" Then
                    mailExtracted = Mid$(hrefVal, 8)
                    mailExtracted = Replace(mailExtracted, "%20", " ")
                    matchCount = matchCount + 1
                    Debug.Print "[ExtractTelMailto] mailto: gefunden -> '" & mailExtracted & "' (Pos " & pos & ")"

                    tagStart = InStrRev(html, "<", pos)
                    closeAnchor = InStr(hrefEnd, html, "</a>", vbTextCompare)

                    If tagStart > 0 And closeAnchor > 0 Then
                        html = Left$(html, tagStart - 1) & " " & mailExtracted & " " & Mid$(html, closeAnchor + 4)
                        pos = tagStart + Len(mailExtracted) + 2
                    Else
                        pos = hrefEnd + 1
                    End If
                Else
                    pos = hrefEnd + 1
                End If
            End If
        End If
    Loop

    Debug.Print "[ExtractTelMailto] Ende, " & matchCount & " tel/mailto-Links extrahiert"
    ExtractTelMailtoLinks = html
End Function

Private Function ExtractTelTextPattern(ByVal html As String) As String
    ' Zweck: Text-Muster "( tel:NUMBER )" durch die Telefonnummer ersetzen.
    '        SendGrid-Mails enthalten tel: als Text-Inhalt statt als href-Attribut.
    '        Muster: ( tel:004961312652011 ) -> 004961312652011
    '        Leere Muster: ( tel: ) -> werden entfernt
    Dim pos As Long
    Dim closeP As Long
    Dim telVal As String
    Dim fullPattern As String
    Dim matchCount As Long

    matchCount = 0
    pos = 1
    Do
        pos = InStr(pos, html, "( tel:", vbTextCompare)
        If pos = 0 Then Exit Do

        closeP = InStr(pos, html, ")")
        If closeP = 0 Then Exit Do

        ' Vollstaendiges Pattern extrahieren: "( tel:NUMBER )"
        fullPattern = Mid$(html, pos, closeP - pos + 1)
        ' Telefonnummer zwischen "tel:" und ")" extrahieren
        telVal = Trim$(Mid$(fullPattern, InStr(1, fullPattern, "tel:", vbTextCompare) + 4))
        ' Abschliessendes ) entfernen
        If Right$(telVal, 1) = ")" Then telVal = Left$(telVal, Len(telVal) - 1)
        telVal = Trim$(telVal)

        If Len(telVal) > 3 Then
            ' Nummer gefunden -> Pattern durch Nummer ersetzen
            matchCount = matchCount + 1
            Debug.Print "[ExtractTelText] '" & fullPattern & "' -> '" & telVal & "'"
            html = Left$(html, pos - 1) & telVal & Mid$(html, closeP + 1)
            pos = pos + Len(telVal)
        Else
            ' Leeres tel: Pattern -> entfernen
            Debug.Print "[ExtractTelText] Leeres Pattern entfernt: '" & fullPattern & "'"
            html = Left$(html, pos - 1) & Mid$(html, closeP + 1)
        End If
    Loop

    If matchCount > 0 Then
        Debug.Print "[ExtractTelText] " & matchCount & " tel-Text-Pattern ersetzt"
    End If
    ExtractTelTextPattern = html
End Function

Private Function IsLikelyPhoneNumber(ByVal textIn As String) As Boolean
    ' Zweck: Erkennt ob ein Text wie eine Telefonnummer aussieht.
    ' Prüft: Länge 6-25, beginnt mit +/0/(, mindestens 7 Ziffern.
    Dim cleaned As String
    Dim digitCount As Long
    Dim j As Long

    cleaned = Trim$(textIn)
    If Len(cleaned) < 6 Or Len(cleaned) > 25 Then Exit Function
    If Left$(cleaned, 1) <> "+" _
       And Not (Left$(cleaned, 1) >= "0" And Left$(cleaned, 1) <= "9") _
       And Left$(cleaned, 1) <> "(" Then Exit Function

    For j = 1 To Len(cleaned)
        If Mid$(cleaned, j, 1) >= "0" And Mid$(cleaned, j, 1) <= "9" Then
            digitCount = digitCount + 1
        End If
    Next j

    IsLikelyPhoneNumber = (digitCount >= 7)
End Function

Private Function DecodeHtmlEntities(ByVal textIn As String) As String
    ' Zweck: Gängige HTML-Entities dekodieren.
    Dim s As String
    s = textIn

    ' Benannte Entities
    s = Replace(s, "&nbsp;", " ", , , vbTextCompare)
    s = Replace(s, "&amp;", "&", , , vbTextCompare)
    s = Replace(s, "&lt;", "<", , , vbTextCompare)
    s = Replace(s, "&gt;", ">", , , vbTextCompare)
    s = Replace(s, "&quot;", Chr$(34), , , vbTextCompare)
    s = Replace(s, "&apos;", "'", , , vbTextCompare)
    s = Replace(s, "&shy;", "", , , vbTextCompare)
    s = Replace(s, "&ndash;", ChrW$(8211), , , vbTextCompare)
    s = Replace(s, "&mdash;", ChrW$(8212), , , vbTextCompare)
    s = Replace(s, "&ouml;", ChrW$(246), , , vbTextCompare)
    s = Replace(s, "&auml;", ChrW$(228), , , vbTextCompare)
    s = Replace(s, "&uuml;", ChrW$(252), , , vbTextCompare)
    s = Replace(s, "&Ouml;", ChrW$(214), , , vbTextCompare)
    s = Replace(s, "&Auml;", ChrW$(196), , , vbTextCompare)
    s = Replace(s, "&Uuml;", ChrW$(220), , , vbTextCompare)
    s = Replace(s, "&szlig;", ChrW$(223), , , vbTextCompare)
    s = Replace(s, "&euro;", ChrW$(8364), , , vbTextCompare)

    ' Numerische Entities: &#NNN; und &#xHH;
    s = DecodeNumericEntities(s)

    ' Soft-Hyphens entfernen (U+00AD)
    s = Replace(s, ChrW$(&HAD), "")

    DecodeHtmlEntities = s
End Function

Private Function DecodeNumericEntities(ByVal textIn As String) As String
    ' Zweck: Numerische HTML-Entities (&#NNN; und &#xHH;) dekodieren.
    Dim result As String
    Dim i As Long
    Dim semiPos As Long
    Dim entityText As String
    Dim codeVal As Long

    result = vbNullString
    i = 1
    Do While i <= Len(textIn)
        If Mid$(textIn, i, 2) = "&#" Then
            semiPos = InStr(i + 2, textIn, ";")
            If semiPos > 0 And semiPos - i <= 8 Then
                entityText = Mid$(textIn, i + 2, semiPos - i - 2)
                On Error Resume Next
                If Left$(entityText, 1) = "x" Or Left$(entityText, 1) = "X" Then
                    codeVal = CLng("&H" & Mid$(entityText, 2))
                Else
                    codeVal = CLng(entityText)
                End If
                If Err.Number = 0 And codeVal > 0 And codeVal <= &HFFFF& Then
                    result = result & ChrW$(codeVal)
                    i = semiPos + 1
                Else
                    Err.Clear
                    result = result & Mid$(textIn, i, 1)
                    i = i + 1
                End If
                On Error GoTo 0
            Else
                result = result & Mid$(textIn, i, 1)
                i = i + 1
            End If
        Else
            result = result & Mid$(textIn, i, 1)
            i = i + 1
        End If
    Loop

    DecodeNumericEntities = result
End Function

Private Function LegacyExtractBody(ByVal contentText As String) As String
    Dim splitMarker As String
    Dim pos As Long

    splitMarker = vbLf & vbLf
    pos = InStr(1, contentText, splitMarker)
    If pos > 0 Then
        LegacyExtractBody = Mid$(contentText, pos + Len(splitMarker))
    Else
        LegacyExtractBody = vbNullString
    End If
End Function

Private Function DecodeBase64ToString(ByVal base64Data As String, ByVal charset As String) As String
    Dim clean As String
    Dim base64Chars As String
    Dim i As Long
    Dim ch As String
    Dim c0 As Long, c1 As Long, c2 As Long, c3 As Long
    Dim padBlock As Long
    Dim out() As Byte
    Dim outPos As Long
    Dim totalLen As Long

    base64Chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"

    ' nur gültige Zeichen behalten (inkl. "=") - Mac-kompatibel via Mid$ (kein StrConv)
    clean = String$(Len(base64Data), vbNullChar)
    Dim cleanLen As Long
    cleanLen = 0
    For i = 1 To Len(base64Data)
        ch = Mid$(base64Data, i, 1)
        Select Case ch
            Case "A" To "Z", "a" To "z", "0" To "9", "+", "/", "="
                cleanLen = cleanLen + 1
                Mid$(clean, cleanLen, 1) = ch
            Case Else
                ' ignorieren (CR, LF, Tab, Space, etc.)
        End Select
    Next i
    If cleanLen = 0 Then Exit Function
    clean = Left$(clean, cleanLen)

    ' Base64-Padding ergänzen falls nötig (statt Exit)
    Do While Len(clean) Mod 4 <> 0
        clean = clean & "="
    Loop

    totalLen = (Len(clean) \ 4) * 3
    If totalLen = 0 Then Exit Function
    ReDim out(totalLen - 1)

    For i = 1 To Len(clean) Step 4
        padBlock = 0

        c0 = InStr(1, base64Chars, Mid$(clean, i, 1), vbBinaryCompare) - 1
        c1 = InStr(1, base64Chars, Mid$(clean, i + 1, 1), vbBinaryCompare) - 1

        ch = Mid$(clean, i + 2, 1)
        If ch = "=" Then
            c2 = 0: padBlock = padBlock + 1
        Else
            c2 = InStr(1, base64Chars, ch, vbBinaryCompare) - 1
        End If

        ch = Mid$(clean, i + 3, 1)
        If ch = "=" Then
            c3 = 0: padBlock = padBlock + 1
        Else
            c3 = InStr(1, base64Chars, ch, vbBinaryCompare) - 1
        End If

        If c0 < 0 Or c1 < 0 Or c2 < 0 Or c3 < 0 Then Exit Function

        If outPos <= UBound(out) Then
            out(outPos) = (c0 * 4 + c1 \ 16) And &HFF
            outPos = outPos + 1
        End If
        If padBlock < 2 And outPos <= UBound(out) Then
            out(outPos) = ((c1 And &HF) * 16 + c2 \ 4) And &HFF
            outPos = outPos + 1
        End If
        If padBlock = 0 And outPos <= UBound(out) Then
            out(outPos) = ((c2 And 3) * 64 + c3) And &HFF
            outPos = outPos + 1
        End If
    Next i

    DecodeBase64ToString = DecodeBytesToString(out, outPos, charset)
End Function

Private Function DecodeQuotedPrintable(ByVal qpText As String, ByVal charset As String) As String
    Dim i As Long
    Dim ch As String
    Dim next2 As String
    Dim bytes() As Byte
    Dim outPos As Long

    qpText = Replace(qpText, "=" & vbCrLf, "")
    qpText = Replace(qpText, "=" & vbLf, "")

    ReDim bytes(Len(qpText)) ' worst case
    i = 1
    Do While i <= Len(qpText)
        ch = Mid$(qpText, i, 1)
        If ch = "=" And i + 2 <= Len(qpText) Then
            next2 = Mid$(qpText, i + 1, 2)
            If IsHexPair(next2) Then
                bytes(outPos) = CByte(CLng("&H" & next2))
                outPos = outPos + 1
                i = i + 3
            Else
                bytes(outPos) = Asc(ch) And &HFF
                outPos = outPos + 1
                i = i + 1
            End If
        Else
            bytes(outPos) = Asc(ch) And &HFF
            outPos = outPos + 1
            i = i + 1
        End If
    Loop

    DecodeQuotedPrintable = DecodeBytesToString(bytes, outPos, charset)
End Function

Private Function IsHexPair(ByVal txt As String) As Boolean
    Dim i As Long
    If Len(txt) <> 2 Then Exit Function
    For i = 1 To 2
        Select Case Mid$(txt, i, 1)
            Case "0" To "9", "A" To "F", "a" To "f"
            Case Else: Exit Function
        End Select
    Next i
    IsHexPair = True
End Function

Private Function Utf8BytesToString(ByRef bytes() As Byte, ByVal lengthBytes As Long) As String
    Dim i As Long
    Dim b1 As Long, b2 As Long, b3 As Long, b4 As Long
    Dim codePoint As Long
    Dim resultText As String

    If lengthBytes = 0 Then Exit Function

    i = LBound(bytes)
    Do While i < lengthBytes
        b1 = bytes(i)
        Select Case b1
            Case Is < &H80
                resultText = resultText & SafeChrW(b1)
            Case &HC0 To &HDF
                If i + 1 >= lengthBytes Then Exit Do
                b2 = bytes(i + 1)
                codePoint = ((b1 And &H1F) * 64) + (b2 And &H3F)
                resultText = resultText & SafeChrW(codePoint)
                i = i + 1
            Case &HE0 To &HEF
                If i + 2 >= lengthBytes Then Exit Do
                b2 = bytes(i + 1)
                b3 = bytes(i + 2)
                codePoint = ((b1 And &HF) * 4096) + ((b2 And &H3F) * 64) + (b3 And &H3F)
                resultText = resultText & SafeChrW(codePoint)
                i = i + 2
            Case Else
                If i + 3 >= lengthBytes Then Exit Do
                b2 = bytes(i + 1)
                b3 = bytes(i + 2)
                b4 = bytes(i + 3)
                codePoint = ((b1 And &H7) * 262144) + ((b2 And &H3F) * 4096) + ((b3 And &H3F) * 64) + (b4 And &H3F)
                resultText = resultText & SafeChrW(codePoint)
                i = i + 3
        End Select
        i = i + 1
    Loop

    Utf8BytesToString = resultText
End Function

Private Function Latin1BytesToString(ByRef bytes() As Byte, ByVal lengthBytes As Long) As String
    Dim i As Long
    Dim s As String
    If lengthBytes = 0 Then Exit Function
    For i = 0 To lengthBytes - 1
        s = s & ChrW$(bytes(i))
    Next i
    Latin1BytesToString = s
End Function

Private Function DecodeBytesToString(ByRef bytes() As Byte, ByVal lengthBytes As Long, ByVal charset As String) As String
    If lengthBytes = 0 Then Exit Function
    If InStr(1, charset, "utf-8", vbTextCompare) > 0 Then
        DecodeBytesToString = Utf8BytesToString(bytes, lengthBytes)
    ElseIf InStr(1, charset, "windows-1252", vbTextCompare) > 0 _
        Or InStr(1, charset, "cp1252", vbTextCompare) > 0 Then
        DecodeBytesToString = Windows1252BytesToString(bytes, lengthBytes)
    Else
        DecodeBytesToString = Latin1BytesToString(bytes, lengthBytes)
    End If
End Function

Private Function Windows1252BytesToString(ByRef bytes() As Byte, ByVal lengthBytes As Long) As String
    ' Zweck: Windows-1252 Bytes in Unicode-String konvertieren.
    ' Im Bereich 0x80-0x9F weicht Windows-1252 von Latin-1/Unicode ab.
    ' Beispiele: 0x80=Euro, 0x9B=rsaquo, 0x85=Ellipsis, 0x93/94=Anführungszeichen.
    Dim i As Long
    Dim s As String
    Dim b As Long
    If lengthBytes = 0 Then Exit Function
    For i = 0 To lengthBytes - 1
        b = bytes(i)
        Select Case b
            Case &H80: s = s & ChrW$(&H20AC)  ' Euro
            Case &H82: s = s & ChrW$(&H201A)  ' Single Low-9 Quote
            Case &H83: s = s & ChrW$(&H192)   ' Latin Small F With Hook
            Case &H84: s = s & ChrW$(&H201E)  ' Double Low-9 Quote
            Case &H85: s = s & ChrW$(&H2026)  ' Horizontal Ellipsis
            Case &H86: s = s & ChrW$(&H2020)  ' Dagger
            Case &H87: s = s & ChrW$(&H2021)  ' Double Dagger
            Case &H88: s = s & ChrW$(&H2C6)   ' Modifier Letter Circumflex
            Case &H89: s = s & ChrW$(&H2030)  ' Per Mille Sign
            Case &H8A: s = s & ChrW$(&H160)   ' Latin Capital S With Caron
            Case &H8B: s = s & ChrW$(&H2039)  ' Single Left Guillemet
            Case &H8C: s = s & ChrW$(&H152)   ' Latin Capital OE
            Case &H8E: s = s & ChrW$(&H17D)   ' Latin Capital Z With Caron
            Case &H91: s = s & ChrW$(&H2018)  ' Left Single Quote
            Case &H92: s = s & ChrW$(&H2019)  ' Right Single Quote
            Case &H93: s = s & ChrW$(&H201C)  ' Left Double Quote
            Case &H94: s = s & ChrW$(&H201D)  ' Right Double Quote
            Case &H95: s = s & ChrW$(&H2022)  ' Bullet
            Case &H96: s = s & ChrW$(&H2013)  ' En Dash
            Case &H97: s = s & ChrW$(&H2014)  ' Em Dash
            Case &H98: s = s & ChrW$(&H2DC)   ' Small Tilde
            Case &H99: s = s & ChrW$(&H2122)  ' Trade Mark Sign
            Case &H9A: s = s & ChrW$(&H161)   ' Latin Small S With Caron
            Case &H9B: s = s & ChrW$(&H203A)  ' Single Right Guillemet (rsaquo)
            Case &H9C: s = s & ChrW$(&H153)   ' Latin Small OE
            Case &H9E: s = s & ChrW$(&H17E)   ' Latin Small Z With Caron
            Case &H9F: s = s & ChrW$(&H178)   ' Latin Capital Y With Diaeresis
            Case &H81, &H8D, &H8F, &H90, &H9D
                ' Undefiniert in Windows-1252 -> ignorieren
            Case Else
                s = s & ChrW$(b)
        End Select
    Next i
    Windows1252BytesToString = s
End Function

Private Function SafeChrW(ByVal codePoint As Long) As String
    If codePoint < 0 Or codePoint > &HFFFF& Then
        SafeChrW = "?"
    Else
        SafeChrW = ChrW$(codePoint)
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
    ' Zweck: Mail-Nachrichten per AppleScript abrufen. Unterstützt mehrere Quellen (Apple Mail + Outlook).
    ' LEAD_MAILBOX und LEAD_FOLDER können per ";" getrennt mehrere Einträge enthalten.
    ' Jeder Eintrag wird anhand des Mailbox-Namens automatisch der richtigen App zugeordnet.
    ' Abhängigkeiten: AppleScriptTask, BuildAppleMailScript, BuildOutlookScript, IsOutlookMailbox.
    ' Rückgabe: zusammengeführter Nachrichtentext oder Leerstring bei Fehler.
    Dim result As String
    Dim mailPath As String
    Dim pathResult As String
    Dim usedFile As Boolean

    Dim mailboxRaw As String
    Dim folderRaw As String
    Dim mailboxes() As String
    Dim folders() As String
    Dim i As Long
    Dim mbName As String
    Dim flName As String
    Dim isOutlook As Boolean
    Dim script As String
    Dim mailResult As String
    Dim sourceLabels As String
    Dim appLabel As String

    mailPath = GetMailPath()
    result = vbNullString

    ' --- 1. Dateiordner (mailpath) ---
    If Len(Trim$(mailPath)) > 0 Then
        If FolderExists(mailPath) Then
            pathResult = FetchMailMessagesFromPath(mailPath)
            If Len(pathResult) > 0 Then
                result = result & pathResult
                usedFile = True
            End If
        Else
            MsgBox "Mailpath ungültig: " & mailPath, vbExclamation
            LogImportError "Mailpath ungültig", mailPath
        End If
    End If

    ' --- 2. Mail-Quellen (LEAD_MAILBOX;LEAD_FOLDER Paare) ---
    mailboxRaw = GetLeadMailbox()
    folderRaw = GetLeadFolder()

    mailboxes = Split(mailboxRaw, ";")
    folders = Split(folderRaw, ";")

    For i = LBound(mailboxes) To UBound(mailboxes)
        mbName = Trim$(mailboxes(i))
        If Len(mbName) > 0 Then
            ' Zugehörigen Folder holen (gleicher Index, oder letzten wiederverwenden)
            If i <= UBound(folders) Then
                flName = Trim$(folders(i))
            Else
                flName = Trim$(folders(UBound(folders)))
            End If
            If Len(flName) = 0 Then flName = LEAD_FOLDER_DEFAULT

            isOutlook = IsOutlookMailbox(mbName)

            If isOutlook Then
                script = BuildOutlookScript(mbName, flName, keywordA, keywordB)
                appLabel = "Outlook"
            Else
                script = BuildAppleMailScript(mbName, flName, keywordA, keywordB)
                appLabel = "Apple Mail"
            End If

            Err.Clear
            On Error Resume Next
            mailResult = AppleScriptTask(APPLESCRIPT_FILE, APPLESCRIPT_HANDLER, script)
            If Err.Number <> 0 Then
                LogImportError "AppleScriptTask-Fehler (" & appLabel & ", " & mbName & ")", Err.Description
                Err.Clear
            ElseIf Left$(mailResult, 6) = "ERROR:" Then
                LogImportError "AppleScript-Fehler (" & appLabel & ", " & mbName & ")", Mid$(mailResult, 7)
            Else
                result = result & mailResult
                If Len(sourceLabels) > 0 Then sourceLabels = sourceLabels & " | "
                sourceLabels = sourceLabels & appLabel & ": " & BuildMailboxSourceLabel(mbName, flName)
            End If
            On Error GoTo 0
        End If
    Next i

    ' --- 3. Source-Note zusammenbauen ---
    gLeadSourceNote = vbNullString
    If usedFile Then gLeadSourceNote = "Dateiordner: " & mailPath
    If Len(sourceLabels) > 0 Then
        If Len(gLeadSourceNote) > 0 Then gLeadSourceNote = gLeadSourceNote & " | "
        gLeadSourceNote = gLeadSourceNote & sourceLabels
    End If

    FetchAppleMailMessages = result
End Function

Private Function BuildAppleMailScript(ByVal mailboxName As String, ByVal folderName As String, ByVal keywordA As String, ByVal keywordB As String) As String
    ' Zweck: AppleScript für Apple Mail generieren.
    ' Rückgabe: Fertiges Script als String.
    Dim script As String
    Dim q As String
    q = Chr$(34)

    script = ""
    script = script & "with timeout of 90 seconds" & vbLf
    script = script & "tell application ""Mail""" & vbLf
    script = script & "set targetBox to missing value" & vbLf
    If Len(mailboxName) > 0 Then
        script = script & "set targetAccountName to " & q & mailboxName & q & vbLf
        script = script & "try" & vbLf
        script = script & "repeat with a in accounts" & vbLf
        script = script & "if (name of a) contains targetAccountName then" & vbLf
        script = script & "try" & vbLf
        script = script & "set targetBox to first mailbox of a whose name is " & q & folderName & q & vbLf
        script = script & "exit repeat" & vbLf
        script = script & "end try" & vbLf
        script = script & "end if" & vbLf
        script = script & "end repeat" & vbLf
        script = script & "end try" & vbLf
    End If
    script = script & "if targetBox is missing value then" & vbLf
    script = script & "try" & vbLf
    script = script & "set targetBox to first mailbox whose name is " & q & folderName & q & vbLf
    script = script & "end try" & vbLf
    script = script & "end if" & vbLf
    script = script & "if targetBox is missing value then error ""Mailbox nicht gefunden: " & folderName & """" & vbLf
    script = script & "set theMessages to (every message of targetBox whose subject contains """ & keywordA & """ or subject contains """ & keywordB & """ or content contains """ & keywordA & """ or content contains """ & keywordB & """ )" & vbLf
    script = script & "if (count of theMessages) > " & MAX_MESSAGES & " then set theMessages to items 1 thru " & MAX_MESSAGES & " of theMessages" & vbLf
    ' Python MIME text/plain Extraktor (base64-kodiert)
    ' Fix: binary mode + message_from_binary_file + charset aus Content-Type Header
    Dim b64Py As String
    b64Py = "aW1wb3J0IGVtYWlsLCBzeXMKd2l0aCBvcGVuKHN5cy5hcmd2WzFdLCAncmInKSBhcyBm" & _
            "OgogICAgbXNnID0gZW1haWwubWVzc2FnZV9mcm9tX2JpbmFyeV9maWxlKGYpCmZvciBw" & _
            "IGluIG1zZy53YWxrKCk6CiAgICBpZiBwLmdldF9jb250ZW50X3R5cGUoKSA9PSAndGV4" & _
            "dC9wbGFpbic6CiAgICAgICAgcGF5bG9hZCA9IHAuZ2V0X3BheWxvYWQoZGVjb2RlPVRy" & _
            "dWUpCiAgICAgICAgaWYgcGF5bG9hZDoKICAgICAgICAgICAgY2hhcnNldCA9IHAuZ2V0" & _
            "X2NvbnRlbnRfY2hhcnNldCgpIG9yICd1dGYtOCcKICAgICAgICAgICAgc3lzLnN0ZG91" & _
            "dC53cml0ZShwYXlsb2FkLmRlY29kZShjaGFyc2V0LCAncmVwbGFjZScpKQogICAgICAg" & _
            "ICAgICBicmVhawo="

    ' Python-Script einmalig auf Platte schreiben
    script = script & "do shell script ""echo '" & b64Py & "' | base64 -D > /tmp/_ep.py""" & vbLf

    script = script & "set outText to """"" & vbLf
    script = script & "repeat with m in theMessages" & vbLf
        script = script & "set outText to outText & """ & MSG_DELIM & """ & linefeed" & vbLf
        script = script & "set outText to outText & """ & DATE_TAG & """ & (date sent of m) & linefeed" & vbLf
        script = script & "set outText to outText & """ & SUBJECT_TAG & """ & (subject of m) & linefeed" & vbLf
        script = script & "set outText to outText & """ & FROM_TAG & """ & (sender of m) & linefeed" & vbLf
        script = script & "set bodyText to """"" & vbLf
        script = script & "set stratDebug to """"" & vbLf
        script = script & "set srcText to """"" & vbLf

        ' =====================================================================
        ' Schritt A: source of m holen (flaches try, keine Verschachtelung)
        ' =====================================================================
        script = script & "try" & vbLf
        script = script & "set srcText to (source of m)" & vbLf
        script = script & "set stratDebug to ""A:len="" & (length of srcText)" & vbLf
        script = script & "on error errA" & vbLf
        script = script & "set stratDebug to ""A:ERR="" & errA" & vbLf
        script = script & "end try" & vbLf

        ' =====================================================================
        ' Schritt B: Falls source vorhanden -> python3 text/plain extrahieren
        ' Datei via printf und quoted form schreiben, dann python3 separat
        ' =====================================================================
        script = script & "if length of srcText > 100 then" & vbLf
        script = script & "try" & vbLf
        ' printf %s statt echo (binary-safe, kein trailing newline)
        script = script & "do shell script ""printf '%s' "" & quoted form of srcText & "" > /tmp/_src.eml""" & vbLf
        script = script & "set plainBody to do shell script ""python3 /tmp/_ep.py /tmp/_src.eml""" & vbLf
        script = script & "if length of plainBody > 50 then" & vbLf
        script = script & "set bodyText to plainBody" & vbLf
        script = script & "set stratDebug to stratDebug & "";B:py_OK,len="" & (length of plainBody)" & vbLf
        script = script & "else" & vbLf
        script = script & "set stratDebug to stratDebug & "";B:py_kurz="" & (length of plainBody)" & vbLf
        script = script & "end if" & vbLf
        script = script & "on error errB" & vbLf
        script = script & "set stratDebug to stratDebug & "";B:ERR="" & errB" & vbLf
        script = script & "end try" & vbLf
        script = script & "end if" & vbLf

        ' =====================================================================
        ' Schritt C: Falls noch kein body -> source als MIME-raw fuer VBA nutzen
        ' =====================================================================
        script = script & "if bodyText is """" and length of srcText > 100 then" & vbLf
        script = script & "set bodyText to srcText" & vbLf
        script = script & "set stratDebug to stratDebug & "";C:raw""" & vbLf
        script = script & "end if" & vbLf

        ' =====================================================================
        ' Schritt D: content of m -> Fallback
        ' =====================================================================
        script = script & "if bodyText is """" then" & vbLf
        script = script & "try" & vbLf
        script = script & "set bodyText to (content of m)" & vbLf
        script = script & "set stratDebug to stratDebug & "";D:content""" & vbLf
        script = script & "on error errD" & vbLf
        script = script & "set stratDebug to stratDebug & "";D:ERR="" & errD" & vbLf
        script = script & "end try" & vbLf
        script = script & "end if" & vbLf

        script = script & "set outText to outText & ""SRCSTRAT:"" & stratDebug & linefeed" & vbLf
        script = script & "set outText to outText & " & q & BODY_TAG & q & " & bodyText & linefeed" & vbLf
    script = script & "end repeat" & vbLf
    script = script & "return outText" & vbLf
    script = script & "end tell" & vbLf
    script = script & "end timeout"

    BuildAppleMailScript = script
End Function

Private Function BuildOutlookScript(ByVal mailboxName As String, ByVal folderName As String, ByVal keywordA As String, ByVal keywordB As String) As String
    ' Zweck: AppleScript für Microsoft Outlook generieren.
    ' LEAD_MAILBOX = Outlook-Account-Name oder E-Mail-Adresse.
    ' LEAD_FOLDER = Ordnername in Outlook (z. B. "Leads", "Posteingang", "Inbox").
    ' Hinweis: "Posteingang"/"Inbox" wird auf die inbox-Property gemappt.
    ' Rückgabe: Fertiges Script als String.
    Dim script As String
    Dim q As String
    Dim isInbox As Boolean
    q = Chr$(34)

    ' Prüfen ob der Ordner ein Posteingang ist
    isInbox = (LCase$(Trim$(folderName)) = "posteingang" Or _
               LCase$(Trim$(folderName)) = "inbox" Or _
               LCase$(Trim$(folderName)) = "posteingang (inbox)")

    script = ""
    script = script & "with timeout of 60 seconds" & vbLf
    script = script & "tell application ""Microsoft Outlook""" & vbLf

    ' --- Account finden ---
    script = script & "set targetAcct to missing value" & vbLf
    script = script & "set targetFolder to missing value" & vbLf

    If Len(Trim$(mailboxName)) > 0 Then
        script = script & "set targetAccountName to " & q & mailboxName & q & vbLf

        ' Helper: Account matchen per Name ODER E-Mail-Adresse
        ' Exchange Accounts
        script = script & "try" & vbLf
        script = script & "repeat with acct in exchange accounts" & vbLf
        script = script & "set matchFound to false" & vbLf
        script = script & "if (name of acct) contains targetAccountName then set matchFound to true" & vbLf
        script = script & "try" & vbLf
        script = script & "if (email address of acct) contains targetAccountName then set matchFound to true" & vbLf
        script = script & "end try" & vbLf
        script = script & "if matchFound then" & vbLf
        script = script & "set targetAcct to acct" & vbLf
        script = script & "exit repeat" & vbLf
        script = script & "end if" & vbLf
        script = script & "end repeat" & vbLf
        script = script & "end try" & vbLf

        ' IMAP Accounts
        script = script & "if targetAcct is missing value then" & vbLf
        script = script & "try" & vbLf
        script = script & "repeat with acct in imap accounts" & vbLf
        script = script & "set matchFound to false" & vbLf
        script = script & "if (name of acct) contains targetAccountName then set matchFound to true" & vbLf
        script = script & "try" & vbLf
        script = script & "if (email address of acct) contains targetAccountName then set matchFound to true" & vbLf
        script = script & "end try" & vbLf
        script = script & "if matchFound then" & vbLf
        script = script & "set targetAcct to acct" & vbLf
        script = script & "exit repeat" & vbLf
        script = script & "end if" & vbLf
        script = script & "end repeat" & vbLf
        script = script & "end try" & vbLf
        script = script & "end if" & vbLf

        ' POP Accounts
        script = script & "if targetAcct is missing value then" & vbLf
        script = script & "try" & vbLf
        script = script & "repeat with acct in pop accounts" & vbLf
        script = script & "set matchFound to false" & vbLf
        script = script & "if (name of acct) contains targetAccountName then set matchFound to true" & vbLf
        script = script & "try" & vbLf
        script = script & "if (email address of acct) contains targetAccountName then set matchFound to true" & vbLf
        script = script & "end try" & vbLf
        script = script & "if matchFound then" & vbLf
        script = script & "set targetAcct to acct" & vbLf
        script = script & "exit repeat" & vbLf
        script = script & "end if" & vbLf
        script = script & "end repeat" & vbLf
        script = script & "end try" & vbLf
        script = script & "end if" & vbLf
    End If

    ' Fallback: Default Account
    script = script & "if targetAcct is missing value then" & vbLf
    script = script & "try" & vbLf
    script = script & "set targetAcct to default account" & vbLf
    script = script & "end try" & vbLf
    script = script & "end if" & vbLf

    script = script & "if targetAcct is missing value then error ""Outlook-Account nicht gefunden: " & mailboxName & """" & vbLf

    ' --- Ordner im Account finden ---
    If isInbox Then
        ' Posteingang/Inbox: direkt die inbox-Property nutzen
        script = script & "try" & vbLf
        script = script & "set targetFolder to inbox of targetAcct" & vbLf
        script = script & "end try" & vbLf
    Else
        ' Benannter Ordner
        script = script & "try" & vbLf
        script = script & "set targetFolder to mail folder " & q & folderName & q & " of targetAcct" & vbLf
        script = script & "end try" & vbLf
    End If

    ' Fallback: auch den anderen Weg versuchen
    If isInbox Then
        script = script & "if targetFolder is missing value then" & vbLf
        script = script & "try" & vbLf
        script = script & "set targetFolder to mail folder " & q & folderName & q & " of targetAcct" & vbLf
        script = script & "end try" & vbLf
        script = script & "end if" & vbLf
    Else
        script = script & "if targetFolder is missing value then" & vbLf
        script = script & "try" & vbLf
        script = script & "set targetFolder to inbox of targetAcct" & vbLf
        script = script & "end try" & vbLf
        script = script & "end if" & vbLf
    End If

    script = script & "if targetFolder is missing value then error ""Outlook-Ordner nicht gefunden: " & folderName & """" & vbLf

    ' --- Nachrichten filtern ---
    script = script & "set theMessages to (every message of targetFolder whose subject contains " & q & keywordA & q & " or subject contains " & q & keywordB & q & ")" & vbLf
    script = script & "if (count of theMessages) > " & MAX_MESSAGES & " then set theMessages to items 1 thru " & MAX_MESSAGES & " of theMessages" & vbLf

    ' --- Nachrichten ausgeben ---
    script = script & "set outText to """"" & vbLf
    script = script & "repeat with m in theMessages" & vbLf
        script = script & "set outText to outText & """ & MSG_DELIM & """ & linefeed" & vbLf
        script = script & "try" & vbLf
        script = script & "set outText to outText & """ & DATE_TAG & """ & (time sent of m) & linefeed" & vbLf
        script = script & "end try" & vbLf
        script = script & "set outText to outText & """ & SUBJECT_TAG & """ & (subject of m) & linefeed" & vbLf
        script = script & "try" & vbLf
        script = script & "set senderAddr to " & q & q & vbLf
        script = script & "set senderObj to sender of m" & vbLf
        script = script & "set senderAddr to address of senderObj" & vbLf
        script = script & "set outText to outText & """ & FROM_TAG & """ & (name of senderObj) & "" <"" & senderAddr & "">"" & linefeed" & vbLf
        script = script & "on error" & vbLf
        script = script & "set outText to outText & """ & FROM_TAG & """  & linefeed" & vbLf
        script = script & "end try" & vbLf
        script = script & "try" & vbLf
        script = script & "set bodyText to plain text content of m" & vbLf
        script = script & "on error" & vbLf
        script = script & "try" & vbLf
        script = script & "set bodyText to content of m" & vbLf
        script = script & "on error" & vbLf
        script = script & "set bodyText to " & q & q & vbLf
        script = script & "end try" & vbLf
        script = script & "end try" & vbLf
        script = script & "set outText to outText & " & q & BODY_TAG & q & " & bodyText & linefeed" & vbLf
    script = script & "end repeat" & vbLf
    script = script & "return outText" & vbLf
    script = script & "end tell" & vbLf
    script = script & "end timeout"

    BuildOutlookScript = script
End Function

Private Function BuildMailboxSourceLabel(ByVal mailboxName As String, ByVal folderName As String) As String
    Dim labelText As String

    labelText = Trim$(mailboxName)
    If Len(labelText) > 0 And Len(Trim$(folderName)) > 0 Then
        labelText = labelText & " / " & Trim$(folderName)
    ElseIf Len(labelText) = 0 Then
        labelText = Trim$(folderName)
    End If

    If Len(labelText) = 0 Then labelText = "(unbekannt)"
    BuildMailboxSourceLabel = labelText
End Function

Public Sub DebugPrintAppleMailFolders()
    ' Zweck: Mailbox-Ordnerstruktur im Direktfenster ausgeben (alle konfigurierten Quellen).
    ' Abhängigkeiten: FetchAppleMailFolderList, FetchOutlookFolderList, IsOutlookMailbox.
    ' Rückgabe: keine (Debug.Print Ausgabe).
    Dim folderText As String
    Dim lines() As String
    Dim i As Long
    Dim mailboxes() As String
    Dim mbName As String
    Dim hasAppleMail As Boolean
    Dim hasOutlook As Boolean

    mailboxes = Split(GetLeadMailbox(), ";")
    For i = LBound(mailboxes) To UBound(mailboxes)
        mbName = Trim$(mailboxes(i))
        If Len(mbName) > 0 Then
            If IsOutlookMailbox(mbName) Then hasOutlook = True Else hasAppleMail = True
        End If
    Next i

    If hasAppleMail Then
        Debug.Print "=== Apple Mail Ordner ==="
        folderText = FetchAppleMailFolderList()
        If Len(folderText) > 0 Then
            lines = Split(folderText, vbLf)
            For i = LBound(lines) To UBound(lines)
                If Len(Trim$(lines(i))) > 0 Then Debug.Print Trim$(lines(i))
            Next i
        End If
        Debug.Print ""
    End If

    If hasOutlook Then
        Debug.Print "=== Outlook Ordner ==="
        folderText = FetchOutlookFolderList()
        If Len(folderText) > 0 Then
            lines = Split(folderText, vbLf)
            For i = LBound(lines) To UBound(lines)
                If Len(Trim$(lines(i))) > 0 Then Debug.Print Trim$(lines(i))
            Next i
        End If
    End If
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

    Err.Clear
    On Error Resume Next
    result = AppleScriptTask(APPLESCRIPT_FILE, APPLESCRIPT_HANDLER, script)
    If Err.Number <> 0 Then
        MsgBox "AppleScriptTask-Fehler. Prüfe Script-Installation und Automation-Rechte.", vbExclamation
        LogImportError "AppleScriptTask-Fehler", Err.Description
        Err.Clear
        On Error GoTo 0
        FetchAppleMailFolderList = vbNullString
        Exit Function
    End If
    On Error GoTo 0
    If Left$(result, 6) = "ERROR:" Then
        MsgBox "AppleScript-Fehler: " & Mid$(result, 7), vbExclamation
        LogImportError "AppleScript-Fehler", Mid$(result, 7)
        FetchAppleMailFolderList = vbNullString
        Exit Function
    End If
    FetchAppleMailFolderList = result
End Function

Private Function FetchOutlookFolderList() As String
    ' Zweck: Ordnerliste aus Microsoft Outlook via AppleScript abrufen.
    ' Listet ALLE Accounts und deren Ordner (Debug-Funktion).
    ' Abhängigkeiten: AppleScriptTask.
    ' Rückgabe: Textliste der Ordner oder Leerstring bei Fehler.
    Dim script As String
    Dim result As String
    Dim q As String

    q = Chr$(34)

    script = ""
    script = script & "with timeout of 30 seconds" & vbLf
    script = script & "tell application ""Microsoft Outlook""" & vbLf
    script = script & "set outText to " & q & q & vbLf

    ' Exchange Accounts
    script = script & "repeat with acct in exchange accounts" & vbLf
    script = script & "set aName to (name of acct)" & vbLf
    script = script & "set aAddr to " & q & q & vbLf
    script = script & "try" & vbLf
    script = script & "set aAddr to email address of acct" & vbLf
    script = script & "end try" & vbLf
    script = script & "set outText to outText & " & q & "ACCOUNT (Exchange): " & q & " & aName & " & q & " [" & q & " & aAddr & " & q & "]" & q & " & linefeed" & vbLf
    script = script & "try" & vbLf
    script = script & "set outText to outText & " & q & "  -> inbox (Posteingang)" & q & " & linefeed" & vbLf
    script = script & "end try" & vbLf
    script = script & "try" & vbLf
    script = script & "repeat with f in mail folders of acct" & vbLf
    script = script & "set outText to outText & " & q & "  " & q & " & (name of f) & linefeed" & vbLf
    script = script & "end repeat" & vbLf
    script = script & "end try" & vbLf
    script = script & "set outText to outText & linefeed" & vbLf
    script = script & "end repeat" & vbLf

    ' IMAP Accounts
    script = script & "repeat with acct in imap accounts" & vbLf
    script = script & "set aName to (name of acct)" & vbLf
    script = script & "set aAddr to " & q & q & vbLf
    script = script & "try" & vbLf
    script = script & "set aAddr to email address of acct" & vbLf
    script = script & "end try" & vbLf
    script = script & "set outText to outText & " & q & "ACCOUNT (IMAP): " & q & " & aName & " & q & " [" & q & " & aAddr & " & q & "]" & q & " & linefeed" & vbLf
    script = script & "try" & vbLf
    script = script & "set outText to outText & " & q & "  -> inbox (Posteingang)" & q & " & linefeed" & vbLf
    script = script & "end try" & vbLf
    script = script & "try" & vbLf
    script = script & "repeat with f in mail folders of acct" & vbLf
    script = script & "set outText to outText & " & q & "  " & q & " & (name of f) & linefeed" & vbLf
    script = script & "end repeat" & vbLf
    script = script & "end try" & vbLf
    script = script & "set outText to outText & linefeed" & vbLf
    script = script & "end repeat" & vbLf

    ' POP Accounts
    script = script & "repeat with acct in pop accounts" & vbLf
    script = script & "set aName to (name of acct)" & vbLf
    script = script & "set aAddr to " & q & q & vbLf
    script = script & "try" & vbLf
    script = script & "set aAddr to email address of acct" & vbLf
    script = script & "end try" & vbLf
    script = script & "set outText to outText & " & q & "ACCOUNT (POP): " & q & " & aName & " & q & " [" & q & " & aAddr & " & q & "]" & q & " & linefeed" & vbLf
    script = script & "try" & vbLf
    script = script & "set outText to outText & " & q & "  -> inbox (Posteingang)" & q & " & linefeed" & vbLf
    script = script & "end try" & vbLf
    script = script & "try" & vbLf
    script = script & "repeat with f in mail folders of acct" & vbLf
    script = script & "set outText to outText & " & q & "  " & q & " & (name of f) & linefeed" & vbLf
    script = script & "end repeat" & vbLf
    script = script & "end try" & vbLf
    script = script & "set outText to outText & linefeed" & vbLf
    script = script & "end repeat" & vbLf

    script = script & "if outText is " & q & q & " then set outText to " & q & "Keine Outlook-Accounts gefunden." & q & " & linefeed" & vbLf
    script = script & "return outText" & vbLf
    script = script & "end tell" & vbLf
    script = script & "end timeout"

    Err.Clear
    On Error Resume Next
    result = AppleScriptTask(APPLESCRIPT_FILE, APPLESCRIPT_HANDLER, script)
    If Err.Number <> 0 Then
        MsgBox "AppleScriptTask-Fehler (Outlook). Prüfe Script-Installation und Automation-Rechte.", vbExclamation
        LogImportError "AppleScriptTask-Fehler (Outlook)", Err.Description
        Err.Clear
        On Error GoTo 0
        FetchOutlookFolderList = vbNullString
        Exit Function
    End If
    On Error GoTo 0
    If Left$(result, 6) = "ERROR:" Then
        MsgBox "AppleScript-Fehler (Outlook): " & Mid$(result, 7), vbExclamation
        LogImportError "AppleScript-Fehler (Outlook)", Mid$(result, 7)
        FetchOutlookFolderList = vbNullString
        Exit Function
    End If
    FetchOutlookFolderList = result
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

    Err.Clear
    On Error Resume Next

    If Len(Dir$(targetPath)) > 0 Then Kill targetPath
    FileCopy sourcePath, targetPath

    If Err.Number <> 0 Then
        If Err.Number = 75 Then
            MsgBox "Zugriff verweigert. Bitte manuell kopieren nach: " & folderPath & " oder AUTO_INSTALL_APPLESCRIPT aktivieren.", vbExclamation
        Else
            MsgBox "AppleScript konnte nicht installiert werden. Prüfe Rechte.", vbExclamation
        End If
        Err.Clear
    End If

    On Error GoTo 0
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
        On Error Resume Next
        valueOut = store(k)
        TryGetKV = (Err.Number = 0)
        Err.Clear
        On Error GoTo 0
    End If
End Function


' =========================
' Message Parsing
' =========================
Private Function ParseMessageBlock(ByVal blockText As String) As Object
    ' Zweck: Datum/Betreff/Body aus einem Message-Block extrahieren.
    ' Abhängigkeiten: NewKeyValueStore, ParseAppleMailDate, SetKV.
    ' Rückgabe: Key/Value-Store mit "Date", "Subject", "Body", "From".
    ' Fix: Nach BODY:-Tag werden ALLE Folgezeilen dem Body zugeordnet,
    '       ohne erneute Tag-Prüfung. Leerzeilen im Body bleiben erhalten.
    Dim lines() As String
    Dim i As Long
    Dim lineText As String
    Dim payload As Object
    Dim inBody As Boolean
    Dim bodyAccum As String

    Set payload = NewKeyValueStore()
    SetKV payload, "Date", Date
    SetKV payload, "Subject", vbNullString
    SetKV payload, "From", vbNullString
    SetKV payload, "Body", vbNullString

    ' Zeilenenden normalisieren (CRLF/CR -> LF) um \r-Artefakte zu vermeiden
    blockText = Replace(blockText, vbCrLf, vbLf)
    blockText = Replace(blockText, vbCr, vbLf)

    inBody = False
    bodyAccum = vbNullString

    lines = Split(blockText, vbLf)
    For i = LBound(lines) To UBound(lines)
        ' Schleife: jede Zeile des Message-Blocks auswerten.
        If inBody Then
            ' Nach BODY:-Tag: ALLE Zeilen sind Body-Inhalt (inkl. Leerzeilen)
            bodyAccum = bodyAccum & lines(i) & vbLf
        Else
            ' Header-Bereich: Tags prüfen
            lineText = Trim$(lines(i))
            If Len(lineText) > 0 Then
                If Left$(lineText, Len(DATE_TAG)) = DATE_TAG Then
                    SetKV payload, "Date", ParseAppleMailDate(Trim$(Mid$(lineText, Len(DATE_TAG) + 1)))
                ElseIf Left$(lineText, Len(SUBJECT_TAG)) = SUBJECT_TAG Then
                    SetKV payload, "Subject", Trim$(Mid$(lineText, Len(SUBJECT_TAG) + 1))
                ElseIf Left$(lineText, Len(FROM_TAG)) = FROM_TAG Then
                    SetKV payload, "From", Trim$(Mid$(lineText, Len(FROM_TAG) + 1))
                ElseIf Left$(lineText, 9) = "SRCSTRAT:" Then
                    Debug.Print "[AppleScript] " & lineText
                ElseIf Left$(lineText, Len(BODY_TAG)) = BODY_TAG Then
                    bodyAccum = Trim$(Mid$(lineText, Len(BODY_TAG) + 1)) & vbLf
                    inBody = True
                End If
            End If
        End If
    Next i

    SetKV payload, "Body", bodyAccum

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

    On Error Resume Next
    ParseAppleMailDate = CDate(t)
    If Err.Number = 0 Then
        On Error GoTo 0
        Exit Function
    End If
    Err.Clear
    On Error GoTo 0

    parts = Split(t, " ")
    If UBound(parts) < 2 Then
        ParseAppleMailDate = Date
        Exit Function
    End If

    On Error Resume Next
    dayNum = CLng(Replace(parts(0), ".", ""))
    monthNum = GermanMonthToNumber(parts(1))
    yearNum = CLng(parts(2))

    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        ParseAppleMailDate = Date
        Exit Function
    End If
    On Error GoTo 0

    timePart = vbNullString
    If UBound(parts) >= 3 Then timePart = parts(3)

    h = 0: m = 0: s = 0
    If Len(timePart) > 0 Then
        timeParts = Split(timePart, ":")
        On Error Resume Next
        If UBound(timeParts) >= 0 Then h = CLng(timeParts(0))
        If UBound(timeParts) >= 1 Then m = CLng(timeParts(1))
        If UBound(timeParts) >= 2 Then s = CLng(timeParts(2))
        Err.Clear
        On Error GoTo 0
    End If

    On Error Resume Next
    ParseAppleMailDate = DateSerial(yearNum, monthNum, dayNum) + TimeSerial(h, m, s)
    If Err.Number <> 0 Then
        Err.Clear
        ParseAppleMailDate = Date
    End If
    On Error GoTo 0
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
        Case "m" & ChrW$(228) & "rz", "maerz": GermanMonthToNumber = 3
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
    Dim startLine As Long

    Set result = NewKeyValueStore()

    currentSection = "Kontakt"
    pendingKey = vbNullString

    workText = bodyText

    ' Zeilenenden normalisieren (CRLF/CR -> LF) um \r-Artefakte zu vermeiden
    workText = Replace(workText, vbCrLf, vbLf)
    workText = Replace(workText, vbCr, vbLf)

    ' Soft-Hyphens (U+00AD) entfernen, die Sektionserkennung stören können
    workText = Replace(workText, ChrW$(&HAD), "")

    ' Unicode Replacement Character (U+FFFD) entfernen
    ' Entsteht wenn Windows-1252 Bytes fälschlich als UTF-8 dekodiert werden
    workText = Replace(workText, ChrW$(&HFFFD), "")

    ' Non-Breaking Spaces (NBSP) in normale Spaces umwandeln
    workText = Replace(workText, ChrW$(160), " ")

    Debug.Print "[ParseLead] Body-Länge: " & Len(workText) & ", erste 200 Zeichen:"
    Debug.Print Left$(workText, 200)

    lines = Split(workText, vbLf)
    Debug.Print "[ParseLead] Anzahl Zeilen: " & (UBound(lines) - LBound(lines) + 1)

    ' Vorlauf-Rauschen ueberspringen: erste Sektions-Ueberschrift suchen
    ' (Weitergeleitete Mails enthalten Signaturen/Header vor den Lead-Daten)
    startLine = LBound(lines)
    For i = LBound(lines) To UBound(lines)
        lineText = Trim$(Replace(lines(i), ChrW$(160), " "))
        lineText = Replace(lineText, ChrW$(&HAD), "")
        If InStr(1, lineText, "Kontaktinformationen", vbTextCompare) > 0 _
           Or InStr(1, lineText, "Informationen zum Senior", vbTextCompare) > 0 _
           Or InStr(1, lineText, "Senior-Informationen", vbTextCompare) > 0 _
           Or InStr(1, lineText, "Angaben zum Senior", vbTextCompare) > 0 _
           Or InStr(1, lineText, "Zur betreuten Person", vbTextCompare) > 0 _
           Or InStr(1, lineText, "Betreute Person", vbTextCompare) > 0 Then
            startLine = i
            Exit For
        End If
    Next i

    If startLine > LBound(lines) Then
        Debug.Print "[ParseLead] Sektionsheader in Zeile " & startLine & " -> " & startLine & " Vorlauf-Zeilen uebersprungen"
    End If

    For i = startLine To UBound(lines)
        ' Schleife: Zeilen iterieren und Abschnitt/Felder erkennen.
        lineText = Trim$(Replace(lines(i), ChrW$(160), " "))
        ' Asterisken entfernen (HTML-Bold-Marker *Label:* aus SendGrid-Templates)
        lineText = Replace(lineText, "*", "")
        lineText = Trim$(lineText)
        If Len(lineText) > 0 Then
            ' pendingKey hat Vorrang: wenn ein Label wartet, ist naechste Zeile der Wert
            If Len(pendingKey) > 0 Then
                Debug.Print "[ParseLead] MapLabel: '" & pendingKey & "' -> '" & Left$(lineText, 60) & "' [" & currentSection & "] (Zeile " & i & ")"
                MapLabelValue result, pendingKey, lineText, currentSection
                pendingKey = vbNullString
                GoTo NextLineLabel
            End If

            ' Sektionserkennung VOR Noise-Check (Sektionsheader koennen URLs enthalten)
            If InStr(1, lineText, "Kontaktinformationen", vbTextCompare) > 0 Then
                currentSection = "Kontakt"
                pendingKey = vbNullString
                Debug.Print "[ParseLead] Sektion: Kontakt (Zeile " & i & ")"
                GoTo NextLineLabel
            ElseIf InStr(1, lineText, "Informationen zum Senior", vbTextCompare) > 0 _
                Or InStr(1, lineText, "Senior-Informationen", vbTextCompare) > 0 _
                Or InStr(1, lineText, "Angaben zum Senior", vbTextCompare) > 0 _
                Or InStr(1, lineText, "Zur betreuten Person", vbTextCompare) > 0 _
                Or InStr(1, lineText, "Betreute Person", vbTextCompare) > 0 _
                Or InStr(1, lineText, "Pflegebed" & ChrW$(252) & "rftige", vbTextCompare) > 0 Then
                currentSection = "Senior"
                pendingKey = vbNullString
                Debug.Print "[ParseLead] Sektion: Senior (Zeile " & i & ")"
                GoTo NextLineLabel
            ElseIf InStr(1, lineText, "Anfragedetails", vbTextCompare) > 0 Then
                currentSection = "Anfrage"
                pendingKey = vbNullString
                Debug.Print "[ParseLead] Sektion: Anfrage (Zeile " & i & ")"
                GoTo NextLineLabel
            ElseIf InStr(1, lineText, "Datenschutz", vbTextCompare) > 0 _
                And InStr(1, lineText, "zustimmung", vbTextCompare) = 0 Then
                currentSection = "Footer"
                pendingKey = vbNullString
                Debug.Print "[ParseLead] Sektion: Footer (Zeile " & i & ") -> Parsing beendet"
                Exit For
            End If

            ' Rausch-Zeilen ueberspringen (E-Mail-Header, URLs, Footer)
            If IsNoiseLine(lineText) Then
                Debug.Print "[ParseLead] Noise uebersprungen (Zeile " & i & "): '" & Left$(lineText, 60) & "'"
                GoTo NextLineLabel
            End If

            ' Label/Wert-Erkennung
            If Right$(lineText, 1) = ":" And Len(lineText) > 1 Then
                Dim candidateKey As String
                candidateKey = Trim$(Left$(lineText, Len(lineText) - 1))
                ' Wenn der "Key" wie eine Telefonnummer aussieht, direkt als Mobil speichern
                If currentSection = "Kontakt" And IsLikelyPhoneNumber(candidateKey) Then
                    If Len(GetField(result, "Kontakt_Mobil")) = 0 Then
                        Debug.Print "[ParseLead] Tel-als-Label: '" & candidateKey & "' (Zeile " & i & ")"
                        SetKV result, "Kontakt_Mobil", CleanLinkedValue(candidateKey)
                    End If
                    pendingKey = vbNullString
                Else
                    pendingKey = candidateKey
                    If Len(pendingKey) > 0 Then
                        Debug.Print "[ParseLead] PendingKey: '" & pendingKey & "' (Zeile " & i & ")"
                    End If
                End If
            ElseIf InStr(lineText, ":") > 0 Then
                ' tel:NUMBER Pattern in Kontakt-Sektion extrahieren
                Dim telPosInline As Long
                telPosInline = InStr(1, lineText, "tel:", vbTextCompare)
                If telPosInline > 0 And currentSection = "Kontakt" Then
                    Dim telNumInline As String
                    telNumInline = Trim$(Mid$(lineText, telPosInline + 4))
                    telNumInline = Replace(telNumInline, ")", "")
                    telNumInline = Replace(telNumInline, "(", "")
                    telNumInline = Trim$(telNumInline)
                    If IsLikelyPhoneNumber(telNumInline) And Len(GetField(result, "Kontakt_Mobil")) = 0 Then
                        Debug.Print "[ParseLead] Tel-Pattern: '" & telNumInline & "' (Zeile " & i & ")"
                        SetKV result, "Kontakt_Mobil", CleanLinkedValue(telNumInline)
                    Else
                        Debug.Print "[ParseLead] Inline: '" & Left$(lineText, 80) & "' [" & currentSection & "] (Zeile " & i & ")"
                        MapInlinePair result, lineText, currentSection
                    End If
                Else
                    Debug.Print "[ParseLead] Inline: '" & Left$(lineText, 80) & "' [" & currentSection & "] (Zeile " & i & ")"
                    MapInlinePair result, lineText, currentSection
                End If
            ElseIf currentSection = "Kontakt" And IsLikelyPhoneNumber(lineText) Then
                If Len(GetField(result, "Kontakt_Mobil")) = 0 Then
                    Debug.Print "[ParseLead] Auto-Mobil: '" & lineText & "' (Zeile " & i & ")"
                    SetKV result, "Kontakt_Mobil", CleanLinkedValue(lineText)
                End If
            End If
        End If
NextLineLabel:
    Next i

    If Len(Trim$(GetField(result, "Senior_Name"))) = 0 Then
        Dim contactName As String
        contactName = Trim$(GetField(result, "Kontakt_Name"))
        If Len(contactName) = 0 Then
            contactName = Trim$(GetField(result, "Kontakt_Vorname") & " " & GetField(result, "Kontakt_Nachname"))
        End If
        If Len(contactName) > 0 Then SetKV result, "Senior_Name", StripNamePrefix(contactName)
    End If

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
        Case "name", "vollst" & ChrW$(228) & "ndiger name", "kontaktperson", "ihr name"
            If LCase$(sectionName) = "senior" Then
                SetKV fields, "Senior_Name", valueNorm
            Else
                SetKV fields, "Kontakt_Name", valueNorm
            End If
        Case "mobil", "telefonnummer", "festnetz", "telefon", "handynummer", "mobilnummer", "handy", "mobiltelefon", "telefon/mobil", "tel", "tel.": SetKV fields, "Kontakt_Mobil", CleanLinkedValue(valueNorm)
        Case "e-mail", "e-mail-adresse", "email", "emailadresse", "email-adresse", "mail", "e-mailadresse": SetKV fields, "Kontakt_Email", CleanLinkedValue(valueNorm)
        Case "erreichbarkeit": SetKV fields, "Kontakt_Erreichbarkeit", valueNorm
        Case "anschrift", "stra" & ChrW$(223) & "e", "strasse", "adresse": SetKV fields, "Kontakt_Anschrift", valueNorm
        Case "beziehung": SetKV fields, "Senior_Beziehung", valueNorm
        Case "alter": SetKV fields, "Senior_Alter", valueNorm
        Case "pflegegrad status": SetKV fields, "Senior_Pflegegrad_Status", valueNorm
        Case "pflegegrad", "pflegegrad/-stufe": SetKV fields, "Senior_Pflegegrad", valueNorm
        Case "lebenssituation": SetKV fields, "Senior_Lebenssituation", valueNorm
        Case "mobilit" & ChrW$(228) & "t": SetKV fields, "Senior_Mobilitaet", valueNorm
        Case "medizinisches": SetKV fields, "Senior_Medizinisches", valueNorm
        Case "behinderung": SetKV fields, "Senior_Behinderung", valueNorm
        Case "postleitzahl", "plz": SetKV fields, "PLZ", CleanPostalCode(valueNorm)
        Case "bedarfsort", "ort", "stadt", "wohnort": SetBedarfsort fields, valueNorm
        Case "nutzer": SetKV fields, "Nutzer", valueNorm
        Case "alltagshilfe aufgaben": SetKV fields, "Alltagshilfe_Aufgaben", valueNorm
        Case "alltagshilfe h" & ChrW$(228) & "ufigkeit": SetKV fields, "Alltagshilfe_Haeufigkeit", valueNorm
        Case "aufgaben": SetKV fields, "Aufgaben", valueNorm
        Case "w" & ChrW$(246) & "chentlicher umfang": SetKV fields, "Woechentlicher_Umfang", valueNorm
        Case "umfang am st" & ChrW$(252) & "ck", "umfang am stueck": SetKV fields, "Umfang_am_Stueck", valueNorm
        Case "abrechnung " & ChrW$(252) & "ber bet.- & entlastungsleistungen", "abrechnung ueber bet.- & entlastungsleistungen": SetKV fields, "Abrechnung_Betreuungsleistungen", valueNorm
        Case "pflegedienst vorhanden": SetKV fields, "Pflegedienst_Vorhanden", valueNorm
        Case "anfragedetails": SetKV fields, "Anfragedetails", valueNorm
        Case "anfragen-nr", "anfragen-nr.", "anfragen nr": SetKV fields, "Anfrage_ID", valueNorm
        Case "weitere details"
            SetKV fields, "Weitere_Details", valueNorm
            ' Eingebettete Key:Value-Paare parsen (z.B. "Pflegegrad/-stufe: 1")
            If InStr(valueNorm, ":") > 0 Then
                MapInlinePair fields, valueNorm, sectionName
            End If
        Case "bedarf": SetKV fields, "Bedarf", valueNorm
        Case "anfragedetails": SetKV fields, "Anfragedetails", valueNorm
        Case "anfragen-nr:": SetKV fields, "Anfrage_ID", valueNorm
        Case "id", "d": SetKV fields, "Anfrage_ID", valueNorm
        Case "budgetrahmen": SetKV fields, "Budgetrahmen", valueNorm
        Case "geschlecht der betreuungskraft": SetKV fields, "Geschlecht_Betreuungskraft", valueNorm
        Case Else
            ' Nicht erkanntes Label loggen fuer Debugging
            Debug.Print "[MapLabel] UNBEKANNT: '" & keyNorm & "' -> '" & Left$(valueNorm, 60) & "' [" & sectionName & "]"
    End Select
End Sub

Private Function IsNoiseLine(ByVal lineText As String) As Boolean
    ' Zweck: Erkennt ob eine Zeile Rauschen ist (E-Mail-Header, URLs, Footer).
    ' Rückgabe: True wenn Zeile uebersprungen werden soll.
    Dim lt As String
    lt = LCase$(Trim$(lineText))

    ' E-Mail-Weiterleitungs-Header (Von:, Datum:, An:, Betreff: etc.)
    If Left$(lt, 4) = "von:" Or Left$(lt, 6) = "datum:" Or Left$(lt, 3) = "an:" _
       Or Left$(lt, 8) = "betreff:" Or Left$(lt, 10) = "gesendet:" _
       Or Left$(lt, 5) = "from:" Or Left$(lt, 5) = "date:" Or Left$(lt, 3) = "to:" _
       Or Left$(lt, 8) = "subject:" Or Left$(lt, 5) = "sent:" _
       Or Left$(lt, 3) = "cc:" Or Left$(lt, 4) = "bcc:" Then
        ' Nur als Noise werten wenn E-Mail-Muster erkennbar (< oder @ oder WG:/Fwd:)
        If InStr(lineText, "<") > 0 Or InStr(lineText, "@") > 0 _
           Or InStr(1, lineText, "WG:", vbTextCompare) > 0 _
           Or InStr(1, lineText, "Fwd:", vbTextCompare) > 0 _
           Or InStr(1, lineText, "AW:", vbTextCompare) > 0 Then
            IsNoiseLine = True
            Exit Function
        End If
        ' Datum-Zeilen mit Wochentag erkennen (z.B. "Datum: Freitag, 23. Januar")
        If Left$(lt, 6) = "datum:" Or Left$(lt, 5) = "date:" Then
            If InStr(1, lineText, "um ", vbTextCompare) > 0 _
               Or InStr(1, lineText, "at ", vbTextCompare) > 0 Then
                IsNoiseLine = True
                Exit Function
            End If
        End If
    End If

    ' URL-Zeilen (http://, https://, www.)
    If Left$(lt, 7) = "http://" Or Left$(lt, 8) = "https://" Or Left$(lt, 4) = "www." Then
        IsNoiseLine = True
        Exit Function
    End If

    ' [Text]<URL> Muster (SendGrid-Tracking-Links)
    If Left$(lt, 1) = "[" And InStr(lt, "]<http") > 0 Then
        IsNoiseLine = True
        Exit Function
    End If

    ' Impressum/Footer-Muster
    If Left$(lt, 18) = "registergericht:" _
       Or Left$(lt, 24) = "sitz der gesellschaft:" _
       Or InStr(1, lt, "umsatzsteuer-identifikationsnummer", vbTextCompare) > 0 _
       Or InStr(1, lt, "geschäftsführer:", vbTextCompare) > 0 _
       Or InStr(1, lt, "gesch" & ChrW$(228) & "ftsf" & ChrW$(252) & "hrer:", vbTextCompare) > 0 _
       Or InStr(1, lt, "anfragen-manager", vbTextCompare) > 0 _
       Or InStr(1, lt, "anfrage reklamieren", vbTextCompare) > 0 _
       Or InStr(1, lt, "datenschutz<", vbTextCompare) > 0 _
       Or InStr(1, lt, "partnerbetreuung@", vbTextCompare) > 0 Then
        IsNoiseLine = True
        Exit Function
    End If

    ' Zustimmungs-/Datenschutz-Zeilen
    If InStr(1, lt, "zustimmung zur datenschutz", vbTextCompare) > 0 _
       Or InStr(1, lt, "zustimmung zur kontaktweitergabe", vbTextCompare) > 0 Then
        IsNoiseLine = True
        Exit Function
    End If

    ' Inline-Zeilen die nur Zeitstempel enthalten (DD.MM.YYYY HH:MM)
    ' z.B. "21.01.2026 10:04 Uhr" -> kein nützliches Feld
    If Len(lt) <= 25 And InStr(lt, ".") > 0 And InStr(lt, "uhr") > 0 Then
        IsNoiseLine = True
        Exit Function
    End If

    ' Zeilen mit <tel:...> oder <mailto:...> Pattern im Footer (nicht in Kontakt-Sektion)
    ' Diese enthalten Pflegehilfe-Hotline-Nummern etc.
    If Left$(lt, 1) >= "0" And Left$(lt, 1) <= "9" And InStr(lt, "<tel:") > 0 Then
        IsNoiseLine = True
        Exit Function
    End If

    ' E-Mail-Adressen mit <mailto:> Pattern (Footer-Kontaktadressen)
    If InStr(lt, "@") > 0 And InStr(lt, "<mailto:") > 0 Then
        ' Nur als Noise werten wenn NICHT ein bekanntes Label davor steht
        If InStr(1, lt, "e-mail", vbTextCompare) = 0 _
           And InStr(1, lt, "email", vbTextCompare) = 0 _
           And InStr(1, lt, "mail-adresse", vbTextCompare) = 0 Then
            IsNoiseLine = True
            Exit Function
        End If
    End If

    ' Inline-URLs mit <https://...> oder <http://...> Pattern
    If InStr(lt, "<https://") > 0 Or InStr(lt, "<http://") > 0 Then
        IsNoiseLine = True
        Exit Function
    End If
End Function

Private Function NormalizeKey(ByVal rawKey As String) As String
    ' Zweck: Schlüsseltext vereinheitlichen.
    ' Abhängigkeiten: Trim$, Replace, LCase$.
    ' Rückgabe: normalisierter Label-Text.
    Dim k As String
    k = rawKey
    ' NBSP und andere Whitespace-Zeichen normalisieren
    k = Replace(k, ChrW$(160), " ")
    k = Replace(k, vbTab, " ")
    k = LCase$(Trim$(k))
    ' Mehrfache Spaces komprimieren
    Do While InStr(1, k, "  ") > 0
        k = Replace(k, "  ", " ")
    Loop
    NormalizeKey = k
End Function

Private Function StripNamePrefix(ByVal nameText As String) As String
    Dim s As String
    Dim sLower As String

    s = Trim$(nameText)
    sLower = LCase$(s)

    If Left$(sLower, 5) = "herr " Then
        s = Trim$(Mid$(s, 6))
    ElseIf Left$(sLower, 6) = "herrn " Then
        s = Trim$(Mid$(s, 7))
    ElseIf Left$(sLower, 5) = "frau " Then
        s = Trim$(Mid$(s, 6))
    End If

    StripNamePrefix = s
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
    noteText = "Automatischer Import vom: " & Format$(Now, "dd.mm.yy hh.nn") & " | Quelle: "
    If Len(Trim$(gLeadSourceNote)) > 0 Then
        noteText = noteText & gLeadSourceNote
    Else
        noteText = noteText & LEAD_SOURCE
    End If

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

    ' Debug: Alle Tabellen-Header auflisten
    Debug.Print "[AddLeadRow] Tabelle '" & tbl.Name & "' hat " & tbl.ListColumns.Count & " Spalten:"
    Dim hdrIdx As Long
    For hdrIdx = 1 To tbl.ListColumns.Count
        Debug.Print "  Spalte " & hdrIdx & ": '" & tbl.ListColumns(hdrIdx).Name & "'"
    Next hdrIdx

    Dim nameVal As String
    Dim phoneVal As String
    Dim plzVal As String
    Dim pgVal As String
    Dim notesVal As String

    nameVal = ResolveKontaktName(fields)
    phoneVal = CleanPhoneNumber(GetField(fields, "Kontakt_Mobil"))
    plzVal = CleanPostalCode(GetField(fields, "PLZ"))
    pgVal = NormalizePflegegrad(GetField(fields, "Senior_Pflegegrad"))
    notesVal = BuildNotes(fields)

    ' Betreff als erste Zeile in Notizen einfuegen (fuer Zuordnung bei fehlerhaftem Parsing)
    Dim subjectVal As String
    subjectVal = GetField(fields, "MailSubject")
    If Len(subjectVal) > 0 Then
        notesVal = "Betreff: " & subjectVal & vbLf & notesVal
    End If

    Debug.Print "[AddLeadRow] Werte -> Name='" & nameVal & "' Tel='" & phoneVal & "' PLZ='" & plzVal & "' PG='" & pgVal & "'"

    SetCellByHeaderMap newRow, headerMap, "Monat Lead erhalten", DateSerial(Year(msgDate), Month(msgDate), 1)
    Set monthCell = GetCellByHeaderMap(newRow, headerMap, "Monat Lead erhalten")
    SetImportNote monthCell
    SetCellByHeaderMap newRow, headerMap, "Lead-Quelle", ResolveLeadSource(fields)
    SetCellByHeaderMap newRow, headerMap, "Leadtyp", leadType
    SetCellByHeaderMap newRow, headerMap, "Status", "Lead erhalten"
    SetCellByHeaderMap newRow, headerMap, "Name", nameVal
    SetCellByHeaderMap newRow, headerMap, "Telefonnummer", phoneVal
    SetCellByHeaderMap newRow, headerMap, "PLZ", plzVal
    SetCellByHeaderMap newRow, headerMap, "PG", pgVal
    SetCellByHeaderMap newRow, headerMap, "Notizen", notesVal

    ' Adresse / Ort / ID in passende Spalten schreiben
    Dim anschriftVal As String
    anschriftVal = GetField(fields, "Kontakt_Anschrift")
    If Len(anschriftVal) > 0 Then SetCellByHeaderMap newRow, headerMap, "Adresse", anschriftVal

    Dim ortVal As String
    ortVal = GetField(fields, "Bedarfsort_Ort")
    If Len(ortVal) > 0 Then SetCellByHeaderMap newRow, headerMap, "Ort", ortVal

    Dim idVal As String
    idVal = GetField(fields, "Anfrage_ID")
    If Len(idVal) > 0 Then SetCellByHeaderMap newRow, headerMap, "ID", idVal

    ' Info-Spalte: Body mit Zeilenumbruechen in die Zelle schreiben (nur wenn Spalte existiert)
    Dim infoIdx As Long
    infoIdx = GetHeaderIndex(headerMap, "Info")
    If infoIdx > 0 Then
        Dim infoCell As Range
        Dim infoBody As String
        Dim infoSubject As String
        infoBody = GetField(fields, "MailBody")
        infoBody = Replace(infoBody, vbCrLf, vbLf)
        infoBody = Replace(infoBody, vbCr, vbLf)
        infoSubject = GetField(fields, "MailSubject")
        If Len(infoSubject) > 0 Then
            infoBody = "Betreff: " & infoSubject & vbLf & vbLf & infoBody
        End If
        SetCellByHeaderMap newRow, headerMap, "Info", infoBody
        Set infoCell = GetCellByHeaderMap(newRow, headerMap, "Info")
        If Not infoCell Is Nothing Then infoCell.WrapText = True
    End If
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
    ' Rückgabe: Bereinigter Absendername oder LEAD_SOURCE.
    Dim fromVal As String
    Dim sourceName As String

    fromVal = GetField(fields, "From")

    If Len(Trim$(fromVal)) = 0 Then
        ResolveLeadSource = LEAD_SOURCE
        Exit Function
    End If

    ' Absendernamen extrahieren (ohne E-Mail-Adresse in spitzen Klammern)
    sourceName = ExtractSenderName(fromVal)

    ' Falls Ergebnis immer noch eine E-Mail-Adresse ist, Domain extrahieren
    If InStr(sourceName, "@") > 0 Then
        Dim atPos As Long
        Dim domainPart As String
        Dim dotPos As Long
        atPos = InStr(sourceName, "@")
        domainPart = Mid$(sourceName, atPos + 1)
        dotPos = InStrRev(domainPart, ".")
        If dotPos > 1 Then
            domainPart = Left$(domainPart, dotPos - 1)
        End If
        sourceName = domainPart
    End If

    If Len(Trim$(sourceName)) = 0 Then
        ResolveLeadSource = LEAD_SOURCE
    Else
        ResolveLeadSource = Trim$(sourceName)
    End If
End Function

Private Sub SetCellByHeaderMap(ByVal rowItem As ListRow, ByVal headerMap As Object, ByVal headerName As String, ByVal valueToSet As Variant)
    ' Zweck: Zellwert anhand vorberechneter Header-Map setzen.
    ' Abhängigkeiten: GetHeaderIndex.
    ' Rückgabe: keine (schreibt in Zeile).
    Dim idx As Long
    idx = GetHeaderIndex(headerMap, headerName)
    If idx > 0 Then
        rowItem.Range.Cells(1, idx).Value = valueToSet
    Else
        Debug.Print "[SetCell] WARNUNG: Spalte '" & headerName & "' nicht in Tabelle gefunden! Wert='" & Left$(CStr(valueToSet), 40) & "'"
    End If
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

    fullName = StripNamePrefix(Trim$(fullName))
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
    notes = AppendNote(notes, "Anschrift", GetField(fields, "Kontakt_Anschrift"))
    notes = AppendNote(notes, "Budgetrahmen", GetField(fields, "Budgetrahmen"))
    notes = AppendNote(notes, "Geschlecht Betreuungskraft", GetField(fields, "Geschlecht_Betreuungskraft"))
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

Private Function CleanWhitespace(ByVal textIn As String) As String
    ' Zweck: Whitespaces normalisieren (inkl. NBSP) und trimmen.
    Dim s As String
    s = Replace(textIn, ChrW$(160), " ")
    s = Replace(s, vbTab, " ")
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    CleanWhitespace = Trim$(s)
End Function

Private Function CleanPhoneNumber(ByVal rawValue As String) As String
    ' Zweck: Telefonnummer bereinigen.
    ' Entfernt <tel:...> Suffix und URL-Encoding (%20 etc.).
    ' Beispiel: "+49 16097008155<tel:+49%2016097008155>" -> "+49 16097008155"
    Dim s As String
    Dim sLower As String
    Dim closePos As Long
    Dim p As Long
    s = Trim$(rawValue)

    ' Sonderfall: reiner tel-Link ohne sichtbaren Anzeigetext
    sLower = LCase$(s)
    If Left$(sLower, 5) = "tel:" Then
        s = Mid$(s, 5)
    ElseIf Left$(sLower, 6) = "<tel:" Then
        closePos = InStr(1, s, ">")
        If closePos > 0 Then
            s = Mid$(s, 6, closePos - 6)
        Else
            s = Mid$(s, 6)
        End If
    End If

    ' URL-Encoding auflösen
    s = Replace(s, "%20", " ", , , vbTextCompare)
    s = Replace(s, "%2B", "+", , , vbTextCompare)

    ' <tel:...> Suffix entfernen
    p = InStr(1, s, "<tel:", vbTextCompare)
    If p > 0 Then
        s = Left$(s, p - 1)
    End If
    ' <mailto:...> Suffix entfernen (falls versehentlich)
    p = InStr(1, s, "<mailto:", vbTextCompare)
    If p > 0 Then
        s = Left$(s, p - 1)
    End If
    ' Allgemeines <...> Suffix entfernen
    p = InStr(1, s, "<")
    If p > 0 Then
        s = Left$(s, p - 1)
    End If
    CleanPhoneNumber = Trim$(s)
End Function

Private Function CleanLinkedValue(ByVal rawValue As String) As String
    ' Zweck: Werte mit angehaengten Link-Tags bereinigen.
    ' Entfernt <tel:...>, <mailto:...> und sonstige <...> Suffixe.
    ' Beispiel: "m.kaiser@meggy.com<mailto:m.kaiser@meggy.com>" -> "m.kaiser@meggy.com"
    Dim s As String
    Dim sLower As String
    Dim closePos As Long
    Dim p As Long
    s = Trim$(rawValue)

    sLower = LCase$(s)

    ' Sonderfall: Wert besteht nur aus Linkpayload (tel/mailto)
    If Left$(sLower, 5) = "tel:" Then
        s = Mid$(s, 5)
    ElseIf Left$(sLower, 6) = "<tel:" Then
        closePos = InStr(1, s, ">")
        If closePos > 0 Then
            s = Mid$(s, 6, closePos - 6)
        Else
            s = Mid$(s, 6)
        End If
    ElseIf Left$(sLower, 7) = "mailto:" Then
        s = Mid$(s, 8)
    ElseIf Left$(sLower, 8) = "<mailto:" Then
        closePos = InStr(1, s, ">")
        If closePos > 0 Then
            s = Mid$(s, 9, closePos - 9)
        Else
            s = Mid$(s, 9)
        End If
    End If

    ' URL-Encoding auflösen
    s = Replace(s, "%20", " ", , , vbTextCompare)
    s = Replace(s, "%2B", "+", , , vbTextCompare)

    p = InStr(1, s, "<")
    If p > 1 Then
        s = Left$(s, p - 1)
    End If
    CleanLinkedValue = Trim$(s)
End Function

Private Function CleanPostalCode(ByVal rawValue As String) As String
    ' Zweck: PLZ bereinigen (Whitespace/NBSP entfernen, bevorzugt nur Ziffern).
    Dim cleaned As String
    Dim digits As String

    cleaned = CleanWhitespace(rawValue)
    digits = FilterDigits(cleaned)
    If Len(digits) > 0 Then
        CleanPostalCode = digits
    Else
        CleanPostalCode = cleaned
    End If
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

' =========================
' Diagnose-Tool: Fehlende Felder erkennen
' =========================
Public Sub DebugDumpFields(ByVal fields As Object, Optional ByVal subject As String = "")
    ' Zweck: Alle erwarteten Felder ausgeben und fehlende markieren.
    ' Aufruf: Nach ParseLeadContent im Direktfenster (Ctrl+G).
    Dim fieldNames As Variant
    Dim fieldLabels As Variant
    Dim i As Long
    Dim v As String
    Dim missingCount As Long

    fieldNames = Array("Kontakt_Anrede", "Kontakt_Vorname", "Kontakt_Nachname", _
                       "Kontakt_Name", "Kontakt_Mobil", "Kontakt_Email", _
                       "Kontakt_Erreichbarkeit", "Kontakt_Anschrift", _
                       "Senior_Name", "Senior_Beziehung", "Senior_Alter", _
                       "Senior_Pflegegrad", "Senior_Pflegegrad_Status", _
                       "Senior_Lebenssituation", "Senior_Mobilitaet", _
                       "Senior_Medizinisches", "Senior_Behinderung", _
                       "PLZ", "Bedarfsort_Ort", "Nutzer", _
                       "Aufgaben", "Alltagshilfe_Aufgaben", "Alltagshilfe_Haeufigkeit", _
                       "Woechentlicher_Umfang", "Umfang_am_Stueck", _
                       "Abrechnung_Betreuungsleistungen", "Pflegedienst_Vorhanden", _
                       "Anfragedetails", "Weitere_Details", "Bedarf", _
                       "Anfrage_ID", "Budgetrahmen", "Geschlecht_Betreuungskraft", _
                       "From", "MailBody")

    fieldLabels = Array("Anrede", "Vorname", "Nachname", _
                        "Kontakt-Name", "Mobil/Telefon", "E-Mail", _
                        "Erreichbarkeit", "Anschrift", _
                        "Senior Name", "Beziehung", "Alter", _
                        "Pflegegrad", "PG Status", _
                        "Lebenssituation", "Mobilit" & ChrW$(228) & "t", _
                        "Medizinisches", "Behinderung", _
                        "PLZ", "Bedarfsort Ort", "Nutzer", _
                        "Aufgaben", "Alltagshilfe Aufgaben", "Alltagshilfe H" & ChrW$(228) & "ufigkeit", _
                        "W" & ChrW$(246) & "chentl. Umfang", "Umfang am St" & ChrW$(252) & "ck", _
                        "Abrechnung Betreuung", "Pflegedienst", _
                        "Anfragedetails", "Weitere Details", "Bedarf", _
                        "Anfrage-ID", "Budgetrahmen", "Geschlecht BK", _
                        "Absender (From)", "Mail-Body")

    missingCount = 0
    Debug.Print ""
    Debug.Print "=============================="
    Debug.Print "FELD-DIAGNOSE"
    If Len(subject) > 0 Then Debug.Print "Betreff: " & subject
    Debug.Print "=============================="
    For i = LBound(fieldNames) To UBound(fieldNames)
        v = GetField(fields, CStr(fieldNames(i)))
        If Len(Trim$(v)) > 0 Then
            ' OK-Felder nicht ausgeben - nur fehlende zeigen
        Else
            Debug.Print "  LEER : " & CStr(fieldLabels(i)) & " (" & CStr(fieldNames(i)) & ")"
            ' Nur Kernfelder als fehlend zaehlen (Name, Telefon, E-Mail)
            ' Vorname/Nachname nur zaehlen wenn Kontakt_Name auch leer ist
            Select Case CStr(fieldNames(i))
                Case "Kontakt_Name", "Kontakt_Mobil", "Kontakt_Email"
                    missingCount = missingCount + 1
                Case "Kontakt_Vorname"
                    If Len(Trim$(GetField(fields, "Kontakt_Name"))) = 0 Then
                        missingCount = missingCount + 1
                    End If
            End Select
        End If
    Next i
    Debug.Print "------------------------------"
    If missingCount > 0 Then
        Debug.Print "WARNUNG: " & missingCount & " Kern-Kontaktfelder fehlen!"
        Debug.Print "Tipp: Body-Text im Direktfenster pruefen (MailBody-Feld oben)"
    Else
        Debug.Print "Alle Kern-Kontaktfelder vorhanden."
    End If
    Debug.Print "=============================="
    Debug.Print ""
End Sub

' =========================
' Smoke Test
' =========================
Public Sub SmokeTestExtraction()
    ' Zweck: Automatisierter Smoke-Test der Extraktionspipeline.
    ' Prüft 6 Mail-Formate: Klartext, HTML-Tabelle, QP-kodiert, Soft-Hyphen, Inline, HTML+Entities.
    ' Ausgabe: Debug.Print (Direktfenster Ctrl+G).
    Dim passCount As Long
    Dim failCount As Long
    Dim totalCount As Long

    passCount = 0
    failCount = 0

    Debug.Print "=============================="
    Debug.Print "SMOKE TEST - Extraktion"
    Debug.Print "=============================="
    Debug.Print ""

    ' --- Test 1: Klartext (Label/Value auf separaten Zeilen) ---
    Dim body1 As String
    body1 = "Kontaktinformationen des Interessenten" & vbLf & _
            "Vor- und Nachname:" & vbLf & _
            "Max Mustermann" & vbLf & _
            "Mobil:" & vbLf & _
            "+49 171 1234567" & vbLf & _
            "E-Mail-Adresse:" & vbLf & _
            "max@beispiel.de" & vbLf & _
            "Erreichbarkeit:" & vbLf & _
            "vormittags" & vbLf & _
            "Bedarfsort:" & vbLf & _
            "80331 Muenchen" & vbLf & _
            "Informationen zum Senior" & vbLf & _
            "Name:" & vbLf & _
            "Erna Mustermann" & vbLf & _
            "Pflegegrad:" & vbLf & _
            "Pflegegrad 3" & vbLf & _
            "Anfragen-Nr.:" & vbLf & _
            "12345"

    totalCount = totalCount + RunSingleSmokeTest("1: Klartext Label/Value", body1, _
        "Max Mustermann", "+49 171 1234567", "max@beispiel.de", "vormittags", _
        "80331", "Erna Mustermann", "3", "12345", passCount, failCount)

    ' --- Test 2: HTML-Tabelle (SendGrid-Style) ---
    Dim body2 As String
    body2 = "<html><body>" & _
            "<h2>Neue Anfrage</h2>" & _
            "<table>" & _
            "<tr><td colspan='2'><b>Kontakt" & ChrW$(&HAD) & "informationen des Interessenten</b></td></tr>" & _
            "<tr><td>Vor- und Nachname:</td><td>Anna Schmidt</td></tr>" & _
            "<tr><td>Mobil:</td><td>+49 160 9876543<tel:+49%20160%209876543></td></tr>" & _
            "<tr><td>E-Mail-Adresse:</td><td>anna@schmidt.de<mailto:anna@schmidt.de></td></tr>" & _
            "<tr><td>Erreichbarkeit:</td><td>nachmittags</td></tr>" & _
            "<tr><td>Bedarfsort:</td><td>50667 K&ouml;ln</td></tr>" & _
            "<tr><td colspan='2'><b>Informationen zum Senior</b></td></tr>" & _
            "<tr><td>Name:</td><td>Helga Schmidt</td></tr>" & _
            "<tr><td>Pflegegrad:</td><td>Pflegegrad 2</td></tr>" & _
            "<tr><td>Anfragen-Nr.:</td><td>67890</td></tr>" & _
            "</table></body></html>"

    totalCount = totalCount + RunSingleSmokeTest("2: HTML-Tabelle (SendGrid)", body2, _
        "Anna Schmidt", "+49 160 9876543", "anna@schmidt.de", "nachmittags", _
        "50667", "Helga Schmidt", "2", "67890", passCount, failCount)

    ' --- Test 3: Quoted-Printable Body ---
    Dim body3 As String
    body3 = "Kontaktinformationen des Interessenten" & vbLf & _
            "Vor- und Nachname:" & vbLf & _
            "Peter M=C3=BCller" & vbLf & _
            "Mobil:" & vbLf & _
            "+49 152 5551234" & vbLf & _
            "E-Mail-Adresse:" & vbLf & _
            "peter@m=C3=BCller.de" & vbLf & _
            "Erreichbarkeit:" & vbLf & _
            "ganzt=C3=A4gig" & vbLf & _
            "Bedarfsort:" & vbLf & _
            "10115 Berlin" & vbLf & _
            "Informationen zum Senior" & vbLf & _
            "Name:" & vbLf & _
            "Gerda M=C3=BCller" & vbLf & _
            "Pflegegrad:" & vbLf & _
            "Pflegegrad 4" & vbLf & _
            "Anfragen-Nr.:" & vbLf & _
            "24680"

    totalCount = totalCount + RunSingleSmokeTest("3: Quoted-Printable", body3, _
        "Peter M" & ChrW$(252) & "ller", "+49 152 5551234", "peter@m" & ChrW$(252) & "ller.de", _
        "ganzt" & ChrW$(228) & "gig", "10115", "Gerda M" & ChrW$(252) & "ller", "4", "24680", _
        passCount, failCount)

    ' --- Test 4: Soft-Hyphen in Sektionsheader ---
    Dim body4 As String
    body4 = "Kontakt" & ChrW$(&HAD) & "informationen des Interessenten" & vbLf & _
            "Vor- und Nachname:" & vbLf & _
            "Lisa Weber" & vbLf & _
            "Mobil:" & vbLf & _
            "+49 176 9998877" & vbLf & _
            "E-Mail-Adresse:" & vbLf & _
            "lisa@weber.de" & vbLf & _
            "Informationen zum Senior" & vbLf & _
            "Name:" & vbLf & _
            "Otto Weber" & vbLf & _
            "Pflegegrad:" & vbLf & _
            "Pflegegrad 1"

    totalCount = totalCount + RunSingleSmokeTest("4: Soft-Hyphen Header", body4, _
        "Lisa Weber", "+49 176 9998877", "lisa@weber.de", "", _
        "", "Otto Weber", "1", "", passCount, failCount)

    ' --- Test 5: Inline-Format (key: value auf gleicher Zeile) ---
    Dim body5 As String
    body5 = "Kontaktinformationen" & vbLf & _
            "Vor- und Nachname: Karl Braun" & vbLf & _
            "Mobil: +49 151 2223344" & vbLf & _
            "E-Mail-Adresse: karl@braun.de" & vbLf & _
            "Erreichbarkeit: abends" & vbLf & _
            "Bedarfsort: 20095 Hamburg" & vbLf & _
            "Informationen zum Senior" & vbLf & _
            "Name: Gisela Braun" & vbLf & _
            "Pflegegrad: Pflegegrad 5" & vbLf & _
            "Anfragen-Nr.: 99999"

    totalCount = totalCount + RunSingleSmokeTest("5: Inline key: value", body5, _
        "Karl Braun", "+49 151 2223344", "karl@braun.de", "abends", _
        "20095", "Gisela Braun", "5", "99999", passCount, failCount)

    ' --- Test 6: HTML mit Entities (&#NNN; und benannte Entities) ---
    Dim body6 As String
    body6 = "<html><body>" & _
            "<div>Kontaktinformationen des Interessenten</div>" & _
            "<div>Vor- und Nachname:</div><div>Hans M&uuml;ller</div>" & _
            "<div>Mobil:</div><div>+49 170 1112233</div>" & _
            "<div>E-Mail-Adresse:</div><div>hans@mueller.de</div>" & _
            "<div>Bedarfsort:</div><div>60311 Frankfurt&nbsp;am&nbsp;Main</div>" & _
            "<div>Informationen zum Senior</div>" & _
            "<div>Name:</div><div>Inge M&uuml;ller</div>" & _
            "<div>Pflegegrad:</div><div>Pflegegrad 3</div>" & _
            "</body></html>"

    totalCount = totalCount + RunSingleSmokeTest("6: HTML + Entities", body6, _
        "Hans M" & ChrW$(252) & "ller", "+49 170 1112233", "hans@mueller.de", "", _
        "60311", "Inge M" & ChrW$(252) & "ller", "3", "", passCount, failCount)

    ' --- Test 7: SendGrid Bold-Pattern (*Label:*) mit tel:-Link ---
    Dim body7 As String
    body7 = "<html><body>" & _
            "<table>" & _
            "<tr><td colspan=""2""><b>Kontaktinformationen des Interessenten</b></td></tr>" & _
            "<tr><td><b>*Name:*</b></td></tr>" & _
            "<tr><td>Peggy Kaiser</td></tr>" & _
            "<tr><td><b>*<a href=""tel:+491601234567""></a>:*</b></td></tr>" & _
            "<tr><td><b>*E-Mail-Adresse:*</b></td></tr>" & _
            "<tr><td>peggy.kaiser@example.de</td></tr>" & _
            "<tr><td><b>*Erreichbarkeit:*</b></td></tr>" & _
            "<tr><td>Vormittags</td></tr>" & _
            "<tr><td><b>*Bedarfsort:*</b></td></tr>" & _
            "<tr><td>55128 Mainz</td></tr>" & _
            "<tr><td colspan=""2""><b>Informationen zum Senior</b></td></tr>" & _
            "<tr><td><b>*Name:*</b></td></tr>" & _
            "<tr><td>Helga Kaiser</td></tr>" & _
            "<tr><td>Pflegegrad/-stufe: 2</td></tr>" & _
            "</table>" & _
            "</body></html>"

    totalCount = totalCount + RunSingleSmokeTest("7: SendGrid Bold *Label:* + tel:", body7, _
        "Peggy Kaiser", "+491601234567", "peggy.kaiser@example.de", "Vormittags", _
        "55128", "Helga Kaiser", "2", "", passCount, failCount)

    ' --- Test 8: Mobil-Wert nur als <tel:...> (ohne sichtbaren Text) ---
    Dim body8 As String
    body8 = "Kontaktinformationen des Interessenten" & vbLf & _
            "Name:" & vbLf & _
            "Irene Schnell" & vbLf & _
            "Mobil:" & vbLf & _
            "<tel:+49%2016097008155>" & vbLf & _
            "E-Mail-Adresse:" & vbLf & _
            "ireneschnell@web.de" & vbLf & _
            "Informationen zum Senior" & vbLf & _
            "Name:" & vbLf & _
            "Irene Schnell"

    totalCount = totalCount + RunSingleSmokeTest("8: Mobil nur <tel:...>", body8, _
        "Irene Schnell", "+49 16097008155", "ireneschnell@web.de", "", _
        "", "Irene Schnell", "", "", passCount, failCount)

    ' --- Test 9: Inline Mobil mit <tel:...> ---
    Dim body9 As String
    body9 = "Kontaktinformationen" & vbLf & _
            "Name: Meggy Kaiser" & vbLf & _
            "Mobil: <tel:+49%2016097008155>" & vbLf & _
            "E-Mail-Adresse: m.kaiser@meggy-k-munich.com" & vbLf & _
            "Informationen zum Senior" & vbLf & _
            "Name: Helga Kaiser"

    totalCount = totalCount + RunSingleSmokeTest("9: Inline Mobil <tel:...>", body9, _
        "Meggy Kaiser", "+49 16097008155", "m.kaiser@meggy-k-munich.com", "", _
        "", "Helga Kaiser", "", "", passCount, failCount)

    ' --- Zusammenfassung ---
    Debug.Print ""
    Debug.Print "=============================="
    Debug.Print "ERGEBNIS: " & passCount & " PASS / " & failCount & " FAIL (von " & totalCount & " Pruefungen)"
    Debug.Print "=============================="

    If failCount = 0 Then
        MsgBox "Smoke Test BESTANDEN: Alle " & totalCount & " Pruefungen OK.", vbInformation, "Smoke Test"
    Else
        MsgBox "Smoke Test: " & failCount & " von " & totalCount & " Pruefungen FEHLGESCHLAGEN." & vbLf & _
               "Details im Direktfenster (Ctrl+G).", vbExclamation, "Smoke Test"
    End If
End Sub

Private Function RunSingleSmokeTest( _
        ByVal testName As String, _
        ByVal bodyText As String, _
        ByVal expectName As String, _
        ByVal expectMobil As String, _
        ByVal expectEmail As String, _
        ByVal expectErreichbarkeit As String, _
        ByVal expectPLZ As String, _
        ByVal expectSeniorName As String, _
        ByVal expectPflegegrad As String, _
        ByVal expectAnfrageID As String, _
        ByRef passCount As Long, _
        ByRef failCount As Long _
    ) As Long
    ' Zweck: Einzelnen Smoke-Test durchlaufen lassen.
    ' Rückgabe: Anzahl der Prüfungen (pass+fail inkrementiert via ByRef).
    Dim decoded As String
    Dim fields As Object
    Dim checks As Long

    Debug.Print "--- " & testName & " ---"

    decoded = DecodeBodyIfNeeded(bodyText)
    Set fields = ParseLeadContent(decoded)

    checks = 0
    checks = checks + AssertField(fields, "Kontakt_Name", expectName, testName, passCount, failCount)
    checks = checks + AssertField(fields, "Kontakt_Mobil", expectMobil, testName, passCount, failCount)
    checks = checks + AssertField(fields, "Kontakt_Email", expectEmail, testName, passCount, failCount)
    If Len(expectErreichbarkeit) > 0 Then
        checks = checks + AssertField(fields, "Kontakt_Erreichbarkeit", expectErreichbarkeit, testName, passCount, failCount)
    End If
    If Len(expectPLZ) > 0 Then
        checks = checks + AssertField(fields, "PLZ", expectPLZ, testName, passCount, failCount)
    End If
    checks = checks + AssertField(fields, "Senior_Name", expectSeniorName, testName, passCount, failCount)
    If Len(expectPflegegrad) > 0 Then
        checks = checks + AssertFieldContains(fields, "Senior_Pflegegrad", expectPflegegrad, testName, passCount, failCount)
    End If
    If Len(expectAnfrageID) > 0 Then
        checks = checks + AssertField(fields, "Anfrage_ID", expectAnfrageID, testName, passCount, failCount)
    End If

    Debug.Print ""
    RunSingleSmokeTest = checks
End Function

Private Function AssertField( _
        ByVal fields As Object, _
        ByVal fieldName As String, _
        ByVal expected As String, _
        ByVal testName As String, _
        ByRef passCount As Long, _
        ByRef failCount As Long _
    ) As Long
    ' Zweck: Einzelnes Feld prüfen (exakter Vergleich, trimmed).
    ' Rückgabe: 1 (Prüfung gezählt).
    Dim actual As String
    actual = Trim$(GetField(fields, fieldName))
    expected = Trim$(expected)

    If StrComp(actual, expected, vbTextCompare) = 0 Then
        Debug.Print "  PASS: " & fieldName & " = """ & actual & """"
        passCount = passCount + 1
    Else
        Debug.Print "  FAIL: " & fieldName & " erwartet """ & expected & """ aber """ & actual & """"
        failCount = failCount + 1
    End If
    AssertField = 1
End Function

Private Function AssertFieldContains( _
        ByVal fields As Object, _
        ByVal fieldName As String, _
        ByVal expectedPart As String, _
        ByVal testName As String, _
        ByRef passCount As Long, _
        ByRef failCount As Long _
    ) As Long
    ' Zweck: Prüft ob ein Feld den erwarteten Teilstring enthält.
    ' Rückgabe: 1 (Prüfung gezählt).
    Dim actual As String
    actual = Trim$(GetField(fields, fieldName))

    If InStr(1, actual, expectedPart, vbTextCompare) > 0 Then
        Debug.Print "  PASS: " & fieldName & " enthaelt """ & expectedPart & """ (= """ & actual & """)"
        passCount = passCount + 1
    Else
        Debug.Print "  FAIL: " & fieldName & " erwartet """ & expectedPart & """ enthalten, aber """ & actual & """"
        failCount = failCount + 1
    End If
    AssertFieldContains = 1
End Function
