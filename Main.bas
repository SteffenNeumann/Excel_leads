Attribute VB_Name = "Main"
Option Explicit

' ==============================================================
' Main.bas -- Lead-Import aus gespeicherten EML-Dateien
' --------------------------------------------------------------
' Version  : 2.8
' Datum    : 2026-04-19
' Autor    : Steffen
' --------------------------------------------------------------
' Liest .eml-Dateien direkt per VBA -- kein Python, kein Shell.
' MIME-Parsing, Base64-Decode und CSV-Auswertung in purem VBA.
'
' Unterstuetzte CSV-Typen:
'   - Alltagshilfe  (Felder: ID, Postleitzahl, Vorname, ...)
'   - Neue Anfrage  (Felder: RequestNumber, RequestZipCode, ...)
'
' Duplikat-Pruefung : ID-Spalte in Pipeline-Tabelle
' Umlaut-Sicherheit : Base64-Bytes -> Utf8BytesToString (pure VBA)
' Dictionary        : Pure VBA Collection (kein Scripting.Dictionary)
'
' Mac-Sandbox-Loesung (v2.6):
'   Open For Binary direkt versuchen. Falls Sandbox blockiert:
'   AppleScriptTask ruft Shell "cp" auf -> kopiert EML in Temp-Datei
'   mit ASCII-Name im gleichen Ordner -> Binary-Read ohne Dialog.
'   Benoetigt: MailReader.scpt in ~/Library/Application Scripts/com.microsoft.Excel/
'
' Changelog:
'   v2.8 | 2026-04-19 | Status='Lead erhalten', Leadquelle-Praefix entfernt
'   v2.7 | 2026-04-19 | Leadquelle aus From-Header statt Subject
'   v2.6 | 2026-04-19 | Sandbox-Fix: AppleScriptTask Shell-Copy statt GetOpenFilename
'                        Kein manueller Dialog mehr noetig (wie legacy_main)
'   v2.5 | 2026-04-19 | Sandbox-Fix: GetOpenFilename einmalig + CanReadFile-Check
'                        ZugriffErteilen() als eigenstaendiger Public Sub
'   v2.4 | 2026-04-19 | ReadEmlText via MacScript/cat (kein Sandbox-Dialog)
'   v2.3 | 2026-04-19 | Tempfile durch Utf8BytesToString ersetzt
'   v2.2 | 2026-04-19 | FindMimeBodyStart: Leerzeilen toleriert
'   v2.1 | 2026-04-18 | Scripting.Dictionary -> Collection (Mac-Fix)
'   v2.0 | 2026-04-18 | Pure VBA MIME-Parser, kein Shell-Zugriff
'   v1.1 | 2026-04-18 | Python/Perl via do shell script (obsolet)
' ==============================================================

' --- AppleScriptTask (Mac-Sandbox-Workaround) ---
Private Const APPLESCRIPT_FILE    As String = "MailReader.scpt"
Private Const APPLESCRIPT_HANDLER As String = "FetchMessages"

' --- Pfad-Einstellung (wird aus Sheet "Berechnung", Named Range "mailpath" gelesen) ---
Private Const SETTINGS_SHEET   As String = "Berechnung"
Private Const NAME_MAILPATH    As String = "mailpath"

' --- Tabelle ---
Private Const SHEET_NAME As String = "Pipeline"
Private Const TABLE_NAME As String = "Kundenliste"

' --- Pipeline-Spaltenbezeichnungen ---
Private Const C_ID       As String = "ID"
Private Const C_ERHALTEN As String = "Lead erhalten"
Private Const C_PLZ      As String = "PLZ"
Private Const C_STATUS   As String = "Status"
Private Const C_QUELLE   As String = "Lead-Quelle"
Private Const C_NAME     As String = "Name"
Private Const C_ADRESSE  As String = "Adresse"
Private Const C_ORT      As String = "Ort"
Private Const C_TELEFON  As String = "Telefonnummer"
Private Const C_PG       As String = "PG"
Private Const C_NOTIZEN  As String = "Notizen"

' KV-Store: Internes Schluessellisten-Feld (fuer Diagnose-Iteration)
Private Const KV_KEYLIST As String = "__KEYS__"

' ==============================================================
' PFAD-EINSTELLUNG (aus Sheet "Berechnung", Named Range "mailpath")
' ==============================================================

Private Function GetMailsFolder() As String
    ' Liest den Mails-Pfad aus dem Named Range "mailpath" im Berechnung-Sheet.
    ' Entfernt fuehrende/abschliessende Anf" & ChrW(252) & "hrungszeichen (z.B. wenn Nutzer
    ' den Pfad mit Quotes eingibt: '"/Users/.../Mails"').
    Dim ws  As Worksheet
    Dim rng As Range
    Dim raw As String

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SETTINGS_SHEET)
    On Error GoTo 0

    If ws Is Nothing Then
        MsgBox "Einstellungs-Sheet '" & SETTINGS_SHEET & "' nicht gefunden!", _
               vbCritical, "Konfigurationsfehler"
        Exit Function
    End If

    On Error Resume Next
    Set rng = ws.Range(NAME_MAILPATH)
    On Error GoTo 0

    If rng Is Nothing Then
        MsgBox "Benannter Bereich '" & NAME_MAILPATH & "' nicht gefunden!" & vbLf & _
               "Bitte im Sheet '" & SETTINGS_SHEET & "' anlegen.", _
               vbCritical, "Konfigurationsfehler"
        Exit Function
    End If

    raw = Trim$(CStr(rng.Value))

    ' Anfuehrungszeichen entfernen (vorne und hinten)
    Do While Left$(raw, 1) = Chr(34) Or Left$(raw, 1) = "'"
        raw = Mid$(raw, 2)
    Loop
    Do While Right$(raw, 1) = Chr(34) Or Right$(raw, 1) = "'"
        raw = Left$(raw, Len(raw) - 1)
    Loop

    ' Abschliessenden Slash entfernen
    If Right$(raw, 1) = "/" Then raw = Left$(raw, Len(raw) - 1)

    GetMailsFolder = Trim$(raw)
End Function

' ==============================================================
' KV-STORE -- Pure VBA Collection (ersetzt Scripting.Dictionary)
' ==============================================================
Private Function KVNew() As Collection
    Set KVNew = New Collection
End Function

Private Sub KVSet(col As Collection, key As String, val As String)
    Dim keyList As String

    On Error Resume Next
    col.Remove key
    On Error GoTo 0
    col.Add val, key

    ' Schluessel in Keylist nachfuehren (nur fuer Nutzdaten)
    If key = KV_KEYLIST Then Exit Sub
    On Error Resume Next
    keyList = CStr(col(KV_KEYLIST))
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        col.Add key, KV_KEYLIST
    Else
        On Error GoTo 0
        col.Remove KV_KEYLIST
        col.Add keyList & "," & key, KV_KEYLIST
    End If
    On Error GoTo 0
End Sub

Private Function KVExists(col As Collection, key As String) As Boolean
    Dim tmp As String
    On Error Resume Next
    tmp = CStr(col(key))
    KVExists = (Err.Number = 0)
    On Error GoTo 0
End Function

Private Function KVGet(col As Collection, key As String) As String
    On Error Resume Next
    KVGet = CStr(col(key))
    If Err.Number <> 0 Then
        Err.Clear
        KVGet = ""
    End If
    On Error GoTo 0
End Function

' ==============================================================
' EINSTIEGSPUNKT
' ==============================================================

Public Sub ImportLeadsFromMailFolder()
    Dim ws          As Worksheet
    Dim tbl         As ListObject
    Dim kv          As Collection
    Dim fields      As Collection
    Dim leadId      As String
    Dim nameVal     As String
    Dim imported    As Long
    Dim skipped     As Long
    Dim emlFile     As String
    Dim emlPath     As String
    Dim emlPaths()  As String
    Dim pathCount   As Long
    Dim mailsFolder As String

    ' --- Pfad aus Einstellungen lesen ---
    mailsFolder = GetMailsFolder()
    If Len(mailsFolder) = 0 Then Exit Sub   ' Fehler wurde in GetMailsFolder gemeldet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "Blatt '" & SHEET_NAME & "' nicht gefunden.", vbCritical, "Fehler"
        Exit Sub
    End If

    Set tbl = FindTable(ws, TABLE_NAME)
    If tbl Is Nothing Then
        MsgBox "Tabelle '" & TABLE_NAME & "' nicht gefunden.", vbCritical, "Fehler"
        Exit Sub
    End If

    ' --- Alle EML-Pfade sammeln ---
    pathCount = 0
    emlFile = Dir$(mailsFolder & "/*.eml")
    Do While Len(emlFile) > 0
        ReDim Preserve emlPaths(pathCount)
        emlPaths(pathCount) = mailsFolder & "/" & emlFile
        pathCount = pathCount + 1
        emlFile = Dir$
    Loop

    If pathCount = 0 Then
        MsgBox "Keine .eml-Dateien in:" & vbLf & mailsFolder, vbExclamation, "Lead-Import"
        Exit Sub
    End If

    ' Mac-Sandbox: kein vorab-Dialog noetig.
    ' ReadEmlText nutzt AppleScriptTask-Fallback (Shell-Copy) wenn
    ' Open For Binary scheitert -- kein Sandbox-Popup pro Datei.

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim pathIdx As Long
    For pathIdx = 0 To pathCount - 1
        emlPath = emlPaths(pathIdx)

        Set kv     = ParseEmlToKv(emlPath)
        Set fields = BuildLeadFields(kv)

        leadId  = KVGet(fields, "id")
        nameVal = KVGet(fields, "name")

        If Len(leadId) = 0 And Len(nameVal) = 0 Then
            ' Keine verwertbaren Daten -- ueberspringen

        ElseIf LeadAlreadyExists(leadId, tbl) Then
            skipped = skipped + 1

        Else
            AddLeadRow fields, tbl
            imported = imported + 1
        End If

    Next pathIdx

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

    MsgBox "Import abgeschlossen:" & vbLf & _
           imported & " Leads neu importiert" & vbLf & _
           skipped & " Duplikate " & ChrW(252) & "bersprungen", _
           vbInformation, "Lead-Import"
End Sub

' ==============================================================
' SCHRITT 1 -- EML direkt per VBA parsen (kein Shell-Prozess)
' ==============================================================

Private Function ParseEmlToKv(emlPath As String) As Collection
    Dim kv       As Collection
    Dim raw      As String
    Dim boundary As String
    Dim csvText  As String

    Set kv = KVNew()

    raw = ReadEmlText(emlPath)
    If Len(raw) = 0 Then Set ParseEmlToKv = kv: Exit Function

    KVSet kv, "_Subject", GetHeaderValue(raw, "Subject")
    KVSet kv, "_Date",    GetHeaderValue(raw, "Date")
    KVSet kv, "_From",    GetHeaderValue(raw, "From")

    boundary = GetMimeBoundary(raw)
    If Len(boundary) = 0 Then Set ParseEmlToKv = kv: Exit Function

    csvText = GetCsvAttachment(raw, boundary)
    If Len(csvText) = 0 Then Set ParseEmlToKv = kv: Exit Function

    ParseCsvIntoDict csvText, kv

    Set ParseEmlToKv = kv
End Function

' --- EML-Datei einlesen ---
' Strategie (Mac):
'   AppleScriptTask Shell-Copy in Temp-Datei im Workbook-Verzeichnis
'   (dort hat VBA Sandbox-Zugriff). Open For Binary wird auf Mac
'   NICHT versucht -- das loest sonst den Sandbox-Dialog aus.
' Windows: direkt Binary-Read.
Private Function ReadEmlText(filePath As String) As String
    Dim fileNum    As Integer
    Dim fileLen    As Long
    Dim rawBytes() As Byte
    Dim result     As String
    Dim i          As Long

    #If Mac Then
        ' Mac: direkt Shell-Copy (Open For Binary wuerde Sandbox-Dialog ausloesen)
        result = ReadEmlViaShellCopy(filePath)
        If Len(result) > 0 Then
            result = Replace(result, vbCrLf, vbLf)
            result = Replace(result, vbCr,   vbLf)
            ReadEmlText = result
        End If
        Exit Function
    #End If

    ' Windows: Binary-Read direkt
    fileNum = FreeFile()
    On Error Resume Next
    Open filePath For Binary Access Read As #fileNum
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    fileLen = LOF(fileNum)
    If fileLen = 0 Then Close #fileNum: Exit Function

    ReDim rawBytes(0 To fileLen - 1)
    Get #fileNum, , rawBytes
    Close #fileNum

    ' Bytes 1:1 als Latin-1 -- Base64 und EML-Header sind 7-Bit-ASCII-sicher
    result = String$(fileLen, " ")
    For i = 0 To fileLen - 1
        Mid$(result, i + 1, 1) = Chr(rawBytes(i))
    Next i

    result = Replace(result, vbCrLf, vbLf)
    result = Replace(result, vbCr,   vbLf)
    ReadEmlText = result
End Function

' --- Mac-Fallback: Shell-Copy ueber AppleScriptTask ---
' Kopiert die Datei per "cp" in eine Temp-Datei im Workbook-Verzeichnis
' (dort hat VBA immer Sandbox-Zugriff, kein Dialog).
' Danach Binary-Read der Temp-Datei, anschliessend Kill.
Private Function ReadEmlViaShellCopy(filePath As String) As String
    Dim tmpPath    As String
    Dim wbFolder   As String
    Dim script     As String
    Dim cpResult   As String
    Dim fileNum    As Integer
    Dim fileLen    As Long
    Dim rawBytes() As Byte
    Dim result     As String
    Dim i          As Long

    ' Temp-Datei im Workbook-Verzeichnis (Sandbox-sicher)
    wbFolder = ThisWorkbook.Path
    If Right$(wbFolder, 1) <> "/" Then wbFolder = wbFolder & "/"
    tmpPath = wbFolder & "_tmp_eml_import.eml"

    ' AppleScript: Shell-Copy in Temp-Pfad
    ' Pfade mit einfachen Anf.zeichen escapen (Shell-sicher).
    ' Eventuelles ' im Pfad wird zu '\'' escaped.
    Dim srcEsc As String
    Dim dstEsc As String
    srcEsc = Replace(filePath, "'", "'\''")
    dstEsc = Replace(tmpPath, "'", "'\''")
    script = "do shell script ""cp '" & srcEsc & "' '" & dstEsc & "'""" 

    On Error Resume Next
    cpResult = AppleScriptTask(APPLESCRIPT_FILE, APPLESCRIPT_HANDLER, script)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    ' Pruefen ob Shell-Ergebnis ein Fehler ist
    If Left$(cpResult, 6) = "ERROR:" Then Exit Function

    ' Temp-Datei lesen (ASCII-Name -> kein Sandbox-Dialog)
    fileNum = FreeFile()
    On Error Resume Next
    Open tmpPath For Binary Access Read As #fileNum
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    fileLen = LOF(fileNum)
    If fileLen = 0 Then Close #fileNum: GoTo Cleanup

    ReDim rawBytes(0 To fileLen - 1)
    Get #fileNum, , rawBytes
    Close #fileNum

    result = String$(fileLen, " ")
    For i = 0 To fileLen - 1
        Mid$(result, i + 1, 1) = Chr(rawBytes(i))
    Next i

    ReadEmlViaShellCopy = result

Cleanup:
    On Error Resume Next
    Kill tmpPath
    On Error GoTo 0
End Function

' --- Header-Wert auslesen (inkl. Folding ueber mehrere Zeilen) ---
Private Function GetHeaderValue(raw As String, headerName As String) As String
    Dim searchKey As String
    Dim pos       As Long
    Dim startPos  As Long
    Dim lineEnd   As Long
    Dim nextChar  As String
    Dim result    As String

    searchKey = vbLf & headerName & ":"
    pos = InStr(1, raw, searchKey, vbTextCompare)

    If pos > 0 Then
        startPos = pos + Len(searchKey)
    ElseIf StrComp(Left$(raw, Len(headerName) + 1), headerName & ":", vbTextCompare) = 0 Then
        startPos = Len(headerName) + 2
    Else
        Exit Function
    End If

    Do
        lineEnd = InStr(startPos, raw, vbLf)
        If lineEnd = 0 Then lineEnd = Len(raw) + 1
        result = result & Mid$(raw, startPos, lineEnd - startPos)
        If lineEnd < Len(raw) Then
            nextChar = Mid$(raw, lineEnd + 1, 1)
            If nextChar = " " Or nextChar = vbTab Then
                result   = result & " "
                startPos = lineEnd + 2
            Else
                Exit Do
            End If
        Else
            Exit Do
        End If
    Loop

    GetHeaderValue = Trim$(result)
End Function

' --- MIME-Boundary aus Content-Type Header extrahieren ---
Private Function GetMimeBoundary(raw As String) As String
    Dim ctHeader As String
    Dim boundPos As Long
    Dim quoteEnd As Long
    Dim charEnd  As Long
    Dim curChar  As String

    ctHeader = GetHeaderValue(raw, "Content-Type")
    If Len(ctHeader) = 0 Then Exit Function

    boundPos = InStr(1, ctHeader, "boundary=", vbTextCompare)
    If boundPos = 0 Then Exit Function

    boundPos = boundPos + Len("boundary=")

    If Mid$(ctHeader, boundPos, 1) = """" Then
        quoteEnd = InStr(boundPos + 1, ctHeader, """")
        If quoteEnd = 0 Then quoteEnd = Len(ctHeader) + 1
        GetMimeBoundary = Mid$(ctHeader, boundPos + 1, quoteEnd - boundPos - 1)
    Else
        charEnd = boundPos
        Do While charEnd <= Len(ctHeader)
            curChar = Mid$(ctHeader, charEnd, 1)
            If curChar = " " Or curChar = ";" Or curChar = vbLf Then Exit Do
            charEnd = charEnd + 1
        Loop
        GetMimeBoundary = Mid$(ctHeader, boundPos, charEnd - boundPos)
    End If
End Function

' --- MIME-Teil: echten Body-Start finden (ueberbrueckt Leerzeilen zwischen Headern) ---
' Hintergrund: Manche E-Mail-Clients fuegen zwischen MIME-Headern Leerzeilen ein.
' Ein simples InStr(vbLf & vbLf) trifft dann die ERSTE Leerzeile statt die letzte
' (echte Header/Body-Trennzeile). Diese Funktion prueft, ob nach einer Leerzeile
' noch ein weiterer Header (Format "Key: Value") folgt -- und ueberspringt ihn ggf.
' Gibt die 1-basierte Startposition des Bodys zurueck (0 = kein Body gefunden).
Private Function FindMimeBodyStart(partText As String) As Long
    Dim pos      As Long
    Dim lineEnd  As Long
    Dim lineText As String
    Dim peekPos  As Long
    Dim peekEnd  As Long
    Dim peekLine As String
    Dim colonP   As Long
    Dim nameStr  As String
    Dim isHdr    As Boolean
    Dim ci       As Long
    Dim c        As String

    pos = 1
    Do While pos <= Len(partText)
        lineEnd  = InStr(pos, partText, vbLf)
        If lineEnd = 0 Then lineEnd = Len(partText) + 1
        lineText = Mid$(partText, pos, lineEnd - pos)

        If Len(Trim$(lineText)) = 0 Then
            ' Leerzeile -- naechste Nicht-Leer-Zeile bestimmen
            peekPos = lineEnd + 1
            Do While peekPos <= Len(partText)
                peekEnd  = InStr(peekPos, partText, vbLf)
                If peekEnd = 0 Then peekEnd = Len(partText) + 1
                peekLine = Mid$(partText, peekPos, peekEnd - peekPos)
                If Len(Trim$(peekLine)) > 0 Then Exit Do
                peekPos = peekEnd + 1
            Loop

            ' Ende des Strings: Body beginnt nach dieser Leerzeile
            If peekPos > Len(partText) Then
                FindMimeBodyStart = lineEnd + 1
                Exit Function
            End If

            ' Ist die naechste Nicht-Leer-Zeile ein MIME-Header? (Key: Value)
            ' Base64-Daten enthalten kein ":" -- CSV-Rohzeilen selten vor dem
            ' ersten echten Trenner. Sicherheitscheck: nur [A-Za-z0-9-] vor ":"
            colonP  = InStr(Trim$(peekLine), ":")
            nameStr = Left$(Trim$(peekLine), IIf(colonP > 1, colonP - 1, 0))
            isHdr   = (colonP > 1 And Left$(Trim$(peekLine), 1) <> " " And _
                       Left$(Trim$(peekLine), 1) <> vbTab)
            If isHdr And Len(nameStr) > 0 Then
                For ci = 1 To Len(nameStr)
                    c = Mid$(nameStr, ci, 1)
                    If Not ((c >= "A" And c <= "Z") Or (c >= "a" And c <= "z") Or _
                            (c >= "0" And c <= "9") Or c = "-") Then
                        isHdr = False
                        Exit For
                    End If
                Next ci
            End If

            If Not isHdr Then
                ' Body beginnt nach dieser Leerzeile
                FindMimeBodyStart = lineEnd + 1
                Exit Function
            End If
            ' Sonst: Leerzeile liegt innerhalb des Header-Blocks -- weiter scannen
        End If

        pos = lineEnd + 1
    Loop

    FindMimeBodyStart = 0
End Function

' --- CSV-Anhang im MIME-Body finden und dekodieren ---
Private Function GetCsvAttachment(raw As String, boundary As String) As String
    Dim parts()    As String
    Dim partIdx    As Long
    Dim partText   As String
    Dim bodyStart  As Long
    Dim partHeader As String
    Dim partBody   As String
    Dim encPos     As Long
    Dim encEnd     As Long
    Dim encoding   As String

    parts = Split(raw, vbLf & "--" & boundary)

    For partIdx = 1 To UBound(parts)
        partText = parts(partIdx)

        If InStr(1, partText, ".csv", vbTextCompare) > 0 Or _
           InStr(1, partText, "text/csv", vbTextCompare) > 0 Then

            ' FindMimeBodyStart ueberbrueckt Leerzeilen zwischen MIME-Headern
            bodyStart = FindMimeBodyStart(partText)
            If bodyStart = 0 Then GoTo NextPart

            partHeader = Left$(partText, bodyStart - 1)
            partBody   = Trim$(Mid$(partText, bodyStart))

            encoding = ""
            encPos = InStr(1, partHeader, "Content-Transfer-Encoding:", vbTextCompare)
            If encPos > 0 Then
                encEnd   = InStr(encPos, partHeader, vbLf)
                If encEnd = 0 Then encEnd = Len(partHeader) + 1
                encoding = Trim$(Mid$(partHeader, _
                                      encPos + Len("Content-Transfer-Encoding:"), _
                                      encEnd - encPos - Len("Content-Transfer-Encoding:")))
            End If

            If LCase$(encoding) = "base64" Then
                GetCsvAttachment = Base64DecodeToString(partBody)
            Else
                GetCsvAttachment = partBody
            End If
            Exit Function
        End If
NextPart:
    Next partIdx
End Function

' --- Pure VBA UTF-8-Decoder: byteArr(startByte..byteCount-1) -> Unicode-String ---
' Ersetzt den alten Tempfile-Ansatz -- kein Schreibzugriff, kein Sandbox-Prompt.
' Unterstuetzt 1-Byte (ASCII), 2-Byte (Umlaute, Akzente) und 3-Byte (z.B. Euro-Zeichen).
' 4-Byte-Sequenzen (Emoji etc.) sind in deutschen Lead-CSVs nicht zu erwarten.
Private Function Utf8BytesToString(byteArr() As Byte, startByte As Long, byteCount As Long) As String
    Dim result As String
    Dim i      As Long
    Dim b1     As Long
    Dim b2     As Long
    Dim b3     As Long
    Dim cp     As Long

    i = startByte
    Do While i < byteCount
        b1 = byteArr(i)

        If b1 < &H80 Then
            ' 1-Byte ASCII (inkl. LF=10, CR=13)
            If b1 <> 13 Then result = result & Chr(b1)  ' CR ignorieren
            i = i + 1

        ElseIf (b1 And &HE0) = &HC0 Then
            ' 2-Byte-Sequenz (U+0080..U+07FF) -- Umlaute, Akzente
            If i + 1 < byteCount Then
                b2 = byteArr(i + 1)
                cp = ((b1 And &H1F) * 64) Or (b2 And &H3F)
                result = result & ChrW(cp)
                i = i + 2
            Else
                i = i + 1
            End If

        ElseIf (b1 And &HF0) = &HE0 Then
            ' 3-Byte-Sequenz (U+0800..U+FFFF) -- Euro-Zeichen u.a.
            If i + 2 < byteCount Then
                b2 = byteArr(i + 1)
                b3 = byteArr(i + 2)
                cp = ((b1 And &HF) * 4096) Or ((b2 And &H3F) * 64) Or (b3 And &H3F)
                result = result & ChrW(cp)
                i = i + 3
            Else
                i = i + 1
            End If

        Else
            ' 4-Byte oder ungueltig -- ueberspringen
            i = i + 1
        End If
    Loop

    Utf8BytesToString = result
End Function

' --- Pure VBA Base64-Decoder (kein Tempfile, kein Sandbox-Prompt) ---
Private Function Base64DecodeToString(b64 As String) As String
    Dim lookup    As String
    Dim cleaned   As String
    Dim charIdx   As Long
    Dim curChar   As String
    Dim byteArr() As Byte
    Dim byteCount As Long
    Dim nChars    As Long
    Dim v1        As Long
    Dim v2        As Long
    Dim v3        As Long
    Dim v4        As Long
    Dim startByte As Long

    lookup = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"

    For charIdx = 1 To Len(b64)
        curChar = Mid$(b64, charIdx, 1)
        If InStr(lookup, curChar) > 0 Or curChar = "=" Then
            cleaned = cleaned & curChar
        End If
    Next charIdx

    Do While Len(cleaned) Mod 4 <> 0
        cleaned = cleaned & "="
    Loop

    nChars = Len(cleaned)
    If nChars = 0 Then Exit Function

    ReDim byteArr(0 To (nChars \ 4) * 3 - 1)
    byteCount = 0

    For charIdx = 1 To nChars - 3 Step 4
        v1 = InStr(lookup, Mid$(cleaned, charIdx,     1)) - 1
        v2 = InStr(lookup, Mid$(cleaned, charIdx + 1, 1)) - 1
        v3 = InStr(lookup, Mid$(cleaned, charIdx + 2, 1)) - 1
        v4 = InStr(lookup, Mid$(cleaned, charIdx + 3, 1)) - 1

        byteArr(byteCount) = CByte((v1 * 4) Or (v2 \ 16))
        byteCount = byteCount + 1

        If Mid$(cleaned, charIdx + 2, 1) <> "=" Then
            byteArr(byteCount) = CByte(((v2 And 15) * 16) Or (v3 \ 4))
            byteCount = byteCount + 1
        End If

        If Mid$(cleaned, charIdx + 3, 1) <> "=" Then
            byteArr(byteCount) = CByte(((v3 And 3) * 64) Or v4)
            byteCount = byteCount + 1
        End If
    Next charIdx

    If byteCount = 0 Then Exit Function

    ' UTF-8 BOM entfernen (EF BB BF)
    startByte = 0
    If byteCount >= 3 Then
        If byteArr(0) = &HEF And byteArr(1) = &HBB And byteArr(2) = &HBF Then
            startByte = 3
        End If
    End If

    ' Pure VBA UTF-8-Decode -- kein Tempfile, kein Sandbox-Prompt
    Base64DecodeToString = Utf8BytesToString(byteArr, startByte, byteCount)
End Function

' --- CSV in KV-Store einlesen (Header-Zeile + erste Datenzeile) ---
Private Sub ParseCsvIntoDict(csvText As String, kv As Collection)
    Dim csvNorm     As String
    Dim lines()     As String
    Dim lineIdx     As Long
    Dim headerLine  As String
    Dim dataLine    As String
    Dim headers()   As String
    Dim values()    As String
    Dim colIdx      As Long
    Dim fieldName   As String
    Dim fieldVal    As String
    Dim headerFound As Boolean
    Dim dataFound   As Boolean

    csvNorm = Replace(Replace(csvText, vbCrLf, vbLf), vbCr, vbLf)
    lines   = Split(csvNorm, vbLf)

    For lineIdx = 0 To UBound(lines)
        If Len(Trim$(lines(lineIdx))) > 0 Then
            If Not headerFound Then
                headerLine  = lines(lineIdx)
                headerFound = True
            ElseIf Not dataFound Then
                dataLine  = lines(lineIdx)
                dataFound = True
                Exit For
            End If
        End If
    Next lineIdx

    If Not headerFound Or Not dataFound Then Exit Sub

    headers = ParseCsvLine(headerLine)
    values  = ParseCsvLine(dataLine)

    For colIdx = 0 To UBound(headers)
        fieldName = Trim$(headers(colIdx))
        If Len(fieldName) > 0 Then
            If colIdx <= UBound(values) Then
                fieldVal = Trim$(values(colIdx))
            Else
                fieldVal = ""
            End If
            KVSet kv, fieldName, fieldVal
        End If
    Next colIdx
End Sub

' --- CSV-Zeile in String-Array aufteilen (Quote-aware) ---
Private Function ParseCsvLine(csvLine As String) As String()
    Dim result() As String
    Dim fields   As New Collection
    Dim curField As String
    Dim inQuote  As Boolean
    Dim charIdx  As Long
    Dim curChar  As String
    Dim nextChar As String
    Dim fieldIdx As Long

    For charIdx = 1 To Len(csvLine)
        curChar = Mid$(csvLine, charIdx, 1)

        If inQuote Then
            If curChar = """" Then
                If charIdx < Len(csvLine) Then
                    nextChar = Mid$(csvLine, charIdx + 1, 1)
                    If nextChar = """" Then
                        curField = curField & """"
                        charIdx  = charIdx + 1
                    Else
                        inQuote = False
                    End If
                Else
                    inQuote = False
                End If
            Else
                curField = curField & curChar
            End If
        Else
            If curChar = """" Then
                inQuote = True
            ElseIf curChar = "," Then
                fields.Add curField
                curField = ""
            Else
                curField = curField & curChar
            End If
        End If
    Next charIdx

    fields.Add curField

    ReDim result(0 To fields.Count - 1)
    For fieldIdx = 1 To fields.Count
        result(fieldIdx - 1) = fields(fieldIdx)
    Next fieldIdx

    ParseCsvLine = result
End Function

' ==============================================================
' SCHRITT 2 -- CSV-Felder auf Pipeline-Spalten mappen
' ==============================================================

Private Function BuildLeadFields(kv As Collection) As Collection
    Dim fields   As Collection
    Dim subject  As String
    Dim fromHdr  As String
    Dim mailDate As String
    Dim id       As String
    Dim vorname  As String
    Dim nachname As String
    Dim plz      As String
    Dim telefon  As String
    Dim pg       As String
    Dim adresse  As String
    Dim ort      As String
    Dim notizen  As String
    Dim nameVal  As String
    Dim rel      As String
    Dim detail   As String
    Dim nutzer   As String
    Dim aufgab   As String
    Dim haeuf    As String

    Set fields = KVNew()

    subject  = KVGet(kv, "_Subject")
    fromHdr  = KVGet(kv, "_From")
    mailDate = KVGet(kv, "_Date")

    ' -- Typ: Neue Anfrage ------------------------------------------
    If KVExists(kv, "RequestNumber") Then

        id       = KVGet(kv, "RequestNumber")
        vorname  = KVGet(kv, "FirstName")
        nachname = KVGet(kv, "SurName")
        plz      = KVGet(kv, "RequestZipCode")
        telefon  = KVGet(kv, "Phone")
        pg       = NormalizePflegegrad(KVGet(kv, "SeniorCareLevel"))
        adresse  = KVGet(kv, "AddressLine1")
        ort      = KVGet(kv, "City")
        If Len(ort) = 0 Then ort = KVGet(kv, "RequestRegion")

        rel    = KVGet(kv, "SeniorRelationship")
        detail = KVGet(kv, "RequestDetail")
        If Len(rel) > 0 Then notizen = "Nutzer: " & rel
        If Len(detail) > 0 Then
            If Len(notizen) > 0 Then notizen = notizen & vbLf
            notizen = notizen & detail
        End If

    ' -- Typ: Alltagshilfe ------------------------------------------
    ElseIf KVExists(kv, "Postleitzahl") Or KVExists(kv, "ID") Then

        id       = KVGet(kv, "ID")
        vorname  = KVGet(kv, "Vorname")
        nachname = KVGet(kv, "Nachname")
        plz      = KVGet(kv, "Postleitzahl")
        telefon  = KVGet(kv, "Telefonnummer")
        pg       = NormalizePflegegrad(KVGet(kv, "Pflegegrad"))
        adresse  = ""
        ort      = ""

        nutzer = KVGet(kv, "Nutzer")
        aufgab = KVGet(kv, "Alltagshilfe Aufgaben")
        haeuf  = KVGet(kv, "Alltagshilfe H" & ChrW(228) & "ufigkeit")

        If Len(nutzer) > 0 Then notizen = "Nutzer: " & nutzer
        If Len(aufgab) > 0 Then
            If Len(notizen) > 0 Then notizen = notizen & vbLf
            notizen = notizen & "Aufgaben: " & aufgab
        End If
        If Len(haeuf) > 0 Then
            If Len(notizen) > 0 Then notizen = notizen & vbLf
            notizen = notizen & "H" & ChrW(228) & "ufigkeit: " & haeuf
        End If

    End If

    ' Name: "Nachname, Vorname"
    If Len(nachname) > 0 And Len(vorname) > 0 Then
        nameVal = nachname & ", " & vorname
    ElseIf Len(nachname) > 0 Then
        nameVal = nachname
    Else
        nameVal = vorname
    End If

    KVSet fields, "id",         id
    KVSet fields, "mail_date",  mailDate
    KVSet fields, "plz",        plz
    KVSet fields, "leadquelle", ExtractFromName(fromHdr)
    KVSet fields, "name",       nameVal
    KVSet fields, "adresse",    adresse
    KVSet fields, "ort",        ort
    KVSet fields, "telefon",    telefon
    KVSet fields, "pg",         pg
    KVSet fields, "notizen",    notizen

    Set BuildLeadFields = fields
End Function

' ==============================================================
' SCHRITT 3 -- Neue Zeile in Pipeline-Tabelle schreiben
' ==============================================================

Private Sub AddLeadRow(fields As Collection, tbl As ListObject)
    Dim newRow   As ListRow
    Dim hIdx     As Collection
    Dim mailDate As Date

    Set hIdx   = BuildHIdx(tbl)
    Set newRow = tbl.ListRows.Add(AlwaysInsert:=True)

    mailDate = ParseMailDate(KVGet(fields, "mail_date"))

    SetCellDate newRow, hIdx, C_ERHALTEN, mailDate
    SetCell     newRow, hIdx, C_ID,       KVGet(fields, "id")
    SetCell     newRow, hIdx, C_STATUS,   "Lead erhalten"
    SetCell     newRow, hIdx, C_QUELLE,   KVGet(fields, "leadquelle")
    SetCell     newRow, hIdx, C_NAME,     KVGet(fields, "name")
    SetCell     newRow, hIdx, C_PLZ,      KVGet(fields, "plz")
    SetCell     newRow, hIdx, C_TELEFON,  KVGet(fields, "telefon")
    SetCell     newRow, hIdx, C_PG,       KVGet(fields, "pg")
    SetCell     newRow, hIdx, C_NOTIZEN,  KVGet(fields, "notizen")

    If Len(KVGet(fields, "adresse")) > 0 Then SetCell newRow, hIdx, C_ADRESSE, KVGet(fields, "adresse")
    If Len(KVGet(fields, "ort"))     > 0 Then SetCell newRow, hIdx, C_ORT,     KVGet(fields, "ort")
End Sub

' ==============================================================
' DUPLIKAT-PRUEFUNG (per ID-Spalte)
' ==============================================================

Private Function LeadAlreadyExists(leadId As String, tbl As ListObject) As Boolean
    Dim hIdx     As Collection
    Dim idColIdx As Long
    Dim dataRng  As Range
    Dim cell     As Range

    If Len(Trim$(leadId)) = 0 Then Exit Function

    Set hIdx = BuildHIdx(tbl)
    If Not KVExists(hIdx, LCase$(C_ID)) Then Exit Function
    idColIdx = CLng(KVGet(hIdx, LCase$(C_ID)))

    On Error Resume Next
    Set dataRng = tbl.ListColumns(idColIdx).DataBodyRange
    On Error GoTo 0
    If dataRng Is Nothing Then Exit Function

    For Each cell In dataRng
        If StrComp(Trim$(CStr(cell.Value)), leadId, vbTextCompare) = 0 Then
            LeadAlreadyExists = True
            Exit Function
        End If
    Next cell
End Function

' ==============================================================
' HILFSFUNKTIONEN
' ==============================================================

Private Function FindTable(ws As Worksheet, tblName As String) As ListObject
    Dim lo As ListObject
    For Each lo In ws.ListObjects
        If lo.Name = tblName Then Set FindTable = lo: Exit Function
    Next lo
End Function

Private Function BuildHIdx(tbl As ListObject) As Collection
    Dim hIdx As Collection
    Dim col  As ListColumn

    Set hIdx = KVNew()
    For Each col In tbl.ListColumns
        KVSet hIdx, LCase$(col.Name), CStr(col.Index)
    Next col
    Set BuildHIdx = hIdx
End Function

Private Sub SetCell(row As ListRow, hIdx As Collection, colName As String, val As String)
    Dim colKey As String

    colKey = LCase$(colName)
    If Not KVExists(hIdx, colKey) Then Exit Sub
    On Error Resume Next
    row.Range.Cells(1, CLng(KVGet(hIdx, colKey))).Value = val
    On Error GoTo 0
End Sub

Private Sub SetCellDate(row As ListRow, hIdx As Collection, colName As String, val As Date)
    Dim colKey As String

    colKey = LCase$(colName)
    If Not KVExists(hIdx, colKey) Then Exit Sub
    On Error Resume Next
    With row.Range.Cells(1, CLng(KVGet(hIdx, colKey)))
        .Value        = val
        .NumberFormat = "DD.MM.YYYY"
    End With
    On Error GoTo 0
End Sub

Private Function NormalizePflegegrad(s As String) As String
    ' "Pflegegrad 2" -> "2"  |  erste Ziffer extrahieren
    Dim charIdx As Long
    For charIdx = 1 To Len(s)
        If Mid$(s, charIdx, 1) >= "0" And Mid$(s, charIdx, 1) <= "9" Then
            NormalizePflegegrad = Mid$(s, charIdx, 1)
            Exit Function
        End If
    Next charIdx
End Function

Private Function ParseMailDate(dateStr As String) As Date
    ' RFC-2822: "Tue, 31 Mar 2026 07:06:12 +0000" -> Date
    Dim s        As String
    Dim commaPos As Long
    Dim parts()  As String
    Dim dayVal   As Long
    Dim monVal   As Long
    Dim yearVal  As Long

    ParseMailDate = Date

    s        = Trim$(dateStr)
    commaPos = InStr(s, ",")
    If commaPos > 0 Then s = Trim$(Mid$(s, commaPos + 1))

    parts = Split(s, " ")
    If UBound(parts) < 2 Then Exit Function

    On Error Resume Next
    dayVal  = CLng(parts(0))
    monVal  = MonthFromAbbr(parts(1))
    yearVal = CLng(parts(2))
    If Err.Number <> 0 Or monVal = 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    ParseMailDate = DateSerial(yearVal, monVal, dayVal)
End Function

Private Function MonthFromAbbr(abbr As String) As Long
    Select Case LCase$(Left$(Trim$(abbr), 3))
        Case "jan": MonthFromAbbr = 1
        Case "feb": MonthFromAbbr = 2
        Case "mar": MonthFromAbbr = 3
        Case "apr": MonthFromAbbr = 4
        Case "may": MonthFromAbbr = 5
        Case "jun": MonthFromAbbr = 6
        Case "jul": MonthFromAbbr = 7
        Case "aug": MonthFromAbbr = 8
        Case "sep": MonthFromAbbr = 9
        Case "oct": MonthFromAbbr = 10
        Case "nov": MonthFromAbbr = 11
        Case "dec": MonthFromAbbr = 12
        Case Else:  MonthFromAbbr = 0
    End Select
End Function

Private Function ExtractFromName(fromHdr As String) As String
    ' "PflegeHelfer24" <noreply@x.de>  -> PflegeHelfer24
    ' Anfragen - Verbund Pflegehilfe <anfragen@pflegehilfe.de> -> Anfragen - Verbund Pflegehilfe
    Dim ltPos As Long
    Dim s     As String

    s = Trim$(fromHdr)
    ltPos = InStr(s, "<")
    If ltPos > 1 Then
        s = Trim$(Left$(s, ltPos - 1))
    End If
    ' Anf&uuml;hrungszeichen entfernen
    If Left$(s, 1) = """" And Right$(s, 1) = """" And Len(s) > 1 Then
        s = Mid$(s, 2, Len(s) - 2)
    End If
    ' Praefix "Anfragen - " entfernen (z.B. Verbund Pflegehilfe)
    If Left$(s, 12) = "Anfragen - " Then
        s = Mid$(s, 13)
    End If
    ExtractFromName = s
End Function

Private Function TmpBase() As String
    Dim tmpDir As String
    tmpDir = Environ("TMPDIR")
    If Len(tmpDir) = 0 Then tmpDir = "/tmp/"
    If Right$(tmpDir, 1) <> "/" Then tmpDir = tmpDir & "/"
    TmpBase = tmpDir
End Function

' ==============================================================
' MAC-SANDBOX: ORDNER-ZUGRIFF (einmalig pro Session)
' ==============================================================

' CanReadFile und RequestFolderAccess entfernt (v2.6).
' Sandbox-Zugriff wird jetzt ueber AppleScriptTask Shell-Copy geloest
' (ReadEmlViaShellCopy) -- kein manueller Dialog mehr noetig.

' Public Sub: Zugriff manuell vorab erteilen (z.B. einmal nach Excel-Neustart).
' Danach kann ImportLeadsFromMailFolder ohne Dialoge laufen.
Public Sub ZugriffErteilen()
    ' Seit v2.6 nicht mehr noetig -- AppleScriptTask umgeht Sandbox.
    ' Bleibt als Stub fuer Rueckwaertskompatibilitaet.
    MsgBox "Seit v2.6 nicht mehr n" & ChrW(246) & "tig." & vbLf & _
           "Der Import nutzt jetzt AppleScriptTask und ben" & ChrW(246) & "tigt" & vbLf & _
           "keinen manuellen Zugriffsdialog mehr." & vbLf & vbLf & _
           "Voraussetzung: MailReader.scpt liegt in:" & vbLf & _
           "~/Library/Application Scripts/com.microsoft.Excel/", _
           vbInformation, "Info"
End Sub

' ==============================================================
' DIAGNOSE -- Cursor in Sub setzen, F5 druecken
' ==============================================================

Public Sub DiagnoseImport()
    Dim msg        As String
    Dim emlCount   As Long
    Dim umlCount   As Long
    Dim emlFile    As String
    Dim emlPath    As String
    Dim testKv     As Collection
    Dim ws         As Worksheet
    Dim tbl        As ListObject
    Dim hIdx       As Collection
    Dim colNames   As String
    Dim colCheck   As String
    Dim lc         As ListColumn
    Dim constNames As Variant
    Dim constIdx   As Long
    Dim constItem  As String
    Dim keyList    As String
    Dim keyParts() As String
    Dim keyStr     As String
    Dim maxKeys    As Long
    Dim keyIdx     As Long
    Dim mailsFolder As String

    msg = "=== Lead-Import Diagnose v2.8 ===" & vbLf & vbLf

    ' 0) Pfad aus Einstellungen lesen
    mailsFolder = GetMailsFolder()
    If Len(mailsFolder) = 0 Then
        msg = msg & "[0] mailpath: NICHT KONFIGURIERT (Sheet Berechnung pruefen)" & vbLf
        MsgBox msg, vbCritical, "Diagnose Lead-Import"
        Exit Sub
    End If
    msg = msg & "[0] mailpath      : " & mailsFolder & vbLf & vbLf

    ' 1) Mails-Ordner: alle EML zaehlen + Umlaut-Dateien erkennen
    On Error Resume Next
    emlFile = Dir$(mailsFolder & "/*.eml")
    Do While Len(emlFile) > 0
        emlCount = emlCount + 1
        If emlFile Like "*[" & ChrW(196) & ChrW(214) & ChrW(220) & _
                              ChrW(228) & ChrW(246) & ChrW(252) & ChrW(223) & "]*" Then
            umlCount = umlCount + 1
        End If
        emlFile = Dir$
    Loop
    On Error GoTo 0
    msg = msg & "[1] Mails-Ordner  : " & mailsFolder & vbLf
    msg = msg & "    EML gefunden  : " & emlCount & vbLf
    If umlCount > 0 Then
        msg = msg & "    davon Umlaut  : " & umlCount & " (Dir$ sieht diese)" & vbLf
    End If
    msg = msg & vbLf

    ' 2) Erste EML parsen und Ergebnis zeigen
    If emlCount > 0 Then
        emlFile = Dir$(mailsFolder & "/*.eml")
        emlPath = mailsFolder & "/" & emlFile
        Set testKv = ParseEmlToKv(emlPath)
        msg = msg & "[2] Parse-Test   : " & emlFile & vbLf
        msg = msg & "    Subject      : " & KVGet(testKv, "_Subject") & vbLf
        msg = msg & "    Date         : " & KVGet(testKv, "_Date") & vbLf
        keyList = KVGet(testKv, KV_KEYLIST)
        If Len(keyList) > 0 Then
            keyParts = Split(keyList, ",")
            keyStr   = ""
            maxKeys  = WorksheetFunction.Min(UBound(keyParts), 7)
            For keyIdx = 0 To maxKeys
                keyStr = keyStr & keyParts(keyIdx) & ", "
            Next keyIdx
            msg = msg & "    CSV-Felder   : " & (UBound(keyParts) + 1) & vbLf
            msg = msg & "    Feldnamen    : " & keyStr & vbLf
        Else
            msg = msg & "    CSV-Felder   : 0 (Parsing fehlgeschlagen!)" & vbLf
        End If
        msg = msg & vbLf
    End If

    ' 3) Tabelle pruefen
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)
    On Error GoTo 0
    If ws Is Nothing Then
        msg = msg & "[3] Blatt '" & SHEET_NAME & "': NICHT GEFUNDEN" & vbLf
        MsgBox msg, vbCritical, "Diagnose Lead-Import"
        Exit Sub
    End If
    msg = msg & "[3] Blatt '" & SHEET_NAME & "': OK" & vbLf

    Set tbl = FindTable(ws, TABLE_NAME)
    If tbl Is Nothing Then
        msg = msg & "[4] Tabelle '" & TABLE_NAME & "': NICHT GEFUNDEN" & vbLf
        MsgBox msg, vbCritical, "Diagnose Lead-Import"
        Exit Sub
    End If
    msg = msg & "[4] Tabelle '" & TABLE_NAME & "': OK (" & tbl.ListColumns.Count & " Spalten)" & vbLf & vbLf

    ' 5) Alle Spaltennamen anzeigen
    Set hIdx = BuildHIdx(tbl)
    colNames = "[5] Spaltennamen in Tabelle:" & vbLf
    For Each lc In tbl.ListColumns
        colNames = colNames & "    " & lc.Name & vbLf
    Next lc
    msg = msg & colNames & vbLf

    ' 6) Konstanten-Abgleich
    constNames = Array(C_ID, C_ERHALTEN, C_PLZ, C_STATUS, C_QUELLE, _
                       C_NAME, C_ADRESSE, C_ORT, C_TELEFON, C_PG, C_NOTIZEN)
    colCheck = "[6] Spalten-Konstanten (OK = gefunden):" & vbLf
    For constIdx = 0 To UBound(constNames)
        constItem = CStr(constNames(constIdx))
        If KVExists(hIdx, LCase$(constItem)) Then
            colCheck = colCheck & "    [OK] " & constItem & vbLf
        Else
            colCheck = colCheck & "    [XX] " & constItem & " <-- NICHT GEFUNDEN" & vbLf
        End If
    Next constIdx
    msg = msg & colCheck

    MsgBox msg, vbInformation, "Diagnose Lead-Import"
End Sub

Public Sub DiagnoseEmlContent()
    Dim emlFile  As String
    Dim emlPath  As String
    Dim raw      As String
    Dim boundary As String
    Dim parts()  As String
    Dim partIdx  As Long
    Dim maxPart  As Long
    Dim snippet  As String
    Dim msg      As String
    Dim testKv   As Collection
    Dim keyList  As String
    Dim keyArr() As String
    Dim keyIdx   As Long

    Dim mailsFolder As String
    mailsFolder = GetMailsFolder()
    If Len(mailsFolder) = 0 Then Exit Sub

    emlFile = Dir$(mailsFolder & "/*.eml")
    If Len(emlFile) = 0 Then
        MsgBox "Keine .eml-Dateien in: " & mailsFolder, vbExclamation
        Exit Sub
    End If

    emlPath  = mailsFolder & "/" & emlFile
    raw      = ReadEmlText(emlPath)
    boundary = GetMimeBoundary(raw)

    msg = "Datei    : " & emlFile & vbLf
    msg = msg & "Subject  : " & GetHeaderValue(raw, "Subject") & vbLf
    msg = msg & "From     : " & GetHeaderValue(raw, "From") & vbLf
    msg = msg & "Date     : " & GetHeaderValue(raw, "Date") & vbLf
    msg = msg & "Boundary : " & IIf(Len(boundary) > 0, boundary, "(kein multipart)") & vbLf & vbLf

    If Len(boundary) > 0 Then
        parts = Split(raw, vbLf & "--" & boundary)
        ' Teile zaehlen: parts(0)=Praeambel, letzte=schliessende Boundary "--"
        Dim realParts As Long
        For partIdx = 1 To UBound(parts)
            If Left$(Trim$(parts(partIdx)), 2) <> "--" Then
                realParts = realParts + 1
            End If
        Next partIdx
        msg = msg & "MIME-Teile: " & realParts & vbLf
        Dim diagNum As Long
        diagNum = 0
        For partIdx = 1 To UBound(parts)
            ' Schliessende Boundary ("--\n...") ueberspringen
            If Left$(Trim$(parts(partIdx)), 2) = "--" Then GoTo SkipDiagPart
            diagNum = diagNum + 1
            If diagNum > 4 Then GoTo SkipDiagPart
            snippet = Left$(Trim$(parts(partIdx)), 150)
            msg = msg & "--- Teil " & diagNum & ":" & vbLf & snippet & vbLf
            SkipDiagPart:
        Next partIdx
        msg = msg & vbLf
    End If

    msg = msg & "=== ParseEmlToKv ===" & vbLf
    Set testKv = ParseEmlToKv(emlPath)
    msg = msg & "_Subject : " & KVGet(testKv, "_Subject") & vbLf
    msg = msg & "_From    : " & KVGet(testKv, "_From") & vbLf
    msg = msg & "_Date    : " & KVGet(testKv, "_Date") & vbLf

    keyList = KVGet(testKv, KV_KEYLIST)
    If Len(keyList) > 0 Then
        keyArr = Split(keyList, ",")
        msg    = msg & "Felder   : " & (UBound(keyArr) + 1) & vbLf
        For keyIdx = 0 To UBound(keyArr)
            msg = msg & CStr(keyArr(keyIdx)) & ": " & KVGet(testKv, CStr(keyArr(keyIdx))) & vbLf
        Next keyIdx
    Else
        msg = msg & "Felder   : 0 (kein CSV-Anhang erkannt)" & vbLf
    End If

    MsgBox msg, vbInformation, "EML-Diagnose: " & emlFile
End Sub
