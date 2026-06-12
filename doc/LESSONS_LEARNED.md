# Lessons Learned – Excel Leads Projekt

Dieses Log dokumentiert Fehler, die im Projektverlauf aufgetreten sind, damit sie sich nicht wiederholen.

---

## LL-001 · 2026-04-18 · KRITISCH: openpyxl zerstört .xlsm-Datei

### Was ist passiert?
Ein Python-Skript mit `openpyxl` hat die Produktionsdatei `Pipeline-Leads.xlsm` beschädigt.
Die Datei konnte danach nicht mehr geöffnet werden. Wiederherstellung war nur über ein Backup möglich.

### Ursache
`openpyxl` unterstützt das **Lesen** von `.xlsm`-Dateien, aber **nicht das vollständige Schreiben**.
Beim Speichern via `workbook.save()` werden folgende Excel-Bestandteile gelöscht oder korrumpiert:

- VBA-Makro-Module (`.bas`, `.cls`, `.frm`)
- Named Ranges / benannte Bereiche
- Datenvalidierungen (`Data Validation`)
- Tabellenformatierungen (ListObject-Stile)
- Worksheet-Events und Workbook-Events

Zusätzlich war die Datei zum Zeitpunkt des Skript-Aufrufs **in Excel geöffnet**, was zu einer Schreibkollision führte.

### Konsequenzen
- Datei war korrupt und nicht öffenbar
- Alle VBA-Module mussten aus dem Git-Repo neu eingespielt werden
- Datenvalidierungen mussten neu konfiguriert werden

### Regeln (dauerhaft gültig für dieses Projekt)

| # | Regel | Begründung |
|---|-------|-----------|
| 🚫 1 | **Kein `openpyxl`-Write auf `.xlsm`** | VBA + Metadaten werden gelöscht |
| 🚫 2 | **Keine externe Dateiänderung bei geöffneter Datei** | Schreibkollision → Korruption |
| ✅ 3 | **Änderungen an `.xlsm` nur via VBA** (innerhalb Excel) | Einziger sicherer Write-Pfad |
| ✅ 4 | **Vor jedem externen Skript: Backup in `Backup/`** | Fallback sicherstellen |
| ✅ 5 | **openpyxl nur mit `read_only=True`** wenn überhaupt | Lesen ist sicher, Schreiben nicht |

### Erlaubte Alternativen für externe Datenweitergabe

```
Extern → Excel:
  Python schreibt in separate .xlsx (kein Makro) → VBA liest diese ein

Excel → Extern:
  VBA exportiert CSV → Python verarbeitet CSV

Skript soll Makro auslösen:
  osascript -e 'tell application "Microsoft Excel" to run macro "MeinMakro"'
```

### Betroffene Dateien
- `Excel_files/Pipeline-Leads.xlsm` (Produktionsdatei)

---

---

## LL-002 · 2026-06-11 · xlwings Web Extension verursacht "Fehler beim Speichern"

### Was ist passiert?
Excel zeigte beim Speichern von `Pipeline-Leads-26_06_11.xlsm` den Dialog:
> „Durch Entfernen oder Reparieren einiger Features kann die Datei von Microsoft Excel möglicherweise gespeichert werden."

### Ursache
Das xlwings-Add-in hatte beim aktiven Einsatz Metadaten in die `.xlsm`-Datei eingebettet:

| Datei im ZIP | Inhalt |
|---|---|
| `xl/webextensions/webextension1.xml` | xlwings Add-in (Store-ID `wa200008175`) mit UDFs `_xldudf_WINGMAN`, `_xldudf_CORREL2` etc. |
| `xl/webextensions/taskpanes.xml` | Taskpane mit `visibility="1"` (automatisch öffnen) |
| `xl/webextensions/_rels/taskpanes.xml.rels` | Querverweise |
| `xl/workbook.xml` (Named Ranges) | `_xleta.ISNUMBER` und `_xleta.TODAY` mit Wert `#NAME?` – broken UDF-Cache |

Excel versucht beim Speichern, die Taskpane-Extension aufzulösen. Schlägt das fehl (kein aktives xlwings-Backend), bricht der Speichervorgang ab.

### Diagnose
```python
import zipfile
with zipfile.ZipFile("Pipeline-Leads.xlsm", "r") as z:
    with z.open("xl/workbook.xml") as f:
        print(f.read().decode())
    # => _xleta.ISNUMBER / _xleta.TODAY mit #NAME?
    # => webextension1.xml vorhanden
```

### Fix (Python ZIP-Manipulation – Datei muss geschlossen sein)
```python
import zipfile, shutil, re, os

# Backup anlegen
shutil.copy2(src, bak)

DROP = {
    "xl/webextensions/webextension1.xml",
    "xl/webextensions/taskpanes.xml",
    "xl/webextensions/_rels/taskpanes.xml.rels",
}

with zipfile.ZipFile(src, "r") as zin, zipfile.ZipFile(tmp, "w", ZIP_DEFLATED) as zout:
    for item in zin.infolist():
        if item.filename in DROP:
            continue
        data = zin.read(item.filename)
        if item.filename == "_rels/.rels":
            data = re.sub(r'\s*<Relationship[^>]*webextensiontaskpanes[^>]*/>', "", data.decode()).encode()
        if item.filename == "[Content_Types].xml":
            data = re.sub(r'\s*<Override[^>]*webextension[^>]*/>', "", data.decode()).encode()
        if item.filename == "xl/workbook.xml":
            data = re.sub(r'\s*<definedName name="_xleta\.[^"]*"[^>]*/>', "", data.decode()).encode()
        zout.writestr(item, data)
os.replace(tmp, src)
```

### Regeln

| # | Regel | Begründung |
|---|---|---|
| 🚫 1 | **xlwings nach Gebrauch aus der .xlsm entfernen** (oder nie aktivieren) | Hinterlässt webextension-Metadaten selbst wenn Python-UDFs nicht mehr genutzt werden |
| 🚫 2 | **Nie "Weiter" im Excel-Reparatur-Dialog klicken** ohne Backup | Excel legt eine neue Datei an und löscht ggf. Features ohne Rückmeldung |
| ✅ 3 | **Diagnose zuerst via Python/zipfile** | Zeigt exakt welche Features Excel stören |
| ✅ 4 | **Fix via ZIP-Manipulation** (Datei in Excel geschlossen!) | Sauberster Weg – VBA-Module bleiben unangetastet |

### Betroffene Dateien
- `Excel_files/Pipeline-Leads-26_06_11.xlsm`
- Backup: `Excel_files/Pipeline-Leads-26_06_11_BACKUP_before_xlwings_fix.xlsm`

---

## LL-003 · 2026-06-11 · KRITISCH: `run script` in AppleScript schlägt bei Umlaut-Pfaden fehl

### Was ist passiert?
`ReadEmlViaShellCopy` konnte EML-Dateien mit Umlauten im Dateinamen nicht lesen (Err 75).
Betroffen waren alle 5 Dateien mit `ö`, `ü` oder `ß` im Dateinamen:
```
Neue Anfrage_ Hans Höbel.eml
Neue Anfrage_ Heidemarie Mühl.eml
Neue Anfrage_ Hülya Yasin.eml
Neue Anfrage_ Veronika Rückschloß.eml
WG_ Neue Anfrage_ Heidemarie Mühl.eml
```

### Ursache
`MailReader.scpt` verwendete den Handler `FetchMessages(scriptText)` mit `run script scriptText`.
`run script` **kompiliert den übergebenen String als AppleScript-Quellcode zur Laufzeit**.
Enthält der String ein Umlaut-Zeichen als String-Literal (eingebettet durch VBA), schlägt die Compilation fehl:

```
Syntax error (-2741): Zeilenende erwartet, Identifier gefunden
```

**Test, der den Bug beweist:**
```python
# Direkt via osascript -e → funktioniert
osascript -e 'do shell script "cp " & quoted form of "/Pfad/Höbel.eml" & " " & quoted form of "/tmp/out.eml"'
# → cp erfolgreich

# Via run script (wie FetchMessages es macht) → FEHLER
osascript -e 'run script "do shell script \"cp \" & quoted form of \"/Pfad/Höbel.eml\" ..."'
# → execution error: syntax error -2741
```

**Warum:** macOS APFS speichert Dateinamen in NFD-Form (`Höbel`). VBA-`Dir$` gibt den Namen möglicherweise in NFC zurück (`Höbel` = U+00F6). Beim Einbetten als Literal in AppleScript-Quellcode interpretiert der AppleScript-Compiler das Nicht-ASCII-Zeichen als fehlerhaftes Token.

### Fix
**Dedizierte Handler ohne `run script`** in `MailReader.applescript` – Pfade werden als Parameter übergeben, nie als Source-Code-Literale kompiliert:

```applescript
-- NEU: Pfad kommt als Parameter, quoted form of wird direkt angewendet
on CopyFile(params)
    set delimPos to offset of "|" in params
    set srcPath to text 1 thru (delimPos - 1) of params
    set dstPath to text (delimPos + 1) thru -1 of params
    try
        do shell script "cp " & quoted form of srcPath & " " & quoted form of dstPath
        return "OK"
    on error errMsg number errNum
        return "ERROR:" & errNum & ":" & errMsg
    end try
end CopyFile

on RemoveXattr(folderPath)
    try
        do shell script "xattr -rd com.apple.quarantine " & quoted form of folderPath
        return "OK"
    on error errMsg number errNum
        return "ERROR:" & errNum & ":" & errMsg
    end try
end RemoveXattr
```

In VBA (`Main.bas` v3.2):
```vba
' ALT (fehlerhaft bei Umlauten):
script = "do shell script ""cp "" & quoted form of " & Chr(34) & filePath & Chr(34) & ...
cpResult = AppleScriptTask(APPLESCRIPT_FILE, "FetchMessages", script)

' NEU (korrekt):
cpResult = AppleScriptTask(APPLESCRIPT_FILE, "CopyFile", filePath & "|" & tmpPath)
```

### Regeln

| # | Regel | Begründung |
|---|---|---|
| 🚫 1 | **Niemals Nicht-ASCII-Zeichen als Literale in `run script`-Strings einbetten** | AppleScript-Compiler wirft Syntax-Error -2741 |
| 🚫 2 | **Niemals Dateipfade per String-Konkatenation in AppleScript-Quellcode einbauen** | Funktioniert zufällig für ASCII, bricht bei Umlauten/Sonderzeichen |
| ✅ 3 | **Pfade immer als Handler-Parameter übergeben** | AppleScript empfängt sie als Unicode-String, `quoted form of` funktioniert korrekt |
| ✅ 4 | **`EnsureMailReaderScptInstalled` auf neuen Handler testen** | Altes .scpt ohne `CopyFile` triggert automatischen Neuinstall |
| ✅ 5 | **Nach .scpt-Änderung: `~/Library/Application Scripts/com.microsoft.Excel/MailReader.scpt` löschen** | Excel nutzt gecachte Version – manuelles Löschen erzwingt Neuinstall |

### Betroffene Dateien
- `Excel_leads/Excel_files/MailReader.applescript` / `MailReader.scpt`
- `Excel_leads/Excel_files/bas/Main.bas` (v3.2)

---

## LL-004 · 2026-06-11 · `EnsureMailReaderScptInstalled` installiert .scpt nicht zuverlässig

### Was ist passiert?
Nach dem Löschen des gecachten `MailReader.scpt` und dem Import der neuen `Main.bas` lief der Import erneut durch — mit identischen Err-75-Fehlern für alle 5 Umlaut-Dateien. Der neue `CopyFile`-Handler war nicht verfügbar, weil die Auto-Installation lautlos scheiterte.

**Diagnostischer Hinweis:** Die Fehlergeschwindigkeit verriet das Problem:
- Vor dem Fix: 13 Sekunden für 5 Fehler (AppleScriptTask wartete auf Timeout)
- Nach kaputtem Install: **1 Sekunde** für 5 Fehler → `ReadEmlViaShellCopy` exitete sofort (kein `.scpt` vorhanden)

### Ursache(n)
Zwei mögliche Ursachen, die zusammenwirken können:

**1. `Static alreadyTried` – einmal True, immer True**
`EnsureMailReaderScptInstalled` verwendet `Static alreadyTried As Boolean`. In VBA werden Static-Variablen beim Neuimport eines Moduls *meist* zurückgesetzt — aber nicht garantiert, wenn das VBA-Projekt nicht vollständig neu kompiliert wird (z. B. wenn Excel die Compile-State zwischen Sessions cached). Wird `alreadyTried = True` aus einem früheren Aufruf in der Session beibehalten, bricht die Funktion sofort ab ohne Installationsversuch.

**2. `FileCopy` in Mac-Sandbox zu Application Scripts unzuverlässig**
VBA's `FileCopy` nach `~/Library/Application Scripts/com.microsoft.Excel/` funktioniert **manchmal** (legacy bestätigt), aber nicht immer. Schlägt FileCopy still fehl (Err-Nummer wird geloggt aber nicht weitergegeben), endet die Funktion ohne installiertes .scpt — und ohne Fehlermeldung ans Makro.

### Symptom-Checkliste
| Symptom | Bedeutung |
|---|---|
| Alle Fehler in < 2 Sekunden | `.scpt` nicht installiert — `ReadEmlViaShellCopy` exitiert sofort |
| Fehler spread über > 10 Sekunden | `.scpt` vorhanden, aber `AppleScriptTask` läuft in Timeout (altes .scpt ohne Handler) |
| `~/Library/Application Scripts/com.microsoft.Excel/` leer | Auto-Install scheiterte — manuelle Installation nötig |

### Fix (Main.bas v3.3)

**Sofort-Workaround** (ohne Excel-Neustart):
```
Im VBA-Direktbereich eintippen: InstallMailReaderScpt
```
Dieser Public Sub bypasses das Static-Flag vollständig.

**Oder manuell im Terminal:**
```bash
cp "/Users/steffen/Documents/GitHub/Excel Leads/Excel_files/MailReader.scpt" \
   ~/Library/Application\ Scripts/com.microsoft.Excel/
```

**Strukturelle Lösung v3.3:**
- `EnsureMailReaderScptInstalled` nur noch dünner Static-Guard → delegiert an `InstallMailReaderScpt`
- `InstallMailReaderScpt` ist `Public` → direkt aufrufbar im Direktbereich
- 4-stufige Installations-Strategie:

```
Versuch 1: AppleScriptTask(HANDLER_COPY, "_test_|_test_") → Err.Number = 0? → schon vorhanden, fertig
Versuch 2: FileCopy sourcePath → contPath
Versuch 3: MacScript "do shell script cp ..." (Pfade sind ASCII → run script sicher)
Versuch 4: MsgBox mit Terminal-Befehl für manuellen Copy
```

### Regeln

| # | Regel | Begründung |
|---|---|---|
| ⚠️ 1 | **Nach .scpt-Änderung: Excel neu starten** oder `InstallMailReaderScpt` im Direktbereich aufrufen | Static-Flag wird erst bei VBA-Reset sicher zurückgesetzt |
| ✅ 2 | **Fehlergeschwindigkeit als Diagnosewerkzeug nutzen** | < 2 s → kein .scpt; > 10 s → falscher Handler |
| ✅ 3 | **`InstallMailReaderScpt` als Public Sub** ermöglicht manuellen Neuinstall ohne Excel-Neustart | Static-Flag umgehen |
| ✅ 4 | **MacScript als Fallback** für FileCopy (Pfade zur Application Scripts sind ASCII → kein Umlaut-Problem) | FileCopy in Sandbox unzuverlässig |
| ✅ 5 | **MsgBox mit Terminal-Befehl** als letzte Eskalationsstufe | Nutzer bekommt konkreten Befehl, kein stilles Scheitern |

### Betroffene Dateien
- `Excel_leads/Excel_files/bas/Main.bas` (v3.3)

---

## LL-005 · 2026-06-12 · `InstallMailReaderScpt` schlägt fehl wenn .xlsm aus Outlook geöffnet wird

### Was ist passiert?
Karim (Kunde) erhielt die MsgBox „Installation fehlgeschlagen – MailReader.scpt fehlt im Workbook-Ordner" beim ersten Import-Start. Der gezeigte Pfad war der Outlook-Temp-Ordner:
```
/Users/maghrebikarim/Library/Containers/com.microsoft.Outlook/Data/tmp/OutlookTemp/MailReader.scpt
```

### Ursache
`InstallMailReaderScpt` (v3.3) suchte `MailReader.scpt` in `ThisWorkbook.Path`. Karim hatte die `.xlsm` direkt aus einer Outlook-E-Mail geöffnet (ohne zu speichern) → `ThisWorkbook.Path` = Outlook-Temp-Ordner → `.scpt` nie dort vorhanden → MsgBox statt Auto-Install.

### Fix (Main.bas v3.4)
`MailReader.scpt` wird nun als Base64-String direkt im VBA eingebettet (`GetMailReaderScptBase64()`). Kein externer Datei-Copy mehr nötig.

**Neue Installations-Strategie:**
```
Versuch 1: AppleScriptTask(CopyFile, "_test_|_test_") → bereits installiert? → fertig
Versuch 2: GetMailReaderScptBase64() → VBA schreibt in $TMPDIR (kein Sandbox-Block)
           MacScript "base64 -D -i tmpFile -o ~/Library/Application Scripts/..." → installiert
Versuch 3: MsgBox mit Terminal-Befehl (letzte Eskalation)
```

**Warum TMPDIR + Shell:**
- VBA hat Schreibrecht auf `$TMPDIR` (kein Sandbox-Block)
- MacScript `base64 -D` umgeht Application-Scripts-Sandbox
- Pfade sind ASCII → kein Umlaut-Problem in `do shell script`

### Regeln

| # | Regel | Begründung |
|---|---|---|
| 🚫 1 | **Niemals `ThisWorkbook.Path` für externe Ressourcen nutzen** | Pfad ist Outlook-Temp wenn Datei aus Mail geöffnet wird |
| ✅ 2 | **Binäre Ressourcen als Base64 einbetten** | Macht .xlsm selbsttragend, kein Deployment-Problem |
| ✅ 3 | **TMPDIR als Staging-Bereich** für Shell-Operationen | Einzige Sandbox-sichere VBA-Write-Location |

### Betroffene Dateien
- `Excel_leads/Excel_files/bas/Main.bas` (v3.4)

---

*Weitere Einträge werden fortlaufend ergänzt.*
