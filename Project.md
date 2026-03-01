# Workflow: Apple Mail / Outlook Leads → Excel (macOS, Excel VBA)

## Ziel
E-Mails in Apple Mail oder Microsoft Outlook mit den Schlägwörtern **„Lead“** oder **„Neue Anfrage“** finden, Inhalte analysieren, strukturierte Daten extrahieren und in eine Excel-Tabelle in die nächste freie Zeile einfügen.

## Voraussetzungen
- macOS, Apple Mail **oder** Microsoft Outlook, Microsoft Excel (Mac)
- Excel-Datei mit Tabelle (ListObject) und fixen Spaltenüberschriften
- Arbeitsblatt: **Pipeline**
- Tabellenname: **Kundenliste** (intelligente Tabelle)
- Makros aktiviert

## Einstellungen (Sheet "Berechnung")
| Benannter Bereich | Default | Beschreibung |
|---|---|---|
| `LEAD_MAILBOX` | `iCloud` | Account-Name(n), per `;` trennbar. Enthält `@`/`outlook`/`exchange` → Outlook, sonst → Apple Mail |
| `LEAD_FOLDER` | `Leads` | Ordnername(n), per `;` trennbar (gleiche Reihenfolge wie LEAD_MAILBOX) |
| `mailpath` | *(leer)* | Optionaler lokaler Pfad zu .eml-Dateien |

**Beispiel für beide Apps gleichzeitig:**
- `LEAD_MAILBOX` = `iCloud;steffen.neumann@dlh.de`
- `LEAD_FOLDER` = `Leads;Posteingang`
- → Sucht in Apple Mail (iCloud/Leads) **und** Outlook (steffen.neumann@dlh.de/Posteingang)

## Datenfelder (Zielspalten)
- Kontakt: Anrede, Vorname, Nachname, Name (falls als Vollname), Mobil, E-Mail-Adresse, Erreichbarkeit
- Senior: Name, Beziehung, Alter, Pflegegrad Status, Pflegegrad, Lebenssituation, Mobilität, Medizinisches
- Anfrage: PLZ, Nutzer, Alltagshilfe Aufgaben, Alltagshilfe Häufigkeit, ID (falls vorhanden)

## Workflow-Schritte
1) **Mail-App durchsuchen (Apple Mail oder Outlook)**
	- Gesteuert über `LEAD_MAIL_APP` Einstellung im Sheet "Berechnung".
	- Filter: Betreff enthält **„Lead“** oder **„Neue Anfrage“**.
	- Nur ungelesene oder letzte 24/48h (optional) zur Duplikatvermeidung.
	- Ergebnis: Liste der passenden Nachrichten inkl. Absender, Datum, Betreff, Body.
	- **Apple Mail**: `tell application "Mail"` – durchsucht `mailbox` im Account.
	- **Outlook**: `tell application "Microsoft Outlook"` – durchsucht `mail folder` im Account (Exchange/IMAP/POP).

2) **Nachrichteninhalt extrahieren**
	- Body als reiner Text lesen.
	- Erkennen, ob Format A (Kontakt/Senior-Block) oder Format B (ID/Einzelfelder).

3) **Parsing (VBA-Regeln)**
	- Zeilenweise auswerten und bekannte Labels mappen.
	- Beispiele:
	  - „Name:“ → `Kontakt_Name` (oder `Senior_Name` je nach Abschnitt)
	  - „Mobil:“ → `Kontakt_Mobil`
	  - „E-Mail-Adresse:“ oder „E-Mail:“ → `Kontakt_Email`
	  - „Pflegegrad:“ → `Senior_Pflegegrad`
	  - „Pflegegrad Status:“ → `Senior_Pflegegrad_Status`
	  - „PLZ/Postleitzahl:“ → `PLZ`
	  - „Alltagshilfe Aufgaben:“ → `Alltagshilfe_Aufgaben`
	  - „Alltagshilfe Häufigkeit:“ → `Alltagshilfe_Haeufigkeit`
	  - „ID:“ → `Anfrage_ID`
	- Abschnittswechsel erkennen (z. B. „Informationen zum Senior“).

4) **Excel-Tabelle befüllen**
	- Tabelle **Kundenliste** auf Blatt **Pipeline** als ListObject nutzen.
	- Nächste freie Zeile: `ListObject.ListRows.Add`.
	- Spalten über Headernamen finden und setzen.

5) **Duplikat-Handling**
	- Eindeutigkeit über Kombination aus `Anfrage_ID` oder (E-Mail + Datum).
	- Falls bereits vorhanden: überspringen oder aktualisieren.

6) **Protokollierung**
	- Optional: Log-Sheet mit Zeitstempel und Message-ID.

## VBA-Implementierung (Outline)
```vba
' 1) Apple Mail abfragen (AppleScript)
Dim script As String, result As String
script = "" & _
"tell application \"Mail\"" & vbLf & _
"set theMessages to (every message of inbox whose subject contains \"Lead\" or subject contains \"Neue Anfrage\" or content contains \"Lead\" or content contains \"Neue Anfrage\")" & vbLf & _
"set outText to \"\"" & vbLf & _
"repeat with m in theMessages" & vbLf & _
"set outText to outText & (content of m) & \"\n---MSG---\n\"" & vbLf & _
"end repeat" & vbLf & _
"return outText" & vbLf & _
"end tell"
result = MacScript(script)

' 2) Parsing
' - Split by "---MSG---"
' - Parse lines, map labels, fill Dictionary

' 3) Excel-Insert
' - Find table, add row, set cells by header name
```

## Tabellenstruktur (Kundenliste)
Kopf-Überschriften in der Tabelle:
- Monat Lead erhalten
- Status
- Lead-Quelle
- Name
- Adresse
- PLZ
- Ort
- Telefonnummer
- PG
- Letzter Kontakt
- Nächster Kontakt
- Notizen
- Abschluss
- Abgesprungen nach
- Grund zum Absprung
- Learning
- Reklamierung Verbund:
- Spend
- Leadtyp

## Spalten-Mapping (Kundenliste)
| Quelle | Zielspalte |
|---|---|
| Kontakt: Name (voll oder aus Vor-/Nachname) | Name |
| Kontakt: Mobil / Telefonnummer | Telefonnummer |
| Postleitzahl / PLZ | PLZ |
| Senior: Pflegegrad | PG |
| Lead-Quelle | Lead-Quelle |
| Monat Lead erhalten (aus Mail-Datum) | Monat Lead erhalten |
| Leadtyp | Leadtyp |
| Notizen (Restinfos wie Erreichbarkeit, Beziehung, Lebenssituation, Mobilität, Medizinisches, Aufgaben, Häufigkeit, ID, E-Mail) | Notizen |

Nicht aus den Maildaten befüllbar (bleiben leer, bis manuell gepflegt):
- Status, Adresse, Ort, Letzter Kontakt, Nächster Kontakt, Abschluss, Abgesprungen nach, Grund zum Absprung, Learning, Reklamierung Verbund:, Spend

## Ergebnis
Bei jedem Lauf werden neue Leads aus Apple Mail erkannt, die relevanten Felder extrahiert und in die nächste freie Zeile der Excel-Tabelle geschrieben.

---

## Dashboard (Dashboard.bas)

### Übersicht
Automatisch generiertes Analytics-Dashboard auf dem Blatt **Dashboard**. Datenquelle ist die Tabelle **Kundenliste** auf dem Blatt **Pipeline**. Aufruf über `BuildDashboard`.

### Design
- Card-basiertes Layout mit abgerundeten Rechtecken (`msoShapeRoundedRectangle`)
- Weiße Karten mit Schatten (Blur 8, Transparency 0.6)
- Hintergrund: Blue White `RGB(245, 248, 252)`
- Schrift: Avenir Next
- Farbpalette: Dezente, monochromatische Blautöne

### Farbpalette
| Farbe | RGB | Verwendung |
|---|---|---|
| Dark Navy | `RGB(25, 55, 95)` | Titel, Lead-Werte |
| Ocean Blue | `RGB(50, 110, 165)` | Primärakzent, Balken, Chart 1 Serie 1 |
| Steel Blue | `RGB(95, 145, 190)` | Sekundärakzent, Laufend, Chart 1 Serie 2 |
| Fog Blue | `RGB(140, 160, 185)` | Absprünge, Chart 2 Balken |
| Blue-Grey | `RGB(100, 120, 150)` | Labels, Spaltenköpfe |
| Slate Blue | `RGB(65, 85, 110)` | Tabellen-Body-Text |
| Ice Blue | `RGB(215, 225, 235)` | Trennlinien |
| Blue White | `RGB(245, 248, 252)` | Hintergrund |

### Komponenten
1. **KPI-Karten** (4 Stück, obere Reihe):
   - Gesamt Leads, Abschlussrate, Absprünge, Laufend
   - Jeweils mit farbigem Seitenstreifen

2. **Charts** (2 Stück, zweite Reihe):
   - **Leads & Abschluss Trend** – Liniendiagramm mit Markern (Leads + Abgeschlossen)
   - **Absprung Trend** – Säulendiagramm (Abgesprungen pro Monat)
   - Chart-Daten in Spalten T–Y (sichtbar), `.SetSourceData` für Mac-Kompatibilität

3. **Absprunggruende** (Tabelle, dritte Reihe links):
   - Top-Gründe mit Anzahl und Anteil
   - Dynamische Card-Höhe (`abgCardH`)

4. **Insights & Empfehlungen** (dritte Reihe rechts):
   - Top Absprunggrund, Peak Leads Monat, Schwächste Abschlussrate, Häufigster Absprungzeitpunkt

5. **Abgesprungen nach Zeitpunkt** (vierte Reihe links):
   - Zeitpunkt-Verteilung mit Anzahl und Anteil

6. **Monatsübersicht** (vierte Reihe rechts):
   - Monat, Leads, Abgeschlossen, Abgesprungen, Rate

### Genutzte Tabellenspalten
- `Monat Lead erhalten` – Gruppierung nach Monat (Year*100+Month)
- `Abschluss` – ja/nein/laufend
- `Grund zum Absprung` – Absprunggrund-Statistik
- `Abgesprungen nach` – Zeitpunkt-Statistik
- `Status` – Statusauswertung

### Hilfsfunktionen
- `FormatCard(shp)` – Weißer Fill, Schatten, abgerundete Ecken
- `AddLabel(ws, x, y, w, h, txt, fontSize, isBold, clr)` – Transparente Textbox

### Layout-Konstanten
| Konstante | Wert | Bedeutung |
|---|---|---|
| LM | 20 | Linker Rand |
| CW | 195 | KPI-Kartenbreite |
| CH | 85 | KPI-Kartenhöhe |
| CG | 15 | Kartenabstand |
| CHW | 410 | Chart-/Tabellenbreite |
| CHH | 230 | Chart-Höhe |
| SG | 20 | Sektionsabstand |
| TH | 230 | Tabellen-Kartenhöhe |
| dataCol | 20 | Erste Datenspalte (T) |

### Mac-Kompatibilität
- `.SetSourceData` statt `.SeriesCollection.NewSeries` (funktioniert zuverlässig auf Mac Excel)
- Chart 2 nutzt eigene Spalten (dataCol2 = 24), da nicht-zusammenhängende Ranges auf Mac problematisch sind
- `On Error Resume Next` um Chart-Formatierung für Mac-spezifische Einschränkungen
- Keine `Paragraphs()`-API, keine `ChartTitle`-Objekte

### Änderungshistorie
| Datum | Commit | Änderung |
|---|---|---|
| 2026-02-07 | bf0fadb | Fix: Spalten erst nach Chart-Erstellung verstecken |
| 2026-02-07 | b114c15 | Fix: SetSourceData statt NewSeries für Mac-Charts |
| 2026-02-07 | 161d75f | Datenspalten sichtbar lassen, Palette Teal/Rose/Steel |
| 2026-02-07 | 143b62e | Design: Dezente Blauton-Palette (Navy/Ocean/Steel/Fog) |
| 2026-02-07 | da7a052 | Update Pipeline-Leads.xlsm |

---

## Troubleshooting Guide

Übersicht aller bekannten Probleme, Lösungsversuche und deren Status.

### Legende
- ✅ **Gelöst** – Fix bestätigt und produktiv
- 🔧 **Implementiert** – Fix committet, wartet auf Bestätigung beim Kunden
- ❌ **Fehlgeschlagen** – Ansatz verworfen
- ⏳ **Offen** – Noch nicht gelöst

---

### 1. Umlaut-Dateien (.eml) können nicht gelesen werden

**Symptom:** `Nachricht X fehlgeschlagen [Datei: WG_ Neue Anfrage_ Sabine Bäuml.eml] (Ergebnis: 0)`
VBA `Dir$`, `Open For Binary`, `MacScript` können NFD-kodierte Umlaute (ö = o + U+0308) in Dateinamen nicht verarbeiten.

| # | Ansatz | Commit | Status |
|---|---|---|---|
| 1 | VBA `Dir$` mit Umlaut-Sonderbehandlung | — | ❌ Dir$ gibt NFD-Namen nicht zurück |
| 2 | Python als primärer EML-Reader | `f315924` | ❌ Python hilft nicht wenn Dateiname selbst das Problem ist |
| 3 | **Ansatz A+: perl-basierte regelkonforme Umbenennung (ä→ae, ö→oe, ü→ue, ß→ss)** | `0f98747` | 🔧 Perl-Skript base64-kodiert, `SanitizeEmlFileNames` VOR `Dir$` ausgeführt |

**Aktueller Stand:** SanitizeEmlFileNames läuft via `RunShellCommand` (MacScript → AppleScriptTask Fallback). Perl ist auf jedem Mac vorinstalliert. Umbenennung erfolgt VOR dem VBA-Import. Warte auf Kundenbestätigung.

---

### 2. MacScript funktioniert nicht auf 64-bit Excel

**Symptom:** `MacScript` schlägt fehl mit Laufzeitfehler auf neueren 64-bit Excel-for-Mac Installationen. Alle Shell-Aufrufe (Perl-Rename, Python-Fallback, EML-Lesen) brechen ab.

| # | Ansatz | Commit | Status |
|---|---|---|---|
| 1 | **RunShellCommand Helper mit MacScript → AppleScriptTask Fallback** | `8ea347b` | ✅ Bestätigt: 3 Renames via MacScript, 3 via AppleScriptTask |

**Lösung:** `RunShellCommand()` versucht zuerst `MacScript("do shell script ...")`, bei Fehler Fallback auf `AppleScriptTask(MailReader.scpt, FetchMessages, "do shell script ...")`. Alle Shell-Aufrufe (SanitizeEmlFileNames, ReadTextFileViaShell, PythonReadEmlFile) nutzen diesen Helper.

---

### 3. "AppleScript Quelle fehlt" beim Kunden

**Symptom:** `MailReader.scpt` und `.applescript` liegen nur im Repo-Root, nicht im `Excel_files/`-Ordner, wo das Workbook liegt. `AppleScriptTask` findet die .scpt nicht.

| # | Ansatz | Commit | Status |
|---|---|---|---|
| 1 | Dateien nach `Excel_files/` kopiert + osacompile via MacScript statt `Shell` | `12e719a` | ❌ Hilft nicht wenn Kunde nur die .xlsm hat |
| 2 | **MailReader.scpt als Base64 direkt im VBA eingebettet (1138 Bytes)** | `a375de7` | ❌ MWriteBase64ToFile nutzte MSXML2.DOMDocument (Windows-only) |
| 3 | **Pure-VBA DecodeBase64() Decoder** (ersetzt MSXML2) | `a375de7` | ❌ VBA `Open For Binary` kann in Sandbox nicht in Application Scripts schreiben |
| 4 | **Shell-basierte Installation: Base64 → TMPDIR → `base64 -D` via Shell → Ziel** | `bfebe45` | 🔧 TMPDIR-Write per VBA (erlaubt), Shell dekodiert zum Ziel (umgeht Sandbox) |

**Aktueller Stand:** Strategie-Reihenfolge in `InstallAppleScript`:
1. Base64 → TMPDIR (VBA) → `base64 -D` via MacScript Shell → Ziel
2. osacompile via MacScript (`.applescript` → `.scpt`)
3. FileCopy (nur nicht-sandboxed Excel)
4. MsgBox mit manuellem Terminal-Befehl

Existenzprüfung via `FileExistsViaShell()` (`test -f` via MacScript) statt `Dir$()`.
Warte auf Kundenbestätigung.

---

### 4. VBA Dir$/Open/FileCopy versagt in Sandbox für Application Scripts

**Symptom:** `Dir$(targetPath)` gibt immer leeren String zurück für `~/Library/Application Scripts/com.microsoft.Excel/`. VBA `Open For Binary Access Write` und `FileCopy` schlagen ebenfalls fehl. Die .scpt-Installation meldet jedes Mal "nicht vorhanden" und scheitert an allen Strategien.

| # | Ansatz | Commit | Status |
|---|---|---|---|
| 1 | **FileExistsViaShell()** – `test -f` via MacScript statt Dir$ | `bfebe45` | 🔧 |
| 2 | **Shell-Write statt VBA I/O** – `base64 -D > target` via MacScript | `bfebe45` | 🔧 |

**Aktueller Stand:** Alle Dateizugriffe auf den Application-Scripts-Ordner laufen jetzt über Shell-Kommandos. VBA-I/O wird nur noch für TMPDIR verwendet (dort hat VBA Zugriff). Warte auf Kundenbestätigung.

---

### Commit-Historie (chronologisch)
| Datum | SHA | Beschreibung |
|---|---|---|
| 2026-02-26 | `b6553ce` | Structured debug logging system |
| 2026-02-26 | `f315924` | Python als primärer EML-Reader |
| 2026-02-26 | `0f98747` | SanitizeEmlFileNames (perl base64, Ansatz A+) |
| 2026-02-27 | `12e719a` | AppleScript-Dateien nach Excel_files/ |
| 2026-02-27 | `8ea347b` | RunShellCommand MacScript→AppleScriptTask Fallback |
| 2026-03-01 | `a375de7` | Embedded MailReader.scpt als Base64, DecodeBase64 |
| 2026-03-01 | `bfebe45` | Shell-basierte .scpt Installation (Sandbox-Fix) |

