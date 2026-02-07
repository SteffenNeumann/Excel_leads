# Workflow: Apple Mail Leads → Excel (macOS, Excel VBA)

## Ziel
E-Mails in Apple Mail mit den Schlagwörtern **„Lead“** oder **„Neue Anfrage“** finden, Inhalte analysieren, strukturierte Daten extrahieren und in eine Excel-Tabelle in die nächste freie Zeile einfügen.

## Voraussetzungen
- macOS, Apple Mail, Microsoft Excel (Mac)
- Excel-Datei mit Tabelle (ListObject) und fixen Spaltenüberschriften
- Arbeitsblatt: **Pipeline**
- Tabellenname: **Kundenliste** (intelligente Tabelle)
- Makros aktiviert

## Datenfelder (Zielspalten)
- Kontakt: Anrede, Vorname, Nachname, Name (falls als Vollname), Mobil, E-Mail-Adresse, Erreichbarkeit
- Senior: Name, Beziehung, Alter, Pflegegrad Status, Pflegegrad, Lebenssituation, Mobilität, Medizinisches
- Anfrage: PLZ, Nutzer, Alltagshilfe Aufgaben, Alltagshilfe Häufigkeit, ID (falls vorhanden)

## Workflow-Schritte
1) **Apple Mail durchsuchen**
	- Filter: Betreff oder Body enthält **„Lead“** oder **„Neue Anfrage“**.
	- Nur ungelesene oder letzte 24/48h (optional) zur Duplikatvermeidung.
	- Ergebnis: Liste der passenden Nachrichten inkl. Absender, Datum, Betreff, Body.

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

