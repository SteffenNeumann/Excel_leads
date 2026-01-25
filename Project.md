# Workflow: Apple Mail Leads → Excel (macOS, Excel VBA)

## Ziel
E-Mails in Apple Mail mit den Schlagwörtern **„Lead“** oder **„Neue Anfrage“** finden, Inhalte analysieren, strukturierte Daten extrahieren und in eine Excel-Tabelle in die nächste freie Zeile einfügen.

## Voraussetzungen
- macOS, Apple Mail, Microsoft Excel (Mac)
- Excel-Datei mit Tabelle (ListObject) und fixen Spaltenüberschriften
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
	- Tabelle als ListObject nutzen (z. B. `tblLeads`).
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

## Spalten-Mapping (Beispiel)
| Quelle | Zielspalte |
|---|---|
| Anrede | Kontakt_Anrede |
| Vorname | Kontakt_Vorname |
| Nachname | Kontakt_Nachname |
| Name (Kontakt) | Kontakt_Name |
| Mobil | Kontakt_Mobil |
| E-Mail / E-Mail-Adresse | Kontakt_Email |
| Erreichbarkeit | Kontakt_Erreichbarkeit |
| Senior Name | Senior_Name |
| Beziehung | Senior_Beziehung |
| Alter | Senior_Alter |
| Pflegegrad Status | Senior_Pflegegrad_Status |
| Pflegegrad | Senior_Pflegegrad |
| Lebenssituation | Senior_Lebenssituation |
| Mobilität | Senior_Mobilitaet |
| Medizinisches | Senior_Medizinisches |
| Postleitzahl / PLZ | PLZ |
| Nutzer | Nutzer |
| Alltagshilfe Aufgaben | Alltagshilfe_Aufgaben |
| Alltagshilfe Häufigkeit | Alltagshilfe_Haeufigkeit |
| ID | Anfrage_ID |

## Ergebnis
Bei jedem Lauf werden neue Leads aus Apple Mail erkannt, die relevanten Felder extrahiert und in die nächste freie Zeile der Excel-Tabelle geschrieben.

