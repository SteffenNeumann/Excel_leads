# Konzept: E-Mail-Parsing & SeaTable-Import – Vitalis Seniorendienst

**Erstellt:** 2026-03-23
**Status:** Konzeptphase

---

## Ausgangslage

Vitalis Seniorendienst GmbH (Ismaning) erhält Kundenanfragen über die Plattform **Verbund Pflegehilfe** (`anfragen@pflegehilfe.de`). Diese Anfragen werden als weitergeleitete E-Mails an das Outlook-Postfach (`Info@vitalis-seniorendienst.de`) zugestellt.

Ziel: Neue Anfragen automatisch aus dem Postfach auslesen, Daten extrahieren und in SeaTable eintragen — ohne manuelle Übertragung.

---

## E-Mail-Format

- **Absender:** `anfragen@pflegehilfe.de`
- **Betreff-Muster:** `Neue Anfrage: [Kundenname]` (weitergeleitet als `WG: Neue Anfrage: [Kundenname]`)
- **Kodierung:** `quoted-printable`, Zeichensatz `Windows-1252`
- **Struktur:** Konsistentes `Label:\nWert\n`-Muster — kein Freitext, kein KI-Parsing nötig

### Extrahierbare Felder

| Kategorie | Felder |
|---|---|
| Interessent | Name, Telefon (Mobil/Festnetz), E-Mail, Erreichbarkeit, Anschrift |
| Senior | Name, Beziehung, Alter, Pflegegrad, Lebenssituation, Mobilität, Medizinisches |
| Anfrage | Anfragen-Nr., Bedarfsort, Aufgaben, Wöchentlicher Umfang, Umfang am Stück, Abrechnung über Bet.- & Entlastungsleistungen, Pflegedienst vorhanden, Bedarf |
| Datenschutz | Zustimmungsdatum |

---

## Technischer Ansatz

**Sprache:** Python 3
**Parsing:** Standard-`email`-Bibliothek (Quoted-Printable + Windows-1252 Decode), Regex für Feldextraktion
**Ziel-API:** SeaTable REST API (Base-Token, gleicher Ansatz wie bei Suzananda WF1b)
**E-Mail-Zugriff:** IMAP oder Microsoft Graph API (Exchange Online)

### Ablauf

1. Script verbindet sich per IMAP/Graph API mit dem Exchange-Postfach
2. Neue E-Mails mit Betreff "Neue Anfrage:" werden gefiltert
3. Body wird dekodiert (quoted-printable, Windows-1252)
4. Relevanter Block ab "Neue Anfrage: Haushaltshilfe" wird isoliert
5. Felder per Regex extrahiert
6. Zeile in SeaTable per HTTP POST angelegt
7. E-Mail als "verarbeitet" markieren (z. B. verschieben in Unterordner)

---

## Plattform

**GitHub Actions** (kostenlos, keine lokale Installation nötig)

- Kein Python-Install auf Kundenseite erforderlich
- Bedienung komplett im Browser (Mac-kompatibel)
- Cron-Scheduler: z. B. stündlich
- Zugangsdaten als verschlüsselte GitHub Secrets hinterlegt
- Logs direkt in der GitHub-Oberfläche einsehbar
- Free-Plan reicht aus (weit unter 2.000 Min/Monat)

**Alternative:** PythonAnywhere (einfacheres Interface, ebenfalls kostenlos)

---

## Beispieldaten

Zwei Test-EML-Dateien vorhanden:

- `Mails/WG_ Neue Anfrage_ Monika Schmidt.eml`
- `Mails/WG_ Neue Anfrage_ Edeltraud Hörath.eml`

Beide Dateien bestätigen identisches Format — Parsing-Logik kann direkt darauf entwickelt und getestet werden.

---

## Offene Punkte

- [ ] SeaTable Base und Tabellenstruktur definieren (Spalten je Feld)
- [ ] IMAP-Zugangsdaten / Graph API OAuth klären (Exchange Online?)
- [ ] Duplikat-Erkennung: Anfragen-Nr. als eindeutiger Schlüssel verwenden
- [ ] Fehlerbehandlung: Was passiert bei unvollstaendigen E-Mails?
- [ ] GitHub-Repo anlegen und Actions konfigurieren
