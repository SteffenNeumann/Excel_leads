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

*Weitere Einträge werden fortlaufend ergänzt.*
