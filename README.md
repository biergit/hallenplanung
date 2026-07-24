# Hallen/Spielplanung Tischtennis

Google Apps Script für die Hallen/Spielplanung eines Tischtennisvereins.

## Enthaltene Blätter

| Blatt | Zweck |
|-------|-------|
| *Team-Blätter* | Ein Blatt pro Team – Mannschaftsführer tragen hier ihre Spiele ein |
| Sperrungen/Anderweitige Belegungen | Gesperrte Tage, Bereiche und Zeiträume |
| Hallen/Spielplan | Öffentliche Kalenderansicht (für Web-Veröffentlichung) |
| Setup | Team-Konfiguration (Rang, Gruppe, Name) und Einstellungen |

## Funktionen

- Pro Team ein eigenes Eingabe-Blatt – Team ergibt sich aus dem Blatt-Namen
- Ausgeblendete Endzeit-Spalte (wird automatisch aus Startzeit + Spieldauer berechnet)
- Status-Spalte früh im sichtbaren Bereich für kleinere Bildschirme
- Hallenbereiche: Große Halle links, Kleine Halle, Große Halle rechts, Große Halle Mitte
- Spieltage: Di, Mi, Fr, Sa, So
- Mittwoch: nur "Kleine Halle" buchbar
- Di + Fr: maximal 2 Bereiche (Platz für Training)
- Konflikt-Prüfung auch teamübergreifend (alle Team-Blätter + Sperrungen)
- Heim-/Auswärtsspiele farblich unterschieden (grün/blau)
- Sperrungen mit optionaler Start-/Endzeit (zeitbasierte oder ganztägige Sperrung)
- Konfigurierbare Standard-Spieldauer (in Setup-Blatt)
- Kalender-Datepicker für Datumseingabe

## Hallen/Spielplan – Filterung

Im veröffentlichten Web-View sind keine Script-Trigger aktiv. Für interaktive Filterung:
1. `Hallen/Spielplan`-Tab im Google Sheet öffnen
2. **Daten → Slicer** → je einen Slicer anlegen für:
   - **Team** (Spalte E)
   - **Heim/Auswärts** (Spalte F: Heim / Auswärts)
   - **Typ** (Spalte H: `gesperrt` oder leer für Spieleinträge)
3. Slicer funktionieren auch in der veröffentlichten HTML-Ansicht

## Installation

1. Google Sheet erstellen (Tabellenkalkulation)
2. **Erweiterungen → Apps Script**
3. Code aus `hallenspielplan.gs` vollständig einfügen *(bzw. `build / hallenspielplan.gs` mit Seed-Daten)*
4. Funktion `setupSheet` auswählen und ausführen (▶️)
5. Berechtigungen erteilen
6. Zurück zum Sheet – Team-Blätter werden aus Setup angelegt

## Web-Veröffentlichung

- **Datei → Für das Web veröffentlichen**
- Blatt "Hallen/Spielplan" auswählen
- Link teilen

## Freigabe an Mannschaftsführer

- **Datei → Freigeben**
- E-Mail-Adressen der Mannschaftsführer hinzufügen
- Berechtigung: "Bearbeiter"

## Build (Seed-Daten einbetten)

```
python3 tools/build.py
```

Liest `data/*.tsv` und schreibt `build/hallenspielplan.gs` mit eingebetteten Seed-Daten.
