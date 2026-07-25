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
- Di + Mi: nur "Kleine Halle" buchbar
- Fr: max. 2 Bereiche, bei 2 Heimspielen muss "Kleine Halle" dabei sein
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
3. Code aus **`build/hallenspielplan.gs`** vollständig einfügen
4. Funktion **`resetAll`** auswählen und ausführen (▶️)
5. Berechtigungen erteilen
6. Zurück zum Sheet – alle Blätter sind eingerichtet

`resetAll()` löscht alle bestehenden Blätter (außer Hallen/Spielplan) und
legt sie inkl. Seed-Daten neu an. Für nicht-destruktive Struktur-Updates
(ohne Datenverlust) gibt es `setupSheet()`.

## Web-Veröffentlichung

- **Datei → Für das Web veröffentlichen**
- Blatt "Hallen/Spielplan" auswählen
- Link teilen
- **Achtung:** Der Link enthält die Tab-ID (GID) des Plan-Blatts. Solange
  dieser Tab nicht gelöscht wird, bleibt die URL stabil — auch über
  `resetAll()` und `setupSheet()` hinweg.

## Freigabe an Mannschaftsführer

- **Datei → Freigeben**
- E-Mail-Adressen der Mannschaftsführer hinzufügen
- Berechtigung: "Bearbeiter"

## Build (Seed-Daten einbetten)

```
python3 tools/build.py          # Produktivdaten
python3 tools/build.py --testdata # Testdaten (alle Fehlerfälle)
python3 tools/build.py --clean    # Leere SEED-Arrays
```

Liest `data/*.tsv` (bzw. `data/test/*.tsv`) und schreibt
`build/hallenspielplan.gs` mit eingebetteten Seed-Daten.

## Testdaten

```
python3 tools/build.py --testdata
```

Erzeugt `build/hallenspielplan.gs` mit 3 Teams und 15 Einträgen,
die alle Validierungsfehler auslösen — zum Durchtesten des Regelwerks:

| Fehlerfall | Test-Eintrag |
|---|---|
| H/A fehlt, Bereich fehlt, Montag | `Gegner …` |
| Endzeit < Startzeit | 16:00–14:00 |
| Di/Mi nur Kleine Halle | Falscher Bereich |
| Doppelbuchung + Zeitüberlappung | Gleicher Bereich, überlappend |
| Datum fehlt, Cross-Team-Konflikt |    |
| Sperrung-Überlappung, Ganztagssperrung |    |
| Endzeit ohne Startzeit, ungültiges Datum |    |

Zurück zu Produktivdaten: `python3 tools/build.py` (ohne `--testdata`),
dann Code neu kopieren und `resetAll()` ausführen.
