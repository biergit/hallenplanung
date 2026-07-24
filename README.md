# Hallen/Spielplanung Tischtennis

Google Apps Script für die Hallen/Spielplanung eines Tischtennisvereins.

## Enthaltene Blätter

| Blatt | Zweck |
|-------|-------|
| Setup | Team-Konfiguration (Rang, Gruppe, Name/Kurzname per Formel) und Einstellungen |
| Sperrungen/Anderweitige Belegungen | Gesperrte Tage, Bereiche und Zeiträume |
| Eingabe | Dateneingabe durch Mannschaftsführer (mit Validierung) |
| Hallen/Spielplan | Öffentliche Kalenderansicht (für Web-Veröffentlichung) |

## Funktionen

- Hallenbereiche: Große Halle links, Kleine Halle, Große Halle rechts, Große Halle Mitte
- Spieltage: Di, Mi, Fr, Sa, So
- Mittwoch: nur "Kleine Halle" buchbar
- Di + Fr: maximal 2 Bereiche (Platz für Training)
- Heim-/Auswärtsspiele farblich unterschieden (grün/blau)
- Ersatzspieler-Hinweis: benachbarte Teams (gleiche Gruppe, Rangunterschied 1) am selben Tag werden markiert
- Sperrungen mit optionaler Start-/Endzeit (zeitbasierte oder ganztägige Sperrung)
- Konfigurierbare Standard-Spieldauer (in Setup-Blatt), verwendet für Zeitüberlappungsprüfung
- Web-Veröffentlichung mit zwei Umschaltern:
  - "Nur Hallenbelegung" (blendet Auswärtsspiele aus)
  - "Sperrungen anzeigen" (blendet Sperrungen ein)
- Kalender-Datepicker für Datumseingabe

## Installation

1. Google Sheet erstellen (Tabellenkalkulation)
2. **Erweiterungen → Apps Script**
3. Code aus `hallenbelegung.gs` vollständig einfügen
4. Funktion `setupSheet` auswählen und ausführen (▶️)
5. Zweimal Berechtigungen erteilen
6. Zurück zum Sheet – alles ist eingerichtet

## Web-Veröffentlichung

- **Datei → Für das Web veröffentlichen**
- Blatt "Hallen/Spielplan" auswählen
- Link teilen

## Freigabe an Mannschaftsführer

- **Datei → Freigeben**
- E-Mail-Adressen der Mannschaftsführer hinzufügen
- Berechtigung: "Bearbeiter"
