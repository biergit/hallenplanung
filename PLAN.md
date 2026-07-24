# Verbesserungsplan

## Prioritaet 1: Stabilitaet und Datenintegritaet

### Hallen/Spielplan skriptbasiert erzeugen

Die aktuelle Google-Sheets-Formel kombiniert mehrere `QUERY`-Ergebnisse per Array-Literal. Das ist wegen leerer Ergebnisse, Fehlerwerten und lokalisierungsabhaengiger Formeltrennzeichen fragil.

Ziel:

- `Eingabe` und `Sperrungen/Anderweitige Belegungen` mit Apps Script lesen.
- Je nach Checkboxen nur Heimspiele bzw. auch Sperrungen einbeziehen.
- Daten im Skript vereinheitlichen, nach Datum und Startzeit sortieren und gesammelt nach `Hallen/Spielplan!A4:I` schreiben.
- Vorhandene Planzeilen vor dem Schreiben leeren.

Nutzen:

- Keine `ARRAY_LITERAL`- oder `QUERY`-Fehler mehr.
- Unabhaengig von Spreadsheet-Locale und Formeltrennzeichen.
- Die Planlogik ist testbar und leichter zu erweitern.

### Trigger gezielt verwalten

`createTrigger()` darf nicht alle Trigger des Apps-Script-Projekts loeschen.

Ziel:

- Nur vorhandene `handleEdit`-Trigger fuer dieses Spreadsheet entfernen.
- Den eigenen Trigger idempotent neu anlegen.

### Schutz und Rollen trennen

Aktuell sind geschuetzte Bereiche nur Warnungen. Bearbeiter koennen dadurch Status-, Hilfs- und Planformeln ueberschreiben.

Ziel:

- `Setup`, Hilfsspalten, Statusspalte und `Hallen/Spielplan` tatsaechlich schuetzen.
- Fuer Mannschaftsfuehrer nur die Eingabespalten in `Eingabe` freigeben.
- Admins als erlaubte Bearbeiter der geschuetzten Bereiche setzen.

### Validierung vervollstaendigen

Fehlende oder ungueltige Eingaben duerfen nicht stillschweigend uebersprungen werden.

Ziel:

- Pflichtfelder fuer Team, Datum und Heim/Auswaerts pruefen.
- Bei Heimspielen Bereich und gueltige Start-/Endzeit verlangen.
- `Endzeit > Startzeit` erzwingen.
- Eine nur teilweise eingegebene Sperrzeit als Fehler markieren, statt sie als ganztagige Sperrung zu deuten.

## Prioritaet 2: Performance und Bedienung

### Validierung auf echte Daten begrenzen

Hilfsformeln bis Zeile 1001 bewirken, dass die Validierung regelmaessig 1.000 Zeilen verarbeitet.

Ziel:

- Letzte inhaltlich belegte Eingabezeile bestimmen, nicht `getLastRow()` der Hilfsspalten verwenden.
- Konflikte einmal nach Datum und Bereich indexieren, statt Zeilen mehrfach gegeneinander zu pruefen.
- Statuswerte in einem einzelnen `setValues()`-Aufruf schreiben.

### Endzeit bei allen relevanten Aenderungen aktualisieren

Die Endzeit wird derzeit nur nach einer Aenderung der Startzeit berechnet.

Ziel:

- Bei Aenderungen an Team, Startzeit oder teambezogener Spieldauer die betroffenen Zeilen aktualisieren.
- Mehrzellen-Paste vollstaendig behandeln.
- Manuell gesetzte Endzeiten entweder eindeutig verbieten oder als bewusstes Override kennzeichnen.

### Dynamische Bereiche statt fester Zeilengrenzen

Validierungen, Formeln und Formatierungen sind meist auf 1.000 Zeilen begrenzt.

Ziel:

- Bereiche beim Einfuegen oder Bearbeiten neuer Datenzeilen automatisch erweitern.
- Gemeinsame Standardgroessen und Spaltennummern zentral konfigurieren.

### Bedienung verbessern

- Hinweise sichtbar ausserhalb technischer Hilfsspalten platzieren.
- Statusmeldungen kurz, konkret und pro Eingabe nachvollziehbar halten.
- Einen klaren Hinweis im Plan anzeigen, wenn keine passenden Eintraege vorhanden sind.
- README auf den Namen `Hallen/Spielplan` aktualisieren.

## Prioritaet 3: Wartbare Weiterentwicklung

### Zentrales Schema einfuehren

Spaltennummern und Feldnamen werden derzeit an vielen Stellen direkt verwendet.

Ziel:

- Ein Schema in `CONFIG` fuer alle Blattspalten pflegen.
- Funktionen ueber Feldnamen bzw. zentrale Spaltenkonstanten arbeiten lassen.
- Neue Spalten an einer Stelle definieren und in Plan, Validierung, Import und Export gezielt abbilden.

### Versionierte, idempotente Migrationen

`setupSheet()` soll vorhandene Daten bewahren und nur fehlende Strukturen ergaenzen.

Ziel:

- Eine `SCHEMA_VERSION` im Setup-Blatt speichern.
- Jede Erweiterung als Migration implementieren: fehlende Spalten, Formeln, Validierungen, Formate oder Schutzregeln ergaenzen.
- Migrationen mehrfach sicher ausfuehrbar machen.
- `resetAll()` ausschliesslich als klar destruktive Test-/Neuaufbau-Funktion behalten.

### Seed-Daten klar abgrenzen

- TSV-Dateien nur fuer Demo, Tests und einen bewussten Neuaufbau verwenden.
- Produktivdaten bei `setupSheet()` niemals ueberschreiben.
- Den TSV-Builder robust gegen Apostrophe, Tabs und Zeilenumbrueche machen.
- `downloadTSV()` entweder als echten Export bereitstellen oder eindeutig als Log-Diagnose benennen.

### Testbare Kernlogik extrahieren

Ziel:

- Datum-/Zeit-Normalisierung, Konfliktpruefung und Planzeilenerzeugung von Apps-Script-I/O trennen.
- Diese Funktionen mit Node-Tests und kleinen TSV-Fixtures absichern.
- Regressionstests fuer Heimspiele, Auswaertsspiele, ganztagige Sperrungen, zeitliche Sperrungen und leere Daten ergaenzen.

## Empfohlene Umsetzungsreihenfolge

1. Planerzeugung von der Formel in Apps Script verlagern.
2. Trigger-Verwaltung und Bereichsschutz korrigieren.
3. Validierung auf echte Daten beschraenken und gesammelt schreiben.
4. Pflichtfeld- und Zeitvalidierung vervollstaendigen.
5. Spaltenschema und versionierte Migrationen einfuehren.
6. Tests, TSV-Haertung und Dokumentation ergaenzen.
