# Verbesserungsplan

## Naechste Schritte (Top-Prioritaet)

Die folgenden Punkte werden vor allen anderen umgesetzt, da sie die Grundlage fuer alle weiteren Verbesserungen legen.

### Alternativname fuer vereinsinterne Matches

Bei vereinsinternen Begegnungen verwenden Teammitglieder im Gegner-Feld oft Kurznamen (z.B. "E II") statt der offiziellen Teamnamen aus dem Setup. Die kuerzlich eingebauten Validierungs-Ausnahmen (Doppelbuchung bei Gegnern, Nachbarschaftswarnung) greifen dann nicht, weil der Gegner-String nicht mit dem Teamnamen uebereinstimmt.

Die Setup-Spalte D ("Kurzname") ist bisher ungenutzt. Sie wird in `Alternativname` umbenannt und in die Validierung einbezogen.

Ziel:

- Spaltenueberschrift "Kurzname" in "Alternativname" aendern (Setup-Sheet `createSetupSheet()`, TSV-Export).
- `readTeams()` um das Feld `alternativName` (Spalte D) erweitern.
- In `validateAllEntries()` und `checkAdjacentTeams()` den Gegner-Abgleich (`data[i][0] === data[j][1]`) erweitern auf: Teamname ODER Alternativname des Gegners.
- Alternativnamen koennen manuell vergeben oder per Sheet-Formel abgeleitet werden.

Aufwand: ca. 15 Zeilen, kein Risiko.

### TypeScript-Migration und Clasp-Deployment

Der Code wird von einer monolithischen `.gs`-Datei auf mehrere TypeScript-Module umgestellt, angelehnt an die Architektur des `einsatzplaner`-Projekts. Build und Deployment erfolgen ueber `tsc` + `clasp push` statt manuellem Copy-Paste in den Apps-Script-Editor.

Dadurch werden mehrere PLAN-Punkte gleichzeitig adressiert:

- **Zentrales Schema**: Spalten-Enums in `ConfigTypes.ts` (Prioritaet 3)
- **Testbare Kernlogik**: Reine Funktionen extrahiert, vitest-Tests mit TSV-Fixtures (Prioritaet 3)
- **Seed-Daten-Abgrenzung**: `build.py` schreibt `dist/Config.js` -- separate `data/` und `test-data/` (Prioritaet 3)

#### Dateistruktur

```
src/
├── ConfigTypes.ts       # Enums (COL_SETUP, COL_TEAM, COL_SPERRUNGEN, COL_PLAN),
│                          Interfaces (TeamEntry, Sperrung, TeamInfo), Weekday-Typen
├── Config.ts            # CONFIG-Objekt, SEED-Daten (von build.py ueberschrieben)
├── Utils.ts             # isValidDate, isValidTime, timeToFraction, datumToKey,
│                          formatTime, protectRange, createTrigger
├── DataReader.ts        # readTeamNames, readTeams, readSperrungen
├── Validation.ts        # handleEdit, validateAllEntries, validateSperrungen,
│                          checkAdjacentTeams, computeEndzeit
├── SheetBuilder.ts      # setupSheet, createTeamSheet, createSperrungenSheet,
│                          createPlanSheet, seedSheets, resetAll
├── PlanGenerator.ts     # generatePlan
├── DataExporter.ts      # downloadTSV
└── Main.ts              # onOpen, repairValidations, Menue-Handler

test/
├── validation.test.ts   # Reine Logik-Tests: Konfliktpruefung, Datum/Zeit-Helfer,
│                          Nachbarschaftscheck mit vitest + TSV-Fixtures
└── build.test.ts        # Prueft tsc-Output und Config.js-Inhalt
```

#### Build-Pipeline

```
src/*.ts            ──tsc──►  dist/*.js           ──clasp push──►  GAS
build.py + data/   ───────►  dist/Config.js
build.py + test-data/ ────►  dist/Config.js (Test-Daten)
```

- `tsconfig.json`: Target ES2019, module None, strict, `@types/google-apps-script`
- `tsconfig.test.json`: Target ES2022, module ESNext, bundler-Resolution (fuer vitest)
- `/// <reference path>` Direktiven statt ES-Imports (GAS kennt kein Modulsystem)
- `appsscript.json` mit Dateireihenfolge im `dist/`-Ordner

#### Clasp und Umgebungen

- `.clasp.json.prod` / `.clasp.json.test`: gitignored, per npm-Scripts wechselbar
- `npm run deploy` / `npm run deploy:test`: switch + build + push
- `.claspignore`: Kontrolliert, welche `dist/`-Dateien deployed werden

#### package.json

- `@google/clasp`, `typescript`, `@types/google-apps-script`, `vitest`
- Scripts: `build`, `build:test`, `watch`, `deploy`, `deploy:test`, `test`, `test:watch`, `switch:prod`, `switch:test`

#### Besonderheiten

- Reihenfolge der Dateien in `appsscript.json` entspricht den Abhaengigkeiten (ConfigTypes -> Config -> Utils -> DataReader -> Validation -> SheetBuilder -> PlanGenerator -> DataExporter -> Main).
- GAS-spezifische APIs (`SpreadsheetApp`, `getRange`, etc.) sind in Unit-Tests nicht verfuegbar und werden gemockt.
- Der Alternativname-Punkt (s.o.) wird noch vor der Migration auf main umgesetzt, damit er in der Migration bereits als TypeScript-Code vorliegt.

Aufwand: ca. 1-2 Tage, separater Branch `migration/typescript`.

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

1. **Alternativname** fuer vereinsinterne Matches umsetzen.
2. **TypeScript-Migration** auf Branch `migration/typescript` durchfuehren.
3. Planerzeugung von der Formel in Apps Script verlagern.
4. Trigger-Verwaltung und Bereichsschutz korrigieren.
5. Validierung auf echte Daten beschraenken und gesammelt schreiben.
6. Pflichtfeld- und Zeitvalidierung vervollstaendigen.
7. Spaltenschema und versionierte Migrationen einfuehren.
8. Tests, TSV-Haertung und Dokumentation ergaenzen.
