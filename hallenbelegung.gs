/**
 * Hallenbelegung Tischtennis
 * =========================
 * Google Apps Script zur Einrichtung und Verwaltung der Hallenbelegung.
 *
 * Verwendung:
 * 1. Neues Google Sheet erstellen (Tabellenkalkulation)
 * 2. Erweiterungen → Apps Script
 * 3. Diesen Code vollständig einfügen (bestehenden Code ersetzen)
 * 4. setupSheet() auswählen und ausführen (▶️)
 * 5. Berechtigungen erteilen
 * 6. Zurück zum Sheet – alles ist eingerichtet
 *
 * Blätter:
 *  Teams          – Rang, Gruppe, generierter Teamname
 *  Sperrungen     – Gesperrte Tage und Bereiche
 *  Eingabe        – Dateneingabe durch Mannschaftsführer
 *  Belegungsplan  – Öffentliche Kalenderansicht
 */

// ==================== KONFIGURATION ====================

var CONFIG = {
  // Die vier Hallenbereiche
  AREAS: [
    'Große Halle links',
    'kleine Halle',
    'große Halle rechts',
    'große Halle mitte'
  ],

  // Erlaubte Wochentage (JavaScript getDay: 0=So, 1=Mo, 2=Di, 3=Mi, 4=Do, 5=Fr, 6=Sa)
  ALLOWED_WEEKDAYS: [2, 3, 5, 6, 0],

  // Deutsche Wochentags-Namen für Anzeige
  WEEKDAY_NAMES: ['So', 'Mo', 'Di', 'Mi', 'Do', 'Fr', 'Sa'],

  // Mittwochs ist nur dieser Bereich buchbar
  WEDNESDAY_AREA: 'kleine Halle',

  // Dienstags und freitags maximal so viele Bereiche
  MAX_AREAS_TUE_FRI: 2,

  // Vordefinierte Gruppen (kann in Spalte G des Teams-Blattes erweitert werden)
  INITIAL_GROUPS: ['Erwachsene', 'Damen', 'Jugend 19', 'Jugend 15']
};

// ==================== SETUP ====================

function setupSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var locale = ss.getSpreadsheetLocale();
  var isGerman = locale.startsWith('de');
  var sep = isGerman ? ';' : ',';
  var arrSep = isGerman ? '\\' : ',';

  // Bestehende Blätter löschen
  var sheetNames = ['Belegungsplan', 'Eingabe', 'Sperrungen', 'Teams'];
  sheetNames.forEach(function(name) {
    var sheet = ss.getSheetByName(name);
    if (sheet) ss.deleteSheet(sheet);
  });

  // Teams zuerst (Eingabe referenziert Teams für Dropdown)
  createTeamsSheet(ss, sep);
  createSperrungenSheet(ss, sep);
  createEingabeSheet(ss, sep);
  createBelegungsplanSheet(ss, sep, arrSep);

  // Installierbaren Trigger einrichten
  createTrigger();

  var ui = SpreadsheetApp.getUi();
  ui.alert(
    'Setup abgeschlossen!',
    'Die folgenden Blätter wurden erstellt:\n\n' +
    '  1. Teams – Team-Konfiguration (bitte Teams eintragen)\n' +
    '  2. Sperrungen – Gesperrte Tage / Bereiche\n' +
    '  3. Eingabe – Dateneingabe für Mannschaftsführer\n' +
    '  4. Belegungsplan – Kalenderansicht\n\n' +
    'Die Validierung läuft automatisch bei jeder Eingabe.\n\n' +
    'Für die Web-Veröffentlichung:\n' +
    '  Datei → Für das Web veröffentlichen → Belegungsplan',
    ui.ButtonSet.OK
  );
}

// -------------------- Teams-Blatt --------------------

function createTeamsSheet(ss, sep) {
  var sheet = ss.insertSheet('Teams', 0);

  // Überschriften
  var headers = ['Rang', 'Gruppe', 'Teamname', 'Kurzname'];
  sheet.getRange(1, 1, 1, 4)
    .setValues([headers])
    .setFontWeight('bold')
    .setBackground('#E8E8E8');

  // Gruppenliste in Spalte G (für Dropdown-Validierung)
  sheet.getRange(1, 7).setValue('Gruppenliste');
  sheet.getRange(1, 7).setFontWeight('bold');
  var groups = CONFIG.INITIAL_GROUPS;
  for (var i = 0; i < groups.length; i++) {
    sheet.getRange(i + 2, 7).setValue(groups[i]);
  }
  sheet.getRange(1, 7).setNote(
    'Verfügbare Gruppen für das Dropdown.\nNeue Gruppen einfach in Spalte G ergänzen.'
  );

  // Beispiel-Teams (Rang und Gruppe, Namen werden per Formel generiert)
  var exampleData = [
    [1, 'Erwachsene'],
    [2, 'Erwachsene'],
    [3, 'Erwachsene'],
    [1, 'Damen'],
    [2, 'Damen'],
    [1, 'Jugend 19'],
    [1, 'Jugend 15']
  ];
  sheet.getRange(2, 1, exampleData.length, 2).setValues(exampleData);

  // Formel für Teamname (C): Nur berechnen wenn Rang eingetragen ist
  // Rang 1 = nur Gruppenname, sonst "Gruppe RÖMISCH(Rang)"
  sheet.getRange(2, 3, 999, 1)
    .setFormula('=IF(A2=""' + sep + ' ""' + sep + ' IF(A2=1' + sep + ' B2' + sep + ' B2 & " " & ROMAN(A2)))');

  // Formel für Kurzname (D): Nur berechnen wenn Rang eingetragen ist
  sheet.getRange(2, 4, 999, 1)
    .setFormula('=IF(A2=""' + sep + ' ""' + sep + ' LEFT(B2' + sep + '1) & IFERROR(REGEXEXTRACT(B2' + sep + '"\\d+")' + sep + ' "") & " " & ROMAN(A2))');

  // Datenvalidierung: Gruppe (B) = Dropdown aus G2:G
  var gruppeRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(sheet.getRange('G2:G100'))
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 2, 1000, 1).setDataValidation(gruppeRule);

  // Datenvalidierung: Rang (A) = Zahl >= 1
  var rangRule = SpreadsheetApp.newDataValidation()
    .requireNumberGreaterThan(0)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 1, 1000, 1).setDataValidation(rangRule);

  // Kopfzeile schützen
  protectRange(sheet, 'A1:D1');

  // Spaltenbreiten
  sheet.setColumnWidth(1, 60);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(3, 200);
  sheet.setColumnWidth(4, 120);
  sheet.setColumnWidth(7, 150);

  // Spalte G ausblenden
  sheet.hideColumns(7);

  // Zeile 1 fixieren
  sheet.setFrozenRows(1);
}

// -------------------- Sperrungen-Blatt --------------------

function createSperrungenSheet(ss, sep) {
  var sheet = ss.insertSheet('Sperrungen', 1);

  var headers = ['Datum', 'Bereich', 'Grund'];
  sheet.getRange(1, 1, 1, 3)
    .setValues([headers])
    .setFontWeight('bold')
    .setBackground('#E8E8E8');

  // Datenvalidierung: Bereich
  var areas = CONFIG.AREAS.concat(['Alle']);
  var bereichRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(areas)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 2, 1000, 1).setDataValidation(bereichRule);

  // Datumsformat
  sheet.getRange(2, 1, 1000, 1).setNumberFormat('DD.MM.YYYY');

  // Kopfzeile schützen
  protectRange(sheet, 'A1:C1');

  // Spaltenbreiten
  sheet.setColumnWidth(1, 120);
  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(3, 300);

  // Hinweis
  sheet.getRange(1, 3).setNote(
    '"Alle" = kompletter Tag gesperrt.\n' +
    'Einzelner Bereich = nur dieser Bereich gesperrt.\n' +
    'Grund ist optional.'
  );

  sheet.setFrozenRows(1);
}

// -------------------- Eingabe-Blatt --------------------

function createEingabeSheet(ss, sep) {
  var sheet = ss.insertSheet('Eingabe', 2);

  var headers = ['Datum', 'Startzeit', 'Team', 'Heim/Auswärts', 'Gegner', 'Bereich', 'Kommentar', '\u26A0\uFE0F Status'];
  sheet.getRange(1, 1, 1, 8)
    .setValues([headers])
    .setFontWeight('bold')
    .setBackground('#E8E8E8');

  // Spaltenformatierung
  sheet.getRange(2, 1, 1000, 1).setNumberFormat('DD.MM.YYYY');
  sheet.getRange(2, 2, 1000, 1).setNumberFormat('HH:MM');

  // Datenvalidierung: Datum (A) – nur erlaubte Wochentage
  // WOCHENTAG: 1=So, 2=Mo, 3=Di, 4=Mi, 5=Do, 6=Fr, 7=Sa
  // Erlaubt: 3(Di), 4(Mi), 6(Fr), 7(Sa), 1(So)
  var weekdayFormula = '=ODER(WOCHENTAG(A2)=3' + sep + ' WOCHENTAG(A2)=4' + sep +
    ' WOCHENTAG(A2)=6' + sep + ' WOCHENTAG(A2)=7' + sep + ' WOCHENTAG(A2)=1)';
  var weekdayRule = SpreadsheetApp.newDataValidation()
    .requireFormulaSatisfied(weekdayFormula)
    .setAllowInvalid(true)
    .setHelpText('Nur Di, Mi, Fr, Sa, So erlaubt')
    .build();
  sheet.getRange(2, 1, 1000, 1).setDataValidation(weekdayRule);

  // Datenvalidierung: Team (C) – aus Teams-Blatt, Spalte C (Teamname)
  var teamsSheet = ss.getSheetByName('Teams');
  var teamRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(teamsSheet.getRange('C2:C1000'))
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 3, 1000, 1).setDataValidation(teamRule);

  // Datenvalidierung: Heim/Auswärts (D)
  var haRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Heim', 'Auswärts'])
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 4, 1000, 1).setDataValidation(haRule);

  // Datenvalidierung: Bereich (F) – bei Auswärtsspiel optional
  var areaRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(CONFIG.AREAS)
    .setAllowInvalid(true)
    .build();
  sheet.getRange(2, 6, 1000, 1).setDataValidation(areaRule);

  // Bedingte Formatierung: Heim = grün
  var heimRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$D2="Heim"')
    .setBackground('#C8E6C9')
    .setRanges([sheet.getRange('A2:H1000')])
    .build();

  // Bedingte Formatierung: Auswärts = blau
  var auswaertsRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$D2="Auswärts"')
    .setBackground('#BBDEFB')
    .setRanges([sheet.getRange('A2:H1000')])
    .build();

  // Bedingte Formatierung: Validierungsfehler = rot
  var fehlerRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=REGEXMATCH($H2' + sep + ' "❌")')
    .setBackground('#FFCDD2')
    .setRanges([sheet.getRange('A2:H1000')])
    .build();

  // Bedingte Formatierung: Ersatzspieler-Warnung = orange
  var warnungRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=REGEXMATCH($H2' + sep + ' "⚠️")')
    .setBackground('#FFE0B2')
    .setRanges([sheet.getRange('A2:H1000')])
    .build();

  var rules = sheet.getConditionalFormatRules();
  rules.push(heimRule, auswaertsRule, fehlerRule, warnungRule);
  sheet.setConditionalFormatRules(rules);

  // Kopfzeile schützen
  protectRange(sheet, 'A1:H1');

  // Spaltenbreiten
  sheet.setColumnWidth(1, 110);
  sheet.setColumnWidth(2, 90);
  sheet.setColumnWidth(3, 180);
  sheet.setColumnWidth(4, 120);
  sheet.setColumnWidth(5, 200);
  sheet.setColumnWidth(6, 200);
  sheet.setColumnWidth(7, 250);
  sheet.setColumnWidth(8, 350);

  // Kopfzeile fixieren
  sheet.setFrozenRows(1);

  // Hinweis auf Status-Spalte
  sheet.getRange(1, 8).setNote(
    'Automatisch vom Script befüllt.\n' +
    '❌ = Validierungsfehler\n' +
    '⚠️ = Benachbartes Team spielt am selben Tag'
  );

  // Infotext in Zelle I1 (rechts neben Status)
  sheet.getRange(1, 9).setValue(
    'Hinweise:\n' +
    '- Nur Di/Mi/Fr/Sa/So erlaubt\n' +
    '- Mi: nur "kleine Halle"\n' +
    '- Di+Fr: max. 2 Bereiche\n' +
    '- Heimspiel: Bereich erforderlich\n' +
    '- Auswärtsspiel: Bereich frei lassen'
  );
  sheet.getRange(1, 9).setFontSize(9);
  sheet.getRange(1, 9).setFontColor('#999999');
  sheet.setColumnWidth(9, 220);
}

// -------------------- Belegungsplan-Blatt --------------------

function createBelegungsplanSheet(ss, sep, arrSep) {
  var sheet = ss.insertSheet('Belegungsplan', 3);

  // Checkbox in A1 zum Umschalten Heim/Alle
  var checkbox = sheet.getRange('A1');
  checkbox.insertCheckboxes();
  checkbox.setValue(false);

  // Label neben Checkbox
  sheet.getRange('B1').setValue('Nur Hallenbelegung (Heimspiele)');
  sheet.getRange('B1').setFontStyle('italic');
  sheet.getRange('B1').setFontColor('#555555');

  // Überschriften Zeile 2
  var headers = ['Datum', 'Tag', 'Startzeit', 'Bereich', 'Team', 'H/A', 'Gegner', 'Kommentar', 'Status'];
  sheet.getRange(2, 1, 1, 9)
    .setValues([headers])
    .setFontWeight('bold')
    .setBackground('#E8E8E8');

  // QUERY-Formel in A3: zeigt alle Spiele oder nur Heimspiele je nach Checkbox
  // Die Formel erzeugt eine zusätzliche Spalte (Col9) mit dem Wochentag
  // arrSep = Array-Spaltentrenner (\ für DE, , für EN)
  // sep    = Funktionsargument-Trenner (; für DE, , für EN)
  var formula = '=IF(A1' + sep +
    'QUERY({Eingabe!A2:H' + arrSep + ' ARRAYFORMULA(TEXT(Eingabe!A2:A' + sep + '"TTT"))}' + sep +
    '"SELECT Col1, Col9, Col2, Col6, Col3, Col4, Col5, Col7, Col8 ' +
    'WHERE Col1 IS NOT NULL AND Col4=\'Heim\' ' +
    'ORDER BY Col1, Col2"' + sep + ' 1)' + sep +
    'QUERY({Eingabe!A2:H' + arrSep + ' ARRAYFORMULA(TEXT(Eingabe!A2:A' + sep + '"TTT"))}' + sep +
    '"SELECT Col1, Col9, Col2, Col6, Col3, Col4, Col5, Col7, Col8 ' +
    'WHERE Col1 IS NOT NULL ' +
    'ORDER BY Col1, Col2"' + sep + ' 1)' +
    ')';

  sheet.getRange(3, 1).setFormula(formula);

  // Datumsformat für QUERY-Ausgabe
  sheet.getRange(3, 1, 1000, 1).setNumberFormat('DD.MM.YYYY');
  sheet.getRange(3, 3, 1000, 1).setNumberFormat('HH:MM');

  // Bedingte Formatierung: Heim = grün
  var heimRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$F3="Heim"')
    .setBackground('#C8E6C9')
    .setRanges([sheet.getRange('A3:I1000')])
    .build();

  // Bedingte Formatierung: Auswärts = blau
  var auswaertsRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$F3="Auswärts"')
    .setBackground('#BBDEFB')
    .setRanges([sheet.getRange('A3:I1000')])
    .build();

  // Bedingte Formatierung: Ersatzspieler-Warnung = orange
  var warnungRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=REGEXMATCH($I3' + sep + ' "⚠️")')
    .setBackground('#FFE0B2')
    .setRanges([sheet.getRange('A3:I1000')])
    .build();

  var rules = sheet.getConditionalFormatRules();
  rules.push(heimRule, auswaertsRule, warnungRule);
  sheet.setConditionalFormatRules(rules);

  // Spaltenbreiten
  sheet.setColumnWidth(1, 110);
  sheet.setColumnWidth(2, 50);
  sheet.setColumnWidth(3, 90);
  sheet.setColumnWidth(4, 200);
  sheet.setColumnWidth(5, 180);
  sheet.setColumnWidth(6, 60);
  sheet.setColumnWidth(7, 200);
  sheet.setColumnWidth(8, 250);
  sheet.setColumnWidth(9, 300);

  // Status-Spalte (I) ausblenden – wird nur für bedingte Formatierung gebraucht
  sheet.hideColumns(9);

  // Zeilen 1-2 fixieren (Checkbox + Überschriften)
  sheet.setFrozenRows(2);

  // Schutz für Kopfbereich
  protectRange(sheet, 'A1:I2');

  // Hinweis
  sheet.getRange(2, 9).setNote(
    'Diese Spalte ist ausgeblendet.\n' +
    'Sie enthält den Status für die farbliche Markierung.'
  );
}

// ==================== VALIDIERUNG ====================

function onEdit(e) {
  var sheet = e.range.getSheet();
  if (sheet.getName() !== 'Eingabe') return;
  if (e.range.getRow() < 2) return;

  var col = e.range.getColumn();
  // Relevante Spalten: Datum(1), Startzeit(2), Team(3), H/A(4), Bereich(6)
  var relevantCols = [1, 2, 3, 4, 6];
  if (relevantCols.indexOf(col) === -1) return;

  validateAllEntries();
}

function validateAllEntries() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Eingabe');
  if (!sheet) return;

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  var data = sheet.getRange(2, 1, lastRow - 1, 8).getValues();

  // Teams und Sperrungen vorab lesen
  var teams = readTeams(ss);
  var sperrungen = readSperrungen(ss);

  // Sammle alle Fehler pro Zeile
  var statusMessages = {};
  for (var i = 0; i < data.length; i++) {
    statusMessages[i] = [];
  }

  // --- 1. Einzelvalidierung jeder Zeile ---
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var datum = row[0];
    var team = row[2];
    var ha = row[3];
    var bereich = row[5];

    if (!isValidDate(datum)) continue;

    var weekday = datum.getDay();
    var dateKey = datumToKey(datum);

    // 1a. Wochentag prüfen (nur für Heimspiele relevant)
    if (ha === 'Heim' && CONFIG.ALLOWED_WEEKDAYS.indexOf(weekday) === -1) {
      statusMessages[i].push('❌ Nur Di, Mi, Fr, Sa, So erlaubt');
    }

    // 1b. Mittwoch: nur kleine Halle
    if (weekday === 3 && ha === 'Heim' && bereich && bereich !== CONFIG.WEDNESDAY_AREA) {
      statusMessages[i].push('❌ Mittwochs nur "' + CONFIG.WEDNESDAY_AREA + '" buchbar');
    }

    // 1c. Dienstag/Freitag: max. 2 Bereiche
    if ((weekday === 2 || weekday === 5) && ha === 'Heim' && bereich) {
      var areasOnDate = {};
      for (var j = 0; j < data.length; j++) {
        if (isValidDate(data[j][0]) && datumToKey(data[j][0]) === dateKey &&
            data[j][3] === 'Heim' && data[j][5]) {
          areasOnDate[data[j][5]] = true;
        }
      }
      // Aktuelle Zeile ist bereits enthalten (data[j] includes current row)
      var uniqueCount = Object.keys(areasOnDate).length;
      if (uniqueCount > CONFIG.MAX_AREAS_TUE_FRI) {
        statusMessages[i].push('❌ Di/Fr: max. ' + CONFIG.MAX_AREAS_TUE_FRI +
          ' Bereiche buchbar (' + uniqueCount + ' wären belegt)');
      }
    }

    // 1d. Doppelbuchung prüfen
    if (ha === 'Heim' && bereich) {
      for (var k = 0; k < data.length; k++) {
        if (k !== i && isValidDate(data[k][0]) &&
            datumToKey(data[k][0]) === dateKey &&
            data[k][3] === 'Heim' &&
            data[k][5] === bereich) {
          statusMessages[i].push('❌ "' + bereich + '" an diesem Tag bereits belegt');
          break;
        }
      }
    }

    // 1e. Sperrungen prüfen
    for (var s = 0; s < sperrungen.length; s++) {
      var sDatum = sperrungen[s].datum;
      var sBereich = sperrungen[s].bereich;
      var sGrund = sperrungen[s].grund;
      if (isValidDate(sDatum) && datumToKey(sDatum) === dateKey) {
        var blocked = false;
        if (sBereich === 'Alle') {
          blocked = true;
        } else if (bereich && sBereich === bereich) {
          blocked = true;
        }
        if (blocked) {
          var msg = '❌ ';
          if (sBereich === 'Alle') {
            msg += 'Tag gesperrt';
          } else {
            msg += '"' + sBereich + '" gesperrt';
          }
          if (sGrund) msg += ': ' + sGrund;
          statusMessages[i].push(msg);
        }
      }
    }
  }

  // --- 2. Benachbarte Teams prüfen (gleiche Gruppe, Rangunterschied 1) ---
  var adjacentMessages = checkAdjacentTeams(data, teams);

  // --- 3. Status-Meldungen schreiben ---
  for (var i = 0; i < data.length; i++) {
    var allMessages = statusMessages[i].slice();

    // Adjacent-Warnungen anfügen
    if (adjacentMessages[i]) {
      allMessages = allMessages.concat(adjacentMessages[i]);
    }

    var statusText = allMessages.join(' | ');
    sheet.getRange(i + 2, 8).setValue(statusText);
  }
}

// -------------------- Hilfsfunktionen Validierung --------------------

function checkAdjacentTeams(data, teams) {
  var messages = {}; // index -> [message1, message2, ...]

  // Gruppiere Einträge nach Datum
  var entriesByDate = {};
  for (var i = 0; i < data.length; i++) {
    if (!isValidDate(data[i][0]) || !data[i][2]) continue;
    var key = datumToKey(data[i][0]);
    if (!entriesByDate[key]) entriesByDate[key] = [];
    entriesByDate[key].push({ index: i, teamName: data[i][2] });
  }

  // Für jedes Datum: Teams gleicher Gruppe mit Rangunterschied 1 finden
  for (var dateKey in entriesByDate) {
    var entries = entriesByDate[dateKey];

    // Nach Gruppe gruppieren
    var byGroup = {};
    for (var e = 0; e < entries.length; e++) {
      var teamName = entries[e].teamName;
      var teamInfo = teams[teamName];
      if (!teamInfo) continue;
      var g = teamInfo.gruppe;
      if (!byGroup[g]) byGroup[g] = [];
      byGroup[g].push({
        index: entries[e].index,
        rang: teamInfo.rang,
        name: teamName
      });
    }

    // In jeder Gruppe nach benachbarten Rängen suchen
    for (var gruppe in byGroup) {
      var gEntries = byGroup[gruppe];
      for (var a = 0; a < gEntries.length; a++) {
        for (var b = a + 1; b < gEntries.length; b++) {
          if (Math.abs(gEntries[a].rang - gEntries[b].rang) === 1) {
            var idxA = gEntries[a].index;
            var idxB = gEntries[b].index;

            if (!messages[idxA]) messages[idxA] = [];
            if (!messages[idxB]) messages[idxB] = [];

            var msgA = '⚠️ Am selben Tag spielt: ' + gEntries[b].name;
            var msgB = '⚠️ Am selben Tag spielt: ' + gEntries[a].name;

            if (messages[idxA].indexOf(msgA) === -1) messages[idxA].push(msgA);
            if (messages[idxB].indexOf(msgB) === -1) messages[idxB].push(msgB);
          }
        }
      }
    }
  }

  return messages;
}

// -------------------- Daten lesen --------------------

function readTeams(ss) {
  var sheet = ss.getSheetByName('Teams');
  if (!sheet) return {};

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return {};

  // Spalten: A=Rang, B=Gruppe, C=Teamname, D=Kurzname
  var raw = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
  var teams = {};
  for (var i = 0; i < raw.length; i++) {
    var rang = raw[i][0];
    var gruppe = raw[i][1];
    var teamname = raw[i][2]; // per Formel generiert
    if (rang && gruppe && teamname) {
      teams[teamname] = { rang: rang, gruppe: gruppe };
    }
  }
  return teams;
}

function readSperrungen(ss) {
  var sheet = ss.getSheetByName('Sperrungen');
  if (!sheet) return [];

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  var raw = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  var sperrungen = [];
  for (var i = 0; i < raw.length; i++) {
    if (raw[i][0]) {
      sperrungen.push({
        datum: raw[i][0],
        bereich: raw[i][1],
        grund: raw[i][2]
      });
    }
  }
  return sperrungen;
}

// -------------------- Allgemeine Hilfsfunktionen --------------------

function isValidDate(d) {
  return d instanceof Date && !isNaN(d.getTime());
}

function datumToKey(datum) {
  if (!isValidDate(datum)) return '';
  return datum.getFullYear() + '-' +
    ('0' + (datum.getMonth() + 1)).slice(-2) + '-' +
    ('0' + datum.getDate()).slice(-2);
}

function protectRange(sheet, rangeA1) {
  var protection = sheet.getRange(rangeA1).protect();
  protection.setDescription('Kopfzeile geschützt');
  protection.setWarningOnly(true);
}

function createTrigger() {
  // Bestehende Trigger für diese Funktion entfernen
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(t) {
    if (t.getHandlerFunction() === 'onEdit') {
      ScriptApp.deleteTrigger(t);
    }
  });

  // Neuen installierbaren Trigger erstellen
  ScriptApp.newTrigger('onEdit')
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onEdit()
    .create();
}

// ==================== MENÜ ====================

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Hallenbelegung')
    .addItem('Setup (neu einrichten)', 'setupSheet')
    .addItem('Jetzt validieren', 'validateAllEntries')
    .addSeparator()
    .addItem('Web-Veröffentlichung einrichten...', 'showPublishHelp')
    .addToUi();
}

function showPublishHelp() {
  var ui = SpreadsheetApp.getUi();
  ui.alert(
    'Web-Veröffentlichung',
    'So veröffentlichen Sie den Belegungsplan:\n\n' +
    '1. Datei → Freigeben → Für das Web veröffentlichen\n' +
    '2. Blatt "Belegungsplan" auswählen\n' +
    '3. "Gesamtes Dokument" → "Belegungsplan"\n' +
    '4. "Veröffentlichen" klicken\n' +
    '5. Den generierten Link kopieren und teilen\n\n' +
    'Tipp: Vor dem Veröffentlichen die Checkbox in A1 ' +
    'auf "Nur Hallenbelegung" setzen, wenn nur die ' +
    'Hallenbelegung (ohne Auswärtsspiele) sichtbar sein soll.',
    ui.ButtonSet.OK
  );
}
