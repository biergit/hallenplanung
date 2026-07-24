/**
 * Hallen/Spielplanung Tischtennis
 * ===============================
 * Google Apps Script zur Einrichtung und Verwaltung der Hallen/Spielplanung.
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
 *  Teams                             – Rang, Gruppe, generierter Teamname
 *  Sperrungen/Anderweitige Belegungen – Gesperrte Tage, Bereiche und Zeiträume
 *  Eingabe                           – Dateneingabe durch Mannschaftsführer
 *  Hallen/Spielplan                  – Öffentliche Kalenderansicht
 */

// ==================== KONFIGURATION ====================

var CONFIG = {
  AREAS: [
    'Große Halle links',
    'Kleine Halle',
    'Große Halle rechts',
    'Große Halle Mitte'
  ],

  // Erlaubte Wochentage (JavaScript getDay: 0=So, 1=Mo, 2=Di, 3=Mi, 4=Do, 5=Fr, 6=Sa)
  ALLOWED_WEEKDAYS: [2, 3, 5, 6, 0],

  // Mittwochs ist nur dieser Bereich buchbar
  WEDNESDAY_AREA: 'Kleine Halle',

  // Dienstags und freitags maximal so viele Bereiche
  MAX_AREAS_TUE_FRI: 2,

  // Vordefinierte Gruppen (kann in Spalte G des Teams-Blattes erweitert werden)
  INITIAL_GROUPS: ['Erwachsene', 'Damen', 'Jugend 19', 'Jugend 15'],

  // Blattnamen
  SHEET_SETUP: 'Setup',
  SHEET_SPERRUNGEN: 'Sperrungen/Anderweitige Belegungen',
  SHEET_PLAN: 'Hallen/Spielplan',

  // Standard-Spieldauer in Stunden (konfigurierbar in Setup!F1)
  GAME_DURATION_HOURS: 2.5,

  // Maximale Zeilenanzahl für Datenbereiche, Formatierungen und Validierungen
  MAX_ROWS: 1000
};
// --- SEED DATA BEGIN ---
var SEED_SETUP = JSON.parse("[]");
var SEED_EINGABE = JSON.parse("[]");
var SEED_SPERRUNGEN = JSON.parse("[]");
// --- SEED DATA END ---

// ==================== SETUP ====================

function setupSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var locale = ss.getSpreadsheetLocale();
  var isGerman = locale.startsWith('de');
  var sep = isGerman ? ';' : ',';

  if (!ss.getSheetByName(CONFIG.SHEET_SETUP)) {
    createSetupSheet(ss, sep);
  }

  SpreadsheetApp.flush();
  var teamNames = readTeamNames(ss);
  for (var i = 0; i < teamNames.length; i++) {
    var ts = ss.getSheetByName(teamNames[i]);
    if (ts) {
      upgradeTeamSheet(ts, sep);
    } else {
      createTeamSheet(ss, teamNames[i], sep);
    }
  }

  var sperrSheet = ss.getSheetByName(CONFIG.SHEET_SPERRUNGEN);
  if (sperrSheet) {
    upgradeSperrungenSheet(ss);
  } else {
    createSperrungenSheet(ss);
  }

  var oldPlan = ss.getSheetByName('Belegungsplan');
  if (oldPlan) ss.deleteSheet(oldPlan);
  var planSheet = ss.getSheetByName(CONFIG.SHEET_PLAN);
  if (planSheet) {
    upgradeHallenSpielplanSheet(ss, sep);
  } else {
    createHallenSpielplanSheet(ss, sep);
  }

  reorderSheets(ss, teamNames);

  generatePlan();
  createTrigger();

  var msg = 'Setup abgeschlossen!\n\n' +
    teamNames.length + ' Team-Blätter\n' +
    '  ' + CONFIG.SHEET_SPERRUNGEN + '\n' +
    '  ' + CONFIG.SHEET_PLAN + '\n' +
    '  ' + CONFIG.SHEET_SETUP + '\n\n' +
    'Für die Web-Veröffentlichung:\n' +
    '  Datei → Für das Web veröffentlichen → ' + CONFIG.SHEET_PLAN;
  try {
    SpreadsheetApp.getUi().alert(msg);
  } catch (e) {
    Logger.log(msg);
  }
}

function reorderSheets(ss, teamNames) {
  var ordered = [];

  for (var i = 0; i < teamNames.length; i++) {
    var s = ss.getSheetByName(teamNames[i]);
    if (s) ordered.push(s);
  }

  var sperr = ss.getSheetByName(CONFIG.SHEET_SPERRUNGEN);
  if (sperr) ordered.push(sperr);
  var plan = ss.getSheetByName(CONFIG.SHEET_PLAN);
  if (plan) ordered.push(plan);
  var setup = ss.getSheetByName(CONFIG.SHEET_SETUP);
  if (setup) ordered.push(setup);

  for (var i = ordered.length - 1; i >= 0; i--) {
    ss.setActiveSheet(ordered[i]);
    ss.moveActiveSheet(i + 1);
  }
}

// -------------------- Seed / Reset --------------------

function resetAll() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var locale = ss.getSpreadsheetLocale();
  var isGerman = locale.startsWith('de');
  var sep = isGerman ? ';' : ',';

  var keep = [CONFIG.SHEET_PLAN];
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (keep.indexOf(sheets[i].getName()) < 0) {
      ss.deleteSheet(sheets[i]);
    }
  }

  createSetupSheet(ss, sep);
  _seedSetup(ss);
  SpreadsheetApp.flush();

  var teamNames = readTeamNames(ss);
  for (var i = 0; i < teamNames.length; i++) {
    createTeamSheet(ss, teamNames[i], sep);
  }

  createSperrungenSheet(ss);
  _seedSperrungen(ss);

  var planSheet = ss.getSheetByName(CONFIG.SHEET_PLAN);
  if (planSheet) {
    upgradeHallenSpielplanSheet(ss, sep);
  } else {
    createHallenSpielplanSheet(ss, sep);
  }

  reorderSheets(ss, teamNames);
  createTrigger();
  _seedEingabe(ss);
  validateSperrungen();
  validateAllEntries();
  generatePlan();

  var msg = 'Alles neu angelegt inkl. Seed-Daten.\n\n' +
    teamNames.length + ' Team-Blätter\n' +
    '  ' + CONFIG.SHEET_SPERRUNGEN + '\n' +
    '  ' + CONFIG.SHEET_PLAN + '\n' +
    '  ' + CONFIG.SHEET_SETUP;
  try {
    SpreadsheetApp.getUi().alert(msg);
  } catch (e) {
    Logger.log(msg);
  }
}

function createSheetsOnly() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var locale = ss.getSpreadsheetLocale();
  var isGerman = locale.startsWith('de');
  var sep = isGerman ? ';' : ',';

  var keep = [CONFIG.SHEET_PLAN];
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (keep.indexOf(sheets[i].getName()) < 0) {
      ss.deleteSheet(sheets[i]);
    }
  }

  createSetupSheet(ss, sep);
  var teamNames = readTeamNames(ss);
  for (var i = 0; i < teamNames.length; i++) {
    createTeamSheet(ss, teamNames[i], sep);
  }

  createSperrungenSheet(ss);

  var planSheet = ss.getSheetByName(CONFIG.SHEET_PLAN);
  if (planSheet) {
    upgradeHallenSpielplanSheet(ss, sep);
  } else {
    createHallenSpielplanSheet(ss, sep);
  }

  createTrigger();
  reorderSheets(ss, teamNames);
  generatePlan();

  var msg = 'Blätter neu angelegt (ohne Seed-Daten).';
  try {
    SpreadsheetApp.getUi().alert(msg);
  } catch (e) {
    Logger.log(msg);
  }
}

function seedSheets(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  if (SEED_SETUP.length > 0) {
    _seedSetup(ss);
    SpreadsheetApp.flush();
  }
  if (SEED_EINGABE.length > 0) _seedEingabe(ss);
  if (SEED_SPERRUNGEN.length > 0) _seedSperrungen(ss);
  validateSperrungen();
  validateAllEntries();
}

function _seedSetup(ss) {
  var sheet = ss.getSheetByName(CONFIG.SHEET_SETUP);
  if (!sheet) return;

  var abData = [];
  var hData = [];
  for (var i = 0; i < SEED_SETUP.length; i++) {
    abData.push([parseInt(SEED_SETUP[i][0]), SEED_SETUP[i][1]]);
    hData.push([SEED_SETUP[i][4] || '']);
  }
  if (abData.length > 0) {
    var abRange = sheet.getRange(2, 1, abData.length, 2);
    abRange.clearDataValidations();
    abRange.setValues(abData);
    // Validierungen wiederherstellen
    sheet.getRange(2, 1, abData.length, 1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireNumberGreaterThan(0).setAllowInvalid(false).build());
    sheet.getRange(2, 2, abData.length, 1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInRange(sheet.getRange('G2:G100')).setAllowInvalid(false).build());
  }
  if (hData.length > 0) {
    sheet.getRange(2, 8, hData.length, 1).setValues(hData);
  }
}

function _seedEingabe(ss) {
  var teamNames = readTeamNames(ss);
  var byTeam = {};
  for (var i = 0; i < SEED_EINGABE.length; i++) {
    var row = SEED_EINGABE[i];
    var team = row[0];
    if (teamNames.indexOf(team) < 0) continue;
    if (!byTeam[team]) byTeam[team] = [];
    byTeam[team].push(row);
  }

  for (var t = 0; t < teamNames.length; t++) {
    var team = teamNames[t];
    var entries = byTeam[team] || [];
    if (entries.length === 0) continue;
    var sheet = ss.getSheetByName(team);
    if (!sheet) continue;

    var data = [];
    for (var i = 0; i < entries.length; i++) {
      var row = entries[i];
      var datum = _parseDate(row[2]);
      var startzeit = _parseTime(row[3]);
      var endzeit = _parseTime(row[4]);
      data.push([row[1], datum, startzeit, endzeit, row[5], row[6], row[7], '']);
    }
    if (data.length > 0) {
      var range = sheet.getRange(2, 1, data.length, 8);
      range.clearDataValidations();
      range.setValues(data);
      sheet.getRange(2, 2, data.length, 1).setDataValidation(
        SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(true).setHelpText('Datum eingeben oder auswählen').build());
      sheet.getRange(2, 5, data.length, 1).setDataValidation(
        SpreadsheetApp.newDataValidation().requireValueInList(['Heim', 'Auswärts']).setAllowInvalid(false).build());
      sheet.getRange(2, 6, data.length, 1).setDataValidation(
        SpreadsheetApp.newDataValidation().requireValueInList(CONFIG.AREAS).setAllowInvalid(false).build());
    }
  }
}

function _seedSperrungen(ss) {
  var sheet = ss.getSheetByName(CONFIG.SHEET_SPERRUNGEN);
  if (!sheet) return;

  var data = [];
  for (var i = 0; i < SEED_SPERRUNGEN.length; i++) {
    var row = SEED_SPERRUNGEN[i];
    var datum = _parseDate(row[0]);
    var startzeit = _parseTime(row[1]);
    var endzeit = _parseTime(row[2]);
    data.push([datum, startzeit, endzeit, row[3], row[4]]);
  }
  if (data.length > 0) {
    var range = sheet.getRange(2, 1, data.length, 5);
    range.clearDataValidations();
    range.setValues(data);
    // Validierungen wiederherstellen
    sheet.getRange(2, 1, data.length, 1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(true).setHelpText('Datum eingeben oder auswählen').build());
    sheet.getRange(2, 4, data.length, 1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(CONFIG.AREAS.concat(['Alle'])).setAllowInvalid(false).build());
  }
}

function _parseDate(str) {
  if (!str || typeof str !== 'string') return '';
  var parts = str.split('.');
  if (parts.length === 3) {
    return new Date(parseInt(parts[2]), parseInt(parts[1]) - 1, parseInt(parts[0]));
  }
  return '';
}

function _parseTime(str) {
  if (!str || typeof str !== 'string') return '';
  var parts = str.split(':');
  if (parts.length === 2) {
    return (parseInt(parts[0]) * 60 + parseInt(parts[1])) / 1440;
  }
  var num = parseFloat(str.replace(',', '.'));
  if (!isNaN(num) && num > 0 && num < 24) {
    var h = Math.floor(num);
    var m = Math.round((num - h) * 60);
    return (h * 60 + m) / 1440;
  }
  return '';
}

// -------------------- Setup-Blatt --------------------

function createSetupSheet(ss, sep) {
  var sheet = ss.insertSheet(CONFIG.SHEET_SETUP);

  var headers = ['Rang', 'Gruppe', 'Teamname', 'Kurzname'];
  sheet.getRange(1, 1, 1, 4)
    .setValues([headers])
    .setFontWeight('bold')
    .setBackground('#E8E8E8');

  // Spieldauer-Konfiguration in E1/F1
  sheet.getRange('E1').setValue('Standard-Spieldauer (h):');
  sheet.getRange('E1').setFontWeight('bold');
  sheet.getRange('F1').setValue(CONFIG.GAME_DURATION_HOURS);
  sheet.getRange('F1').setNumberFormat('0.0');
  sheet.getRange('F1').setNote('Standard-Spieldauer für alle Teams. Kann pro Team in Spalte H überschrieben werden.');

  // Spalte H: Team-spezifische Spieldauer
  sheet.getRange(1, 8).setValue('Spieldauer (h)');
  sheet.getRange(1, 8).setFontWeight('bold').setBackground('#E8E8E8');
  sheet.getRange(1, 8).setNote('Optionale team-spezifische Spieldauer in Stunden. Leer lassen für Standard aus F1.');
  sheet.getRange(2, 8, CONFIG.MAX_ROWS, 1).setNumberFormat('0.0');

  sheet.getRange(1, 7).setValue('Gruppenliste');
  sheet.getRange(1, 7).setFontWeight('bold');
  var groups = CONFIG.INITIAL_GROUPS;
  for (var i = 0; i < groups.length; i++) {
    sheet.getRange(i + 2, 7).setValue(groups[i]);
  }
  sheet.getRange(1, 7).setNote(
    'Verfügbare Gruppen für das Dropdown.\nNeue Gruppen einfach in Spalte G ergänzen.'
  );

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

  sheet.getRange(2, 3, 999, 1)
    .setFormula('=IF(A2=""' + sep + ' ""' + sep + ' IF(A2=1' + sep + ' B2' + sep + ' B2 & " " & ROMAN(A2)))');

  sheet.getRange(2, 4, 999, 1)
    .setFormula('=IF(A2=""' + sep + ' ""' + sep + ' LEFT(B2' + sep + '1) & IFERROR(REGEXEXTRACT(B2' + sep + '"\\d+")' + sep + ' "") & " " & ROMAN(A2))');

  var gruppeRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(sheet.getRange('G2:G100'))
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 2, CONFIG.MAX_ROWS, 1).setDataValidation(gruppeRule);

  var rangRule = SpreadsheetApp.newDataValidation()
    .requireNumberGreaterThan(0)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 1, CONFIG.MAX_ROWS, 1).setDataValidation(rangRule);

  protectRange(sheet, 'A1:D1');

  sheet.setColumnWidth(1, 60);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(3, 200);
  sheet.setColumnWidth(4, 120);
  sheet.setColumnWidth(7, 150);

  sheet.hideColumns(7);

  var instructions = [
    ['SLICER-FILTER (Web-Ansicht):', ''],
    ['1.', 'Hallen/Spielplan-Tab öffnen.'],
    ['2.', 'Menü Daten → Slicer.'],
    ['3.', 'Slicer für Team (Spalte E) anlegen.'],
    ['4.', 'Slicer für H/A (Spalte F) anlegen.'],
    ['5.', 'Slicer für Typ (Spalte I: leer=Spiel, »gesperrt«=Sperrung).'],
    ['', 'Slicer funktionieren in der veröffentlichten HTML-Ansicht.'],
    ['', ''],
    ['WEB-VERÖFFENTLICHUNG:', 'Datei → Freigeben → Für das Web veröffentlichen'],
    ['', 'Blatt "Hallen/Spielplan" auswählen.'],
    ['', ''],
    ['FREIGABE:', 'Datei → Freigeben → E-Mail der Mannschaftsführer hinzufügen.'],
    ['', 'Berechtigung: "Bearbeiter".'],
  ];

  var range = sheet.getRange(2, 10, instructions.length, 2);
  range.setValues(instructions);
  range.setFontSize(10);
  // Titel-Spalte fett
  for (var r = 0; r < instructions.length; r++) {
    if (instructions[r][0]) {
      sheet.getRange(r + 2, 10).setFontWeight('bold');
    }
  }

  // Spalte K breiter machen
  sheet.setColumnWidth(11, 480);

  sheet.setFrozenRows(1);
}

// -------------------- Sperrungen-Blatt --------------------

function createSperrungenSheet(ss) {
  var sheet = ss.insertSheet(CONFIG.SHEET_SPERRUNGEN);

  var headers = ['Datum', 'Startzeit', 'Endzeit', 'Bereich', 'Kommentar', '\u26A0\uFE0F Status'];
  sheet.getRange(1, 1, 1, 6)
    .setValues([headers])
    .setFontWeight('bold')
    .setBackground('#E8E8E8');

  var dateRule = SpreadsheetApp.newDataValidation()
    .requireDate()
    .setAllowInvalid(true)
    .setHelpText('Datum eingeben oder auswählen')
    .build();
  sheet.getRange(2, 1, CONFIG.MAX_ROWS, 1).setDataValidation(dateRule);

  sheet.getRange(2, 2, CONFIG.MAX_ROWS, 1).setNumberFormat('HH:MM');
  sheet.getRange(2, 3, CONFIG.MAX_ROWS, 1).setNumberFormat('HH:MM');

  var areas = CONFIG.AREAS.concat(['Alle']);
  var bereichRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(areas)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 4, CONFIG.MAX_ROWS, 1).setDataValidation(bereichRule);

  sheet.getRange(2, 1, CONFIG.MAX_ROWS, 1).setNumberFormat('DD.MM.YYYY');

  protectRange(sheet, 'A1:F1');

  sheet.setColumnWidth(1, 120);
  sheet.setColumnWidth(2, 100);
  sheet.setColumnWidth(3, 100);
  sheet.setColumnWidth(4, 200);
  sheet.setColumnWidth(5, 350);

  sheet.setFrozenRows(1);

  sheet.getRange(1, 5).setNote(
    'Startzeit/Endzeit optional (Format HH:MM).\n' +
    'Ohne Zeitangabe = ganztägig gesperrt.\n' +
    'Nur Endzeit ohne Startzeit = nicht erlaubt.\n' +
    '"Alle" = gesamte Halle an diesem Tag/in diesem Zeitraum gesperrt.'
  );
}

function upgradeSperrungenSheet(ss) {
  var sheet = ss.getSheetByName(CONFIG.SHEET_SPERRUNGEN);
  if (!sheet) return;

  sheet.showColumns(6, 2);
  var lastRow = sheet.getLastRow();
  if (lastRow >= 1) {
    sheet.getRange(1, 6, 1, 2).clearContent();
    if (lastRow >= 2) {
      sheet.getRange(2, 6, lastRow - 1, 2).clearContent();
    }
  }
  sheet.getRange(1, 6).setValue('\u26A0\uFE0F Status').setFontWeight('bold').setBackground('#E8E8E8');
  protectRange(sheet, 'A1:F1');
}

// -------------------- Team-Blätter --------------------

function createTeamSheet(ss, teamName, sep) {
  var sheet = ss.insertSheet(teamName);

  var headers = ['Gegner', 'Datum', 'Startzeit', 'späteste Endzeit', 'Heim/Auswärts', 'Bereich', '\u26A0\uFE0F Status', 'Kommentar'];
  sheet.getRange(1, 1, 1, 8)
    .setValues([headers])
    .setFontWeight('bold')
    .setBackground('#E8E8E8');

  sheet.getRange(2, 2, CONFIG.MAX_ROWS, 1).setNumberFormat('DD.MM.YYYY');
  sheet.getRange(2, 3, CONFIG.MAX_ROWS, 1).setNumberFormat('HH:MM');
  sheet.getRange(2, 4, CONFIG.MAX_ROWS, 1).setNumberFormat('HH:MM');

  sheet.getRange(1, 4).setNote('Wird automatisch beim Eintragen der Startzeit berechnet.');

  var dateRule = SpreadsheetApp.newDataValidation()
    .requireDate()
    .setAllowInvalid(true)
    .setHelpText('Datum eingeben oder auswählen')
    .build();
  sheet.getRange(2, 2, CONFIG.MAX_ROWS, 1).setDataValidation(dateRule);

  var haRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Heim', 'Auswärts'])
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 5, CONFIG.MAX_ROWS, 1).setDataValidation(haRule);

  var areaRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(CONFIG.AREAS)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 6, CONFIG.MAX_ROWS, 1).setDataValidation(areaRule);

  sheet.hideColumns(4, 1);

  var heimRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$E2="Heim"')
    .setBackground('#C8E6C9')
    .setRanges([sheet.getRange('A2:H' + CONFIG.MAX_ROWS)])
    .build();

  var auswaertsRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$E2="Auswärts"')
    .setBackground('#BBDEFB')
    .setRanges([sheet.getRange('A2:H' + CONFIG.MAX_ROWS)])
    .build();

  var fehlerRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=REGEXMATCH($G2' + sep + ' "❌")')
    .setBackground('#FFCDD2')
    .setRanges([sheet.getRange('A2:H' + CONFIG.MAX_ROWS)])
    .build();

  var warnungRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=REGEXMATCH($G2' + sep + ' "⚠️")')
    .setBackground('#FFE0B2')
    .setRanges([sheet.getRange('A2:H' + CONFIG.MAX_ROWS)])
    .build();

  var rules = sheet.getConditionalFormatRules();
  rules.push(heimRule, auswaertsRule, fehlerRule, warnungRule);
  sheet.setConditionalFormatRules(rules);

  protectRange(sheet, 'A1:H1');

  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 110);
  sheet.setColumnWidth(3, 90);
  sheet.setColumnWidth(5, 120);
  sheet.setColumnWidth(6, 200);
  sheet.setColumnWidth(7, 350);
  sheet.setColumnWidth(8, 250);

  sheet.setFrozenRows(1);
}

function upgradeTeamSheet(sheet, sep) {
  sheet.showColumns(4);
  sheet.getRange(1, 4).setValue('späteste Endzeit').setFontWeight('bold').setBackground('#E8E8E8');
  sheet.hideColumns(4, 1);

  if (sheet.getRange('G1').getValue() !== '\u26A0\uFE0F Status') {
    sheet.getRange(1, 1, 1, 8).clearContent();
    sheet.getRange(1, 1, 1, 8)
      .setValues([['Gegner', 'Datum', 'Startzeit', 'späteste Endzeit', 'Heim/Auswärts', 'Bereich', '\u26A0\uFE0F Status', 'Kommentar']])
      .setFontWeight('bold')
      .setBackground('#E8E8E8');
  }

  protectRange(sheet, 'A1:H1');
}

// -------------------- Hallen/Spielplan-Blatt --------------------

function createHallenSpielplanSheet(ss, sep) {
  initHallenSpielplanSheet(ss.insertSheet(CONFIG.SHEET_PLAN), sep);
}

/**
 * Aktualisiert die Struktur des Hallen/Spielplan-Tabs, ohne ihn zu löschen.
 * Die GID (Tab-ID) bleibt erhalten — Voraussetzung für eine stabile,
 * veröffentlichte Web-URL. Ein Löschen+Neuanlegen würde die URL brechen.
 *
 * Wird von setupSheet(), resetAll() und createSheetsOnly() aufgerufen,
 * wenn der Tab bereits existiert.
 */
function upgradeHallenSpielplanSheet(ss, sep) {
  var sheet = ss.getSheetByName(CONFIG.SHEET_PLAN);
  if (!sheet) {
    createHallenSpielplanSheet(ss, sep);
    return;
  }
  initHallenSpielplanSheet(sheet, sep);
}

function initHallenSpielplanSheet(sheet, sep) {

  sheet.deleteRows(1, 2);

  var headers = ['Datum', 'Tag', 'Startzeit', 'Bereich', 'Team', 'H/A', 'Gegner', 'Kommentar', 'Status'];
  sheet.getRange(1, 1, 1, 9)
    .setValues([headers])
    .setFontWeight('bold')
    .setBackground('#E8E8E8');

  var heimRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$F2="Heim"')
    .setBackground('#C8E6C9')
    .setRanges([sheet.getRange('A2:I' + CONFIG.MAX_ROWS)])
    .build();

  var auswaertsRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$F2="Auswärts"')
    .setBackground('#BBDEFB')
    .setRanges([sheet.getRange('A2:I' + CONFIG.MAX_ROWS)])
    .build();

  var fehlerRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=REGEXMATCH($I2' + sep + ' "❌")')
    .setBackground('#FFCDD2')
    .setRanges([sheet.getRange('A2:I' + CONFIG.MAX_ROWS)])
    .build();

  var warnungRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=REGEXMATCH($I2' + sep + ' "⚠️")')
    .setBackground('#FFE0B2')
    .setRanges([sheet.getRange('A2:I' + CONFIG.MAX_ROWS)])
    .build();

  var sperrRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=REGEXMATCH($I2' + sep + ' "gesperrt")')
    .setBackground('#E0E0E0')
    .setRanges([sheet.getRange('A2:I' + CONFIG.MAX_ROWS)])
    .build();

  var rules = sheet.getConditionalFormatRules();
  rules.push(heimRule, auswaertsRule, fehlerRule, warnungRule, sperrRule);
  sheet.setConditionalFormatRules(rules);

  sheet.setColumnWidth(1, 110);
  sheet.setColumnWidth(2, 40);
  sheet.setColumnWidth(3, 50);
  sheet.setColumnWidth(4, 150);
  sheet.setColumnWidth(5, 140);
  sheet.setColumnWidth(6, 65);
  sheet.setColumnWidth(7, 200);
  sheet.setColumnWidth(8, 250);
  sheet.setColumnWidth(9, 300);

  sheet.setFrozenRows(1);
  protectRange(sheet, 'A1:I1');
}

function generatePlan() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var planSheet = ss.getSheetByName(CONFIG.SHEET_PLAN);
  if (!planSheet) return;

  var rows = [];
  var teamNames = readTeamNames(ss);

  for (var t = 0; t < teamNames.length; t++) {
    var teamSheet = ss.getSheetByName(teamNames[t]);
    if (!teamSheet) continue;
    var lastRow = teamSheet.getLastRow();
    if (lastRow < 2) continue;
    var data = teamSheet.getRange(2, 1, lastRow - 1, 8).getValues();
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      if (!isValidDate(row[1])) continue;
      rows.push([
        row[1],
        weekdayName(row[1].getDay()),
        row[2],
        row[5] || '',
        teamNames[t],
        row[4],
        row[0],
        row[7] || '',
        row[6] || ''
      ]);
    }
  }

  var sperrSheet = ss.getSheetByName(CONFIG.SHEET_SPERRUNGEN);
  if (sperrSheet) {
    var lastRow = sperrSheet.getLastRow();
    if (lastRow >= 2) {
      var data = sperrSheet.getRange(2, 1, lastRow - 1, 5).getValues();
      for (var i = 0; i < data.length; i++) {
        var row = data[i];
        if (!isValidDate(row[0])) continue;
        if (isValidTime(row[2]) && !isValidTime(row[1])) continue;
        var zeitraum = '';
        if (!isValidTime(row[1]) && !isValidTime(row[2])) {
          zeitraum = 'ganztägig';
        } else if (isValidTime(row[1]) && isValidTime(row[2])) {
          zeitraum = formatTime(row[1]) + ' - ' + formatTime(row[2]);
        } else if (isValidTime(row[1])) {
          zeitraum = 'ab ' + formatTime(row[1]);
        } else {
          zeitraum = 'bis ' + formatTime(row[2]);
        }
        rows.push([
          row[0],
          weekdayName(row[0].getDay()),
          row[1],
          row[3] || '',
          '',
          '',
          zeitraum,
          row[4] || '',
          'gesperrt'
        ]);
      }
    }
  }

  rows.sort(function(a, b) {
    var da = a[0] instanceof Date ? a[0].getTime() : 0;
    var db = b[0] instanceof Date ? b[0].getTime() : 0;
    if (da !== db) return da - db;
    var ta = timeToFraction(a[2]) || 0;
    var tb = timeToFraction(b[2]) || 0;
    return ta - tb;
  });

  var maxRow = Math.max(planSheet.getLastRow(), 1);
  if (maxRow >= 2) {
    planSheet.getRange(2, 1, maxRow - 1, 9).clearContent();
  }

  if (rows.length > 0) {
    planSheet.getRange(2, 1, rows.length, 9).setValues(rows);
    planSheet.getRange(2, 1, rows.length, 1).setNumberFormat('DD.MM.YYYY');
    planSheet.getRange(2, 3, rows.length, 1).setNumberFormat('HH:MM');
  } else {
    planSheet.getRange(2, 5).setValue('Keine Einträge vorhanden.');
  }
}

function weekdayName(day) {
  return ['So', 'Mo', 'Di', 'Mi', 'Do', 'Fr', 'Sa'][day];
}

function formatTime(t) {
  if (t instanceof Date) {
    return ('0' + t.getHours()).slice(-2) + ':' + ('0' + t.getMinutes()).slice(-2);
  }
  if (typeof t === 'number' && t >= 0 && t < 1) {
    var total = Math.round(t * 1440);
    var h = Math.floor(total / 60);
    var m = total % 60;
    return ('0' + h).slice(-2) + ':' + ('0' + m).slice(-2);
  }
  return '';
}

// ==================== VALIDIERUNG ====================

function handleEdit(e) {
  var sheet = e.range.getSheet();
  var name = sheet.getName();
  var teamNames = readTeamNames(SpreadsheetApp.getActiveSpreadsheet());

  if (teamNames.indexOf(name) >= 0) {
    if (e.range.getRow() < 2) return;
    var col = e.range.getColumn();

  if (col === 3) {
    var rawVal = e.range.getValue();
    if (typeof rawVal === 'number' && rawVal >= 1 && rawVal < 24) {
      var h = Math.floor(rawVal);
      var m = Math.round((rawVal - h) * 60);
      e.range.setValue((h * 60 + m) / 1440);
      e.range.setNumberFormat('HH:MM');
    }
    computeEndzeit(e.range);
  }

    var relevantCols = [1, 2, 3, 4, 5, 6, 8];
    if (relevantCols.indexOf(col) === -1) return;
    validateAllEntries();
    generatePlan();
  } else if (name === CONFIG.SHEET_SPERRUNGEN) {
    if (e.range.getRow() < 2) return;
    validateSperrungen();
    validateAllEntries();
    generatePlan();
  }
}

function computeEndzeit(startCell) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = startCell.getSheet();
  var row = startCell.getRow();
  var teamName = sheet.getName();
  var teams = readTeams(ss);
  var spieldauer = CONFIG.GAME_DURATION_HOURS;
  if (teams[teamName] && typeof teams[teamName].spieldauer === 'number' && teams[teamName].spieldauer > 0) {
    spieldauer = teams[teamName].spieldauer;
  }
  var startVal = startCell.getValue();
  var frac = timeToFraction(startVal);
  if (frac !== null) {
    sheet.getRange(row, 4).setValue(frac + spieldauer / 24).setNumberFormat('HH:MM');
  } else {
    sheet.getRange(row, 4).clearContent();
  }
}

function validateAllEntries() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var teamNames = readTeamNames(ss);
  var sperrungen = readSperrungen(ss);
  var teams = readTeams(ss);

  var data = [];
  var entrySheets = [];
  for (var t = 0; t < teamNames.length; t++) {
    var sheet = ss.getSheetByName(teamNames[t]);
    if (!sheet) continue;
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) continue;
    var teamData = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
    for (var i = 0; i < teamData.length; i++) {
      var r = teamData[i];
      data.push([
        teamNames[t], r[0], r[1], r[2], r[3], r[4], r[5], r[7], r[6]
      ]);
      entrySheets.push({ sheet: sheet, rowIndex: i });
    }
  }

  if (data.length === 0) return;

  var statusMessages = {};
  for (var i = 0; i < data.length; i++) {
    statusMessages[i] = [];
  }

  var bookingsByDateAndArea = {};
  var areaCountByDate = {};
  for (var i = 0; i < data.length; i++) {
    if (!isValidDate(data[i][2])) continue;
    if (data[i][5] !== 'Heim' || !data[i][6]) continue;
    var dk = datumToKey(data[i][2]);
    var ba = data[i][6];
    var key = dk + '|' + ba;
    if (!bookingsByDateAndArea[key]) bookingsByDateAndArea[key] = [];
    bookingsByDateAndArea[key].push(i);
    if (!areaCountByDate[dk]) areaCountByDate[dk] = {};
    areaCountByDate[dk][ba] = (areaCountByDate[dk][ba] || 0) + 1;
  }

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var team = row[0];
    var datum = row[2];
    var startzeit = row[3];
    var endzeit = row[4];
    var ha = row[5];
    var bereich = row[6];

    if (!team) {
      statusMessages[i].push('❌ Team fehlt');
    }
    if (!isValidDate(datum)) {
      statusMessages[i].push('❌ Datum fehlt');
      continue;
    }
    if (ha !== 'Heim' && ha !== 'Auswärts') {
      statusMessages[i].push('❌ H/A fehlt');
      continue;
    }

    var weekday = datum.getDay();
    var dateKey = datumToKey(datum);

    if (ha === 'Heim' && CONFIG.ALLOWED_WEEKDAYS.indexOf(weekday) === -1) {
      statusMessages[i].push('❌ Nur Di, Mi, Fr, Sa, So erlaubt');
    }

    if (ha === 'Heim' && !bereich) {
      statusMessages[i].push('❌ Bereich fehlt');
    }

    if (weekday === 3 && ha === 'Heim' && bereich && bereich !== CONFIG.WEDNESDAY_AREA) {
      statusMessages[i].push('❌ Mittwochs nur "' + CONFIG.WEDNESDAY_AREA + '" buchbar');
    }

    if ((weekday === 2 || weekday === 5) && ha === 'Heim' && bereich) {
      var areasOnDate = areaCountByDate[dateKey] || {};
      if (Object.keys(areasOnDate).length > CONFIG.MAX_AREAS_TUE_FRI) {
        statusMessages[i].push('❌ Di/Fr: max. ' + CONFIG.MAX_AREAS_TUE_FRI +
          ' Bereiche buchbar (' + Object.keys(areasOnDate).length + ' wären belegt)');
      }
    }

    if (ha === 'Heim' && bereich) {
      var key = dateKey + '|' + bereich;
      var sameSlot = bookingsByDateAndArea[key] || [];
      var others = sameSlot.filter(function(j) { return j !== i; });
      if (others.length > 0) {
        if (weekday === 6 || weekday === 0) {
          var s1 = timeToFraction(startzeit), e1 = timeToFraction(endzeit);
          var hasOverlap = false;
          for (var k = 0; k < others.length; k++) {
            var s2 = timeToFraction(data[others[k]][3]);
            var e2 = timeToFraction(data[others[k]][4]);
            if (s1 !== null && e1 !== null && s2 !== null && e2 !== null) {
              if (s1 < e2 && e1 > s2) {
                statusMessages[i].push('❌ Zeitüberlappung in "' + bereich + '"');
                hasOverlap = true;
                break;
              }
            } else {
              if (!hasOverlap) {
                statusMessages[i].push('❌ "' + bereich + '" bereits belegt (Startzeit fehlt)');
                hasOverlap = true;
              }
            }
          }
        } else {
          statusMessages[i].push('❌ "' + bereich + '" an diesem Tag bereits belegt');
        }
      }
    }

    if (ha === 'Heim' && bereich) {
      for (var s = 0; s < sperrungen.length; s++) {
        var sDatum = sperrungen[s].datum;
        if (!isValidDate(sDatum) || datumToKey(sDatum) !== dateKey) continue;
        if (sperrungen[s].bereich !== 'Alle' && sperrungen[s].bereich !== bereich) continue;

        var sStart = sperrungen[s].startzeit;
        var sEnd = sperrungen[s].endzeit;

        if (isValidTime(sStart) && isValidTime(sEnd)) {
          if (isValidTime(startzeit) && isValidTime(endzeit)) {
            var bs = timeToFraction(startzeit), be = timeToFraction(endzeit);
            var ss = timeToFraction(sStart), se = timeToFraction(sEnd);
            if (bs !== null && be !== null && ss !== null && se !== null) {
              if (bs < se && be > ss) {
                var msg = '❌ Überschneidung mit Sperrung';
                if (sperrungen[s].kommentar) msg += ': ' + sperrungen[s].kommentar;
                statusMessages[i].push(msg);
              }
            }
          }
        } else if (isValidTime(sStart)) {
          if (isValidTime(startzeit) && isValidTime(endzeit)) {
            var bs = timeToFraction(startzeit), be = timeToFraction(endzeit);
            var ss = timeToFraction(sStart), se = 1;
            if (bs !== null && be !== null && ss !== null) {
              if (bs < se && be > ss) {
                var msg = '❌ Überschneidung mit Sperrung (ab ' + formatTime(sStart) + ')';
                if (sperrungen[s].kommentar) msg += ': ' + sperrungen[s].kommentar;
                statusMessages[i].push(msg);
              }
            }
          } else {
            var msg = '❌ ';
            if (sperrungen[s].bereich === 'Alle') {
              msg += 'Tag ab ' + formatTime(sStart) + ' gesperrt';
            } else {
              msg += '"' + sperrungen[s].bereich + '" ab ' + formatTime(sStart) + ' gesperrt';
            }
            if (sperrungen[s].kommentar) msg += ': ' + sperrungen[s].kommentar;
            statusMessages[i].push(msg);
          }
        } else {
          var msg = '❌ ';
          if (sperrungen[s].bereich === 'Alle') {
            msg += 'Tag gesperrt';
          } else {
            msg += '"' + sperrungen[s].bereich + '" gesperrt';
          }
          if (sperrungen[s].kommentar) msg += ': ' + sperrungen[s].kommentar;
          statusMessages[i].push(msg);
        }
      }
    }

    if (ha === 'Heim' && isValidTime(startzeit) && isValidTime(endzeit)) {
      var st = timeToFraction(startzeit);
      var et = timeToFraction(endzeit);
      if (st !== null && et !== null && et <= st) {
        statusMessages[i].push('❌ Endzeit muss nach Startzeit liegen');
      }
    }
  }

  var adjacentMessages = checkAdjacentTeams(data, teams);

  var perSheet = {};
  for (var i = 0; i < data.length; i++) {
    var es = entrySheets[i];
    var name = es.sheet.getName();
    if (!perSheet[name]) perSheet[name] = {};
    var allMessages = statusMessages[i].slice();
    if (adjacentMessages[i]) {
      allMessages = allMessages.concat(adjacentMessages[i]);
    }
    perSheet[name][es.rowIndex] = allMessages.join(' | ');
  }

  for (var key in perSheet) {
    var s = ss.getSheetByName(key);
    if (!s) continue;
    var rows = perSheet[key];
    var max = 0;
    for (var ri in rows) {
      max = Math.max(max, parseInt(ri) + 1);
    }
    var vals = [];
    for (var vi = 0; vi < max; vi++) {
      vals.push([rows[vi] || '']);
    }
    if (vals.length > 0) {
      s.getRange(2, 7, vals.length, 1).setValues(vals);
    }
  }
}

function validateSperrungen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_SPERRUNGEN);
  if (!sheet) return;

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  var data = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
  var statusValues = [];
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var messages = [];
    if (!isValidDate(row[0]) && row[0]) {
      messages.push('❌ Ungültiges Datum');
    }
    if (isValidTime(row[2]) && !isValidTime(row[1])) {
      messages.push('❌ Endzeit ohne Startzeit');
    }
    statusValues.push([messages.join(' | ')]);
  }
  if (statusValues.length > 0) {
    sheet.getRange(2, 6, statusValues.length, 1).setValues(statusValues);
  }
}

// -------------------- Hilfsfunktionen Validierung --------------------

function checkAdjacentTeams(data, teams) {
  var messages = {};

  var entriesByDate = {};
  for (var i = 0; i < data.length; i++) {
    if (!isValidDate(data[i][2]) || !data[i][0]) continue;
    var key = datumToKey(data[i][2]);
    if (!entriesByDate[key]) entriesByDate[key] = [];
    entriesByDate[key].push({ index: i, teamName: data[i][0] });
  }

  for (var dateKey in entriesByDate) {
    var entries = entriesByDate[dateKey];

    var byGroup = {};
    for (var e = 0; e < entries.length; e++) {
      var teamName = entries[e].teamName;
      var teamInfo = teams[teamName];
      if (!teamInfo) continue;
      var g = teamInfo.gruppe;
      if (!byGroup[g]) byGroup[g] = [];
      byGroup[g].push({ index: entries[e].index, rang: teamInfo.rang, name: teamName });
    }

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

function readTeamNames(ss) {
  var setupSheet = ss.getSheetByName(CONFIG.SHEET_SETUP);
  if (!setupSheet) return [];
  var lastRow = setupSheet.getLastRow();
  if (lastRow < 2) return [];
  var data = setupSheet.getRange(2, 3, lastRow - 1, 1).getValues();
  var names = [];
  for (var i = 0; i < data.length; i++) {
    if (data[i][0]) names.push(String(data[i][0]));
  }
  return names;
}

function readTeams(ss) {
  var sheet = ss.getSheetByName(CONFIG.SHEET_SETUP);
  if (!sheet) return {};

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return {};

  var raw = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
  var defaultDuration = sheet.getRange('F1').getValue();
  if (typeof defaultDuration !== 'number' || defaultDuration <= 0) defaultDuration = CONFIG.GAME_DURATION_HOURS;

  var teams = {};
  for (var i = 0; i < raw.length; i++) {
    if (raw[i][0] && raw[i][1] && raw[i][2]) {
      var spieldauer = raw[i][7]; // Spalte H
      if (typeof spieldauer !== 'number' || spieldauer <= 0) spieldauer = defaultDuration;
      teams[raw[i][2]] = { rang: raw[i][0], gruppe: raw[i][1], spieldauer: spieldauer };
    }
  }
  return teams;
}

function readSperrungen(ss) {
  var sheet = ss.getSheetByName(CONFIG.SHEET_SPERRUNGEN);
  if (!sheet) return [];

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  var raw = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
  var sperrungen = [];
  for (var i = 0; i < raw.length; i++) {
    if (raw[i][0]) {
      sperrungen.push({
        datum: raw[i][0],
        startzeit: raw[i][1],
        endzeit: raw[i][2],
        bereich: raw[i][3],
        kommentar: raw[i][4]
      });
    }
  }
  return sperrungen;
}

// -------------------- Allgemeine Hilfsfunktionen --------------------

function isValidDate(d) {
  return d instanceof Date && !isNaN(d.getTime());
}

function isValidTime(t) {
  if (typeof t === 'number' && t >= 0 && t < 1) return true;
  if (t instanceof Date && !isNaN(t.getTime())) {
    return true;
  }
  return false;
}

function timeToFraction(t) {
  if (t instanceof Date && !isNaN(t.getTime())) {
    return (t.getHours() * 3600 + t.getMinutes() * 60 + t.getSeconds()) / 86400;
  }
  if (typeof t === 'number' && t >= 0 && t < 1) {
    return t;
  }
  return null;
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
  protection.removeEditors(protection.getEditors());
  if (protection.canDomainEdit()) protection.setDomainEdit(false);
}

function createTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'handleEdit') {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger('handleEdit')
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onEdit()
    .create();
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Hallenplanung')
    .addItem('Plan aktualisieren', 'generatePlan')
    .addToUi();
}

// ==================== TSV-EXPORT ====================

function downloadTSV() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var result = '';

  // Setup Teams
  result += '# data/setup_teams.tsv\n';
  result += 'Rang\tGruppe\tTeamname\tKurzname\tSpieldauer\n';
  var setupSheet = ss.getSheetByName(CONFIG.SHEET_SETUP);
  if (setupSheet) {
    var lastRow = Math.max(setupSheet.getLastRow(), 1);
    if (lastRow >= 2) {
      var data = setupSheet.getRange(2, 1, lastRow - 1, 8).getValues();
      for (var i = 0; i < data.length; i++) {
        if (data[i][0]) {
          result += [data[i][0], data[i][1], data[i][2], data[i][3], data[i][7] || ''].join('\t') + '\n';
        }
      }
    }
  }
  result += '\n';

  // Eingabe
  result += '# data/eingabe.tsv\n';
  result += 'Team\tGegner\tDatum\tStartzeit\tspäteste Endzeit\tHeim/Auswärts\tBereich\tKommentar\n';
  var teamNames = readTeamNames(ss);
  for (var t = 0; t < teamNames.length; t++) {
    var sheet = ss.getSheetByName(teamNames[t]);
    if (!sheet) continue;
    var lastRow2 = Math.max(sheet.getLastRow(), 1);
    if (lastRow2 >= 2) {
      var data2 = sheet.getRange(2, 1, lastRow2 - 1, 8).getValues();
      for (var j = 0; j < data2.length; j++) {
        if (data2[j][1]) {
          var d = data2[j][1];
          var dateStr = d instanceof Date ? ('0' + d.getDate()).slice(-2) + '.' + ('0' + (d.getMonth() + 1)).slice(-2) + '.' + d.getFullYear() : d;
          var st = _formatTime(data2[j][2]);
          var et = _formatTime(data2[j][3]);
          result += [teamNames[t], data2[j][0], dateStr, st, et, data2[j][4], data2[j][5], data2[j][7] || ''].join('\t') + '\n';
        }
      }
    }
  }
  result += '\n';

  // Sperrungen
  result += '# data/sperrungen.tsv\n';
  result += 'Datum\tStartzeit\tEndzeit\tBereich\tKommentar\n';
  var sperrSheet = ss.getSheetByName(CONFIG.SHEET_SPERRUNGEN);
  if (sperrSheet) {
    var lastRow3 = Math.max(sperrSheet.getLastRow(), 1);
    if (lastRow3 >= 2) {
      var data3 = sperrSheet.getRange(2, 1, lastRow3 - 1, 5).getValues();
      for (var k = 0; k < data3.length; k++) {
        if (data3[k][0]) {
          var ds = data3[k][0];
          var dateStr = ds instanceof Date ? ('0' + ds.getDate()).slice(-2) + '.' + ('0' + (ds.getMonth() + 1)).slice(-2) + '.' + ds.getFullYear() : ds;
          result += [dateStr, _formatTime(data3[k][1]), _formatTime(data3[k][2]), data3[k][3], data3[k][4]].join('\t') + '\n';
        }
      }
    }
  }

  Logger.log(result);
}

function _formatTime(val) {
  if (typeof val === 'number' && val >= 0 && val < 1) {
    var totalMin = Math.round(val * 1440);
    var h = Math.floor(totalMin / 60);
    var m = totalMin % 60;
    return ('0' + h).slice(-2) + ':' + ('0' + m).slice(-2);
  }
  if (val instanceof Date) {
    return ('0' + val.getHours()).slice(-2) + ':' + ('0' + val.getMinutes()).slice(-2);
  }
  return val || '';
}
