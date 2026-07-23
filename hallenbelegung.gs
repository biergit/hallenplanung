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
 *  Teams                             – Rang, Gruppe, generierter Teamname
 *  Sperrungen/Anderweitige Belegungen – Gesperrte Tage, Bereiche und Zeiträume
 *  Eingabe                           – Dateneingabe durch Mannschaftsführer
 *  Belegungsplan                     – Öffentliche Kalenderansicht
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
  SHEET_EINGABE: 'Eingabe',
  SHEET_PLAN: 'Hallen/Spielplan',

  // Standard-Spieldauer in Stunden (konfigurierbar in Setup!F1)
  GAME_DURATION_HOURS: 4
};
// --- SEED DATA BEGIN ---
var SEED_SETUP = JSON.parse('[]');
var SEED_EINGABE = JSON.parse('[]');
var SEED_SPERRUNGEN = JSON.parse('[]');
// --- SEED DATA END ---

// ==================== SETUP ====================

function setupSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var locale = ss.getSpreadsheetLocale();
  var isGerman = locale.startsWith('de');
  var sep = isGerman ? ';' : ',';
  var arrSep = isGerman ? '\\' : ',';

  // Plan sheet immer neu (auch alten Reiter-Namen aufräumen)
  var oldPlan = ss.getSheetByName('Belegungsplan');
  if (oldPlan) ss.deleteSheet(oldPlan);
  var planSheet = ss.getSheetByName(CONFIG.SHEET_PLAN);
  if (planSheet) ss.deleteSheet(planSheet);
  createBelegungsplanSheet(ss, sep, arrSep);

  // Setup: nur anlegen, wenn nicht vorhanden
  if (!ss.getSheetByName(CONFIG.SHEET_SETUP)) {
    createSetupSheet(ss, sep);
  }

  // Sperrungen: upgraden wenn vorhanden, sonst neu
  var sperrSheet = ss.getSheetByName(CONFIG.SHEET_SPERRUNGEN);
  if (sperrSheet) {
    upgradeSperrungenSheet(ss, sep);
  } else {
    createSperrungenSheet(ss, sep);
  }

  // Eingabe: nur anlegen, wenn nicht vorhanden
  if (!ss.getSheetByName(CONFIG.SHEET_EINGABE)) {
    createEingabeSheet(ss, sep);
  }

  createTrigger();
  seedSheets(ss);

  var msg = 'Setup abgeschlossen!\n\n' +
    'Blätter:\n' +
    '  1. ' + CONFIG.SHEET_SETUP + ' – Team-Konfiguration und Einstellungen\n' +
    '  2. ' + CONFIG.SHEET_SPERRUNGEN + ' – Gesperrte Tage, Bereiche, Zeiträume\n' +
    '  3. ' + CONFIG.SHEET_EINGABE + ' – Dateneingabe für Mannschaftsführer\n' +
    '  4. ' + CONFIG.SHEET_PLAN + ' – Kalenderansicht\n\n' +
    'Die Validierung läuft automatisch bei jeder Eingabe.\n\n' +
    'Für die Web-Veröffentlichung:\n' +
    '  Datei → Für das Web veröffentlichen → ' + CONFIG.SHEET_PLAN;
  try {
    SpreadsheetApp.getUi().alert(msg);
  } catch (e) {
    Logger.log(msg);
  }
}

// -------------------- Seed / Reset --------------------

function resetAll() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var locale = ss.getSpreadsheetLocale();
  var isGerman = locale.startsWith('de');
  var sep = isGerman ? ';' : ',';
  var arrSep = isGerman ? '\\' : ',';

  var sheetNames = [CONFIG.SHEET_PLAN, CONFIG.SHEET_EINGABE, CONFIG.SHEET_SPERRUNGEN, CONFIG.SHEET_SETUP];
  sheetNames.forEach(function(name) {
    var sheet = ss.getSheetByName(name);
    if (sheet) ss.deleteSheet(sheet);
  });
  // Alten Reiter-Namen aufräumen
  var oldPlan = ss.getSheetByName('Belegungsplan');
  if (oldPlan) ss.deleteSheet(oldPlan);

  createSetupSheet(ss, sep);
  createSperrungenSheet(ss, sep);
  createEingabeSheet(ss, sep);
  createBelegungsplanSheet(ss, sep, arrSep);
  createTrigger();
  seedSheets(ss);

  var msg = 'Alles neu angelegt inkl. Seed-Daten.\n\n' +
    '  1. ' + CONFIG.SHEET_SETUP + '\n' +
    '  2. ' + CONFIG.SHEET_SPERRUNGEN + '\n' +
    '  3. ' + CONFIG.SHEET_EINGABE + '\n' +
    '  4. ' + CONFIG.SHEET_PLAN;
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
  var arrSep = isGerman ? '\\' : ',';

  var sheetNames = [CONFIG.SHEET_PLAN, CONFIG.SHEET_EINGABE, CONFIG.SHEET_SPERRUNGEN, CONFIG.SHEET_SETUP];
  sheetNames.forEach(function(name) {
    var sheet = ss.getSheetByName(name);
    if (sheet) ss.deleteSheet(sheet);
  });
  var oldPlan = ss.getSheetByName('Belegungsplan');
  if (oldPlan) ss.deleteSheet(oldPlan);

  createSetupSheet(ss, sep);
  createSperrungenSheet(ss, sep);
  createEingabeSheet(ss, sep);
  createBelegungsplanSheet(ss, sep, arrSep);
  createTrigger();

  var msg = 'Blätter neu angelegt (ohne Seed-Daten).';
  try {
    SpreadsheetApp.getUi().alert(msg);
  } catch (e) {
    Logger.log(msg);
  }
}

function seedSheets(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  if (SEED_SETUP.length > 0) _seedSetup(ss);
  if (SEED_EINGABE.length > 0) _seedEingabe(ss);
  if (SEED_SPERRUNGEN.length > 0) _seedSperrungen(ss);
}

function _seedSetup(ss) {
  var sheet = ss.getSheetByName(CONFIG.SHEET_SETUP);
  if (!sheet) return;
  if (sheet.getLastRow() >= 2) return; // Bereits Daten vorhanden

  var data = [];
  for (var i = 0; i < SEED_SETUP.length; i++) {
    data.push([parseInt(SEED_SETUP[i][0]), SEED_SETUP[i][1]]);
  }
  if (data.length > 0) {
    sheet.getRange(2, 1, data.length, 2).setValues(data);
  }
}

function _seedEingabe(ss) {
  var sheet = ss.getSheetByName(CONFIG.SHEET_EINGABE);
  if (!sheet) return;
  if (sheet.getLastRow() >= 2) return;

  var data = [];
  for (var i = 0; i < SEED_EINGABE.length; i++) {
    var row = SEED_EINGABE[i];
    var datum = _parseDate(row[0]);
    var startzeit = _parseTime(row[1]);
    var endzeit = _parseTime(row[2]);
    data.push([datum, startzeit, endzeit, row[3], row[4], row[5], row[6], row[7], '']);
  }
  if (data.length > 0) {
    sheet.getRange(2, 1, data.length, 9).setValues(data);
  }
}

function _seedSperrungen(ss) {
  var sheet = ss.getSheetByName(CONFIG.SHEET_SPERRUNGEN);
  if (!sheet) return;
  if (sheet.getLastRow() >= 2) return;

  var data = [];
  for (var i = 0; i < SEED_SPERRUNGEN.length; i++) {
    var row = SEED_SPERRUNGEN[i];
    var datum = _parseDate(row[0]);
    var startzeit = _parseTime(row[1]);
    var endzeit = _parseTime(row[2]);
    data.push([datum, startzeit, endzeit, row[3], row[4]]);
  }
  if (data.length > 0) {
    sheet.getRange(2, 1, data.length, 5).setValues(data);
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
  var sheet = ss.insertSheet(CONFIG.SHEET_SETUP, 0);

  var headers = ['Rang', 'Gruppe', 'Teamname', 'Kurzname'];
  sheet.getRange(1, 1, 1, 4)
    .setValues([headers])
    .setFontWeight('bold')
    .setBackground('#E8E8E8');

  // Spieldauer-Konfiguration in E1/F1
  sheet.getRange('E1').setValue('Spieldauer (h):');
  sheet.getRange('E1').setFontWeight('bold');
  sheet.getRange('F1').setValue(CONFIG.GAME_DURATION_HOURS);
  sheet.getRange('F1').setNumberFormat('0.0');
  sheet.getRange('F1').setNote('Standard-Spieldauer in Stunden. Wird für die Berechnung der spätesten Endzeit verwendet.');

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
  sheet.getRange(2, 2, 1000, 1).setDataValidation(gruppeRule);

  var rangRule = SpreadsheetApp.newDataValidation()
    .requireNumberGreaterThan(0)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 1, 1000, 1).setDataValidation(rangRule);

  protectRange(sheet, 'A1:D1');

  sheet.setColumnWidth(1, 60);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(3, 200);
  sheet.setColumnWidth(4, 120);
  sheet.setColumnWidth(7, 150);

  sheet.hideColumns(7);

  // Bedienungsanleitung in Spalten J-L
  sheet.getRange('J1').setValue('📋 BEDIENUNGSANLEITUNG');
  sheet.getRange('J1').setFontWeight('bold');
  sheet.getRange('J1').setFontSize(12);
  sheet.setColumnWidth(10, 500);

  var instructions = [
    ['SPIELDAUER ÄNDERN:', 'Wert in Zelle F1 ändern (z.B. 4.5 für 4h 30min).'],
    ['', 'Gilt nur für neue Einträge – bestehende Endzeiten bleiben unverändert.'],
    ['', ''],
    ['STARTZEIT EINGEBEN:', 'Einfach "12" statt "12:00" eingeben – das Script wandelt es um.'],
    ['', '"12.5" = 12:30, "14.25" = 14:15, usw.'],
    ['', ''],
    ['TEAMS:', 'Rang und Gruppe in Spalten A+B eintragen.'],
    ['', 'Teamname und Kurzname werden automatisch generiert.'],
    ['', 'Neue Gruppen: in Spalte G ergänzen (ausgeblendet).'],
    ['', ''],
    ['EINGABE:', 'Mannschaftsführer tragen Heim-/Auswärtsspiele ein.'],
    ['', 'Heimspiel: Bereich auswählen (Pflicht).'],
    ['', 'Auswärtsspiel: Bereich frei lassen.'],
    ['', 'Fehler und Warnungen erscheinen in der Status-Spalte.'],
    ['', ''],
    ['SPERRUNGEN:', 'Gesperrte Tage/Bereiche/Zeiträume eintragen.'],
    ['', 'Ohne Start-/Endzeit: ganztägig gesperrt.'],
    ['', 'Mit Start-/Endzeit: nur dieser Zeitraum gesperrt.'],
    ['', ''],
    ['HALLEN/SPIELPLAN:', 'Kalenderansicht für die Web-Veröffentlichung.'],
    ['', 'Checkbox A1: Nur Hallenbelegung (Heimspiele).'],
    ['', 'Checkbox A2: Sperrungen/Anderw. Belegungen einblenden.'],
    ['', ''],
    ['WEB-VERÖFFENTLICHUNG:', 'Datei → Freigeben → Für das Web veröffentlichen'],
    ['', 'Blatt "Hallen/Spielplan" auswählen.'],
    ['', 'Checkboxen VOR dem Veröffentlichen nach Wunsch setzen.'],
    ['', ''],
    ['FREIGABE:', 'Datei → Freigeben → E-Mail der Mannschaftsführer hinzufügen'],
    ['', 'Berechtigung: "Bearbeiter".'],
    ['', 'Andere Mitglieder: Link aus Web-Veröffentlichung.'],
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

function createSperrungenSheet(ss, sep) {
  var sheet = ss.insertSheet(CONFIG.SHEET_SPERRUNGEN, 1);

  var headers = ['Datum', 'Startzeit', 'Endzeit', 'Bereich', 'Kommentar', 'Wochentag', 'Zeitraum Anzeige'];
  sheet.getRange(1, 1, 1, 7)
    .setValues([headers])
    .setFontWeight('bold')
    .setBackground('#E8E8E8');

  var dateRule = SpreadsheetApp.newDataValidation()
    .requireDate()
    .setAllowInvalid(true)
    .setHelpText('Datum eingeben oder auswählen')
    .build();
  sheet.getRange(2, 1, 1000, 1).setDataValidation(dateRule);

  sheet.getRange(2, 2, 1000, 1).setNumberFormat('HH:MM');
  sheet.getRange(2, 3, 1000, 1).setNumberFormat('HH:MM');

  var areas = CONFIG.AREAS.concat(['Alle']);
  var bereichRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(areas)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 4, 1000, 1).setDataValidation(bereichRule);

  sheet.getRange(2, 1, 1000, 1).setNumberFormat('DD.MM.YYYY');

  // F: Wochentag (Hilfsspalte für QUERY im Hallen/Spielplan)
  sheet.getRange(2, 6, 1000, 1).setFormula('=IF(A2="";;TEXT(A2' + sep + '"ddd"))');

  // G: Zeitraum-Anzeige (Hilfsspalte für QUERY im Hallen/Spielplan)
  sheet.getRange(2, 7, 1000, 1).setFormula(
    '=IF(A2="";;IF(B2="";"ganztägig";TEXT(B2' + sep + '"HH:MM")&" - "&TEXT(C2' + sep + '"HH:MM")))');

  sheet.hideColumns(6, 2);

  protectRange(sheet, 'A1:G1');

  sheet.setColumnWidth(1, 120);
  sheet.setColumnWidth(2, 100);
  sheet.setColumnWidth(3, 100);
  sheet.setColumnWidth(4, 200);
  sheet.setColumnWidth(5, 350);

  sheet.setFrozenRows(1);

  sheet.getRange(1, 5).setNote(
    'Startzeit/Endzeit optional (Format HH:MM).\n' +
    'Ohne Zeitangabe = ganztägig gesperrt.\n' +
    '"Alle" = gesamte Halle an diesem Tag/in diesem Zeitraum gesperrt.'
  );
}

function upgradeSperrungenSheet(ss, sep) {
  var sheet = ss.getSheetByName(CONFIG.SHEET_SPERRUNGEN);
  if (!sheet) return;

  var lastRow = sheet.getLastRow();
  var maxRow = Math.max(lastRow, 2);

  // Spalte F: Wochentag (Hilfsspalte für QUERY)
  sheet.getRange(1, 6).setValue('Wochentag').setFontWeight('bold').setBackground('#E8E8E8');
  sheet.getRange(2, 6, maxRow - 1, 1).setFormula('=IF(A2="";;TEXT(A2' + sep + '"ddd"))');

  // Spalte G: Zeitraum-Anzeige (Hilfsspalte für QUERY)
  sheet.getRange(1, 7).setValue('Zeitraum Anzeige').setFontWeight('bold').setBackground('#E8E8E8');
  sheet.getRange(2, 7, maxRow - 1, 1).setFormula(
    '=IF(A2="";;IF(B2="";"ganztägig";TEXT(B2' + sep + '"HH:MM")&" - "&TEXT(C2' + sep + '"HH:MM")))');

  sheet.hideColumns(6, 2);
  protectRange(sheet, 'A1:G1');
}

// -------------------- Eingabe-Blatt --------------------

function createEingabeSheet(ss, sep) {
  var sheet = ss.insertSheet(CONFIG.SHEET_EINGABE, 2);

  // A=Datum, B=Startzeit, C=späteste Endzeit (Formel), D=Team, E=Heim/Auswärts,
  // F=Gegner, G=Bereich, H=Kommentar, I=Status
  var headers = ['Datum', 'Startzeit', 'späteste Endzeit', 'Team', 'Heim/Auswärts', 'Gegner', 'Bereich', 'Kommentar', '\u26A0\uFE0F Status'];
  sheet.getRange(1, 1, 1, 9)
    .setValues([headers])
    .setFontWeight('bold')
    .setBackground('#E8E8E8');

  sheet.getRange(2, 1, 1000, 1).setNumberFormat('DD.MM.YYYY');
  sheet.getRange(2, 2, 1000, 1).setNumberFormat('HH:MM');
  sheet.getRange(2, 3, 1000, 1).setNumberFormat('HH:MM');

  // Spalte C: späteste Endzeit (wird per Script beim Eintrag der Startzeit berechnet)
  sheet.getRange(1, 3).setNote('Wird automatisch beim Eintragen der Startzeit berechnet (Startzeit + Spieldauer aus Setup).');

  // Date picker
  var dateRule = SpreadsheetApp.newDataValidation()
    .requireDate()
    .setAllowInvalid(true)
    .setHelpText('Datum eingeben oder auswählen (Kalender)')
    .build();
  sheet.getRange(2, 1, 1000, 1).setDataValidation(dateRule);

  // Team-Dropdown aus Setup-Blatt
  var setupSheet = ss.getSheetByName(CONFIG.SHEET_SETUP);
  var teamRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(setupSheet.getRange('C2:C1000'))
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 4, 1000, 1).setDataValidation(teamRule);

  // Heim/Auswärts
  var haRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Heim', 'Auswärts'])
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 5, 1000, 1).setDataValidation(haRule);

  // Bereich
  var areaRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(CONFIG.AREAS)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 7, 1000, 1).setDataValidation(areaRule);

  // Bedingte Formatierungen
  var heimRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$E2="Heim"')
    .setBackground('#C8E6C9')
    .setRanges([sheet.getRange('A2:I1000')])
    .build();

  var auswaertsRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$E2="Auswärts"')
    .setBackground('#BBDEFB')
    .setRanges([sheet.getRange('A2:I1000')])
    .build();

  var fehlerRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=REGEXMATCH($I2' + sep + ' "❌")')
    .setBackground('#FFCDD2')
    .setRanges([sheet.getRange('A2:I1000')])
    .build();

  var warnungRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=REGEXMATCH($I2' + sep + ' "⚠️")')
    .setBackground('#FFE0B2')
    .setRanges([sheet.getRange('A2:I1000')])
    .build();

  var rules = sheet.getConditionalFormatRules();
  rules.push(heimRule, auswaertsRule, fehlerRule, warnungRule);
  sheet.setConditionalFormatRules(rules);

  protectRange(sheet, 'A1:I1');

  sheet.setColumnWidth(1, 110);
  sheet.setColumnWidth(2, 90);
  sheet.setColumnWidth(3, 110);
  sheet.setColumnWidth(4, 180);
  sheet.setColumnWidth(5, 120);
  sheet.setColumnWidth(6, 200);
  sheet.setColumnWidth(7, 200);
  sheet.setColumnWidth(8, 250);
  sheet.setColumnWidth(9, 350);

  sheet.setFrozenRows(1);

  sheet.getRange(1, 9).setNote(
    'Automatisch vom Script befüllt.\n' +
    '❌ = Validierungsfehler\n' +
    '⚠️ = Benachbartes Team spielt am selben Tag'
  );

  sheet.getRange(1, 10).setValue(
    'Hinweise:\n' +
    '- Nur Di/Mi/Fr/Sa/So\n' +
    '- Mi: nur "Kleine Halle"\n' +
    '- Di+Fr: max. 2 Bereiche\n' +
    '- Heimspiel: Bereich nötig\n' +
    '- Auswärtsspiel: Bereich leer'
  );
  sheet.getRange(1, 10).setFontSize(9);
  sheet.getRange(1, 10).setFontColor('#999999');
  sheet.setColumnWidth(10, 220);
}

// -------------------- Belegungsplan-Blatt --------------------

function createBelegungsplanSheet(ss, sep, arrSep) {
  var sheet = ss.insertSheet(CONFIG.SHEET_PLAN, 3);
  var sperrName = CONFIG.SHEET_SPERRUNGEN;

  // Checkbox 1 in A1: Nur Hallenbelegung
  sheet.getRange('A1').insertCheckboxes();
  sheet.getRange('A1').setValue(false);
  sheet.getRange('B1').setValue('Nur Hallenbelegung (Heimspiele)');
  sheet.getRange('B1').setFontStyle('italic');
  sheet.getRange('B1').setFontColor('#555555');

  // Checkbox 2 in A2: Sperrungen anzeigen
  sheet.getRange('A2').insertCheckboxes();
  sheet.getRange('A2').setValue(false);
  sheet.getRange('B2').setValue('Sperrungen/Anderweitige Belegungen anzeigen');
  sheet.getRange('B2').setFontStyle('italic');
  sheet.getRange('B2').setFontColor('#555555');

  // Überschriften Zeile 3
  var headers = ['Datum', 'Tag', 'Startzeit', 'Bereich', 'Team', 'H/A', 'Gegner', 'Kommentar', 'Status'];
  sheet.getRange(3, 1, 1, 9)
    .setValues([headers])
    .setFontWeight('bold')
    .setBackground('#E8E8E8');

  // Formel: =SORT({IF(A1; QH; QA); IF(A2; QS; dummy)}; 1; TRUE; 3; TRUE)
  // QH/QA: QUERY aus Eingabe, QS: COUNTA-gesicherte Sperrungen-QUERY
  // dummy: 9 leere Werte, sortiert ans Ende

  var qBase = 'QUERY({Eingabe!A2:I' + arrSep +
    ' ARRAYFORMULA(TEXT(Eingabe!A2:A' + sep + '"ddd"))}' + sep;

  var sqlHeim = '"SELECT Col1, Col10, Col2, Col7, Col4, Col5, Col6, Col8, Col9 ' +
    'WHERE Col1 IS NOT NULL AND Col5=\'Heim\'"';
  var sqlAll = '"SELECT Col1, Col10, Col2, Col7, Col4, Col5, Col6, Col8, Col9 ' +
    'WHERE Col1 IS NOT NULL"';

  var QH = qBase + sqlHeim + sep + ' 0)';
  var QA = qBase + sqlAll + sep + ' 0)';

  var dummyRow = '{""' + arrSep + '""' + arrSep + '""' + arrSep + '""' + arrSep + '""' + arrSep + '""' + arrSep + '""' + arrSep + '""' + arrSep + '""}';

  var QS = 'IF(COUNTA(\'' + sperrName + '\'!A2:A)=0' + sep + ' ' + dummyRow + sep +
    ' QUERY(\'' + sperrName + '\'!A2:G' + sep +
    '"SELECT Col1, Col6, Col2, Col4, \' \', \' \', Col7, Col5, \' \' ' +
    'WHERE Col1 IS NOT NULL AND Col4 IS NOT NULL"' + sep + ' 0))';

  var formula = '=QUERY(SORT({' +
    'IF(A1' + sep + ' ' + QH + sep + ' ' + QA + ')' + sep +
    ' IF(A2' + sep + ' ' + QS + sep + ' ' + dummyRow + ')' +
    '}' + sep + ' 1' + sep + ' TRUE' + sep + ' 3' + sep + ' TRUE)' + sep +
    '"SELECT * WHERE Col1 IS NOT NULL"' + sep + ' 0)';

  sheet.getRange(4, 1).setFormula(formula);

  sheet.getRange(4, 1, 1000, 1).setNumberFormat('DD.MM.YYYY');
  sheet.getRange(4, 3, 1000, 1).setNumberFormat('HH:MM');

  // Bedingte Formatierungen
  var heimRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$F4="Heim"')
    .setBackground('#C8E6C9')
    .setRanges([sheet.getRange('A4:I1000')])
    .build();

  var auswaertsRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$F4="Auswärts"')
    .setBackground('#BBDEFB')
    .setRanges([sheet.getRange('A4:I1000')])
    .build();

  var fehlerRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=REGEXMATCH($I4' + sep + ' "❌")')
    .setBackground('#FFCDD2')
    .setRanges([sheet.getRange('A4:I1000')])
    .build();

  var warnungRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=REGEXMATCH($I4' + sep + ' "⚠️")')
    .setBackground('#FFE0B2')
    .setRanges([sheet.getRange('A4:I1000')])
    .build();

  var sperrRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=REGEXMATCH($I4' + sep + ' "Sperrung")')
    .setBackground('#E0E0E0')
    .setRanges([sheet.getRange('A4:I1000')])
    .build();

  var rules = sheet.getConditionalFormatRules();
  rules.push(heimRule, auswaertsRule, fehlerRule, warnungRule, sperrRule);
  sheet.setConditionalFormatRules(rules);

  sheet.setColumnWidth(1, 110);
  sheet.setColumnWidth(2, 50);
  sheet.setColumnWidth(3, 90);
  sheet.setColumnWidth(4, 200);
  sheet.setColumnWidth(5, 180);
  sheet.setColumnWidth(6, 60);
  sheet.setColumnWidth(7, 200);
  sheet.setColumnWidth(8, 250);
  sheet.setColumnWidth(9, 300);

  sheet.setFrozenRows(3);
  protectRange(sheet, 'A3:I3');
}

// ==================== VALIDIERUNG ====================

function handleEdit(e) {
  var sheet = e.range.getSheet();
  var name = sheet.getName();

  if (name === CONFIG.SHEET_EINGABE) {
    if (e.range.getRow() < 2) return;
    var col = e.range.getColumn();

  // Spalte B (Startzeit): "12" → 12:00 konvertieren + Endzeit berechnen
  if (col === 2) {
    var rawVal = e.range.getValue();
    if (typeof rawVal === 'number' && rawVal >= 1 && rawVal < 24) {
      var h = Math.floor(rawVal);
      var m = Math.round((rawVal - h) * 60);
      e.range.setValue((h * 60 + m) / 1440);
      e.range.setNumberFormat('HH:MM');
    }
    computeEndzeit(e.range);
  }

    var relevantCols = [1, 2, 3, 4, 5, 7];
    if (relevantCols.indexOf(col) === -1 && col !== 6 && col !== 8) return;
    validateAllEntries();
  } else if (name === CONFIG.SHEET_SPERRUNGEN) {
    if (e.range.getRow() < 2) return;
    validateAllEntries();
  }
}

function computeEndzeit(startCell) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = startCell.getSheet();
  var row = startCell.getRow();
  var setupSheet = ss.getSheetByName(CONFIG.SHEET_SETUP);
  if (!setupSheet) return;
  var spieldauer = setupSheet.getRange('F1').getValue();
  if (typeof spieldauer !== 'number' || spieldauer <= 0) spieldauer = CONFIG.GAME_DURATION_HOURS;
  var startVal = startCell.getValue();
  var frac = timeToFraction(startVal);
  if (frac !== null) {
    sheet.getRange(row, 3).setValue(frac + spieldauer / 24).setNumberFormat('HH:MM');
  } else {
    sheet.getRange(row, 3).clearContent();
  }
}

function validateAllEntries() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_EINGABE);
  if (!sheet) return;

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  var data = sheet.getRange(2, 1, lastRow - 1, 9).getValues();

  var teams = readTeams(ss);
  var sperrungen = readSperrungen(ss);

  var statusMessages = {};
  for (var i = 0; i < data.length; i++) {
    statusMessages[i] = [];
  }

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var datum = row[0];
    var startzeit = row[1];
    var endzeit = row[2];
    var ha = row[4];
    var bereich = row[6];

    if (!isValidDate(datum)) continue;

    var weekday = datum.getDay();
    var dateKey = datumToKey(datum);

    // 1a. Wochentag (nur Heimspiele)
    if (ha === 'Heim' && CONFIG.ALLOWED_WEEKDAYS.indexOf(weekday) === -1) {
      statusMessages[i].push('❌ Nur Di, Mi, Fr, Sa, So erlaubt');
    }

    // 1a2. Bereich Pflicht für Heimspiele
    if (ha === 'Heim' && !bereich) {
      statusMessages[i].push('❌ Bereich fehlt');
    }

    // 1b. Mittwoch: nur Kleine Halle
    if (weekday === 3 && ha === 'Heim' && bereich && bereich !== CONFIG.WEDNESDAY_AREA) {
      statusMessages[i].push('❌ Mittwochs nur "' + CONFIG.WEDNESDAY_AREA + '" buchbar');
    }

    // 1c. Di/Fr: max. 2 Bereiche
    if ((weekday === 2 || weekday === 5) && ha === 'Heim' && bereich) {
      var areasOnDate = {};
      for (var j = 0; j < data.length; j++) {
        if (isValidDate(data[j][0]) && datumToKey(data[j][0]) === dateKey &&
            data[j][4] === 'Heim' && data[j][6]) {
          areasOnDate[data[j][6]] = true;
        }
      }
      if (Object.keys(areasOnDate).length > CONFIG.MAX_AREAS_TUE_FRI) {
        statusMessages[i].push('❌ Di/Fr: max. ' + CONFIG.MAX_AREAS_TUE_FRI +
          ' Bereiche buchbar (' + Object.keys(areasOnDate).length + ' wären belegt)');
      }
    }

    // 1d. Doppelbuchung prüfen
    if (ha === 'Heim' && bereich) {
      for (var k = 0; k < data.length; k++) {
          if (k !== i && isValidDate(data[k][0]) &&
              datumToKey(data[k][0]) === dateKey &&
              data[k][4] === 'Heim' &&
              data[k][6] === bereich) {

            if (weekday === 6 || weekday === 0) {
              var otherStart = data[k][1];
              var otherEnd = data[k][2];
              var s1 = timeToFraction(startzeit), e1 = timeToFraction(endzeit);
              var s2 = timeToFraction(otherStart), e2 = timeToFraction(otherEnd);
              if (s1 !== null && e1 !== null && s2 !== null && e2 !== null) {
                if (s1 < e2 && e1 > s2) {
                  statusMessages[i].push('❌ Zeitüberlappung in "' + bereich + '"');
                  break;
                }
              } else {
                statusMessages[i].push('❌ "' + bereich + '" bereits belegt (Startzeit fehlt)');
                break;
              }
            } else {
              statusMessages[i].push('❌ "' + bereich + '" an diesem Tag bereits belegt');
              break;
            }
          }
      }
    }

    // 1e. Sperrungen prüfen
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
  }
  }

  // 2. Benachbarte Teams
  var adjacentMessages = checkAdjacentTeams(data, teams);

  // 3. Status schreiben (Spalte I)
  for (var i = 0; i < data.length; i++) {
    var allMessages = statusMessages[i].slice();
    if (adjacentMessages[i]) {
      allMessages = allMessages.concat(adjacentMessages[i]);
    }
    sheet.getRange(i + 2, 9).setValue(allMessages.join(' | '));
  }
}

// -------------------- Hilfsfunktionen Validierung --------------------

function checkAdjacentTeams(data, teams) {
  var messages = {};

  var entriesByDate = {};
  for (var i = 0; i < data.length; i++) {
    if (!isValidDate(data[i][0]) || !data[i][3]) continue;
    var key = datumToKey(data[i][0]);
    if (!entriesByDate[key]) entriesByDate[key] = [];
    entriesByDate[key].push({ index: i, teamName: data[i][3] });
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

function readTeams(ss) {
  var sheet = ss.getSheetByName(CONFIG.SHEET_SETUP);
  if (!sheet) return {};

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return {};

  var raw = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
  var teams = {};
  for (var i = 0; i < raw.length; i++) {
    if (raw[i][0] && raw[i][1] && raw[i][2]) {
      teams[raw[i][2]] = { rang: raw[i][0], gruppe: raw[i][1] };
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
  protection.setWarningOnly(true);
}

function createTrigger() {
  // Alle bestehenden Trigger löschen (auch Altlasten aus früheren Versionen)
  ScriptApp.getProjectTriggers().forEach(function(t) {
    ScriptApp.deleteTrigger(t);
  });

  ScriptApp.newTrigger('handleEdit')
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onEdit()
    .create();
}

// ==================== TSV-EXPORT ====================

function downloadTSV() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var result = '';

  // Setup Teams
  result += '# data/setup_teams.tsv\n';
  result += 'Rang\tGruppe\tTeamname\tKurzname\n';
  var setupSheet = ss.getSheetByName(CONFIG.SHEET_SETUP);
  if (setupSheet) {
    var lastRow = Math.max(setupSheet.getLastRow(), 1);
    if (lastRow >= 2) {
      var data = setupSheet.getRange(2, 1, lastRow - 1, 4).getValues();
      for (var i = 0; i < data.length; i++) {
        if (data[i][0]) {
          result += [data[i][0], data[i][1], data[i][2], data[i][3]].join('\t') + '\n';
        }
      }
    }
  }
  result += '\n';

  // Eingabe
  result += '# data/eingabe.tsv\n';
  result += 'Datum\tStartzeit\tspäteste Endzeit\tTeam\tHeim/Auswärts\tGegner\tBereich\tKommentar\n';
  var eingabeSheet = ss.getSheetByName(CONFIG.SHEET_EINGABE);
  if (eingabeSheet) {
    var lastRow2 = Math.max(eingabeSheet.getLastRow(), 1);
    if (lastRow2 >= 2) {
      var data2 = eingabeSheet.getRange(2, 1, lastRow2 - 1, 9).getValues();
      for (var j = 0; j < data2.length; j++) {
        if (data2[j][0]) {
          var d = data2[j][0];
          var dateStr = d instanceof Date ? ('0' + d.getDate()).slice(-2) + '.' + ('0' + (d.getMonth() + 1)).slice(-2) + '.' + d.getFullYear() : d;
          var st = _formatTime(data2[j][1]);
          var et = _formatTime(data2[j][2]);
          result += [dateStr, st, et, data2[j][3], data2[j][4], data2[j][5], data2[j][6], data2[j][7]].join('\t') + '\n';
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
