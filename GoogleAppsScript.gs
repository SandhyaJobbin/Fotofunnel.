/**
 * FotoFunnel Training — Google Sheets backend  (v3: waves + sessions)
 * ------------------------------------------------------------------
 * SETUP (one time):
 * 1. Create a Google Sheet. You do NOT need to add tabs — this script
 *    creates "Records", "Waves" and "Sessions" automatically.
 * 2. In the Sheet:  Extensions ▸ Apps Script.
 * 3. Delete any starter code, paste ALL of this file, press Save.
 * 4. Set SHEET_ID below to your Sheet's ID (the long string in the URL
 *    between /d/ and /edit), or leave '' if the script is bound to it.
 * 5. Deploy ▸ New deployment ▸ type "Web app".
 *      - Execute as:  Me
 *      - Who has access:  Anyone
 *    Deploy, authorise, and COPY the "/exec" Web app URL.
 * 6. Paste that URL into the GAS_URL constant in index.html.
 * ------------------------------------------------------------------
 */

var SHEET_ID = '';               // optional: paste Sheet ID, or leave '' if bound
var RECORDS_TAB  = 'Records';
var WAVES_TAB    = 'Waves';
var SESSIONS_TAB = 'Sessions';

var RECORD_HEADERS = ['attempt_id','employee_id','name','country','trainer_name','wave','session','queue',
                      'level','attempt_number','score','correct','total_questions','percentage',
                      'htq','result','responses','completed_at'];
var WAVE_HEADERS    = ['wave_name','trainer_name','created_at','question_count'];
var SESSION_HEADERS = ['session_name','trainer_name','time','queue','created_at','question_count','questions'];

function ss_() { return SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet(); }
function getSheet_(name, headers) {
  var ss = ss_(); var sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); sh.appendRow(headers); }
  if (sh.getLastRow() === 0) sh.appendRow(headers);
  return sh;
}
function readTab_(name, headers) {
  var sh = getSheet_(name, headers);
  var values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  var head = values[0];
  return values.slice(1).map(function (row) {
    var o = {}; head.forEach(function (h, i) { o[h] = row[i]; }); return o;
  });
}
function json_(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }

// Sessions store their questions as a JSON string in the sheet; parse on read.
function parseSessions_(rows) {
  return rows.map(function (r) {
    var q = [];
    if (r.questions) { try { q = JSON.parse(r.questions); } catch (e) { q = []; } }
    r.questions = q;
    return r;
  });
}

function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) || 'getAll';
  var out = {};
  if (action === 'getAll' || action === 'getRecords')  out.records  = readTab_(RECORDS_TAB, RECORD_HEADERS);
  if (action === 'getAll' || action === 'getWaves')    out.waves    = readTab_(WAVES_TAB, WAVE_HEADERS);
  if (action === 'getAll' || action === 'getSessions') out.sessions = parseSessions_(readTab_(SESSIONS_TAB, SESSION_HEADERS));
  return json_(out);
}

function doPost(e) {
  var body = {};
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'bad json' }); }
  var action = body.action;
  if (action === 'startAttempt')  return startAttempt_(body);
  if (action === 'finishAttempt') return finishAttempt_(body);
  if (action === 'createWave')    return createWave_(body);
  if (action === 'createSession') return createSession_(body);
  return json_({ ok: false, error: 'unknown action' });
}

function startAttempt_(b) {
  var sh = getSheet_(RECORDS_TAB, RECORD_HEADERS);
  sh.appendRow(RECORD_HEADERS.map(function (h) { return b[h] !== undefined ? b[h] : ''; }));
  return json_({ ok: true, attempt_id: b.attempt_id });
}
function finishAttempt_(b) {
  var sh = getSheet_(RECORDS_TAB, RECORD_HEADERS);
  var values = sh.getDataRange().getValues();
  var head = values[0]; var idCol = head.indexOf('attempt_id');
  for (var r = 1; r < values.length; r++) {
    if (String(values[r][idCol]) === String(b.attempt_id)) {
      RECORD_HEADERS.forEach(function (h, i) { if (b[h] !== undefined) sh.getRange(r + 1, i + 1).setValue(b[h]); });
      return json_({ ok: true, updated: true });
    }
  }
  sh.appendRow(RECORD_HEADERS.map(function (h) { return b[h] !== undefined ? b[h] : ''; }));
  return json_({ ok: true, appended: true });
}
function createWave_(b) {
  var sh = getSheet_(WAVES_TAB, WAVE_HEADERS);
  sh.appendRow(WAVE_HEADERS.map(function (h) { return b[h] !== undefined ? b[h] : ''; }));
  return json_({ ok: true });
}
function createSession_(b) {
  var sh = getSheet_(SESSIONS_TAB, SESSION_HEADERS);
  var row = SESSION_HEADERS.map(function (h) {
    if (h === 'questions') return JSON.stringify(b.questions || []);
    return b[h] !== undefined ? b[h] : '';
  });
  sh.appendRow(row);
  return json_({ ok: true });
}
