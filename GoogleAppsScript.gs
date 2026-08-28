/**
 * FotoFunnel Training — Google Sheets backend  (v6: aggregated scores + Completed status)
 * ------------------------------------------------------------------
 * SETUP (one time):
 * 1. Create a Google Sheet. Do NOT add tabs manually — the script
 *    creates "Records", "Waves" and "Sessions" automatically.
 * 2. Extensions ▸ Apps Script → delete starter code → paste this → Save.
 * 3. Set SHEET_ID below (the ID in the Sheet URL), or leave '' if bound.
 * 4. Deploy ▸ New deployment ▸ Web app
 *      Execute as: Me  |  Who has access: Anyone
 *    Copy the /exec URL and paste as GAS_URL in index.html.
 * 5. When redeploying: always choose "New version".
 * ------------------------------------------------------------------
 *
 * RECORDS TAB COLUMNS (auto-created on first run):
 *   attempt_id, employee_id, name, country, trainer_name, wave, session, queue,
 *   level, attempt_number, score, correct, total_questions, percentage,
 *   htq, result, responses, completed_at,
 *   level1_score, level1_pct, level2_score, level2_pct, level3_score, level3_pct
 *
 *   result values: "In Progress" | "Failed" | "Passed" | "Completed"
 *   ("Completed" = all 3 levels passed for this queue)
 *
 * WAVES TAB COLUMNS:
 *   wave_name, trainer_name, created_at, question_count
 *
 * SESSIONS TAB COLUMNS:
 *   session_name, trainer_name, time, queue, created_at, question_count, questions
 *   (questions is a JSON array; each item: image, scenario, correct, policy,
 *    feedbackCorrect, feedbackWrong)
 * ------------------------------------------------------------------
 */

var SHEET_ID     = '';
var RECORDS_TAB  = 'Records';
var WAVES_TAB    = 'Waves';
var SESSIONS_TAB = 'Sessions';

var RECORD_HEADERS = [
  'attempt_id','employee_id','name','country','trainer_name','wave','session','queue',
  'level','attempt_number','score','correct','total_questions','percentage',
  'htq','result','responses','completed_at',
  'level1_score','level1_pct','level2_score','level2_pct',
  'level3_score','level3_pct'
];
var WAVE_HEADERS    = ['wave_name','trainer_name','created_at','question_count'];
var SESSION_HEADERS = ['session_name','trainer_name','time','queue','created_at','question_count','questions'];

/* ─────────── helpers ─────────── */
function ss_() {
  return SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
}

function getSheet_(name, headers) {
  var ss = ss_();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(headers);
    styleHeaderRow_(sh, headers.length);
    return sh;
  }
  if (sh.getLastRow() === 0) {
    sh.appendRow(headers);
    styleHeaderRow_(sh, headers.length);
  }
  return sh;
}

function styleHeaderRow_(sh, colCount) {
  try {
    var r = sh.getRange(1, 1, 1, colCount);
    r.setBackground('#04140F')
     .setFontColor('#FFFFFF')
     .setFontWeight('bold')
     .setFontSize(10)
     .setHorizontalAlignment('center')
     .setVerticalAlignment('middle')
     .setWrap(true);
    sh.setFrozenRows(1);
    sh.setRowHeight(1, 36);
  } catch(e) {}
}

function readTab_(name, headers) {
  var sh = getSheet_(name, headers);
  var values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  var head = values[0];
  return values.slice(1).map(function(row) {
    var o = {};
    head.forEach(function(h, i) { o[h] = row[i]; });
    return o;
  });
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function parseSessions_(rows) {
  return rows.map(function(r) {
    var q = [];
    if (r.questions) { try { q = JSON.parse(r.questions); } catch(e) { q = []; } }
    r.questions = q;
    return r;
  });
}

/* ─────────── HTTP handlers ─────────── */
function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) || 'getAll';
  var out = {};
  if (action === 'getAll' || action === 'getRecords')
    out.records  = readTab_(RECORDS_TAB, RECORD_HEADERS);
  if (action === 'getAll' || action === 'getWaves')
    out.waves    = readTab_(WAVES_TAB, WAVE_HEADERS);
  if (action === 'getAll' || action === 'getSessions')
    out.sessions = parseSessions_(readTab_(SESSIONS_TAB, SESSION_HEADERS));
  return json_(out);
}

function doPost(e) {
  var body = {};
  try { body = JSON.parse(e.postData.contents); }
  catch(err) { return json_({ ok: false, error: 'bad json' }); }
  var action = body.action;
  if (action === 'startAttempt')  return startAttempt_(body);
  if (action === 'finishAttempt') return finishAttempt_(body);
  if (action === 'createWave')    return createWave_(body);
  if (action === 'createSession') return createSession_(body);
  return json_({ ok: false, error: 'unknown action' });
}

/* ─────────── record actions ─────────── */
function startAttempt_(b) {
  var sh = getSheet_(RECORDS_TAB, RECORD_HEADERS);
  sh.appendRow(RECORD_HEADERS.map(function(h) {
    return b[h] !== undefined ? b[h] : '';
  }));
  colorLastRow_(sh);
  return json_({ ok: true, attempt_id: b.attempt_id });
}

function finishAttempt_(b) {
  var sh     = getSheet_(RECORDS_TAB, RECORD_HEADERS);
  var values = sh.getDataRange().getValues();
  var head   = values[0];
  var idCol  = head.indexOf('attempt_id');

  for (var r = 1; r < values.length; r++) {
    if (String(values[r][idCol]) === String(b.attempt_id)) {
      RECORD_HEADERS.forEach(function(h, i) {
        if (b[h] !== undefined && b[h] !== '') {
          sh.getRange(r + 1, i + 1).setValue(b[h]);
        }
      });
      // After updating, check if all 3 levels for this employee+queue are passed
      // and update result to "Completed" on this row if so
      var empId = String(b.employee_id || '').trim();
      var queue = String(b.queue || '').trim().toLowerCase();
      if (empId && queue) {
        updateCompletedStatus_(sh, values, head, empId, queue, r + 1);
      }
      colorResultCell_(sh, r + 1, head.indexOf('result') + 1, b.result);
      return json_({ ok: true, updated: true });
    }
  }

  // Row not found — append
  sh.appendRow(RECORD_HEADERS.map(function(h) {
    return b[h] !== undefined ? b[h] : '';
  }));
  colorLastRow_(sh);
  colorResultCell_(sh, sh.getLastRow(), RECORD_HEADERS.indexOf('result') + 1, b.result);
  return json_({ ok: true, appended: true });
}

/**
 * After finishing a level attempt, check if this employee has now passed
 * all 3 levels for this queue. If so, mark the current row as "Completed".
 */
function updateCompletedStatus_(sh, values, head, empId, queue, currentSheetRow) {
  try {
    var empIdCol  = head.indexOf('employee_id');
    var queueCol  = head.indexOf('queue');
    var levelCol  = head.indexOf('level');
    var resultCol = head.indexOf('result');
    var pctCol    = head.indexOf('percentage');
    if (empIdCol < 0 || queueCol < 0 || levelCol < 0 || resultCol < 0) return;

    var PASS_MARK = 80;
    var levelsPassed = {};

    // Scan all rows for this employee + queue
    for (var r = 1; r < values.length; r++) {
      var rowEmpId = String(values[r][empIdCol] || '').trim();
      var rowQueue = String(values[r][queueCol]  || '').trim().toLowerCase();
      var rowLevel = String(values[r][levelCol]  || '').trim();
      var rowPct   = parseFloat(values[r][pctCol]) || 0;
      if (rowEmpId !== empId || rowQueue !== queue) continue;
      if (rowPct >= PASS_MARK && (rowLevel === '1' || rowLevel === '2' || rowLevel === '3')) {
        levelsPassed[rowLevel] = true;
      }
    }

    if (levelsPassed['1'] && levelsPassed['2'] && levelsPassed['3']) {
      // Mark current row as Completed
      sh.getRange(currentSheetRow, resultCol + 1).setValue('Completed');
      sh.getRange(currentSheetRow, resultCol + 1)
        .setBackground('#D4EDDA').setFontColor('#155724').setFontWeight('bold');
    }
  } catch(e) {}
}

/* ─────────── wave / session actions ─────────── */
function createWave_(b) {
  var sh = getSheet_(WAVES_TAB, WAVE_HEADERS);
  sh.appendRow(WAVE_HEADERS.map(function(h) { return b[h] !== undefined ? b[h] : ''; }));
  colorLastRow_(sh);
  return json_({ ok: true });
}

function createSession_(b) {
  var sh = getSheet_(SESSIONS_TAB, SESSION_HEADERS);
  var row = SESSION_HEADERS.map(function(h) {
    if (h === 'questions') return JSON.stringify(b.questions || []);
    return b[h] !== undefined ? b[h] : '';
  });
  sh.appendRow(row);
  colorLastRow_(sh);
  return json_({ ok: true });
}

/* ─────────── sheet styling helpers ─────────── */
function colorLastRow_(sh) {
  try {
    var row  = sh.getLastRow();
    var cols = sh.getLastColumn();
    if (row < 2) return;
    var bg = (row % 2 === 0) ? '#F0F7F4' : '#FFFFFF';
    sh.getRange(row, 1, 1, cols).setBackground(bg);
  } catch(e) {}
}

function colorResultCell_(sh, rowNum, colNum, result) {
  if (!colNum || colNum < 1) return;
  try {
    var cell = sh.getRange(rowNum, colNum);
    if (result === 'Completed' || result === 'Passed') {
      cell.setBackground('#D4EDDA').setFontColor('#155724').setFontWeight('bold');
    } else if (result === 'Failed') {
      cell.setBackground('#F8D7DA').setFontColor('#721C24').setFontWeight('bold');
    } else {
      cell.setBackground('#FFF3CD').setFontColor('#856404').setFontWeight('bold');
    }
  } catch(e) {}
}
