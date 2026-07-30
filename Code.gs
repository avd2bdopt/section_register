/**
 * AVD-II (B) CASE REGISTER — Apps Script backend
 * ------------------------------------------------------
 * Deploy this bound to your Google Sheet (Extensions > Apps Script),
 * then deploy as a Web App (Deploy > New deployment > Web app).
 *   - Execute as: Me
 *   - Who has access: Anyone with the link
 * Copy the resulting /exec URL into CONFIG.API_URL in index.html.
 *
 * Expected sheet tabs (create these exact names in the spreadsheet).
 * You do NOT need to type headers by hand — getSheet_() creates them
 * automatically the first time each tab is used, in the exact order
 * listed below.
 *
 * SLP / Appeals / Withdrawal tabs — identical columns:
 *   ID | Date of Receipt | Email Date | Email Time | File No | Computer No |
 *   RC No./Case No. | Subject | Status | Date Communicated to CBI | Remarks
 *
 * Disciplinary tab columns:
 *   ID | Date of Receipt | Checklist Compliance | Date of Return of Proposal |
 *   Compliance Received from CBI | Date of Receiving Compliance | File No |
 *   Computer No | Subject | Stage | Date of IO Change Request |
 *   Date of PO Change Request | Status | Remarks
 */

var LITIGATION_COLS = ['ID', 'Date of Receipt', 'Email Date', 'Email Time', 'File No', 'Computer No', 'RC No./Case No.', 'Subject', 'Status', 'Date Communicated to CBI', 'Remarks'];

var SHEETS = {
  SLP: LITIGATION_COLS,
  Appeals: LITIGATION_COLS,
  Withdrawal: LITIGATION_COLS,
  Disciplinary: ['ID', 'Date of Receipt', 'Checklist Compliance', 'Date of Return of Proposal', 'Compliance Received from CBI', 'Date of Receiving Compliance', 'File No', 'Computer No', 'Subject', 'Stage', 'Date of IO Change Request', 'Date of PO Change Request', 'Status', 'Remarks']
};

function doGet(e) {
  try {
    var action = e.parameter.action || 'list';
    if (action === 'list') {
      var all = {};
      Object.keys(SHEETS).forEach(function (name) {
        all[name] = readSheet_(name);
      });
      return jsonOut_({ ok: true, data: all });
    }
    return jsonOut_({ ok: false, error: 'Unknown action' });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    var action = body.action;
    var sheetName = body.sheet;

    if (!SHEETS[sheetName]) {
      return jsonOut_({ ok: false, error: 'Unknown sheet: ' + sheetName });
    }

    if (action === 'add') {
      var row = addRow_(sheetName, body.data);
      return jsonOut_({ ok: true, data: row });
    }
    if (action === 'update') {
      var updated = updateRow_(sheetName, body.id, body.data);
      return jsonOut_({ ok: updated, data: body.data });
    }
    if (action === 'delete') {
      var deleted = deleteRow_(sheetName, body.id);
      return jsonOut_({ ok: deleted });
    }
    return jsonOut_({ ok: false, error: 'Unknown action: ' + action });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}

function getSheet_(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(SHEETS[name]);
  }
  return sh;
}

function readSheet_(name) {
  var sh = getSheet_(name);
  var values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  var headers = values[0];
  var out = [];
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    if (row.join('') === '') continue;
    var obj = {};
    for (var c = 0; c < headers.length; c++) {
      var val = row[c];
      if (val instanceof Date) val = Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      obj[headers[c]] = val;
    }
    out.push(obj);
  }
  return out;
}

function addRow_(name, data) {
  var sh = getSheet_(name);
  var headers = SHEETS[name];
  var id = 'R' + new Date().getTime() + Math.floor(Math.random() * 1000);
  data['ID'] = id;
  var row = headers.map(function (h) { return data[h] !== undefined ? data[h] : ''; });
  sh.appendRow(row);
  return data;
}

function updateRow_(name, id, data) {
  var sh = getSheet_(name);
  var headers = SHEETS[name];
  var idCol = headers.indexOf('ID') + 1;
  var values = sh.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][idCol - 1]) === String(id)) {
      var rowNum = i + 1;
      var row = headers.map(function (h) {
        return data[h] !== undefined ? data[h] : values[i][headers.indexOf(h)];
      });
      sh.getRange(rowNum, 1, 1, headers.length).setValues([row]);
      return true;
    }
  }
  return false;
}

function deleteRow_(name, id) {
  var sh = getSheet_(name);
  var headers = SHEETS[name];
  var idCol = headers.indexOf('ID') + 1;
  var values = sh.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][idCol - 1]) === String(id)) {
      sh.deleteRow(i + 1);
      return true;
    }
  }
  return false;
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
