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
 *   Date of PO Change Request | Status | Remarks | Name & Designation of
 *   Delinquent | Present Status | Under Stay (CAT/High Court) | Charge Memo Date
 */

/**
 * IMPORTANT: this script must point at your actual case register Sheet.
 * Paste the Sheet's ID here (the long string in its URL, between
 * /d/ and /edit) — this removes any ambiguity about which spreadsheet
 * the script is bound to.
 */
var SPREADSHEET_ID = "1I5WKmv3-hJNcc8kutou_SEtwpz8jOQMEm3yljrwL9OM";

var LITIGATION_COLS = ['ID', 'Date of Receipt', 'Email Date', 'Email Time', 'File No', 'Computer No', 'RC No./Case No.', 'Subject', 'Status', 'Date Communicated to CBI', 'Remarks', 'Date received for DoPT action', 'Case Timeline', 'Date file sent to DoLA', 'Date DoLA advice received'];

var SHEETS = {
  SLP: LITIGATION_COLS,
  Appeals: LITIGATION_COLS,
  Withdrawal: LITIGATION_COLS,
  Disciplinary: ['ID', 'Date of Receipt', 'Checklist Compliance', 'Date of Return of Proposal', 'Compliance Received from CBI', 'Date of Receiving Compliance', 'File No', 'Computer No', 'Subject', 'Stage', 'Date of IO Change Request', 'Date of PO Change Request', 'Status', 'Remarks', 'Name & Designation of Delinquent', 'Present Status', 'Under Stay (CAT/High Court)', 'Charge Memo Date', 'Date received for DoPT action', 'Date final order issued', 'Case Timeline', 'Date chargesheet issued', 'Date representation sought from CO', 'Date representation received from CO', 'Date IO/PO names sought from CBI', 'Date IO/PO appointed', 'Date inquiry commenced', 'Date inquiry report received', 'Date representation sought on inquiry report', 'Date representation received on inquiry report', 'Date case sent to UPSC', 'Date UPSC advice received', 'Date representation sought on UPSC advice', 'Date representation received on UPSC advice']
};

function doGet(e) {
  try {
    var action = e.parameter.action || 'list';
    if (action === 'list') {
      var cache = CacheService.getScriptCache();
      var cached = cache.get('list_response');
      if (cached) {
        return ContentService.createTextOutput(cached).setMimeType(ContentService.MimeType.JSON);
      }
      var all = {};
      Object.keys(SHEETS).forEach(function (name) {
        all[name] = readSheet_(name);
      });
      var payload = JSON.stringify({ ok: true, data: all });
      // Read-only response cache: makes repeat dashboard loads much faster.
      // It is cleared immediately after every dashboard add/update/delete.
      cache.put('list_response', payload, 60); // seconds
      return ContentService.createTextOutput(payload).setMimeType(ContentService.MimeType.JSON);
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

    var result;
    if (action === 'add') {
      var row = addRow_(sheetName, body.data);
      result = jsonOut_({ ok: true, data: row });
    } else if (action === 'update') {
      var updated = updateRow_(sheetName, body.id, body.data);
      result = jsonOut_({ ok: updated, data: body.data });
    } else if (action === 'delete') {
      var deleted = deleteRow_(sheetName, body.id);
      result = jsonOut_({ ok: deleted });
    } else {
      return jsonOut_({ ok: false, error: 'Unknown action: ' + action });
    }
    // Any successful write invalidates the cached list so the next
    // load reflects the change immediately instead of stale data.
    CacheService.getScriptCache().remove('list_response');
    return result;
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}

var _ssCache_ = null;
function getSpreadsheet_() {
  if (!_ssCache_) _ssCache_ = SpreadsheetApp.openById(SPREADSHEET_ID);
  return _ssCache_;
}

function getSheet_(name) {
  var ss = getSpreadsheet_();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(SHEETS[name]);
  }
  return sh;
}

/**
 * Reads a tab's data. Also verifies the header row matches SHEETS[name]
 * and repairs it in place if not (e.g. data was entered before headers
 * existed, or new columns were added to SHEETS) — using the data already
 * fetched by this same call, so the common case (headers already
 * correct) costs exactly one read, no separate round-trip.
 */
function readSheet_(name) {
  var sh = getSheet_(name);
  var values = sh.getDataRange().getValues();
  var headers = SHEETS[name];

  if (values.length === 0) {
    sh.appendRow(headers);
    return [];
  }
  var currentHeaders = values[0];
  var matches = headers.every(function (h, i) { return currentHeaders[i] === h; }) && currentHeaders.length === headers.length;
  if (!matches) {
    // Safe schema extension: when new fields are appended to the expected
    // layout, append only their header cells. Existing columns and every
    // existing row remain exactly where they are.
    var isPrefix = currentHeaders.length < headers.length && currentHeaders.every(function (h, i) { return headers[i] === h; });
    if (isPrefix) {
      var additions = headers.slice(currentHeaders.length);
      sh.getRange(1, currentHeaders.length + 1, 1, additions.length).setValues([additions]);
    } else {
      throw new Error('Sheet headers for ' + name + ' do not match the expected layout. No data was changed.');
    }
  } else {
    headers = currentHeaders;
  }

  if (values.length < 2) return [];
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
