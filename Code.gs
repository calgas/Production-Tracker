/**
 * CALGAS CAPACITORS — PRODUCTION TRACKER · BACKEND
 *
 * Order-to-dispatch production tracking. One Apps Script file bound to the
 * Production Tracker spreadsheet, serving a GitHub Pages PWA over a
 * single POST endpoint.
 *
 * This project issues no logins of its own. Sign-in happens in Stock
 * Management, and every request here is validated against that
 * spreadsheet's Sessions tab. Anything that moves stock is posted to Stock
 * Management's API rather than written here — its lock only protects
 * writes that go through its own code, so a second script writing to
 * RM_Current_Stock directly would corrupt balances.
 *
 * FRESH INSTALL
 *   1. Create the spreadsheet, open Extensions > Apps Script, paste this file.
 *   2. Check STOCK_SHEET_ID and STOCK_API_URL below point at the live
 *      Stock Management sheet and deployment.
 *   3. Run ADMIN_setupSheet() once from the editor and accept the
 *      permission prompts.
 *   4. Deploy > New deployment > Web app, Execute as: Me,
 *      Who has access: Anyone. Put the /exec URL in the frontend.
 *   5. Later code changes: Deploy > Manage deployments > Edit > New
 *      version, which keeps the same /exec URL.
 */

var APP_VERSION = '3.1.0';

/* =========================================================================
   CONFIG
   ========================================================================= */

// Stock Management's spreadsheet — read directly for Sessions, Users and
// the RM/FG catalogs.
var STOCK_SHEET_ID = '1iCFlMmzjhBiMsS47ahy56gm35LZc1GidJP5yVPnYAbc';

// Stock Management's deployment — the write path for anything that moves
// stock. Server-to-server via UrlFetchApp, a different execution context
// from the browser, so this URL is needed even though the frontend has
// its own copy.
var STOCK_API_URL = 'https://script.google.com/macros/s/AKfycbyxTvERmZY63TC9Pz4tMXb9x3mbT9hOe0YDuq2dNr5D5tkku2kYffFtmEzILuP1G9DrKA/exec';

// ELE Tracker's spreadsheet, where element batches are tracked now. Trace
// looks there for a code this sheet's own older Elements tab doesn't have.
// Left blank, Trace says where to look instead. The ID is the long part of
// the ELE spreadsheet's URL, between /d/ and /edit.
var ELE_SHEET_ID = '';

// How long a request waits for the write lock before giving up. The
// frontend's own timeout must be longer than this, or it stops listening
// while the write is still going ahead.
var WRITE_LOCK_WAIT_MS = 20000;

// How far back the idempotency log is searched for a repeated request.
// Only one column is read, so the window costs the same one round trip
// whatever its size — wide enough that a request left unresolved
// overnight is still recognised the next morning.
var IDEMPOTENCY_SCAN_ROWS = 2000;

// Actions that run without a session.
var PUBLIC_ACTIONS = ['ping'];

/* =========================================================================
   WEB APP ROUTER

   Shaped like Stock Management's: one POST endpoint, an action name in the
   body, a session token on everything but the public allowlist.

   Two rules carried over from there, both learned from real duplicates:

   1. Writes are serialised. The lock is taken once, here at the request
      boundary, rather than inside individual functions — several actions
      read a row, check its status and then write it back, and without a
      lock around the whole sequence two clicks can both pass the check.

   2. A write carrying a clientTxnId runs at most once. The ID identifies
      the entry, not the attempt, so a resend after a dropped connection
      replays the original answer instead of producing a second row. The
      frontend must mint one per entry and keep it until the entry is
      confirmed — see checkTransaction_.
   ========================================================================= */

function doPost(e) { return handleRequest_(e); }

// GET never reaches the action handler, so nothing — token included — can
// travel in a query string where it would land in server logs and browser
// history. The frontend only ever POSTs.
function doGet(e) {
  return jsonResponse_({ success: false, error: 'This API accepts POST requests only.' });
}

function handleRequest_(e) {
  try {
    var params = {};
    if (e && e.postData && e.postData.contents) params = JSON.parse(e.postData.contents);
    else if (e && e.parameter) params = e.parameter;

    var action = params.action;
    if (!action) throw new Error('Missing action');

    var session = null;
    if (PUBLIC_ACTIONS.indexOf(action) === -1) {
      session = validateSharedSession_(params.token);
      if (!session) throw new Error('Invalid or expired session — please log in again');
    }

    var data = isReadOnlyAction_(action)
      ? routeAction_(action, session, params)
      : (action === 'uploadHandoverFile'
          // Its own short locks around a slow Drive write, rather than the
          // one lock every other write holds throughout.
          ? uploadHandoverFile_(session, params)
          : action === 'receiveHandover'
            // Likewise around the call to Stock Management — see
            // receiveHandover_.
            ? receiveHandover_(session, params)
            : runWrite_(action, session, params));

    return jsonResponse_({ success: true, data: data });
  } catch (err) {
    return jsonResponse_({ success: false, error: err.message });
  }
}

// Anything not recognisably a read is treated as a write and locked. A new
// action that forgets to announce itself is therefore slow, never unsafe.
function isReadOnlyAction_(action) {
  if (action === 'ping' || action === 'whoami' || action === 'checkTransaction') return true;
  if (action === 'traceElement') return true;
  return /^get[A-Z]/.test(action);
}

function runWrite_(action, session, params) {
  return withWriteLock_(function () {
    var clientTxnId = String(params.clientTxnId || '').trim();

    // Checked inside the lock, so two arrivals of the same request can't
    // both find nothing and both write.
    if (clientTxnId) {
      var prior = findIdempotentResult_(clientTxnId);
      if (prior) {
        checkReplay_(prior, action, session);
        var replay = prior.result;
        if (replay && typeof replay === 'object' && !Array.isArray(replay)) replay.duplicate = true;
        return replay;
      }
    }

    var data = routeAction_(action, session, params);
    if (clientTxnId) recordIdempotentResult_(clientTxnId, action, session ? session.Username : '', data);
    return data;
  });
}

function withWriteLock_(fn) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(WRITE_LOCK_WAIT_MS);
  } catch (err) {
    throw new Error('The system is busy with another entry — please try again in a moment.');
  }
  try {
    return fn();
  } finally {
    // Commit before the lock goes. A write still pending when the next
    // request takes the lock is how two batches end up with one code.
    try { SpreadsheetApp.flush(); } catch (err) {}
    lock.releaseLock();
  }
}

function routeAction_(action, session, params) {
  switch (action) {
    // --- plumbing ---
    case 'ping':
      // No session: proves the deployment itself is reachable.
      return { ok: true, version: APP_VERSION, time: new Date() };
    case 'whoami':
      return getWhoAmI_(session);
    case 'checkTransaction':
      return checkTransaction_(params);

    // --- orders ---
    case 'placeOrder':
      return placeOrder_(session, params);
    case 'updateOrderStatus':
      return updateOrderStatus_(session, params);
    case 'getOrders':
      return getOrders_();
    case 'getFGCatalog':
      return getFGCatalog_();

    // --- the workflow ---
    case 'getWorkflow':
      return getWorkflow_(session);
    case 'getItems':
      return getItems_(params);
    case 'getHandovers':
      return getHandovers_(session, params);
    case 'getRMCatalog':
      return getRMCatalog_();
    case 'requestHandover':
      return requestHandover_(session, params);
    case 'sendHandover':
      return sendHandover_(session, params);
    // receiveHandover is routed in handleRequest_, outside runWrite_: it
    // takes and lets go of the lock itself.
    case 'rejectHandover':
      return closeHandover_(session, params, 'reject');
    case 'declineHandover':
      return closeHandover_(session, params, 'decline');
    case 'cancelHandover':
      return closeHandover_(session, params, 'cancel');
    case 'recordWIP':
      return recordWIP_(session, params);

    // --- files on handovers ---
    case 'getHandoverFiles':
      return getHandoverFiles_(session, params);
    case 'getHandoverFile':
      return getHandoverFile_(session, params);
    case 'removeHandoverFile':
      return removeHandoverFile_(session, params);

    // --- history: elements tracked before the workflow ---
    case 'traceElement':
      return traceElement_(params);

    default:
      throw new Error('Unknown action: ' + action);
  }
}

/* =========================================================================
   IDEMPOTENCY

   One tab recording every completed write against the client's transaction
   ID, with the answer that was sent back. A repeat of the same ID replays
   that answer rather than doing the work twice.

   A dedicated log rather than an ID column on each sheet, because writes
   here land in many different tabs and several change an existing row
   instead of adding one — there is no single row to hang the ID on.
   ========================================================================= */

var IDEMPOTENCY_TAB = 'Idempotency';

function findIdempotentResult_(clientTxnId) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(IDEMPOTENCY_TAB);
  if (!sheet) return null;
  var meta = sheetMeta_(sheet);
  if (meta.map['ClientTxnId'] === undefined) return null;

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  var startRow = Math.max(2, lastRow - IDEMPOTENCY_SCAN_ROWS + 1);
  var ids = sheet.getRange(startRow, meta.map['ClientTxnId'] + 1, lastRow - startRow + 1, 1).getValues();

  for (var i = ids.length - 1; i >= 0; i--) {
    if (String(ids[i][0]) !== clientTxnId) continue;
    var values = sheet.getRange(startRow + i, 1, 1, meta.lastCol).getValues()[0];
    var row = {};
    meta.headers.forEach(function (h, c) { if (h) row[h] = values[c]; });
    var result = null;
    try { result = JSON.parse(row.Result || 'null'); } catch (err) { result = null; }
    return { clientTxnId: clientTxnId, action: row.Action, user: row.User, at: row.Timestamp, result: result };
  }
  return null;
}

// A replay answers only the request it was recorded for. IDs are random per
// entry, so a mismatch means a client bug, not a duplicate — and replaying
// someone else's answer would report work that wasn't done.
function checkReplay_(prior, action, session) {
  if (prior.action && prior.action !== action) {
    throw new Error('This entry\u2019s ID was already used for a different action (' + prior.action + '). Reload the app and try again.');
  }
  if (prior.user && session && prior.user !== session.Username) {
    throw new Error('This entry\u2019s ID belongs to someone else\u2019s entry. Reload the app and try again.');
  }
}

function recordIdempotentResult_(clientTxnId, action, username, result) {
  var sheet = getOrCreateSheet_(SpreadsheetApp.getActiveSpreadsheet(), IDEMPOTENCY_TAB, TAB_SCHEMA[IDEMPOTENCY_TAB]);
  // Dates become ISO strings here, which is what they'd be over the wire
  // anyway — a replayed answer reads the same as the original did.
  appendRowByHeader_(sheet, {
    ClientTxnId: clientTxnId,
    Action: action,
    User: username || '',
    Timestamp: new Date(),
    Result: JSON.stringify(result === undefined ? null : result)
  });
}

// Read-only: lets the frontend ask whether a write it never got an answer
// for actually happened, instead of guessing or sending it again.
function checkTransaction_(params) {
  var clientTxnId = String(params.clientTxnId || '').trim();
  if (!clientTxnId) throw new Error('clientTxnId is required');
  var prior = findIdempotentResult_(clientTxnId);
  return prior
    ? { found: true, action: prior.action, at: prior.at, result: prior.result }
    : { found: false };
}

/* =========================================================================
   AUTH — sessions borrowed from Stock Management

   Login happens there; this project only validates. The session row's
   cached Role and Department are trusted rather than re-read from Users,
   the same trade-off Stock Management makes itself (a role change takes
   effect at next login), kept identical on purpose rather than making
   this project quietly stricter than the one issuing the token.

   Deliberately NOT cached. This script's cache is separate from Stock
   Management's, so a sign-out there could not clear an entry held here,
   and a revoked session would keep working until the cache expired. The
   lookup is one indexed search, so the saving wouldn't be worth it.
   ========================================================================= */

var _stockSpreadsheet_ = null;
function getStockSpreadsheet_() {
  if (!_stockSpreadsheet_) _stockSpreadsheet_ = SpreadsheetApp.openById(STOCK_SHEET_ID);
  return _stockSpreadsheet_;
}

function validateSharedSession_(token) {
  token = String(token || '').trim();
  if (!token) return null;

  var sheet = getStockSpreadsheet_().getSheetByName('Sessions');
  if (!sheet) return null;

  // Searching for the token beats reading the whole Sessions tab: the cost
  // is one row rather than one row per session ever issued.
  var cell = sheet.createTextFinder(token).matchEntireCell(true).findNext();
  if (!cell) return null;

  var meta = sheetMeta_(sheet);
  var values = sheet.getRange(cell.getRow(), 1, 1, meta.lastCol).getValues()[0];
  var row = {};
  meta.headers.forEach(function (h, c) { if (h) row[h] = values[c]; });

  // The search matches anywhere in the row; only the Token column counts.
  if (String(row.Token) !== token) return null;
  if (!row.ExpiresAt || new Date(row.ExpiresAt).getTime() < Date.now()) return null;
  return row;
}

// Stock Management's Admin role is the one piece of Role that means
// anything here. Everything else is decided by department.
function isAdmin_(session) {
  return session && session.Role === 'Admin';
}

/* =========================================================================
   ACCESS

   What someone may do here is decided in Stock Management — tiers from the
   Departments tab, per-person overrides, all managed in its Team & Access
   with everything else. Their grants travel on the session row this app
   already reads to check the token, so no lookup is needed. Admin holds
   everything.
   ========================================================================= */

var PT_FEATURES = {
  'pt.handover.send':    'send or return from your department',
  'pt.handover.receive': 'receive into your department',
  'pt.handover.any':     'act for any department',
  'pt.wip.record':       'record work in progress',
  'pt.view.all':         'see every department',
  'pt.orders.place':     'place orders'
};

function can_(session, key) {
  if (isAdmin_(session)) return true;
  return String(session.Perms || '').split(',').indexOf(key) > -1;
}

function requirePerm_(session, key) {
  if (!can_(session, key)) {
    throw new Error('You don\u2019t have access to ' + (PT_FEATURES[key] || key) +
                    '. Ask an Admin to grant it in Team & Access.');
  }
}

// Acting for a department: your own, or any, if you've been granted that.
function requireDept_(session, dept) {
  if (isAdmin_(session) || can_(session, 'pt.handover.any') || session.Department === dept) return;
  throw new Error('You can only do this for your own department (' + session.Department +
                  '). Ask an Admin for wider access.');
}

function canSee_(session, dept) {
  return isAdmin_(session) || can_(session, 'pt.view.all') || can_(session, 'pt.handover.any') || session.Department === dept;
}

function getWhoAmI_(session) {
  var usersSheet = getStockSpreadsheet_().getSheetByName('Users');
  var user = readRowByValue_(usersSheet, 'Username', session.Username);
  var name = session.Username;
  if (user && user.map['Name'] !== undefined) name = user.values[user.map['Name']] || session.Username;

  // Drives what the frontend offers. Each action re-checks for itself, so
  // this is convenience, not the security boundary.
  var permissions = {};
  Object.keys(PT_FEATURES).forEach(function (k) { permissions[k] = can_(session, k); });
  return {
    username: session.Username,
    name: name,
    role: session.Role,
    department: session.Department,
    isAdmin: isAdmin_(session),
    permissions: permissions
  };
}

/* =========================================================================
   STOCK MANAGEMENT BRIDGE

   Every stock movement goes through Stock Management's API so its lock and
   its no-negative-stock guard apply. Three things make that safe across a
   network that drops:

   - Each post carries a clientTxnId built from what it represents (the
     request, or the batch and BOM line), so it is the same ID every time
     that movement is sent. Stock Management recognises a repeat and
     returns the original row instead of writing a second one.
   - A reply that can't be read is treated as unknown, not failed. Apps
     Script answers some requests with a redirect, and a redirected POST
     arrives as a GET, which its doGet refuses — that says nothing about
     whether the write happened.
   - An unknown outcome is settled by asking (checkTransaction), never by
     posting again and hoping.
   ========================================================================= */

function postToStock_(body) {
  var res = UrlFetchApp.fetch(STOCK_API_URL, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(body),
    muteHttpExceptions: true,
    followRedirects: true
  });
  var text = res.getContentText();
  var json = null;
  try { json = JSON.parse(text); } catch (err) { json = null; }
  return { code: res.getResponseCode(), json: json, text: text };
}

function callStockApi_(action, params, clientTxnId) {
  var body = Object.assign({ action: action }, params);
  if (clientTxnId) body.clientTxnId = clientTxnId;

  for (var attempt = 0; attempt < 2; attempt++) {
    var res = postToStock_(body);

    if (res.json && res.json.success) return res.json.data;

    // A refusal that came back intact: Stock Management ran and said no.
    // Nothing was written, so report it as-is.
    if (res.json && res.json.error && !/POST requests only/i.test(res.json.error)) {
      throw new Error('Stock Management: ' + res.json.error);
    }

    // Anything else is an unknown outcome. Ask whether it landed.
    if (clientTxnId && params.token) {
      var check = postToStock_({ action: 'checkTransaction', token: params.token, clientTxnId: clientTxnId });
      if (check.json && check.json.success && check.json.data && check.json.data.found) {
        var found = check.json.data;
        found.duplicate = true;
        return found;
      }
    }
  }

  throw new Error('Stock Management could not be reached, and the movement was not recorded. ' +
                  'Nothing has been changed — try again once the connection is back.');
}

/* =========================================================================
   SHEET HELPERS

   Every getValue/setValue/getValues is a round trip, and they are what a
   request's time is actually spent on. Header rows are read once per
   execution, rows are read whole rather than cell by cell, and the cells
   that change are written back in as few ranges as possible.

   These mirror Stock Management's helpers rather than sharing a library —
   one dependency isn't worth it for a handful of stable functions. They
   are independent copies: changing one here does not change that one.
   ========================================================================= */

function jsonResponse_(payload) {
  return ContentService.createTextOutput(JSON.stringify(payload)).setMimeType(ContentService.MimeType.JSON);
}

// Keyed by spreadsheet as well as sheet, because this project reads two
// spreadsheets and a sheet ID is only unique within its own file.
var _sheetMetaCache_ = {};
function sheetMeta_(sheet) {
  var key = sheet.getParent().getId() + '!' + sheet.getSheetId();
  if (_sheetMetaCache_[key]) return _sheetMetaCache_[key];
  var lastCol = sheet.getLastColumn();
  var headers = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
  var map = {};
  headers.forEach(function (h, i) { if (h) map[h] = i; });
  var meta = { map: map, headers: headers, lastCol: lastCol };
  _sheetMetaCache_[key] = meta;
  return meta;
}

function invalidateSheetMeta_(sheet) {
  delete _sheetMetaCache_[sheet.getParent().getId() + '!' + sheet.getSheetId()];
}

function headerIndexMap_(sheet) { return sheetMeta_(sheet).map; }

// Finds a row by a key column and returns it whole, so the caller doesn't
// pay a round trip per cell. Returns null when there's no match.
function readRowByValue_(sheet, headerName, value) {
  if (!sheet) return null;
  var meta = sheetMeta_(sheet);
  if (meta.map[headerName] === undefined) return null;
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;

  var keys = sheet.getRange(2, meta.map[headerName] + 1, lastRow - 1, 1).getValues();
  var target = String(value);
  for (var i = 0; i < keys.length; i++) {
    if (String(keys[i][0]) !== target) continue;
    var rowIdx = i + 2;
    return {
      rowIdx: rowIdx,
      map: meta.map,
      headers: meta.headers,
      lastCol: meta.lastCol,
      values: sheet.getRange(rowIdx, 1, 1, meta.lastCol).getValues()[0]
    };
  }
  return null;
}

// Turns a row read by readRowByValue_ into a plain object.
function rowToObject_(row) {
  var obj = {};
  row.headers.forEach(function (h, c) { if (h) obj[h] = row.values[c]; });
  return obj;
}

// Writes several cells of one row, merging adjacent columns into a single
// range write. Columns that aren't changing are never rewritten, so a
// formula in a neighbouring column survives.
// Sheets reads a cell that starts with = + - or @ as a formula. Text typed
// by someone — a remark, a reference — that happens to start that way would
// become #ERROR!, or a formula running inside the spreadsheet. A leading
// apostrophe tells Sheets "this is text"; Sheets drops it, so reading the
// cell back gives exactly what was written.
function sheetSafe_(v) {
  return (typeof v === 'string' && /^[=+\-@]/.test(v)) ? "'" + v : v;
}

// A form sends a date as "2026-09-24". new Date() reads that as midnight
// UTC, which Sheets shows in India as 05:30. Built from its parts, it is
// midnight in the script's own time zone.
function parseDay_(v) {
  var m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(v || '').trim());
  return m ? new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])) : new Date(v);
}

function writeRowCells_(sheet, rowIdx, changes) {
  var cols = Object.keys(changes).map(Number).sort(function (a, b) { return a - b; });
  var i = 0;
  while (i < cols.length) {
    var startAt = i;
    while (i + 1 < cols.length && cols[i + 1] === cols[i] + 1) i++;
    var run = cols.slice(startAt, i + 1);
    sheet.getRange(rowIdx, run[0] + 1, 1, run.length)
         .setValues([run.map(function (c) { return sheetSafe_(changes[c]); })]);
    i++;
  }
}

// Builds the { columnIndex: value } map writeRowCells_ wants from
// { header: value }, skipping headers the sheet doesn't have.
function stageCells_(map, fields) {
  var changes = {};
  Object.keys(fields).forEach(function (header) {
    if (map[header] !== undefined) changes[map[header]] = fields[header];
  });
  return changes;
}

function readSheetAsObjects_(sheet) {
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var meta = sheetMeta_(sheet);
  var values = sheet.getRange(2, 1, lastRow - 1, meta.lastCol).getValues();
  return values.map(function (row) {
    var obj = {};
    meta.headers.forEach(function (h, i) { if (h) obj[h] = row[i]; });
    return obj;
  }).filter(function (obj) {
    return Object.keys(obj).some(function (k) { return obj[k] !== '' && obj[k] !== null; });
  });
}

// Appends a row built from a { header: value } object. Keys with no
// matching header are dropped, so callers must ensureColumns_ first when
// a field is new.
function appendRowByHeader_(sheet, rowObj) {
  var meta = sheetMeta_(sheet);
  var row = new Array(meta.lastCol).fill('');
  Object.keys(rowObj).forEach(function (key) {
    if (meta.map[key] !== undefined) row[meta.map[key]] = sheetSafe_(rowObj[key]);
  });
  sheet.appendRow(row);
}

// Adds any columns the sheet doesn't have, to the right of the existing
// headers. A no-op once they exist, so it's safe on every request.
function ensureColumns_(sheet, names) {
  var meta = sheetMeta_(sheet);
  var missing = names.filter(function (n) { return meta.map[n] === undefined; });
  if (missing.length === 0) return;
  sheet.getRange(1, meta.lastCol + 1, 1, missing.length).setValues([missing]);
  invalidateSheetMeta_(sheet);
}

function findRowIndexByValue_(sheet, header, value) {
  var row = readRowByValue_(sheet, header, value);
  return row ? row.rowIdx : -1;
}

function getOrCreateSheet_(spreadsheet, name, headers) {
  var sheet = spreadsheet.getSheetByName(name);
  if (sheet) return sheet;
  sheet = spreadsheet.insertSheet(name);
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeader_(sheet, headers.length);
  invalidateSheetMeta_(sheet);
  return sheet;
}

function styleHeader_(sheet, numCols) {
  sheet.getRange(1, 1, 1, numCols)
       .setFontWeight('bold').setBackground('#0f172a').setFontColor('#f8fafc');
  sheet.setFrozenRows(1);
}

// The time to the second, then six random hex characters. Three random
// digits gave two records in the same second a 1-in-900 chance of sharing
// an ID; this makes it about 1 in 16 million.
function generateId_(prefix) {
  var ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMddHHmmss');
  var rand = Utilities.getUuid().replace(/-/g, '').slice(0, 6);
  return prefix + '-' + ts + '-' + rand;
}

// Scaling a BOM line by a batch quantity turns 0.1 x 3 into
// 0.30000000000000004. Posting that would carry the noise into Stock
// Management's ledger and its balances, so quantities are settled to six
// decimals — well beyond anything weighed on the floor — before they go.
function roundQty_(n) {
  return Math.round((Number(n) || 0) * 1e6) / 1e6;
}

function isActiveValue_(val) {
  if (val === false) return false;
  if (typeof val === 'string' && val.toUpperCase() === 'FALSE') return false;
  return true;
}

/* =========================================================================
   CODE GENERATION

   Every physical thing — MPP batch, Element, Order, request — takes its
   code from here rather than each module inventing its own numbering, so
   two departments creating one in the same second can't collide.

   No lock of its own: writes already hold the request lock (see
   withWriteLock_), and taking a second one inside it would be at best
   redundant. ADMIN_ functions run by hand from the editor are the only
   callers outside that lock, and they are single-user by nature.
   ========================================================================= */

var CODE_SEQUENCES_HEADERS = ['Prefix', 'LastNumber'];

// generateCode_('MPP') -> 'MPP-00001'
function generateCode_(prefix, pad) {
  pad = pad || 5;
  var sheet = getOrCreateSheet_(SpreadsheetApp.getActiveSpreadsheet(), 'Code_Sequences', CODE_SEQUENCES_HEADERS);
  var row = readRowByValue_(sheet, 'Prefix', prefix);
  var next;
  if (!row) {
    next = 1;
    appendRowByHeader_(sheet, { Prefix: prefix, LastNumber: next });
  } else {
    next = (Number(row.values[row.map['LastNumber']]) || 0) + 1;
    sheet.getRange(row.rowIdx, row.map['LastNumber'] + 1).setValue(next);
  }
  return prefix + '-' + ('0000000000' + next).slice(-pad);
}

/* =========================================================================
   SETUP

   Run ADMIN_setupSheet() once after pasting this file into a fresh
   spreadsheet. Safe to run again later — existing tabs are left alone, so
   it also serves to add a tab introduced by a later change.

   Deliberately not created here: Users, Departments, RM_Master,
   RM_Stock_Ledger, FG_Master, FG_Stock_Ledger. Those live in the Stock
   Management sheet and are read and written there.
   ========================================================================= */

var TAB_SCHEMA = {
  'Workflow': ['FromDept', 'ToDept', 'Kind'],
  'Handovers': ['HandoverID', 'Date', 'Kind', 'Direction', 'FromDept', 'ToDept', 'ItemCode', 'ItemName', 'Qty', 'Unit',
                'Status', 'RequestedBy', 'SentBy', 'SentDate', 'ReceivedBy', 'ReceivedDate', 'ClosedBy', 'ClosedDate',
                'CloseReason', 'OrderRef', 'Reference', 'Remarks', 'StockTxnId'],
  'WIP_Entries': ['EntryID', 'Date', 'Dept', 'ItemCode', 'Type', 'Qty', 'By', 'Remarks'],
  'Attachments': ['AttachmentID', 'HandoverID', 'FileID', 'FileName', 'MimeType', 'Size', 'Thumb', 'Dept', 'UploadedBy',
                  'UploadedAt', 'Status', 'RemovedBy', 'RemovedAt', 'RemoveReason', 'ClientTxnId', 'LinkedTo'],
  // Kept for the move-over and for history; nothing new is written here.
  'Internal_Requests': ['ReqID', 'Date', 'Type', 'FromDept', 'ToDept', 'ItemCode', 'ItemName', 'Qty', 'Status', 'OrderRef', 'RequestedBy', 'ApprovedBy', 'IssuedDate', 'ReceivedDate', 'MigratedTo'],
  'ELE_BOM': ['iVersion', 'BOM', 'Input/Output', 'ProductName', 'ProductCode', 'UnitName', 'Qty', 'Rate', 'Gross'],
  'FG_BOM': ['iVersion', 'BOM', 'Input/Output', 'ProductName', 'ProductCode', 'UnitName', 'Qty', 'Rate', 'Gross'],
  'Orders': ['OrderID', 'Date', 'Customer', 'FGCode', 'FGName', 'Qty', 'Source', 'PlacedBy', 'Status', 'DueDate'],
  'Elements': ['ElementCode', 'SlitMPPCode', 'WindingDate', 'WindMachine', 'Customer', 'DrawingRef', 'Spec', 'SprayDate', 'TestResult', 'CurrentStage', 'CurrentStatus'],
  'Element_Stage_Log': ['LogID', 'ElementCode', 'Stage', 'Date', 'Operator', 'Result', 'Notes'],
  'Code_Sequences': CODE_SEQUENCES_HEADERS,
  'Idempotency': ['ClientTxnId', 'Action', 'User', 'Timestamp', 'Result'],
  'AuditLog': ['Timestamp', 'User', 'Action', 'Module', 'RecordID', 'OldValue', 'NewValue']
};

// Who changed what, and what it was before. Created on first use, so a
// sheet set up before the log existed gains it without a setup run.
function logAudit_(user, action, module, recordId, oldValue, newValue) {
  try {
    var sheet = getOrCreateSheet_(SpreadsheetApp.getActiveSpreadsheet(), 'AuditLog', TAB_SCHEMA['AuditLog']);
    appendRowByHeader_(sheet, {
      Timestamp: new Date(), User: user, Action: action, Module: module, RecordID: recordId,
      OldValue: oldValue === undefined ? '' : oldValue, NewValue: newValue === undefined ? '' : newValue
    });
  } catch (err) {
    // The log never blocks the work it records.
    Logger.log('Audit log write failed: ' + err.message);
  }
}

function ADMIN_setupSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var created = [];
  var existing = [];

  Object.keys(TAB_SCHEMA).forEach(function (name) {
    if (ss.getSheetByName(name)) {
      // Tabs made before a column was added get it now, rather than
      // silently dropping that field on every write.
      ensureColumns_(ss.getSheetByName(name), TAB_SCHEMA[name]);
      existing.push(name);
      return;
    }
    getOrCreateSheet_(ss, name, TAB_SCHEMA[name]);
    created.push(name);
  });

  // The route starts as the standard one; after that it is edited in the
  // sheet, and never overwritten here.
  var wf = ss.getSheetByName('Workflow');
  if (wf && wf.getLastRow() < 2) {
    DEFAULT_WORKFLOW.forEach(function (l) { appendRowByHeader_(wf, { FromDept: l[0], ToDept: l[1], Kind: l[2] }); });
  }

  // The Sheet1 every new spreadsheet comes with is clutter once the real
  // tabs exist.
  var sheet1 = ss.getSheetByName('Sheet1');
  if (sheet1 && ss.getSheets().length > 1 && sheet1.getLastRow() === 0) ss.deleteSheet(sheet1);

  Logger.log('Created: ' + (created.join(', ') || '(none — all tabs already existed)'));
  Logger.log('Already existed: ' + (existing.join(', ') || '(none)'));
  return { created: created, existing: existing };
}

// Quick check from the editor that this script can see Stock Management.
// Run it after a fresh install, before pointing the frontend at it.
function ADMIN_checkStockLink() {
  var ss = getStockSpreadsheet_();
  var tabs = ['Sessions', 'Users', 'RM_Master', 'FG_Master'];
  var missing = tabs.filter(function (t) { return !ss.getSheetByName(t); });
  var report = {
    stockSheet: ss.getName(),
    missingTabs: missing,
    api: null
  };
  var ping = postToStock_({ action: 'ping' });
  report.api = ping.json && ping.json.success ? 'reachable'
    : ping.json ? 'reachable, but too old to answer ping — deploy Stock Management 3.16 or later'
    : 'NOT reachable — check STOCK_API_URL and that the deployment is set to Anyone';
  Logger.log(JSON.stringify(report, null, 2));
  return report;
}

/* =========================================================================
   ORDERS

   Placing an order is permission-gated to Marketing, Dispatch or Admin.
   Viewing the list is open to anyone logged in — what's in the pipeline is
   useful planning context for every department, not just the ones that can
   create orders.
   ========================================================================= */

var ORDER_SOURCES = ['Mail', 'WhatsApp', 'Call', 'Other'];

function placeOrder_(session, params) {
  requirePerm_(session, 'pt.orders.place');

  var customer = String(params.customer || '').trim();
  var fgCode = String(params.fgCode || '').trim();
  var qty = Number(params.qty);
  var source = String(params.source || '').trim();

  if (!customer) throw new Error('Customer is required');
  if (!fgCode) throw new Error('FG code is required');
  if (!qty || isNaN(qty) || qty <= 0) throw new Error('Qty must be a positive number');
  if (ORDER_SOURCES.indexOf(source) === -1) throw new Error('Source must be one of: ' + ORDER_SOURCES.join(', '));

  // Checked against Stock Management's live FG_Master rather than trusting
  // what the client sent — the form's dropdown comes from the same place,
  // but the API is reachable directly, so this is the real boundary. Also
  // captures the name as it is today, so a later rename doesn't rewrite
  // the history of this order.
  var fgRow = readRowByValue_(getStockSpreadsheet_().getSheetByName('FG_Master'), 'FG_Code', fgCode);
  if (!fgRow) throw new Error('Unknown FG code: ' + fgCode);
  var fg = rowToObject_(fgRow);
  if (!isActiveValue_(fg.Active)) throw new Error('FG code is inactive: ' + fgCode);

  var ordersSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Orders');
  ensureColumns_(ordersSheet, ['FGName']);

  var rowObj = {
    OrderID: generateCode_('ORD'),
    Date: new Date(),
    Customer: customer,
    FGCode: fgCode,
    FGName: fg.FG_Name || '',
    Qty: qty,
    Source: source,
    PlacedBy: session.Username,
    Status: 'Pending',
    DueDate: params.dueDate ? parseDay_(params.dueDate) : ''
  };
  appendRowByHeader_(ordersSheet, rowObj);
  logAudit_(session.Username, 'PLACE_ORDER', 'Orders', rowObj.OrderID, '',
            JSON.stringify({ customer: rowObj.Customer, fgCode: rowObj.FGCode, qty: rowObj.Qty }));
  return rowObj;
}

// An order used to stay Pending for ever: nothing moved it on. Whoever can
// place orders can set its status by hand, and a dispatch received against
// it sets it on its own (updateOrderFromDispatch_).
var ORDER_STATUSES = ['Pending', 'In production', 'Part dispatched', 'Dispatched', 'Cancelled'];

function updateOrderStatus_(session, params) {
  requirePerm_(session, 'pt.orders.place');
  var id = String(params.orderId || '').trim();
  var status = String(params.status || '').trim();
  if (ORDER_STATUSES.indexOf(status) === -1) throw new Error('Status must be one of: ' + ORDER_STATUSES.join(', '));
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Orders');
  var row = readRowByValue_(sheet, 'OrderID', id);
  if (!row) throw new Error('Order not found: ' + id);
  var before = String(row.values[row.map['Status']] || '');
  if (before === status) return { orderId: id, status: status, alreadyDone: true };
  writeRowCells_(sheet, row.rowIdx, stageCells_(row.map, { Status: status }));
  logAudit_(session.Username, 'ORDER_STATUS', 'Orders', id, before, status);
  return { orderId: id, status: status };
}

// A finished-goods handover out of Stores that carries an order's ID moves
// that order on when it's received: Dispatched once everything received
// against it covers the order's quantity, Part dispatched before that. A
// cancelled order is left alone.
function updateOrderFromDispatch_(h, session) {
  if (h.Kind !== 'FG' || h.FromDept !== STOCK_HOLDER || h.Direction !== 'Forward' || !h.OrderRef) return;
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Orders');
  var row = sheet ? readRowByValue_(sheet, 'OrderID', String(h.OrderRef)) : null;
  if (!row) return;
  var order = rowToObject_(row);
  if (order.Status === 'Cancelled') return;
  var sent = 0;
  readHandovers_().forEach(function (x) {
    if (x.Kind === 'FG' && x.FromDept === STOCK_HOLDER && x.Direction === 'Forward' &&
        x.Status === 'Received' && String(x.OrderRef) === String(h.OrderRef)) sent += Number(x.Qty) || 0;
  });
  var status = sent + 1e-9 >= (Number(order.Qty) || 0) ? 'Dispatched' : 'Part dispatched';
  if (status === order.Status) return;
  writeRowCells_(sheet, row.rowIdx, stageCells_(row.map, { Status: status }));
  logAudit_(session.Username, 'ORDER_STATUS', 'Orders', order.OrderID, order.Status, status + ' (' + h.HandoverID + ')');
}

function getOrders_() {
  var orders = readSheetAsObjects_(SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Orders'));
  orders.sort(function (a, b) { return new Date(b.Date).getTime() - new Date(a.Date).getTime(); });
  return orders;
}

// Active FG items, trimmed to what the place-order search needs — no
// reason to ship every BOM and label field to populate one dropdown.
function getFGCatalog_() {
  return readSheetAsObjects_(getStockSpreadsheet_().getSheetByName('FG_Master'))
    .filter(function (r) { return isActiveValue_(r.Active); })
    .map(function (r) { return { code: String(r.FG_Code), name: r.FG_Name, unit: r.Unit || '' }; });
}

/* =========================================================================
   WORKFLOW

   The route material takes, one handover at a time:

     Purchase → Stores → Winding → Spray → Short Clearance → Assembly →
     Stores (finished goods) → Dispatch

   Every step can also run backwards, as a return. The route lives in the
   Workflow tab — one row per link, From / To / Kind — so a department
   renamed or added is a sheet edit, not a code change. Kind says what
   travels on the link:

     RM   raw material, held in Stock Management; the handover posts to it
     WIP  elements between Winding and Assembly, counted here by spec
     FG   finished goods, held in Stock Management; the handover posts to it
   ========================================================================= */

// Where raw material and finished goods sit in Stock Management. Links into
// or out of this department are the ones that move Stock Management stock.
var STOCK_HOLDER = 'Stores';
var WORKFLOW_KINDS = ['RM', 'WIP', 'FG'];

var DEFAULT_WORKFLOW = [
  ['Purchase',        'Stores',          'RM'],
  ['Stores',          'Winding',         'RM'],
  ['Winding',         'Spray',           'WIP'],
  ['Spray',           'Short Clearance', 'WIP'],
  ['Short Clearance', 'Assembly-MFD',    'WIP'],
  ['Short Clearance', 'Assembly-kVAr',   'WIP'],
  ['Assembly-MFD',    'Stores',          'FG'],
  ['Assembly-kVAr',   'Stores',          'FG'],
  ['Stores',          'Dispatch',        'FG']
];

var _workflow = null;
function readWorkflow_() {
  if (_workflow) return _workflow;
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Workflow');
  var links = (sheet ? readSheetAsObjects_(sheet) : []).map(function (r) {
    return { from: String(r.FromDept || '').trim(), to: String(r.ToDept || '').trim(), kind: String(r.Kind || '').trim().toUpperCase() };
  }).filter(function (l) { return l.from && l.to && WORKFLOW_KINDS.indexOf(l.kind) > -1; });
  // Until the tab is set up, the standard route applies.
  if (!links.length) links = DEFAULT_WORKFLOW.map(function (l) { return { from: l[0], to: l[1], kind: l[2] }; });
  _workflow = links;
  return links;
}

// The link a handover travels: forward along the workflow, or back along
// it as a return. Anything else isn't a step in the route.
function findLink_(from, to) {
  var links = readWorkflow_();
  for (var i = 0; i < links.length; i++) {
    if (links[i].from === from && links[i].to === to) return { kind: links[i].kind, direction: 'Forward' };
  }
  for (var k = 0; k < links.length; k++) {
    if (links[k].from === to && links[k].to === from) return { kind: links[k].kind, direction: 'Return' };
  }
  throw new Error(from + ' doesn\u2019t hand over to ' + to + ' \u2014 only neighbouring departments in the workflow do.');
}

// What a department does with elements, read off the route: the first
// element department makes them from raw material (Winding); one that turns
// them into finished goods uses them up (Assembly); any that has elements
// passing through can scrap them.
function wipRoles_(dept) {
  var links = readWorkflow_();
  var has = function (fn) { return links.some(fn); };
  var inWip = has(function (l) { return l.to === dept && l.kind === 'WIP'; });
  var outWip = has(function (l) { return l.from === dept && l.kind === 'WIP'; });
  var inRm = has(function (l) { return l.to === dept && l.kind === 'RM'; });
  var outFg = has(function (l) { return l.from === dept && l.kind === 'FG'; });
  return { produces: outWip && inRm, uses: inWip && outFg, holds: inWip || outWip };
}

function workflowDepartments_() {
  var seen = [];
  readWorkflow_().forEach(function (l) {
    [l.from, l.to].forEach(function (d) { if (seen.indexOf(d) === -1) seen.push(d); });
  });
  return seen;
}

/* =========================================================================
   ITEMS

   Raw material and finished goods come from Stock Management's masters.
   Elements are counted by spec — the specs Design defines in ELE_BOM — and
   Winding can start a new one when it produces the first of them.
   ========================================================================= */

function getRMCatalog_() {
  return readSheetAsObjects_(getStockSpreadsheet_().getSheetByName('RM_Master'))
    .filter(function (r) { return isActiveValue_(r.Active); })
    .map(function (r) { return { code: String(r.RM_Code), name: r.RM_Name, unit: r.Unit || '' }; });
}

function getElementSpecs_() {
  var specs = [];
  var add = function (name) {
    name = String(name || '').replace(/\s+/g, ' ').trim();
    if (name && !specs.some(function (x) { return x.toLowerCase() === name.toLowerCase(); })) specs.push(name);
  };
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  readSheetAsObjects_(ss.getSheetByName('ELE_BOM')).forEach(function (r) { add(r.BOM); });
  // Specs Winding has started since, so they can travel on down the line.
  readSheetAsObjects_(ss.getSheetByName('WIP_Entries')).forEach(function (r) { add(r.ItemCode); });
  return specs.sort().map(function (n) { return { code: n, name: n, unit: 'NOS' }; });
}

function itemsFor_(kind) {
  if (kind === 'RM') return getRMCatalog_();
  if (kind === 'FG') return getFGCatalog_();
  if (kind === 'WIP') return getElementSpecs_();
  throw new Error('Unknown kind: ' + kind);
}

function findItem_(kind, code) {
  code = String(code || '').trim();
  if (!code) throw new Error('Pick the item to hand over');
  var hit = itemsFor_(kind).filter(function (x) { return String(x.code).toLowerCase() === code.toLowerCase(); })[0];
  if (!hit) throw new Error((kind === 'WIP' ? 'Unknown element spec: ' : 'Unknown or inactive item: ') + code);
  return hit;
}

function positiveQty_(v) {
  var qty = roundQty_(Number(v));
  if (!qty || isNaN(qty) || qty <= 0) throw new Error('Quantity must be more than zero');
  return qty;
}

/* =========================================================================
   HANDOVERS

   One record per movement between neighbouring departments:

     Requested  the receiving side asked for it (optional)
     Sent       the sending side has sent it
     Received   the receiving side confirmed it — the moment custody
                changes, and so the moment Stock Management is posted
     Rejected / Declined / Cancelled   closed without moving anything

   Stock Management is posted only on receipt, by the person receiving, to
   their own department, with a fixed ID per handover — so a lost reply
   can't post twice, and a rejected handover never posts at all.
   ========================================================================= */

function handoversSheet_() {
  return getOrCreateSheet_(SpreadsheetApp.getActiveSpreadsheet(), 'Handovers', TAB_SCHEMA['Handovers']);
}

// Read once per execution and reused. A send, a WIP entry and a screen
// load each used to read the whole tab two or three times (on-hand totals,
// the list, the stock check). Every write below clears it.
var _handoversCache_ = null;
function readHandovers_() {
  if (!_handoversCache_) _handoversCache_ = readSheetAsObjects_(SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Handovers'));
  return _handoversCache_;
}

// A handover moved to the archive (ADMIN_archiveHandovers) is still found,
// so its files and history keep opening. Archived ones are closed, so
// nothing writes to them.
function findHandover_(id) {
  var sheet = handoversSheet_();
  var key = String(id || '').trim();
  var row = readRowByValue_(sheet, 'HandoverID', key);
  if (!row) {
    var archive = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HANDOVER_ARCHIVE_TAB);
    var old = archive ? readRowByValue_(archive, 'HandoverID', key) : null;
    if (old) return { sheet: archive, row: old, h: rowToObject_(old), archived: true };
    throw new Error('Handover not found: ' + id);
  }
  return { sheet: sheet, row: row, h: rowToObject_(row) };
}

function updateHandover_(found, fields) {
  if (found.archived) throw new Error(found.h.HandoverID + ' is archived and can\u2019t be changed.');
  writeRowCells_(found.sheet, found.row.rowIdx, stageCells_(found.row.map, fields));
  _handoversCache_ = null;
}

function newHandover_(session, link, from, to, item, qty, status, params) {
  var now = new Date();
  var rec = {
    HandoverID: generateCode_('HO', 5), Date: now, Kind: link.kind, Direction: link.direction,
    FromDept: from, ToDept: to, ItemCode: item.code, ItemName: item.name, Qty: qty, Unit: item.unit || '',
    Status: status, OrderRef: String(params.orderRef || '').trim(),
    Reference: String(params.reference || '').trim(), Remarks: String(params.remarks || '').trim()
  };
  if (status === 'Requested') rec.RequestedBy = session.Username;
  if (status === 'Sent') { rec.SentBy = session.Username; rec.SentDate = now; }
  appendRowByHeader_(handoversSheet_(), rec);
  _handoversCache_ = null;
  return rec;
}

// The receiving department asks its neighbour to send.
function requestHandover_(session, params) {
  var from = String(params.fromDept || '').trim(), to = String(params.toDept || '').trim();
  requirePerm_(session, 'pt.handover.receive');
  requireDept_(session, to);
  var link = findLink_(from, to);
  var rec = newHandover_(session, link, from, to, findItem_(link.kind, params.itemCode), positiveQty_(params.qty), 'Requested', params);
  logAudit_(session.Username, 'REQUEST', 'Handovers', rec.HandoverID, '', describeHandover_(rec));
  return rec;
}

// A fresh send, or the answer to a request (handoverId).
function sendHandover_(session, params) {
  requirePerm_(session, 'pt.handover.send');
  if (params.handoverId) {
    var found = findHandover_(params.handoverId), h = found.h;
    if (h.Status === 'Sent' || h.Status === 'Received') return { handoverId: h.HandoverID, status: h.Status, alreadyDone: true };
    if (h.Status !== 'Requested') throw new Error('This request is ' + h.Status.toLowerCase() + ' and can\u2019t be sent');
    requireDept_(session, h.FromDept);
    var qty = params.qty !== undefined && params.qty !== '' ? positiveQty_(params.qty) : Number(h.Qty);
    if (h.Kind === 'WIP') assertOnHand_(h.FromDept, h.ItemCode, qty);
    else assertStockAvailable_(h.Kind, h.FromDept, h.ItemCode, qty, h.HandoverID);
    var fields = { Status: 'Sent', Qty: qty, SentBy: session.Username, SentDate: new Date() };
    if (params.reference) fields.Reference = String(params.reference).trim();
    if (params.remarks) fields.Remarks = String(params.remarks).trim();
    updateHandover_(found, fields);
    logAudit_(session.Username, 'SEND', 'Handovers', h.HandoverID, 'Requested ' + h.Qty, 'Sent ' + qty);
    return { handoverId: h.HandoverID, status: 'Sent', qty: qty };
  }
  var from = String(params.fromDept || '').trim(), to = String(params.toDept || '').trim();
  requireDept_(session, from);
  var link = findLink_(from, to);
  var item = findItem_(link.kind, params.itemCode);
  var qty2 = positiveQty_(params.qty);
  if (link.kind === 'WIP') assertOnHand_(from, item.code, qty2);
  else assertStockAvailable_(link.kind, from, item.code, qty2, null);
  var rec = newHandover_(session, link, from, to, item, qty2, 'Sent', params);
  logAudit_(session.Username, 'SEND', 'Handovers', rec.HandoverID, '', describeHandover_(rec));
  return { handoverId: rec.HandoverID, status: 'Sent', qty: qty2 };
}

// Receiving can call Stock Management, which takes seconds. Holding this
// app's write lock through that made every other entry wait behind it, and
// after WRITE_LOCK_WAIT_MS fail with "the system is busy". So the lock is
// held for the checks, let go for the call, and taken again to record the
// result. Two people confirming the same handover at once are safe: the
// stock post carries a fixed ID per handover, so Stock Management writes it
// once, and whoever records second finds it already received.
function receiveHandover_(session, params) {
  var clientTxnId = String(params.clientTxnId || '').trim();

  var first = withWriteLock_(function () {
    if (clientTxnId) {
      var prior = findIdempotentResult_(clientTxnId);
      if (prior) {
        checkReplay_(prior, 'receiveHandover', session);
        var replay = prior.result;
        if (replay && typeof replay === 'object' && !Array.isArray(replay)) replay.duplicate = true;
        return { done: replay };
      }
    }
    requirePerm_(session, 'pt.handover.receive');
    var h = findHandover_(params.handoverId).h;
    // Safe to send twice: a second copy finds it already received.
    if (h.Status === 'Received') return { done: { handoverId: h.HandoverID, status: 'Received', alreadyDone: true } };
    if (h.Status !== 'Sent') throw new Error('Only a sent handover can be received (this one is ' + h.Status.toLowerCase() + ')');
    requireDept_(session, h.ToDept);
    var posts = (h.Kind === 'RM' || h.Kind === 'FG') && !String(h.StockTxnId || '');
    if (!posts) return { done: finishReceipt_(session, h.HandoverID, String(h.StockTxnId || ''), clientTxnId) };
    return { h: h };
  });
  if (first.done) return first.done;

  var h = first.h;
  var mv = stockMovement_(h);
  var result = callStockApi_(mv.action, Object.assign({ token: session.Token }, mv.params), 'pt:' + h.HandoverID + ':recv');
  var stockTxn = (result && (result.txnId || result.TxnID)) || 'posted';

  return withWriteLock_(function () { return finishReceipt_(session, h.HandoverID, stockTxn, clientTxnId); });
}

// Marks a handover received, under the write lock.
function finishReceipt_(session, handoverId, stockTxn, clientTxnId) {
  var found = findHandover_(handoverId), h = found.h;
  var out;
  if (h.Status === 'Received') {
    out = { handoverId: h.HandoverID, status: 'Received', stockTxnId: String(h.StockTxnId || stockTxn), alreadyDone: true };
  } else {
    if (h.Status !== 'Sent') {
      // Rejected or cancelled in the moment Stock Management was being
      // posted. The stock has already moved, so the receipt stands; the
      // log says what it overrode, for someone to look at.
      logAudit_(session.Username, 'RECEIVE_AFTER_CLOSE', 'Handovers', h.HandoverID, h.Status, 'Received · ' + stockTxn);
    }
    updateHandover_(found, { Status: 'Received', ReceivedBy: session.Username, ReceivedDate: new Date(), StockTxnId: stockTxn });
    var received = Object.assign({}, h, { Status: 'Received', StockTxnId: stockTxn });
    // The stock movement is already posted; a problem filing its paperwork
    // must not undo or fail the receipt. ADMIN_setupAttachments() finishes
    // any link left undone.
    try {
      linkToStock_(received, activeHandoverFiles_(h.HandoverID));
    } catch (err) {
      Logger.log('Linking files for ' + h.HandoverID + ' failed: ' + err.message);
    }
    try {
      updateOrderFromDispatch_(received, session);
    } catch (err) {
      Logger.log('Updating the order for ' + h.HandoverID + ' failed: ' + err.message);
    }
    logAudit_(session.Username, 'RECEIVE', 'Handovers', h.HandoverID, h.Status, 'Received' + (stockTxn ? ' · ' + stockTxn : ''));
    out = { handoverId: h.HandoverID, status: 'Received', stockTxnId: stockTxn };
  }
  if (clientTxnId) recordIdempotentResult_(clientTxnId, 'receiveHandover', session.Username, out);
  return out;
}

// Raw material and finished goods leaving Stores come out of Stock
// Management when the other side receives them — which used to be the first
// moment anyone found out Stores didn't have them: the receiver's confirm
// failed, and the handover sat as Sent with nobody able to clear it. Checked
// here too, against the balance Stock Management holds less what Stores has
// already sent and not yet had received, so the sender hears it instead.
// Stock Management's own guard is still the boundary; this only moves the
// message to the person who can act on it.
var STOCK_TABS = {
  RM: { sheet: 'RM_Current_Stock', code: 'RM_Code' },
  FG: { sheet: 'FG_Current_Stock', code: 'FG_Code' }
};

function assertStockAvailable_(kind, from, code, qty, exceptHandoverId) {
  if (from !== STOCK_HOLDER || !STOCK_TABS[kind]) return;
  var tab = STOCK_TABS[kind];
  var row = readRowByValue_(getStockSpreadsheet_().getSheetByName(tab.sheet), tab.code, code);
  var balance = row && row.map['CurrentBalance'] !== undefined ? Number(row.values[row.map['CurrentBalance']]) || 0 : 0;
  var inTransit = 0;
  readHandovers_().forEach(function (h) {
    if (h.Status !== 'Sent' || h.Kind !== kind || h.FromDept !== STOCK_HOLDER) return;
    if (String(h.ItemCode) !== String(code) || h.HandoverID === exceptHandoverId) return;
    if (h.StockTxnId) return; // already posted (moved over from a request)
    inTransit += Number(h.Qty) || 0;
  });
  var free = roundQty_(balance - inTransit);
  if (free + 1e-9 < qty) {
    throw new Error('Stores has ' + roundQty_(balance) + ' of ' + code + ' in Stock Management' +
                    (inTransit ? ', ' + roundQty_(inTransit) + ' of it already sent and waiting to be received' : '') +
                    ', so ' + qty + ' can\u2019t be sent. Post the missing receipt in Stock Management first.');
  }
}

// The movement a received handover makes in Stock Management, booked to the
// receiving department — the person confirming is posting for their own
// department, which is what their Stock Management access allows. The
// other side is named in the remarks.
function stockMovement_(h) {
  var qty = Number(h.Qty);
  var note = 'Production Tracker ' + h.HandoverID + ': ' + h.FromDept + ' \u2192 ' + h.ToDept + (h.Remarks ? ' \u2014 ' + h.Remarks : '');
  var base = { code: h.ItemCode, qty: qty, reference: String(h.OrderRef || h.Reference || h.HandoverID), remarks: note };
  var into = h.ToDept === STOCK_HOLDER;
  if (h.Kind === 'RM') {
    if (into) return { action: 'addRMTransaction', params: Object.assign(base, { txnType: h.Direction === 'Forward' ? 'Receipt' : 'Return', department: STOCK_HOLDER }) };
    return { action: 'addRMTransaction', params: Object.assign(base, { txnType: 'Issue', department: h.ToDept }) };
  }
  if (into) return { action: 'addFGTransaction', params: Object.assign(base, { txnType: h.Direction === 'Forward' ? 'Production' : 'Return', department: STOCK_HOLDER }) };
  if (h.Direction === 'Forward') return { action: 'addFGTransaction', params: Object.assign(base, { txnType: 'Dispatch', department: h.ToDept }) };
  // Finished goods going back to Assembly leave stock as a negative adjustment.
  return { action: 'addFGTransaction', params: Object.assign(base, { txnType: 'Adjustment', qty: -qty, department: h.ToDept }) };
}

// The receiver refuses something sent; the sender declines a request; either
// side withdraws what it started. None of these moves anything.
function closeHandover_(session, params, how) {
  var found = findHandover_(params.handoverId), h = found.h;
  var target = { reject: 'Rejected', decline: 'Declined', cancel: 'Cancelled' }[how];
  if (h.Status === target) return { handoverId: h.HandoverID, status: target, alreadyDone: true };
  if (how === 'reject') {
    if (h.Status !== 'Sent') throw new Error('Only a sent handover can be rejected');
    requirePerm_(session, 'pt.handover.receive'); requireDept_(session, h.ToDept);
  } else if (how === 'decline') {
    if (h.Status !== 'Requested') throw new Error('Only a request can be declined');
    requirePerm_(session, 'pt.handover.send'); requireDept_(session, h.FromDept);
  } else if (h.Status === 'Sent') {
    requirePerm_(session, 'pt.handover.send'); requireDept_(session, h.FromDept);
  } else if (h.Status === 'Requested') {
    requirePerm_(session, 'pt.handover.receive'); requireDept_(session, h.ToDept);
  } else {
    throw new Error('This handover is ' + h.Status.toLowerCase() + ' and can\u2019t be cancelled');
  }
  updateHandover_(found, { Status: target, ClosedBy: session.Username, ClosedDate: new Date(), CloseReason: String(params.reason || '').trim() });
  logAudit_(session.Username, target.toUpperCase(), 'Handovers', h.HandoverID, h.Status, target + (params.reason ? ': ' + params.reason : ''));
  return { handoverId: h.HandoverID, status: target };
}

function describeHandover_(h) {
  return h.Kind + ' ' + h.Direction + ' ' + h.FromDept + ' \u2192 ' + h.ToDept + ' · ' + h.ItemCode + ' × ' + h.Qty;
}

/* =========================================================================
   WORK IN PROGRESS

   Elements between Winding and Assembly aren't in Stock Management, so they
   are counted here. A department's elements on hand, per spec, are:

     received by handover + produced
     − sent on (in transit counts as gone, so it can't be sent twice)
     − scrapped − used

   Winding produces them from the raw material it received; Assembly uses
   them up making finished goods; anyone holding them can scrap them.
   ========================================================================= */

function wipOnHand_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var totals = {};
  var bump = function (dept, code, n) {
    if (!totals[dept]) totals[dept] = {};
    totals[dept][code] = roundQty_((totals[dept][code] || 0) + n);
  };
  readHandovers_().forEach(function (h) {
    if (String(h.Kind) !== 'WIP') return;
    var q = Number(h.Qty) || 0;
    if (h.Status === 'Received') bump(h.ToDept, h.ItemCode, q);
    if (h.Status === 'Sent' || h.Status === 'Received') bump(h.FromDept, h.ItemCode, -q);
  });
  readSheetAsObjects_(ss.getSheetByName('WIP_Entries')).forEach(function (e) {
    var q = Number(e.Qty) || 0;
    // Carried: what archived handovers had added or taken away, written by
    // ADMIN_archiveHandovers with its own sign.
    bump(e.Dept, e.ItemCode, e.Type === 'Produced' || e.Type === 'Carried' ? q : -q);
  });
  return totals;
}

function assertOnHand_(dept, code, qty) {
  var have = ((wipOnHand_()[dept] || {})[code]) || 0;
  if (have + 1e-9 < qty) {
    throw new Error('Only ' + have + ' of ' + code + ' on hand in ' + dept + ' \u2014 can\u2019t move ' + qty + '.');
  }
}

function recordWIP_(session, params) {
  requirePerm_(session, 'pt.wip.record');
  var dept = String(params.dept || session.Department || '').trim();
  requireDept_(session, dept);
  var type = String(params.type || '').trim();
  var roles = wipRoles_(dept);
  if (type === 'Produced' && !roles.produces) throw new Error(dept + ' doesn\u2019t make elements \u2014 they reach it by handover.');
  if (type === 'Used' && !roles.uses) throw new Error(dept + ' doesn\u2019t use elements up.');
  if (type === 'Scrapped' && !roles.holds) throw new Error(dept + ' doesn\u2019t hold elements.');
  if (['Produced', 'Scrapped', 'Used'].indexOf(type) === -1) throw new Error('Type must be Produced, Scrapped or Used');

  var qty = positiveQty_(params.qty);
  var code = String(params.itemCode || '').replace(/\s+/g, ' ').trim();
  if (!code) throw new Error('Pick the element spec');
  var specs = getElementSpecs_();
  var known = specs.filter(function (x) { return x.code.toLowerCase() === code.toLowerCase(); })[0];
  if (known) code = known.code;
  else {
    // Only the department that makes elements can start a spec nobody has yet.
    if (type !== 'Produced') throw new Error('Unknown element spec: ' + code);
    // A typed name that differs from an existing spec only in spaces, dots,
    // dashes or case is that spec mistyped. Started anyway, it would sit in
    // every dropdown for good.
    var squash = function (t) { return String(t).toLowerCase().replace(/[\s._\-\/]+/g, ''); };
    var near = specs.filter(function (x) { return squash(x.code) === squash(code); })[0];
    if (near) throw new Error('Did you mean ' + near.code + '? Pick it from the list rather than typing it.');
    // A genuinely new spec is a Manager's or Admin's call; specs normally
    // come from ELE_BOM.
    if (!isAdmin_(session) && !can_(session, 'pt.handover.any')) {
      throw new Error(code + ' isn\u2019t an element spec yet. Specs come from ELE_BOM \u2014 ask a Manager or Admin to start a new one.');
    }
  }
  if (type !== 'Produced') assertOnHand_(dept, code, qty);

  var rec = { EntryID: generateId_('WIP'), Date: new Date(), Dept: dept, ItemCode: code, Type: type, Qty: qty,
              By: session.Username, Remarks: String(params.remarks || '').trim() };
  appendRowByHeader_(getOrCreateSheet_(SpreadsheetApp.getActiveSpreadsheet(), 'WIP_Entries', TAB_SCHEMA['WIP_Entries']), rec);
  logAudit_(session.Username, 'WIP_' + type.toUpperCase(), 'WIP_Entries', rec.EntryID, '', dept + ' · ' + code + ' × ' + qty);
  return rec;
}

/* =========================================================================
   READING

   One call gives a department everything its screen needs: the route, its
   handovers either way, and its elements on hand. Someone who can see only
   their own department asking for another's is refused.
   ========================================================================= */

function getWorkflow_(session) {
  var depts = workflowDepartments_();
  var roles = {};
  depts.forEach(function (d) { roles[d] = wipRoles_(d); });
  return {
    links: readWorkflow_(), departments: depts, stockHolder: STOCK_HOLDER, roles: roles,
    myDept: session.Department,
    visible: depts.filter(function (d) { return canSee_(session, d); })
  };
}

function getItems_(params) {
  var kind = String(params.kind || '').trim().toUpperCase();
  return itemsFor_(kind);
}

function getHandovers_(session, params) {
  var dept = String(params.dept || session.Department || '').trim();
  if (!canSee_(session, dept)) throw new Error('You can only see your own department (' + session.Department + ').');
  var limit = Math.min(Number(params.limit) || 300, 1000);
  var rows = readHandovers_()
    .filter(function (h) { return h.FromDept === dept || h.ToDept === dept; });
  rows.sort(function (a, b) { return new Date(b.Date).getTime() - new Date(a.Date).getTime(); });
  var counts = handoverFileCounts_();
  rows.forEach(function (h) { h.AttachmentCount = counts[h.HandoverID] || 0; });
  return {
    dept: dept,
    handovers: rows.slice(0, limit),
    onHand: wipOnHand_()[dept] || {},
    roles: wipRoles_(dept)
  };
}

/* =========================================================================
   FILES ON HANDOVERS

   Photos and PDFs on any handover — an invoice with raw material from
   Purchase, a photo of rejected elements, a drawing with a request.
   Either department on the handover can add them, at any point until it's
   closed without moving.

   They're stored where Stock Management keeps its files — the "CALGAS
   Stock Attachments" folder, under "Production Tracker", a folder per
   month — never shared by link. When a handover is received and posts a
   Stock Management entry, its files are recorded on that entry too: the
   same files, not copies, so the invoice is on the Receipt in Stock
   Management's history. Files added after the receipt follow them there.
   ========================================================================= */

var ATTACH_MAX_FILES = 10;
var ATTACH_MAX_BYTES = 10 * 1024 * 1024;
var ATTACH_TYPES = { 'image/jpeg': 'jpg', 'image/png': 'png', 'image/webp': 'webp', 'image/gif': 'gif', 'application/pdf': 'pdf' };
var ATTACH_ROOT_NAME = 'CALGAS Stock Attachments';
var ATTACH_SUBFOLDER = 'Production Tracker';
// Stock Management's Attachments tab, which this app adds to when a
// handover's files go onto an entry. The same columns Stock Management
// itself uses (its ATTACH_HEADERS).
var SM_ATTACH_HEADERS = ['AttachmentID', 'Ledger', 'TxnID', 'FileID', 'FileName', 'MimeType', 'Size', 'Thumb', 'UploadedBy',
                         'UploadedAt', 'Status', 'RemovedBy', 'RemovedAt', 'RemoveReason', 'ClientTxnId', 'Source'];

function handoverFilesSheet_() {
  return getOrCreateSheet_(SpreadsheetApp.getActiveSpreadsheet(), 'Attachments', TAB_SCHEMA['Attachments']);
}

function readHandoverFiles_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Attachments');
  return sheet ? readSheetAsObjects_(sheet) : [];
}

function activeHandoverFiles_(handoverId) {
  return readHandoverFiles_().filter(function (a) { return a.HandoverID === handoverId && a.Status === 'Active'; });
}

function handoverFileCounts_() {
  var counts = {};
  readHandoverFiles_().forEach(function (a) { if (a.Status === 'Active') counts[a.HandoverID] = (counts[a.HandoverID] || 0) + 1; });
  return counts;
}

function publicHandoverFile_(a) {
  return {
    attachmentId: a.AttachmentID, handoverId: a.HandoverID, fileName: a.FileName, mimeType: a.MimeType,
    size: Number(a.Size) || 0, thumb: a.Thumb || '', dept: a.Dept, uploadedBy: a.UploadedBy, uploadedAt: a.UploadedAt,
    onStock: !!a.LinkedTo
  };
}

function cleanFileName_(name, mime) {
  var base = String(name || '').replace(/[\\/:*?"<>|\u0000-\u001f]/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 120);
  var ext = ATTACH_TYPES[mime];
  if (!base) base = 'file';
  if (!new RegExp('\\.' + ext + '$', 'i').test(base) && !(ext === 'jpg' && /\.jpe?g$/i.test(base))) {
    base = base.replace(/\.[A-Za-z0-9]{1,5}$/, '') + '.' + ext;
  }
  return base;
}

function formatBytes_(n) {
  return n >= 1048576 ? (Math.round(n / 104857.6) / 10) + ' MB' : Math.max(1, Math.round(n / 1024)) + ' KB';
}

// The shared attachments folder, found through Stock Management's Settings
// (made here if Stock Management hasn't stored a file yet, and recorded
// there so both apps use the one folder), then this app's folder in it,
// then the month's.
function attachmentsFolder_(when) {
  var settings = getStockSpreadsheet_().getSheetByName('Settings');
  var rows = readSheetAsObjects_(settings);
  var saved = rows.filter(function (r) { return r.Key === 'ATTACHMENTS_FOLDER_ID'; })[0];
  var root = null;
  if (saved && saved.Value) {
    try { root = DriveApp.getFolderById(String(saved.Value)); } catch (err) { root = null; }
  }
  if (!root) {
    root = DriveApp.createFolder(ATTACH_ROOT_NAME);
    // Stock Management may have made the folder in the same moment (this
    // script can't take its lock). Look again before saving ours: if it
    // got there first, use its folder and bin this one, so both apps keep
    // one folder.
    invalidateSheetMeta_(settings);
    var again = readSheetAsObjects_(settings).filter(function (r) { return r.Key === 'ATTACHMENTS_FOLDER_ID'; })[0];
    var theirs = null;
    if (again && again.Value && (!saved || String(again.Value) !== String(saved.Value))) {
      try { theirs = DriveApp.getFolderById(String(again.Value)); } catch (err) { theirs = null; }
    }
    if (theirs) {
      try { root.setTrashed(true); } catch (err) {}
      root = theirs;
    } else {
      var at = rows.indexOf(saved);
      if (at > -1) writeRowCells_(settings, at + 2, stageCells_(sheetMeta_(settings).map, { Value: root.getId() }));
      else appendRowByHeader_(settings, { Key: 'ATTACHMENTS_FOLDER_ID', Value: root.getId() });
    }
  }
  var sub = function (parent, name) {
    var it = parent.getFoldersByName(name);
    return it.hasNext() ? it.next() : parent.createFolder(name);
  };
  var month = Utilities.formatDate(when, Session.getScriptTimeZone(), 'yyyy-MM-dd').slice(0, 7);
  return sub(sub(root, ATTACH_SUBFOLDER), month);
}

// Either department on the handover, or someone who can act for any.
function requireHandoverParty_(session, h) {
  if (isAdmin_(session) || can_(session, 'pt.handover.any')) return;
  if (!can_(session, 'pt.handover.send') && !can_(session, 'pt.handover.receive')) {
    throw new Error('You don\u2019t have access to handovers. Ask an Admin to grant it in Team & Access.');
  }
  if (session.Department !== h.FromDept && session.Department !== h.ToDept) {
    throw new Error('Only ' + h.FromDept + ' and ' + h.ToDept + ' can add files to ' + h.HandoverID + '.');
  }
}

function isClosedWithoutMoving_(h) {
  return ['Rejected', 'Declined', 'Cancelled'].indexOf(String(h.Status)) > -1;
}

// One file per request. Checked under the lock, written to Drive outside
// it, recorded under the lock again — re-checking in case a retry of this
// same upload got there first.
function uploadHandoverFile_(session, params) {
  var id = String(params.handoverId || '').trim();
  var clientTxnId = String(params.clientTxnId || '').trim();
  var mime = String(params.mimeType || '').toLowerCase();
  if (!ATTACH_TYPES[mime]) throw new Error('Only photos (JPEG, PNG, WebP, GIF) and PDFs can be attached.');
  var data = String(params.data || '');
  if (!data) throw new Error('The file is empty.');
  if (Math.floor(data.length * 3 / 4) > ATTACH_MAX_BYTES * 1.01) {
    throw new Error('That file is larger than ' + formatBytes_(ATTACH_MAX_BYTES) + ', the most one file can be.');
  }
  var name = cleanFileName_(params.fileName, mime);
  var thumb = String(params.thumb || '');
  if (thumb.length > 45000 || !/^[A-Za-z0-9+\/=]*$/.test(thumb)) thumb = '';

  var sheet, folder, h;
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    sheet = handoverFilesSheet_();
    if (clientTxnId) {
      var prior = readHandoverFiles_().filter(function (a) { return String(a.ClientTxnId) === clientTxnId; })[0];
      if (prior) return publicHandoverFile_(prior);
    }
    h = findHandover_(id).h;
    if (isClosedWithoutMoving_(h)) throw new Error(h.HandoverID + ' was ' + h.Status.toLowerCase() + ', so files can\u2019t be added to it.');
    requireHandoverParty_(session, h);
    if (activeHandoverFiles_(h.HandoverID).length >= ATTACH_MAX_FILES) {
      throw new Error(h.HandoverID + ' already has ' + ATTACH_MAX_FILES + ' files, the most one handover can hold.');
    }
    folder = attachmentsFolder_(new Date());
  } finally {
    lock.releaseLock();
  }

  var bytes = Utilities.base64Decode(data);
  if (bytes.length > ATTACH_MAX_BYTES) {
    throw new Error('That file is larger than ' + formatBytes_(ATTACH_MAX_BYTES) + ', the most one file can be.');
  }
  var file = folder.createFile(Utilities.newBlob(bytes, mime, h.HandoverID + ' - ' + name));

  var rec;
  lock.waitLock(30000);
  try {
    if (clientTxnId) {
      var raced = readHandoverFiles_().filter(function (a) { return String(a.ClientTxnId) === clientTxnId; })[0];
      if (raced) {
        try { file.setTrashed(true); } catch (err) {}
        return publicHandoverFile_(raced);
      }
    }
    if (activeHandoverFiles_(h.HandoverID).length >= ATTACH_MAX_FILES) {
      try { file.setTrashed(true); } catch (err) {}
      throw new Error(h.HandoverID + ' already has ' + ATTACH_MAX_FILES + ' files, the most one handover can hold.');
    }
    rec = {
      AttachmentID: generateId_('ATT'), HandoverID: h.HandoverID, FileID: file.getId(), FileName: name, MimeType: mime,
      Size: bytes.length, Thumb: thumb, Dept: session.Department, UploadedBy: session.Username, UploadedAt: new Date(),
      Status: 'Active', ClientTxnId: clientTxnId
    };
    appendRowByHeader_(sheet, rec);
    // Already received: this one goes straight onto its Stock Management entry.
    var now = findHandover_(h.HandoverID).h;
    if (now.Status === 'Received') {
      try { linkToStock_(now, [rec]); } catch (err) { Logger.log('Linking ' + rec.AttachmentID + ' failed: ' + err.message); }
    }
  } finally {
    lock.releaseLock();
  }
  logAudit_(session.Username, 'ATTACH_FILE', 'Handovers', h.HandoverID, '', name + ' (' + formatBytes_(bytes.length) + ')');
  return publicHandoverFile_(rec);
}

// Recorded on the Stock Management entry the receipt posted — only for raw
// material and finished goods, and only when there's a real entry (one
// moved over from a request was posted at issue, under no entry of its own).
function linkToStock_(h, files) {
  var txnId = String(h.StockTxnId || '');
  if ((h.Kind !== 'RM' && h.Kind !== 'FG') || !/^(RMT|FGT)-/.test(txnId)) return 0;
  var pending = (files || []).filter(function (a) { return a.Status === 'Active' && !a.LinkedTo; });
  if (!pending.length) return 0;

  var ss = getStockSpreadsheet_();
  var smSheet = ss.getSheetByName('Attachments');
  if (!smSheet) {
    smSheet = ss.insertSheet('Attachments');
    smSheet.getRange(1, 1, 1, SM_ATTACH_HEADERS.length).setValues([SM_ATTACH_HEADERS]);
  } else {
    ensureColumns_(smSheet, ['Source']);
  }
  var already = {};
  readSheetAsObjects_(smSheet).forEach(function (r) { already[String(r.ClientTxnId)] = true; });
  var ptSheet = handoverFilesSheet_();
  var linked = 0;
  pending.forEach(function (a) {
    var key = 'pt:' + a.AttachmentID;
    if (!already[key]) {
      appendRowByHeader_(smSheet, {
        AttachmentID: generateId_('ATT'), Ledger: h.Kind, TxnID: txnId, FileID: a.FileID, FileName: a.FileName,
        MimeType: a.MimeType, Size: a.Size, Thumb: a.Thumb, UploadedBy: a.UploadedBy, UploadedAt: a.UploadedAt,
        Status: 'Active', ClientTxnId: key, Source: 'Production Tracker ' + h.HandoverID
      });
    }
    var row = readRowByValue_(ptSheet, 'AttachmentID', a.AttachmentID);
    if (row) writeRowCells_(ptSheet, row.rowIdx, stageCells_(row.map, { LinkedTo: h.Kind + ':' + txnId }));
    a.LinkedTo = h.Kind + ':' + txnId;
    linked++;
  });
  return linked;
}

function findHandoverFile_(attachmentId) {
  var id = String(attachmentId || '').trim();
  var a = readHandoverFiles_().filter(function (r) { return r.AttachmentID === id; })[0];
  if (!a) throw new Error('That file wasn\u2019t found.');
  return a;
}

// Whoever can see either department on the handover can see its files.
function requireSeesHandover_(session, h) {
  if (canSee_(session, h.FromDept) || canSee_(session, h.ToDept)) return;
  throw new Error('You can only see handovers of your own department (' + session.Department + ').');
}

function getHandoverFiles_(session, params) {
  var h = findHandover_(params.handoverId).h;
  requireSeesHandover_(session, h);
  return activeHandoverFiles_(h.HandoverID).map(publicHandoverFile_);
}

function getHandoverFile_(session, params) {
  var a = findHandoverFile_(params.attachmentId);
  if (a.Status !== 'Active') throw new Error('That file isn\u2019t available any more.');
  requireSeesHandover_(session, findHandover_(a.HandoverID).h);
  var blob = DriveApp.getFileById(String(a.FileID)).getBlob();
  return { attachmentId: a.AttachmentID, fileName: a.FileName, mimeType: a.MimeType, size: Number(a.Size) || 0,
           data: Utilities.base64Encode(blob.getBytes()) };
}

// Before the receipt, the department that added a file can take it off;
// after it, only someone who can act for any department, since by then it
// is also on a Stock Management entry. That entry keeps its copy of the
// record — removing it there is a Stock Management correction. The file
// itself stays in Drive either way.
function removeHandoverFile_(session, params) {
  var a = findHandoverFile_(params.attachmentId);
  if (a.Status === 'Removed') return { attachmentId: a.AttachmentID, alreadyDone: true };
  var h = findHandover_(a.HandoverID).h;
  var wide = isAdmin_(session) || can_(session, 'pt.handover.any');
  if (h.Status === 'Received') {
    if (!wide) throw new Error(h.HandoverID + ' has been received, so only a Manager or Admin can remove its files.');
  } else if (!wide) {
    requireHandoverParty_(session, h);
    if (session.Department !== a.Dept) throw new Error('Only ' + a.Dept + ', who added it, can remove this file.');
  }
  var sheet = handoverFilesSheet_();
  var row = readRowByValue_(sheet, 'AttachmentID', a.AttachmentID);
  writeRowCells_(sheet, row.rowIdx, stageCells_(row.map, {
    Status: 'Removed', RemovedBy: session.Username, RemovedAt: new Date(), RemoveReason: String(params.reason || '').trim()
  }));
  logAudit_(session.Username, 'REMOVE_FILE', 'Handovers', h.HandoverID, a.FileName, 'Removed' + (params.reason ? ': ' + params.reason : ''));
  return { attachmentId: a.AttachmentID, removed: true, stillOnStock: !!a.LinkedTo };
}

// Run once from the editor after deploying: it asks for Google Drive
// permission (the first time), makes the tab and folders, says where
// files will go, and finishes any receipt whose files weren't linked yet.
function ADMIN_setupAttachments() {
  handoverFilesSheet_();
  var folder = attachmentsFolder_(new Date());
  var relinked = 0;
  var byId = {};
  readHandoverFiles_().forEach(function (a) {
    if (a.Status === 'Active' && !a.LinkedTo) (byId[a.HandoverID] = byId[a.HandoverID] || []).push(a);
  });
  Object.keys(byId).forEach(function (id) {
    var h = findHandover_(id).h;
    if (h.Status === 'Received') relinked += linkToStock_(h, byId[id]);
  });
  var report = { folder: ATTACH_ROOT_NAME + ' / ' + ATTACH_SUBFOLDER + ' / ' + folder.getName(), relinked: relinked };
  Logger.log(JSON.stringify(report));
  return report;
}

/* =========================================================================
   MOVING OVER FROM REQUESTS

   Open RM requests become handovers, so nothing in flight is lost. One
   already issued was posted to Stock Management when it was issued, so it
   comes across as Sent with that posting recorded — receiving it confirms
   custody without posting a second time. Requests for Metallization and
   Slitting material have nowhere to go and are listed, not moved. Safe to
   run again: a moved request is marked and skipped.
   ========================================================================= */

function ADMIN_migrateToHandovers() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var reqSheet = ss.getSheetByName('Internal_Requests');
  var report = { moved: 0, alreadyMoved: 0, notMoved: [] };
  if (!reqSheet) { Logger.log(JSON.stringify(report)); return report; }
  ensureColumns_(reqSheet, ['MigratedTo']);
  handoversSheet_();
  var rows = readSheetAsObjects_(reqSheet);
  rows.forEach(function (r, idx) {
    if (!r.ReqID) return;
    if (r.MigratedTo) { report.alreadyMoved++; return; }
    var open = ['Requested', 'Approved', 'Issued'].indexOf(String(r.Status)) > -1;
    if (!open) return;
    if (String(r.Type) !== 'RM') { report.notMoved.push(r.ReqID + ' (' + r.Type + ', ' + r.Status + ')'); return; }
    var to = String(r.FromDept) === 'Store' ? 'Stores' : String(r.FromDept);
    var issued = String(r.Status) === 'Issued';
    var rec = {
      HandoverID: generateCode_('HO', 5), Date: r.Date || new Date(), Kind: 'RM', Direction: 'Forward',
      FromDept: STOCK_HOLDER, ToDept: to, ItemCode: r.ItemCode, ItemName: r.ItemName, Qty: r.Qty,
      Status: issued ? 'Sent' : 'Requested', RequestedBy: r.RequestedBy, OrderRef: r.OrderRef || '',
      Reference: r.ReqID, Remarks: 'Moved over from request ' + r.ReqID,
      StockTxnId: issued ? 'posted-at-issue:' + r.ReqID : ''
    };
    if (issued) { rec.SentBy = r.ApprovedBy || ''; rec.SentDate = r.IssuedDate || r.Date; }
    appendRowByHeader_(handoversSheet_(), rec);
    var meta = sheetMeta_(reqSheet);
    writeRowCells_(reqSheet, idx + 2, stageCells_(meta.map, { MigratedTo: rec.HandoverID }));
    report.moved++;
  });
  Logger.log(JSON.stringify(report, null, 2));
  return report;
}

/* =========================================================================
   ARCHIVING OLD HANDOVERS

   On-hand totals and every department's screen read the whole Handovers
   tab, so its size is what a send or a screen load costs. Run
   ADMIN_archiveHandovers() from the editor once the tab runs to tens of
   thousands of rows: it moves closed handovers older than monthsToKeep
   (default 12) to Handovers_Archive. Elements a received WIP handover
   moved are carried over as one "Carried" WIP entry per department and
   spec, so on-hand totals don't change. Archived handovers still open, with
   their files, through findHandover_.

   Safe to stop and run again. Rows are copied to the archive (stamped with
   the run's ID) before the carried entries are written and before anything
   leaves Handovers, and a row found in both tabs is finished off rather
   than copied twice. If a run does stop partway, on-hand totals count those
   rows twice until the next run finishes them, so run it again straight
   away. A large first archive can take several runs; the log says
   runAgain: true while there is more to move.
   ========================================================================= */

var HANDOVER_ARCHIVE_TAB = 'Handovers_Archive';
var CLOSED_STATUSES = ['Received', 'Rejected', 'Declined', 'Cancelled'];

function ADMIN_archiveHandovers(monthsToKeep) {
  monthsToKeep = Number(monthsToKeep) || 12;
  return withWriteLock_(function () {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = handoversSheet_();
    var meta = sheetMeta_(sheet);
    var lastRow = sheet.getLastRow();
    var report = { archived: 0, finishedFromEarlierRun: 0, carriedEntries: 0, kept: 0 };
    if (lastRow < 2) { Logger.log(JSON.stringify(report)); return report; }

    var archive = getOrCreateSheet_(ss, HANDOVER_ARCHIVE_TAB, meta.headers.filter(Boolean).concat(['ArchiveRun']));
    ensureColumns_(archive, meta.headers.filter(Boolean).concat(['ArchiveRun']));
    var aMeta = sheetMeta_(archive);
    var archivedRun = {};
    readSheetAsObjects_(archive).forEach(function (a) { if (a.HandoverID) archivedRun[a.HandoverID] = String(a.ArchiveRun || ''); });

    var wipSheet = getOrCreateSheet_(ss, 'WIP_Entries', TAB_SCHEMA['WIP_Entries']);
    var carriedRuns = {};
    readSheetAsObjects_(wipSheet).forEach(function (e) {
      if (e.Type === 'Carried') carriedRuns[String(e.Remarks || '')] = true;
    });

    var cutoff = new Date();
    cutoff.setMonth(cutoff.getMonth() - monthsToKeep);
    var runId = generateId_('ARC');
    var values = sheet.getRange(2, 1, lastRow - 1, meta.lastCol).getValues();
    var toCopy = [], removeRows = [];
    var carriedByRun = {}; // run -> dept -> spec -> signed qty
    // Each contiguous block of rows is one delete, and deletes are the slow
    // part, so a run stops taking rows after MAX_BLOCKS blocks and says to
    // run it again. That keeps each run well inside Apps Script's time limit.
    var MAX_BLOCKS = 200, blocks = 0, lastTaken = -1;
    var take = function (rowNum) {
      if (rowNum !== lastTaken + 1) {
        if (blocks >= MAX_BLOCKS) return false;
        blocks++;
      }
      lastTaken = rowNum;
      removeRows.push(rowNum);
      return true;
    };

    var carry = function (run, h) {
      if (String(h.Kind) !== 'WIP' || h.Status !== 'Received') return;
      var q = Number(h.Qty) || 0;
      var byDept = carriedByRun[run] = carriedByRun[run] || {};
      [[h.ToDept, q], [h.FromDept, -q]].forEach(function (pair) {
        var d = byDept[pair[0]] = byDept[pair[0]] || {};
        d[h.ItemCode] = roundQty_((d[h.ItemCode] || 0) + pair[1]);
      });
    };

    values.forEach(function (row, i) {
      var h = {};
      meta.headers.forEach(function (hd, c) { if (hd) h[hd] = row[c]; });
      if (!h.HandoverID) { report.kept++; return; }
      if (archivedRun[h.HandoverID] !== undefined) {
        // Copied by an earlier run that stopped before removing it.
        var run = archivedRun[h.HandoverID];
        if (!take(i + 2)) { report.kept++; report.runAgain = true; return; }
        if (!carriedRuns['Archive ' + run]) carry(run, h);
        report.finishedFromEarlierRun++;
        return;
      }
      var when = new Date(h.ReceivedDate || h.ClosedDate || h.Date);
      if (CLOSED_STATUSES.indexOf(String(h.Status)) === -1 || isNaN(when.getTime()) || when >= cutoff) { report.kept++; return; }
      if (!take(i + 2)) { report.kept++; report.runAgain = true; return; }
      toCopy.push(h);
      carry(runId, h);
    });

    // 1. Copy to the archive, stamped with this run.
    if (toCopy.length) {
      var aRows = toCopy.map(function (h) {
        var out = new Array(aMeta.lastCol).fill('');
        Object.keys(h).forEach(function (k) { if (aMeta.map[k] !== undefined) out[aMeta.map[k]] = sheetSafe_(h[k]); });
        out[aMeta.map['ArchiveRun']] = runId;
        return out;
      });
      archive.getRange(archive.getLastRow() + 1, 1, aRows.length, aMeta.lastCol).setValues(aRows);
      report.archived = toCopy.length;
    }

    // 2. Carried WIP entries, one per department and spec, per run.
    var wMeta = sheetMeta_(wipSheet);
    var wRows = [];
    Object.keys(carriedByRun).forEach(function (run) {
      Object.keys(carriedByRun[run]).forEach(function (dept) {
        Object.keys(carriedByRun[run][dept]).forEach(function (spec) {
          var q = carriedByRun[run][dept][spec];
          if (!q) return;
          var rec = { EntryID: generateId_('WIP'), Date: new Date(), Dept: dept, ItemCode: spec, Type: 'Carried',
                      Qty: q, By: 'archive', Remarks: 'Archive ' + run };
          var out = new Array(wMeta.lastCol).fill('');
          Object.keys(rec).forEach(function (k) { if (wMeta.map[k] !== undefined) out[wMeta.map[k]] = sheetSafe_(rec[k]); });
          wRows.push(out);
        });
      });
    });
    if (wRows.length) {
      wipSheet.getRange(wipSheet.getLastRow() + 1, 1, wRows.length, wMeta.lastCol).setValues(wRows);
      report.carriedEntries = wRows.length;
    }
    SpreadsheetApp.flush();

    // 3. Remove from Handovers, bottom-up in contiguous blocks, so each
    //    delete leaves the tab consistent if the run stops here.
    removeRows.sort(function (a, b) { return b - a; });
    var i = 0;
    while (i < removeRows.length) {
      var end = removeRows[i], start = end;
      while (i + 1 < removeRows.length && removeRows[i + 1] === start - 1) { start = removeRows[++i]; }
      sheet.deleteRows(start, end - start + 1);
      i++;
    }
    _handoversCache_ = null;
    logAudit_('admin', 'ARCHIVE_HANDOVERS', 'Handovers', runId, '', JSON.stringify(report));
    Logger.log(JSON.stringify(report, null, 2));
    return report;
  });
}

/* =========================================================================
   TRACEABILITY

   No permission gate — this is read-only, and tracing an element is useful
   to every department, not only whoever is handling it now. Returns the
   element's current state, its stage history in order, and the batch it
   was wound from.
   ========================================================================= */

function traceElement_(params) {
  var elementCode = String(params.elementCode || '').trim();
  if (!elementCode) throw new Error('Element code is required');

  // Elements tracked here before ELE Tracker existed.
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var elementRow = readRowByValue_(ss.getSheetByName('Elements'), 'ElementCode', elementCode);
  if (elementRow) {
    var element = rowToObject_(elementRow);
    var log = readSheetAsObjects_(ss.getSheetByName('Element_Stage_Log'))
      .filter(function (r) { return r.ElementCode === elementCode; });
    log.sort(function (a, b) { return new Date(a.Date).getTime() - new Date(b.Date).getTime(); });
    var batchRow = readRowByValue_(ss.getSheetByName('MPP_Batches'), 'MPPCode', element.SlitMPPCode);
    return { element: element, log: log, sourceBatch: batchRow ? rowToObject_(batchRow) : null };
  }

  // Everything since: ELE Tracker's batches. Nothing writes the tabs above
  // any more, so without this every new code came back "not found".
  var fromEle = traceEleBatch_(elementCode);
  if (fromEle) return fromEle;
  throw new Error(ELE_SHEET_ID
    ? 'No element or ELE batch found for ' + elementCode + '.'
    : 'Element not found: ' + elementCode + '. Batches are tracked in ELE Tracker now \u2014 look it up there, ' +
      'or ask an Admin to set ELE_SHEET_ID in this app\u2019s script so Trace finds them here.');
}

// ELE's stage tabs, in card order, with the stage name shown in the trace.
var ELE_STAGE_TABS = [
  ['Winding', 'Winding'], ['PaperMaskingSpray', 'Spray & Core Cleaning'], ['HeatStabilization', 'Heat Stabilization'],
  ['ShortClearing', 'Short Clearing'], ['Welding', 'Welding & Soldering'], ['PouringCuring', 'Pouring & Curing'],
  ['PouringElectricalPacking', 'Electrical Testing & Packing'], ['BDVFinalTesting', 'BDV & Final Testing']
];

function isVoided_(r) {
  return r.Voided === true || String(r.Voided).toUpperCase() === 'TRUE';
}

// An ELE batch shaped like an element trace, so the Trace screen shows it
// unchanged: the batch's state, then one line per stage entry (and per
// stage marked not applicable), oldest first.
function traceEleBatch_(code) {
  if (!ELE_SHEET_ID) return null;
  var ele = SpreadsheetApp.openById(ELE_SHEET_ID);
  var wanted = code.toLowerCase();
  var batch = readSheetAsObjects_(ele.getSheetByName('BatchMaster')).filter(function (b) {
    return String(b.MasterBatchNo).toLowerCase() === wanted || String(b.ElementBatchNo).toLowerCase() === wanted;
  })[0];
  if (!batch) return null;
  var no = String(batch.MasterBatchNo);
  var has = function (v) { return v !== '' && v !== null && v !== undefined; };

  var log = [];
  ELE_STAGE_TABS.forEach(function (t) {
    readSheetAsObjects_(ele.getSheetByName(t[0])).forEach(function (r) {
      if (String(r.MasterBatchNo) !== no || isVoided_(r)) return;
      var parts = [];
      if (has(r.TotalQty)) parts.push('Total ' + r.TotalQty);
      if (has(r.GoodQty)) parts.push('Good ' + r.GoodQty);
      if (has(r.RejectQty)) parts.push('Reject ' + r.RejectQty);
      log.push({ Date: r.Date, Stage: t[1] + (has(r.Shift) ? ' (shift ' + r.Shift + ')' : ''), Result: parts.join(' \u00b7 '),
                 Operator: r.OperatorName || '', Notes: r.Remarks || '' });
    });
  });
  readSheetAsObjects_(ele.getSheetByName('Stage_Skips')).forEach(function (r) {
    if (String(r.MasterBatchNo) !== no || isVoided_(r)) return;
    log.push({ Date: r.Date, Stage: r.Stage, Result: 'Not applicable', Operator: r.By || '', Notes: r.Reason || '' });
  });
  var qa = readSheetAsObjects_(ele.getSheetByName('FinalQA_Results')).filter(function (r) {
    return String(r.MasterBatchNo) === no && !isVoided_(r);
  });
  if (qa.length) {
    log.push({ Date: batch.CompletedDate || '', Stage: 'Final QA', Result: batch.QAVerdict || (qa.length + ' reading' + (qa.length === 1 ? '' : 's')),
               Operator: qa[0].FOAInspectorName || '', Notes: qa[0].Notes || '' });
  }
  log.sort(function (a, b) {
    var x = new Date(a.Date).getTime(), y = new Date(b.Date).getTime();
    return (isNaN(x) ? Infinity : x) - (isNaN(y) ? Infinity : y);
  });

  var spec = [has(batch.MFD) ? batch.MFD + ' MFD' : '', has(batch.Voltage) ? batch.Voltage + ' V' : ''].filter(Boolean).join(' \u00b7 ');
  return {
    element: {
      ElementCode: no, Spec: spec || batch.FGCode || '', CurrentStage: batch.CurrentStage, CurrentStatus: batch.Status,
      Customer: batch.CustomerName || '', WindingDate: batch.DatePlanning || batch.CreatedDate || '', SlitMPPCode: batch.FilmBatchNo || ''
    },
    log: log,
    sourceBatch: null,
    source: 'ELE Tracker'
  };
}
