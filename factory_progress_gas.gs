// ============================================================
// 工場進捗管理 GAS サーバー v5.0
// ============================================================

function doGet(e) {
  try {
    const action = e.parameter.action || '';
    if (action === 'ping') return ok({ message: 'pong' });
    if (action === 'getAll') return getAll_();
    if (action === 'getShipping') return getShipping_();
    return errRes('unknown action: ' + action);
  } catch (ex) {
    return errRes(ex.message);
  }
}

function doPost(e) {
  try {
    const p = JSON.parse(e.postData.contents);
    const action = p.action || '';
    if (action === 'saveItemMaster')    return saveItemMaster_(p);
    if (action === 'savePlansProgress') return savePlansProgress_(p);
    if (action === 'saveShiftAttendance') return saveShiftAttendance_(p);
    if (action === 'saveOperationLog')  return saveOperationLog_(p);
    if (action === 'saveAll')           return saveAll_(p);
    if (action === 'saveShipping')       return saveShipping_(p);
    return errRes('unknown action: ' + action);
  } catch (ex) {
    return errRes(ex.message);
  }
}

// ============================================================
// getAll
// ============================================================
function getAll_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ok({
    itemMaster:   readJsonSheet_(ss, 'itemMaster'),
    plans:        readPlansSheet_(ss),
    progress:     readProgressSheet_(ss),
    shiftAttend:  readShiftAttend_(ss),
    operationLog: readJsonSheet_(ss, 'operationLog')
  });
}

// ============================================================
// itemMaster
// ============================================================
function saveItemMaster_(p) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  writeJsonSheet_(ss, 'itemMaster', p.itemMaster || []);
  return ok({});
}

// ============================================================
// plans + progress
// ============================================================
function savePlansProgress_(p) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  writePlansSheet_(ss, p.plans || []);
  writeProgressSheet_(ss, p.progress || {});
  return ok({});
}

// ============================================================
// shiftAttendance (v5.0: startTime/endTime対応)
// ============================================================
function saveShiftAttendance_(p) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  var shName = 'shiftAttendance';
  var sheet = ss.getSheetByName(shName);
  if (!sheet) sheet = ss.insertSheet(shName);
  sheet.clearContents();

  var rows = [
    ['baseShift',    JSON.stringify(p.baseShift    || [])],
    ['weeklyAttend', JSON.stringify(p.weeklyAttend || {})]
  ];
  sheet.getRange(1, 1, rows.length, 2).setValues(rows);
  return ok({});
}

// ============================================================
// operationLog (v5.0新規)
// ============================================================
function saveOperationLog_(p) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  var shName = 'operationLog';
  var sheet = ss.getSheetByName(shName);
  if (!sheet) {
    sheet = ss.insertSheet(shName);
    sheet.getRange(1,1,1,8).setValues([['timestamp','deviceName','planId','itemName','color','proc','prevVal','newVal']]);
  }
  var log = p.log;
  if (!log) return ok({});
  sheet.appendRow([
    log.timestamp || new Date().toISOString(),
    log.deviceName || '',
    log.planId     || '',
    log.itemName   || '',
    log.color      || '',
    log.proc       || '',
    log.prevVal !== undefined ? log.prevVal : '',
    log.newVal  !== undefined ? log.newVal  : ''
  ]);
  return ok({});
}

// ============================================================
// saveAll
// ============================================================
function saveAll_(p) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (p.itemMaster)   writeJsonSheet_(ss, 'itemMaster', p.itemMaster);
  if (p.plans)        writePlansSheet_(ss, p.plans);
  if (p.progress)     writeProgressSheet_(ss, p.progress);
  if (p.baseShift !== undefined || p.weeklyAttend !== undefined) {
    saveShiftAttendance_({ baseShift: p.baseShift || [], weeklyAttend: p.weeklyAttend || {} });
  }
  return ok({});
}

// ============================================================
// シート読み書きヘルパー
// ============================================================

// plansシート: planId,date(YYYY-MM-DD),itemId,qty,memo,priority
function readPlansSheet_(ss) {
  var sheet = ss.getSheetByName('plans');
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  var result = [];
  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    if (!r[0]) continue;
    result.push({
      id:       String(r[0]),
      date:     toDateString_(r[1]),
      itemId:   String(r[2] || ''),
      qty:      Number(r[3] || 0),
      memo:     String(r[4] || ''),
      priority: r[5] === true || r[5] === 'TRUE' || r[5] === 1
    });
  }
  return result;
}

function writePlansSheet_(ss, plans) {
  var sheet = ss.getSheetByName('plans');
  if (!sheet) sheet = ss.insertSheet('plans');
  sheet.clearContents();
  var rows = [['id','date','itemId','qty','memo','priority']];
  for (var i = 0; i < plans.length; i++) {
    var p = plans[i];
    rows.push([p.id, p.date, p.itemId, p.qty, p.memo || '', p.priority ? 'TRUE' : 'FALSE']);
  }
  if (rows.length > 1) sheet.getRange(1, 1, rows.length, 6).setValues(rows);
  else sheet.getRange(1,1,1,6).setValues(rows);
}

// progressシート: planId,preDone,postDone(JSON)
function readProgressSheet_(ss) {
  var sheet = ss.getSheetByName('progress');
  if (!sheet) return {};
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return {};
  var result = {};
  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    if (!r[0]) continue;
    var postDone = {};
    try { postDone = JSON.parse(String(r[2] || '{}')); } catch(e) { postDone = {}; }
    result[String(r[0])] = {
      preDone:  Number(r[1] || 0),
      postDone: postDone
    };
  }
  return result;
}

function writeProgressSheet_(ss, progress) {
  var sheet = ss.getSheetByName('progress');
  if (!sheet) sheet = ss.insertSheet('progress');
  sheet.clearContents();
  var rows = [['planId','preDone','postDone']];
  var keys = Object.keys(progress);
  for (var i = 0; i < keys.length; i++) {
    var k = keys[i];
    var v = progress[k];
    rows.push([k, v.preDone || 0, JSON.stringify(v.postDone || {})]);
  }
  sheet.getRange(1, 1, rows.length, 3).setValues(rows);
}

// shiftAttendance (key-value形式)
function readShiftAttend_(ss) {
  var sheet = ss.getSheetByName('shiftAttendance');
  if (!sheet) return { baseShift: [], weeklyAttend: {} };
  var data = sheet.getDataRange().getValues();
  var kv = {};
  for (var i = 0; i < data.length; i++) {
    if (data[i][0]) kv[String(data[i][0])] = String(data[i][1] || '');
  }
  var baseShift = [];
  var weeklyAttend = {};
  try { baseShift = JSON.parse(kv['baseShift'] || '[]'); } catch(e) {}
  try { weeklyAttend = JSON.parse(kv['weeklyAttend'] || '{}'); } catch(e) {}
  return { baseShift: baseShift, weeklyAttend: weeklyAttend };
}

// JSON単列シート (itemMaster, operationLog)
function readJsonSheet_(ss, sheetName) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  var result = [];
  for (var i = 1; i < data.length; i++) {
    try { result.push(JSON.parse(String(data[i][0]))); } catch(e) {}
  }
  return result;
}

function writeJsonSheet_(ss, sheetName, arr) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);
  sheet.clearContents();
  if (!arr || !arr.length) return;
  var rows = [['json']];
  for (var i = 0; i < arr.length; i++) rows.push([JSON.stringify(arr[i])]);
  sheet.getRange(1, 1, rows.length, 1).setValues(rows);
}

// 日付→YYYY-MM-DD文字列
function toDateString_(v) {
  if (!v) return '';
  if (typeof v === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(v)) return v;
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var d = new Date(String(v));
  if (!isNaN(d.getTime())) return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return String(v);
}

// レスポンスヘルパー
function ok(data) {
  data.ok = true;
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}
function errRes(msg) {
  return ContentService.createTextOutput(JSON.stringify({ ok: false, error: msg })).setMimeType(ContentService.MimeType.JSON);
}
function getShipping_() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('deliveryHistory');
    if (!sheet) {
      sheet = ss.insertSheet('deliveryHistory');
      sheet.appendRow(['id','timestamp','deviceName','deliveryNo','memo','palletsJson','shippingDate']);
    }
const rows = sheet.getDataRange().getValues();
    const headers = rows[0];
    const records = rows.slice(1).map(row => {
      const obj = {};
      headers.forEach((h, i) => obj[h] = row[i]);
      return obj;
    });
    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, deliveryHistory: records }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
function saveShipping_(p) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('deliveryHistory');
    if (!sheet) {
      sheet = ss.insertSheet('deliveryHistory');
      sheet.appendRow(['id','timestamp','deviceName','deliveryNo','memo','palletsJson','shippingDate']);
    }  
 const records = p.records || [];
    records.forEach(r => {
      sheet.appendRow([
        r.id || '',
        r.timestamp || '',
        r.deviceName || '',
        r.deliveryNo || '',
        r.memo || '',
        r.palletsJson || '',
        r.shippingDate || ''
      ]);
    });
return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}