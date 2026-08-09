/**
 * Aom-Training — Google Apps Script backend
 * ใช้คู่กับแดชบอร์ด https://nestapitchayut2-cloud.github.io/Aom-Training/
 *
 * วิธีติดตั้ง: ดูขั้นตอนใน README.md (SETUP-GOOGLE-SHEET)
 */

var SHEET_LOG  = 'Log';
var SHEET_MOCK = 'Mock';

var HEAD_LOG  = ['timestamp','date','type','hours_oet','hours_stage3','note'];
var HEAD_MOCK = ['timestamp','date','listening','reading','writing','speaking','calc','applied','law','osce','note'];

/** สร้างชีตถ้ายังไม่มี + ใส่หัวตาราง */
function ensureSheet_(name, head) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
  }
  if (sh.getLastRow() === 0) {
    sh.appendRow(head);
    sh.getRange(1, 1, 1, head.length).setFontWeight('bold').setBackground('#e8eef7');
    sh.setFrozenRows(1);
  }
  return sh;
}

/** อ่านชีตออกมาเป็น array ของ object */
function readSheet_(name, head) {
  var sh = ensureSheet_(name, head);
  var last = sh.getLastRow();
  if (last < 2) return [];
  var values = sh.getRange(2, 1, last - 1, head.length).getValues();
  return values.map(function (r) {
    var o = {};
    head.forEach(function (k, i) {
      var v = r[i];
      if (v instanceof Date) v = Utilities.formatDate(v, 'Asia/Bangkok', 'yyyy-MM-dd');
      o[k] = v;
    });
    return o;
  }).filter(function (o) { return o.date; });
}

/** GET — ดึงข้อมูลทั้งหมดกลับไปแสดงในแดชบอร์ด */
function doGet(e) {
  var out;
  try {
    out = {
      ok: true,
      log:  readSheet_(SHEET_LOG,  HEAD_LOG),
      mock: readSheet_(SHEET_MOCK, HEAD_MOCK)
    };
  } catch (err) {
    out = { ok: false, error: String(err) };
  }
  var cb = e && e.parameter && e.parameter.callback;
  if (cb) {
    return ContentService
      .createTextOutput(cb + '(' + JSON.stringify(out) + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

/** POST — บันทึกแถวใหม่ (หรือลบแถว) */
function doPost(e) {
  var res = { ok: false };
  try {
    var body = JSON.parse(e.postData.contents);

    if (body.action === 'delete') {
      var head = body.sheet === SHEET_MOCK ? HEAD_MOCK : HEAD_LOG;
      var sh = ensureSheet_(body.sheet, head);
      var last = sh.getLastRow();
      if (last >= 2) {
        var stamps = sh.getRange(2, 1, last - 1, 1).getValues();
        for (var i = stamps.length - 1; i >= 0; i--) {
          if (String(stamps[i][0]) === String(body.timestamp)) {
            sh.deleteRow(i + 2);
            break;
          }
        }
      }
      res = { ok: true, deleted: true };

    } else if (body.sheet === SHEET_MOCK) {
      var m = ensureSheet_(SHEET_MOCK, HEAD_MOCK);
      var ts = body.timestamp || new Date().getTime();
      m.appendRow([ts, body.date, body.listening, body.reading, body.writing,
                   body.speaking, body.calc, body.applied, body.law, body.osce, body.note || '']);
      res = { ok: true, timestamp: ts };

    } else {
      var l = ensureSheet_(SHEET_LOG, HEAD_LOG);
      var ts2 = body.timestamp || new Date().getTime();
      l.appendRow([ts2, body.date, body.type, body.hours_oet || 0, body.hours_stage3 || 0, body.note || '']);
      res = { ok: true, timestamp: ts2 };
    }
  } catch (err) {
    res = { ok: false, error: String(err) };
  }
  return ContentService
    .createTextOutput(JSON.stringify(res))
    .setMimeType(ContentService.MimeType.JSON);
}

/** รันครั้งเดียวเพื่อสร้างชีตให้พร้อม (กด Run ใน editor) */
function setup() {
  ensureSheet_(SHEET_LOG, HEAD_LOG);
  ensureSheet_(SHEET_MOCK, HEAD_MOCK);
  SpreadsheetApp.getActiveSpreadsheet().toast('สร้างชีต Log และ Mock เรียบร้อย', 'Aom-Training', 5);
}
