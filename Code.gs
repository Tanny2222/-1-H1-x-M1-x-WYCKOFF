// ============================================================
//  Backtest Journal — Google Apps Script Backend
//  วิธีใช้:
//  1. เปิด Google Sheets ใหม่
//  2. Extensions → Apps Script → วางโค้ดนี้ทั้งหมด
//  3. แก้ SHEET_NAME ถ้าต้องการ
//  4. Deploy → New deployment → Web app
//     - Execute as: Me
//     - Who has access: Anyone
//  5. Copy URL ที่ได้ไปใส่ใน index.html (ตรง APPS_SCRIPT_URL)
// ============================================================

const SHEET_NAME = 'trades';
// timezone สำหรับ format วันที่ตอนอ่านกลับ — ถ้า Sheet ของคุณตั้ง timezone ไว้แล้วถูกต้อง
// (File → Settings → Time zone) จะใช้ Session.getScriptTimeZone() ตามนั้นอัตโนมัติ
// ถ้าต้องการ fix ตายตัวเป็นไทย ให้เปลี่ยนเป็น 'Asia/Bangkok'
const TIMEZONE = Session.getScriptTimeZone();

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    // สร้าง header row
    sheet.appendRow(['id','tradeNo','status','tradeDate','entryTime','session','tags','before','after','rr','note','duration','pcImage','date']);
    sheet.getRange(1,1,1,14).setFontWeight('bold').setBackground('#1a1a2e').setFontColor('#C9A84C');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getHeaders(sheet) {
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
}

function colIndex(headers, name) {
  return headers.indexOf(name); // 0-based, -1 ถ้าไม่พบ
}

function rowToObj(headers, row) {
  const obj = {};
  headers.forEach((h, i) => {
    let val = row[i];
    // parse JSON fields
    if (h === 'tags' || h === 'before' || h === 'after') {
      try { val = JSON.parse(val || '[]'); } catch(e) { val = []; }
    }
    if (h === 'rr') val = (val === '' || val === null || val === undefined) ? '' : Number(val);
    if (h === 'duration') val = (val === '' || val === null || val === undefined) ? '' : Number(val);
    if (h === 'pcImage') val = (val === '' || val === null || val === undefined) ? true : Boolean(val);
    // tradeDate ต้องออกมาเป็น 'yyyy-MM-dd' ล้วนๆ เสมอ ไม่ว่า cell จะเก็บเป็น
    // Date object (Sheets auto-convert) หรือ string ที่ติด timestamp มา
    if (h === 'tradeDate') {
      if (val instanceof Date) {
        val = Utilities.formatDate(val, TIMEZONE, 'yyyy-MM-dd');
      } else if (val) {
        val = String(val).slice(0, 10);
      }
    }
    // entryTime ต้องออกมาเป็น 'HH:mm' ล้วนๆ เสมอ — Sheets อาจ auto-convert
    // string เวลาให้กลายเป็น Date object (กับวันที่ 1899-12-30 ติดมา) ต้อง format กลับ
    if (h === 'entryTime') {
      if (val instanceof Date) {
        val = Utilities.formatDate(val, TIMEZONE, 'HH:mm');
      } else if (val) {
        val = String(val).slice(0, 5);
      }
    }
    obj[h] = val;
  });
  return obj;
}

function objToRow(headers, obj) {
  return headers.map(h => {
    let val = obj[h];
    if (h === 'tags' || h === 'before' || h === 'after') {
      val = JSON.stringify(Array.isArray(val) ? val : []);
    }
    if (val === undefined || val === null) val = '';
    return val;
  });
}

// ── CORS headers ──
function corsHeaders() {
  return {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET,POST,OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type',
    'Content-Type': 'application/json'
  };
}

function doOptions(e) {
  return ContentService.createTextOutput('')
    .setMimeType(ContentService.MimeType.TEXT);
}

// ── GET: ดึงข้อมูลทั้งหมด ──
function doGet(e) {
  try {
    const sheet = getSheet();
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return respond({ trades: [] });
    }
    const headers = getHeaders(sheet);
    const rows = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
    const trades = rows
      .filter(row => row[0] !== '') // กรอง row ว่าง
      .map(row => rowToObj(headers, row));
    return respond({ trades });
  } catch(err) {
    return respond({ error: err.message }, 500);
  }
}

// ── POST: รับ action จาก body ──
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;

    if (action === 'add') return actionAdd(body.trade);
    if (action === 'update') return actionUpdate(body.trade);
    if (action === 'delete') return actionDelete(body.id);
    if (action === 'togglePcImage') return actionTogglePcImage(body.id, body.value);
    if (action === 'deleteAll') return actionDeleteAll();

    return respond({ error: 'Unknown action' }, 400);
  } catch(err) {
    return respond({ error: err.message }, 500);
  }
}

function actionAdd(trade) {
  const sheet = getSheet();
  const headers = getHeaders(sheet);
  // สร้าง id ถ้าไม่มี
  if (!trade.id) trade.id = Date.now().toString(36) + Math.random().toString(36).slice(2,6);
  if (!trade.date) trade.date = new Date().toLocaleDateString('th-TH', { day:'2-digit', month:'short', year:'numeric' });
  sheet.appendRow(objToRow(headers, trade));
  forceDateColumnAsText(sheet, headers, sheet.getLastRow());
  return respond({ ok: true, trade });
}

function actionUpdate(trade) {
  const sheet = getSheet();
  const headers = getHeaders(sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return respond({ error: 'Not found' }, 404);

  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
  const rowIdx = ids.indexOf(trade.id);
  if (rowIdx === -1) return respond({ error: 'Not found' }, 404);

  const sheetRow = rowIdx + 2;
  forceDateColumnAsText(sheet, headers, sheetRow);
  sheet.getRange(sheetRow, 1, 1, headers.length).setValues([objToRow(headers, trade)]);
  return respond({ ok: true, trade });
}

// บังคับ cell ของ column 'tradeDate' และ 'entryTime' ในแถวนี้ให้เป็น plain text
// format (@) ก่อนเขียนค่า ป้องกัน Google Sheets auto-convert เป็น Date/Time object
function forceDateColumnAsText(sheet, headers, rowNum) {
  ['tradeDate', 'entryTime'].forEach(colName => {
    const idx = colIndex(headers, colName);
    if (idx === -1) return;
    sheet.getRange(rowNum, idx + 1).setNumberFormat('@');
  });
}

function actionDelete(id) {
  const sheet = getSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return respond({ error: 'Not found' }, 404);

  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
  const rowIdx = ids.indexOf(id);
  if (rowIdx === -1) return respond({ error: 'Not found' }, 404);

  sheet.deleteRow(rowIdx + 2);
  return respond({ ok: true });
}

function actionTogglePcImage(id, value) {
  const sheet = getSheet();
  const headers = getHeaders(sheet);
  const pcIdx = headers.indexOf('pcImage');
  const lastRow = sheet.getLastRow();
  if (lastRow < 2 || pcIdx === -1) return respond({ error: 'Not found' }, 404);

  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
  const rowIdx = ids.indexOf(id);
  if (rowIdx === -1) return respond({ error: 'Not found' }, 404);

  sheet.getRange(rowIdx + 2, pcIdx + 1).setValue(value);
  return respond({ ok: true });
}

function actionDeleteAll() {
  const sheet = getSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.deleteRows(2, lastRow - 1);
  }
  return respond({ ok: true });
}

// ── Migration: รันครั้งเดียวจาก Apps Script editor (เลือกฟังก์ชันนี้แล้วกด Run) ──
// แปลง column tradeDate ทั้งหมดที่มีอยู่แล้วให้เป็น plain text 'yyyy-MM-dd'
// แก้ปัญหาแถวเก่าที่ถูก Sheets auto-convert เป็น Date ไปแล้ว (ติด timestamp)
// หมายเหตุ: ถ้าเพิ่ม column 'entryTime' เข้าไปใน sheet เองด้วยมือ และพบว่า
// แถวเก่าๆ ก็ถูก auto-convert เป็น time เหมือนกัน ให้เพิ่ม 'entryTime' ใน
// อาเรย์ COLUMNS_TO_FIX ด้านล่างแล้วรันฟังก์ชันนี้อีกครั้ง
function migrateFixTradeDates() {
  const COLUMNS_TO_FIX = ['tradeDate']; // เพิ่ม 'entryTime' ที่นี่ถ้าจำเป็น
  const sheet = getSheet();
  const headers = getHeaders(sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  COLUMNS_TO_FIX.forEach(colName => {
    const idx = colIndex(headers, colName);
    if (idx === -1) return;
    const range = sheet.getRange(2, idx + 1, lastRow - 1, 1);
    const values = range.getValues();
    range.setNumberFormat('@'); // บังคับ format เป็น text ก่อน set ค่าใหม่
    const fmt = colName === 'entryTime' ? 'HH:mm' : 'yyyy-MM-dd';
    const sliceLen = colName === 'entryTime' ? 5 : 10;
    const fixed = values.map(r => {
      let v = r[0];
      if (v instanceof Date) v = Utilities.formatDate(v, TIMEZONE, fmt);
      else if (v) v = String(v).slice(0, sliceLen);
      return [v];
    });
    range.setValues(fixed);
  });
}

function respond(data, code) {
  const output = ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
  return output;
}
