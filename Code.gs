// ============================================================
// Clyde's Confections — Accounting System Backend
// Google Apps Script Web App
// Deploy as: Execute as ME, Anyone can access (even anonymous)
// ============================================================

const SHEET_NAME = 'CC_Data';

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.getRange(1, 1, 1, 2).setValues([['key', 'value']]);
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 220);
    sh.setColumnWidth(2, 800);
  }
  return sh;
}

// ── Build response with full CORS headers ──
function makeResponse(data) {
  const output = ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
  return output;
}

// ── OPTIONS preflight (CORS) ──
function doOptions(e) {
  return makeResponse({ ok: true });
}

// ── GET: read data ──
function doGet(e) {
  try {
    const p = e.parameter;
    const action = p.action || '';

    if (action === 'ping') {
      return makeResponse({ ok: true, msg: 'CC Accounting API is live', ts: new Date().toISOString() });
    }

    if (action === 'get' && p.key) {
      const sh = getSheet();
      const data = sh.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(p.key)) {
          return makeResponse({ ok: true, key: p.key, value: String(data[i][1]) });
        }
      }
      return makeResponse({ ok: true, key: p.key, value: null });
    }

    if (action === 'list') {
      const sh = getSheet();
      const data = sh.getDataRange().getValues();
      const keys = data.slice(1).map(r => String(r[0])).filter(Boolean);
      return makeResponse({ ok: true, keys: keys });
    }

    return makeResponse({ ok: false, error: 'Unknown GET action: ' + action });

  } catch (err) {
    return makeResponse({ ok: false, error: err.toString() });
  }
}

// ── POST: write data ──
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action || '';
    const key = body.key;
    const value = body.value;

    if (action === 'set') {
      if (key === undefined || value === undefined) {
        return makeResponse({ ok: false, error: 'Missing key or value' });
      }
      const sh = getSheet();
      const data = sh.getDataRange().getValues();
      const valStr = (typeof value === 'string') ? value : JSON.stringify(value);
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(key)) {
          sh.getRange(i + 1, 2).setValue(valStr);
          return makeResponse({ ok: true, action: 'updated', key: key });
        }
      }
      sh.appendRow([key, valStr]);
      return makeResponse({ ok: true, action: 'inserted', key: key });
    }

    if (action === 'delete') {
      if (!key) return makeResponse({ ok: false, error: 'Missing key' });
      const sh = getSheet();
      const data = sh.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(key)) {
          sh.deleteRow(i + 1);
          return makeResponse({ ok: true, action: 'deleted', key: key });
        }
      }
      return makeResponse({ ok: true, action: 'not_found', key: key });
    }

    if (action === 'batch_set' && Array.isArray(body.pairs)) {
      const sh = getSheet();
      const data = sh.getDataRange().getValues();
      const keyIndex = {};
      for (let i = 1; i < data.length; i++) {
        keyIndex[String(data[i][0])] = i + 1;
      }
      const toAppend = [];
      for (const pair of body.pairs) {
        const k = String(pair.key);
        const v = (typeof pair.value === 'string') ? pair.value : JSON.stringify(pair.value);
        if (keyIndex[k]) {
          sh.getRange(keyIndex[k], 2).setValue(v);
        } else {
          toAppend.push([k, v]);
        }
      }
      if (toAppend.length > 0) {
        sh.getRange(sh.getLastRow() + 1, 1, toAppend.length, 2).setValues(toAppend);
      }
      return makeResponse({ ok: true, action: 'batch_set', count: body.pairs.length });
    }

    return makeResponse({ ok: false, error: 'Unknown POST action: ' + action });

  } catch (err) {
    return makeResponse({ ok: false, error: err.toString() });
  }
}

// ── Manual test — run this in the Apps Script editor to verify ──
function testAPI() {
  Logger.log('Sheet: ' + getSheet().getName());
  Logger.log('Rows: ' + getSheet().getLastRow());
  Logger.log('Ping response: ' + JSON.stringify(makeResponse({ ok: true, msg: 'test' })));
}
