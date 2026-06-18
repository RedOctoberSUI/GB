// ═══════════════════════════════════════════════════════
//  Fest-Anmeldung · Apps Script  CLv0.006
//  Sheet-ID: 1BLJSVqF2s-WJ8pZvKQyPxYSimrxOjQrVOFZJDHKE8p0
// ═══════════════════════════════════════════════════════

const SHEET_NAME = 'Anmeldungen';
const HEADERS    = ['timestamp','typ','source','vorname','nachname','email','tel','b_vorname','b_nachname','bier','kommentar'];

function getSheet() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let   sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    const hdr = sheet.getRange(1, 1, 1, HEADERS.length);
    hdr.setValues([['Timestamp','Typ','Quelle','Vorname','Nachname','E-Mail','Telefon','Begl. Vorname','Begl. Nachname','Bier','Kommentar']]);
    hdr.setFontWeight('bold').setBackground('#1a1a18').setFontColor('#ffffff');
  }
  return sheet;
}

// ── GET: load all rows ────────────────────────────────
function doGet(e) {
  const action = e && e.parameter && e.parameter.action;
  if (action === 'getAll') {
    const sheet  = getSheet();
    const data   = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      return jsonResponse({ rows: [] });
    }
    const rows = data.slice(1).map((row, i) => {
      const obj = {};
      HEADERS.forEach((h, j) => obj[h] = row[j] !== undefined ? String(row[j]) : '');
      obj._rowIndex = i + 2;
      return obj;
    });
    return jsonResponse({ rows });
  }
  return jsonResponse({ status: 'ok' });
}

// ── POST: new entry or update ─────────────────────────
function doPost(e) {
  try {
    const data   = JSON.parse(e.postData.contents);
    const sheet  = getSheet();
    const action = data.action;

    if (action === 'update') {
      const rowIdx = parseInt(data.rowIndex);
      const d      = data.data;
      sheet.getRange(rowIdx, 1, 1, HEADERS.length).setValues([[
        d.timestamp  || '',
        d.typ        || '',
        d.source     || '',
        d.vorname    || '',
        d.nachname   || '',
        d.email      || '',
        d.tel        || '',
        d.b_vorname  || '',
        d.b_nachname || '',
        d.bier        || 0,
        d.kommentar  || '',
      ]]);
      return jsonResponse({ status: 'ok', action: 'updated', row: rowIdx });
    }

    // Default: new registration
    sheet.appendRow([
      data.timestamp  || new Date().toLocaleString('de-CH'),
      data.typ        || '',
      data.source     || 'form',
      data.vorname    || '',
      data.nachname   || '',
      data.email      || '',
      data.tel        || '',
      data.b_vorname  || '',
      data.b_nachname || '',
      data.bier       || 0,
      data.kommentar  || '',
    ]);
    return jsonResponse({ status: 'ok', action: 'created' });

  } catch (err) {
    return jsonResponse({ status: 'error', message: err.toString() });
  }
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Test ──────────────────────────────────────────────
function testInsert() {
  const fake = { postData: { contents: JSON.stringify({
    typ: 'Mit Begleitung', source: 'form',
    vorname: 'Anna', nachname: 'Muster',
    email: 'anna@beispiel.ch', tel: '+41 79 000 00 00',
    b_vorname: 'Ben', b_nachname: 'Muster', bier: 3, kommentar: 'Test',
    timestamp: new Date().toLocaleString('de-CH'),
  })}};
  Logger.log(doPost(fake).getContent());
}
