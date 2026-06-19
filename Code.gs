/*
 * Wirtschaft zum Grünen Baum 2026 — Apps Script Backend
 * CLv0.041
 *
 * Sheet: 1u31CdcQt4BFJkpZDC5_ShIzkHcHh23bJBt2L8mTDOgc
 *
 * Spalten (1-indexiert):
 *   A=1  Timestamp
 *   B=2  Vorname
 *   C=3  Nachname
 *   D=4  Email
 *   E=5  Tel
 *   F=6  Begleitung Vorname
 *   G=7  Begleitung Nachname
 *   H=8  Anzahl Kinder
 *   I=9  Einladend
 *   J=10 Kommentar
 *   K=11 2025
 *   L=12 2026
 */

const SHEET_ID    = '1u31CdcQt4BFJkpZDC5_ShIzkHcHh23bJBt2L8mTDOgc';
const SHEET_NAME  = 'Anmeldungen GB';   // ggf. anpassen wenn das Tab anders heisst
const HEADER_ROW  = 1;
const SECRET_TOKEN= 'oEmhvp6yDaS7XVAp3q2MyhsELMClhtq-';

const COL = {
  timestamp:   1,
  vorname:     2,
  nachname:    3,
  email:       4,
  tel:         5,
  b_vorname:   6,
  b_nachname:  7,
  kinder:      8,
  einladend:   9,
  kommentar:  10,
  year_2025:  11,
  year_2026:  12,
};

// ─── Entry points ─────────────────────────────────────────────────────────

function doGet(e) {
  const params = (e && e.parameter) || {};
  if (params.token !== SECRET_TOKEN) {
    return jsonResponse({ ok: false, error: 'Unauthorized' });
  }
  const action = params.action || 'getAll';
  if (action === 'getAll') return jsonResponse({ ok: true, rows: getAllRows() });
  return jsonResponse({ ok: false, error: 'Unknown action: ' + action });
}

function doPost(e) {
  try {
    const body   = JSON.parse(e.postData.contents);
    if (body.token !== SECRET_TOKEN) {
      return jsonResponse({ ok: false, error: 'Unauthorized' });
    }
    const action = body.action || 'insert';

    if (action === 'update') return handleUpdate(body);
    return handleUpsert(body);
  } catch (err) {
    return jsonResponse({ ok: false, error: String(err) });
  }
}

// ─── Handlers ─────────────────────────────────────────────────────────────

/*
 * Anmeldung/Absage vom Form oder manuelle Erfassung im Admin.
 * Sucht bestehende Zeile per Vor-/Nachname (case-insensitive). Wenn gefunden:
 * updated Felder + 2026-Status. Sonst neue Zeile unten anhängen.
 */
function handleUpsert(data) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  if (!sheet) return jsonResponse({ ok: false, error: 'Sheet "' + SHEET_NAME + '" nicht gefunden' });

  const status2026 = data.typ === 'Absage' ? 'Abgemeldet' : 'Angemeldet';
  const vn = (data.vorname  || '').toString().trim();
  const nn = (data.nachname || '').toString().trim();
  const timestamp = data.timestamp || new Date().toLocaleString('de-CH');

  // Suche nach passender Zeile
  const lastRow = sheet.getLastRow();
  let targetRow = -1;
  if (lastRow > HEADER_ROW) {
    const values = sheet.getRange(HEADER_ROW + 1, COL.vorname, lastRow - HEADER_ROW, 2).getValues();
    for (let i = 0; i < values.length; i++) {
      const rowVn = (values[i][0] || '').toString().trim().toLowerCase();
      const rowNn = (values[i][1] || '').toString().trim().toLowerCase();
      if (rowVn === vn.toLowerCase() && rowNn === nn.toLowerCase()) {
        targetRow = HEADER_ROW + 1 + i;
        break;
      }
    }
  }

  if (targetRow > 0) {
    // Update bestehende Zeile — nur nicht-leere Felder überschreiben
    setIfFilled(sheet, targetRow, COL.timestamp,  timestamp);
    setIfFilled(sheet, targetRow, COL.email,      data.email);
    setIfFilled(sheet, targetRow, COL.tel,        data.tel);
    setIfFilled(sheet, targetRow, COL.b_vorname,  data.b_vorname);
    setIfFilled(sheet, targetRow, COL.b_nachname, data.b_nachname);
    setIfFilled(sheet, targetRow, COL.kinder,     data.bier);  // 'bier' Feld trägt jetzt Kinderzahl
    setIfFilled(sheet, targetRow, COL.kommentar,  data.kommentar);
    sheet.getRange(targetRow, COL.year_2026).setValue(status2026);
    return jsonResponse({ ok: true, action: 'updated', row: targetRow });
  }

  // Neue Zeile unten anfügen
  const newRow = [
    timestamp,
    vn,
    nn,
    data.email     || '',
    data.tel       || '',
    data.b_vorname || '',
    data.b_nachname|| '',
    data.bier      || '',
    data.einladend || '',
    data.kommentar || '',
    '',                // 2025 leer
    status2026,        // 2026
  ];
  sheet.appendRow(newRow);
  return jsonResponse({ ok: true, action: 'inserted', row: sheet.getLastRow() });
}

/*
 * Manuelles Update via Admin-Edit-Modal. data enthält die kompletten Felder
 * und rowIndex zeigt die Zeile (1-basiert wie im Sheet).
 */
function handleUpdate(body) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  if (!sheet) return jsonResponse({ ok: false, error: 'Sheet "' + SHEET_NAME + '" nicht gefunden' });

  const rowIndex = parseInt(body.rowIndex, 10);
  if (!rowIndex || rowIndex <= HEADER_ROW) {
    return jsonResponse({ ok: false, error: 'Ungültiger rowIndex' });
  }
  const d = body.data || {};
  sheet.getRange(rowIndex, COL.vorname).setValue(d.vorname     || '');
  sheet.getRange(rowIndex, COL.nachname).setValue(d.nachname    || '');
  sheet.getRange(rowIndex, COL.email).setValue(d.email       || '');
  sheet.getRange(rowIndex, COL.tel).setValue(d.tel         || '');
  sheet.getRange(rowIndex, COL.b_vorname).setValue(d.b_vorname   || '');
  sheet.getRange(rowIndex, COL.b_nachname).setValue(d.b_nachname  || '');
  sheet.getRange(rowIndex, COL.kinder).setValue(d.bier        || '');
  sheet.getRange(rowIndex, COL.einladend).setValue(d.einladend   || '');
  sheet.getRange(rowIndex, COL.kommentar).setValue(d.kommentar   || '');
  if (d.year_2026 !== undefined) sheet.getRange(rowIndex, COL.year_2026).setValue(d.year_2026 || '');
  return jsonResponse({ ok: true, action: 'updated', row: rowIndex });
}

// ─── Read ─────────────────────────────────────────────────────────────────

function getAllRows() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow <= HEADER_ROW) return [];
  const lastCol = sheet.getLastColumn();
  const data = sheet.getRange(HEADER_ROW + 1, 1, lastRow - HEADER_ROW, lastCol).getValues();

  const rows = [];
  for (let i = 0; i < data.length; i++) {
    const r = data[i];
    const vorname  = (r[COL.vorname  - 1] || '').toString().trim();
    const nachname = (r[COL.nachname - 1] || '').toString().trim();
    // Komplett leere Zeilen überspringen
    if (!vorname && !nachname) continue;

    const year2026 = (r[COL.year_2026 - 1] || '').toString().trim();
    const status2025 = (r[COL.year_2025 - 1] || '').toString().trim();

    // 'typ' aus 2026 ableiten — für Admin-Tool
    let typ;
    if (year2026 === 'Abgemeldet') typ = 'Absage';
    else if (year2026 === 'Angemeldet') {
      const hasBegl = (r[COL.b_vorname - 1] || '').toString().trim() !== '';
      typ = hasBegl ? 'Mit Begleitung' : 'Solo';
    } else {
      typ = 'Offen';  // noch keine Antwort 2026
    }

    rows.push({
      _rowIndex:  HEADER_ROW + 1 + i,
      timestamp:  (r[COL.timestamp  - 1] || '').toString(),
      vorname:    vorname,
      nachname:   nachname,
      email:      (r[COL.email      - 1] || '').toString(),
      tel:        (r[COL.tel        - 1] || '').toString(),
      b_vorname:  (r[COL.b_vorname  - 1] || '').toString(),
      b_nachname: (r[COL.b_nachname - 1] || '').toString(),
      bier:       (r[COL.kinder     - 1] || '').toString(),  // Compat: weiterhin 'bier'
      einladend:  (r[COL.einladend  - 1] || '').toString(),
      kommentar:  (r[COL.kommentar  - 1] || '').toString(),
      status_2025:status2025,
      status_2026:year2026,
      typ:        typ,
      source:     status2025 ? 'liste' : 'form', // Vorab-Eintrag oder Live-Anmeldung
    });
  }
  return rows;
}

// ─── Helpers ──────────────────────────────────────────────────────────────

function setIfFilled(sheet, row, col, value) {
  if (value === undefined || value === null) return;
  const s = value.toString().trim();
  if (!s) return;
  sheet.getRange(row, col).setValue(value);
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

