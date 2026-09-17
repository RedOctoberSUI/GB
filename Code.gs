/*
 * Wirtschaft zum Grünen Baum 2026 — Apps Script Backend
 * CLv0.077
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

const SHEET_ID       = '1u31CdcQt4BFJkpZDC5_ShIzkHcHh23bJBt2L8mTDOgc';
const SHEET_NAME     = 'Anmeldungen GB';   // ggf. anpassen wenn das Tab anders heisst
const SETTINGS_SHEET = 'Settings';         // 2. Tab für Key/Value-Einstellungen
const HEADER_ROW     = 1;
const SECRET_TOKEN   = 'oEmhvp6yDaS7XVAp3q2MyhsELMClhtq-';

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
  if (action === 'getAll')      return jsonResponse({ ok: true, rows: getAllRows() });
  if (action === 'getSettings') return jsonResponse({ ok: true, settings: getAllSettings() });
  return jsonResponse({ ok: false, error: 'Unknown action: ' + action });
}

function doPost(e) {
  try {
    const body   = JSON.parse(e.postData.contents);
    if (body.token !== SECRET_TOKEN) {
      return jsonResponse({ ok: false, error: 'Unauthorized' });
    }
    const action = body.action || 'insert';

    if (action === 'update')     return handleUpdate(body);
    if (action === 'setSetting') return handleSetSetting(body);
    if (action === 'replaceAll') return handleReplaceAll(body);
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

  let status2026;
  if (data.year_2026 !== undefined && data.year_2026 !== null) {
    // Explizit gesetzter Wert (auch '' für Offen) wird 1:1 verwendet
    status2026 = data.year_2026;
  } else if (data.typ === 'Absage')            status2026 = 'Abgemeldet';
  else if (data.typ === 'PNG')                 status2026 = 'PNG';
  else if (data.typ === 'Einzuladen')          status2026 = 'Einzuladen';
  else if (data.typ === 'Einladung gesendet')  status2026 = 'Einladung gesendet';
  else if (data.typ === 'Offen')               status2026 = '';
  else                                          status2026 = 'Angemeldet';
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
    // Bei Solo-Anmeldung: alte Begleitungs- und Kinderfelder aus dem Vorjahr aktiv leeren
    if (data.typ === 'Solo') {
      sheet.getRange(targetRow, COL.b_vorname).setValue('');
      sheet.getRange(targetRow, COL.b_nachname).setValue('');
      sheet.getRange(targetRow, COL.kinder).setValue('');
    } else {
      setIfFilled(sheet, targetRow, COL.b_vorname,  data.b_vorname);
      setIfFilled(sheet, targetRow, COL.b_nachname, data.b_nachname);
      setIfFilled(sheet, targetRow, COL.kinder,     data.bier);
    }
    // Kommentar immer überschreiben (auch mit leer) — sonst bleibt alter Kommentar aus Vorjahr
    sheet.getRange(targetRow, COL.kommentar).setValue(data.kommentar || '');
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
    if (year2026 === 'Abgemeldet')            typ = 'Absage';
    else if (year2026 === 'PNG')              typ = 'PNG';
    else if (year2026 === 'Einzuladen')       typ = 'Einzuladen';
    else if (year2026 === 'Einladung gesendet') typ = 'Einladung gesendet';
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


// ─── Settings (2. Tab "Settings" mit Key/Value) ──────────────────────────

function getSettingsSheet_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(SETTINGS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(SETTINGS_SHEET);
    sheet.getRange(1, 1, 1, 2).setValues([['Key', 'Value']]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getAllSettings() {
  const sheet = getSettingsSheet_();
  const last = sheet.getLastRow();
  if (last <= 1) return {};
  const data = sheet.getRange(2, 1, last - 1, 2).getValues();
  const out = {};
  for (const r of data) {
    const k = (r[0] || '').toString().trim();
    if (k) out[k] = formatSettingValue_(r[1]);
  }
  return out;
}

function formatSettingValue_(v) {
  if (v === null || v === undefined || v === '') return '';
  if (Object.prototype.toString.call(v) === '[object Date]') {
    if (v.getFullYear() === 1899) {
      const hh = String(v.getHours()).padStart(2, '0');
      const mm = String(v.getMinutes()).padStart(2, '0');
      return hh + ':' + mm;
    }
    const y  = v.getFullYear();
    const mo = String(v.getMonth() + 1).padStart(2, '0');
    const d  = String(v.getDate()).padStart(2, '0');
    return y + '-' + mo + '-' + d;
  }
  return v.toString();
}

function handleSetSetting(body) {
  const key   = (body.key   || '').toString().trim();
  const value = (body.value || '').toString();
  if (!key) return jsonResponse({ ok: false, error: 'Key fehlt' });
  const sheet = getSettingsSheet_();
  const last = sheet.getLastRow();
  let targetRow = -1;
  if (last > 1) {
    const keys = sheet.getRange(2, 1, last - 1, 1).getValues();
    for (let i = 0; i < keys.length; i++) {
      if ((keys[i][0] || '').toString().trim() === key) {
        targetRow = i + 2;
        break;
      }
    }
  }
  if (targetRow > 0) {
    sheet.getRange(targetRow, 2).setNumberFormat('@').setValue(value);
  } else {
    sheet.appendRow([key, '']);
    sheet.getRange(sheet.getLastRow(), 2).setNumberFormat('@').setValue(value);
  }
  return jsonResponse({ ok: true, key: key });
}


// ─── ReplaceAll (CSV-Import: löscht alle Daten und schreibt neue Zeilen) ─
function handleReplaceAll(body) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  if (!sheet) return jsonResponse({ ok: false, error: 'Sheet "' + SHEET_NAME + '" nicht gefunden' });

  const rows = body.rows;
  if (!Array.isArray(rows)) return jsonResponse({ ok: false, error: 'rows fehlt oder kein Array' });

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow > HEADER_ROW) {
    // Alle Daten-Zeilen löschen
    sheet.getRange(HEADER_ROW + 1, 1, lastRow - HEADER_ROW, lastCol).clearContent();
  }

  if (rows.length === 0) return jsonResponse({ ok: true, action: 'replaced', rows: 0 });

  // Zeilen aufbauen im festen Spalten-Layout
  const data = rows.map(r => [
    r.timestamp   || '',
    r.vorname     || '',
    r.nachname    || '',
    r.email       || '',
    r.tel         || '',
    r.b_vorname   || '',
    r.b_nachname  || '',
    r.bier        || '',
    r.einladend   || '',
    r.kommentar   || '',
    r.status_2025 || '',
    r.status_2026 || '',
  ]);
  sheet.getRange(HEADER_ROW + 1, 1, data.length, 12).setValues(data);
  return jsonResponse({ ok: true, action: 'replaced', rows: rows.length });
}
