/**
 * Zitherinventar – Backend in Google Apps Script
 *
 * Speichert die Erfassungen im Tabellenblatt „Bestand“, liest die alte Liste
 * aus dem Blatt „Alte Liste“ und legt Fotos in einem privaten Drive-Ordner ab.
 * Zugriff nur mit dem Schlüssel (TOKEN), den setup() erzeugt.
 */

const SHEET_BESTAND = 'Bestand';
const SHEET_ALT = 'Alte Liste';
const TZ = 'Europe/Berlin';

// Spalten im Blatt „Bestand“ – Reihenfolge = Reihenfolge in der Tabelle
const COLS = [
  ['zeitstempel', 'Zeitstempel'],
  ['barcode', 'neue Signatur'],
  ['altSig', 'Alte Signatur'],
  ['hersteller', 'Hersteller'],
  ['fabriknummer', 'Fabriknummer'],
  ['baujahr', 'Baujahr'],
  ['zustand', 'Zustand'],
  ['beschaedigungen', 'Beschädigungen'],
  ['saiten', 'Zustand der Saiten'],
  ['koffer', 'Koffer'],
  ['foto', 'Foto'],
  ['klang', 'Wie findet Sarah den Klang?'],
  ['besonderheiten', 'Besonderheiten und Sonstiges'],
  ['gattung', 'Gattung'],
  ['mensur', 'Mensur'],
  ['lfd', 'Lfd. Nr. (alt)'],
  ['regal', 'Regal-Nr.'],
  ['regalort', 'Regalort'],
  ['verkauf', 'Verkauf'],
  ['kaufwert', 'Kaufwert (DM)'],
  ['wert', 'tatsächlicher Wert (DM)'],
  ['fotoIds', 'Foto-IDs'],
  ['erstellt', 'Erstellt']
];

const OLD_COLS = ['lfd', 'sig', 'regal', 'regalort', 'hersteller', 'fabriknummer', 'baujahr', 'zustand',
  'beschaedigungen', 'saiten', 'gattung', 'mensur', 'besonderheiten', 'kaufwert', 'wert', 'verkauf', 'pruefen'];
const OLD_HEAD = ['Lfd. Nr.', 'Alte Signatur', 'Regal-Nr.', 'Regalort', 'Hersteller', 'Fabriknummer', 'Baujahr',
  'Zustand', 'Beschädigungen', 'Zustand der Saiten', 'Gattung', 'Mensur', 'Besonderheiten',
  'Kaufwert (DM)', 'tatsächlicher Wert (DM)', 'Verkauf', 'Lesung prüfen'];

/* ------------------------------------------------------------------ */
/* Einmalig im Editor ausführen: legt Blätter, Fotoordner und Schlüssel an */
function setup() {
  const ss = SpreadsheetApp.getActive();
  const props = PropertiesService.getScriptProperties();

  if (!props.getProperty('TOKEN')) props.setProperty('TOKEN', Utilities.getUuid());
  if (!props.getProperty('FOLDER_ID')) {
    const folder = DriveApp.createFolder('Zitherinventar Fotos');
    props.setProperty('FOLDER_ID', folder.getId());
  }

  let b = ss.getSheetByName(SHEET_BESTAND);
  if (!b) b = ss.insertSheet(SHEET_BESTAND);
  if (b.getLastRow() === 0) {
    b.getRange(1, 1, 1, COLS.length).setValues([COLS.map(c => c[1])]).setFontWeight('bold');
    b.setFrozenRows(1);
  }
  b.getRange(1, 1, b.getMaxRows(), COLS.length).setNumberFormat('@'); // alles als Text (führende Nullen im Barcode!)

  let a = ss.getSheetByName(SHEET_ALT);
  if (!a) {
    a = ss.insertSheet(SHEET_ALT);
    const rows = ALTELISTE.map(r => OLD_COLS.map(k => r[k] == null ? '' : String(r[k])));
    a.getRange(1, 1, rows.length + 1, OLD_COLS.length).setNumberFormat('@').setValues([OLD_HEAD].concat(rows));
    a.getRange(1, 1, 1, OLD_COLS.length).setFontWeight('bold');
    a.setFrozenRows(1);
  }

  Logger.log('Schlüssel für die Anwendung: ' + props.getProperty('TOKEN'));
  Logger.log('Fotoordner: https://drive.google.com/drive/folders/' + props.getProperty('FOLDER_ID'));
}

/* ------------------------------------------------------------------ */
function doGet(e) {
  const p = (e && e.parameter) || {};
  if (!tokenOk(p.token)) return json({ error: 'Zugriff verweigert: Schlüssel falsch.' });
  try {
    if (p.action === 'photo') return json(readPhoto(p.id));
    return json({ old: readOld(), records: readRecords() });
  } catch (err) {
    return json({ error: String(err && err.message || err) });
  }
}

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json({ error: 'Ungültige Anfrage.' }); }
  if (!tokenOk(body.token)) return json({ error: 'Zugriff verweigert: Schlüssel falsch.' });

  if (body.action === 'upload') {
    try { return json(uploadPhoto(body)); } catch (err) { return json({ error: String(err && err.message || err) }); }
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (body.action === 'save') return json(saveRecord(body.record || {}, body.previousBarcode, !!body.overwrite));
    if (body.action === 'delete') return json(deleteRecord(body.barcode));
    return json({ error: 'Unbekannte Aktion.' });
  } catch (err) {
    return json({ error: String(err && err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

/* ------------------------------------------------------------------ */
function tokenOk(t) {
  const token = PropertiesService.getScriptProperties().getProperty('TOKEN');
  return !!token && t === token;
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function sheet(name) {
  const s = SpreadsheetApp.getActive().getSheetByName(name);
  if (!s) throw new Error('Tabellenblatt „' + name + '“ fehlt. Bitte setup() ausführen.');
  return s;
}

function readOld() {
  const s = sheet(SHEET_ALT);
  if (s.getLastRow() < 2) return [];
  const vals = s.getRange(2, 1, s.getLastRow() - 1, OLD_COLS.length).getDisplayValues();
  return vals.filter(r => r.join('') !== '').map(r => {
    const o = {}; OLD_COLS.forEach((k, i) => o[k] = r[i]); return o;
  });
}

function readRecords() {
  const s = sheet(SHEET_BESTAND);
  if (s.getLastRow() < 2) return [];
  const vals = s.getRange(2, 1, s.getLastRow() - 1, COLS.length).getDisplayValues();
  return vals.filter(r => r[1] !== '').map(rowToRecord);
}

function rowToRecord(r) {
  const o = {};
  COLS.forEach((c, i) => o[c[0]] = r[i]);
  o.fotoIds = o.fotoIds ? o.fotoIds.split(',').filter(String) : [];
  delete o.foto;
  return o;
}

function findRow(s, code) {
  if (!code || s.getLastRow() < 2) return -1;
  const col = COLS.findIndex(c => c[0] === 'barcode') + 1;
  const vals = s.getRange(2, col, s.getLastRow() - 1, 1).getDisplayValues();
  for (let i = 0; i < vals.length; i++) if (vals[i][0] === String(code)) return i + 2;
  return -1;
}

function saveRecord(rec, previousBarcode, overwrite) {
  const code = String(rec.barcode || '').trim();
  if (!code) throw new Error('Barcode fehlt.');
  const s = sheet(SHEET_BESTAND);

  const rowNew = findRow(s, code);
  const rowPrev = previousBarcode ? findRow(s, previousBarcode) : -1;
  if (rowNew > 0 && rowNew !== rowPrev && !overwrite) {
    return { error: 'Barcode ' + code + ' ist bereits vergeben.', conflict: true };
  }

  const now = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm');
  let target = rowPrev > 0 ? rowPrev : rowNew;
  let erstellt = now;
  if (target > 0) {
    const col = COLS.findIndex(c => c[0] === 'erstellt') + 1;
    erstellt = s.getRange(target, col).getDisplayValue() || now;
  }

  const ids = (rec.fotoIds || []).filter(String);
  const values = Object.assign({}, rec, {
    barcode: code,
    zeitstempel: now,
    erstellt: erstellt,
    fotoIds: ids.join(','),
    foto: ids.map(id => 'https://drive.google.com/file/d/' + id + '/view').join('\n')
  });
  const row = COLS.map(c => values[c[0]] == null ? '' : String(values[c[0]]));

  // Barcode umbenannt und neuer Barcode existiert schon (Überschreiben bestätigt): alte Zeile entfernen
  if (rowPrev > 0 && rowNew > 0 && rowNew !== rowPrev) {
    s.getRange(rowNew, 1, 1, COLS.length).setNumberFormat('@').setValues([row]);
    s.deleteRow(rowPrev);
  } else if (target > 0) {
    s.getRange(target, 1, 1, COLS.length).setNumberFormat('@').setValues([row]);
  } else {
    target = s.getLastRow() + 1;
    s.getRange(target, 1, 1, COLS.length).setNumberFormat('@').setValues([row]);
  }
  return { record: rowToRecord(row) };
}

function deleteRecord(code) {
  const s = sheet(SHEET_BESTAND);
  const r = findRow(s, code);
  if (r < 0) return { error: 'Eintrag nicht gefunden.' };
  s.deleteRow(r); // Fotos bleiben im Drive-Ordner erhalten
  return { ok: true };
}

function folder() {
  const id = PropertiesService.getScriptProperties().getProperty('FOLDER_ID');
  if (!id) throw new Error('Fotoordner fehlt. Bitte setup() ausführen.');
  return DriveApp.getFolderById(id);
}

function uploadPhoto(body) {
  if (!body.data) throw new Error('Keine Bilddaten.');
  const blob = Utilities.newBlob(Utilities.base64Decode(body.data), body.mime || 'image/jpeg', body.name || 'foto.jpg');
  const file = folder().createFile(blob);
  return { id: file.getId() };
}

function readPhoto(id) {
  const f = DriveApp.getFileById(id);
  const fid = folder().getId();
  const parents = f.getParents();
  let inFolder = false;
  while (parents.hasNext()) if (parents.next().getId() === fid) inFolder = true;
  if (!inFolder) throw new Error('Foto gehört nicht zum Inventar.');
  const blob = f.getBlob();
  return { mime: blob.getContentType(), data: Utilities.base64Encode(blob.getBytes()) };
}
