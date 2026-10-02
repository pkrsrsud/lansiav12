const TZ = 'Asia/Jakarta';
const SHEETS = {
  PARTICIPANTS: 'Peserta',
  EVENTS: 'Kegiatan',
  ATTENDANCE: 'Kehadiran'
};

const CACHE_TTL_SECONDS = { EVENTS: 30, PARTICIPANTS: 30, ATTENDANCE: 5 };

function cache_() { return CacheService.getScriptCache(); }
function cacheGetJson_(key) { try { const v = cache_().get(key); return v ? JSON.parse(v) : null; } catch (e) { return null; } }
function cachePutJson_(key, value, seconds) { try { cache_().put(key, JSON.stringify(value), seconds); } catch (e) {} }
function clearCache_(prefix) {
  try {
    if (prefix === 'all' || prefix === 'events') cache_().remove('lansia:events');
    if (prefix === 'all' || prefix === 'participants') cache_().remove('lansia:participants');
  } catch (e) {}
}
function clearAttendanceCache_(eventId) { try { cache_().remove('lansia:attendance:' + String(eventId)); } catch (e) {} }

function props_() {
  return PropertiesService.getScriptProperties();
}

function config_() {
  return {
    spreadsheetId: props_().getProperty('SPREADSHEET_ID'),
    apiSecret: props_().getProperty('API_SECRET'),
    adminPassword: props_().getProperty('ADMIN_PASSWORD'),
    templateId: props_().getProperty('DOC_TEMPLATE_ID') || '',
    docxTemplateId: props_().getProperty('DOCX_TEMPLATE_ID') || '',
    exportFolderId: props_().getProperty('EXPORT_FOLDER_ID') || ''
  };
}

function setupConfig() {
  // EDIT VALUES BELOW ONCE, THEN RUN setupConfig().
  props_().setProperties({
    SPREADSHEET_ID: 'PASTE_SPREADSHEET_ID',
    API_SECRET: 'GANTI_DENGAN_SECRET_PANJANG_ACAK',
    ADMIN_PASSWORD: 'GANTI_DENGAN_PASSWORD_ADMIN',
    DOC_TEMPLATE_ID: '',
    DOCX_TEMPLATE_ID: 'PASTE_WORD_TEMPLATE_ID',
    EXPORT_FOLDER_ID: ''
  }, true);
  Logger.log('Config tersimpan. Jangan bagikan source Apps Script ini.');
}

function setupSpreadsheet() {
  const c = config_();
  if (!c.spreadsheetId || c.spreadsheetId === 'PASTE_SPREADSHEET_ID') throw new Error('Isi SPREADSHEET_ID dulu.');
  const ss = SpreadsheetApp.openById(c.spreadsheetId);
  ensureSheet_(ss, SHEETS.PARTICIPANTS, ['id','full_name','address','is_active','created_at','updated_at']);
  ensureSheet_(ss, SHEETS.EVENTS, ['id','name','event_date','event_time','location','open_at','close_at','is_active','created_at','updated_at']);
  ensureSheet_(ss, SHEETS.ATTENDANCE, ['id','event_id','participant_id','created_at','display_order']);
  Logger.log('Spreadsheet siap: Peserta, Kegiatan, Kehadiran.');
}

function importDocxTemplate() {
  const c = config_();
  const docxId = c.docxTemplateId || c.templateId;
  if (!docxId || docxId === 'PASTE_WORD_TEMPLATE_ID') throw new Error('Isi DOCX_TEMPLATE_ID dengan ID file Word (.docx) di Google Drive.');
  const source = DriveApp.getFileById(docxId);
  const mime = source.getMimeType();
  if (mime === MimeType.GOOGLE_DOCS || mime === 'application/vnd.google-apps.document') {
    props_().setProperty('DOC_TEMPLATE_ID', source.getId());
    return json_({ok:true, id:source.getId(), url:'https://docs.google.com/document/d/' + source.getId() + '/edit', message:'File sudah merupakan Google Docs.'});
  }
  const metadata = {name: source.getName().replace(/\.[^.]+$/, '') + ' - Google Docs', mimeType:'application/vnd.google-apps.document'};
  const converted = Drive.Files.create(metadata, source.getBlob(), {fields:'id,name,mimeType'});
  if (!converted || !converted.id) throw new Error('Konversi Word ke Google Docs gagal.');
  props_().setProperty('DOC_TEMPLATE_ID', converted.id);
  Logger.log('Google Docs template ID: ' + converted.id);
  Logger.log('Buka: https://docs.google.com/document/d/' + converted.id + '/edit');
  return json_({ok:true, id:converted.id, url:'https://docs.google.com/document/d/' + converted.id + '/edit', message:'Template Word berhasil dikonversi ke Google Docs.'});
}

function ensureDocTemplate_() {
  const c = config_();
  if (c.templateId) {
    const file = DriveApp.getFileById(c.templateId);
    const mime = file.getMimeType();
    if (mime === MimeType.GOOGLE_DOCS || mime === 'application/vnd.google-apps.document') return c.templateId;
  }
  if (c.docxTemplateId && c.docxTemplateId !== 'PASTE_WORD_TEMPLATE_ID') return importDocxTemplate_().id;
  throw new Error('Template belum dikonfigurasi. Isi DOCX_TEMPLATE_ID lalu jalankan importDocxTemplate().');
}

function importDocxTemplate_() {
  const c = config_();
  const docxId = c.docxTemplateId || c.templateId;
  if (!docxId || docxId === 'PASTE_WORD_TEMPLATE_ID') throw new Error('Isi DOCX_TEMPLATE_ID terlebih dahulu.');
  const source = DriveApp.getFileById(docxId);
  const mime = source.getMimeType();
  if (mime === MimeType.GOOGLE_DOCS || mime === 'application/vnd.google-apps.document') {
    props_().setProperty('DOC_TEMPLATE_ID', source.getId());
    return {id:source.getId(), url:'https://docs.google.com/document/d/' + source.getId() + '/edit'};
  }
  const metadata = {name: source.getName().replace(/\.[^.]+$/, '') + ' - Google Docs', mimeType:'application/vnd.google-apps.document'};
  const converted = Drive.Files.create(metadata, source.getBlob(), {fields:'id,name,mimeType'});
  if (!converted || !converted.id) throw new Error('Konversi Word ke Google Docs gagal.');
  props_().setProperty('DOC_TEMPLATE_ID', converted.id);
  return {id:converted.id, url:'https://docs.google.com/document/d/' + converted.id + '/edit'};
}

function prepareGoogleDocTemplate() {
  const templateId = ensureDocTemplate_();
  const doc = DocumentApp.openById(templateId);
  const replacements = {
    'Selasa': '{{HARI}}',
    '29 September 2026': '{{TANGGAL}}',
    '15.30 WIB - Selesai': '{{WAKTU}}',
    'Ruang Tunggu Ruang Bersalin': '{{TEMPAT}}',
    'Kelas Ibu Hamil dan Prenatal Gentle Yoga': '{{ACARA}}'
  };
  replaceTemplateText_(doc.getBody(), replacements, false);
  const header = doc.getHeader();
  if (header) replaceTemplateText_(header, replacements, false);
  const footer = doc.getFooter();
  if (footer) replaceTemplateText_(footer, replacements, false);
  doc.saveAndClose();
  Logger.log('Template disiapkan dengan field dinamis: HARI, TANGGAL, WAKTU, TEMPAT, ACARA.');
}

function replaceTemplateText_(section, replacements, escapedKeys) {
  Object.keys(replacements).forEach(function(key) {
    const pattern = escapedKeys ? key : key.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    section.replaceText(pattern, replacements[key]);
  });
}

function ensureSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) sh.getRange(1,1,1,headers.length).setValues([headers]);
  else {
    const first = sh.getRange(1,1,1,headers.length).getValues()[0];
    if (first.join('|') !== headers.join('|')) sh.getRange(1,1,1,headers.length).setValues([headers]);
  }
  sh.setFrozenRows(1);
}

function doGet(e) {
  return json_({ ok: true, service: 'Lansia SMART Google Apps Script' });
}

function doPost(e) {
  try {
    const request = JSON.parse(e.postData.contents || '{}');
    const c = config_();
    if (!request.secret || request.secret !== c.apiSecret) return json_({ error: true, message: 'Secret backend tidak valid.' });
    const action = request.action || '';
    const payload = request.payload || {};
    const adminToken = request.adminToken || '';

    if (action === 'adminLogin') return adminLogin_(payload);
    if (action === 'getActiveEvent') return getActiveEvent_();
    if (action === 'searchParticipant') return searchParticipant_(payload);
    if (action === 'submitAttendance') return submitAttendance_(payload);

    requireAdmin_(adminToken);
    if (action === 'adminLoad') return adminLoad_(payload);
    if (action === 'saveParticipant') return saveParticipant_(payload);
    if (action === 'toggleParticipant') return toggleParticipant_(payload);
    if (action === 'importParticipants') return importParticipants_(payload);
    if (action === 'saveEvent') return saveEvent_(payload);
    if (action === 'toggleEvent') return toggleEvent_(payload);
    if (action === 'deleteEvent') return deleteEvent_(payload);
    if (action === 'removeAttendance') return removeAttendance_(payload);
    if (action === 'exportReport') return exportReport_(payload);

    return json_({ error: true, message: 'Action tidak dikenal.' });
  } catch (err) {
    return json_({ error: true, message: err instanceof Error ? err.message : String(err) });
  }
}

function json_(value) {
  return ContentService.createTextOutput(JSON.stringify(value)).setMimeType(ContentService.MimeType.JSON);
}

function adminLogin_(p) {
  const c = config_();
  if (!p.password || p.password !== c.adminPassword) throw new Error('Email atau password salah.');
  // Email hanya identitas; password adalah kontrol utama pada project kecil ini.
  const token = Utilities.getUuid() + Utilities.getUuid().replace(/-/g,'');
  CacheService.getScriptCache().put('admin:' + token, '1', 1800);
  return json_({ token: token });
}

function requireAdmin_(token) {
  if (!token || CacheService.getScriptCache().get('admin:' + token) !== '1') throw new Error('Sesi admin tidak valid atau sudah kedaluwarsa.');
}

function sheetData_(name) {
  const sh = SpreadsheetApp.openById(config_().spreadsheetId).getSheetByName(name);
  if (!sh) throw new Error('Sheet ' + name + ' tidak ditemukan. Jalankan setupSpreadsheet().');
  const values = sh.getDataRange().getValues();
  const headers = values.shift() || [];
  return { sh, rows: values, headers };
}

function records_(name) {
  const d = sheetData_(name);
  return d.rows.map(rowToObject_(d.headers));
}

function rowToObject_(headers) {
  return function(row) {
    const obj = {};
    headers.forEach((h,i) => obj[h] = row[i]);
    return obj;
  };
}

function iso_(value) {
  if (!value) return '';
  if (value instanceof Date) return value.toISOString();
  return String(value);
}

function bool_(value) {
  return value === true || String(value).toLowerCase() === 'true';
}

function dateOnly_(value) {
  if (!value) return '';
  if (value instanceof Date) {
    return Utilities.formatDate(value, TZ, 'yyyy-MM-dd');
  }
  const text = String(value).trim();
  const match = text.match(/(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})/);
  if (match) {
    return match[1] + '-' + ('0' + match[2]).slice(-2) + '-' + ('0' + match[3]).slice(-2);
  }
  return text.slice(0, 10);
}

function getEvents_() {
  const cached = cacheGetJson_('lansia:events');
  if (cached) return cached;
  const result = records_(SHEETS.EVENTS).map(e => ({ ...e, event_date: dateOnly_(e.event_date), open_at: iso_(e.open_at), close_at: iso_(e.close_at), is_active: bool_(e.is_active), created_at: iso_(e.created_at), updated_at: iso_(e.updated_at) }));
  cachePutJson_('lansia:events', result, CACHE_TTL_SECONDS.EVENTS);
  return result;
}

function getParticipants_() {
  const cached = cacheGetJson_('lansia:participants');
  if (cached) return cached;
  const result = records_(SHEETS.PARTICIPANTS).map(p => ({ ...p, is_active: bool_(p.is_active), created_at: iso_(p.created_at), updated_at: iso_(p.updated_at) }));
  cachePutJson_('lansia:participants', result, CACHE_TTL_SECONDS.PARTICIPANTS);
  return result;
}

function getAttendance_(eventId) {
  const cacheKey = 'lansia:attendance:' + String(eventId);
  const cached = cacheGetJson_(cacheKey);
  if (cached) return cached;
  const attendance = records_(SHEETS.ATTENDANCE).filter(a => String(a.event_id) === String(eventId)).sort((a,b) => Number(a.display_order) - Number(b.display_order));
  const participants = getParticipants_();
  const map = Object.fromEntries(participants.map(p => [p.id, p]));
  const result = attendance.map(a => ({
    id: a.id, event_id: a.event_id, participant_id: a.participant_id, created_at: iso_(a.created_at), display_order: Number(a.display_order || 0),
    participant: map[a.participant_id] ? { id: map[a.participant_id].id, full_name: map[a.participant_id].full_name, address: map[a.participant_id].address } : null
  }));
  cachePutJson_(cacheKey, result, CACHE_TTL_SECONDS.ATTENDANCE);
  return result;
}


function getActiveEvent_() {
  const events = getEvents_().filter(e => e.is_active);
  events.sort((a,b) => new Date(b.open_at).getTime() - new Date(a.open_at).getTime());
  return json_({ event: events[0] || null });
}

function normalize_(value) {
  return String(value || '').trim().replace(/\s+/g,' ').toLowerCase();
}

function searchParticipant_(p) {
  const events = getEvents_().filter(e => e.id === p.event_id && e.is_active);
  if (!events.length) return json_({ participant: null, already_present: false });
  const event = events[0];
  const now = Date.now();
  if (now < new Date(event.open_at).getTime() || now >= new Date(event.close_at).getTime()) return json_({ participant: null, already_present: false });
  const q = normalize_(p.full_name);
  const participant = getParticipants_().find(x => x.is_active && normalize_(x.full_name) === q);
  if (!participant) return json_({ participant: null, already_present: false });
  const attendance = getAttendance_(event.id);
  return json_({ participant: participant, already_present: attendance.some(a => a.participant_id === participant.id) });
}

function submitAttendance_(p) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const events = getEvents_().filter(e => e.id === p.event_id && e.is_active);
    if (!events.length) return json_({ success:false, message:'Kegiatan tidak tersedia.' });
    const event = events[0];
    const now = Date.now();
    if (now < new Date(event.open_at).getTime() || now >= new Date(event.close_at).getTime()) return json_({ success:false, message:'Pendaftaran sudah ditutup atau belum dibuka.' });
    const participants = getParticipants_();
    const participant = participants.find(x => x.id === p.participant_id && x.is_active);
    if (!participant) return json_({ success:false, message:'Peserta tidak valid.' });
    const d = sheetData_(SHEETS.ATTENDANCE);
    const existing = d.rows.find(r => String(r[1]) === String(p.event_id) && String(r[2]) === String(p.participant_id));
    if (existing) return json_({ success:false, duplicate:true, message:'Peserta ini sudah tercatat hadir.' });
    const order = d.rows.filter(r => String(r[1]) === String(p.event_id)).length + 1;
    const sh = d.sh;
    sh.appendRow([Utilities.getUuid(), p.event_id, p.participant_id, new Date(), order]);
    clearAttendanceCache_(p.event_id);
    return json_({ success:true, message:'Kehadiran berhasil dicatat.', display_order:order });
  } finally {
    lock.releaseLock();
  }
}

function adminLoad_(p) {
  const events = getEvents_();
  const participants = getParticipants_();
  const target = p.event_id || (events.find(e => e.is_active)?.id || events[0]?.id || '');
  return json_({ events, participants, attendance: target ? getAttendance_(target) : [] });
}

function findRowById_(sheetName, id) {
  const d = sheetData_(sheetName);
  const idx = d.headers.indexOf('id');
  for (let i=0;i<d.rows.length;i++) if (String(d.rows[i][idx]) === String(id)) return { sh:d.sh, row:i+2, headers:d.headers, values:d.rows[i] };
  return null;
}

function saveParticipant_(p) {
  const fullName = String(p.full_name || '').trim().replace(/\s+/g,' ');
  const address = String(p.address || '').trim();
  if (!fullName || !address) throw new Error('Nama lengkap dan alamat wajib diisi.');
  const participants = getParticipants_();
  const duplicate = participants.some(x => x.id !== p.id && x.is_active && normalize_(x.full_name) === normalize_(fullName));
  if (duplicate) throw new Error('Peserta dengan nama tersebut sudah ada.');
  const sh = SpreadsheetApp.openById(config_().spreadsheetId).getSheetByName(SHEETS.PARTICIPANTS);
  const now = new Date();
  if (p.id) {
    const row = findRowById_(SHEETS.PARTICIPANTS, p.id); if (!row) throw new Error('Peserta tidak ditemukan.');
    const values = row.values.slice();
    values[1] = fullName; values[2] = address; values[5] = now;
    row.sh.getRange(row.row,1,1,row.headers.length).setValues([values]);
    clearCache_('participants');
    return json_({ ok:true });
  }
  sh.appendRow([Utilities.getUuid(), fullName, address, true, now, now]);
  clearCache_('participants');
  return json_({ ok:true });
}

function toggleParticipant_(p) {
  const row = findRowById_(SHEETS.PARTICIPANTS, p.id); if (!row) throw new Error('Peserta tidak ditemukan.');
  const idx = row.headers.indexOf('is_active');
  row.sh.getRange(row.row, idx+1).setValue(Boolean(p.is_active));
  row.sh.getRange(row.row, row.headers.indexOf('updated_at')+1).setValue(new Date());
  clearCache_('participants');
  return json_({ ok:true });
}

function importParticipants_(p) {
  const url = String(p.url || '').trim();
  const requestedType = String(p.source_type || 'auto').toLowerCase();
  const updateAddress = p.update_address !== false;
  const reactivate = p.reactivate_existing !== false;
  if (!url) throw new Error('Tempel URL Google Spreadsheet atau Google Docs terlebih dahulu.');
  if (!['auto','spreadsheet','docs'].includes(requestedType)) throw new Error('Jenis sumber import tidak valid.');

  let parsed = null;
  const errors = [];
  const typesToTry = requestedType === 'auto' ? ['spreadsheet','docs'] : [requestedType];

  for (let i=0; i<typesToTry.length && !parsed; i++) {
    const type = typesToTry[i];
    try {
      if (type === 'spreadsheet') parsed = readParticipantsFromSpreadsheet_(url);
      else parsed = readParticipantsFromDocs_(url);
    } catch (err) {
      errors.push((err instanceof Error ? err.message : String(err)));
    }
  }

  if (!parsed) {
    throw new Error('Sumber data tidak bisa dibaca. Pastikan file adalah Google Spreadsheet/Google Docs dan akun Apps Script memiliki akses. ' + errors.join(' | '));
  }
  if (!parsed.rows.length) throw new Error('Tidak ditemukan baris peserta yang berisi nama dan alamat.');

  const result = upsertImportedParticipants_(parsed.rows, updateAddress, reactivate);
  result.source_type = parsed.source_type;
  result.source_name = parsed.source_name;
  result.source_url = parsed.source_url;
  return json_({ok:true, ...result});
}

function extractFileId_(url, type) {
  const text = String(url || '').trim();
  const patterns = type === 'spreadsheet'
    ? [/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/, /^([a-zA-Z0-9_-]{20,})$/]
    : [/\/document\/d\/([a-zA-Z0-9_-]+)/, /^([a-zA-Z0-9_-]{20,})$/];
  for (let i=0;i<patterns.length;i++) {
    const m = text.match(patterns[i]);
    if (m) return m[1];
  }
  return '';
}

function extractGid_(url) {
  const m = String(url || '').match(/[?#&]gid=(\d+)/);
  return m ? m[1] : '';
}

function headerKey_(value) {
  return normalize_(value).replace(/[^a-z0-9]/g, '');
}

function participantColumns_(headerRow) {
  const keys = headerRow.map(headerKey_);
  let nameIndex = -1;
  let addressIndex = -1;
  keys.forEach(function(key, i) {
    if (nameIndex < 0 && (key === 'nama' || key === 'namalengkap' || key === 'namapeserta' || key.indexOf('nama') === 0)) nameIndex = i;
    if (addressIndex < 0 && (key === 'alamat' || key === 'alamatlengkap' || key === 'domisili' || key.indexOf('alamat') >= 0 || key === 'jabatanalamat')) addressIndex = i;
  });
  return { nameIndex:nameIndex, addressIndex:addressIndex };
}

function cleanImportedParticipant_(nameValue, addressValue) {
  const fullName = String(nameValue || '').trim().replace(/\s+/g, ' ');
  const address = String(addressValue || '').trim().replace(/\s+/g, ' ');
  if (!fullName || !address) return null;
  const lower = normalize_(fullName);
  if (['nama','nama lengkap','nama peserta'].indexOf(lower) >= 0) return null;
  if (['alamat','alamat lengkap','domisili','jabatan/alamat'].indexOf(normalize_(address)) >= 0) return null;
  return { full_name: fullName, address: address };
}

function readParticipantsFromSpreadsheet_(url) {
  const id = extractFileId_(url, 'spreadsheet');
  if (!id) throw new Error('ID Google Spreadsheet tidak ditemukan dari URL.');
  const ss = SpreadsheetApp.openById(id);
  const gid = extractGid_(url);
  let sh = null;
  if (gid) sh = ss.getSheets().find(function(sheet) { return String(sheet.getSheetId()) === String(gid); }) || null;
  if (!sh) sh = ss.getActiveSheet() || ss.getSheets()[0];
  if (!sh) throw new Error('Google Spreadsheet tidak memiliki sheet.');

  const values = sh.getDataRange().getDisplayValues();
  if (!values.length) throw new Error('Google Spreadsheet kosong.');

  const cols = participantColumns_(values[0]);
  let start = 1;
  if (cols.nameIndex < 0 || cols.addressIndex < 0) {
    cols.nameIndex = 0;
    cols.addressIndex = 1;
    start = 0;
  }

  const rows = [];
  for (let r=start; r<values.length; r++) {
    const item = cleanImportedParticipant_(values[r][cols.nameIndex], values[r][cols.addressIndex]);
    if (item) rows.push(item);
  }

  return {
    source_type:'spreadsheet',
    source_name:ss.getName() + ' / ' + sh.getName(),
    source_url:'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/edit#gid=' + sh.getSheetId(),
    rows:dedupeImportedRows_(rows)
  };
}

function readParticipantsFromDocs_(url) {
  const id = extractFileId_(url, 'docs');
  if (!id) throw new Error('ID Google Docs tidak ditemukan dari URL.');
  const doc = DocumentApp.openById(id);
  const body = doc.getBody();
  const rows = [];
  const tables = body.getTables();

  for (let t=0; t<tables.length; t++) {
    const table = tables[t];
    let headerRow = -1;
    let cols = {nameIndex:-1, addressIndex:-1};
    const scanLimit = Math.min(table.getNumRows(), 6);
    for (let r=0; r<scanLimit; r++) {
      const values = [];
      for (let c=0; c<table.getRow(r).getNumCells(); c++) values.push(table.getCell(r,c).getText());
      const found = participantColumns_(values);
      if (found.nameIndex >= 0 && found.addressIndex >= 0) {
        headerRow = r;
        cols = found;
        break;
      }
    }
    if (headerRow < 0) continue;

    for (let r=headerRow+1; r<table.getNumRows(); r++) {
      const row = table.getRow(r);
      if (cols.nameIndex >= row.getNumCells() || cols.addressIndex >= row.getNumCells()) continue;
      const item = cleanImportedParticipant_(row.getCell(cols.nameIndex).getText(), row.getCell(cols.addressIndex).getText());
      if (item) rows.push(item);
    }
  }

  if (!rows.length) {
    const paragraphs = body.getParagraphs();
    for (let i=0; i<paragraphs.length; i++) {
      const text = paragraphs[i].getText().trim();
      if (!text || text.indexOf('|') < 0) continue;
      const parts = text.split('|');
      if (parts.length < 2) continue;
      const item = cleanImportedParticipant_(parts[0], parts.slice(1).join('|'));
      if (item) rows.push(item);
    }
  }

  return {
    source_type:'docs',
    source_name:doc.getName(),
    source_url:'https://docs.google.com/document/d/' + doc.getId() + '/edit',
    rows:dedupeImportedRows_(rows)
  };
}

function dedupeImportedRows_(rows) {
  const seen = {};
  const out = [];
  rows.forEach(function(item) {
    const key = normalize_(item.full_name);
    if (!key || seen[key]) return;
    seen[key] = true;
    out.push(item);
  });
  return out;
}

function upsertImportedParticipants_(rows, updateAddress, reactivate) {
  const d = sheetData_(SHEETS.PARTICIPANTS);
  const sh = d.sh;
  const headerIndex = Object.fromEntries(d.headers.map(function(h,i){ return [h,i]; }));
  const existing = {};
  d.rows.forEach(function(row, i) {
    const name = row[headerIndex.full_name];
    const key = normalize_(name);
    if (!key) return;
    const candidate = {row:i+2, values:row, isActive:bool_(row[headerIndex.is_active])};
    if (!existing[key] || candidate.isActive) existing[key] = candidate;
  });

  let inserted = 0, updated = 0, unchanged = 0, skipped = 0;
  const now = new Date();
  const append = [];

  rows.forEach(function(item) {
    if (!item || !item.full_name || !item.address) { skipped++; return; }
    const key = normalize_(item.full_name);
    const found = existing[key];
    if (!found) {
      const row = [];
      row[headerIndex.id] = Utilities.getUuid();
      row[headerIndex.full_name] = item.full_name;
      row[headerIndex.address] = item.address;
      row[headerIndex.is_active] = true;
      row[headerIndex.created_at] = now;
      row[headerIndex.updated_at] = now;
      append.push(row);
      existing[key] = {row:-1, values:row, isActive:true};
      inserted++;
      return;
    }

    const current = found.values.slice();
    let changed = false;
    if (updateAddress && String(current[headerIndex.address] || '').trim() !== item.address) {
      current[headerIndex.address] = item.address;
      changed = true;
    }
    if (reactivate && !bool_(current[headerIndex.is_active])) {
      current[headerIndex.is_active] = true;
      changed = true;
    }
    if (changed) {
      current[headerIndex.updated_at] = now;
      if (found.row > 0) sh.getRange(found.row, 1, 1, d.headers.length).setValues([current]);
      found.values = current;
      found.isActive = bool_(current[headerIndex.is_active]);
      updated++;
    } else unchanged++;
  });

  if (append.length) sh.getRange(sh.getLastRow()+1, 1, append.length, d.headers.length).setValues(append);
  clearCache_('participants');
  return {total:rows.length, inserted:inserted, updated:updated, unchanged:unchanged, skipped:skipped};
}

function parseLocalDateTime_(value) {
  return Utilities.parseDate(String(value), TZ, "yyyy-MM-dd'T'HH:mm");
}

function saveEvent_(p) {
  const name = String(p.name || '').trim();
  const eventDate = String(p.event_date || '').trim();
  const eventTime = String(p.event_time || '').trim();
  const location = String(p.location || '').trim();
  if (!name || !eventDate || !location || !p.open_at || !p.close_at) throw new Error('Nama, tanggal, tempat, buka, dan tutup pendaftaran wajib diisi.');
  const openAt = parseLocalDateTime_(p.open_at);
  const closeAt = parseLocalDateTime_(p.close_at);
  if (closeAt.getTime() <= openAt.getTime()) throw new Error('Waktu tutup harus setelah waktu buka.');
  const sh = SpreadsheetApp.openById(config_().spreadsheetId).getSheetByName(SHEETS.EVENTS);
  const now = new Date();
  if (Boolean(p.is_active)) {
    const d = sheetData_(SHEETS.EVENTS);
    const activeCol = d.headers.indexOf('is_active') + 1;
    d.rows.forEach((r,i) => { if (String(r[0]) !== String(p.id || '')) d.sh.getRange(i+2, activeCol).setValue(false); });
  }
  const openIso = openAt.toISOString();
  const closeIso = closeAt.toISOString();
  if (p.id) {
    const row = findRowById_(SHEETS.EVENTS, p.id); if (!row) throw new Error('Kegiatan tidak ditemukan.');
    const values = row.values.slice();
    values[1]=name; values[2]=eventDate; values[3]=eventTime; values[4]=location; values[5]=openIso; values[6]=closeIso; values[7]=Boolean(p.is_active); values[9]=now;
    row.sh.getRange(row.row,1,1,row.headers.length).setValues([values]);
    clearCache_('events');
    return json_({ event: getEvents_().find(e=>e.id===p.id) });
  }
  const id = Utilities.getUuid();
  sh.appendRow([id,name,eventDate,eventTime,location,openIso,closeIso,Boolean(p.is_active),now,now]);
  clearCache_('events');
  return json_({ event: getEvents_().find(e=>e.id===id) });
}

function toggleEvent_(p) {
  const row = findRowById_(SHEETS.EVENTS,p.id); if (!row) throw new Error('Kegiatan tidak ditemukan.');
  const activeIdx = row.headers.indexOf('is_active') + 1;
  const target = Boolean(p.is_active);
  if (target) {
    const d = sheetData_(SHEETS.EVENTS);
    d.rows.forEach((r,i)=>{ if(String(r[0])!==String(p.id)) d.sh.getRange(i+2,activeIdx).setValue(false); });
  }
  row.sh.getRange(row.row,activeIdx).setValue(target);
  row.sh.getRange(row.row,row.headers.indexOf('updated_at')+1).setValue(new Date());
  clearCache_('events');
  return json_({ ok:true });
}

function deleteEvent_(p) {
  const row = findRowById_(SHEETS.EVENTS,p.id); if (!row) throw new Error('Kegiatan tidak ditemukan.');
  const att = sheetData_(SHEETS.ATTENDANCE);
  for (let i=att.rows.length-1;i>=0;i--) if (String(att.rows[i][1])===String(p.id)) att.sh.deleteRow(i+2);
  row.sh.deleteRow(row.row);
  clearCache_('events');
  clearAttendanceCache_(p.id);
  return json_({ ok:true });
}

function removeAttendance_(p) {
  const row = findRowById_(SHEETS.ATTENDANCE,p.id); if (!row) throw new Error('Data kehadiran tidak ditemukan.');
  const eventId = row.values[1];
  row.sh.deleteRow(row.row);
  clearAttendanceCache_(eventId);
  // Re-number remaining rows for the event.
  const d = sheetData_(SHEETS.ATTENDANCE);
  let n=1;
  const targetRows=[];
  d.rows.forEach((r,i)=>{ if(String(r[1])===String(eventId)) targetRows.push(i+2); });
  targetRows.forEach(r=>{ d.sh.getRange(r,5).setValue(n++); });
  return json_({ ok:true });
}

function exportReport_(p) {
  const event = getEvents_().find(e=>e.id===p.event_id);
  if (!event) throw new Error('Kegiatan tidak ditemukan.');
  const attendance = getAttendance_(event.id);
  if (!attendance.length) throw new Error('Belum ada peserta hadir.');
  const c = config_();
  const templateId = ensureDocTemplate_();
  const template = DriveApp.getFileById(templateId);
  const name = 'Daftar Hadir - ' + sanitizeFile_(event.name) + ' - ' + event.event_date;
  const folder = c.exportFolderId ? DriveApp.getFolderById(c.exportFolderId) : template.getParents().hasNext() ? template.getParents().next() : DriveApp.getRootFolder();
  const copy = template.makeCopy(name, folder);
  fillTemplate_(copy.getId(), event, attendance);

  if (p.format === 'pdf') {
    const pdfBlob = DriveApp.getFileById(copy.getId()).getAs(MimeType.PDF).setName(name + '.pdf');
    const pdfFile = folder.createFile(pdfBlob);
    return json_({ url: pdfFile.getUrl(), docUrl: 'https://docs.google.com/document/d/' + copy.getId() + '/edit', pdfUrl: pdfFile.getUrl() });
  }
  return json_({ url: 'https://docs.google.com/document/d/' + copy.getId() + '/edit' });
}

function sanitizeFile_(value) {
  return String(value).replace(/[^\p{L}\p{N}\s_-]/gu,'').trim().replace(/\s+/g,'-').slice(0,80);
}

function dayId_(dateValue) {
  const days = ['Minggu','Senin','Selasa','Rabu','Kamis','Jumat','Sabtu'];
  return days[Number(Utilities.formatDate(new Date(dateValue + 'T00:00:00+07:00'), TZ, 'u')) % 7];
}

function monthId_(m) {
  return ['Januari','Februari','Maret','April','Mei','Juni','Juli','Agustus','September','Oktober','November','Desember'][m-1];
}

function formatTanggalId_(value) {
  const d = new Date(value + 'T00:00:00+07:00');
  return Utilities.formatDate(d, TZ, 'd') + ' ' + monthId_(Number(Utilities.formatDate(d, TZ, 'M'))) + ' ' + Utilities.formatDate(d, TZ, 'yyyy');
}

function fillTemplate_(docId, event, attendance) {
  const doc = DocumentApp.openById(docId);
  const body = doc.getBody();

  const replacements = {
    '{{HARI}}': dayId_(event.event_date),
    '{{TANGGAL}}': formatTanggalId_(event.event_date),
    '{{WAKTU}}': event.event_time || '',
    '{{TEMPAT}}': event.location || '',
    '{{ACARA}}': event.name || ''
  };

  Object.keys(replacements).forEach(k => body.replaceText(k.replace(/[{}]/g,'\\$&'), replacements[k]));
  const header = doc.getHeader();
  if (header) Object.keys(replacements).forEach(k => header.replaceText(k.replace(/[{}]/g,'\\$&'), replacements[k]));
  const footer = doc.getFooter();
  if (footer) Object.keys(replacements).forEach(k => footer.replaceText(k.replace(/[{}]/g,'\\$&'), replacements[k]));

  // Template asli berisi beberapa blok daftar hadir lama (masing-masing dengan
  // tanda tangan). Untuk satu kegiatan, gunakan blok pertama sebagai master.
  // Isi setelah tanda tangan blok pertama dihapus agar tidak muncul blok lama
  // berulang, sementara kop, keterangan kegiatan, dan tanda tangan tetap ada.
  const tables = body.getTables();
  const dataTables = tables.filter(t => isAttendanceTable_(t));
  if (!dataTables.length) throw new Error('Tabel daftar hadir tidak ditemukan di template.');

  const masterTable = dataTables[0];
  trimTemplateAfterFirstSignature_(body, masterTable);

  // Bersihkan semua baris data yang ada di tabel master, tetapi pertahankan
  // dua baris header + satu baris contoh sebagai cetakan format data. Baris
  // contoh ini kemudian dicopy supaya border, tinggi, dan format tabel tetap
  // mengikuti template saat jumlah peserta bertambah.
  while (masterTable.getNumRows() > 3) {
    masterTable.removeRow(masterTable.getNumRows() - 1);
  }

  const seedRow = masterTable.getRow(2);

  // Tambahkan/isi baris sesuai jumlah kehadiran. Tidak ada lagi batas 120
  // peserta. Google Docs akan meneruskan tabel ke halaman berikutnya otomatis.
  attendance.forEach((entry, index) => {
    const row = index === 0 ? seedRow : masterTable.appendTableRow(seedRow.copy());
    const number = Number(entry.display_order || (index + 1));
    row.getCell(0).setText(String(number));
    row.getCell(1).setText(String(entry.participant?.full_name || ''));
    row.getCell(2).setText(String(entry.participant?.address || ''));
    // Kolom HADIR dan MEWAKILI sengaja dikosongkan.
    row.getCell(3).setText('');
    row.getCell(4).setText('');
  });

  doc.saveAndClose();
}

function isAttendanceTable_(table) {
  if (table.getNumRows() < 2 || table.getRow(0).getNumCells() < 5) return false;
  const h0 = String(table.getCell(0,0).getText()).trim().toUpperCase();
  const h1 = String(table.getCell(0,1).getText()).trim().toUpperCase();
  return h0 === 'NO' && h1 === 'NAMA';
}

function trimTemplateAfterFirstSignature_(body, masterTable) {
  // Cari tabel daftar hadir kedua. Konten setelah tabel kedua dan seterusnya
  // merupakan blok daftar hadir lama beserta tanda tangan berikutnya.
  const allTables = body.getTables();
  let secondDataTable = null;
  for (let i = 0; i < allTables.length; i++) {
    if (allTables[i] === masterTable) {
      for (let j = i + 1; j < allTables.length; j++) {
        if (isAttendanceTable_(allTables[j])) {
          secondDataTable = allTables[j];
          break;
        }
      }
      break;
    }
  }

  if (!secondDataTable) return;

  const secondTableIndex = body.getChildIndex(secondDataTable);
  // Hapus semua elemen sesudah tanda tangan pertama dengan menemukan paragraf
  // terakhir sebelum tabel kedua. Struktur template asli memiliki blok tanda
  // tangan tepat sebelum tabel kedua.
  const lastKeepIndex = Math.max(0, secondTableIndex - 1);
  for (let i = body.getNumChildren() - 1; i > lastKeepIndex; i--) {
    body.getChild(i).removeFromParent();
  }
}

