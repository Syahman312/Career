// NextRole (formerly Ladder) – Google Apps Script backend, version 2 (adds result slip uploads).
// Sheets sync, Calendar (exam + goal dates), daily Drive backups and an error log.
//
// Setup:
// 1) Make a new Google Sheet > Extensions > Apps Script. Paste this file.
// 2) Put your own long secret word in TOKEN below (the same one you type in NextRole > Settings).
// 3) Choose "authorize" in the toolbar and click Run once.
// 4) Deploy > New deployment > Web app. Execute as: Me. Who has access: Anyone. Copy the /exec link into NextRole.
// After edits: Deploy > Manage deployments > pencil icon > Version: New version > Deploy.
const TOKEN = 'CHANGE_ME';
const VERSION = 2;
const CAL_NAME = 'Career';

function authorize() {
  LockService.getScriptLock(); CacheService.getScriptCache(); SpreadsheetApp.getActiveSpreadsheet();
  CalendarApp.getDefaultCalendar(); DriveApp.getRootFolder();
}

function doPost(e) {
  let b;
  try { b = JSON.parse(e.postData.contents); } catch (x) { return out({ error: 'Bad request' }); }
  const cache = CacheService.getScriptCache();
  const fails = +(cache.get('fails') || 0);
  if (fails >= 20) return out({ error: 'Too many wrong secret words. Locked for 10 minutes.' });
  if (TOKEN === 'CHANGE_ME') return out({ error: 'Set your secret word in the first lines of Code.gs' });
  if (typeof b.token !== 'string' || b.token !== TOKEN) {
    cache.put('fails', String(fails + 1), 600);
    logErr('auth', 'Wrong secret word');
    return out({ error: 'Wrong token' });
  }
  try { return out(route(b)); }
  catch (x) { logErr(b.action || 'sync', (x && x.stack) || x); return out({ error: String((x && x.message) || x) }); }
}

function route(b) {
  switch (b.action) {
    case 'health': return health();
    case 'backups': return listBackups();
    case 'getbackup': return getBackup(b.fid);
    case 'upload': return uploadFile(b);
    case 'getfile': return getFile(b.fid);
    case 'delfile': return delFile(b.fid);
    case 'snap': return withLock(() => ({ name: saveBackup(Object.values(readData(dataSheet())), 'before-restore') }));
    default: return withLock(() => syncAll(b));
  }
}

function withLock(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(25000)) return { error: 'Busy, try again' };
  try { return fn(); } finally { lock.releaseLock(); }
}

// ---------- Sheets sync ----------
function dataSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName('Data') || ss.insertSheet('Data', 0);
}
function readData(sh) {
  const m = {};
  sh.getDataRange().getValues().forEach(r => {
    if (!r[0] || r[0] === 'id') return;
    try { m[r[0]] = JSON.parse(r[1]); } catch (x) { logErr('sheet', 'Skipped an unreadable row: ' + r[0]); }
  });
  return m;
}
function syncAll(b) {
  const sh = dataSheet(), m = readData(sh);
  (Array.isArray(b.items) ? b.items : []).forEach(i => {
    if (!i || typeof i.id !== 'string' || !/^[\w.:-]{1,80}$/.test(i.id)) return;
    if (!m[i.id] || (+m[i.id].updated || 0) < (+i.updated || 0)) m[i.id] = i;
  });
  const all = Object.values(m);
  sh.clear();
  sh.appendRow(['id', 'json (do not edit)', 'type', 'date', 'title', 'detail', 'progress / score', 'deleted']);
  if (all.length) sh.getRange(2, 1, all.length, 8).setValues(all.map(i => [
    i.id, JSON.stringify(i), i.kind,
    i.date || i.due || i.start || '',
    i.title || i.name || i.n || i.qual || (i.kind === 'sit' ? i.c + ' attempt' : ''),
    i.impact || i.next || i.st || i.res || i.co || i.inst || i.type || '',
    i.pct !== undefined ? i.pct : (i.score || i.hrs || i.lvl || ''),
    i.deleted ? 'yes' : '']));
  sh.setFrozenRows(1);
  let cal = '';
  if (b.events) { try { syncCal(b.events); cal = 'ok'; } catch (x) { cal = String((x && x.message) || x); logErr('calendar', x); } }
  let bk = '';
  try { bk = autoBackup(all); } catch (x) { logErr('backup', x); }
  return { items: all, cal: cal, bk: bk, v: VERSION };
}

// ---------- Calendar ----------
function syncCal(ev) {
  const cal = CalendarApp.getCalendarsByName(CAL_NAME)[0] || CalendarApp.createCalendar(CAL_NAME);
  const P = PropertiesService.getScriptProperties();
  const map = JSON.parse(P.getProperty('cal') || '{}'), seen = {};
  (Array.isArray(ev) ? ev : []).forEach(e => {
    if (!e || !/^\d{4}-\d{2}-\d{2}$/.test(e.date)) return;
    const title = String(e.title || '').slice(0, 200), key = String(e.key || '').slice(0, 120);
    seen[key] = 1;
    const m = map[key];
    if (m && m.d === e.date && m.t === title) return;
    if (m) { try { cal.getEventById(m.id).deleteEvent(); } catch (x) {} }
    const x = cal.createAllDayEvent(title, new Date(e.date + 'T00:00:00'), { description: 'Added by NextRole' });
    x.removeAllReminders(); x.addPopupReminder(7 * 24 * 60); x.addPopupReminder(24 * 60);
    map[key] = { id: x.getId(), d: e.date, t: title };
  });
  Object.keys(map).forEach(k => {
    if (!seen[k]) { try { cal.getEventById(map[k].id).deleteEvent(); } catch (x) {} delete map[k]; }
  });
  P.setProperty('cal', JSON.stringify(map));
}

// ---------- Daily backups: Drive / Ladder Backups ----------
// Folder and file names keep the old "Ladder" name so existing backups still restore.
function topFolder(name) {
  const it = DriveApp.getFoldersByName(name);
  return it.hasNext() ? it.next() : DriveApp.createFolder(name);
}
function backupFolder() { return topFolder('Ladder Backups'); }
function stamp(fmt) { return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), fmt); }
// Refreshes today's backup at most every 30 minutes.
function autoBackup(all) {
  const P = PropertiesService.getScriptProperties();
  if (Date.now() - +(P.getProperty('bkAt') || 0) < 30 * 60 * 1000) return P.getProperty('bkDay') || '';
  saveBackup(all, '');
  const day = stamp('yyyy-MM-dd');
  P.setProperty('bkAt', String(Date.now())); P.setProperty('bkDay', day);
  return day;
}
function saveBackup(all, label) {
  const fo = backupFolder();
  const name = 'ladder-backup-' + stamp('yyyy-MM-dd') + (label ? '-' + label + '-' + stamp('HHmm') : '') + '.json';
  const body = JSON.stringify({ v: 1, app: 'Ladder', saved: new Date().toISOString(), items: all });
  const it = fo.getFilesByName(name);
  if (it.hasNext()) it.next().setContent(body); else fo.createFile(name, body, 'application/json');
  prune(fo);
  return name;
}
// Keeps the 40 newest backups; older ones go to the Drive bin.
function prune(fo) {
  const files = [], it = fo.getFiles();
  while (it.hasNext()) { const f = it.next(); if (/^ladder-backup-.*\.json$/.test(f.getName())) files.push(f); }
  files.sort((a, b) => b.getDateCreated() - a.getDateCreated()).slice(40).forEach(f => f.setTrashed(true));
}
function listBackups() {
  const fo = backupFolder(), it = fo.getFiles(), o = [];
  while (it.hasNext()) {
    const f = it.next();
    if (/^ladder-backup-.*\.json$/.test(f.getName())) o.push({ fid: f.getId(), name: f.getName(), at: f.getLastUpdated().getTime(), size: f.getSize() });
  }
  o.sort((a, b) => b.at - a.at);
  return { backups: o.slice(0, 40) };
}
function getBackup(fid) {
  const f = DriveApp.getFileById(String(fid)), ps = f.getParents();
  if (!ps.hasNext() || ps.next().getId() !== backupFolder().getId()) return { error: 'Not a Ladder backup' };
  return { data: f.getBlob().getDataAsString() };
}

// ---------- Result slips: Drive / NextRole Results ----------
function resultsFolder() { return topFolder('NextRole Results'); }
function inResults(f) { const id = resultsFolder().getId(), ps = f.getParents(); while (ps.hasNext()) if (ps.next().getId() === id) return true; return false; }
function uploadFile(b) {
  const name = String(b.name || 'result').replace(/[\\/:*?"<>|]/g, ' ').slice(0, 150);
  const mime = /^(application\/pdf|image\/[\w.+-]+)$/.test(String(b.mime)) ? b.mime : 'application/pdf';
  const bytes = Utilities.base64Decode(String(b.data || ''));
  if (!bytes.length) return { error: 'Empty file' };
  if (bytes.length > 10 * 1024 * 1024) return { error: 'File is larger than 10 MB' };
  const f = resultsFolder().createFile(Utilities.newBlob(bytes, mime, name));
  return { fid: f.getId(), url: f.getUrl() };
}
function getFile(fid) {
  const f = DriveApp.getFileById(String(fid));
  if (!inResults(f)) return { error: 'Not a NextRole result file' };
  return { data: Utilities.base64Encode(f.getBlob().getBytes()), mime: f.getMimeType(), name: f.getName() };
}
function delFile(fid) {
  const f = DriveApp.getFileById(String(fid));
  if (!inResults(f)) return { error: 'Not a NextRole result file' };
  f.setTrashed(true);
  return { ok: 1 };
}

// ---------- Connection check ----------
function health() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), sh = dataSheet(), P = PropertiesService.getScriptProperties();
  const bf = DriveApp.getFoldersByName('Ladder Backups');
  const lg = ss.getSheetByName('Log'), n = lg ? Math.max(0, lg.getLastRow() - 1) : 0;
  return {
    v: VERSION,
    sheet: { name: ss.getName(), url: ss.getUrl(), rows: Math.max(0, sh.getLastRow() - 1) },
    backups: bf.hasNext() ? { url: bf.next().getUrl(), last: P.getProperty('bkDay') || '' } : null,
    calendar: !!CalendarApp.getCalendarsByName(CAL_NAME)[0],
    errors: n
  };
}

// ---------- Error log: a "Log" tab in this Google Sheet ----------
function logErr(area, msg) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sh = ss.getSheetByName('Log');
    if (!sh) { sh = ss.insertSheet('Log', ss.getNumSheets()); sh.appendRow(['time', 'area', 'message']); }
    sh.appendRow([new Date(), String(area), String(msg).slice(0, 500)]);
    if (sh.getLastRow() > 600) sh.deleteRows(2, sh.getLastRow() - 500);
  } catch (x) {}
}

function out(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
