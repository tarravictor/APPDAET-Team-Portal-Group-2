/**
 * Project 06 Team Portal: submissions backend (Google Apps Script).
 *
 * Deploy: Deploy → New deployment → Web app → Execute as: Me, Who has access: Anyone.
 * Run setup() once first so Google asks for Drive and Sheets permission.
 *
 * Every submission is saved in the member's Drive folder (files) and logged in
 * the "Project 06 Submissions" sheet and SUBMISSIONS_LOG.md (date and time, PH time).
 */

var ROOT = '1kdE8-4N8xwryf6zu5rShL_P4kyEP05Ix';
var FOLDERS = {
  armas: '1uX9z4rQu_YVDsI6BqfcYYYBpZbBOdy3t',
  manalili: '1-FPA8Dn6A-ez6fwgSf4icXiPVbZl80Wy',
  diego: '1vIMYtHZhU62jbqSpRqF2Tm1BBQPpt20P',
  tumulac: '1By3FZKTVJwRrkgjjC0qKzFUj6DAtuRiP',
  vik: '1HOO6iKbrLaKwZbUrYnkPhzJQIqJ2kNJb'
};
var NAMES = { armas: 'Armas', manalili: 'Manalili', diego: 'Diego', tumulac: 'Tumulac', vik: 'Vik' };
var SLOTS = {
  manalili: ['pr1-scope'], armas: ['pr1-algo'], tumulac: ['pr1-decomp', 'pr1-report'],
  diego: ['pr1-env', 'pr1-wire'], vik: ['pr1-repo', 'pr1-gh', 'pr1-pm', 'pr1-slides']
};
var EXT_OK = ['pdf', 'png', 'jpg', 'jpeg', 'gif', 'webp', 'mp4', 'py', 'txt', 'md', 'csv', 'json', 'docx', 'pptx', 'xlsx', 'zip'];
var HEAD = ['id', 'at', 'member', 'slot', 'kind', 'fileName', 'link', 'url', 'token', 'removed'];
var TZ = 'Asia/Manila';

function setup() {
  sheet_();
  DriveApp.getFolderById(ROOT).getName();
  Logger.log('Ready. Sheet: ' + SpreadsheetApp.openById(prop_('SHEET_ID')).getUrl());
}

function prop_(k, v) {
  var p = PropertiesService.getScriptProperties();
  if (v !== undefined) p.setProperty(k, v);
  return p.getProperty(k);
}

function sheet_() {
  var id = prop_('SHEET_ID');
  if (id) { try { return SpreadsheetApp.openById(id).getSheets()[0]; } catch (e) {} }
  var ss = SpreadsheetApp.create('Project 06 Submissions');
  DriveApp.getFileById(ss.getId()).moveTo(DriveApp.getFolderById(ROOT));
  var sh = ss.getSheets()[0];
  sh.setName('Submissions');
  sh.appendRow(HEAD);
  sh.setFrozenRows(1);
  // The submissions already in the Claude portal, carried over.
  [
    ['2026-11-08T10:21:21.280Z', 'vik', 'pr1-repo', 'https://github.com/tarravictor/Game-of-Life-Final-Project-Group-2'],
    ['2026-11-08T10:21:43.895Z', 'vik', 'pr1-pm', 'https://drive.google.com/file/d/1AcudlZVBOmv_lqjbJc0nu5vmb1Mg8_UA/view'],
    ['2026-11-08T10:34:35.830Z', 'vik', 'pr1-gh', 'https://drive.google.com/file/d/1vmMuJMojaKHbQ352lO2j4iqLMapWe_qP/view'],
    ['2026-11-08T10:46:07.769Z', 'vik', 'pr1-slides', 'https://canva.link/wt84adyla8p85z9']
  ].forEach(function (r) { sh.appendRow([Utilities.getUuid(), r[0], r[1], r[2], 'link', '', r[3], r[3], '', '']); });
  prop_('SHEET_ID', ss.getId());
  return sh;
}

function rows_() {
  var v = sheet_().getDataRange().getValues();
  var out = [];
  for (var i = 1; i < v.length; i++) {
    var o = {};
    HEAD.forEach(function (h, j) { o[h] = v[i][j]; });
    o._row = i + 1;
    if (o.at instanceof Date) o.at = o.at.toISOString();
    out.push(o);
  }
  return out;
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  var subs = rows_().filter(function (r) { return !r.removed; }).map(function (r) {
    return { id: r.id, at: r.at, member: r.member, slot: r.slot, kind: r.kind, fileName: r.fileName, link: r.link, url: r.url };
  });
  return json_({ ok: true, subs: subs, config: { pr1: '2026-11-11' } });
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var b = JSON.parse(e.postData.contents);
    if (b.action === 'remove') return json_(remove_(b));
    return json_(submit_(b));
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

function submit_(b) {
  if (!FOLDERS[b.member]) return { ok: false, error: 'Unknown member.' };
  if ((SLOTS[b.member] || []).indexOf(b.slot) < 0) return { ok: false, error: 'Unknown item.' };
  var at = new Date().toISOString();
  var id = Utilities.getUuid(), token = Utilities.getUuid();
  var row;
  if (b.kind === 'link') {
    var link = String(b.link || '').trim();
    if (!/^https?:\/\/\S+\.\S+/i.test(link) || link.length > 2000) return { ok: false, error: 'Paste a full link starting with https://' };
    row = [id, at, b.member, b.slot, 'link', '', link, link, token, ''];
  } else if (b.kind === 'file') {
    var name = String(b.fileName || 'file').replace(/[\\\/:*?"<>|]/g, '_').slice(0, 150);
    var ext = (name.split('.').pop() || '').toLowerCase();
    if (EXT_OK.indexOf(ext) < 0) return { ok: false, error: 'This file type isn\'t accepted. Export it as PDF.' };
    var bytes = Utilities.base64Decode(String(b.data || ''));
    if (!bytes.length) return { ok: false, error: 'The file is empty.' };
    if (bytes.length > 20 * 1024 * 1024) return { ok: false, error: 'The file is over 20 MB. Paste a Drive link instead.' };
    var stamp = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH.mm');
    var blob = Utilities.newBlob(bytes, String(b.mime || 'application/octet-stream'), stamp + ' - ' + name);
    var f = DriveApp.getFolderById(FOLDERS[b.member]).createFile(blob);
    try { f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (x) { /* school policy may block; owner still has it */ }
    row = [id, at, b.member, b.slot, 'file', name, '', f.getUrl(), token, ''];
  } else {
    return { ok: false, error: 'Nothing to submit.' };
  }
  sheet_().appendRow(row);
  writeLog_();
  return { ok: true, id: id, token: token, url: row[7], at: at };
}

function remove_(b) {
  var r = rows_().filter(function (x) { return x.id === b.id; })[0];
  if (!r || r.removed) return { ok: true };
  if (!b.token || String(r.token) !== String(b.token)) return { ok: false, error: 'Only the person who submitted this can remove it (from the same browser). Ask Vik.' };
  sheet_().getRange(r._row, HEAD.indexOf('removed') + 1).setValue(new Date().toISOString());
  writeLog_();
  return { ok: true };
}

function writeLog_() {
  var live = rows_().filter(function (r) { return !r.removed; });
  var fmt = function (iso, p) { return Utilities.formatDate(new Date(iso), TZ, p); };
  var lines = ['# Project 06 · Submissions log', '',
    '_Game of Life with Music · APPDAET · generated automatically by the Team Portal. All times are Philippine time (UTC+8)._', '',
    '- **Last updated:** ' + Utilities.formatDate(new Date(), TZ, 'EEE, MMM d, yyyy h:mm a'),
    '- **Total submissions:** ' + live.length, '',
    '| # | Date uploaded | Time (PHT) | Member | Item | File or link |', '|---|---|---|---|---|---|'];
  live.forEach(function (r, i) {
    var what = r.kind === 'file' ? '[' + r.fileName + '](' + r.url + ')' : '[' + r.link + '](' + r.link + ')';
    lines.push('| ' + (i + 1) + ' | ' + fmt(r.at, 'MMM d, yyyy') + ' | ' + fmt(r.at, 'h:mm a') + ' | ' + (NAMES[r.member] || r.member) + ' | ' + r.slot + ' | ' + what + ' |');
  });
  var text = lines.join('\n') + '\n';
  var logId = prop_('LOG_ID');
  if (logId) { try { DriveApp.getFileById(logId).setContent(text); return; } catch (e) {} }
  var f = DriveApp.getFolderById(ROOT).createFile('SUBMISSIONS_LOG.md', text, 'text/markdown');
  prop_('LOG_ID', f.getId());
}
