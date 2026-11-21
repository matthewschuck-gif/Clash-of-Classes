/**
 * Clash of Classes -- Apps Script backend
 * File 61: hardening. Things that make the site harder to break into or vandalize, all checked on
 * the server (a changed page in someone's browser can't skip them).
 *
 * SECURITY ALERTS: an email to SECURITY_ALERT_EMAILS (blank = the Sheet owner) when
 *   - someone keeps typing the wrong shared password (5 wrong, and the 10-try lockout),
 *   - an account that isn't on the Staff Access list tries to sign in,
 *   - someone signs in with their Google account for the first time,
 *   - one save moves more than BIG_CHANGE_POINTS points in total.
 *   The same alert goes out at most once an hour.
 * SAVE LIMITS: a save bigger than MAX_SAVE_KB, more than 5,000 events, a point value that isn't a
 *   whole number, or a new/changed event worth more than MAX_EVENT_POINTS to one squad is refused.
 *   Rows already on the events tab are never judged, only what this save adds or changes.
 * TRACKING LIMIT: the anonymous page-open counter takes at most TRACK_PER_MINUTE calls a minute
 *   for the whole site, so nobody can flood the Usage tab or the Apps Script quota.
 * SIGN OUT: ends the session on the server (the token stops working everywhere).
 * NIGHTLY BACKUP: after 2 a.m. a copy of this Sheet goes into the "Clash Backups" Drive folder;
 *   the newest BACKUP_KEEP copies are kept, older ones go to the trash. Restore = open a copy and
 *   copy its tabs back (or File › Make a copy of the backup).
 */

const HARDENING_SETTINGS_ = [
  ['SECURITY_ALERTS', 'ON', 'ON = security alert emails (wrong passwords, refused sign-ins, first sign-ins, big point changes). OFF = none are sent; the Admin Actions tab still records everything.'],
  ['NIGHTLY_BACKUP', 'ON', 'ON = a copy of this Sheet goes to the Clash Backups folder each night after 2 a.m. OFF = no nightly copies (the menu\'s Back up this Sheet now still works).'],
  ['SECURITY_ALERT_EMAILS', '', 'Who gets security alert emails (comma separated). Blank = the Sheet owner.'],
  ['BIG_CHANGE_POINTS', '200', 'A save that adds, changes or removes more than this many points in total sends a security alert.'],
  ['MAX_EVENT_POINTS', '1000', 'Most points one new or edited event may give one squad. Bigger values are refused on the server.'],
  ['MAX_SAVE_KB', '8000', 'Largest site save the server accepts, in KB (gallery photos make saves big).'],
  ['TRACK_PER_MINUTE', '600', 'Most page-open counts the server takes per minute for the whole site.'],
  ['BACKUP_KEEP', '14', 'How many nightly backup copies of this Sheet to keep in the Clash Backups folder.'],
];

// ── ALERTS ──────────────────────────────────────────────────────────
function securityAlert_(key, subject, lines) {
  try {
    if (String(getSetting_('SECURITY_ALERTS') || 'ON').toUpperCase() === 'OFF') return false;
    const cache = CacheService.getScriptCache(), k = 'clash-alert-' + key;
    if (cache.get(k)) return false;
    cache.put(k, '1', 3600);
    const to = String(getSetting_('SECURITY_ALERT_EMAILS') || '').split(',').map(function (s) { return s.trim(); }).filter(String);
    if (!to.length) { const o = ownerEmail_(); if (o) to.push(o); }
    if (!to.length) return false;
    const esc = function (t) { return String(t == null ? '' : t).replace(/[<>&"]/g, function (c) { return { '<': '&lt;', '>': '&gt;', '&': '&amp;', '"': '&quot;' }[c]; }); };
    const when = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'EEE MMM d, h:mm a');
    const html = '<div style="font-family:Arial,sans-serif;max-width:560px;border-top:6px solid #490e6f;padding:16px">' +
      '<div style="font-size:12px;color:#490e6f;font-weight:bold;letter-spacing:1px">CLASH OF CLASSES · SECURITY</div>' +
      '<h2 style="margin:6px 0 12px;color:#222">' + esc(subject) + '</h2>' +
      lines.map(function (l) { return '<p style="margin:6px 0;color:#333">' + esc(l) + '</p>'; }).join('') +
      '<p style="margin-top:14px;color:#777;font-size:12px">' + esc(when) + '. Full record: the Admin Actions tab. Together We Climb.</p></div>';
    MailApp.sendEmail({ to: to.join(','), subject: 'Clash security: ' + subject, htmlBody: html, name: 'Clash of Classes' });
    return true;
  } catch (e) { console.warn('Security alert: ' + e.message); return false; }
}

// ── SAVE LIMITS ─────────────────────────────────────────────────────
const HARDEN_SQUADS_ = ['BLACK', 'GOLD', 'GREY', 'PURPLE'];
function eventPoints_(e) {
  return HARDEN_SQUADS_.map(function (k) { return e[k] != null && e[k] !== '' ? e[k] : (e.pts ? e.pts[k] : 0); });
}

/** Refuses a save that is too big, before it is even read. */
function checkSaveSize_(valueJson) {
  const maxKb = Math.max(100, Number(getSetting_('MAX_SAVE_KB')) || 8000);
  if (typeof valueJson !== 'string' || valueJson.length > maxKb * 1024) throw new Error('This save is too big (over ' + maxKb + ' KB). Remove some gallery photos and try again.');
}

/** Refuses a save that carries events or points that can't be right. Returns the point swing. */
function checkSave_(value, stored) {
  if (!value || typeof value !== 'object' || Array.isArray(value)) throw new Error('That save was not site data.');
  const events = value.events == null ? [] : value.events;
  if (!Array.isArray(events)) throw new Error('That save has no events list.');
  if (events.length > 5000) throw new Error('That save has too many events.');
  const maxPts = Math.max(1, Number(getSetting_('MAX_EVENT_POINTS')) || 1000);
  const was = {};
  (stored || []).forEach(function (s) { was[s.id] = s; });
  let swing = 0;
  events.forEach(function (e) {
    if (!e || typeof e !== 'object') throw new Error('That save has a broken event.');
    if (String(e.name == null ? '' : e.name).length > 300 || String(e.note == null ? '' : e.note).length > 1000) throw new Error('An event name or note is too long.');
    const now = eventPoints_(e).map(function (v) {
      const n = Number(v === '' || v == null ? 0 : v);
      if (!isFinite(n) || Math.floor(n) !== n) throw new Error('"' + String(e.name).slice(0, 60) + '" has a point value that is not a whole number.');
      return n;
    });
    const o = was[e.id];
    const before = o ? eventPoints_(o).map(function (v) { return Number(v) || 0; }) : [0, 0, 0, 0];
    const changed = !o || now.some(function (n, i) { return n !== before[i]; });
    if (changed && now.some(function (n) { return Math.abs(n) > maxPts; })) {
      throw new Error('"' + String(e.name).slice(0, 60) + '" gives one squad more than ' + maxPts + ' points. If that is right, raise MAX_EVENT_POINTS in Settings.');
    }
    now.forEach(function (n, i) { swing += Math.abs(n - before[i]); });
  });
  return swing;
}

// Removed rows count toward the swing too (only rows the page had seen; see clashSave).
function removedSwing_(stored, events, seenKey) {
  const keep = {};
  (events || []).forEach(function (e) { keep[e.id] = true; });
  let swing = 0;
  (stored || []).forEach(function (s) {
    if (keep[s.id] || !seenKey(s)) return;
    eventPoints_(s).forEach(function (v) { swing += Math.abs(Number(v) || 0); });
  });
  return swing;
}

function bigChangeAlert_(who, swing, summary) {
  const limit = Math.max(1, Number(getSetting_('BIG_CHANGE_POINTS')) || 200);
  if (swing <= limit) return;
  securityAlert_('big-' + (who.email || who.name), 'A save moved ' + swing + ' points', [
    (who.name || 'Someone') + (who.email ? ' (' + who.email + ')' : '') + ', signed in with ' + (who.via === 'google' ? 'their school account' : 'the shared password') + ', made one save that moved ' + swing + ' points in total (alert above ' + limit + ').',
    summary || '',
    'If this wasn\'t expected: Admin › System shows the latest admin actions, and last night\'s copy is in the Clash Backups folder.']);
}

// ── TRACKING LIMIT ──────────────────────────────────────────────────
function trackAllowed_() {
  try {
    const cache = CacheService.getScriptCache(), k = 'clash-track-min-' + Math.floor(Date.now() / 60000);
    const n = (Number(cache.get(k)) || 0) + 1;
    cache.put(k, String(n), 120);
    return n <= Math.max(10, Number(getSetting_('TRACK_PER_MINUTE')) || 600);
  } catch (e) { return true; }
}

// ── SIGN OUT ────────────────────────────────────────────────────────
function clashSignOut(token) {
  const who = adminSession_(token);
  if (!who) return JSON.stringify({ ok: true });
  CacheService.getScriptCache().remove('clash-admin-' + token);
  logAdminAction_(who, 'Signed out', '');
  return JSON.stringify({ ok: true });
}

// ── NIGHTLY BACKUP ──────────────────────────────────────────────────
function maybeBackup_() {
  if (String(getSetting_('NIGHTLY_BACKUP') || 'ON').toUpperCase() === 'OFF') return;
  const props = PropertiesService.getScriptProperties();
  const day = todayYmd_(), hour = Number(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'H'));
  if (hour < 2 || props.getProperty('clash-backup-day') === day) return;
  props.setProperty('clash-backup-day', day);
  backupNow_();
}

function backupNow_() {
  const props = PropertiesService.getScriptProperties();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let folder = null;
  const fid = props.getProperty('clash-backup-folder');
  if (fid) { try { folder = DriveApp.getFolderById(fid); if (folder.isTrashed()) folder = null; } catch (e) { folder = null; } }
  if (!folder) { folder = DriveApp.createFolder('Clash Backups'); props.setProperty('clash-backup-folder', folder.getId()); }
  DriveApp.getFileById(ss.getId()).makeCopy('Clash backup ' + todayYmd_(), folder);
  const keep = Math.max(1, Number(getSetting_('BACKUP_KEEP')) || 14);
  const files = [];
  const it = folder.getFiles();
  while (it.hasNext()) { const f = it.next(); if (/^Clash backup /.test(f.getName())) files.push(f); }
  files.sort(function (a, b) { return b.getDateCreated() - a.getDateCreated(); });
  files.slice(keep).forEach(function (f) { f.setTrashed(true); });
  noteRun_('backup');
  return folder.getUrl();
}

/** Sheet menu: back up now. */
function menuBackupNow_() {
  const url = backupNow_();
  try { SpreadsheetApp.getUi().alert('Backed up. The copies are in the Clash Backups folder:\n' + url); } catch (e) { console.log(url); }
}

function ensureHardeningSettings_() {
  const sheet = getSheetSafe_(SETTINGS_TAB);
  if (!sheet) return;
  const have = {};
  sheet.getDataRange().getValues().forEach(function (r) { have[String(r[0]).trim()] = true; });
  HARDENING_SETTINGS_.forEach(function (d) { if (!have[d[0]]) sheet.appendRow(d); });
}

// file-stamp (tools/stamp-files.js): lets Admin › System tell whether this file is the newest one
var FILE_STAMP_61_Hardening = '41d295703e';
