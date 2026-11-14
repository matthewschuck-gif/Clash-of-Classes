/**
 * Clash of Classes -- Apps Script backend
 * File 40: Service. A student tells us about service or a kind act they did, with a reflection, and
 * earns points for their squad once staff approve it.
 *   - On My Squad (800#): what they did, when and where, a reflection (what it meant, what they
 *     learned), and evidence if they have it: a link, or a photo / PDF from their Chromebook
 *     (saved to a private Drive folder, SERVICE_FOLDER_ID; only staff open it).
 *   - Admin › ✅ Service: approve for 5, 10 or 20 points (SERVICE_POINTS), ask for more information
 *     (the student sees the question on My Squad and answers it there; it comes back to the list),
 *     or deny with a note.
 *   - Approved = an events row "Service: Ann L." (type service) for the squad.
 *   - A student can have SERVICE_OPEN (3) waiting at once.
 * SERVICE SUBMISSIONS tab (staff only): ID | Submitted | 800# | Student | Squad | What | When / Where |
 *   Reflection | Link | File | Status (Waiting, More info, Approved, Denied) | Points | Staff Note |
 *   Student Reply | Decided By | Decided At.
 */

const SERVICE_HEADERS_ = ['ID', 'Submitted', '800#', 'Student', 'Squad', 'What', 'When / Where', 'Reflection', 'Link', 'File', 'Status', 'Points', 'Staff Note', 'Student Reply', 'Decided By', 'Decided At'];
const MAX_SERVICE_BYTES_ = 6 * 1024 * 1024;

function ensureServiceTab_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getSheetByName('Service Submissions')) return;
  const sh = ss.insertSheet('Service Submissions');
  sh.getRange(1, 1, 1, SERVICE_HEADERS_.length).setValues([SERVICE_HEADERS_]).setFontWeight('bold').setBackground('#490e6f').setFontColor('#ffffff');
  sh.setFrozenRows(1);
}

function servicePoints_() { const l = listOf_(getSetting_('SERVICE_POINTS') || '5, 10, 20').map(Number).filter(function (n) { return n > 0; }); return l.length ? l : [5, 10, 20]; }

function serviceFolder_() {
  const id = driveIdFrom_(getSetting_('SERVICE_FOLDER_ID'));
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) { /* make one */ } }
  const folder = DriveApp.createFolder('Clash of Classes — Service evidence (private)');
  setSettingValue_('SERVICE_FOLDER_ID', folder.getId());
  return folder;
}

function serviceStudent_(id) {
  const sid = id800_(id);
  const r = sheetRows_('Students').filter(function (x) { return id800_(x['800#']) === sid; })[0];
  if (r) return { id: sid, first: String(r['First Name'] || '').trim(), last: String(r['Last Name'] || '').trim(), squad: squadOf_(r['Squad']) };
  const b = sheetRows_(getSetting_('STUDENT_BOARD_TAB') || 'Student Board').filter(function (x) { return id800_(x['800#']) === sid; })[0];
  if (b) { const p = String(b['Name'] || '').trim().split(/\s+/); return { id: sid, first: p[0] || '', last: p.slice(1).join(' '), squad: squadOf_(b['Squad']) }; }
  return null;
}

function cleanLink_(v) {
  const s = String(v || '').trim();
  if (!s) return '';
  if (!/^https:\/\/[^\s<>"]+$/i.test(s)) throw new Error('The link must start with https:\/\/ (or leave it blank).');
  return s.slice(0, 500);
}

// Saves an uploaded photo or PDF (a data: URL from the page) to the private folder. Returns the file ID.
function saveServiceFile_(dataUrl, label) {
  if (!dataUrl) return '';
  const m = String(dataUrl).match(/^data:(image\/(?:png|jpeg|gif|webp)|application\/pdf);base64,([A-Za-z0-9+/=]+)$/);
  if (!m) throw new Error('Upload a photo (PNG or JPG) or a PDF.');
  const bytes = Utilities.base64Decode(m[2]);
  if (bytes.length > MAX_SERVICE_BYTES_) throw new Error('That file is too big (6 MB most). Try a smaller photo, or paste a link instead.');
  const ext = { 'image/png': '.png', 'image/jpeg': '.jpg', 'image/gif': '.gif', 'image/webp': '.webp', 'application/pdf': '.pdf' }[m[1]];
  return serviceFolder_().createFile(Utilities.newBlob(bytes, m[1], label + ext)).getId();
}

/** A student's own submissions (no other student's). */
function clashServiceMine(id) {
  const sid = id800_(id);
  if (!sid) return '[]';
  return JSON.stringify(sheetRows_('Service Submissions').filter(function (r) { return id800_(r['800#']) === sid; })
    .sort(function (a, b) { return String(b['Submitted']).localeCompare(String(a['Submitted'])); }).slice(0, 10)
    .map(function (r) {
      return { id: String(r['ID']), submitted: String(r['Submitted']).slice(0, 10), what: String(r['What'] || ''), status: String(r['Status'] || 'Waiting'),
        points: Number(r['Points']) || 0, note: String(r['Status']) === 'Approved' ? '' : String(r['Staff Note'] || ''), reply: String(r['Student Reply'] || '') };
    }));
}

/** { id, what, when, reflection, link, file (data: URL) } from My Squad. */
function clashServiceSubmit(id, json) {
  const st = serviceStudent_(id);
  if (!st) throw new Error('We couldn\'t find that 800 number. Check it and try again.');
  if (SQUAD_KEYS_.indexOf(st.squad) === -1) throw new Error('You don\'t have a squad yet. Ask your homeroom teacher.');
  const p = JSON.parse(json || '{}');
  const what = String(p.what || '').trim().slice(0, 400), when = String(p.when || '').trim().slice(0, 200), refl = String(p.reflection || '').trim().slice(0, 1500);
  if (what.length < 10) throw new Error('Tell us what you did (a sentence or two).');
  if (refl.length < 40) throw new Error('Add a reflection: why it mattered or what you learned (a few sentences).');
  const link = cleanLink_(p.link);
  ensureServiceTab_();
  const open = sheetRows_('Service Submissions').filter(function (r) { return id800_(r['800#']) === st.id && /^(waiting|more info)/i.test(String(r['Status'] || 'Waiting')); }).length;
  if (open >= Math.max(1, settingNum_('SERVICE_OPEN', 3))) throw new Error('You have ' + open + ' waiting for staff already. Once they\'re looked at you can send more.');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sub = 'S' + Date.now().toString(36);
    const file = p.file ? saveServiceFile_(p.file, sub + ' ' + st.squad + ' ' + st.first) : '';
    getSheet_('Service Submissions').appendRow([sub, Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'), st.id,
      st.first + ' ' + st.last, st.squad, what, when, refl, link, file, 'Waiting', '', '', '', '', '']);
  } finally { lock.releaseLock(); }
  return clashServiceMine(st.id);
}

/** The student answers a "More info" question (and can add a link or a file). */
function clashServiceReply(id, subId, json) {
  const sid = id800_(id), p = JSON.parse(json || '{}');
  const r = sheetRows_('Service Submissions').filter(function (x) { return String(x['ID']) === String(subId) && id800_(x['800#']) === sid; })[0];
  if (!r) throw new Error('We couldn\'t find that submission.');
  if (String(r['Status']) !== 'More info') throw new Error('This one isn\'t waiting for an answer.');
  const reply = String(p.reply || '').trim().slice(0, 1000);
  if (reply.length < 5) throw new Error('Write your answer.');
  const link = cleanLink_(p.link) || String(r['Link'] || '');
  const file = p.file ? saveServiceFile_(p.file, r['ID'] + ' reply') : String(r['File'] || '');
  const sh = getSheet_('Service Submissions');
  const set = function (h, v) { const c = headerCol_(sh, h); if (c) sh.getRange(r._row, c).setValue(v); };
  set('Student Reply', (String(r['Student Reply'] || '') ? String(r['Student Reply']) + '\n' : '') + reply);
  set('Link', link); set('File', file); set('Status', 'Waiting');
  return clashServiceMine(sid);
}

// ── Staff (Admin › ✅ Service) ──
function clashServiceList(token) {
  requireAdmin_(token);
  ensureServiceTab_();
  const rows = sheetRows_('Service Submissions').filter(function (r) { return String(r['ID'] || '').trim(); });
  const view = function (r) {
    return { id: String(r['ID']), submitted: String(r['Submitted']), student: String(r['Student'] || ''), squad: squadOf_(r['Squad']), what: String(r['What'] || ''),
      when: String(r['When / Where'] || ''), reflection: String(r['Reflection'] || ''), link: String(r['Link'] || ''), file: String(r['File'] || ''),
      status: String(r['Status'] || 'Waiting'), points: Number(r['Points']) || 0, note: String(r['Staff Note'] || ''), reply: String(r['Student Reply'] || ''),
      by: String(r['Decided By'] || ''), at: String(r['Decided At'] || '') };
  };
  return JSON.stringify({
    points: servicePoints_(),
    open: rows.filter(function (r) { return /^(waiting|more info|)$/i.test(String(r['Status'] || '')); }).map(view)
      .sort(function (a, b) { return (a.status === 'Waiting' ? 0 : 1) - (b.status === 'Waiting' ? 0 : 1) || a.submitted.localeCompare(b.submitted); }),
    done: rows.filter(function (r) { return /^(approved|denied)$/i.test(String(r['Status'] || '')); }).map(view)
      .sort(function (a, b) { return b.at.localeCompare(a.at); }).slice(0, 20),
  });
}

/** The evidence file as a data: URL for the panel (so staff never need Drive access to it). */
function clashServiceFile(token, subId) {
  requireAdmin_(token);
  const r = sheetRows_('Service Submissions').filter(function (x) { return String(x['ID']) === String(subId); })[0];
  if (!r || !String(r['File'] || '')) throw new Error('No file on this one.');
  const blob = DriveApp.getFileById(driveIdFrom_(r['File'])).getBlob();
  return JSON.stringify({ name: blob.getName(), type: blob.getContentType(), data: 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes()) });
}

/** decision: approve (points), info (note = the question), deny (note required). */
function clashServiceDecide(token, subId, decision, points, note) {
  requireAdmin_(token);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const r = sheetRows_('Service Submissions').filter(function (x) { return String(x['ID']) === String(subId); })[0];
    if (!r) throw new Error('That submission is gone.');
    if (/^(approved|denied)$/i.test(String(r['Status']))) throw new Error('Already ' + String(r['Status']).toLowerCase() + '.');
    note = String(note || '').trim().slice(0, 600);
    const sh = getSheet_('Service Submissions');
    const set = function (h, v) { const c = headerCol_(sh, h); if (c) sh.getRange(r._row, c).setValue(v); };
    const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
    if (decision === 'info') {
      if (note.length < 5) throw new Error('Write the question for the student.');
      set('Status', 'More info'); set('Staff Note', note); set('Decided At', stamp);
    } else if (decision === 'deny') {
      if (note.length < 5) throw new Error('Write a short reason (the student sees it).');
      set('Status', 'Denied'); set('Staff Note', note); set('Decided By', 'Admin'); set('Decided At', stamp);
    } else if (decision === 'approve') {
      points = Number(points);
      if (servicePoints_().indexOf(points) === -1) throw new Error('Pick ' + servicePoints_().join(', ') + ' points.');
      const squad = squadOf_(r['Squad']);
      if (SQUAD_KEYS_.indexOf(squad) === -1) throw new Error('That student has no squad.');
      const parts = String(r['Student'] || '').trim().split(/\s+/), short = parts[0] + (parts.length > 1 ? ' ' + parts[parts.length - 1].charAt(0) + '.' : '');
      const pts = { BLACK: '', GOLD: '', GREY: '', PURPLE: '' }; pts[squad] = points;
      getSheet_('events').appendRow(['', 'Service: ' + short, 'service', todayYmd_(), '', '', pts.BLACK, pts.GOLD, pts.GREY, pts.PURPLE]);
      set('Status', 'Approved'); set('Points', points); if (note) set('Staff Note', note); set('Decided By', 'Admin'); set('Decided At', stamp);
    } else throw new Error('Unknown decision.');
  } finally { lock.releaseLock(); }
  return clashServiceList(token);
}

// file-stamp (tools/stamp-files.js): lets Admin › System tell whether this file is the newest one
var FILE_STAMP_40_Service = '24a25fc269';
