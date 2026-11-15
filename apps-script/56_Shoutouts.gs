/**
 * Clash of Classes -- Apps Script backend
 * File 56: staff shout-outs and the event planner.
 *
 * STAFF SHOUT-OUTS (Settings › SHOUTOUTS ON)
 *   Teachers signed in on the Virtual Trait Dough page write a one-line shout-out to a student.
 *   Nothing shows until staff tick Approved (Sheet or Admin › System), unless
 *   SHOUTOUT_AUTO_APPROVE is ON. Approved ones show on the home page and the morning
 *   announcement for SHOUTOUT_DAYS (14) days, first name + last initial only, and the student
 *   sees their own on My Climb. No points: it's a thank-you, not a score.
 *   Staff Shout-outs tab: When | From | 800# | Student | Squad | Message | Approved.
 *
 * EVENT PLANNER
 *   Planned Events tab: Show | Date | Event | What Happens | Type | BLACK | GOLD | GREY | PURPLE | Posted
 *   Plan a Clash event ahead of time. Ticked rows show on the home page the week before ("Coming
 *   up") and in the morning announcement the day before and the day of. After the day, type each
 *   squad's points (Sheet, or Admin › System › Planned events) and it posts to the events tab once,
 *   as a normal event (Type, blank = event). Posted fills in.
 */

const SHOUT_HEADERS_ = ['When', 'From', '800#', 'Student', 'Squad', 'Message', 'Approved'];
const PLANNED_HEADERS_ = ['Show', 'Date', 'Event', 'What Happens', 'Type', 'BLACK', 'GOLD', 'GREY', 'PURPLE', 'Posted'];

function ensureShoutoutTabs_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  [['Staff Shout-outs', SHOUT_HEADERS_], ['Planned Events', PLANNED_HEADERS_]].forEach(function (t) {
    if (ss.getSheetByName(t[0])) return;
    const sh = ss.insertSheet(t[0]);
    sh.getRange(1, 1, 1, t[1].length).setValues([t[1]]).setFontWeight('bold').setBackground('#490e6f').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  });
}

// ── Shout-outs ──
// Student search for the shout-out box: every squad (Trait Dough only goes to the other squads).
function clashShoutoutFind(token, query) {
  doughSession_(token);
  const q = String(query || '').trim().toLowerCase();
  if (q.length < 2) return '[]';
  return JSON.stringify(sheetRows_('Students').filter(function (r) {
    const name = (String(r['First Name'] || '') + ' ' + String(r['Last Name'] || '')).toLowerCase();
    return SQUAD_KEYS_.indexOf(squadOf_(r['Squad'])) > -1 && (name.indexOf(q) > -1 || id800_(r['800#']) === id800_(q));
  }).slice(0, 15).map(function (r) {
    return { id: id800_(r['800#']), name: String(r['First Name'] || '').trim() + ' ' + String(r['Last Name'] || '').trim(), grade: String(r['Grade'] || ''), squad: squadOf_(r['Squad']) };
  }));
}

function clashShoutoutPost(token, id, message) {
  if (!settingOn_('SHOUTOUTS', 'ON')) throw new Error('Shout-outs are off right now.');
  const me = doughSession_(token), sid = id800_(id);
  message = String(message || '').replace(/\s+/g, ' ').trim();
  if (message.length < 10) throw new Error('Write a sentence about what they did.');
  if (message.length > 160) throw new Error('Keep it to 160 characters.');
  const s = sheetRows_('Students').filter(function (r) { return id800_(r['800#']) === sid; })[0] ||
    boardRows_().filter(function (r) { return id800_(r['800#']) === sid; })[0];
  if (!s) throw new Error('Pick a student from the list.');
  ensureShoutoutTabs_();
  const today = todayYmd_();
  const mineToday = sheetRows_('Staff Shout-outs').filter(function (r) { return String(r['From']) === me.name && String(r['When']).slice(0, 10) === today; }).length;
  if (mineToday >= settingNum_('SHOUTOUTS_PER_DAY', 10)) throw new Error('That\'s today\'s shout-outs. More tomorrow!');
  const name = s['First Name'] ? String(s['First Name']).trim() + ' ' + String(s['Last Name'] || '').trim() : String(s['Name'] || '');
  getSheet_('Staff Shout-outs').appendRow([Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'), me.name, sid, name.trim(),
    squadOf_(s['Squad']), message, settingOn_('SHOUTOUT_AUTO_APPROVE', 'OFF')]);
  return JSON.stringify({ ok: true, waiting: !settingOn_('SHOUTOUT_AUTO_APPROVE', 'OFF') });
}

/** Approved shout-outs from the last SHOUTOUT_DAYS days, newest first: [{ who, squad, from, message, date }]. */
function shoutoutsPublic_(limit) {
  if (!settingOn_('SHOUTOUTS', 'ON')) return [];
  const since = ymdShift_(todayYmd_(), -settingNum_('SHOUTOUT_DAYS', 14));
  return sheetRows_('Staff Shout-outs').filter(function (r) { return ticked_(r['Approved']) && String(r['When']).slice(0, 10) >= since && String(r['Message'] || '').trim(); })
    .map(function (r) { return { who: shortName_(r['Student']), squad: squadOf_(r['Squad']), from: String(r['From'] || ''), message: String(r['Message']).trim(), date: String(r['When']).slice(0, 10) }; })
    .reverse().slice(0, limit || 6);
}

/** One student's approved shout-outs (My Climb). */
function myShoutouts_(id) {
  return sheetRows_('Staff Shout-outs').filter(function (r) { return id800_(r['800#']) === id && ticked_(r['Approved']); })
    .map(function (r) { return { from: String(r['From'] || ''), message: String(r['Message'] || ''), date: String(r['When']).slice(0, 10) }; }).reverse().slice(0, 5);
}

// ── Event planner ──
function plannedRows_() {
  return sheetRows_('Planned Events').filter(function (r) { return ticked_(r['Show']) && asYmd_(r['Date']) && String(r['Event'] || '').trim(); }).map(function (r) {
    const pts = {}; let any = false;
    SQUAD_KEYS_.forEach(function (s) { const v = String(r[s] == null ? '' : r[s]).trim(); pts[s] = v === '' ? '' : Number(v) || 0; if (v !== '') any = true; });
    return { row: r._row, date: asYmd_(r['Date']).slice(0, 10), name: String(r['Event']).trim(), what: String(r['What Happens'] || '').trim(),
      type: String(r['Type'] || '').trim() || 'event', pts: pts, scored: any, posted: String(r['Posted'] || '').trim() };
  });
}

/** Coming up in the next 7 days (today included): [{ date, name, what }]. */
function plannedSoon_() {
  const today = todayYmd_(), soon = ymdShift_(today, 7);
  return plannedRows_().filter(function (p) { return p.date >= today && p.date <= soon; }).sort(function (a, b) { return a.date.localeCompare(b.date); })
    .map(function (p) { return { date: p.date, name: p.name, what: p.what }; });
}

function postPlanned_(p) {
  getSheet_('events').appendRow(['', p.name, p.type, p.date, p.what, '', p.pts.BLACK, p.pts.GOLD, p.pts.GREY, p.pts.PURPLE]);
  getSheet_('Planned Events').getRange(p.row, PLANNED_HEADERS_.indexOf('Posted') + 1).setValue(todayYmd_());
}

/** Daily: posts every past planned event that has points typed in. */
function postPlannedEvents_() {
  const today = todayYmd_();
  let n = 0;
  plannedRows_().forEach(function (p) { if (p.date <= today && p.scored && !p.posted) { postPlanned_(p); n++; } });
  return n;
}

/** Admin › System: past events still waiting for points, and the next ones. */
function clashPlanned(token) {
  requireAdmin_(token);
  ensureShoutoutTabs_();
  const today = todayYmd_();
  const rows = plannedRows_();
  return JSON.stringify({
    toScore: rows.filter(function (p) { return p.date <= today && !p.posted; }).map(function (p) { return { row: p.row, date: p.date, name: p.name, pts: p.pts }; }),
    soon: plannedSoon_(),
  });
}

function clashPlannedPost(token, row, ptsJson) {
  requireAdmin_(token);
  const p = plannedRows_().filter(function (x) { return x.row === Number(row); })[0];
  if (!p) throw new Error('That planned event isn\'t on the tab (is Show ticked?).');
  if (p.posted) throw new Error('Already posted on ' + p.posted + '.');
  if (p.date > todayYmd_()) throw new Error('Post the points on or after ' + p.date + '.');
  const pts = JSON.parse(ptsJson || '{}');
  SQUAD_KEYS_.forEach(function (s) { const v = Math.round(Number(pts[s]) || 0); if (v < 0 || v > 1000) throw new Error('Points are 0 to 1000.'); p.pts[s] = v || ''; });
  getSheet_('Planned Events').getRange(p.row, PLANNED_HEADERS_.indexOf('BLACK') + 1, 1, 4).setValues([[p.pts.BLACK, p.pts.GOLD, p.pts.GREY, p.pts.PURPLE]]);
  postPlanned_(p);
  return clashPlanned(token);
}

// file-stamp (tools/stamp-files.js): lets Admin › System tell whether this file is the newest one
var FILE_STAMP_56_Shoutouts = '';

// file-stamp (tools/stamp-files.js): lets Admin › System tell whether this file is the newest one
var FILE_STAMP_56_Shoutouts = 'a57689a8e2';
