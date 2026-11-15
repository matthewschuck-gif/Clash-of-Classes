/**
 * Clash of Classes -- Apps Script backend
 * File 39: Virtual Trait Dough. Teachers hand out a digital Trait Dough bill from any Chromebook:
 * clashofclasses.org/dough (?page=dough).
 *   - Each teacher signs in with their own password (their room number, set by staff). A teacher can
 *     only give dough to students in the OTHER squads, never their own (Mrs. Example is Gold, so she
 *     gives to Grey, Black and Purple students).
 *   - A bill: the student, the trait they showed (VIRTUAL_DOUGH_TRAITS, blank = the LRG traits), what
 *     they did (required), and 1, 3 or 5 points (VIRTUAL_DOUGH_POINTS). It goes on the events tab as
 *     Trait Dough (type trait_dough), so it counts for the Trait Dough Drive area like paper bills.
 *   - VIRTUAL_DOUGH_DAILY (10) bills per teacher per day. 5 wrong passwords = that name is locked
 *     for 15 minutes.
 *   - VIRTUAL_DOUGH_PER_STUDENT (2) bills per student per day, from all teachers together.
 *   - The points balance check (29_Balance.gs) flags lopsided giving: one teacher's bills mostly to one
 *     student or one squad, or a student whose dough all comes from one teacher.
 *   - Fridays (DOUGH_NUDGE ON) each teacher with an Email gets a short note: their bills this week and
 *     the squads they haven't given to this month.
 * DOUGH TEACHERS tab (staff only): Active | Teacher | Squad | Password | Email.
 * VIRTUAL DOUGH tab: every bill (When | Teacher | Teacher Squad | 800# | Student | Squad | Trait |
 * What They Did | Points). The public site only ever shows "Virtual Trait Dough: Ann L.".
 */

const DOUGH_TEACHER_HEADERS_ = ['Active', 'Teacher', 'Squad', 'Password', 'Email'];
const DOUGH_LOG_HEADERS_ = ['When', 'Teacher', 'Teacher Squad', '800#', 'Student', 'Squad', 'Trait', 'What They Did', 'Points'];

function ensureDoughTabs_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  [['Dough Teachers', DOUGH_TEACHER_HEADERS_, [[false, 'Example Teacher (untick or replace)', 'GOLD', '74', '']]], ['Virtual Dough', DOUGH_LOG_HEADERS_, null]].forEach(function (t) {
    const old = ss.getSheetByName(t[0]);
    if (old) { // an older tab gets the newer columns (Email)
      const head = old.getRange(1, 1, 1, Math.max(1, old.getLastColumn())).getValues()[0].map(function (h) { return String(h).trim(); });
      t[1].forEach(function (h) { if (head.indexOf(h) === -1) { head.push(h); old.getRange(1, head.length).setValue(h).setFontWeight('bold').setBackground('#490e6f').setFontColor('#ffffff'); } });
      return;
    }
    const sh = ss.insertSheet(t[0]);
    sh.getRange(1, 1, 1, t[1].length).setValues([t[1]]).setFontWeight('bold').setBackground('#490e6f').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    if (t[2]) sh.getRange(2, 1, t[2].length, t[1].length).setValues(t[2]);
    if (t[0] === 'Dough Teachers') { try { sh.getRange(2, 1, 200, 1).insertCheckboxes(); } catch (e) { /* fine */ } }
  });
}

function doughTeachers_() {
  return sheetRows_('Dough Teachers').filter(function (r) { return ticked_(r['Active']) && String(r['Teacher'] || '').trim(); }).map(function (r) {
    return { name: String(r['Teacher']).trim(), squad: squadOf_(r['Squad']), pw: String(r['Password'] == null ? '' : r['Password']).trim(), email: String(r['Email'] || '').trim() };
  });
}

function doughPoints_() { const l = listOf_(getSetting_('VIRTUAL_DOUGH_POINTS') || '1, 3, 5').map(Number).filter(function (n) { return n > 0; }); return l.length ? l : [1, 3, 5]; }

function doughTraits_() {
  const own = String(getSetting_('VIRTUAL_DOUGH_TRAITS') || '').split(',').map(function (s) { return s.trim(); }).filter(String);
  if (own.length) return own;
  try { const t = siteTraits_(); if (t && t.length) return t.map(function (x) { return x.trait; }); } catch (e) { /* built-in list */ }
  return ['Growth Mindset', 'Problem Solving', 'Communication', 'Collaboration', 'Responsibility'];
}

/** The sign-in screen: teacher names only (never squads or passwords). */
function clashDoughStart() {
  return JSON.stringify({ teachers: doughTeachers_().map(function (t) { return t.name; }).sort(), points: doughPoints_(), traits: doughTraits_(),
    shoutouts: (function () { try { return settingOn_('SHOUTOUTS', 'ON'); } catch (e) { return false; } })() }); // 56_Shoutouts.gs
}

function doughSession_(token) {
  const raw = token ? CacheService.getScriptCache().get('dough-' + token) : null;
  if (!raw) throw new Error('Please sign in again.');
  return JSON.parse(raw);
}

function doughGivenToday_(teacher) {
  const today = todayYmd_();
  return sheetRows_('Virtual Dough').filter(function (r) { return String(r['Teacher']) === teacher && String(r['When']).slice(0, 10) === today; }).length;
}

/** name + password -> a 2-hour session. */
function clashDoughSignIn(name, password) {
  const cache = CacheService.getScriptCache(), key = 'dough-fail-' + String(name || '').toLowerCase().replace(/\W/g, '');
  const fails = Number(cache.get(key)) || 0;
  if (fails >= 5) throw new Error('Too many wrong passwords. Try again in 15 minutes.');
  const t = doughTeachers_().filter(function (x) { return x.name === String(name || '').trim(); })[0];
  if (!t || !t.pw || String(password || '').trim().toLowerCase() !== t.pw.toLowerCase()) {
    cache.put(key, String(fails + 1), 900);
    throw new Error('That password doesn\'t match. (It\'s set by the Clash coordinator.)');
  }
  if (SQUAD_KEYS_.indexOf(t.squad) === -1) throw new Error('Your squad isn\'t set yet. Ask the Clash coordinator.');
  cache.remove(key);
  const token = Utilities.getUuid().replace(/-/g, '');
  cache.put('dough-' + token, JSON.stringify({ name: t.name, squad: t.squad }), 7200);
  const max = Math.max(1, settingNum_('VIRTUAL_DOUGH_DAILY', 10));
  return JSON.stringify({ token: token, name: t.name, squad: t.squad, left: Math.max(0, max - doughGivenToday_(t.name)), max: max });
}

/** Students from the other squads matching a name or 800#. */
function clashDoughFind(token, query) {
  const me = doughSession_(token), q = String(query || '').trim().toLowerCase();
  if (q.length < 2) return '[]';
  return JSON.stringify(sheetRows_('Students').filter(function (r) {
    const s = squadOf_(r['Squad']);
    if (SQUAD_KEYS_.indexOf(s) === -1 || s === me.squad) return false;
    const name = (String(r['First Name'] || '') + ' ' + String(r['Last Name'] || '')).toLowerCase();
    return name.indexOf(q) > -1 || id800_(r['800#']) === id800_(q);
  }).slice(0, 15).map(function (r) {
    return { id: id800_(r['800#']), name: String(r['First Name'] || '').trim() + ' ' + String(r['Last Name'] || '').trim(), grade: String(r['Grade'] || ''), squad: squadOf_(r['Squad']) };
  }));
}

/** Gives one bill. */
function clashDoughGive(token, id, trait, what, points) {
  const me = doughSession_(token);
  const sid = id800_(id), stu = sheetRows_('Students').filter(function (r) { return id800_(r['800#']) === sid; })[0];
  if (!sid || !stu) throw new Error('Pick a student from the list.');
  const squad = squadOf_(stu['Squad']);
  if (SQUAD_KEYS_.indexOf(squad) === -1) throw new Error('That student has no squad yet.');
  if (squad === me.squad) throw new Error('Virtual Trait Dough goes to students in the other squads, not your own.');
  points = Number(points);
  if (doughPoints_().indexOf(points) === -1) throw new Error('Pick ' + doughPoints_().join(', ') + ' points.');
  trait = String(trait || '').trim();
  if (doughTraits_().indexOf(trait) === -1) throw new Error('Pick the trait they showed.');
  what = String(what || '').trim().slice(0, 300);
  if (what.length < 8) throw new Error('Write what they did (a sentence is plenty).');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const max = Math.max(1, settingNum_('VIRTUAL_DOUGH_DAILY', 10)), given = doughGivenToday_(me.name);
    if (given >= max) throw new Error('That\'s all ' + max + ' of today\'s bills. More tomorrow!');
    const perStudent = Math.max(1, settingNum_('VIRTUAL_DOUGH_PER_STUDENT', 2)), today = todayYmd_();
    const theirs = sheetRows_('Virtual Dough').filter(function (r) { return id800_(r['800#']) === sid && String(r['When']).slice(0, 10) === today; }).length;
    if (theirs >= perStudent) throw new Error('That student already got ' + theirs + ' bills today (the most per day). Catch them again tomorrow!');
    const first = String(stu['First Name'] || '').trim(), last = String(stu['Last Name'] || '').trim();
    const full = first + ' ' + last, short = first + (last ? ' ' + last.charAt(0) + '.' : '');
    ensureDoughTabs_();
    getSheet_('Virtual Dough').appendRow([Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'), me.name, me.squad, sid, full, squad, trait, what, points]);
    const pts = { BLACK: '', GOLD: '', GREY: '', PURPLE: '' }; pts[squad] = points;
    getSheet_('events').appendRow(['', 'Virtual Trait Dough: ' + short, 'trait_dough', todayYmd_(), trait + ' · from ' + me.name, '', pts.BLACK, pts.GOLD, pts.GREY, pts.PURPLE]);
    return JSON.stringify({ ok: true, student: short, squad: squad, points: points, left: Math.max(0, max - given - 1) });
  } finally { lock.releaseLock(); }
}

function serveDough_() {
  return HtmlService.createHtmlOutputFromFile('Dough').setTitle('Virtual Trait Dough · Clash of Classes')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ── Fairness flags (for the points balance check) ──
/** Lines about lopsided Virtual Dough between from and to. Students named first name + last initial. */
function doughFlags_(from, to) {
  const rows = sheetRows_('Virtual Dough').filter(function (r) { const d = String(r['When']).slice(0, 10); return d >= from && d <= to; });
  const flags = [], byTeacher = {}, byStudent = {};
  rows.forEach(function (r) {
    const t = String(r['Teacher']), id = id800_(r['800#']);
    (byTeacher[t] = byTeacher[t] || []).push(r);
    (byStudent[id] = byStudent[id] || []).push(r);
  });
  Object.keys(byTeacher).forEach(function (t) {
    const list = byTeacher[t];
    if (list.length < 8) return;
    const per = {}, sq = {};
    list.forEach(function (r) { const id = id800_(r['800#']); per[id] = (per[id] || 0) + 1; sq[squadOf_(r['Squad'])] = (sq[squadOf_(r['Squad'])] || 0) + 1; });
    const topId = Object.keys(per).sort(function (a, b) { return per[b] - per[a]; })[0];
    if (per[topId] / list.length > 0.4) flags.push('Virtual Dough: ' + t + ' gave ' + per[topId] + ' of ' + list.length + ' bills to one student (' + shortName_(list.filter(function (r) { return id800_(r['800#']) === topId; })[0]['Student']) + ').');
    const topSq = Object.keys(sq).sort(function (a, b) { return sq[b] - sq[a]; })[0];
    if (sq[topSq] / list.length > 0.7) flags.push('Virtual Dough: ' + t + ' gave ' + sq[topSq] + ' of ' + list.length + ' bills to ' + titleSquad_(topSq) + '.');
  });
  Object.keys(byStudent).forEach(function (id) {
    const list = byStudent[id], teachers = {};
    list.forEach(function (r) { teachers[String(r['Teacher'])] = 1; });
    if (list.length >= 5 && Object.keys(teachers).length === 1) flags.push('Virtual Dough: ' + shortName_(list[0]['Student']) + ' got all ' + list.length + ' bills from ' + Object.keys(teachers)[0] + '.');
  });
  return flags;
}

// ── The Friday nudge ──
/** Emails each active teacher with an Email: bills this week, squads not given to this month. Returns lines. */
function doughNudge_() {
  const today = todayYmd_(), mon = ymdShift_(today, -((new Date(today + 'T12:00:00Z').getUTCDay() + 6) % 7)), month = today.slice(0, 7);
  const rows = sheetRows_('Virtual Dough'), lines = [], site = siteUrl_();
  doughTeachers_().filter(function (t) { return /@/.test(t.email); }).forEach(function (t) {
    const mine = rows.filter(function (r) { return String(r['Teacher']) === t.name; });
    const week = mine.filter(function (r) { return String(r['When']).slice(0, 10) >= mon; }).length;
    const monthRows = mine.filter(function (r) { return String(r['When']).slice(0, 7) === month; });
    const others = SQUAD_KEYS_.filter(function (s) { return s !== t.squad; });
    const none = others.filter(function (s) { return !monthRows.some(function (r) { return squadOf_(r['Squad']) === s; }); });
    const body = '<div style="font:15px/1.6 Arial,sans-serif;color:#22123b">' +
      '<p>Hi ' + t.name + ',</p><p>' + (week ? 'You gave <b>' + week + '</b> Virtual Trait Dough bill' + (week === 1 ? '' : 's') + ' this week. Thank you!' : 'No Virtual Trait Dough from you yet this week.') + '</p>' +
      (none.length ? '<p>This month, <b>' + none.map(titleSquad_).join(' and ') + '</b> ' + (none.length === 1 ? 'hasn\'t' : 'haven\'t') + ' gotten any from you. Keep an eye out for a student showing a trait.</p>' : '<p>You\'ve given to every other squad this month. 🙌</p>') +
      '<p><a href="' + site + '/dough" style="background:#1f7a4d;color:#fff;padding:9px 16px;border-radius:999px;text-decoration:none;font-weight:700">Give Trait Dough</a></p>' +
      '<p style="color:#888;font-size:12px">Sent on Fridays by Clash of Classes. Ask the Clash coordinator to take you off (Dough Teachers tab, Email).</p></div>';
    try { MailApp.sendEmail({ to: t.email, subject: 'Virtual Trait Dough this week', htmlBody: body, name: 'Clash of Classes' }); lines.push(t.name + ': sent'); }
    catch (e) { lines.push(t.name + ': not sent (' + e.message + ')'); }
  });
  return lines;
}

function maybeDoughNudge_() {
  if (String(getSetting_('DOUGH_NUDGE') || 'ON').toUpperCase() !== 'ON') return;
  const now = new Date(), today = todayYmd_();
  if (new Date(today + 'T12:00:00Z').getUTCDay() !== 5 || Number(Utilities.formatDate(now, Session.getScriptTimeZone(), 'H')) < 13) return;
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('dough-nudge') === today) return;
  props.setProperty('dough-nudge', today);
  doughNudge_();
}

// file-stamp (tools/stamp-files.js): lets Admin › System tell whether this file is the newest one
var FILE_STAMP_39_VirtualDough = '6b48d03bc1';
