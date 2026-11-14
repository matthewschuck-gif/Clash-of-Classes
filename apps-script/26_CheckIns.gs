/**
 * Clash of Classes -- Apps Script backend
 * File 26: Squad event check-ins (Move · Make · Talk). How Gym 1 on the map is won.
 *
 * Sign-up Events tab gains a Kind column: Move (active), Make (creative) or Talk (trivia,
 *   speaking). An event with a Kind and a Date opens for check-in on that Date and stays open
 *   CHECKIN_DAYS more days (Settings, default 1), for students who forgot.
 * Event Check-ins tab: Event | Date | Kind | 800# | Name | Squad | Checked In | Confirmed | Confirmed By
 *   Students check in on the site with their 800# (My Squad and the home page). A General
 *   confirms each one in Generals HQ › Check-ins (or removes it), and can add a member who forgot.
 *   Only confirmed check-ins count: an 800# typed by someone else never scores on its own.
 *
 * Gym 1 (rule "Move · Make · Talk", 18_Conquest.gs) = last month's confirmed participants plus the
 * points staff gave for squad events typed Move / Make / Talk (squad_move, squad_make, squad_talk).
 */

const CHECKIN_HEADERS_ = ['Event', 'Date', 'Kind', '800#', 'Name', 'Squad', 'Checked In', 'Confirmed', 'Confirmed By'];
const EVENT_KINDS_ = ['Move', 'Make', 'Talk'];
const SQUAD_EVENT_TYPES_ = 'squad_move, squad_make, squad_talk';

function ensureCheckinTabs_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName('Event Check-ins')) {
    const sh = ss.insertSheet('Event Check-ins');
    sh.getRange(1, 1, 1, CHECKIN_HEADERS_.length).setValues([CHECKIN_HEADERS_]).setFontWeight('bold').setBackground('#490e6f').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  }
  const ev = ss.getSheetByName('Sign-up Events');
  if (ev && ev.getLastColumn()) {
    const head = ev.getRange(1, 1, 1, ev.getLastColumn()).getValues()[0].map(function (h) { return String(h).trim(); });
    if (head.indexOf('Kind') === -1) {
      ev.getRange(1, head.filter(String).length + 1).setValue('Kind').setFontWeight('bold').setBackground('#490e6f').setFontColor('#ffffff');
      try {
        ev.getRange(2, head.filter(String).length + 1, 199, 1)
          .setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(EVENT_KINDS_, true).build());
      } catch (e) { console.warn('Kind dropdown: ' + e.message); }
    }
  }
}

function eventKind_(v) {
  const k = String(v || '').trim().toLowerCase();
  return EVENT_KINDS_.filter(function (x) { return x.toLowerCase() === k; })[0] || '';
}

// Squad events with a Kind: [{ event, date, kind, until, open }]; open = check-in open today.
function checkinEvents_() {
  const today = todayYmd_(), days = Math.max(0, Math.min(7, Number(getSetting_('CHECKIN_DAYS')) || 1));
  return sheetRows_('Sign-up Events').map(function (r) {
    const date = String(r['Date'] || '').slice(0, 10), kind = eventKind_(r['Kind']);
    if (!String(r['Event'] || '').trim() || !kind || !/^\d{4}-\d{2}-\d{2}$/.test(date)) return null;
    const until = ymdShift_(date, days);
    return { event: String(r['Event']).trim(), date: date, kind: kind, until: until, open: today >= date && today <= until };
  }).filter(Boolean);
}

/** Student side: events open for check-in today, and what this 800# has done. */
function clashCheckins(id) {
  const sid = id800_(id);
  const mine = sid ? sheetRows_('Event Check-ins').filter(function (r) { return id800_(r['800#']) === sid; }) : [];
  return JSON.stringify({ events: checkinEvents_().filter(function (e) { return e.open; }).map(function (e) {
    const c = mine.filter(function (r) { return String(r['Event']) === e.event && String(r['Date'] || '').slice(0, 10) === e.date; })[0];
    return { event: e.event, date: e.date, kind: e.kind, status: !c ? '' : ticked_(c['Confirmed']) ? 'confirmed' : 'waiting' };
  }) });
}

/** A student checks in with their 800#. A General confirms it in HQ before it counts. */
function clashEventCheckIn(event, id) {
  const cache = CacheService.getScriptCache();
  const used = Number(cache.get('checkins')) || 0;
  if (used >= (Number(getSetting_('POLL_VOTES_PER_10MIN')) || 1500)) throw new Error('So many check-ins right now! Try again in a minute.');
  cache.put('checkins', String(used + 1), 600);
  const sid = id800_(id);
  if (sid.length < 5) throw new Error('Type your 800 number.');
  const ev = checkinEvents_().filter(function (e) { return e.event === String(event).trim() && e.open; })[0];
  if (!ev) throw new Error('Check-in for that event isn\'t open.');
  const me = sheetRows_(getSetting_('STUDENT_BOARD_TAB') || 'Student Board').filter(function (r) { return id800_(r['800#']) === sid; })[0];
  if (!me) throw new Error('That 800 number isn\'t on the list. Check it, or ask your homeroom teacher.');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    ensureCheckinTabs_();
    const dup = sheetRows_('Event Check-ins').some(function (r) { return id800_(r['800#']) === sid && String(r['Event']) === ev.event && String(r['Date'] || '').slice(0, 10) === ev.date; });
    if (!dup) {
      getSheet_('Event Check-ins').appendRow([ev.event, ev.date, ev.kind, sid, String(me['Name'] || ''), squadOf_(me['Squad']),
        Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'), false, '']);
    }
  } finally {
    lock.releaseLock();
  }
  return clashCheckins(sid);
}

// For Generals HQ: this squad's check-ins from the last 45 days, newest first,
// plus the events a General can still add someone to (any Kind event in that window).
function hqCheckins_(squad) {
  const from = ymdShift_(todayYmd_(), -45);
  const rows = sheetRows_('Event Check-ins').filter(function (r) { return squadOf_(r['Squad']) === squad && String(r['Date'] || '').slice(0, 10) >= from; })
    .map(function (r) { return { row: r._row, event: String(r['Event']), date: String(r['Date']).slice(0, 10), kind: String(r['Kind'] || ''), id: id800_(r['800#']),
      name: String(r['Name'] || ''), confirmed: ticked_(r['Confirmed']), by: String(r['Confirmed By'] || '') }; })
    .sort(function (a, b) { return b.date.localeCompare(a.date) || Number(a.confirmed) - Number(b.confirmed) || a.name.localeCompare(b.name); });
  const events = checkinEvents_().filter(function (e) { return e.date >= from && e.date <= todayYmd_(); })
    .sort(function (a, b) { return b.date.localeCompare(a.date); });
  return { rows: rows, events: events, waiting: rows.filter(function (r) { return !r.confirmed; }).length };
}

/** A General confirms (ok = true) or removes (ok = false) one of their squad's check-ins. */
function clashGeneralsConfirmCheckin(token, row, ok) {
  const me = hqSession_(token);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const r = sheetRows_('Event Check-ins').filter(function (x) { return x._row === Number(row); })[0];
    if (!r || squadOf_(r['Squad']) !== me.squad) throw new Error('That check-in isn\'t on your squad.');
    const sh = getSheet_('Event Check-ins');
    if (ok) sh.getRange(r._row, CHECKIN_HEADERS_.indexOf('Confirmed') + 1, 1, 2).setValues([[true, me.name || me.id]]);
    else sh.deleteRow(r._row);
  } finally {
    lock.releaseLock();
  }
  return clashGeneralsHq(token);
}

/** A General adds a squad member who took part but didn't check in (confirmed straight away). */
function clashGeneralsAddCheckin(token, event, id) {
  const me = hqSession_(token);
  const sid = id800_(id);
  const ev = hqCheckins_(me.squad).events.filter(function (e) { return e.event === String(event).trim(); })[0];
  if (!ev) throw new Error('Pick a Move, Make or Talk event from the last 45 days.');
  const student = sheetRows_(getSetting_('STUDENT_BOARD_TAB') || 'Student Board').filter(function (r) { return id800_(r['800#']) === sid && squadOf_(r['Squad']) === me.squad; })[0];
  if (!student) throw new Error('That student isn\'t on your squad.');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    ensureCheckinTabs_();
    const sh = getSheet_('Event Check-ins');
    const have = sheetRows_('Event Check-ins').filter(function (r) { return id800_(r['800#']) === sid && String(r['Event']) === ev.event && String(r['Date'] || '').slice(0, 10) === ev.date; })[0];
    if (have) sh.getRange(have._row, CHECKIN_HEADERS_.indexOf('Confirmed') + 1, 1, 2).setValues([[true, me.name || me.id]]);
    else sh.appendRow([ev.event, ev.date, ev.kind, sid, String(student['Name'] || ''), me.squad,
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'), true, me.name || me.id]);
  } finally {
    lock.releaseLock();
  }
  return clashGeneralsHq(token);
}

// Confirmed participants per squad for events dated in [from, to] (one per student per event).
function eventParticipants_(from, to) {
  const out = { BLACK: 0, GOLD: 0, GREY: 0, PURPLE: 0 }, seen = {};
  sheetRows_('Event Check-ins').forEach(function (r) {
    const d = String(r['Date'] || '').slice(0, 10), s = squadOf_(r['Squad']), k = id800_(r['800#']) + '|' + r['Event'] + '|' + d;
    if (d < from || d > to || out[s] === undefined || !ticked_(r['Confirmed']) || seen[k]) return;
    seen[k] = 1; out[s]++;
  });
  return out;
}
