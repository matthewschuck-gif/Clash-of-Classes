/**
 * Clash of Classes -- Apps Script backend
 * File 11: Mountaineer Mayhem Week -- event sign-ups with a limit, students signing themselves
 * up, captain votes, shirt orders, and the Squad Night door.
 *
 * Sign-up Events tab (Generals HQ already uses it) gains: Slot | Type | Self Sign-up
 *   Slot: PASS, Lunch, Core Extension, Mayhem Day, Squad Night (anything you like).
 *   Type: Gaming, Brain, Creative, Performance, Active, Table, Service, Low-key...
 *   Self Sign-up ticked = students can sign up for it themselves on My Squad (when
 *   SELF_SIGNUPS is ON). Everything else is signed up by Generals.
 * Rules for every sign-up (Generals or students):
 *   - at most SIGNUP_LIMIT events per student, counting only events that are Open
 *   - not two events on the same Date in the same Slot
 *   - the squad's Spots Per Squad
 *
 * Shirt Orders tab: Date | 800# | Name | Squad | Size | Paid | Picked Up | Points Sent
 *   Clash of Classes › Send shirt sale points adds SHIRT_POINTS per paid shirt to that squad,
 *   once (Points Sent is filled in). A paid shirt also means free Squad Night entry.
 *
 * Squad Night tab: Time | 800# | Name | Squad | Entry | Paid | By  (written by the door page)
 *   Free for anyone signed up for an event or on the MLC Roster, free with a paid shirt,
 *   otherwise SQUAD_NIGHT_PRICE. Clash of Classes › Send Squad Night attendance points adds
 *   SQUAD_NIGHT_POINTS per student who came, once.
 */

const SIGNUP_EXTRA_HEADERS_ = ['Slot', 'Type', 'Self Sign-up'];
const SHIRT_HEADERS_ = ['Date', '800#', 'Name', 'Squad', 'Size', 'Paid', 'Picked Up', 'Points Sent'];
const DOOR_HEADERS_ = ['Time', '800#', 'Name', 'Squad', 'Entry', 'Paid', 'By'];

function ensureMayhemWeekTabs_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ev = ss.getSheetByName('Sign-up Events');
  if (ev) {
    const have = ev.getRange(1, 1, 1, Math.max(ev.getLastColumn(), 1)).getValues()[0].map(function (h) { return String(h).trim(); });
    const missing = SIGNUP_EXTRA_HEADERS_.filter(function (h) { return have.indexOf(h) === -1; });
    if (missing.length) ev.getRange(1, have.filter(String).length + 1, 1, missing.length).setValues([missing]).setFontWeight('bold').setBackground('#490e6f').setFontColor('#ffffff');
  }
  [['Shirt Orders', SHIRT_HEADERS_], ['Squad Night', DOOR_HEADERS_]].forEach(function (t) {
    if (ss.getSheetByName(t[0])) return;
    const sh = ss.insertSheet(t[0]);
    sh.getRange(1, 1, 1, t[1].length).setValues([t[1]]).setFontWeight('bold').setBackground('#490e6f').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  });
}

function boardRow_(sid) {
  return sheetRows_(getSetting_('STUDENT_BOARD_TAB') || 'Student Board').filter(function (r) { return id800_(r['800#']) === sid; })[0] || null;
}
function eventOpen_(ev) { return String(ev['Open']).trim() === '' || ticked_(ev['Open']); }

/** Throws a plain-words reason if this student can't take this event. */
function signupGuard_(sid, ev, squad) {
  const events = sheetRows_('Sign-up Events');
  const mine = sheetRows_('Sign-ups').filter(function (r) { return id800_(r['800#']) === sid; });
  const limit = Number(getSetting_('SIGNUP_LIMIT')) || 0;
  const openMine = mine.filter(function (r) {
    const e = events.filter(function (x) { return String(x['Event']) === String(r['Event']); })[0];
    return e && eventOpen_(e);
  });
  if (limit && openMine.length >= limit) throw new Error('Already signed up for ' + limit + ' events (the most anyone can do). Drop one first.');
  const date = String(ev['Date'] || ''), slot = String(ev['Slot'] || '').trim().toLowerCase();
  if (date && slot) {
    const clash = mine.filter(function (r) {
      const e = events.filter(function (x) { return String(x['Event']) === String(r['Event']); })[0];
      return e && String(e['Date'] || '') === date && String(e['Slot'] || '').trim().toLowerCase() === slot;
    })[0];
    if (clash) throw new Error('Already in "' + clash['Event'] + '" at the same time (' + ev['Slot'] + ', ' + date + ').');
  }
  const spots = Number(ev['Spots Per Squad']) || 0;
  const taken = sheetRows_('Sign-ups').filter(function (r) { return String(r['Event']) === String(ev['Event']) && squadOf_(r['Squad']) === squad; }).length;
  if (spots && taken >= spots) throw new Error('Your squad\'s spots for that event are full.');
}

// What a student sees on My Squad: open events for their squad, and what they're in.
function mySignups_(sid, squad) {
  const on = String(getSetting_('SELF_SIGNUPS')).toUpperCase() === 'ON';
  const signups = sheetRows_('Sign-ups');
  const mine = signups.filter(function (r) { return id800_(r['800#']) === sid; }).map(function (r) { return String(r['Event']); });
  const events = sheetRows_('Sign-up Events').filter(function (e) { return String(e['Event'] || '').trim() && eventOpen_(e); }).map(function (e) {
    const spots = Number(e['Spots Per Squad']) || 0;
    const taken = signups.filter(function (r) { return String(r['Event']) === String(e['Event']) && squadOf_(r['Squad']) === squad; }).length;
    return { event: String(e['Event']), date: String(e['Date'] || ''), slot: String(e['Slot'] || ''), type: String(e['Type'] || ''),
      self: on && ticked_(e['Self Sign-up']), spotsLeft: spots ? Math.max(0, spots - taken) : null, mine: mine.indexOf(String(e['Event'])) > -1 };
  });
  return { on: on, limit: Number(getSetting_('SIGNUP_LIMIT')) || 0, count: events.filter(function (e) { return e.mine; }).length, events: events };
}

/** A student signs up (or drops) on My Squad with their 800#. */
function clashStudentSignUp(id, event, remove) {
  if (String(getSetting_('SELF_SIGNUPS')).toUpperCase() !== 'ON') throw new Error('Sign-ups are done by your Generals right now.');
  const cache = CacheService.getScriptCache();
  const used = Number(cache.get('clash-selfsign')) || 0;
  if (used >= 600) throw new Error('The site is busy. Try again in a few minutes.');
  cache.put('clash-selfsign', String(used + 1), 600);
  const sid = id800_(id);
  const me = boardRow_(sid);
  if (!me) throw new Error('That 800 number isn\'t on the list.');
  const squad = squadOf_(me['Squad']);
  const ev = sheetRows_('Sign-up Events').filter(function (r) { return String(r['Event']) === String(event); })[0];
  if (!ev || !eventOpen_(ev) || !ticked_(ev['Self Sign-up'])) throw new Error('That event isn\'t open for sign-ups.');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const existing = sheetRows_('Sign-ups').filter(function (r) { return String(r['Event']) === String(event) && id800_(r['800#']) === sid; })[0];
    if (remove) { if (existing) getSheet_('Sign-ups').deleteRow(existing._row); }
    else if (!existing) {
      signupGuard_(sid, ev, squad);
      getSheet_('Sign-ups').appendRow([event, sid, String(me['Name'] || ''), squad, 'self',
        Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm')]);
    }
  } finally {
    lock.releaseLock();
  }
  return JSON.stringify(mySignups_(sid, squad));
}

// ── CAPTAIN VOTES ───────────────────────────────────────────────────
// What Generals can nominate for (Settings › NOMINATE_FOR, separated by ;).
function nominateFor_() {
  const v = String(getSetting_('NOMINATE_FOR') || '').split(';').map(function (s) { return s.trim(); }).filter(String);
  return v.length ? v : NOMINATE_FOR_;
}

// This squad's votes for anything with "Captain" in it: [{ for, nominee, id, votes }].
function captainTally_(squad) {
  const out = {};
  sheetRows_('Nominations').filter(function (n) { return squadOf_(n['Squad']) === squad && /captain/i.test(String(n['For'])); }).forEach(function (n) {
    const k = n['For'] + '|' + id800_(n['Nominee 800#']);
    const t = out[k] = out[k] || { for: String(n['For']), nominee: String(n['Nominee'] || ''), id: id800_(n['Nominee 800#']), voters: {} };
    t.voters[String(n['General'])] = true;
  });
  return Object.keys(out).map(function (k) { const t = out[k]; return { for: t.for, nominee: t.nominee, id: t.id, votes: Object.keys(t.voters).length }; })
    .sort(function (a, b) { return a.for < b.for ? -1 : a.for > b.for ? 1 : b.votes - a.votes; });
}

// ── SHIRTS ──────────────────────────────────────────────────────────
function paidShirt_(sid) {
  return sheetRows_('Shirt Orders').some(function (r) { return id800_(r['800#']) === sid && ticked_(r['Paid']); });
}

/** Paid shirts not yet counted → SHIRT_POINTS each, one event per squad. Returns what was sent. */
function sendShirtPoints() {
  const per = Number(getSetting_('SHIRT_POINTS')) || 0;
  if (!per) return ['SHIRT_POINTS is 0.'];
  const sh = getSheetSafe_('Shirt Orders');
  if (!sh) return ['No Shirt Orders tab yet. Run setup.'];
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sentCol = SHIRT_HEADERS_.indexOf('Points Sent') + 1, squadCol = SHIRT_HEADERS_.indexOf('Squad') + 1;
    const pts = { BLACK: 0, GOLD: 0, GREY: 0, PURPLE: 0 }, rows = [];
    sheetRows_('Shirt Orders').forEach(function (r) {
      if (!ticked_(r['Paid']) || String(r['Points Sent'] || '').trim()) return;
      let squad = squadOf_(r['Squad']);
      if (pts[squad] === undefined) { const b = boardRow_(id800_(r['800#'])); squad = b ? squadOf_(b['Squad']) : ''; if (squad) sh.getRange(r._row, squadCol).setValue(squad); }
      if (pts[squad] === undefined) return;
      pts[squad] += per;
      rows.push(r._row);
    });
    if (!rows.length) return ['No new paid shirts.'];
    const today = todayYmd_();
    getSheet_(TABS.EVENTS).appendRow(['', 'Mayhem shirt sales', 'shirt_sales', today, rows.length + ' shirts × ' + per, '',
      pts.BLACK || '', pts.GOLD || '', pts.GREY || '', pts.PURPLE || '']);
    rows.forEach(function (row) { sh.getRange(row, sentCol).setValue(today); });
    return [rows.length + ' shirts: ' + SQUAD_KEYS_.map(function (s) { return s + ' +' + pts[s]; }).join(', ')];
  } finally {
    lock.releaseLock();
  }
}
function menuSendShirtPoints_() { SpreadsheetApp.getUi().alert('Shirt sale points', sendShirtPoints().join('\n'), SpreadsheetApp.getUi().ButtonSet.OK); }

// ── SQUAD NIGHT DOOR (staff, signed in as admin on the site) ────────
function doorStatus_(sid) {
  const me = boardRow_(sid);
  if (!me) throw new Error('That 800 number isn\'t on the list.');
  const competitor = sheetRows_('Sign-ups').some(function (r) { return id800_(r['800#']) === sid; }) ||
    sheetRows_('MLC Roster').some(function (r) { return id800_(r['800#']) === sid; });
  const shirt = paidShirt_(sid);
  const already = sheetRows_('Squad Night').filter(function (r) { return id800_(r['800#']) === sid; })[0];
  return { id: sid, name: String(me['Name'] || ''), squad: squadOf_(me['Squad']),
    entry: competitor ? 'Free (competitor)' : shirt ? 'Free (shirt)' : 'Pay ' + (getSetting_('SQUAD_NIGHT_PRICE') || '$5'),
    free: competitor || shirt, already: already ? String(already['Time']) : '' };
}
function doorCounts_() {
  const c = { BLACK: 0, GOLD: 0, GREY: 0, PURPLE: 0 };
  sheetRows_('Squad Night').forEach(function (r) { const s = squadOf_(r['Squad']); if (c[s] !== undefined) c[s]++; });
  return c;
}
function clashDoorCheck(token, id) {
  requireAdmin_(token);
  return JSON.stringify({ student: doorStatus_(id800_(id)), counts: doorCounts_() });
}
function clashDoorAdmit(token, id, paid) {
  requireAdmin_(token);
  const sid = id800_(id);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const s = doorStatus_(sid);
    if (s.already) throw new Error(s.name + ' already came in at ' + s.already + '.');
    if (!s.free && !paid) throw new Error('Mark them as paid first.');
    getSheet_('Squad Night').appendRow([Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'), sid, s.name, s.squad,
      s.entry, s.free ? '' : (getSetting_('SQUAD_NIGHT_PRICE') || '$5'), 'door']);
  } finally {
    lock.releaseLock();
  }
  return JSON.stringify({ student: doorStatus_(sid), counts: doorCounts_() });
}

/** SQUAD_NIGHT_POINTS per student who came, as one event, once. */
function sendSquadNightPoints() {
  const per = Number(getSetting_('SQUAD_NIGHT_POINTS')) || 0;
  if (!per) return ['SQUAD_NIGHT_POINTS is 0.'];
  if (readEventsTab_().some(function (e) { return String(e.type) === 'squad_night' && String(e.date) === todayYmd_(); })) return ['Squad Night points were already sent today.'];
  const c = doorCounts_();
  const total = SQUAD_KEYS_.reduce(function (n, s) { return n + c[s]; }, 0);
  if (!total) return ['Nobody has been checked in yet.'];
  getSheet_(TABS.EVENTS).appendRow(['', 'Squad Night attendance', 'squad_night', todayYmd_(), total + ' students × ' + per, '',
    c.BLACK * per || '', c.GOLD * per || '', c.GREY * per || '', c.PURPLE * per || '']);
  return [SQUAD_KEYS_.map(function (s) { return s + ' ' + c[s] + ' → +' + c[s] * per; }).join(', ')];
}
function menuSendSquadNightPoints_() { SpreadsheetApp.getUi().alert('Squad Night points', sendSquadNightPoints().join('\n'), SpreadsheetApp.getUi().ButtonSet.OK); }
