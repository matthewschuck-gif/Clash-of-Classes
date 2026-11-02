/**
 * Clash of Classes -- Apps Script backend
 * File 5: Generals HQ and the Mountaineer Leadership Corps (from the Squad Generals Handbook
 * and the MLC Handbook).
 *
 * Tabs (all edited in the Sheet; Run setup makes them):
 *   MLC Roster      800# | Name | Squad | Role | Show On My Squad
 *                   Role: General, Event Captain, Hype Captain, Participation Captain, Athletics
 *                   Captain, Art Captain, Music Captain, Dance Captain, Spirit Captain, Delegate,
 *                   Day Camp Counselor, Pathfinder, Steward, Scout, Vanguard. Vanguard stays
 *                   private (Trailback Panel confidentiality) unless Show On My Squad is ticked.
 *   Squad Themes    Squad | Theme | What It Means | Service Project | Service Goal | Service Progress |
 *                   Trait | Lore Title | Lore | Lore Image | Show Lore
 *                   The lore columns are space for each squad's story (colors and traits are still
 *                   being decided). Nothing shows until Show Lore is ticked; then My Squad, Generals
 *                   HQ and the squad meeting slides show it. Lore Image: an https:// image link.
 *   Sign-up Events  Event | Date | Spots Per Squad | Open
 *   Sign-ups        Event | 800# | Name | Squad | Added By | Added At
 *   Mayhem Checklist Squad | Item | Done | Done By
 *   Nominations     Date | General | Squad | For | Nominee 800# | Nominee | Why   (Trailback Panel each
 *                   marking period, next year's Generals; staff make the final choice)
 *   General Letters Updated | General | Squad | Letter | Text   (the two end-of-year letters)
 *
 * Generals sign in with their 800# (Role = General on MLC Roster) plus GENERALS_CODE from
 * Settings. They only ever see and change their own squad. The squad roster comes from the
 * Student Board (first name + last initial only).
 */

const MLC_HEADERS_ = ['800#', 'Name', 'Squad', 'Role', 'Show On My Squad'];
const THEME_HEADERS_ = ['Squad', 'Theme', 'What It Means', 'Service Project', 'Service Goal', 'Service Progress',
  'Trait', 'Lore Title', 'Lore', 'Lore Image', 'Show Lore', 'Mascot', 'Mascot Name', 'Mascot Image'];
const SIGNUP_EVENT_HEADERS_ = ['Event', 'Date', 'Spots Per Squad', 'Open'];
const SIGNUP_HEADERS_ = ['Event', '800#', 'Name', 'Squad', 'Added By', 'Added At'];
const MAYHEM_HEADERS_ = ['Squad', 'Item', 'Done', 'Done By'];
const NOMINATION_HEADERS_ = ['Date', 'General', 'Squad', 'For', 'Nominee 800#', 'Nominee', 'Why'];
const LETTER_HEADERS_ = ['Updated', 'General', 'Squad', 'Letter', 'Text'];
const NOMINATE_FOR_ = ['Trailback Panel (Vanguard)', 'Next year\'s Squad General'];
const LETTERS_ = ['To future Generals', 'To staff advisors'];
const MAYHEM_ITEMS_ = [
  'Captains assigned and briefed at least two weeks before Mayhem Week',
  'Attended all required after-school prep sessions',
  'Squad roster confirmed with Participation Captain',
  'Sign-ups submitted for every event by staff deadline',
  'Every squad member placed in at least one event before anyone doubles',
  'All participants notified in writing at least twice',
  'Morning announcements written and submitted each day of the week',
  'T-shirt sales tracked by Spirit Captain and submitted for bonus points',
  'Poster submitted to Art Captain by the judging deadline',
  'Squad Dance rehearsal organized by Dance Captain',
  'Video or hype content created and shared before the week begins',
  'Staff advisor check-in completed before Friday closing assembly',
];
const DEFAULT_THEMES_ = [
  ['BLACK', 'Freudenfreude', 'Joy in others\' success. Lifting others up is the standard.', '', '', '', '', '', '', '', false],
  ['GOLD', 'Hope', 'Belief that things get better. Lift the room when things are hard.', '', '', '', '', '', '', '', false],
  ['GREY', 'Belonging', 'Everyone has a place here, especially new students and quieter voices.', '', '', '', '', '', '', '', false],
  ['PURPLE', 'Curiosity', 'Ask more questions. Lean into learning, creativity and trying new things.', '', '', '', '', '', '', '', false],
];

function ensureGeneralsTabs_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const make = function (name, headers, seed) {
    let sh = ss.getSheetByName(name);
    if (sh) return sh;
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#490e6f').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    if (seed && seed.length) sh.getRange(2, 1, seed.length, headers.length).setValues(seed);
    return sh;
  };
  make('MLC Roster', MLC_HEADERS_);
  const themes = make('Squad Themes', THEME_HEADERS_, DEFAULT_THEMES_);
  // Sheets made before the lore columns existed: add the missing headers at the end.
  const have = themes.getRange(1, 1, 1, Math.max(themes.getLastColumn(), 1)).getValues()[0].map(function (h) { return String(h).trim(); });
  const missing = THEME_HEADERS_.filter(function (h) { return have.indexOf(h) === -1; });
  if (missing.length) {
    themes.getRange(1, have.filter(String).length + 1, 1, missing.length).setValues([missing]).setFontWeight('bold').setBackground('#490e6f').setFontColor('#ffffff');
  }
  themes.getRange(2, THEME_HEADERS_.indexOf('Show Lore') + 1, 4, 1).insertCheckboxes();
  make('Sign-up Events', SIGNUP_EVENT_HEADERS_);
  make('Sign-ups', SIGNUP_HEADERS_);
  const seed = [];
  SQUAD_KEYS_.forEach(function (s) { MAYHEM_ITEMS_.forEach(function (i) { seed.push([s, i, false, '']); }); });
  make('Mayhem Checklist', MAYHEM_HEADERS_, seed);
  make('Nominations', NOMINATION_HEADERS_);
  make('General Letters', LETTER_HEADERS_);
}

function sheetRows_(name) {
  const sh = getSheetSafe_(name);
  if (!sh || sh.getLastRow() < 2) return [];
  const values = sh.getDataRange().getValues();
  const h = values[0].map(function (x) { return String(x).trim(); });
  return values.slice(1).map(function (r, i) {
    const o = { _row: i + 2 };
    h.forEach(function (k, j) { o[k] = Object.prototype.toString.call(r[j]) === '[object Date]' ? Utilities.formatDate(r[j], Session.getScriptTimeZone(), 'yyyy-MM-dd') : r[j]; });
    return o;
  });
}

function squadOf_(v) { const s = String(v || '').trim().toUpperCase(); return s === 'GRAY' ? 'GREY' : s; }
function ticked_(v) { return v === true || /^(true|yes|y|x|✓)$/i.test(String(v || '').trim()); }
function id800_(v) { return String(v || '').replace(/\D/g, ''); }

// ── MLC roles (also used by My Squad) ───────────────────────────────
// { 800#: ['General', 'Hype Captain'] } for roles that may show publicly.
function publicRoles_() {
  const out = {};
  sheetRows_('MLC Roster').forEach(function (r) {
    const id = id800_(r['800#']);
    const role = String(r['Role'] || '').trim();
    if (!id || !role) return;
    const show = String(r['Show On My Squad']).trim() === '' ? !/vanguard/i.test(role) : ticked_(r['Show On My Squad']);
    if (show) (out[id] = out[id] || []).push(role);
  });
  return out;
}

function squadTheme_(squad) {
  const t = sheetRows_('Squad Themes').filter(function (r) { return squadOf_(r['Squad']) === squad; })[0];
  if (!t) return null;
  const out = { theme: String(t['Theme'] || ''), meaning: String(t['What It Means'] || ''), project: String(t['Service Project'] || ''),
    goal: Number(t['Service Goal']) || 0, progress: Number(t['Service Progress']) || 0, trait: '', lore: null };
  // Trait and lore only once Show Lore is ticked, so drafts stay in the Sheet.
  if (ticked_(t['Show Lore'])) {
    out.trait = String(t['Trait'] || '').trim();
    const text = String(t['Lore'] || '').trim();
    const image = String(t['Lore Image'] || '').trim();
    if (text || image) out.lore = { title: String(t['Lore Title'] || '').trim(), text: text, image: /^https:\/\//i.test(image) ? image : '' };
  }
  return out;
}

// ── GENERALS SIGN-IN ────────────────────────────────────────────────
function clashGeneralsSignIn(id, code) {
  const cache = CacheService.getScriptCache();
  const fails = Number(cache.get('hq-fails')) || 0;
  if (fails >= 15) throw new Error('Too many tries. Wait 10 minutes.');
  const want = String(getSetting_('GENERALS_CODE') || '').trim();
  const sid = id800_(id);
  const me = sheetRows_('MLC Roster').filter(function (r) { return id800_(r['800#']) === sid && /general/i.test(String(r['Role'])); })[0];
  if (!want || !me || String(code || '').trim().toLowerCase() !== want.toLowerCase()) {
    cache.put('hq-fails', String(fails + 1), 600);
    throw new Error(!want ? 'Generals HQ isn\'t open yet (no GENERALS_CODE on Settings).' : 'That 800# and code don\'t match a General.');
  }
  const token = Utilities.getUuid();
  const squad = squadOf_(me['Squad']);
  cache.put('hq-' + token, JSON.stringify({ id: sid, name: String(me['Name'] || ''), squad: squad }), 6 * 3600);
  return JSON.stringify({ token: token, name: String(me['Name'] || ''), squad: squad });
}

function hqSession_(token) {
  const s = token ? CacheService.getScriptCache().get('hq-' + token) : null;
  if (!s) throw new Error('Please sign in to Generals HQ again.');
  return JSON.parse(s);
}

// ── HQ DATA ─────────────────────────────────────────────────────────
function clashGeneralsHq(token) {
  const me = hqSession_(token);
  const squad = me.squad;
  // Squad roster from the Student Board: first name + last initial only.
  const board = sheetRows_(getSetting_('STUDENT_BOARD_TAB') || 'Student Board').filter(function (r) { return squadOf_(r['Squad']) === squad; });
  const signups = sheetRows_('Sign-ups').filter(function (r) { return squadOf_(r['Squad']) === squad; });
  const events = sheetRows_('Sign-up Events').filter(function (r) { return String(r['Event']).trim(); });
  const placed = {};
  signups.forEach(function (r) { const id = id800_(r['800#']); (placed[id] = placed[id] || []).push(String(r['Event'])); });
  const roles = publicRoles_();
  const roster = board.map(function (r) {
    const id = id800_(r['800#']);
    return { id: id, name: String(r['Name'] || ''), grade: String(r['Grade'] || ''), events: placed[id] || [], roles: roles[id] || [] };
  }).sort(function (a, b) { return a.events.length - b.events.length || a.name.localeCompare(b.name); });
  const totals = squadTotals_();
  const recent = readEventsTab_().filter(function (e) { return (Number(e.pts[squad]) || 0) > 0; })
    .sort(function (a, b) { return String(b.date).localeCompare(String(a.date)); }).slice(0, 5)
    .map(function (e) { return { name: publicEvent_(e).name, date: e.date, pts: Number(e.pts[squad]) || 0 }; }); // no student names (03_Web.gs)
  const leaders = sheetRows_('MLC Roster').filter(function (r) { return squadOf_(r['Squad']) === squad && !/vanguard/i.test(String(r['Role'])); })
    .map(function (r) { return { name: String(r['Name'] || ''), role: String(r['Role'] || '') }; });
  let month = '', trait = '';
  try { const v = loadAppData_('coc').value || {}; month = (v.metrics && v.metrics.month) || ''; trait = v.currentTrait || ''; } catch (e) { /* none */ }
  return JSON.stringify({
    me: me, squad: squad, theme: squadTheme_(squad), standing: standingFor_(squad, totals), totals: totals,
    roster: roster, notPlaced: roster.filter(function (s) { return !s.events.length; }).length,
    events: events.map(function (r) {
      const n = signups.filter(function (s) { return String(s['Event']) === String(r['Event']); }).length;
      return { event: String(r['Event']), date: String(r['Date'] || ''), slot: String(r['Slot'] || ''), type: String(r['Type'] || ''),
        spots: Number(r['Spots Per Squad']) || 0, open: String(r['Open']).trim() === '' || ticked_(r['Open']), taken: n };
    }),
    mayhem: sheetRows_('Mayhem Checklist').filter(function (r) { return squadOf_(r['Squad']) === squad; })
      .map(function (r) { return { row: r._row, item: String(r['Item']), done: ticked_(r['Done']) }; }),
    recent: recent, leaders: leaders, month: month, trait: trait,
    challenges: challengeProgress_().filter(function (c) { return c.active; }),
    duels: JSON.parse(clashDuels()).filter(function (d) { return d.squads.indexOf(squad) > -1; }),
    lore: loreFor_(squad),
    poll: (function () { try { return pollAnnouncement_(); } catch (e) { return ''; } })(), // 06_Polls.gs
    nominateFor: nominateFor_(), // Settings › NOMINATE_FOR (11_MayhemWeek.gs)
    captainTally: captainTally_(squad),
    designs: hqDesigns_(squad), // 12_ShirtDesigns.gs
    shirtVoting: String(getSetting_('SHIRT_VOTING')).toUpperCase() === 'ON',
    signupLimit: Number(getSetting_('SIGNUP_LIMIT')) || 0,
    nominations: sheetRows_('Nominations').filter(function (r) { return id800_(r['General']) === me.id || String(r['General']) === me.name; })
      .map(function (r) { return { date: String(r['Date']), for: String(r['For']), nominee: String(r['Nominee']), why: String(r['Why']) }; }),
    letters: LETTERS_.map(function (l) {
      const r = sheetRows_('General Letters').filter(function (x) { return id800_(x['General']) === me.id && String(x['Letter']) === l; })[0];
      return { letter: l, text: r ? String(r['Text']) : '', updated: r ? String(r['Updated']) : '' };
    }),
  });
}

// A General nominates a squad member (Trailback Panel or next year's General). Staff decide.
function clashGeneralsNominate(token, forWhat, id, why) {
  const me = hqSession_(token);
  if (nominateFor_().indexOf(forWhat) === -1) throw new Error('Pick what you are nominating for.');
  if (sheetRows_('Nominations').some(function (n) { return String(n['General']) === me.id && String(n['For']) === forWhat && id800_(n['Nominee 800#']) === id800_(id); })) {
    throw new Error('You already nominated that student for ' + forWhat + '.');
  }
  if (String(why || '').trim().length < 10) throw new Error('Say why in a sentence or two.');
  const sid = id800_(id);
  const s = sheetRows_(getSetting_('STUDENT_BOARD_TAB') || 'Student Board').filter(function (r) { return id800_(r['800#']) === sid && squadOf_(r['Squad']) === me.squad; })[0];
  if (!s) throw new Error('Pick a student from your squad.');
  getSheet_('Nominations').appendRow([Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'), me.id, me.squad, forWhat, sid, String(s['Name'] || ''), String(why).trim()]);
  return clashGeneralsHq(token);
}

// Saves (or replaces) one of the General's two end-of-year letters.
function clashGeneralsLetter(token, letter, text) {
  const me = hqSession_(token);
  if (LETTERS_.indexOf(letter) === -1) throw new Error('Unknown letter.');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = getSheet_('General Letters');
    const r = sheetRows_('General Letters').filter(function (x) { return id800_(x['General']) === me.id && String(x['Letter']) === letter; })[0];
    const row = [Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'), me.id, me.squad, letter, String(text || '')];
    if (r) sh.getRange(r._row, 1, 1, row.length).setValues([row]); else sh.appendRow(row);
  } finally {
    lock.releaseLock();
  }
  return clashGeneralsHq(token);
}

// Add (or remove) a squad member for a sign-up event.
function clashGeneralsSignUp(token, event, id, remove) {
  const me = hqSession_(token);
  const sid = id800_(id);
  const ev = sheetRows_('Sign-up Events').filter(function (r) { return String(r['Event']) === String(event); })[0];
  if (!ev) throw new Error('That event isn\'t on the Sign-up Events tab.');
  const student = sheetRows_(getSetting_('STUDENT_BOARD_TAB') || 'Student Board').filter(function (r) { return id800_(r['800#']) === sid && squadOf_(r['Squad']) === me.squad; })[0];
  if (!student) throw new Error('That student isn\'t on your squad.');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = getSheet_('Sign-ups');
    const existing = sheetRows_('Sign-ups').filter(function (r) { return String(r['Event']) === String(event) && id800_(r['800#']) === sid; })[0];
    if (remove) { if (existing) sh.deleteRow(existing._row); return clashGeneralsHq(token); }
    if (existing) return clashGeneralsHq(token);
    if (!(String(ev['Open']).trim() === '' || ticked_(ev['Open']))) throw new Error('Sign-ups for that event are closed.');
    signupGuard_(sid, ev, me.squad); // spots, the SIGNUP_LIMIT and same-time events (11_MayhemWeek.gs)
    sh.appendRow([event, sid, String(student['Name'] || ''), me.squad, me.name || me.id,
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm')]);
  } finally {
    lock.releaseLock();
  }
  return clashGeneralsHq(token);
}

function clashGeneralsChecklist(token, row, done) {
  const me = hqSession_(token);
  const r = sheetRows_('Mayhem Checklist').filter(function (x) { return x._row === Number(row); })[0];
  if (!r || squadOf_(r['Squad']) !== me.squad) throw new Error('That item isn\'t on your squad\'s checklist.');
  getSheet_('Mayhem Checklist').getRange(r._row, 3, 1, 2).setValues([[!!done, done ? (me.name || me.id) : '']]);
  return clashGeneralsHq(token);
}

function clashGeneralsService(token, progress) {
  const me = hqSession_(token);
  const r = sheetRows_('Squad Themes').filter(function (x) { return squadOf_(x['Squad']) === me.squad; })[0];
  if (!r) throw new Error('Your squad has no row on the Squad Themes tab.');
  getSheet_('Squad Themes').getRange(r._row, THEME_HEADERS_.indexOf('Service Progress') + 1).setValues([[Math.max(0, Number(progress) || 0)]]);
  return clashGeneralsHq(token);
}
