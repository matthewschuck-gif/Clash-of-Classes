/**
 * Clash of Classes -- Apps Script backend
 * File 47: families take part. Four pieces, each OFF until you turn it on in Settings:
 *
 * FAMILY_CHALLENGE  A weekly at-home challenge (FAMILY CHALLENGES tab: Show | Week Of | Challenge).
 *   It shows on the family update and on My Squad. The student tells us they did it (My Squad);
 *   it goes on the Admin › ✅ Service list marked Family, and staff approve it for points (type family).
 * FAMILY_SHOUTOUTS  Families send a shout-out about their student from the family update page
 *   (FAMILY SHOUT-OUTS tab). Staff approve in Admin › ✅ Service: 1, 3 or 5 points, and tick
 *   "Spotlight" to put it on the Squad Spotlight tab for the week (you still tick Show there).
 * FAMILY_PULSE  A question for families (FAMILY PULSE tab: Show | Question | Choices | From | Through).
 *   They answer on the family update page with their student's first and last name and grade. Each
 *   family's answer adds FAMILY_PULSE_POINTS to that student's squad (paid once a day as one row per
 *   squad). Families see the results after answering. Answers on FAMILY PULSE ANSWERS (staff only).
 * FAMILY_CHECKIN  Family Night / Squad Night check-in at the door: clashofclasses.org/?page=familydoor
 *   (admin password). Find the student, check the family in. Each family adds FAMILY_CHECKIN_POINTS
 *   to the squad (paid once a day). FAMILY CHECK-INS tab.
 * Families never see another student's name: they type their own student's name, and only counts
 * come back.
 */

const FAMILY_TABS_ = [
  ['Family Challenges', ['Show', 'Week Of', 'Challenge']],
  ['Family Shout-outs', ['ID', 'Sent', 'Student', 'Grade', 'Squad', '800#', 'What', 'Trait', 'From', 'Status', 'Points', 'Spotlight', 'Decided At']],
  ['Family Pulse', ['Show', 'Question', 'Choices', 'From', 'Through']],
  ['Family Pulse Answers', ['Question', 'Student', 'Grade', 'Squad', '800#', 'Choice', 'At', 'Paid']],
  ['Family Check-ins', ['Event', 'Date', 'Student', 'Squad', '800#', 'Adults', 'At', 'Paid']],
];

function familyOn_(key) { return String(getSetting_(key) || 'OFF').toUpperCase() === 'ON'; }

function ensureFamilyTabs_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  FAMILY_TABS_.forEach(function (t) {
    if (ss.getSheetByName(t[0])) return;
    const sh = ss.insertSheet(t[0]);
    sh.getRange(1, 1, 1, t[1].length).setValues([t[1]]).setFontWeight('bold').setBackground('#490e6f').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    if (t[0] === 'Family Challenges') sh.appendRow([false, todayYmd_(), 'Example: eat one dinner together this week with no phones at the table. (Untick or replace.)']);
  });
}

// Finds a student by first + last name and grade (families type these). null = no single match.
function findStudentByName_(first, last, grade) {
  const f = String(first || '').trim().toLowerCase(), l = String(last || '').trim().toLowerCase(), g = String(grade || '').replace(/\D/g, '');
  if (!f || !l) return null;
  const hits = sheetRows_('Students').filter(function (r) {
    return String(r['First Name'] || '').trim().toLowerCase() === f && String(r['Last Name'] || '').trim().toLowerCase() === l && (!g || String(r['Grade'] || '').replace(/\D/g, '') === g);
  });
  if (hits.length !== 1) return null;
  const r = hits[0];
  return { id: id800_(r['800#']), first: String(r['First Name']).trim(), last: String(r['Last Name']).trim(), grade: String(r['Grade'] || ''), squad: squadOf_(r['Squad']) };
}

function familyLimit_(key, max, seconds) {
  const cache = CacheService.getScriptCache(), n = Number(cache.get(key)) || 0;
  if (n >= max) throw new Error('That\'s a lot at once. Please try again later.');
  cache.put(key, String(n + 1), seconds);
}

// ── 1. The weekly family challenge ──
function familyChallenge_() {
  if (!familyOn_('FAMILY_CHALLENGE')) return null;
  const today = todayYmd_();
  const r = sheetRows_('Family Challenges').filter(function (x) { const w = String(x['Week Of']).slice(0, 10); return ticked_(x['Show']) && String(x['Challenge'] || '').trim() && w <= today; })
    .sort(function (a, b) { return String(b['Week Of']).localeCompare(String(a['Week Of'])); })[0];
  return r ? { week: String(r['Week Of']).slice(0, 10), text: String(r['Challenge']).trim() } : null;
}
function clashFamilyChallenge() { return JSON.stringify(familyChallenge_()); }

// ── 2. Family shout-outs ──
function clashFamilyShoutout(json) {
  if (!familyOn_('FAMILY_SHOUTOUTS')) throw new Error('Family shout-outs are not open right now.');
  const p = JSON.parse(json || '{}');
  const first = String(p.first || '').trim().slice(0, 40), last = String(p.last || '').trim().slice(0, 40), what = String(p.what || '').trim().slice(0, 600);
  if (!first || !last) throw new Error('Type your student\'s first and last name.');
  if (what.length < 15) throw new Error('Tell us what they did (a sentence or two).');
  familyLimit_('fam-shout-' + (first + last).toLowerCase().replace(/\W/g, ''), 3, 86400);
  familyLimit_('fam-shout-all', 120, 3600);
  ensureFamilyTabs_();
  const st = findStudentByName_(first, last, p.grade);
  getSheet_('Family Shout-outs').appendRow(['F' + Date.now().toString(36), Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'),
    st ? st.first + ' ' + st.last : first + ' ' + last, st ? st.grade : String(p.grade || ''), st ? st.squad : '', st ? st.id : '', what,
    String(p.trait || '').slice(0, 60), String(p.from || '').trim().slice(0, 60), 'Waiting', '', false, '']);
  return JSON.stringify({ ok: true });
}

function familyShoutList_() {
  return sheetRows_('Family Shout-outs').filter(function (r) { return String(r['ID'] || '').trim(); }).map(function (r) {
    return { id: String(r['ID']), sent: String(r['Sent']), student: String(r['Student'] || ''), grade: String(r['Grade'] || ''), squad: squadOf_(r['Squad']),
      what: String(r['What'] || ''), trait: String(r['Trait'] || ''), from: String(r['From'] || ''), status: String(r['Status'] || 'Waiting'), points: Number(r['Points']) || 0 };
  });
}

/** Staff: decision approve (points 1/3/5, spotlight true/false) or deny. */
function clashFamilyShoutDecide(token, id, decision, points, spotlight) {
  requireAdmin_(token);
  const r = sheetRows_('Family Shout-outs').filter(function (x) { return String(x['ID']) === String(id); })[0];
  if (!r) throw new Error('That shout-out is gone.');
  if (String(r['Status']) !== 'Waiting') throw new Error('Already ' + String(r['Status']).toLowerCase() + '.');
  const sh = getSheet_('Family Shout-outs'), set = function (h, v) { const c = headerCol_(sh, h); if (c) sh.getRange(r._row, c).setValue(v); };
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  if (decision === 'deny') { set('Status', 'Denied'); set('Decided At', stamp); return JSON.stringify({ list: familyShoutList_() }); }
  const squad = squadOf_(r['Squad']);
  points = Number(points) || 0;
  if (points && SQUAD_KEYS_.indexOf(squad) === -1) throw new Error('No squad matched this student. Type their squad on the Family Shout-outs tab first.');
  if (points && [1, 3, 5].indexOf(points) === -1) throw new Error('Pick 1, 3 or 5 points.');
  if (points) {
    const pts = { BLACK: '', GOLD: '', GREY: '', PURPLE: '' }; pts[squad] = points;
    getSheet_('events').appendRow(['', 'Family shout-out: ' + shortName_(r['Student']), 'family', todayYmd_(), '', '', pts.BLACK, pts.GOLD, pts.GREY, pts.PURPLE]);
  }
  if (spotlight) {
    ensureSpotlightTab_();
    getSheet_('Squad Spotlight').appendRow([false, mondayOf_(todayYmd_()), squad, shortName_(r['Student']), String(r['What']).trim(), 'Family shout-out' + (r['From'] ? ' from ' + r['From'] : '')]);
  }
  set('Status', 'Approved'); set('Points', points); set('Spotlight', !!spotlight); set('Decided At', stamp);
  return JSON.stringify({ list: familyShoutList_() });
}

// ── 3. The Family Pulse ──
function familyPulse_() {
  if (!familyOn_('FAMILY_PULSE')) return null;
  const today = todayYmd_();
  const r = sheetRows_('Family Pulse').filter(function (x) {
    const a = String(x['From'] || '').slice(0, 10), b = String(x['Through'] || '').slice(0, 10);
    return ticked_(x['Show']) && String(x['Question'] || '').trim() && pollChoices_(x['Choices']).length >= 2 && (!a || a <= today) && (!b || b >= today);
  })[0];
  if (!r) return null;
  const q = String(r['Question']).trim(), choices = pollChoices_(r['Choices']), counts = {};
  choices.forEach(function (c) { counts[c] = 0; });
  const answers = sheetRows_('Family Pulse Answers').filter(function (a) { return String(a['Question']) === q; });
  answers.forEach(function (a) { const c = String(a['Choice']); if (counts[c] !== undefined) counts[c]++; });
  return { question: q, choices: choices, counts: counts, total: answers.length };
}

function clashFamilyPulseAnswer(json) {
  const fp = familyPulse_();
  if (!fp) throw new Error('There is no family question open right now.');
  const p = JSON.parse(json || '{}');
  if (fp.choices.indexOf(String(p.choice)) === -1) throw new Error('Pick an answer.');
  const st = findStudentByName_(p.first, p.last, p.grade);
  if (!st) throw new Error('We couldn\'t find that student. Check the spelling of their first and last name and their grade.');
  familyLimit_('fam-pulse-all', 300, 3600);
  ensureFamilyTabs_();
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (sheetRows_('Family Pulse Answers').some(function (a) { return String(a['Question']) === fp.question && id800_(a['800#']) === st.id; })) throw new Error('Your family already answered this one. Thank you!');
    getSheet_('Family Pulse Answers').appendRow([fp.question, st.first + ' ' + st.last, st.grade, st.squad, st.id, String(p.choice), Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'), '']);
  } finally { lock.releaseLock(); }
  return JSON.stringify(familyPulse_());
}

// ── 4. Family Night check-ins (door, staff) ──
function clashFamilyDoorFind(token, query) {
  requireAdmin_(token);
  const q = String(query || '').trim().toLowerCase();
  if (q.length < 2) return '[]';
  return JSON.stringify(sheetRows_('Students').filter(function (r) {
    return (String(r['First Name'] || '') + ' ' + String(r['Last Name'] || '')).toLowerCase().indexOf(q) > -1 || id800_(r['800#']) === id800_(q);
  }).slice(0, 12).map(function (r) { return { id: id800_(r['800#']), name: String(r['First Name'] || '').trim() + ' ' + String(r['Last Name'] || '').trim(), grade: String(r['Grade'] || ''), squad: squadOf_(r['Squad']) }; }));
}

function clashFamilyDoorCheckIn(token, eventName, id, adults) {
  requireAdmin_(token);
  if (!familyOn_('FAMILY_CHECKIN')) throw new Error('Family check-in is off (Settings › FAMILY_CHECKIN).');
  eventName = String(eventName || '').trim().slice(0, 80);
  if (!eventName) throw new Error('Name the event first (e.g. Squad Night).');
  const sid = id800_(id), r = sheetRows_('Students').filter(function (x) { return id800_(x['800#']) === sid; })[0];
  if (!r) throw new Error('Pick a student from the list.');
  ensureFamilyTabs_();
  const today = todayYmd_();
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (sheetRows_('Family Check-ins').some(function (c) { return String(c['Event']) === eventName && String(c['Date']).slice(0, 10) === today && id800_(c['800#']) === sid; })) throw new Error('That family is already checked in. Welcome back!');
    getSheet_('Family Check-ins').appendRow([eventName, today, String(r['First Name']).trim() + ' ' + String(r['Last Name']).trim(), squadOf_(r['Squad']), sid, Math.max(1, Math.min(9, Number(adults) || 1)), Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'HH:mm'), '']);
  } finally { lock.releaseLock(); }
  const all = sheetRows_('Family Check-ins').filter(function (c) { return String(c['Event']) === eventName && String(c['Date']).slice(0, 10) === today; });
  const bySquad = { BLACK: 0, GOLD: 0, GREY: 0, PURPLE: 0 };
  all.forEach(function (c) { const s = squadOf_(c['Squad']); if (bySquad[s] !== undefined) bySquad[s]++; });
  return JSON.stringify({ total: all.length, bySquad: bySquad });
}

// ── Paying the family points (Pulse answers, check-ins): once a day, one row per squad each ──
function payFamilyPoints_() {
  const today = todayYmd_(), lines = [];
  [['Family Pulse Answers', 'FAMILY_PULSE_POINTS', 1, 'Family Pulse'], ['Family Check-ins', 'FAMILY_CHECKIN_POINTS', 2, 'Family Night']].forEach(function (t) {
    const sh = getSheetSafe_(t[0]);
    if (!sh) return;
    const per = Math.max(0, settingNum_(t[1], t[2]));
    const rows = sheetRows_(t[0]).filter(function (r) { return !String(r['Paid'] || '').trim() && SQUAD_KEYS_.indexOf(squadOf_(r['Squad'])) > -1; });
    if (!rows.length || !per) return;
    const pts = { BLACK: 0, GOLD: 0, GREY: 0, PURPLE: 0 };
    rows.forEach(function (r) { pts[squadOf_(r['Squad'])] += per; });
    getSheet_('events').appendRow(['', t[3] + ': ' + rows.length + ' famil' + (rows.length === 1 ? 'y' : 'ies'), 'family', today, '', '', pts.BLACK || '', pts.GOLD || '', pts.GREY || '', pts.PURPLE || '']);
    const c = headerCol_(sh, 'Paid');
    rows.forEach(function (r) { if (c) sh.getRange(r._row, c).setValue(today); });
    lines.push(t[3] + ': ' + rows.length);
  });
  return lines;
}

function maybePayFamilyPoints_() {
  if (!familyOn_('FAMILY_PULSE') && !familyOn_('FAMILY_CHECKIN')) return;
  const props = PropertiesService.getScriptProperties(), today = todayYmd_();
  if (props.getProperty('family-paid') === today) return;
  props.setProperty('family-paid', today);
  payFamilyPoints_();
}

/** What the family page shows (only the pieces that are on). */
function familyExtras_() {
  return { challenge: familyChallenge_(), shoutouts: familyOn_('FAMILY_SHOUTOUTS'), pulse: familyPulse_(), traits: (function () { try { return doughTraits_(); } catch (e) { return []; } })() };
}

function serveFamilyDoor_() {
  return HtmlService.createHtmlOutputFromFile('FamilyDoor').setTitle('Family check-in · Clash of Classes')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// file-stamp (tools/stamp-files.js): lets Admin › System tell whether this file is the newest one
var FILE_STAMP_47_Families = '534d2d23f5';
