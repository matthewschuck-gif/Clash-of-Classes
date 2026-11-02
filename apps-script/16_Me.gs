/**
 * Clash of Classes -- Apps Script backend
 * File 16: the student dashboard ("My Climb"), opened at the STUDENT link with ?page=me.
 *
 * Sign-in: students open it signed in with their school Google account. That deployment is
 * "Execute as: Me" and "Who has access: Anyone within EASD", so Session.getActiveUser() is the
 * student. Their 800# comes from the Student Board's Email column (Climb-er fills it from
 * PowerSchool), or from STUDENT_EMAIL_PATTERN (e.g. {800}@easdpa.org). No one can open another
 * student's dashboard by typing an 800#. Staff signed in as admin can preview any student.
 *
 * What a student sees is only about themselves, plus squad AVERAGES to compare against (never
 * another student's numbers): positive contributions, STAR growth, attendance %.
 *
 * DESTINATION TICKETS (tabs Destination Tickets, Student Tickets): like the board game, a
 * student draws TICKET_CHOICES tickets and keeps at least one (up to TICKETS_HELD at a time).
 * A ticket is a personal goal measured from when it was drawn: bills, summits, shoutouts, star,
 * polls, signups, buddy, tools (different kinds of things), streak. Join with + to need both.
 * Finishing one earns the squad its Points (once GAME_AWARDS is ON) and frees a slot.
 *
 * MY MASCOT (tab Avatars): each student makes their own version of their squad's animal (fur, eyes,
 * mood, hat, color); the gear on it comes from what they've done (AVATAR_GEAR).
 */

const TICKET_HEADERS_ = ['Show', 'Ticket', 'Text', 'Requirement', 'Points'];
const STUDENT_TICKET_HEADERS_ = ['800#', 'Squad', 'Ticket', 'Drawn', 'Baseline', 'Status', 'Completed', 'Awarded'];
const AVATAR_HEADERS_ = ['800#', 'Look', 'Updated'];
const MASCOTS_ = { GOLD: 'Cottontail rabbit', PURPLE: 'Ruffled grouse', GREY: 'Fisher', BLACK: 'White-tailed deer' };
const LOOK_CHOICES_ = { fur: 3, eyes: 4, mood: 4, hat: 5, accent: 6 }; // GameUi.html LOOK: the student's own version of their squad's animal

function ensureMeTabs_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  [['Destination Tickets', TICKET_HEADERS_, [
      [true, 'First Steps', 'Turn in a Trait Dough bill.', 'bills: 1', 2],
      [true, 'Dough Run', 'Turn in 3 Trait Dough bills.', 'bills: 3', 4],
      [true, 'Echo Chamber', 'Vote in 5 Daily Polls.', 'polls: 5', 3],
      [true, 'Summit Seeker', 'Reach a Mountaineer Mindset summit.', 'summits: 1', 4],
      [true, 'Kind Words', 'Earn a shout-out.', 'shoutouts: 1', 3],
      [true, 'Growth Spurt', 'Earn STAR growth points.', 'star: 1', 4],
      [true, 'Squad Up', 'Sign up for a squad event.', 'signups: 1', 3],
      [true, 'Trail Mix', 'Add something 3 different ways (bills, polls, summits, shout-outs, events...).', 'tools: 3', 6],
      [true, 'Double Switchback', 'A Trait Dough bill and 3 poll votes.', 'bills: 1 + polls: 3', 4],
      [true, 'Steady Climber', 'Keep a 3-week contribution streak.', 'streak: 3', 5]]],
   ['Student Tickets', STUDENT_TICKET_HEADERS_, null],
   ['Avatars', AVATAR_HEADERS_, null]].forEach(function (t) {
    if (ss.getSheetByName(t[0])) return;
    const sh = ss.insertSheet(t[0]);
    sh.getRange(1, 1, 1, t[1].length).setValues([t[1]]).setFontWeight('bold').setBackground('#490e6f').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    if (t[2]) sh.getRange(2, 1, t[2].length, t[1].length).setValues(t[2]);
  });
}

// ── WHO IS SIGNED IN ────────────────────────────────────────────────
function boardRows_() { return sheetRows_(getSetting_('STUDENT_BOARD_TAB') || 'Student Board'); }

function signedInId_() {
  const email = String(Session.getActiveUser().getEmail() || '').trim().toLowerCase();
  if (!email) throw new Error('SIGN_IN: Open My Climb with the student link, signed in with your school Google account.');
  const row = boardRows_().filter(function (r) { return String(r['Email'] || '').trim().toLowerCase() === email; })[0];
  if (row) return id800_(row['800#']);
  const pattern = String(getSetting_('STUDENT_EMAIL_PATTERN') || '').trim().toLowerCase();
  if (pattern.indexOf('{800}') > -1) {
    const re = new RegExp('^' + pattern.replace(/[.*+?^$()|[\]\\]/g, '\\$&').replace('\\{800\\}', '(\\d{5,})').replace('{800}', '(\\d{5,})') + '$');
    const m = email.match(re);
    if (m) return m[1];
  }
  throw new Error('NOT_FOUND: ' + email + ' isn\'t matched to a student yet. Ask your homeroom teacher.');
}

// My Climb is for the school day: open MY_CLIMB_HOURS (07:00-16:00) on school days only
// (weekdays not on the School Calendar tab). Staff previews are never locked.
function climbOpen_(now) {
  now = now || new Date();
  const tz = Session.getScriptTimeZone();
  const hours = String(getSetting_('MY_CLIMB_HOURS') || '07:00-16:00').match(/(\d{1,2}):(\d{2})\s*-\s*(\d{1,2}):(\d{2})/);
  if (!hours) return true;
  const hm = Utilities.formatDate(now, tz, 'HH:mm');
  const pad = function (h, m) { return ('0' + h).slice(-2) + ':' + m; };
  if (hm < pad(hours[1], hours[2]) || hm >= pad(hours[3], hours[4])) return false;
  if (String(getSetting_('MY_CLIMB_SCHOOL_DAYS') || 'ON').toUpperCase() === 'OFF') return true;
  const dow = Number(Utilities.formatDate(now, tz, 'u')); // 1 = Monday … 7 = Sunday
  if (dow > 5) return false;
  try { return schoolCalendar_().noSchool.indexOf(Utilities.formatDate(now, tz, 'yyyy-MM-dd')) === -1; } catch (e) { return true; }
}

// Which student: the signed-in one, or (staff signed in as admin) any student to preview.
function meId_(asId, adminToken) {
  if (asId && adminToken) { requireAdmin_(adminToken); return id800_(asId); }
  if (!climbOpen_()) throw new Error('CLOSED: My Climb is open ' + String(getSetting_('MY_CLIMB_HOURS') || '07:00-16:00').replace('-', ' to ') + ' on school days. See you on the trail!');
  return signedInId_();
}

// ── NUMBERS: ME VS. MY SQUAD ────────────────────────────────────────
const METRICS_ = [
  ['bills', 'Trait Dough Bills', 'Trait Dough bills'], ['mindset', 'Mindset Summits', 'Mindset summits'],
  ['shoutouts', 'Shout-outs', 'Shout-outs'], ['star', 'STAR Growth Points', 'STAR growth points'],
  ['streak', 'Streak (weeks)', 'Week streak'], ['attendance', 'Attendance %', 'Attendance']];

function metricsFor_(id, squad, rows) {
  rows = rows || boardRows_();
  const mine = rows.filter(function (r) { return id800_(r['800#']) === id; })[0] || {};
  const mates = rows.filter(function (r) { return squadOf_(r['Squad']) === squad; });
  return METRICS_.map(function (m) {
    const vals = mates.map(function (r) { return r[m[1]]; }).filter(function (v) { return String(v).trim() !== ''; }).map(Number).filter(function (v) { return !isNaN(v); });
    const avg = vals.length ? Math.round(10 * vals.reduce(function (a, b) { return a + b; }, 0) / vals.length) / 10 : null;
    const raw = mine[m[1]];
    return { key: m[0], label: m[2], me: String(raw).trim() === '' || raw === undefined ? null : Number(raw), squadAvg: avg, pct: m[0] === 'attendance' };
  }).filter(function (m) { return m.me !== null || m.squadAvg !== null; });
}

// ── DESTINATION TICKETS ─────────────────────────────────────────────
function ticketDefs_() {
  return sheetRows_('Destination Tickets').filter(function (r) { return ticked_(r['Show']) && String(r['Ticket'] || '').trim(); }).map(function (r) {
    return { ticket: String(r['Ticket']).trim(), text: String(r['Text'] || ''), req: String(r['Requirement'] || ''), points: Number(r['Points']) || 0 };
  });
}

// Counts the tickets measure, for one student, right now.
function ticketCounts_(id, me) {
  return {
    bills: Number(me.bills) || 0, summits: Number(me.mindset) || 0, shoutouts: Number(me.shoutouts) || 0, star: Number(me.star) || 0,
    buddy: Number(me.ambassador) || 0, streak: Number(me.streak) || 0,
    polls: sheetRows_('Poll Votes').filter(function (v) { return id800_(v['800#']) === id; }).length,
    signups: sheetRows_('Sign-ups').filter(function (s) { return id800_(s['800#']) === id; }).length,
  };
}

// [{ key, need, have }] for "bills: 1 + polls: 3" against a baseline.
function ticketProgress_(req, now, base) {
  const kinds = ['bills', 'summits', 'shoutouts', 'star', 'buddy', 'polls', 'signups'];
  return String(req || '').split('+').map(function (p) {
    const m = p.trim().toLowerCase().match(/^(bills|summits|shoutouts|star|buddy|polls|signups|tools|streak)\s*:\s*(\d+)$/);
    if (!m) return null;
    const k = m[1], need = Number(m[2]);
    let have;
    if (k === 'streak') have = now.streak;
    else if (k === 'tools') have = kinds.filter(function (x) { return (now[x] || 0) > (base[x] || 0); }).length;
    else have = Math.max(0, (now[k] || 0) - (base[k] || 0));
    return { key: k, need: need, have: Math.min(have, need) };
  }).filter(Boolean);
}

function myTickets_(id, squad, me, write) {
  const defs = ticketDefs_(), now = ticketCounts_(id, me);
  const sheet = write ? getSheetSafe_('Student Tickets') : null;
  const today = todayYmd_();
  return sheetRows_('Student Tickets').filter(function (t) { return id800_(t['800#']) === id; }).map(function (t) {
    const def = defs.filter(function (d) { return d.ticket === String(t['Ticket']); })[0] || { ticket: String(t['Ticket']), text: '', req: '', points: 0 };
    let base = {}; try { base = JSON.parse(String(t['Baseline'] || '{}')); } catch (e) { base = {}; }
    const prog = ticketProgress_(def.req, now, base);
    let status = String(t['Status'] || 'Holding'), completed = String(t['Completed'] || '');
    if (status === 'Holding' && prog.length && prog.every(function (p) { return p.have >= p.need; })) {
      status = 'Completed'; completed = today;
      if (sheet) sheet.getRange(t._row, STUDENT_TICKET_HEADERS_.indexOf('Status') + 1, 1, 2).setValues([[status, completed]]);
    }
    return { row: t._row, ticket: def.ticket, text: def.text, points: def.points, progress: prog, status: status, drawn: String(t['Drawn'] || ''),
      completed: completed, awarded: String(t['Awarded'] || '') };
  });
}

// Completed tickets not yet paid: points to the squad (when GAME_AWARDS is ON), once.
function awardTickets_() {
  if (!awardsOn_()) return 0;
  const sheet = getSheetSafe_('Student Tickets');
  if (!sheet) return 0;
  const defs = ticketDefs_(), engine = enginePowers_(), board = boardRows_();
  let n = 0;
  sheetRows_('Student Tickets').forEach(function (t) {
    if (String(t['Status']) !== 'Completed' || String(t['Awarded'] || '')) return;
    const squad = squadOf_(t['Squad']);
    const def = defs.filter(function (d) { return d.ticket === String(t['Ticket']); })[0];
    const pts = (def ? def.points : 0) + engineBonus_(engine, squad, 'tickets');
    const who = board.filter(function (r) { return id800_(r['800#']) === id800_(t['800#']); })[0];
    const o = {}; o[squad] = pts;
    if (pts > 0) addPointsEvent_('Destination ticket: ' + String(who ? who['Name'] : 'a ' + titleSquad_(squad) + ' climber'), 'destination_ticket', String(t['Ticket']), o);
    sheet.getRange(t._row, STUDENT_TICKET_HEADERS_.indexOf('Awarded') + 1).setValue(todayYmd_() + ' +' + pts);
    n++;
  });
  return n;
}

function clashMeDrawTickets(asId, adminToken) {
  const id = meId_(asId, adminToken);
  const held = sheetRows_('Student Tickets').filter(function (t) { return id800_(t['800#']) === id && String(t['Status'] || 'Holding') === 'Holding'; });
  const max = Math.max(1, Number(getSetting_('TICKETS_HELD')) || 3);
  if (held.length >= max) throw new Error('You\'re holding ' + held.length + ' tickets. Finish one to draw more.');
  const have = held.map(function (t) { return String(t['Ticket']); });
  const pool = ticketDefs_().filter(function (d) { return have.indexOf(d.ticket) === -1; });
  for (let i = pool.length - 1; i > 0; i--) { const j = Math.floor(Math.random() * (i + 1)); const x = pool[i]; pool[i] = pool[j]; pool[j] = x; }
  const offer = pool.slice(0, Math.max(1, Number(getSetting_('TICKET_CHOICES')) || 3));
  CacheService.getScriptCache().put('tickets-' + id, JSON.stringify(offer.map(function (d) { return d.ticket; })), 900);
  return JSON.stringify({ offer: offer, keepUpTo: max - held.length });
}

function clashMeKeepTickets(names, asId, adminToken) {
  const id = meId_(asId, adminToken);
  const offered = JSON.parse(CacheService.getScriptCache().get('tickets-' + id) || '[]');
  const keep = (names || []).filter(function (n) { return offered.indexOf(n) > -1; });
  if (!keep.length) throw new Error('Keep at least one ticket.');
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const row = boardRows_().filter(function (r) { return id800_(r['800#']) === id; })[0];
    if (!row) throw new Error('You\'re not on the Student Board yet.');
    const me = { bills: row['Trait Dough Bills'], mindset: row['Mindset Summits'], shoutouts: row['Shout-outs'], star: row['STAR Growth Points'],
      ambassador: row['Ambassador Buddies'], streak: row['Streak (weeks)'] };
    const held = sheetRows_('Student Tickets').filter(function (t) { return id800_(t['800#']) === id && String(t['Status'] || 'Holding') === 'Holding'; }).length;
    const max = Math.max(1, Number(getSetting_('TICKETS_HELD')) || 3);
    const base = JSON.stringify(ticketCounts_(id, me));
    const sheet = getSheet_('Student Tickets');
    keep.slice(0, Math.max(0, max - held)).forEach(function (n) { sheet.appendRow([id, squadOf_(row['Squad']), n, todayYmd_(), base, 'Holding', '', '']); });
    CacheService.getScriptCache().remove('tickets-' + id);
  } finally { lock.releaseLock(); }
  return clashMe(asId, adminToken);
}

// ── MY MASCOT ───────────────────────────────────────────────────────
// Gear rules: "boots=bills>=1; pack=bills>=5; ..." checked against the student's numbers.
function gearFor_(me, extra) {
  const rules = String(getSetting_('AVATAR_GEAR') || '').split(';').map(function (p) {
    const m = p.trim().match(/^([a-z]+)\s*=\s*([a-z]+)\s*>=\s*(\d+)$/i); return m ? { gear: m[1].toLowerCase(), key: m[2].toLowerCase(), min: Number(m[3]) } : null;
  }).filter(Boolean);
  const vals = Object.assign({ bills: me.bills, mindset: me.mindset, shoutouts: me.shoutouts, star: me.star, streak: me.streak, buddy: me.ambassador }, extra || {});
  return rules.map(function (r) { return { gear: r.gear, have: (Number(vals[r.key]) || 0) >= r.min, need: r.key + ' ' + r.min }; });
}

function lookFor_(id) {
  const r = sheetRows_('Avatars').filter(function (a) { return id800_(a['800#']) === id; })[0];
  try { return r ? JSON.parse(String(r['Look'] || '{}')) : null; } catch (e) { return null; }
}

function clashMeSaveLook(look, asId, adminToken) {
  const id = meId_(asId, adminToken);
  const clean = {};
  Object.keys(LOOK_CHOICES_).forEach(function (k) { const v = Math.floor(Number(look && look[k])); clean[k] = v >= 0 && v < LOOK_CHOICES_[k] ? v : 0; });
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const sheet = getSheet_('Avatars');
    const r = sheetRows_('Avatars').filter(function (a) { return id800_(a['800#']) === id; })[0];
    const row = [id, JSON.stringify(clean), todayYmd_()];
    if (r) sheet.getRange(r._row, 1, 1, 3).setValues([row]); else sheet.appendRow(row);
  } finally { lock.releaseLock(); }
  return JSON.stringify({ look: clean });
}

// ── THE DASHBOARD ───────────────────────────────────────────────────
function clashMe(asId, adminToken) {
  const id = meId_(asId, adminToken);
  const page = JSON.parse(clashMyPage(id)); // everything My Squad already shows (04_Students.gs)
  const rows = boardRows_();
  const row = rows.filter(function (r) { return id800_(r['800#']) === id; })[0] || {};
  const squad = page.me.squad;
  page.me.id = id;
  page.me.attendance = String(row['Attendance %'] || '').trim() === '' ? null : Number(row['Attendance %']);
  page.metrics = metricsFor_(id, squad, rows);
  page.tickets = myTickets_(id, squad, page.me, true);
  awardTickets_();
  page.ticketsHeld = page.tickets.filter(function (t) { return t.status === 'Holding'; }).length;
  page.ticketsMax = Math.max(1, Number(getSetting_('TICKETS_HELD')) || 3);
  page.look = lookFor_(id);
  page.gear = gearFor_(page.me, { tickets: page.tickets.filter(function (t) { return t.status === 'Completed'; }).length, passport: page.passport && page.passport.complete ? 1 : 0 });
  page.mascot = mascotFor_(squad);
  page.game = gameState_();
  page.preview = !!(asId && adminToken);
  return JSON.stringify(page);
}

function mascotFor_(squad) {
  const t = sheetRows_('Squad Themes').filter(function (r) { return squadOf_(r['Squad']) === squad; })[0] || {};
  const image = String(t['Mascot Image'] || '').trim();
  return { squad: squad, animal: String(t['Mascot'] || '').trim() || MASCOTS_[squad] || '', name: String(t['Mascot Name'] || '').trim(),
    image: /^https:\/\//i.test(image) ? image : '' };
}

function clashMascots() {
  const out = {};
  SQUAD_KEYS_.forEach(function (s) { out[s] = mascotFor_(s); });
  return JSON.stringify(out);
}

// The student page itself: ?page=me on the student link.
function serveMe_() {
  const t = HtmlService.createTemplateFromFile('Me');
  const shield = String(getSetting_('SHIELD_LOGO_URL') || '').trim();
  t.shield = /^https:\/\//i.test(shield) ? shield : '';
  return t.evaluate()
    .setTitle('My Climb · Clash of Classes')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
