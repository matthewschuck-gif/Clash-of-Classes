/**
 * Clash of Classes -- Apps Script backend
 * File 22: the week. Everything that runs on its own and tells people what happened.
 *
 * THIS WEEK IN CLASH: what happened Monday to today (standings, the week's points, highlights,
 *   map captures, creatures evolved, the Opponent, polls, duels and goals) and what's up next
 *   week (events, goals and duels ending, the Opponent, days off, the marking period). Shown on
 *   the site's This Week page, and emailed as an HTML digest every week to DIGEST_EMAILS
 *   (Clash of Classes › Turn on the weekly digest email; DIGEST_DAY at DIGEST_HOUR).
 * MORNING ANNOUNCEMENT: a ready-to-read script (yesterday's points, map captures, the Opponent,
 *   Hot Streaks, creatures evolved, the Daily Poll). Menu, admin panel, and the digest.
 * MARKING PERIODS (Settings › MARKING_PERIODS "MP1=2026-08-26..2026-11-06; MP2=…"): squad meeting
 *   slides show each squad's marking period, and BEAT_MP_BONUS points go to every squad that
 *   out-earns its own last marking period (once per marking period, with the game awards).
 * HEALTH (Admin › System): when each automatic job last ran, the last error, triggers, how
 *   fresh the Student Board and metrics are, settings that still need filling, and which pages
 *   people open (USAGE: anonymous counts per page per day, no names, on the Usage tab).
 */

const USAGE_HEADERS_ = ['Date', 'Page', 'Opens'];
const DIGEST_LOG_HEADERS_ = ['Sent', 'Week Of', 'Recipients', 'Subject'];
const USAGE_PAGES_ = ['pulse', 'homerooms', 'year', 'home', 'mysquad', 'trait', 'squads', 'tracker', 'mindset', 'wheel', 'traitdough', 'sotm', 'kickoff', 'mayhem', 'champs', 'gallery',
  'staff', 'spin', 'growth', 'leaders', 'week', 'generals', 'hq', 'me-me', 'me-tickets', 'me-squad', 'me-map', 'me-game'];

function ensureWeekTabs_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  [['Usage', USAGE_HEADERS_], ['Weekly Digest Log', DIGEST_LOG_HEADERS_]].forEach(function (t) {
    if (ss.getSheetByName(t[0])) return;
    const sh = ss.insertSheet(t[0]);
    sh.getRange(1, 1, 1, t[1].length).setValues([t[1]]).setFontWeight('bold').setBackground('#490e6f').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  });
}

// ── RUN LOG (for the health panel) ──────────────────────────────────
function noteRun_(job) { try { PropertiesService.getScriptProperties().setProperty('ran-' + job, new Date().toISOString()); } catch (e) { /* never block a job */ } }
function noteError_(job, err) {
  try { PropertiesService.getScriptProperties().setProperty('last-error', JSON.stringify({ job: job, message: String(err && err.message || err), at: new Date().toISOString() })); } catch (e) { /* ignore */ }
}

// ── MARKING PERIODS ─────────────────────────────────────────────────
function markingPeriods_() {
  return String(getSetting_('MARKING_PERIODS') || '').split(';').map(function (p) {
    const m = p.trim().match(/^([^=]+)=\s*(\d{4}-\d{2}-\d{2})\s*\.\.\s*(\d{4}-\d{2}-\d{2})$/);
    return m ? { name: m[1].trim(), start: m[2], end: m[3] } : null;
  }).filter(Boolean).sort(function (a, b) { return a.start.localeCompare(b.start); });
}

// { now, prev } for a date (today by default); either can be null.
function mpFor_(ymd) {
  ymd = ymd || todayYmd_();
  const list = markingPeriods_();
  let i = list.findIndex(function (p) { return ymd >= p.start && ymd <= p.end; });
  if (i < 0) i = list.filter(function (p) { return p.end < ymd; }).length - 1; // between periods: the last one that ended
  return { now: i >= 0 ? list[i] : null, prev: i > 0 ? list[i - 1] : null, next: list.filter(function (p) { return p.start > ymd; })[0] || null };
}

// Points a squad earned between two dates (game bonuses and private rows left out).
function squadPointsBetween_(events, from, to) { return pointsBy_(events, from, to, null).points; }

// Once a marking period has ended: BEAT_MP_BONUS to every squad that out-earned its previous one.
function payBeatMp_(events) {
  const bonus = Number(getSetting_('BEAT_MP_BONUS')) || 0;
  if (bonus <= 0) return [];
  const list = markingPeriods_(), today = todayYmd_(), props = PropertiesService.getScriptProperties(), lines = [];
  list.forEach(function (mp, i) {
    if (i === 0 || mp.end >= today || props.getProperty('beat-mp-' + mp.name)) return;
    const now = squadPointsBetween_(events, mp.start, mp.end), before = squadPointsBetween_(events, list[i - 1].start, list[i - 1].end);
    const o = {};
    SQUAD_KEYS_.forEach(function (s) { if (now[s] > before[s] && now[s] > 0) o[s] = bonus; });
    if (Object.keys(o).length) addPointsEvent_('Beat our own marking period: ' + mp.name, 'mp_bonus', Object.keys(o).map(titleSquad_).join(', ') + ' out-earned ' + list[i - 1].name, o);
    props.setProperty('beat-mp-' + mp.name, today);
    lines.push(mp.name + ': ' + (Object.keys(o).length ? Object.keys(o).join(', ') + ' +' + bonus : 'nobody beat their last marking period'));
  });
  return lines;
}

// What the squad meeting slides add for one squad.
function meetingData_(squad) {
  const events = readEventsTab_(), today = todayYmd_(), mp = mpFor_(today);
  const out = { mp: null, zones: [], wings: [], evolved: 0, summit: 0, students: 0, metrics: null, firsts: [], fresh: null, doubles: [], pulse: '' };
  // January (or the 3 weeks after Settings › FRESH_START): the fresh-start and freeze-and-thaw slides.
  const fs = String(getSetting_('FRESH_START') || '').trim().slice(0, 10), hasFs = /^\d{4}-\d{2}-\d{2}$/.test(fs);
  if ((hasFs && today >= ymdShift_(fs, -14) && today <= ymdShift_(fs, 21)) || today.slice(5, 7) === '01') out.fresh = { date: hasFs ? fs : '' };
  if (mp.now) {
    const end = mp.now.end < today ? mp.now.end : today;
    const now = squadPointsBetween_(events, mp.now.start, end);
    const before = mp.prev ? squadPointsBetween_(events, mp.prev.start, mp.prev.end) : null;
    out.mp = { name: mp.now.name, start: mp.now.start, end: mp.now.end, points: now[squad], prevName: mp.prev ? mp.prev.name : '', prevPoints: before ? before[squad] : null,
      bonus: Number(getSetting_('BEAT_MP_BONUS')) || 0 };
  }
  try {
    const zones = mapZones_(events);
    out.zones = zones.filter(function (z) { return z.holder === squad; }).map(function (z) { return z.zone; });
    out.doubles = zones.filter(function (z) { return z.mult > 1; }).map(function (z) { return { zone: z.zone, rule: z.rule, holder: z.holder, mult: z.mult }; });
    const w = heldWings_(zones); out.wings = Object.keys(w).filter(function (k) { return w[k] === squad; });
  } catch (e) { /* the map is optional */ }
  boardRows_().filter(function (r) { return squadOf_(r['Squad']) === squad; }).forEach(function (r) {
    out.students++;
    const st = stageFor_(climbXp_({ bills: boardNumber_(r['Trait Dough Bills']), mindset: boardNumber_(r['Mindset Summits']), star: boardNumber_(r['STAR Growth Points']),
      shoutouts: boardNumber_(r['Shout-outs']), buddy: boardNumber_(r['Ambassador Buddies']), streak: boardNumber_(r['Streak (weeks)']), journal: boardNumber_(r['Trail Journal Growth']) }));
    if (st >= 2) out.evolved++;
    if (st >= 4) out.summit++;
  });
  try {
    const m = readMetricsTab_();
    if (m) out.metrics = { attendance: m.attendance[squad], referral: m.referral[squad], grades: m.grades[squad],
      rank: ['attendance', 'referral', 'grades'].reduce(function (o, k) { o[k] = 1 + SQUAD_KEYS_.filter(function (s) { return (m[k][s] || 0) > (m[k][squad] || 0); }).length; return o; }, {}) };
  } catch (e) { /* no metrics yet */ }
  try { const p = JSON.parse(clashPulseLive()).live[0]; out.pulse = p ? p.question : ''; } catch (e) { /* no Pulse open */ }
  try { out.firsts = leadersData_().boards.filter(function (b) { return SQUAD_KEYS_.every(function (s) { return b.scores[squad] >= b.scores[s]; }) && b.scores[squad] > 0; }).map(function (b) { return b.icon + ' ' + b.label; }); } catch (e) { /* ignore */ }
  return out;
}

// ── CREATURE SNAPSHOTS (how many evolved, day by day) ───────────────
function stageCounts_() {
  const c = { BLACK: [0, 0, 0, 0, 0], GOLD: [0, 0, 0, 0, 0], GREY: [0, 0, 0, 0, 0], PURPLE: [0, 0, 0, 0, 0] };
  boardRows_().forEach(function (r) {
    const s = squadOf_(r['Squad']);
    if (!c[s]) return;
    c[s][stageFor_(climbXp_({ bills: boardNumber_(r['Trait Dough Bills']), mindset: boardNumber_(r['Mindset Summits']), star: boardNumber_(r['STAR Growth Points']),
      shoutouts: boardNumber_(r['Shout-outs']), buddy: boardNumber_(r['Ambassador Buddies']), streak: boardNumber_(r['Streak (weeks)']), journal: boardNumber_(r['Trail Journal Growth']) }))]++;
  });
  return c;
}
// Evolutions = steps up a stage, summed (a student going from 1 to 3 counts twice).
function evolutionScore_(c) { const o = {}; SQUAD_KEYS_.forEach(function (s) { o[s] = c[s][2] + 2 * c[s][3] + 3 * c[s][4]; }); return o; }

function snapshotStages_() {
  const props = PropertiesService.getScriptProperties(), today = todayYmd_();
  let log = {}; try { log = JSON.parse(props.getProperty('stage-log') || '{}'); } catch (e) { log = {}; }
  if (log[today]) return log;
  log[today] = evolutionScore_(stageCounts_());
  Object.keys(log).sort().slice(0, -21).forEach(function (d) { delete log[d]; }); // keep three weeks
  props.setProperty('stage-log', JSON.stringify(log));
  return log;
}
// Evolutions per squad since a date, from the snapshots.
function evolvedSince_(ymd) {
  let log = {}; try { log = JSON.parse(PropertiesService.getScriptProperties().getProperty('stage-log') || '{}'); } catch (e) { log = {}; }
  const days = Object.keys(log).sort(), base = days.filter(function (d) { return d <= ymd; }).pop() || days[0];
  const now = evolutionScore_(stageCounts_()), out = { BLACK: 0, GOLD: 0, GREY: 0, PURPLE: 0 };
  if (!base) return out;
  SQUAD_KEYS_.forEach(function (s) { out[s] = Math.max(0, now[s] - (log[base][s] || 0)); });
  return out;
}

// ── THE MORNING ANNOUNCEMENT ────────────────────────────────────────
function lastSchoolDay_(ymd) {
  let off = []; try { off = schoolCalendar_().noSchool; } catch (e) { off = []; }
  let d = ymdShift_(ymd || todayYmd_(), -1);
  for (let i = 0; i < 10; i++) {
    const dow = new Date(d + 'T12:00:00').getDay();
    if (dow !== 0 && dow !== 6 && off.indexOf(d) === -1) return d;
    d = ymdShift_(d, -1);
  }
  return d;
}

function announcementScript_() {
  const events = readEventsTab_(), y = lastSchoolDay_(), pts = squadPointsBetween_(events, y, y), parts = [];
  const order = SQUAD_KEYS_.slice().sort(function (a, b) { return pts[b] - pts[a]; });
  if (pts[order[0]] > 0) parts.push('Good morning, Mountaineers! Yesterday ' + titleSquad_(order[0]) + ' Squad earned the most points with ' + pts[order[0]] + '. ' +
    order.slice(1).filter(function (s) { return pts[s] > 0; }).map(function (s) { return titleSquad_(s) + ' earned ' + pts[s]; }).join(', ') + (order.slice(1).some(function (s) { return pts[s] > 0; }) ? '.' : ''));
  else parts.push('Good morning, Mountaineers!');
  const totals = squadTotals_(), lead = SQUAD_KEYS_.slice().sort(function (a, b) { return totals[b] - totals[a]; });
  parts.push(titleSquad_(lead[0]) + ' leads Clash of Classes, ' + (totals[lead[0]] - totals[lead[1]]).toLocaleString() + ' points ahead of ' + titleSquad_(lead[1]) + '.');
  const caps = sheetRows_('Map Battles').filter(function (b) { return String(b['Date'] || '').slice(0, 10) === y; });
  if (caps.length) parts.push('On the map: ' + caps.map(function (b) { return titleSquad_(squadOf_(b['Taken By'])) + ' took the ' + b['Zone']; }).join(', ') + '.');
  try {
    const m = monster_(events);
    if (m) parts.push(m.defeated ? 'We defeated ' + m.name + '! Every squad earns ' + m.reward + ' points.' : m.name + ' is ' + Math.round(100 * (1 - m.health)) + '% defeated. Keep it up to beat it together.');
  } catch (e) { /* no Opponent */ }
  try {
    const st = hotStreaks_(), hot = SQUAD_KEYS_.filter(function (s) { return st[s] && st[s].length >= 2; });
    hot.forEach(function (s) { parts.push(titleSquad_(s) + ' is on a ' + st[s].length + '-day Hot Streak in the Daily Poll!'); });
  } catch (e) { /* no streaks */ }
  const ev = evolvedSince_(lastSchoolDay_(y)), evN = SQUAD_KEYS_.reduce(function (n, s) { return n + ev[s]; }, 0);
  if (evN) parts.push(evN + ' creature' + (evN === 1 ? '' : 's') + ' evolved on My Climb. Keep earning LRG points!');
  try { const poll = pollAnnouncement_(); if (poll && !/^No daily poll/.test(poll)) parts.push(poll); } catch (e) { /* no polls */ }
  return parts.join(' ');
}

function clashAnnouncement(token) { requireAdmin_(token); return JSON.stringify({ text: announcementScript_() }); }
function menuAnnouncement_() { SpreadsheetApp.getUi().alert('Read this in the morning announcements', announcementScript_(), SpreadsheetApp.getUi().ButtonSet.OK); }

// ── USAGE (anonymous page opens) ────────────────────────────────────
function clashTrack(page) {
  page = String(page || '').toLowerCase();
  if (USAGE_PAGES_.indexOf(page) === -1) return '';
  const cache = CacheService.getScriptCache(), key = 'use-' + todayYmd_() + '-' + page;
  cache.put(key, String((Number(cache.get(key)) || 0) + 1), 21600);
  return '';
}

// Writes today's counts from the cache to the Usage tab (from backgroundTick).
function flushUsage_() {
  const sheet = getSheetSafe_('Usage');
  if (!sheet) return;
  const today = todayYmd_(), cache = CacheService.getScriptCache();
  const got = cache.getAll ? cache.getAll(USAGE_PAGES_.map(function (p) { return 'use-' + today + '-' + p; })) : {};
  const rows = sheetRows_('Usage');
  USAGE_PAGES_.forEach(function (p) {
    const n = Number(got['use-' + today + '-' + p]) || 0;
    if (!n) return;
    const r = rows.filter(function (x) { return String(x['Date']).slice(0, 10) === today && String(x['Page']) === p; })[0];
    if (r) sheet.getRange(r._row, 3).setValue(Math.max(n, Number(r['Opens']) || 0)); else sheet.appendRow([today, p, n]);
  });
}

function usageSince_(ymd) {
  const out = {};
  sheetRows_('Usage').forEach(function (r) { if (String(r['Date']).slice(0, 10) >= ymd) out[r['Page']] = (out[r['Page']] || 0) + (Number(r['Opens']) || 0); });
  return Object.keys(out).map(function (p) { return { page: p, opens: out[p] }; }).sort(function (a, b) { return b.opens - a.opens; });
}

// ── HEALTH (Admin › System) ─────────────────────────────────────────
function clashHealth(token) {
  requireAdmin_(token);
  const props = PropertiesService.getScriptProperties(), now = Date.now();
  const ago = function (iso) { if (!iso) return null; const m = Math.round((now - Date.parse(iso)) / 60000); return m < 90 ? m + ' min ago' : m < 2880 ? Math.round(m / 60) + ' hours ago' : Math.round(m / 1440) + ' days ago'; };
  const stale = function (iso, minutes) { return !iso || now - Date.parse(iso) > minutes * 60000; };
  let triggers = [];
  try { triggers = ScriptApp.getProjectTriggers().map(function (t) { return t.getHandlerFunction(); }); } catch (e) { triggers = []; }
  const jobs = [
    ['backgroundTick', 'Background check (map, freezes, fast loading)', 15, true],
    ['gameTick', 'Game awards (territory, goals, the Opponent, boosts)', 90, awardsOn_()],
    ['weeklyDigest', 'Weekly digest email', 60 * 24 * 8, String(getSetting_('DIGEST') || 'OFF').toUpperCase() === 'ON'],
  ].map(function (j) {
    const last = props.getProperty('ran-' + j[0]);
    return { job: j[0], label: j[1], last: last, ago: ago(last), trigger: triggers.indexOf(j[0]) > -1, expected: j[3], ok: !j[3] || (triggers.indexOf(j[0]) > -1 && !stale(last, j[2])) };
  });
  let boardUpdated = '';
  boardRows_().forEach(function (r) { const u = String(r['Updated'] || ''); if (u > boardUpdated) boardUpdated = u; });
  let squadBoard = '';
  try { const sb = getSheetSafe_('Squad Board'); if (sb) squadBoard = String(sb.getRange(1, 1).getValues()[0][0] || ''); } catch (e) { /* none */ }
  const missing = [['STUDENT_LINK', 'My Climb link'], ['GENERALS_CODE', 'Generals HQ code'], ['MARKING_PERIODS', 'Marking period dates'], ['DIGEST_EMAILS', 'Weekly digest recipients']]
    .filter(function (k) { return !String(getSetting_(k[0]) || '').trim(); }).map(function (k) { return k[1] + ' (Settings › ' + k[0] + ')'; });
  let quota = null; try { quota = MailApp.getRemainingDailyQuota(); } catch (e) { quota = null; }
  let err = null; try { err = JSON.parse(props.getProperty('last-error') || 'null'); } catch (e) { err = null; }
  return JSON.stringify({
    jobs: jobs, lastError: err ? Object.assign(err, { ago: ago(err.at) }) : null,
    studentBoard: { updated: boardUpdated, ok: !!boardUpdated }, squadBoard: squadBoard,
    metricsNote: 'Attendance, no-referral and no-F numbers fill in from Define a Climb-er\'s newest PowerSchool upload (its SQUAD_METRICS setting). Finalize in Admin › Metrics to award them.',
    missing: missing, gameAwards: awardsOn_(), emailQuota: quota, usage: usageSince_(ymdShift_(todayYmd_(), -7)),
  });
}

// ── THIS WEEK IN CLASH ──────────────────────────────────────────────
function weekData_(asOf) {
  asOf = asOf || todayYmd_();
  const events = readEventsTab_(), mon = weekStart_(asOf), fri = ymdShift_(mon, 4), nextMon = ymdShift_(mon, 7), nextFri = ymdShift_(mon, 11);
  const week = squadPointsBetween_(events, mon, asOf), totals = squadTotals_();
  const place = SQUAD_KEYS_.slice().sort(function (a, b) { return totals[b] - totals[a]; });
  const highlights = events.filter(function (e) { const d = String(e.date || '').slice(0, 10); return d >= mon && d <= asOf && GAME_TYPES_.indexOf(String(e.type)) === -1; })
    .map(function (e) { const p = publicEvent_(e); const sum = SQUAD_KEYS_.reduce(function (n, s) { return n + Math.max(0, Number(e.pts[s]) || 0); }, 0); return { name: p.name, date: String(e.date).slice(0, 10), total: sum, pts: e.pts }; })
    .filter(function (e) { return e.total > 0; }).sort(function (a, b) { return b.total - a.total; }).slice(0, 6);
  const caps = sheetRows_('Map Battles').filter(function (b) { const d = String(b['Date'] || '').slice(0, 10); return d >= mon && d <= asOf; })
    .map(function (b) { return { zone: String(b['Zone']), to: squadOf_(b['Taken By']), from: squadOf_(b['From']), date: String(b['Date']).slice(0, 10) }; });
  let king = '', zonesHeld = { BLACK: 0, GOLD: 0, GREY: 0, PURPLE: 0 };
  try { mapZones_(events).forEach(function (z) { if (z.holder) zonesHeld[z.holder]++; if (z.summit && z.holder) king = z.holder; }); } catch (e) { /* map optional */ }
  const evolved = evolvedSince_(ymdShift_(mon, -1));
  const boosts = sheetRows_('Boost Ledger').filter(function (b) { const d = String(b['Picked'] || '').slice(0, 10); return d >= mon && d <= asOf; }).length;
  let opponent = null; try { const m = monster_(events); if (m) opponent = { name: m.name, defeated: m.defeated, pct: Math.round(100 * (1 - m.health)), goal: m.goal, end: m.end, reward: m.reward }; } catch (e) { /* none */ }
  const votes = sheetRows_('Poll Votes').filter(function (v) { const d = String(v['Date'] || '').slice(0, 10); return d >= mon && d <= asOf; }).length;
  let streak = null; try { const st = hotStreaks_(); const top = SQUAD_KEYS_.filter(function (s) { return st[s] && st[s].length >= 2; }).sort(function (a, b) { return st[b].length - st[a].length; })[0]; if (top) streak = { squad: top, days: st[top].length }; } catch (e) { /* none */ }
  let firsts = []; try { firsts = leadersData_().boards.filter(function (b) { return b.key !== 'month'; }).map(function (b) { const lead = SQUAD_KEYS_.slice().sort(function (a, c) { return b.scores[c] - b.scores[a]; })[0]; return { label: b.icon + ' ' + b.label, squad: lead, score: b.scores[lead], unit: b.unit || '' }; }).filter(function (f) { return f.score > 0; }); } catch (e) { /* none */ }
  // Up next: things dated next week.
  const inNext = function (d) { d = String(d || '').slice(0, 10); return d >= nextMon && d <= nextFri; };
  const next = [];
  sheetRows_('Sign-up Events').forEach(function (r) { if (inNext(r['Date'])) next.push({ date: String(r['Date']).slice(0, 10), what: String(r['Event']), kind: 'Event' }); });
  sheetRows_('Squad Duels').forEach(function (r) {
    if (!ticked_(r['Show'])) return;
    if (inNext(r['Start'])) next.push({ date: String(r['Start']).slice(0, 10), what: 'Squad Duel starts: ' + r['Title'], kind: 'Duel' });
    if (inNext(r['End'])) next.push({ date: String(r['End']).slice(0, 10), what: 'Squad Duel ends: ' + r['Title'], kind: 'Duel' });
  });
  sheetRows_('Challenges').forEach(function (r) { if (ticked_(r['Show']) && inNext(r['End'])) next.push({ date: String(r['End']).slice(0, 10), what: 'Challenge ends: ' + r['Title'], kind: 'Challenge' }); });
  sheetRows_('Round Goals').forEach(function (r) { if (ticked_(r['Show']) && inNext(r['End'])) next.push({ date: String(r['End']).slice(0, 10), what: 'Round goal ends: ' + r['Goal'], kind: 'Goal' }); });
  sheetRows_('Opponents').forEach(function (r) {
    if (!ticked_(r['Show'])) return;
    if (inNext(r['Start'])) next.push({ date: String(r['Start']).slice(0, 10), what: 'New Opponent: ' + r['Opponent'], kind: 'Opponent' });
    if (inNext(r['End'])) next.push({ date: String(r['End']).slice(0, 10), what: 'Last chance to beat ' + r['Opponent'], kind: 'Opponent' });
  });
  const polls = sheetRows_('Daily Polls').filter(function (r) { return ticked_(r['Show']) && inNext(r['Date']); }).length;
  if (polls) next.push({ date: nextMon, what: polls + ' Daily Poll' + (polls === 1 ? '' : 's') + ' scheduled', kind: 'Poll' });
  try { schoolCalendar_().noSchool.filter(inNext).forEach(function (d) { next.push({ date: d, what: 'No school', kind: 'Day off' }); }); } catch (e) { /* none */ }
  try { yearPlanItems_().forEach(function (it) { if (it.date && inNext(it.date)) next.push({ date: it.date, what: it.title + (it.who && it.who !== 'All' ? ' (' + it.who + ')' : ''), kind: it.kind || 'Year Plan' }); }); } catch (e) { /* 24_YearPlan.gs */ }
  markingPeriods_().forEach(function (p) { if (inNext(p.end)) next.push({ date: p.end, what: p.name + ' ends', kind: 'Marking period' }); if (inNext(p.start)) next.push({ date: p.start, what: p.name + ' starts', kind: 'Marking period' }); });
  next.sort(function (a, b) { return a.date.localeCompare(b.date); });
  return { weekOf: mon, through: asOf < fri ? asOf : fri, nextWeekOf: nextMon, totals: totals, place: place, week: week, highlights: highlights, captures: caps,
    zonesHeld: zonesHeld, king: king, evolved: evolved, boosts: boosts, opponent: opponent, votes: votes, streak: streak, firsts: firsts, next: next,
    site: String(getSetting_('SITE_URL') || 'https://clashofclasses.org') };
}

function clashWeek() {
  const key = 'week-' + dataStamp_();
  const hit = cacheGetBig_(key);
  if (hit) return hit;
  const text = JSON.stringify(weekData_());
  cachePutBig_(key, text, 300);
  return text;
}

// The digest as email-safe HTML: tables and inline styles only.
function digestHtml_(d) {
  const C = { BLACK: '#3f4652', GOLD: '#b8860b', GREY: '#5f6f86', PURPLE: '#6d28d9' }, T = function (s) { return s ? s.charAt(0) + s.slice(1).toLowerCase() : ''; };
  const h = function (s) { return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); };
  const day = function (ymd) { const d = new Date(ymd + 'T12:00:00'); return ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'][d.getDay()] + ' ' + (d.getMonth() + 1) + '/' + d.getDate(); };
  const sec = function (title, body) { return '<tr><td style="padding:18px 24px 4px"><div style="font:700 17px Georgia,serif;color:#2a0f45;border-bottom:2px solid #f2e3ff;padding-bottom:6px">' + title + '</div>' + body + '</td></tr>'; };
  const p = function (t) { return '<p style="margin:8px 0;font:14px/1.5 Arial,sans-serif;color:#2b2b3a">' + t + '</p>'; };
  const standings = '<table width="100%" cellpadding="6" style="border-collapse:collapse;font:14px Arial,sans-serif;margin-top:8px">' +
    '<tr style="color:#6b5a85;font-size:12px"><td>Place</td><td>Squad</td><td align="right">Total</td><td align="right">This week</td></tr>' +
    d.place.map(function (s, i) { return '<tr style="border-top:1px solid #eee"><td>' + (i + 1) + '</td><td style="font-weight:700;color:' + C[s] + '">' + T(s) + '</td><td align="right">' + d.totals[s].toLocaleString() + '</td><td align="right">+' + d.week[s].toLocaleString() + '</td></tr>'; }).join('') + '</table>';
  const hl = d.highlights.length ? '<ul style="margin:8px 0;padding-left:18px;font:14px/1.6 Arial,sans-serif;color:#2b2b3a">' + d.highlights.map(function (e) {
    return '<li>' + h(e.name) + ' <span style="color:#888">(' + day(e.date) + ')</span>: ' + SQUAD_KEYS_.filter(function (s) { return (Number(e.pts[s]) || 0) > 0; }).map(function (s) { return '<span style="color:' + C[s] + '">' + T(s) + ' +' + e.pts[s] + '</span>'; }).join(', ') + '</li>';
  }).join('') + '</ul>' : p('A quiet week for points.');
  const ev = SQUAD_KEYS_.reduce(function (n, s) { return n + d.evolved[s]; }, 0);
  const climb = p((ev ? '<b>' + ev + '</b> creature evolution' + (ev === 1 ? '' : 's') + ' on My Climb (' + SQUAD_KEYS_.filter(function (s) { return d.evolved[s]; }).map(function (s) { return T(s) + ' ' + d.evolved[s]; }).join(', ') + ')' : 'No new evolutions yet this week') +
    (d.boosts ? ', and <b>' + d.boosts + '</b> new evolution boost' + (d.boosts === 1 ? '' : 's') + ' picked.' : '.'));
  const map = p((d.captures.length ? d.captures.map(function (c) { return '<span style="color:' + C[c.to] + ';font-weight:700">' + T(c.to) + '</span> took the ' + h(c.zone) + (c.from ? ' from ' + T(c.from) : ''); }).join('; ') + '. ' : 'No zones changed hands. ') +
    (d.king ? 'King of the Mountain: <b style="color:' + C[d.king] + '">' + T(d.king) + '</b>. ' : '') + 'Zones held: ' + SQUAD_KEYS_.map(function (s) { return T(s) + ' ' + d.zonesHeld[s]; }).join(', ') + '.');
  const opp = d.opponent ? p(d.opponent.defeated ? 'We beat <b>' + h(d.opponent.name) + '</b>! Every squad earns ' + d.opponent.reward + '.' : '<b>' + h(d.opponent.name) + '</b> is ' + d.opponent.pct + '% defeated (goal: ' + h(d.opponent.goal) + ', until ' + day(d.opponent.end) + ').') : '';
  const polls = p('<b>' + d.votes + '</b> Daily Poll votes this week.' + (d.streak ? ' <span style="color:' + C[d.streak.squad] + ';font-weight:700">' + T(d.streak.squad) + '</span> is on a ' + d.streak.days + '-day Hot Streak.' : ''));
  const firsts = d.firsts.length ? '<table width="100%" cellpadding="4" style="border-collapse:collapse;font:13px Arial,sans-serif;margin-top:6px">' + d.firsts.map(function (f) {
    return '<tr style="border-top:1px solid #f0f0f0"><td>' + h(f.label) + '</td><td align="right" style="font-weight:700;color:' + C[f.squad] + '">' + T(f.squad) + ' · ' + f.score.toLocaleString() + f.unit + '</td></tr>'; }).join('') + '</table>' : '';
  const next = d.next.length ? '<table width="100%" cellpadding="5" style="border-collapse:collapse;font:14px Arial,sans-serif;margin-top:6px">' + d.next.map(function (n) {
    return '<tr style="border-top:1px solid #eee"><td style="color:#6b21a8;white-space:nowrap;width:70px">' + day(n.date) + '</td><td>' + h(n.what) + '</td><td align="right" style="color:#999;font-size:12px">' + h(n.kind) + '</td></tr>'; }).join('') + '</table>' : p('Nothing dated on the calendar yet. Add events on the Clash sheet and they\'ll show here.');
  return '<!doctype html><html><body style="margin:0;background:#f4f0fa"><table width="100%" cellpadding="0" style="background:#f4f0fa;padding:20px 0"><tr><td align="center">' +
    '<table width="620" cellpadding="0" style="max-width:620px;width:100%;background:#ffffff;border-radius:14px;overflow:hidden;border:1px solid #e6dcf5">' +
    '<tr><td style="background:#2a0f45;padding:22px 24px"><div style="font:700 12px Arial,sans-serif;letter-spacing:2px;color:#ffe100;text-transform:uppercase">This Week in Clash</div>' +
    '<div style="font:700 24px Georgia,serif;color:#ffffff;margin-top:4px">Week of ' + day(d.weekOf) + '</div>' +
    '<div style="font:13px Arial,sans-serif;color:#d6c2ff;margin-top:4px">' + T(d.place[0]) + ' leads by ' + (d.totals[d.place[0]] - d.totals[d.place[1]]).toLocaleString() + ' points.</div></td></tr>' +
    sec('Standings', standings) + sec('Highlights', hl) + sec('Conquer the school', map) + sec('My Climb', climb) + (opp ? sec('The Opponent', opp) : '') + sec('Daily Polls', polls) +
    (firsts ? sec('Who leads what', firsts) : '') + sec('Up next: week of ' + day(d.nextWeekOf), next) +
    '<tr><td style="padding:18px 24px 24px"><a href="' + h(d.site) + '" style="display:inline-block;background:#ffe100;color:#2a0f45;font:700 14px Arial,sans-serif;padding:10px 18px;border-radius:999px;text-decoration:none">Open Clash of Classes</a>' +
    '<p style="font:12px Arial,sans-serif;color:#999;margin:14px 0 0">Sent automatically by Clash of Classes. Recipients: Settings › DIGEST_EMAILS.</p></td></tr></table></td></tr></table></body></html>';
}

function digestRecipients_() {
  return String(getSetting_('DIGEST_EMAILS') || '').split(/[,;\s]+/).map(function (s) { return s.trim().toLowerCase(); }).filter(function (s) { return /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(s); });
}

// The weekly send (trigger). Once per week even if the trigger fires twice.
function weeklyDigest(force) {
  noteRun_('weeklyDigest');
  const props = PropertiesService.getScriptProperties(), d = weekData_(), to = digestRecipients_();
  if (!force && props.getProperty('digest-week') === d.weekOf) return 'Already sent this week.';
  if (!to.length) return 'No recipients: add emails to Settings › DIGEST_EMAILS.';
  const subject = 'This Week in Clash · week of ' + d.weekOf + ' · ' + titleSquad_(d.place[0]) + ' leads';
  try {
    MailApp.sendEmail({ to: to.join(','), subject: subject, htmlBody: digestHtml_(d), name: 'Clash of Classes' });
  } catch (e) { noteError_('weeklyDigest', e); throw e; }
  props.setProperty('digest-week', d.weekOf);
  const log = getSheetSafe_('Weekly Digest Log');
  if (log) log.appendRow([new Date(), d.weekOf, to.length, subject]);
  return 'Sent to ' + to.length + '.';
}

function weeklyDigestIfOn() { if (String(getSetting_('DIGEST') || 'OFF').toUpperCase() === 'ON') weeklyDigest(false); }

function menuTurnOnDigest_() {
  const ui = SpreadsheetApp.getUi();
  if (!digestRecipients_().length) { ui.alert('Add who should get it first: Settings › DIGEST_EMAILS (emails separated by commas).'); return; }
  setSettingValue_('DIGEST', 'ON');
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'weeklyDigestIfOn') ScriptApp.deleteTrigger(t); });
  const days = { MONDAY: ScriptApp.WeekDay.MONDAY, TUESDAY: ScriptApp.WeekDay.TUESDAY, WEDNESDAY: ScriptApp.WeekDay.WEDNESDAY, THURSDAY: ScriptApp.WeekDay.THURSDAY, FRIDAY: ScriptApp.WeekDay.FRIDAY };
  const dayKey = String(getSetting_('DIGEST_DAY') || 'FRIDAY').toUpperCase(), hour = Math.max(5, Math.min(20, Number(getSetting_('DIGEST_HOUR')) || 14));
  ScriptApp.newTrigger('weeklyDigestIfOn').timeBased().onWeekDay(days[dayKey] || ScriptApp.WeekDay.FRIDAY).atHour(hour).create();
  ui.alert('The weekly digest is ON: every ' + (days[dayKey] ? dayKey.charAt(0) + dayKey.slice(1).toLowerCase() : 'Friday') + ' around ' + hour + ':00 to ' + digestRecipients_().length + ' people.');
}

function menuTurnOffDigest_() {
  setSettingValue_('DIGEST', 'OFF');
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'weeklyDigestIfOn') ScriptApp.deleteTrigger(t); });
  SpreadsheetApp.getUi().alert('The weekly digest email is OFF. The This Week page on the site still updates.');
}

// A preview to whoever is running it, any time.
function menuPreviewDigest_() {
  const me = Session.getEffectiveUser().getEmail(), d = weekData_();
  MailApp.sendEmail({ to: me, subject: '[Preview] This Week in Clash · week of ' + d.weekOf, htmlBody: digestHtml_(d), name: 'Clash of Classes' });
  SpreadsheetApp.getUi().alert('Sent a preview to ' + me + '.');
}
