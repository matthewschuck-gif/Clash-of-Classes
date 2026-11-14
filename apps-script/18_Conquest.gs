/**
 * Clash of Classes -- Apps Script backend
 * File 18: Conquest. How territory on the school map changes hands, all editable without code.
 *
 * CONQUEST RULES tab: one row per way to win a zone. Map Zones › Conquer By names the rule.
 *   Method:      most points | most entries (times a squad earned) | staff pick (staff type Holder)
 *                | highest number | lowest number (a monthly number per squad: see SQUAD MONTHLY below)
 *   Counts:      event types or words in the event name, comma-separated (* = every point).
 *                For highest/lowest number: the measures to add up (attendance, lrr_ld, poll_votes...)
 *   Window Days: how far back it looks (blank = CLAIM_DAYS), or "last month" = the last calendar month.
 *                Highest/lowest number always use the last calendar month.
 *   Per 100:     divide by the squad's students (x 100), so a bigger squad has no edge. For counts
 *                like LRR/LDs, failing grades and interventions.
 *   Must Border: Risk rule. A squad that holds a zone next door attacks with the normal bar. A squad
 *                with no zones at all counts as next door everywhere, so nobody gets shut out.
 *   Defense %:   the holder's edge. A challenger needs this much more than the holder to take it.
 *   Far Attack %: Must Border zones only. A squad from farther away can still take the zone if it
 *                beats the holder by this much (on an open zone: beats the best squad next door by
 *                this much). The same bars for every squad, attacking or defending. Blank = only
 *                squads next door can attack.
 *   A zone with no Conquer By works the old way: most points from its Claimed By, CLAIM_DAYS.
 * THE MAP: zones are boxes on a grid (X, Y, W, H); the grid grows to fit them. Neighbors are
 *   zones whose boxes touch, plus anything listed in Borders (a hallway that connects two areas).
 *   The building layout is areas only (gyms, cafeteria, halls): never room numbers, doors or exits.
 * SQUAD MONTHLY tab: one row per squad per month (Month yyyy-MM | Squad | Students | a column per
 *   measure). Define a Climb-er's squad sync fills it from PowerSchool; staff can type in it too.
 *   A measure is read from (1) its column here, (2) poll_votes = Poll Votes rows, (3) attendance from a
 *   finalized month, or (4) for highest number only, how many events rows match its name.
 *   A highest/lowest zone changes hands only when all four squads have a number for the month.
 * MULTIPLIER (Map Zones): the zone's weekly pay x this (any number, 0 to 5; change it any time in the
 *   Map editor or the cell). Starting values: 2 = check-ins (students show up and check in: Pride Day,
 *   Move · Make · Talk, polls and the Wheel); 1.5 = school records (attendance, grades, LRR/LDs,
 *   interventions, STAR growth); 1 = points and staff results.
 * WINDOW DAYS "this month" = the month so far (the Front Office: last Friday's staff challenge winner
 *   takes it right away).
 * STAFF OVERRIDES (Admin › Map, or the tabs):
 *   Locked (Map Zones): ticked = the holder never changes on its own. Staff set Holder themselves.
 *   ZONE OVERRIDES tab: Month | Zone | Squad | Score | Note | Set By | Set At. A Score here replaces that
 *     squad's computed number for that zone that month (whatever it came from: Climb-er, check-ins,
 *     points). Month = the month the zone is judged on (last month for the monthly rules). The newest
 *     row for a Month + Zone + Squad wins; a row with a blank Score clears the override.
 * MAP BATTLES tab: every capture, shown under the map on the site, HQ and the TV.
 * The admin panel's Map tab (MapEditor.html) edits zones and rules by dragging and picking.
 * HOW THE MAP WORKS tab: the staff guide to all of this, rewritten on every Run setup.
 */

const RULE_HEADERS_ = ['Show', 'Rule', 'How It Works', 'Method', 'Counts', 'Window Days', 'Must Border', 'Defense %', 'Far Attack %', 'Per 100'];
const BATTLE_HEADERS_ = ['Date', 'Zone', 'Taken By', 'From', 'Rule'];
const OVERRIDE_HEADERS_ = ['Month', 'Zone', 'Squad', 'Score', 'Note', 'Set By', 'Set At'];
const RULE_METHODS_ = ['most points', 'most entries', 'staff pick', 'highest number', 'lowest number'];
const MEASURE_METHODS_ = ['highest number', 'lowest number'];
const MONTHLY_TAB_ = 'Squad Monthly';
// What Define a Climb-er (or staff) fills in each month, per squad. Counts, not percents, except
// attendance and pride_day (percent of the squad).
const MONTHLY_HEADERS_ = ['Month', 'Squad', 'Students', 'attendance', 'pride_day', 'lrr_ld', 'failing',
  'interventions_teacher', 'interventions_admin', 'shout_outs'];

const BUILDING_RULES_ = [
  // [Show, Rule, How It Works, Method, Counts, Window Days, Must Border, Defense %, Far Attack %, Per 100]
  // The monthly set (Menu › Use the monthly zone rules). Every one looks at the last calendar month.
  [true, 'Pride Day', 'The biggest share of your squad showing Mountaineer pride on Pride Days last month.', 'highest number', 'pride_day', 'last month', false, 0, '', false],
  [true, 'Full Climb', 'Every point your squad earned last month counts.', 'most points', '*', 'last month', false, 0, '', false],
  [true, 'Move · Make · Talk', 'Clash squad events last month: every member a General confirmed took part, plus the points the squad earned in Move, Make and Talk events.', 'highest number', 'event_participants, squad_event_points', 'last month', false, 0, '', false],
  [true, 'Clean Cafeteria', 'The fewest LRRs and lunch detentions per 100 students last month.', 'lowest number', 'lrr_ld', 'last month', false, 0, '', true],
  [true, 'Wheel & Poll Turnout', 'The most Daily Poll and Pulse votes plus wheel spins last month.', 'highest number', 'poll_votes, wheel', 'last month', false, 0, '', false],
  [true, 'Attendance', 'The highest attendance rate last month.', 'highest number', 'attendance', 'last month', false, 0, '', false],
  [true, 'Shout-out Storm', 'The most shout-outs last month.', 'highest number', 'shout_outs', 'last month', false, 0, '', false],
  [true, 'Growth Grind', 'The most STAR growth and Mountaineer Mindset summit points last month.', 'most points', 'star_growth, mountaineer_mindset', 'last month', false, 0, '', false],
  [true, 'Staff Challenge', 'Your squad\'s staff won the most First Friday Coffee staff challenges this month (it changes hands the day a winner is entered).', 'most entries', 'staff_challenge', 'this month', false, 0, '', false],
  [true, 'Passing Grades', 'The fewest failing grades per 100 students last month.', 'lowest number', 'failing', 'last month', false, 0, '', true],
  [true, 'Trait Dough Drive', 'The most Trait Dough earned last month.', 'most points', 'trait_dough', 'last month', false, 0, '', false],
  [true, 'Fewest Interventions', 'The fewest teacher and admin interventions combined, per 100 students, last month.', 'lowest number', 'interventions_teacher, interventions_admin', 'last month', false, 0, '', true],
  // Still here for anyone who wants them.
  [true, 'Siege', 'Risk rules: a squad next door needs 25% more points than the holder. A squad from farther away can still take it with 75% more.', 'most points', '*', 14, true, 25, 75, false],
  [true, 'Staff Pick', 'Staff decide who holds it.', 'staff pick', '', '', false, 0, '', false],
];

// Areas of the building on a 24 x 16 grid, roughly where they sit on the school map.
// [Zone, Wing, Summit, Conquer By, Weekly Points, X, Y, W, H, Multiplier]
// Multiplier 2 = check-ins: students show up and check in (Pride Day, Move · Make · Talk events, polls
// and the Wheel), so holding the zone pays double and squads push the habits. 1.5 = school records
// (attendance, passing grades, LRR/LDs, interventions, STAR growth). 1 = points and staff results.
const BUILDING_ZONES_ = [
  ['Gym 2', 'Athletics', false, 'Pride Day', 3, 0, 0, 4, 10, 2],
  ['Gym 1', 'Athletics', false, 'Move · Make · Talk', 3, 4, 0, 5, 8, 2],
  ['Cafeteria', 'Heart of the School', true, 'Clean Cafeteria', 5, 4, 8, 5, 4, 1.5],
  ['Music Wing', 'The Arts', false, 'Full Climb', 3, 0, 10, 4, 6, 1],
  ['Auditorium', 'The Arts', false, 'Wheel & Poll Turnout', 3, 4, 12, 5, 4, 2],
  ['50s Hall', 'North Halls', false, 'Attendance', 3, 9, 0, 6, 5, 1.5],
  ['60s Hall', 'North Halls', false, 'Shout-out Storm', 3, 9, 5, 6, 4, 1],
  ['Media Center', 'Heart of the School', false, 'Growth Grind', 3, 9, 9, 6, 3, 1.5],
  ['Front Office', 'Heart of the School', false, 'Staff Challenge', 2, 9, 12, 6, 3, 1],
  ['70s Hall', 'East Halls', false, 'Passing Grades', 3, 15, 6, 6, 4, 1.5],
  ['80s Hall', 'East Halls', false, 'Trait Dough Drive', 3, 15, 10, 6, 5, 1],
  ['90s Hall', 'East Halls', false, 'Fewest Interventions', 3, 21, 6, 3, 9, 1.5],
];

function buildingZoneRows_() {
  return BUILDING_ZONES_.map(function (z) {
    return { 'Show': true, 'Zone': z[0], 'Wing': z[1], 'Summit': z[2], 'Conquer By': z[3], 'Weekly Points': z[4], 'X': z[5], 'Y': z[6], 'W': z[7], 'H': z[8], 'Multiplier': z[9] };
  });
}

// Writes objects into a tab by its own header row (columns can be in any order), replacing rows 2+.
function writeRowsByHeader_(sheet, objs) {
  const head = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function (h) { return String(h).trim(); });
  const last = sheet.getLastRow();
  if (last > 1) sheet.getRange(2, 1, last - 1, head.length).clearContent();
  if (!objs.length) return;
  sheet.getRange(2, 1, objs.length, head.length).setValues(objs.map(function (o) {
    return head.map(function (h) { return o[h] === undefined || o[h] === null ? '' : o[h]; });
  }));
}

function headerCol_(sheet, name) {
  const head = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function (h) { return String(h).trim(); });
  return head.indexOf(name) + 1;
}

// Conquest Rules + Map Battles, and the newer columns on an older Map Zones tab.
function ensureConquestTabs_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  [['Conquest Rules', RULE_HEADERS_, BUILDING_RULES_], ['Map Battles', BATTLE_HEADERS_, null], [MONTHLY_TAB_, MONTHLY_HEADERS_, null], ['Zone Overrides', OVERRIDE_HEADERS_, null]].forEach(function (t) {
    if (ss.getSheetByName(t[0])) return;
    const sh = ss.insertSheet(t[0]);
    sh.getRange(1, 1, 1, t[1].length).setValues([t[1]]).setFontWeight('bold').setBackground('#490e6f').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    if (t[2]) sh.getRange(2, 1, t[2].length, t[1].length).setValues(t[2]);
    if (t[0] === 'Conquest Rules') {
      try {
        sh.getRange(2, RULE_HEADERS_.indexOf('Method') + 1, 200, 1)
          .setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(RULE_METHODS_, true).build());
      } catch (e) { console.warn('Conquest Rules dropdown: ' + e.message); }
    }
  });
  const cr = ss.getSheetByName('Conquest Rules');
  if (cr && cr.getLastColumn()) {
    const head = cr.getRange(1, 1, 1, cr.getLastColumn()).getValues()[0].map(function (h) { return String(h).trim(); });
    ['Far Attack %', 'Per 100'].forEach(function (h) {
      if (head.indexOf(h) > -1) return;
      head.push(h);
      cr.getRange(1, head.length).setValue(h).setFontWeight('bold').setBackground('#490e6f').setFontColor('#ffffff');
    });
  }
  const mz = ss.getSheetByName('Map Zones');
  if (mz && mz.getLastColumn()) {
    const head = mz.getRange(1, 1, 1, mz.getLastColumn()).getValues()[0].map(function (h) { return String(h).trim(); });
    ['Conquer By', 'Borders', 'Multiplier', 'Locked'].forEach(function (h) {
      if (head.indexOf(h) > -1) return;
      head.push(h);
      mz.getRange(1, head.length).setValue(h).setFontWeight('bold').setBackground('#490e6f').setFontColor('#ffffff');
    });
  }
}

function conquestRules_() {
  const out = {};
  sheetRows_('Conquest Rules').forEach(function (r) {
    const name = String(r['Rule'] || '').trim();
    if (!name || !ticked_(r['Show'])) return;
    const method = String(r['Method'] || '').trim().toLowerCase();
    const m = RULE_METHODS_.indexOf(method) > -1 ? method : 'most points';
    const thisMonth = /this\s*month/i.test(String(r['Window Days'] || '')) && MEASURE_METHODS_.indexOf(m) === -1;
    out[name.toLowerCase()] = { thisMonth: thisMonth, rule: name, text: String(r['How It Works'] || ''), method: m,
      counts: String(r['Counts'] || ''), days: Number(r['Window Days']) || 0, border: ticked_(r['Must Border']), defense: Math.max(0, Number(r['Defense %']) || 0),
      month: !thisMonth && (/month/i.test(String(r['Window Days'] || '')) || MEASURE_METHODS_.indexOf(m) > -1), per100: ticked_(r['Per 100']),
      far: String(r['Far Attack %'] == null ? '' : r['Far Attack %']).trim() === '' ? null : Math.max(0, Number(r['Far Attack %']) || 0) };
  });
  return out;
}

// Two boxes are neighbors when they share part of an edge (corners alone don't count).
function zonesTouch_(a, b) {
  const xo = Math.min(a.x + a.w, b.x + b.w) - Math.max(a.x, b.x);
  const yo = Math.min(a.y + a.h, b.y + b.h) - Math.max(a.y, b.y);
  return (xo === 0 && yo > 0) || (yo === 0 && xo > 0);
}

function zoneNeighbors_(zones) {
  const out = {}, byName = {};
  zones.forEach(function (z) { out[z.zone] = []; byName[z.zone.toLowerCase()] = z.zone; });
  function link(a, b) { if (a !== b && out[a].indexOf(b) === -1) { out[a].push(b); out[b].push(a); } }
  zones.forEach(function (a, i) { zones.slice(i + 1).forEach(function (b) { if (zonesTouch_(a, b)) link(a.zone, b.zone); }); });
  zones.forEach(function (z) {
    String(z.borders || '').split(',').forEach(function (n) { const other = byName[n.trim().toLowerCase()]; if (other) link(z.zone, other); });
  });
  return out;
}

function mapGrid_(zones) {
  let cols = 12, rows = 8;
  zones.forEach(function (z) { cols = Math.max(cols, z.x + z.w); rows = Math.max(rows, z.y + z.h); });
  return { cols: cols, rows: rows };
}

// Every shown zone with its holder, worked out from its rule. write = save captures to the Sheet.
function mapZones_(events, write) {
  const claimDays = Math.max(1, Number(getSetting_('CLAIM_DAYS')) || 14);
  const today = todayYmd_();
  events = events || readEventsTab_();
  const sheet = write ? getSheetSafe_('Map Zones') : null;
  const rules = conquestRules_();
  const zones = sheetRows_('Map Zones').filter(function (r) { return ticked_(r['Show']) && String(r['Zone'] || '').trim(); }).map(function (r) {
    const named = String(r['Conquer By'] || '').trim();
    const rule = rules[named.toLowerCase()] ||
      { rule: '', text: '', method: 'most points', counts: String(r['Claimed By'] || ''), days: 0, border: false, defense: 0, far: null, month: false, per100: false };
    let holder = squadOf_(r['Holder']);
    if (SQUAD_KEYS_.indexOf(holder) === -1) holder = '';
    return { row: r._row, zone: String(r['Zone']).trim(), wing: String(r['Wing'] || '').trim(), summit: ticked_(r['Summit']),
      rule: rule.rule, ruleText: rule.text, method: rule.method, claimedBy: rule.counts, days: rule.days || claimDays,
      border: rule.border, defense: rule.defense, far: rule.border ? rule.far : null, weekly: Number(r['Weekly Points']) || 0,
      month: !!rule.month, thisMonth: !!rule.thisMonth, per100: !!rule.per100, low: rule.method === 'lowest number',
      mult: r['Multiplier'] === '' || r['Multiplier'] == null ? 1 : Math.max(0, Number(r['Multiplier']) || 0),
      x: Math.max(0, Number(r['X']) || 0), y: Math.max(0, Number(r['Y']) || 0), w: Math.max(1, Number(r['W']) || 1), h: Math.max(1, Number(r['H']) || 1),
      borders: String(r['Borders'] || ''), holder: holder, since: String(r['Since'] || ''), locked: ticked_(r['Locked']) };
  });
  const nb = zoneNeighbors_(zones);
  const overrides = zoneOverrides_();
  // Borders are judged on who held what before this check, so the order of rows doesn't matter.
  const held = {}, count = { BLACK: 0, GOLD: 0, GREY: 0, PURPLE: 0 };
  zones.forEach(function (z) { held[z.zone] = z.holder; if (z.holder) count[z.holder]++; });
  const battles = [];
  const hCol = sheet ? headerCol_(sheet, 'Holder') : 0, sCol = sheet ? headerCol_(sheet, 'Since') : 0;
  const lastMonth = monthBounds_(-1);
  zones.forEach(function (z) {
    const from = z.month ? lastMonth[0] : z.thisMonth ? today.slice(0, 8) + '01' : ymdShift_(today, -z.days + 1), to = z.month ? lastMonth[1] : today;
    z.window = z.month ? lastMonth[0].slice(0, 7) : '';
    if (MEASURE_METHODS_.indexOf(z.method) > -1) z.score = measureScores_(z.claimedBy, from, to, events, z.per100, z.low);
    else {
      const tally = pointsBy_(events, from, to, eventMatcher_(z.claimedBy));
      z.score = z.method === 'most entries' ? tally.rows : tally.points;
      if (z.per100) z.score = per100_(z.score, from.slice(0, 7));
    }
    // Staff overrides for the month this zone is judged on (Zone Overrides tab / Admin › Map).
    z.judged = z.month ? z.window : today.slice(0, 7);
    z.computed = Object.assign({}, z.score);
    z.overridden = [];
    const ov = overrides[z.judged + '|' + z.zone.toLowerCase()] || {};
    Object.keys(ov).forEach(function (s) { z.score[s] = ov[s]; z.overridden.push(s); });
    z.neighbors = nb[z.zone];
    if (z.method === 'staff pick' || z.locked) return; // staff decide; nothing moves on its own
    // Highest/lowest number: nothing changes until every squad has a number for the month.
    if (SQUAD_KEYS_.some(function (s) { return z.score[s] === null || z.score[s] === undefined; })) return;
    const low = z.low;
    // Next door (or no border rule, or no zones yet): the normal bar. From farther away: the Far
    // Attack bar, the same for every squad (blank = only next door can attack).
    const nextDoor = function (s) { return !z.border || s === z.holder || !count[s] || z.neighbors.some(function (n) { return held[n] === s; }); };
    const qualifies = SQUAD_KEYS_.filter(function (s) {
      // Most: a squad needs something to show. Lowest: zero is the best there is.
      if (s === z.holder || (!low && !(z.score[s] > 0))) return false;
      const near = nextDoor(s);
      if (!near && z.far === null) return false;
      // The bar is the holder's score; on an open zone, the best squad that could walk in.
      const others = SQUAD_KEYS_.filter(function (o) { return o !== s && nextDoor(o); }).map(function (o) { return z.score[o]; });
      const ref = z.holder ? (z.score[z.holder] || 0)
        : low ? Math.min.apply(null, others.concat([Infinity])) : Math.max.apply(null, others.concat([0]));
      const edge = near ? (z.holder ? z.defense : 0) : z.far;
      // Lowest: beat the bar by being that much lower (a 25% edge = 25% fewer than the holder).
      return low ? z.score[s] < (ref === Infinity ? Infinity : ref * (1 - edge / 100)) : z.score[s] > ref * (1 + edge / 100);
    });
    const vals = qualifies.map(function (s) { return z.score[s]; });
    const best = low ? Math.min.apply(null, vals.concat([Infinity])) : Math.max.apply(null, vals.concat([0]));
    const leaders = qualifies.filter(function (s) { return z.score[s] === best; });
    // King of the hill: the holder keeps it on a tie, a quiet stretch, or when nobody clears the bar.
    if (leaders.length !== 1) return;
    battles.push([today, z.zone, leaders[0], z.holder, z.rule || 'Most points']);
    z.holder = leaders[0]; z.since = today;
    if (sheet && hCol) sheet.getRange(z.row, hCol).setValue(z.holder);
    if (sheet && sCol) sheet.getRange(z.row, sCol).setValue(today);
  });
  const log = battles.length && write ? getSheetSafe_('Map Battles') : null;
  if (log) battles.forEach(function (b) { log.appendRow(b); });
  return zones;
}

// { 'yyyy-MM|zone': { SQUAD: score } } from the Zone Overrides tab (newest row wins; blank Score clears).
function zoneOverrides_() {
  const out = {};
  sheetRows_('Zone Overrides').forEach(function (r) {
    const ym = ymOf_(r['Month']), zone = String(r['Zone'] || '').trim().toLowerCase(), s = squadOf_(r['Squad']);
    if (!ym || !zone || SQUAD_KEYS_.indexOf(s) === -1) return;
    const k = ym + '|' + zone, raw = String(r['Score'] == null ? '' : r['Score']).trim();
    out[k] = out[k] || {};
    if (raw === '') delete out[k][s];
    else if (!isNaN(Number(raw.replace(/[%,]/g, '')))) out[k][s] = Number(raw.replace(/[%,]/g, ''));
  });
  return out;
}

// ── SQUAD MONTHLY: the numbers Define a Climb-er sends each month ───
// "2026-12", a Date cell (Sheets turns 2026-12 into 2026-12-01) or "December 2026" -> "2026-12".
function ymOf_(v) {
  const s = String(v == null ? '' : v).trim();
  let m = s.match(/^(\d{4})-(\d{1,2})/);
  if (m) return m[1] + '-' + ('0' + m[2]).slice(-2);
  const names = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'];
  m = s.toLowerCase().match(/([a-z]{3})[a-z]*\.?\s+(\d{4})/);
  if (m && names.indexOf(m[1]) > -1) return m[2] + '-' + ('0' + (names.indexOf(m[1]) + 1)).slice(-2);
  return '';
}

// { SQUAD: { column: number } } for one month. Columns are matched without case, spaces or dashes.
function squadMonthly_(ym) {
  const out = {};
  sheetRows_(MONTHLY_TAB_).forEach(function (r) {
    const s = squadOf_(r['Squad']);
    if (SQUAD_KEYS_.indexOf(s) === -1 || ymOf_(r['Month']) !== ym) return;
    out[s] = out[s] || {};
    Object.keys(r).forEach(function (k) {
      if (k === '_row' || k === 'Month' || k === 'Squad' || r[k] === '' || r[k] == null) return;
      const n = Number(String(r[k]).replace(/[%,\s]/g, ''));
      if (!isNaN(n)) out[s][measureKey_(k)] = n;
    });
  });
  return out;
}

function measureKey_(k) { return String(k || '').trim().toLowerCase().replace(/[\s\-]+/g, '_'); }

// Students per squad: Squad Monthly's Students, else the Students tab (the roster), else the Student Board.
function squadSizes_(ym, monthly) {
  monthly = monthly || squadMonthly_(ym);
  const board = { BLACK: 0, GOLD: 0, GREY: 0, PURPLE: 0 };
  const count = function (rows) { rows.forEach(function (r) { const s = squadOf_(r['Squad']); if (board[s] !== undefined) board[s]++; }); };
  count(sheetRows_('Students'));
  if (!SQUAD_KEYS_.some(function (s) { return board[s]; })) count(sheetRows_(getSetting_('STUDENT_BOARD_TAB') || 'Student Board'));
  const out = {};
  SQUAD_KEYS_.forEach(function (s) { const n = monthly[s] && monthly[s].students; out[s] = n > 0 ? n : board[s]; });
  return out;
}

function per100_(score, ym, monthly) {
  const size = squadSizes_(ym, monthly), out = {};
  SQUAD_KEYS_.forEach(function (s) {
    out[s] = score[s] === null || score[s] === undefined || !(size[s] > 0) ? null : Math.round(1000 * score[s] / size[s] * 100) / 1000;
  });
  return out;
}

// 800# check-ins (Poll Votes) between from and to on polls whose Kind or question has one of the words.
function checkInsBy_(from, to, spec) {
  const words = listOf_(spec), polls = {};
  sheetRows_('Daily Polls').forEach(function (p) {
    if (pollKind_(p['Kind']) !== 'Check-in') return;
    const text = (String(p['Kind'] || '') + ' ' + String(p['Question'] || '')).toLowerCase();
    polls[String(p._row)] = !words.length || words.some(function (w) { return text.indexOf(w) > -1; });
  });
  const out = { BLACK: 0, GOLD: 0, GREY: 0, PURPLE: 0 };
  sheetRows_('Poll Votes').forEach(function (v) {
    const d = String(v['Date'] || v['At'] || '').slice(0, 10), s = squadOf_(v['Squad']);
    if (d >= from && d <= to && out[s] !== undefined && polls[String(v['Poll'])]) out[s]++;
  });
  return out;
}

// A squad's number for the month from a list of measures (added together). null = no number yet.
// Each measure: its Squad Monthly column; poll_votes = Poll Votes rows; attendance = a finalized
// month (Admin › Metrics); for highest-number rules only, how many events rows match its name.
function measureScores_(spec, from, to, events, per100, low) {
  const ym = from.slice(0, 7), monthly = squadMonthly_(ym);
  const out = { BLACK: 0, GOLD: 0, GREY: 0, PURPLE: 0 };
  const add = function (s, v) { if (out[s] !== null) out[s] = v === null || v === undefined ? null : out[s] + v; };
  const names = listOf_(spec).map(measureKey_);
  if (!names.length) return { BLACK: null, GOLD: null, GREY: null, PURPLE: null };
  names.forEach(function (n) {
    if (SQUAD_KEYS_.some(function (s) { return monthly[s] && monthly[s][n] !== undefined; })) {
      SQUAD_KEYS_.forEach(function (s) { add(s, monthly[s] ? monthly[s][n] : null); });
      return;
    }
    // Squad events (26_CheckIns.gs): confirmed participants, and points from Move / Make / Talk events.
    if (n === 'event_participants') { const c = eventParticipants_(from, to); SQUAD_KEYS_.forEach(function (s) { add(s, c[s]); }); return; }
    if (n === 'squad_event_points') { const p = pointsBy_(events, from, to, eventMatcher_(SQUAD_EVENT_TYPES_)).points; SQUAD_KEYS_.forEach(function (s) { add(s, p[s]); }); return; }
    // pride_day with no Squad Monthly number: 800# check-ins on Pride Day Check-in polls, per 100 students.
    if (n === 'pride_day') {
      const c = checkInsBy_(from, to, 'pride');
      if (SQUAD_KEYS_.some(function (s) { return c[s]; })) {
        const size = squadSizes_(ym, monthly);
        SQUAD_KEYS_.forEach(function (s) { add(s, size[s] > 0 ? Math.round(1000 * c[s] / size[s]) / 10 : null); });
        return;
      }
    }
    if (n === 'poll_votes') {
      const c = { BLACK: 0, GOLD: 0, GREY: 0, PURPLE: 0 };
      sheetRows_('Poll Votes').forEach(function (v) { const d = String(v['Date'] || v['At'] || '').slice(0, 10), s = squadOf_(v['Squad']); if (d >= from && d <= to && c[s] !== undefined) c[s]++; });
      SQUAD_KEYS_.forEach(function (s) { add(s, c[s]); });
      return;
    }
    if (n === 'attendance' || n === 'referral' || n === 'grades') {
      let months = [];
      try { months = (loadAppData_('coc').value || {}).monthlyMetrics || []; } catch (e) { months = []; }
      const m = months.filter(function (x) { return ymOf_(x.month) === ym; }).pop();
      SQUAD_KEYS_.forEach(function (s) { add(s, m && m[n] && m[n][s] !== undefined && m[n][s] !== '' ? Number(m[n][s]) : null); });
      return;
    }
    if (low) { SQUAD_KEYS_.forEach(function (s) { add(s, null); }); return; } // a missing count is never "zero problems"
    const forms = [n, n.replace(/_/g, ' '), n.replace(/s$/, ''), n.replace(/s$/, '').replace(/_/g, ' ')];
    const tally = pointsBy_(events, from, to, function (e) {
      // Name or note (wheel spins are noted "Wheel spin").
      const type = measureKey_(e.type), name = (String(e.name || '') + ' ' + String(e.note || '')).toLowerCase().replace(/[\-_]+/g, ' ');
      return forms.some(function (f) { return type === f || type === f.replace(/ /g, '_') || name.indexOf(f.replace(/_/g, ' ')) > -1; });
    });
    SQUAD_KEYS_.forEach(function (s) { add(s, tally.rows[s]); });
  });
  return per100 ? per100_(out, ym, monthly) : out;
}

function recentBattles_(n) {
  const from = ymdShift_(todayYmd_(), -14);
  return sheetRows_('Map Battles').filter(function (b) { return String(b['Date'] || '').slice(0, 10) >= from && b['Zone']; })
    .slice(-(n || 5)).reverse().map(function (b) {
      return { date: String(b['Date']).slice(0, 10), zone: String(b['Zone']), to: squadOf_(b['Taken By']), from: squadOf_(b['From']), rule: String(b['Rule'] || '') };
    });
}

// What every page gets for the map (home, TV, Generals HQ, My Climb).
function mapPayload_(events, zones) {
  if (!gameOn_('GAME_MAP')) return null;
  zones = zones || mapZones_(events);
  const rules = {};
  zones.forEach(function (z) {
    const k = z.rule || 'Most points';
    rules[k] = rules[k] || { rule: k, text: z.ruleText || 'Most points from ' + (z.claimedBy || 'anything') + ' in the last ' + z.days + ' days.',
      border: z.border, defense: z.defense, far: z.far, days: z.days, method: z.method, month: z.month, thisMonth: z.thisMonth, window: z.window, per100: z.per100, low: z.low, zones: [] };
    rules[k].zones.push(z.zone);
  });
  return {
    zones: zones.map(function (z) { const o = Object.assign({}, z); delete o.row; return o; }),
    wings: heldWings_(zones), grid: mapGrid_(zones), rules: Object.keys(rules).map(function (k) { return rules[k]; }),
    battles: recentBattles_(5), style: String(getSetting_('MAP_STYLE') || 'puzzle').toLowerCase(),
    claimDays: Number(getSetting_('CLAIM_DAYS')) || 14,
  };
}

// ── THE ADMIN PANEL'S MAP TAB ───────────────────────────────────────
function clashMapLoad(token) {
  requireAdmin_(token);
  ensureGameTabs_();
  const zones = sheetRows_('Map Zones').filter(function (r) { return String(r['Zone'] || '').trim(); }).map(function (r) {
    return { show: ticked_(r['Show']), zone: String(r['Zone']).trim(), wing: String(r['Wing'] || ''), summit: ticked_(r['Summit']),
      conquerBy: String(r['Conquer By'] || ''), claimedBy: String(r['Claimed By'] || ''), weekly: Number(r['Weekly Points']) || 0,
      x: Number(r['X']) || 0, y: Number(r['Y']) || 0, w: Math.max(1, Number(r['W']) || 1), h: Math.max(1, Number(r['H']) || 1),
      borders: String(r['Borders'] || ''), holder: squadOf_(r['Holder']), since: String(r['Since'] || ''),
      mult: r['Multiplier'] === '' || r['Multiplier'] == null ? 1 : Number(r['Multiplier']) || 0, locked: ticked_(r['Locked']) };
  });
  // Each zone's numbers right now (computed, and any staff override), for the editor's override boxes.
  const scores = {};
  try {
    mapZones_(null, false).forEach(function (z) {
      const ov = {};
      z.overridden.forEach(function (s) { ov[s] = z.score[s]; });
      scores[z.zone] = { month: z.judged, computed: z.computed, overrides: ov, low: z.low, method: z.method };
    });
  } catch (e) { console.warn('Map scores: ' + e.message); }
  const rules = sheetRows_('Conquest Rules').filter(function (r) { return String(r['Rule'] || '').trim(); }).map(function (r) {
    return { show: ticked_(r['Show']), rule: String(r['Rule']).trim(), text: String(r['How It Works'] || ''), method: String(r['Method'] || 'most points').toLowerCase(),
      counts: String(r['Counts'] || ''), days: r['Window Days'] === '' ? '' : /month/i.test(String(r['Window Days'])) ? 'last month' : Number(r['Window Days']) || '',
      border: ticked_(r['Must Border']), defense: Number(r['Defense %']) || 0,
      far: r['Far Attack %'] === '' || r['Far Attack %'] == null ? '' : Number(r['Far Attack %']) || 0, per100: ticked_(r['Per 100']) };
  });
  return JSON.stringify({ zones: zones, rules: rules, scores: scores, methods: RULE_METHODS_, claimDays: Number(getSetting_('CLAIM_DAYS')) || 14,
    building: { zones: buildingZoneRows_(), rules: BUILDING_RULES_.map(function (r) {
      return { show: r[0], rule: r[1], text: r[2], method: r[3], counts: r[4], days: r[5], border: r[6], defense: r[7], far: r[8], per100: r[9] }; }) } });
}

function clashMapSave(token, json) {
  requireAdmin_(token);
  const d = JSON.parse(json || '{}');
  if (!Array.isArray(d.zones) || !Array.isArray(d.rules)) throw new Error('Nothing to save.');
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    ensureGameTabs_();
    const clip = function (v, max) { return String(v == null ? '' : v).trim().slice(0, max || 120); };
    const num = function (v, lo, hi) { const n = Math.round(Number(v) || 0); return Math.max(lo, Math.min(hi, n)); };
    const names = {};
    const rules = d.rules.filter(function (r) { return clip(r.rule); }).map(function (r) {
      const name = clip(r.rule, 60);
      if (names[name.toLowerCase()]) throw new Error('Two rules are named "' + name + '".');
      names[name.toLowerCase()] = 1;
      const method = String(r.method || '').toLowerCase();
      return { 'Show': r.show !== false, 'Rule': name, 'How It Works': clip(r.text, 300), 'Method': RULE_METHODS_.indexOf(method) > -1 ? method : 'most points',
        'Counts': clip(r.counts, 200), 'Window Days': r.days === '' || r.days == null ? '' : /month/i.test(String(r.days)) ? 'last month' : num(r.days, 1, 365),
        'Must Border': !!r.border, 'Defense %': num(r.defense, 0, 500),
        'Far Attack %': r.far === '' || r.far == null ? '' : num(r.far, 0, 1000), 'Per 100': !!r.per100 };
    });
    const old = {};
    sheetRows_('Map Zones').forEach(function (r) { old[String(r['Zone'] || '').trim().toLowerCase()] = r; });
    const seen = {}, today = todayYmd_();
    const zones = d.zones.filter(function (z) { return clip(z.zone); }).map(function (z) {
      const name = clip(z.zone, 40);
      if (seen[name.toLowerCase()]) throw new Error('Two zones are named "' + name + '".');
      seen[name.toLowerCase()] = 1;
      const before = old[name.toLowerCase()] || {};
      let holder = squadOf_(z.holder);
      if (SQUAD_KEYS_.indexOf(holder) === -1) holder = '';
      const since = holder === squadOf_(before['Holder']) ? String(before['Since'] || '') : holder ? today : '';
      return { 'Show': z.show !== false, 'Zone': name, 'Wing': clip(z.wing, 40), 'Summit': !!z.summit, 'Claimed By': clip(z.claimedBy, 200),
        'Weekly Points': num(z.weekly, 0, 100), 'X': num(z.x, 0, 60), 'Y': num(z.y, 0, 40), 'W': num(z.w, 1, 60), 'H': num(z.h, 1, 40),
        'Holder': holder, 'Since': since, 'Conquer By': clip(z.conquerBy, 60), 'Borders': clip(z.borders, 300),
        'Multiplier': z.mult === '' || z.mult == null ? 1 : Math.max(0, Math.min(5, Math.round(20 * (Number(z.mult) || 0)) / 20)), 'Locked': !!z.locked };
    });
    writeRowsByHeader_(getSheetSafe_('Conquest Rules'), rules);
    writeRowsByHeader_(getSheetSafe_('Map Zones'), zones);
    // Score overrides: [{ zone, month, squad, value }] ('' clears). Appended, so the tab keeps the history.
    const ovs = Array.isArray(d.overrides) ? d.overrides : [];
    const osh = ovs.length ? getSheetSafe_('Zone Overrides') : null;
    ovs.forEach(function (o) {
      const s = squadOf_(o.squad), month = ymOf_(o.month), zone = clip(o.zone, 40);
      if (!osh || !zone || !month || SQUAD_KEYS_.indexOf(s) === -1) return;
      const raw = String(o.value == null ? '' : o.value).trim();
      if (raw !== '' && isNaN(Number(raw))) throw new Error('Override for ' + zone + ' (' + s + ') must be a number.');
      osh.appendRow([month, zone, s, raw === '' ? '' : Number(raw), clip(o.note, 200), 'Admin', Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm')]);
    });
    CacheService.getScriptCache().remove('game-tick');
    // Re-judge now so overrides and unlocked zones show the right holder straight away.
    try { mapZones_(null, true); } catch (e) { console.warn('Map re-check: ' + e.message); }
    return JSON.stringify({ ok: true, zones: zones.length, rules: rules.length, overrides: ovs.length });
  } finally { lock.releaseLock(); }
}

// Menu: swap the mountain map for the building layout (claims start fresh; old rules stay).
function menuBuildingMap_() {
  const ui = SpreadsheetApp.getUi();
  if (ui.alert('Use the building map?', 'This replaces every row on Map Zones with areas of the building (gyms, cafeteria, halls). Holders start fresh. ' +
    'Conquest Rules keeps your rules and adds any missing ones. You can move and rename everything in the admin panel\'s Map tab after.', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  applyBuildingMap_();
  ui.alert('Done: the map is the building now. Admin panel › 🗺️ Map to drag zones, rename them and pick how each one is won.');
}

function applyBuildingMap_() {
  ensureGameTabs_();
  const have = {};
  sheetRows_('Conquest Rules').forEach(function (r) { have[String(r['Rule'] || '').trim().toLowerCase()] = 1; });
  const sh = getSheetSafe_('Conquest Rules');
  BUILDING_RULES_.forEach(function (r) { if (!have[r[1].toLowerCase()]) sh.appendRow(r); });
  writeRowsByHeader_(getSheetSafe_('Map Zones'), buildingZoneRows_());
  setupMapGuideTab_();
}

// Menu: the monthly zone rules on the map you already have (zone spots and names stay). Each zone in
// BUILDING_ZONES_ gets its rule and Multiplier; same-named rules are rewritten, others are kept.
// fresh = every zone starts unclaimed (a clean January start).
function menuMonthlyRules_() {
  const ui = SpreadsheetApp.getUi();
  const a = ui.alert('Use the monthly zone rules?', 'Every zone is won on last month\'s numbers (Pride Day, attendance, LRR/LDs, failing grades, interventions, shout-outs...), and ' +
    'zones students make the data for pay 2x. Zones keep their spots.\n\nStart every zone unclaimed too? (Yes = clean start, No = keep holders)', ui.ButtonSet.YES_NO_CANCEL);
  if (a === ui.Button.CANCEL || a === ui.Button.CLOSE) return;
  const r = applyMonthlyRules_(a === ui.Button.YES);
  ui.alert('Done: ' + r.zones + ' zones set, ' + r.rules + ' rules written.' + (r.missing.length ? '\nNot on your map (add them in the Map editor): ' + r.missing.join(', ') : '') +
    '\n\nDefine a Climb-er fills the Squad Monthly tab; until it does, the highest/lowest zones wait (nobody loses a zone over missing numbers).');
}

function applyMonthlyRules_(fresh) {
  ensureGameTabs_();
  const crSheet = getSheetSafe_('Conquest Rules');
  const byName = {};
  BUILDING_RULES_.forEach(function (r) { byName[r[1].toLowerCase()] = r; });
  const kept = sheetRows_('Conquest Rules').filter(function (r) { return String(r['Rule'] || '').trim() && !byName[String(r['Rule']).trim().toLowerCase()]; });
  const rules = BUILDING_RULES_.map(function (r) { const o = {}; RULE_HEADERS_.forEach(function (h, i) { o[h] = r[i]; }); return o; })
    .concat(kept.map(function (r) { const o = Object.assign({}, r); delete o._row; return o; }));
  writeRowsByHeader_(crSheet, rules);
  const plan = {};
  BUILDING_ZONES_.forEach(function (z) { plan[z[0].toLowerCase()] = z; });
  const found = {};
  const zones = sheetRows_('Map Zones').map(function (r) {
    const o = Object.assign({}, r); delete o._row;
    const z = plan[String(r['Zone'] || '').trim().toLowerCase()];
    if (z) { found[z[0]] = 1; o['Conquer By'] = z[3]; o['Multiplier'] = z[9]; }
    if (fresh) { o['Holder'] = ''; o['Since'] = ''; }
    return o;
  });
  writeRowsByHeader_(getSheetSafe_('Map Zones'), zones);
  CacheService.getScriptCache().remove('game-tick');
  setupMapGuideTab_();
  return { zones: Object.keys(found).length, rules: rules.length, missing: BUILDING_ZONES_.map(function (z) { return z[0]; }).filter(function (n) { return !found[n]; }) };
}

// ── THE "HOW THE MAP WORKS" TAB ─────────────────────────────────────
// Staff guide to the map, the ways to conquer and what owning pays. Rewritten every Run setup
// (and by the building-map menu), like "How To Use This Sheet", so it never drifts from the code.
const MAP_GUIDE_TAB_ = 'How The Map Works';

function setupMapGuideTab_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(MAP_GUIDE_TAB_);
  if (!sheet) sheet = ss.insertSheet(MAP_GUIDE_TAB_);
  else sheet.clear();
  const H = '## '; // section heading marker (stripped below and bolded)
  const rows = [
    'Clash of Classes -- How The Map Works (zones, ways to conquer, what owning pays)',
    '',
    'The school map is a Risk-style game laid on top of the points squads already earn. Each area of the',
    'building is a zone. A squad takes a zone by earning the most of the right kind of points; holding zones',
    'pays bonus points every week. Nothing on the map ever takes points away.',
    '',
    H + 'THE EASY WAY: THE ADMIN PANEL',
    'clashofclasses.org -> Admin -> 🗺️ Map edits everything on this page without touching the Sheet:',
    '- Drag a zone to move it. Drag its bottom-right corner to resize it. Zones snap to a grid.',
    '- Click a zone to rename it, set its wing, pick how it is won, set its weekly points, tick it as the',
    '  Summit, or hide it.',
    '- Below the map, "Ways to conquer" is the Conquest Rules tab: edit, add or delete rules there.',
    '- Click "Save map" when you are done. Nothing changes until you save.',
    'Everything below can also be typed straight into the tabs; both ways stay in sync.',
    '',
    H + 'THE TABS',
    'Map Zones -- one row per zone:',
    '     Show          -- untick to hide a zone without deleting it.',
    '     Zone          -- the name students see (e.g. Cafeteria, 60s Hall). Must be unique.',
    '     Wing          -- a color set. Hold every zone in a wing for the wing bonus (see WHAT OWNING PAYS).',
    '     Summit        -- tick ONE zone. Its holder is King of the Mountain and it pays double.',
    '     Conquer By    -- the name of a rule on the Conquest Rules tab. Blank = the old way (Claimed By).',
    '     Claimed By    -- only used when Conquer By is blank: event types or words; * = every point.',
    '     Weekly Points -- what the holder earns each week.',
    '     Multiplier    -- Weekly Points x this (blank = 1). 2 = check-ins, 1.5 = school records. Any number.',
    '     X, Y, W, H    -- where the zone sits on the grid and how big it is (the Map editor sets these).',
    '     Borders       -- zones it connects to that do not touch on the map, comma-separated (a hallway).',
    '     Holder, Since -- filled in automatically. Only type a Holder on a Staff Pick or Locked zone.',
    '     Locked        -- tick to freeze the holder: nothing changes it automatically until you untick it.',
    'Conquest Rules -- one row per way to win a zone (see HOW A ZONE CHANGES HANDS).',
    'Map Battles -- every capture, written automatically. Shown under the map. Do not type here.',
    'Engine Powers -- lasting powers a squad unlocks; the "rent" ones pay extra for every zone held.',
    'Chance Cards / Squad Hand -- the Summit Fever card doubles one week of territory points.',
    '',
    H + 'HOW A ZONE CHANGES HANDS (THE CONQUEST RULES TAB)',
    'Each zone picks one rule. Each rule row has:',
    '     Show          -- untick to switch a rule off (zones using it go back to the old way).',
    '     Rule          -- its name. This is what Conquer By on Map Zones points to.',
    '     How It Works  -- one sentence students see under the map ("How to win each zone").',
    '     Method        -- most points  = most points from the Counts events in the window.',
    '                      most entries = most TIMES a squad earned from those events (rewards turnout,',
    '                                     not one big score).',
    '                      staff pick   = never changes on its own. Type the squad in Holder on Map Zones',
    '                                     (or pick "Held by" in the Map editor).',
    '                      highest number = the biggest monthly number (attendance %, Pride Day %, shout-outs).',
    '                      lowest number  = the smallest monthly number (LRR/LDs, failing grades, interventions).',
    '                                     Both read last calendar month from the Squad Monthly tab (below).',
    '     Counts        -- which points count: event types or words in the event name, comma-separated.',
    '                      * = every point. Common types: trait_dough, daily_poll, star_growth,',
    '                      mountaineer_mindset, duel_bonus, mayhem_week, mayhem, shout, lunch.',
    '                      Game bonuses (territory, chance cards, etc.) and private behavior rows never count.',
    '     Window Days   -- how many days back it looks. Blank = CLAIM_DAYS on the Settings tab (14).',
    '                      "last month" = the last full calendar month (all the monthly rules use this).',
    '     Must Border   -- the Risk rule. Tick it to make the zone a Siege zone (see below).',
    '     Defense %     -- the holder\'s edge: a challenger needs this much MORE than the holder. 0 = just more.',
    '     Far Attack %  -- Siege zones only: how much more a squad from farther away needs. Blank = only',
    '                      squads next door can attack.',
    '     Per 100       -- divide by the squad\'s students and x 100, so a bigger squad has no edge.',
    'The holder always keeps a zone on a tie or when nobody scored (king of the hill).',
    '',
    H + 'THE MONTHLY ZONES (Menu -> Use the monthly zone rules)',
    'Every zone but the Front Office is won on LAST calendar month, so it holds steady all month and changes on the 1st.',
    '     Gym 2 (old gym)  Pride Day             highest Pride Day participation %        pays 2x  (check-ins)',
    '     Gym 1 (new gym)  Move · Make · Talk    confirmed participants + event points    pays 2x  (check-ins)',
    '     Cafeteria        Clean Cafeteria       fewest LRR + LD students per 100         Summit, pays 1.5x',
    '     Music Wing       Full Climb            most total points                        pays 1x',
    '     Auditorium       Wheel & Poll Turnout  most poll/Pulse votes + wheel spins      pays 2x  (check-ins)',
    '     50s Hall         Attendance            highest attendance rate                  pays 1.5x',
    '     60s Hall         Shout-out Storm       most shout-outs                          pays 1x',
    '     Media Center     Growth Grind          most STAR growth + Mindset summit points pays 1.5x',
    '     Front Office     Staff Challenge       most First Friday staff wins THIS month  pays 1x',
    '     70s Hall         Passing Grades        fewest failing grades per 100            pays 1.5x',
    '     80s Hall         Trait Dough Drive     most Trait Dough earned                  pays 1x',
    '     90s Hall         Fewest Interventions  fewest teacher + admin interventions/100 pays 1.5x',
    'Why: 2x zones are won by students showing up and checking in (Pride Day, squad events, polls and the',
    'Wheel), so holding them grows those habits and the data. 1.5x zones come from school records',
    '(attendance, grades, LRR/LDs, interventions, STAR). 1x = points and staff results. Change any',
    'zone\'s Multiplier any time (Map editor or the Map Zones cell, 0 to 5, e.g. 1.25).',
    'Staff Challenge (First Friday Coffee): add one events row for the winning squad with type',
    '     staff_challenge (1 point is plenty). Window Days "this month": the squad with the most',
    '     staff_challenge rows this month takes the Front Office as soon as the row is in.',
    'LRR + LD (Cafeteria): filled into Squad Monthly lrr_ld once a day from the Follow-Up sheets',
    '     (Settings LRR_SHEET, CONSEQUENCE_SHEET; FOLLOWUP_STATS ON). A number you type is kept.',
    'Move · Make · Talk (Gym 1) = confirmed participants + squad event points, last month:',
    '     1. On Sign-up Events, give each Clash squad event a Date and a Kind (Move, Make or Talk).',
    '     2. On that day (and CHECKIN_DAYS after), students check in on the site with their 800#.',
    '     3. Generals confirm or remove each check-in in Generals HQ -> Check-ins (and add anyone who forgot).',
    '        Only confirmed check-ins count (Event Check-ins tab).',
    '     4. When you give the squads points for the event (Admin -> Add Event or Bulk Entry), pick its Type:',
    '        Squad event: Move, Make or Talk. Those points add to the participants.',
    'Pride Day: Climb-er\'s pride_day column (% of the squad). Until Climb-er sends it, add a Pride Day row',
    '     per squad (Type: Pride Day) with how many took part; the most Pride Day rows wins.',
    '',
    H + 'SQUAD MONTHLY (filled by Define a Climb-er)',
    'One row per squad per month. Month = 2026-12 (any date in the month works). Columns:',
    '     Students            -- squad size that month (for Per 100). Blank = count on the Student Board.',
    '     attendance          -- attendance rate, 0-100.        pride_day     -- Pride Day participation %.',
    '     shout_outs          -- shout-outs received.',
    '     lrr_ld              -- LRRs + lunch detentions.       failing       -- failing grades (count).',
    '     interventions_teacher, interventions_admin -- counts; the 90s Hall adds the two.',
    'Add any other column and name it in a rule\'s Counts (several, comma-separated, are added together).',
    'Also built in: poll_votes (Daily Poll + Pulse votes), attendance from a finalized month (Admin -> Metrics),',
    'event_participants (confirmed squad event check-ins), squad_event_points (Move / Make / Talk event points).',
    'A highest/lowest zone waits until ALL FOUR squads have a number that month: missing data never costs',
    'anyone a zone, and a blank count is never read as zero problems.',
    'Claims are checked about every 10 minutes (and on the hourly game check).',
    '',
    H + 'SIEGE ZONES (MUST BORDER) -- RISK RULES',
    '- Zones are neighbors when their boxes touch on the map, or when one lists the other in Borders.',
    '- A squad that holds a zone next door attacks with the Defense % bar.',
    '- A squad with NO zones at all counts as next door everywhere, so nobody is ever shut out.',
    '- A squad from farther away can still take it with the Far Attack % bar (if Far Attack % is filled in).',
    '- On an open zone (nobody holds it), a squad from farther away must beat the best squad next door by',
    '  the Far Attack %.',
    '- It works both ways: whoever holds the zone defends with the same bars, so a zone taken from afar can',
    '  be taken back by a neighbor with just the Defense % edge.',
    'Example (Siege: Defense 25%, Far Attack 75%). GOLD holds the 60s Hall with 40 points this window.',
    '     PURPLE holds the 50s Hall next door and has 52 points -> 52 is more than 40 x 1.25 = 50 -> PURPLE takes it.',
    '     BLACK holds nothing next door and has 65 points -> it needs more than 40 x 1.75 = 70 -> not yet.',
    '',
    H + 'WHAT OWNING PAYS (THE MULTIPLIERS)',
    'Once a week, each squad gets one "Territory" row on the events tab for the week that just ended:',
    '     Weekly Points     -- every zone held pays its Weekly Points (Map Zones column). Change per zone.',
    '     The Summit        -- the zone with Summit ticked pays 2x its Weekly Points. (2x is fixed.)',
    '     Multiplier        -- each zone\'s own Multiplier on top (Map Zones column; 2 check-ins, 1.5 records).',
    '     Wing bonus        -- hold EVERY zone in a wing and that wing\'s weekly points pay again.',
    '                          Settings -> WING_MULTIPLIER (2 = the wing pays double in total, 3 = triple).',
    '     Engine powers     -- Engine Powers rows with Effect "rent" add Value points per zone held, once',
    '                          the squad unlocks them (Unlock When, e.g. "zones: 2").',
    '     Summit Fever card -- a squad that plays it gets that week\'s territory total doubled. (2x is fixed.)',
    'Territory points only go out when Clash of Classes -> Turn on game awards is ON. Until then the map',
    'still shows who holds what, but sends no points.',
    'Example: GOLD holds the Cafeteria (Summit, 5), Music Wing (3) and Auditorium (3) = The Arts wing.',
    '     5 x 2 + 3 + 3 = 16, plus The Arts again (3 + 3) at WING_MULTIPLIER 2 = 22,',
    '     plus a +1-per-zone engine power x 3 zones = 25 points that week.',
    '',
    H + 'STAFF OVERRIDES: YOU ALWAYS HAVE THE LAST WORD',
    'Admin -> Map -> click a zone:',
    '- "Lock holder": the zone keeps whoever you pick in "Held by" until you unlock it.',
    '- "Override this month\'s numbers": type a squad\'s number to replace what was computed (a wrong',
    '  Climb-er count, a missed check-in, a tie you want to break). Clear the box to go back to the data.',
    '  Saved on the Zone Overrides tab (Month, Zone, Squad, Score, Note) with the date; the newest row wins.',
    'Everything else is editable too: rules and Multiplier (Map editor), Squad Monthly numbers, Event',
    'Check-ins (tick or untick Confirmed), Point Freezes, Daily Polls and Poll Ideas. Change a row, and the',
    'site uses it at the next check (about 10 minutes, or right away when you save in the Map editor).',
    '',
    H + 'SETTINGS TAB KEYS FOR THE MAP',
    '     GAME_MAP          -- OFF hides the map everywhere.',
    '     GAME_AWARDS       -- ON sends the weekly territory points (use the menu to turn it on or off).',
    '     CLAIM_DAYS        -- the window for zones and rules with no Window Days.',
    '     WING_MULTIPLIER   -- the wing bonus (above).',
    '     MAP_STYLE         -- puzzle (rough jigsaw cuts, the default), clean (plain areas) or mountain.',
    '',
    H + 'HOW TO...',
    '- Add a new way to conquer: add a row on Conquest Rules (or "+ Way to conquer" in the Map editor),',
    '  then pick it in a zone\'s Conquer By.',
    '- Rename a zone or move it: use the Map editor (renaming in the editor keeps the holder).',
    '- Connect two zones that do not touch: type the other zone\'s name in Borders.',
    '- Hand a zone to a squad yourself: set its rule to Staff Pick and type the squad in Holder.',
    '- Start over with the building layout: Clash of Classes -> Use the building map (holders start fresh;',
    '  your rules are kept).',
    '',
    H + 'KEEP IT SAFE',
    'The map shows areas of the building only (gyms, cafeteria, halls). Never add room numbers, doors,',
    'exits or anything that maps how to move through the building.',
  ];
  const values = rows.map(function (r) { return [r.indexOf(H) === 0 ? r.slice(H.length) : r]; });
  sheet.getRange(1, 1, values.length, 1).setValues(values);
  sheet.getRange(1, 1).setFontWeight('bold').setFontSize(14).setFontColor('#490e6f');
  rows.forEach(function (r, i) { if (r.indexOf(H) === 0) sheet.getRange(i + 1, 1).setFontWeight('bold').setFontColor('#490e6f'); });
  sheet.setColumnWidth(1, 900);
  sheet.setFrozenRows(1);
}
