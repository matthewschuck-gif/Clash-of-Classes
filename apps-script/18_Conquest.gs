/**
 * Clash of Classes -- Apps Script backend
 * File 18: Conquest. How territory on the school map changes hands, all editable without code.
 *
 * CONQUEST RULES tab: one row per way to win a zone. Map Zones › Conquer By names the rule.
 *   Method:      most points   | most entries (times a squad earned)
 *                check-ins     (students who checked in with their 800#: Poll Votes on polls whose Kind or
 *                               question has a Counts word, * = every poll: Pride Day, the Wheel, polls)
 *                school data   (one number per squad: attendance | grades | referral from the metrics
 *                               tab, or any Stat on the Squad Stats tab for that month: LRR + LD Students,
 *                               Interventions, Top Learners ...)
 *                staff pick    (staff type Holder)
 *   Counts:      event types or words in the event name, comma-separated (* = every point); the poll
 *                words for check-ins; the stat name for school data
 *   Window Days: how far back it looks (blank = CLAIM_DAYS). "last month" = the last calendar month,
 *                so a zone is won on last month's numbers and held all of this month. "this month" =
 *                the month so far (the Office: First Friday Coffee winners take it the day they win).
 *   Best Is:     Most (blank) or Least (fewest LRRs, fewest interventions)
 *   Per 100 Students: divide by squad size, so a bigger squad has no edge on counts
 *   Captured By, Multiplier: how the data is collected, and what that does to the zone's weekly pay
 *                (CAPTURE_MULTIPLIERS_ are the starting values; the Multiplier column wins)
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
 * MAP BATTLES tab: every capture, shown under the map on the site, HQ and the TV.
 * The admin panel's Map tab (MapEditor.html) edits zones and rules by dragging and picking.
 * HOW THE MAP WORKS tab: the staff guide to all of this, rewritten on every Run setup.
 */

const RULE_HEADERS_ = ['Show', 'Rule', 'How It Works', 'Method', 'Counts', 'Window Days', 'Must Border', 'Defense %', 'Far Attack %',
  'Best Is', 'Per 100 Students', 'Captured By', 'Multiplier'];
const BATTLE_HEADERS_ = ['Date', 'Zone', 'Taken By', 'From', 'Rule'];
const RULE_METHODS_ = ['most points', 'most entries', 'check-ins', 'school data', 'staff pick'];
const STATS_HEADERS_ = ['Month', 'Stat', 'BLACK', 'GOLD', 'GREY', 'PURPLE', 'Source', 'Updated'];
// How the data is captured sets the weekly pay. 800# check-ins are data the school never had
// before, so they pay double; school records are the hardest to move.
const CAPTURE_MULTIPLIERS_ = { '800# check-in': 2, 'school records': 1.5, 'clash points': 1, 'staff challenge': 1, 'staff pick': 1 };

// One rule per area of the building. Every one looks at the last calendar month.
// [Show, Rule, How It Works, Method, Counts, Window Days, Must Border, Defense %, Far Attack %, Best Is, Per 100 Students, Captured By, Multiplier]
const BUILDING_RULES_ = [
  [true, 'Pride Day', 'Most students who checked in on Mountaineer Pride Day last month (per 100 students).', 'check-ins', 'pride', 'last month', false, 0, '', 'Most', true, '800# check-in', 2],
  [true, 'Full Climb', 'Most total points last month, from anything.', 'most points', '*', 'last month', false, 0, '', 'Most', false, 'Clash points', 1],
  [true, 'Top Learners', 'Most monthly event winners (Students of the Month, top learners) last month.', 'school data', 'Top Learners', 'last month', false, 0, '', 'Most', true, 'School records', 1.5],
  [true, 'Clean Lunch', 'Fewest students with an LRR or a Lunch Detention last month (per 100 students).', 'school data', 'LRR + LD Students', 'last month', false, 0, '', 'Least', true, 'School records', 1.5],
  [true, 'Wheel & Polls', 'Most Spin Wheel and poll check-ins last month (per 100 students).', 'check-ins', '*', 'last month', false, 0, '', 'Most', true, '800# check-in', 2],
  [true, 'Attendance', 'Best attendance rate last month.', 'school data', 'attendance', 'last month', false, 0, '', 'Most', false, 'School records', 1.5],
  [true, 'Shout-outs', 'Most shout-outs last month (per 100 students).', 'most entries', 'shout', 'last month', false, 0, '', 'Most', true, 'Clash points', 1],
  [true, 'Growth Grind', 'Most Mindset summit and STAR growth points last month (per 100 students).', 'most points', 'star_growth, mountaineer_mindset', 'last month', false, 0, '', 'Most', true, 'School records', 1.5],
  [true, 'First Friday Coffee', 'The squad whose teachers and staff won this month\'s First Friday Coffee challenges.', 'school data', 'Staff Challenge Wins', 'this month', false, 0, '', 'Most', false, 'Staff challenge', 1],
  [true, 'Staff Pick', 'Staff decide who holds it.', 'staff pick', '', '', false, 0, '', 'Most', false, 'Staff pick', 1],
  [true, 'Passing Grades', 'Fewest failing grades: the most students with no core F.', 'school data', 'grades', 'last month', false, 0, '', 'Most', false, 'School records', 1.5],
  [true, 'Trait Dough', 'Most Trait Dough earned last month (per 100 students).', 'most points', 'trait_dough', 'last month', false, 0, '', 'Most', true, 'Clash points', 1],
  [true, 'Fewest Interventions', 'Fewest teacher and admin interventions combined last month (per 100 students).', 'school data', 'Interventions', 'last month', false, 0, '', 'Least', true, 'School records', 1.5],
];

// Areas of the building on a 24 x 16 grid, roughly where they sit on the school map.
// [Zone, Wing, Summit, Conquer By, Weekly Points, X, Y, W, H]
const BUILDING_ZONES_ = [
  ['Gym 2', 'Athletics', false, 'Pride Day', 3, 0, 0, 4, 10],
  ['Gym 1', 'Athletics', false, 'Top Learners', 3, 4, 0, 5, 8],
  ['Cafeteria', 'Heart of the School', true, 'Clean Lunch', 5, 4, 8, 5, 4],
  ['Music Wing', 'The Arts', false, 'Full Climb', 3, 0, 10, 4, 6],
  ['Auditorium', 'The Arts', false, 'Wheel & Polls', 3, 4, 12, 5, 4],
  ['50s Hall', 'North Halls', false, 'Attendance', 3, 9, 0, 6, 5],
  ['60s Hall', 'North Halls', false, 'Shout-outs', 3, 9, 5, 6, 4],
  ['Media Center', 'Heart of the School', false, 'Growth Grind', 3, 9, 9, 6, 3],
  ['Front Office', 'Heart of the School', false, 'First Friday Coffee', 2, 9, 12, 6, 3],
  ['70s Hall', 'East Halls', false, 'Passing Grades', 3, 15, 6, 6, 4],
  ['80s Hall', 'East Halls', false, 'Trait Dough', 3, 15, 10, 6, 5],
  ['90s Hall', 'East Halls', false, 'Fewest Interventions', 3, 21, 6, 3, 9],
];

function buildingZoneRows_() {
  return BUILDING_ZONES_.map(function (z) {
    return { 'Show': true, 'Zone': z[0], 'Wing': z[1], 'Summit': z[2], 'Conquer By': z[3], 'Weekly Points': z[4], 'X': z[5], 'Y': z[6], 'W': z[7], 'H': z[8] };
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
  [['Conquest Rules', RULE_HEADERS_, BUILDING_RULES_], ['Map Battles', BATTLE_HEADERS_, null]].forEach(function (t) {
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
    // A tab made before these columns existed: add them at the end.
    RULE_HEADERS_.slice(RULE_HEADERS_.indexOf('Far Attack %')).forEach(function (h) {
      if (head.indexOf(h) > -1) return;
      head.push(h);
      cr.getRange(1, head.length).setValue(h).setFontWeight('bold').setBackground('#490e6f').setFontColor('#ffffff');
    });
    try {
      cr.getRange(2, head.indexOf('Method') + 1, 200, 1)
        .setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(RULE_METHODS_, true).build());
    } catch (e) { console.warn('Conquest Rules dropdown: ' + e.message); }
  }
  if (!ss.getSheetByName('Squad Stats')) {
    const st = ss.insertSheet('Squad Stats');
    st.getRange(1, 1, 1, STATS_HEADERS_.length).setValues([STATS_HEADERS_]).setFontWeight('bold').setBackground('#490e6f').setFontColor('#ffffff');
    st.setFrozenRows(1);
  }
  const mz = ss.getSheetByName('Map Zones');
  if (mz && mz.getLastColumn()) {
    const head = mz.getRange(1, 1, 1, mz.getLastColumn()).getValues()[0].map(function (h) { return String(h).trim(); });
    ['Conquer By', 'Borders'].forEach(function (h) {
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
    out[name.toLowerCase()] = { rule: name, text: String(r['How It Works'] || ''), method: RULE_METHODS_.indexOf(method) > -1 ? method : 'most points',
      counts: String(r['Counts'] || ''), days: Number(r['Window Days']) || 0, lastMonth: /month/i.test(String(r['Window Days'] || '')) && !/this/i.test(String(r['Window Days'] || '')),
      thisMonth: /this\s*month/i.test(String(r['Window Days'] || '')),
      border: ticked_(r['Must Border']), defense: Math.max(0, Number(r['Defense %']) || 0),
      far: String(r['Far Attack %'] == null ? '' : r['Far Attack %']).trim() === '' ? null : Math.max(0, Number(r['Far Attack %']) || 0),
      least: /^least|^fewest|^lowest/i.test(String(r['Best Is'] || '').trim()), per100: ticked_(r['Per 100 Students']),
      captured: String(r['Captured By'] || '').trim(), mult: ruleMultiplier_(r['Multiplier'], r['Captured By']) };
  });
  return out;
}

// The Multiplier column, or the starting value for how the data is captured, or 1.
function ruleMultiplier_(v, captured) {
  const n = Number(v);
  if (String(v == null ? '' : v).trim() !== '' && n > 0) return Math.min(10, n);
  return CAPTURE_MULTIPLIERS_[String(captured || '').trim().toLowerCase()] || 1;
}

// The calendar month before today's: { from, to, month: 'yyyy-MM' }.
function lastMonthWindow_(today) {
  const y = Number(today.slice(0, 4)), m = Number(today.slice(5, 7));
  const py = m === 1 ? y - 1 : y, pm = m === 1 ? 12 : m - 1;
  const month = py + '-' + (pm < 10 ? '0' : '') + pm;
  const days = new Date(Date.UTC(py, pm, 0)).getUTCDate();
  return { from: month + '-01', to: month + '-' + days, month: month };
}

// Students per squad (Students tab, else the Student Board), for Per 100 Students.
function squadSizes_() {
  const out = { BLACK: 0, GOLD: 0, GREY: 0, PURPLE: 0 };
  const count = function (rows) { rows.forEach(function (r) { const s = squadOf_(r['Squad']); if (out[s] !== undefined) out[s]++; }); };
  count(sheetRows_('Students'));
  if (!SQUAD_KEYS_.some(function (s) { return out[s]; })) count(sheetRows_(getSetting_('STUDENT_BOARD_TAB') || 'Student Board'));
  return out;
}

// 800# check-ins (Poll Votes) between from and to on polls whose Kind or question has a word.
function checkInsBy_(from, to, spec) {
  const words = listOf_(spec), all = !words.length || words.indexOf('*') > -1;
  const polls = {};
  sheetRows_('Daily Polls').forEach(function (p) {
    const text = (String(p['Kind'] || '') + ' ' + String(p['Question'] || '')).toLowerCase();
    polls[String(p._row)] = all || words.some(function (w) { return text.indexOf(w) > -1; });
  });
  const out = { BLACK: 0, GOLD: 0, GREY: 0, PURPLE: 0 };
  sheetRows_('Poll Votes').forEach(function (v) {
    const d = String(v['Date'] || '').slice(0, 10), s = squadOf_(v['Squad']);
    if (d >= from && d <= to && out[s] !== undefined && polls[String(v['Poll'])]) out[s]++;
  });
  return out;
}

// One number per squad for a month: attendance | grades | referral (the metrics tab, Climb-er's
// newest numbers) or a Stat row on the Squad Stats tab. A squad with no number is null.
function schoolStat_(name, month) {
  const key = String(name || '').trim().toLowerCase();
  const out = { BLACK: null, GOLD: null, GREY: null, PURPLE: null };
  if (['attendance', 'grades', 'referral'].indexOf(key) > -1) {
    const m = readMetricsTab_();
    if (m && m[key]) SQUAD_KEYS_.forEach(function (s) { const v = Number(m[key][s]); if (v > 0) out[s] = v; });
    return out;
  }
  sheetRows_('Squad Stats').forEach(function (r) {
    if (String(r['Stat'] || '').trim().toLowerCase() !== key) return;
    let mo = r['Month'];
    mo = mo instanceof Date ? Utilities.formatDate(mo, Session.getScriptTimeZone(), 'yyyy-MM') : String(mo || '').trim().slice(0, 7);
    if (month && mo !== month) return;
    SQUAD_KEYS_.forEach(function (s) { if (String(r[s] == null ? '' : r[s]).trim() !== '') out[s] = Number(r[s]) || 0; });
  });
  return out;
}

// What each squad scored on a zone's rule. Counts may be divided by squad size (per 100 students).
function zoneScore_(z, events, today) {
  const win = z.lastMonth ? lastMonthWindow_(today) : z.thisMonth ? { from: today.slice(0, 8) + '01', to: today, month: today.slice(0, 7) }
    : { from: ymdShift_(today, -z.days + 1), to: today, month: '' };
  let score;
  if (z.method === 'school data') score = schoolStat_(z.claimedBy, win.month || today.slice(0, 7));
  else if (z.method === 'check-ins') score = checkInsBy_(win.from, win.to, z.claimedBy);
  else {
    const tally = pointsBy_(events, win.from, win.to, eventMatcher_(z.claimedBy));
    score = z.method === 'most entries' ? tally.rows : tally.points;
  }
  if (z.per100) {
    const size = squadSizes_();
    SQUAD_KEYS_.forEach(function (s) { if (score[s] != null && size[s]) score[s] = Math.round(1000 * score[s] / size[s]) / 10; });
  }
  return score;
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
      { rule: '', text: '', method: 'most points', counts: String(r['Claimed By'] || ''), days: 0, border: false, defense: 0, far: null, least: false, per100: false, captured: '', mult: 1 };
    let holder = squadOf_(r['Holder']);
    if (SQUAD_KEYS_.indexOf(holder) === -1) holder = '';
    return { row: r._row, zone: String(r['Zone']).trim(), wing: String(r['Wing'] || '').trim(), summit: ticked_(r['Summit']),
      rule: rule.rule, ruleText: rule.text, method: rule.method, claimedBy: rule.counts, days: rule.days || claimDays,
      border: rule.border, defense: rule.defense, far: rule.border ? rule.far : null, weekly: Number(r['Weekly Points']) || 0,
      lastMonth: !!rule.lastMonth, thisMonth: !!rule.thisMonth, least: !!rule.least, per100: !!rule.per100, captured: rule.captured || '', mult: rule.mult || 1,
      x: Math.max(0, Number(r['X']) || 0), y: Math.max(0, Number(r['Y']) || 0), w: Math.max(1, Number(r['W']) || 1), h: Math.max(1, Number(r['H']) || 1),
      borders: String(r['Borders'] || ''), holder: holder, since: String(r['Since'] || '') };
  });
  const nb = zoneNeighbors_(zones);
  // Borders are judged on who held what before this check, so the order of rows doesn't matter.
  const held = {}, count = { BLACK: 0, GOLD: 0, GREY: 0, PURPLE: 0 };
  zones.forEach(function (z) { held[z.zone] = z.holder; if (z.holder) count[z.holder]++; });
  const battles = [];
  const hCol = sheet ? headerCol_(sheet, 'Holder') : 0, sCol = sheet ? headerCol_(sheet, 'Since') : 0;
  zones.forEach(function (z) {
    z.score = z.method === 'staff pick' ? { BLACK: 0, GOLD: 0, GREY: 0, PURPLE: 0 } : zoneScore_(z, events, today);
    z.neighbors = nb[z.zone];
    if (z.method === 'staff pick') return;
    // Least (fewest LRRs): lower wins, 0 is the best score, a squad with no number sits out.
    const least = z.least;
    const has = function (s) { return z.score[s] != null && (least || z.score[s] > 0); };
    const beats = function (a, ref, edge) { return least ? a < ref * (1 - Math.min(edge, 100) / 100) : a > ref * (1 + edge / 100); };
    // Next door (or no border rule, or no zones yet): the normal bar. From farther away: the Far
    // Attack bar, the same for every squad (blank = only next door can attack).
    const nextDoor = function (s) { return !z.border || s === z.holder || !count[s] || z.neighbors.some(function (n) { return held[n] === s; }); };
    const qualifies = SQUAD_KEYS_.filter(function (s) {
      if (s === z.holder || !has(s)) return false;
      const near = nextDoor(s);
      if (!near && z.far === null) return false;
      // The bar is the holder's score; on an open zone, the best squad that could walk in.
      const others = SQUAD_KEYS_.filter(function (o) { return o !== s && nextDoor(o) && has(o); }).map(function (o) { return z.score[o]; });
      const ref = z.holder ? (has(z.holder) ? z.score[z.holder] : (least ? Infinity : 0))
        : (least ? Math.min.apply(null, others.concat([Infinity])) : Math.max.apply(null, others.concat([0])));
      const edge = near ? (z.holder ? z.defense : 0) : z.far;
      return beats(z.score[s], ref, edge);
    });
    const scores = qualifies.map(function (s) { return z.score[s]; });
    const best = least ? Math.min.apply(null, scores.concat([Infinity])) : Math.max.apply(null, scores.concat([0]));
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
      border: z.border, defense: z.defense, far: z.far, days: z.days, method: z.method, lastMonth: z.lastMonth, thisMonth: z.thisMonth, least: z.least,
      per100: z.per100, captured: z.captured, mult: z.mult, zones: [] };
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
      borders: String(r['Borders'] || ''), holder: squadOf_(r['Holder']), since: String(r['Since'] || '') };
  });
  const rules = sheetRows_('Conquest Rules').filter(function (r) { return String(r['Rule'] || '').trim(); }).map(function (r) {
    return { show: ticked_(r['Show']), rule: String(r['Rule']).trim(), text: String(r['How It Works'] || ''), method: String(r['Method'] || 'most points').toLowerCase(),
      counts: String(r['Counts'] || ''), days: windowCell_(r['Window Days']), border: ticked_(r['Must Border']), defense: Number(r['Defense %']) || 0,
      far: r['Far Attack %'] === '' || r['Far Attack %'] == null ? '' : Number(r['Far Attack %']) || 0,
      best: /^least|^fewest|^lowest/i.test(String(r['Best Is'] || '')) ? 'Least' : 'Most', per100: ticked_(r['Per 100 Students']),
      captured: String(r['Captured By'] || ''), mult: r['Multiplier'] === '' || r['Multiplier'] == null ? '' : Number(r['Multiplier']) || '' };
  });
  return JSON.stringify({ zones: zones, rules: rules, methods: RULE_METHODS_, claimDays: Number(getSetting_('CLAIM_DAYS')) || 14,
    building: { zones: buildingZoneRows_(), rules: BUILDING_RULES_.map(function (r) {
      return { show: r[0], rule: r[1], text: r[2], method: r[3], counts: r[4], days: r[5], border: r[6], defense: r[7], far: r[8],
        best: r[9], per100: r[10], captured: r[11], mult: r[12] }; }) } });
}

// Window Days as typed: '' | a number of days | 'last month'.
function windowCell_(v) {
  if (v === '' || v == null) return '';
  if (/this\s*month/i.test(String(v))) return 'this month';
  if (/month/i.test(String(v))) return 'last month';
  return Number(v) || '';
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
        'Counts': clip(r.counts, 200), 'Window Days': r.days === '' || r.days == null ? '' : /month/i.test(String(r.days)) ? windowCell_(r.days) : num(r.days, 1, 365),
        'Must Border': !!r.border, 'Defense %': num(r.defense, 0, 500),
        'Far Attack %': r.far === '' || r.far == null ? '' : num(r.far, 0, 1000),
        'Best Is': /^least/i.test(String(r.best || '')) ? 'Least' : 'Most', 'Per 100 Students': !!r.per100, 'Captured By': clip(r.captured, 40),
        'Multiplier': r.mult === '' || r.mult == null ? '' : Math.max(0, Math.min(10, Math.round(4 * (Number(r.mult) || 0)) / 4)) || '' };
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
        'Holder': holder, 'Since': since, 'Conquer By': clip(z.conquerBy, 60), 'Borders': clip(z.borders, 300) };
    });
    writeRowsByHeader_(getSheetSafe_('Conquest Rules'), rules);
    writeRowsByHeader_(getSheetSafe_('Map Zones'), zones);
    CacheService.getScriptCache().remove('game-tick');
    return JSON.stringify({ ok: true, zones: zones.length, rules: rules.length });
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
  writeRowsByHeader_(getSheetSafe_('Map Zones'), buildingZoneRows_());
  applyLocationRules_(); // the building's rules replace older rules with the same names; others stay
}

// Menu: put each building area on its own data (Pride Day, Clean Lunch, Attendance...). Rules with
// these names are replaced, the rest are kept; zones keep their place, size and holder.
function applyLocationRules_() {
  ensureGameTabs_();
  const sh = getSheetSafe_('Conquest Rules');
  const byName = {};
  BUILDING_RULES_.forEach(function (r) { byName[r[1].toLowerCase()] = r; });
  const kept = sheetRows_('Conquest Rules').filter(function (r) { return !byName[String(r['Rule'] || '').trim().toLowerCase()]; });
  const objs = BUILDING_RULES_.map(function (r) {
    const o = {}; RULE_HEADERS_.forEach(function (h, i) { o[h] = r[i]; }); return o;
  }).concat(kept.map(function (r) { const o = Object.assign({}, r); delete o._row; return o; }));
  writeRowsByHeader_(sh, objs);
  const zoneRule = {};
  BUILDING_ZONES_.forEach(function (z) { zoneRule[z[0].toLowerCase()] = z[3]; });
  const mz = getSheetSafe_('Map Zones'), col = headerCol_(mz, 'Conquer By');
  let n = 0;
  sheetRows_('Map Zones').forEach(function (r) {
    const rule = zoneRule[String(r['Zone'] || '').trim().toLowerCase()];
    if (rule && col) { mz.getRange(r._row, col).setValue(rule); n++; }
  });
  setupMapGuideTab_();
  CacheService.getScriptCache().remove('game-tick');
  return n;
}

function menuLocationRules_() {
  const ui = SpreadsheetApp.getUi();
  if (ui.alert('Give every area its own data?', 'Gym 2 = Pride Day check-ins, Gym 1 = top learners, Cafeteria = fewest LRR/LD students, Music = Full Climb, ' +
    'Auditorium = Wheel + poll check-ins, 50s = attendance, 60s = shout-outs, Media Center = Growth Grind, Office = First Friday Coffee staff challenge winners, 70s = fewest failing grades, ' +
    '80s = Trait Dough, 90s = fewest interventions. All judged on last calendar month. Zones keep their spot and holder; your other rules stay.', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  const n = applyLocationRules_();
  ui.alert('Done: ' + n + ' areas now have their own rule. See "How The Map Works" for the list and the multipliers.');
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
    '     X, Y, W, H    -- where the zone sits on the grid and how big it is (the Map editor sets these).',
    '     Borders       -- zones it connects to that do not touch on the map, comma-separated (a hallway).',
    '     Holder, Since -- filled in automatically. Only type a Holder on a Staff Pick zone.',
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
    '                      check-ins    = most students who checked in with their 800# (Poll Votes) on polls',
    '                                     whose Kind or question has a Counts word. pride = Pride Day',
    '                                     Check-in polls; * = every poll (Wheel Check-ins, School Pulse...).',
    '                      school data  = one number per squad. Counts = attendance, grades (% no core F) or',
    '                                     referral (% no office intervention) from the metrics tab, or a Stat',
    '                                     on the Squad Stats tab for that month (LRR + LD Students,',
    '                                     Interventions, Top Learners: Climb-er fills these; staff can type).',
    '                      staff pick   = never changes on its own. Type the squad in Holder on Map Zones',
    '                                     (or pick "Held by" in the Map editor).',
    '     Counts        -- which points count: event types or words in the event name, comma-separated.',
    '                      * = every point. Common types: trait_dough, daily_poll, star_growth,',
    '                      mountaineer_mindset, duel_bonus, mayhem_week, mayhem, shout, lunch.',
    '                      Game bonuses (territory, chance cards, etc.) and private behavior rows never count.',
    '     Window Days   -- how many days back it looks. Blank = CLAIM_DAYS on the Settings tab (14).',
    '                      "last month" = the last calendar month: on the 1st the zone goes to last',
    '                      month\'s winner and stays with them all month. The building rules use this,',
    '                      except the Office: "this month" = the month so far, so it changes hands the',
    '                      day the First Friday Coffee result goes in.',
    '     Best Is       -- Most, or Least (fewest LRR + LD students, fewest interventions). With Least, 0 is',
    '                      the best score and a squad with no number sits out.',
    '     Per 100 Students -- divide by squad size, so a bigger squad has no edge on counts.',
    '     Captured By   -- how the data is collected: 800# check-in, School records, Clash points, Staff pick.',
    '     Multiplier    -- what the zone\'s Weekly Points are multiplied by. Blank = from Captured By:',
    '                      800# check-in 2 (data the school never had before), School records 1.5 (the',
    '                      hardest to move), Clash points 1, Staff pick 1.',
    '     Must Border   -- the Risk rule. Tick it to make the zone a Siege zone (see below).',
    '     Defense %     -- the holder\'s edge: a challenger needs this much MORE than the holder. 0 = just more.',
    '     Far Attack %  -- Siege zones only: how much more a squad from farther away needs. Blank = only',
    '                      squads next door can attack.',
    'The holder always keeps a zone on a tie or when nobody scored (king of the hill).',
    'Claims are checked about every 10 minutes (and on the hourly game check).',
    '',
    H + 'THE BUILDING: WHAT WINS EACH AREA (LAST CALENDAR MONTH)',
    '     Gym 2 (old gym)  -- Pride Day: most Mountaineer Pride Day check-ins per 100 students    x2',
    '     Gym 1 (new gym)  -- Top Learners: most monthly event winners per 100 students          x1.5',
    '     Cafeteria        -- Clean Lunch: fewest students with an LRR or LD per 100 (the Summit) x1.5, Summit x2',
    '     Music Wing       -- Full Climb: most total points                                     x1',
    '     Auditorium       -- Wheel & Polls: most Spin Wheel and poll check-ins per 100          x2',
    '     50s Hall         -- Attendance: best attendance rate                                  x1.5',
    '     60s Hall         -- Shout-outs: most shout-outs per 100 students                      x1',
    '     Media Center     -- Growth Grind: most Mindset summit + STAR growth points per 100     x1.5',
    '     Front Office     -- First Friday Coffee: the squad whose adults won this month\'s staff  x1',
    '                         challenges. Type it on Squad Stats right after coffee: Month (yyyy-MM),',
    '                         Stat "Staff Challenge Wins", a number under each squad (wins or points).',
    '                         The winner takes the Office that day and keeps it until a squad beats it.',
    '     70s Hall         -- Passing Grades: fewest failing grades (most students with no core F) x1.5',
    '     80s Hall         -- Trait Dough: most Trait Dough earned per 100 students              x1',
    '     90s Hall         -- Fewest Interventions: teacher + admin combined, per 100            x1.5',
    'Pride Day and the Wheel need a check-in poll on Daily Polls that day (Kind "Pride Day Check-in" or',
    '"Wheel Check-in"; the Poll Ideas tab has both, ready to schedule).',
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
    '     Capture multiplier-- times its rule\'s Multiplier (800# check-in x2, School records x1.5, others x1).',
    '     The Summit        -- the zone with Summit ticked pays 2x its Weekly Points. (2x is fixed.)',
    '     Wing bonus        -- hold EVERY zone in a wing and that wing\'s weekly points pay again.',
    '                          Settings -> WING_MULTIPLIER (2 = the wing pays double in total, 3 = triple).',
    '     Engine powers     -- Engine Powers rows with Effect "rent" add Value points per zone held, once',
    '                          the squad unlocks them (Unlock When, e.g. "zones: 2").',
    '     Summit Fever card -- a squad that plays it gets that week\'s territory total doubled. (2x is fixed.)',
    'Territory points only go out when Clash of Classes -> Turn on game awards is ON. Until then the map',
    'still shows who holds what, but sends no points.',
    'Example: GOLD holds the Cafeteria (Summit, 5, x1.5), Music Wing (3, x1) and Auditorium (3, x2) = The Arts wing.',
    '     5 x 1.5 x 2 + 3 + 3 x 2 = 24, plus The Arts again (3 + 3) at WING_MULTIPLIER 2 = 30,',
    '     plus a +1-per-zone engine power x 3 zones = 33 points that week.',
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
