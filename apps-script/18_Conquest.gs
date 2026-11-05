/**
 * Clash of Classes -- Apps Script backend
 * File 18: Conquest. How territory on the school map changes hands, all editable without code.
 *
 * CONQUEST RULES tab: one row per way to win a zone. Map Zones › Conquer By names the rule.
 *   Method:      most points | most entries (times a squad earned) | staff pick (staff type Holder)
 *   Counts:      event types or words in the event name, comma-separated (* = every point)
 *   Window Days: how far back it looks (blank = CLAIM_DAYS)
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
 */

const RULE_HEADERS_ = ['Show', 'Rule', 'How It Works', 'Method', 'Counts', 'Window Days', 'Must Border', 'Defense %', 'Far Attack %'];
const BATTLE_HEADERS_ = ['Date', 'Zone', 'Taken By', 'From', 'Rule'];
const RULE_METHODS_ = ['most points', 'most entries', 'staff pick'];

const BUILDING_RULES_ = [
  [true, 'Full Climb', 'Every point your squad earns, anywhere, counts.', 'most points', '*', 14, false, 0, ''],
  [true, 'Game Day', 'Win Mayhem events, gaming and field days, and Squad Duels.', 'most points', 'mayhem, gaming, field, duel_bonus', 21, false, 0, ''],
  [true, 'Poll Turnout', 'Get your squad voting in the Daily Poll.', 'most points', 'daily_poll', 14, false, 0, ''],
  [true, 'Trait Dough Drive', 'Turn in the most Trait Dough.', 'most points', 'trait_dough', 14, false, 0, ''],
  [true, 'Growth Grind', 'STAR growth and Mountaineer Mindset summits.', 'most points', 'star_growth, mountaineer_mindset', 30, false, 0, ''],
  [true, 'Shout-out Storm', 'Recognize each other: the most shout-outs and Trait Dough turn-ins.', 'most entries', 'shout, trait_dough', 14, false, 0, ''],
  [true, 'Lunch Battles', 'Win the lunch challenges.', 'most points', 'lunch, mayhem_week', 21, false, 0, ''],
  [true, 'Siege', 'Risk rules: a squad next door needs 25% more points than the holder. A squad from farther away can still take it with 75% more.', 'most points', '*', 14, true, 25, 75],
  [true, 'Staff Pick', 'Staff decide who holds it.', 'staff pick', '', '', false, 0, ''],
];

// Areas of the building on a 24 x 16 grid, roughly where they sit on the school map.
// [Zone, Wing, Summit, Conquer By, Weekly Points, X, Y, W, H]
const BUILDING_ZONES_ = [
  ['Gym 2', 'Athletics', false, 'Game Day', 3, 0, 0, 4, 10],
  ['Gym 1', 'Athletics', false, 'Game Day', 3, 4, 0, 5, 8],
  ['Cafeteria', 'Heart of the School', true, 'Full Climb', 5, 4, 8, 5, 4],
  ['Music Wing', 'The Arts', false, 'Poll Turnout', 3, 0, 10, 4, 6],
  ['Auditorium', 'The Arts', false, 'Shout-out Storm', 3, 4, 12, 5, 4],
  ['50s Hall', 'North Halls', false, 'Trait Dough Drive', 3, 9, 0, 6, 5],
  ['60s Hall', 'North Halls', false, 'Siege', 3, 9, 5, 6, 4],
  ['Media Center', 'Heart of the School', false, 'Growth Grind', 3, 9, 9, 6, 3],
  ['Front Office', 'Heart of the School', false, 'Staff Pick', 2, 9, 12, 6, 3],
  ['70s Hall', 'East Halls', false, 'Lunch Battles', 3, 15, 6, 6, 4],
  ['80s Hall', 'East Halls', false, 'Siege', 3, 15, 10, 6, 5],
  ['90s Hall', 'East Halls', false, 'Trait Dough Drive', 3, 21, 6, 3, 9],
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
    if (head.indexOf('Far Attack %') === -1) cr.getRange(1, head.length + 1).setValue('Far Attack %').setFontWeight('bold').setBackground('#490e6f').setFontColor('#ffffff');
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
      counts: String(r['Counts'] || ''), days: Number(r['Window Days']) || 0, border: ticked_(r['Must Border']), defense: Math.max(0, Number(r['Defense %']) || 0),
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
      { rule: '', text: '', method: 'most points', counts: String(r['Claimed By'] || ''), days: 0, border: false, defense: 0, far: null };
    let holder = squadOf_(r['Holder']);
    if (SQUAD_KEYS_.indexOf(holder) === -1) holder = '';
    return { row: r._row, zone: String(r['Zone']).trim(), wing: String(r['Wing'] || '').trim(), summit: ticked_(r['Summit']),
      rule: rule.rule, ruleText: rule.text, method: rule.method, claimedBy: rule.counts, days: rule.days || claimDays,
      border: rule.border, defense: rule.defense, far: rule.border ? rule.far : null, weekly: Number(r['Weekly Points']) || 0,
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
    const tally = pointsBy_(events, ymdShift_(today, -z.days + 1), today, eventMatcher_(z.claimedBy));
    z.score = z.method === 'most entries' ? tally.rows : tally.points;
    z.neighbors = nb[z.zone];
    if (z.method === 'staff pick') return;
    // Next door (or no border rule, or no zones yet): the normal bar. From farther away: the Far
    // Attack bar, the same for every squad (blank = only next door can attack).
    const nextDoor = function (s) { return !z.border || s === z.holder || !count[s] || z.neighbors.some(function (n) { return held[n] === s; }); };
    const qualifies = SQUAD_KEYS_.filter(function (s) {
      if (s === z.holder || !(z.score[s] > 0)) return false;
      const near = nextDoor(s);
      if (!near && z.far === null) return false;
      // The bar is the holder's score; on an open zone, the best squad that could walk in.
      const ref = z.holder ? (z.score[z.holder] || 0)
        : Math.max.apply(null, SQUAD_KEYS_.filter(function (o) { return o !== s && nextDoor(o); }).map(function (o) { return z.score[o]; }).concat([0]));
      const edge = near ? (z.holder ? z.defense : 0) : z.far;
      return z.score[s] > ref * (1 + edge / 100);
    });
    const best = Math.max.apply(null, qualifies.map(function (s) { return z.score[s]; }).concat([0]));
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
      border: z.border, defense: z.defense, far: z.far, days: z.days, method: z.method, zones: [] };
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
      counts: String(r['Counts'] || ''), days: r['Window Days'] === '' ? '' : Number(r['Window Days']) || '', border: ticked_(r['Must Border']), defense: Number(r['Defense %']) || 0,
      far: r['Far Attack %'] === '' || r['Far Attack %'] == null ? '' : Number(r['Far Attack %']) || 0 };
  });
  return JSON.stringify({ zones: zones, rules: rules, methods: RULE_METHODS_, claimDays: Number(getSetting_('CLAIM_DAYS')) || 14,
    building: { zones: buildingZoneRows_(), rules: BUILDING_RULES_.map(function (r) {
      return { show: r[0], rule: r[1], text: r[2], method: r[3], counts: r[4], days: r[5], border: r[6], defense: r[7], far: r[8] }; }) } });
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
        'Counts': clip(r.counts, 200), 'Window Days': r.days === '' || r.days == null ? '' : num(r.days, 1, 365), 'Must Border': !!r.border, 'Defense %': num(r.defense, 0, 500),
        'Far Attack %': r.far === '' || r.far == null ? '' : num(r.far, 0, 1000) };
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
  const have = {};
  sheetRows_('Conquest Rules').forEach(function (r) { have[String(r['Rule'] || '').trim().toLowerCase()] = 1; });
  const sh = getSheetSafe_('Conquest Rules');
  BUILDING_RULES_.forEach(function (r) { if (!have[r[1].toLowerCase()]) sh.appendRow(r); });
  writeRowsByHeader_(getSheetSafe_('Map Zones'), buildingZoneRows_());
}
