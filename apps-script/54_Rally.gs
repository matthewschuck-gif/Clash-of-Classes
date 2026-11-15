/**
 * Clash of Classes -- Apps Script backend
 * File 54: Helping Hands, Squad Goals and Squad Banners.
 *
 * HELPING HANDS (Settings › HELPING_HANDS ON, HELPING_HANDS_MULT 1.25)
 *   The squad with the fewest points last week (Monday to Sunday) gets Helping Hands this week:
 *   its points count HELPING_HANDS_MULT times. A tie for last means every tied squad gets it; a
 *   week where nobody scored gives it to no one. The extra is added the Monday after, as one
 *   "Helping Hands" event (type helping_hands). Bonuses and game points never count toward it.
 *   Helping Hands tab: Week | Squad | Last Week | Earned | Bonus | Paid (fills in by itself).
 *
 * SQUAD GOALS (Settings › SQUAD_GOALS ON, SQUAD_GOAL_BONUS 15)
 *   A General (on Generals HQ) or a captain (on My Climb) sets one goal for the squad each week,
 *   Monday to Wednesday: what counts (squad points, service points, Trait Dough bills, poll votes,
 *   Arcade games) and a number that beats last week's. Hit it by Sunday and the squad gets
 *   SQUAD_GOAL_BONUS, added the Monday after (type squad_goal).
 *   Squad Goals tab: Week | Squad | Measure | Target | Note | Set By | Set At | Result | Bonus | Paid.
 *
 * SQUAD BANNERS (always on; nothing shows until staff approve)
 *   A General or captain picks the squad's motto, a banner pattern and an icon. Staff tick
 *   Approved (Sheet or Admin › System). The newest approved one shows on My Climb, Generals HQ,
 *   the home page, and goes into the Hall of Fame at rollover.
 *   Squad Banners tab: Submitted | Squad | Motto | Pattern | Icon | By | Approved.
 */

const HANDS_HEADERS_ = ['Week', 'Squad', 'Last Week', 'Earned', 'Bonus', 'Paid'];
const GOAL_HEADERS_ = ['Week', 'Squad', 'Measure', 'Target', 'Note', 'Set By', 'Set At', 'Result', 'Bonus', 'Paid'];
const BANNER_HEADERS_ = ['Submitted', 'Squad', 'Motto', 'Pattern', 'Icon', 'By', 'Approved'];
const GOAL_MEASURES_ = { points: 'squad points', service: 'service points', dough: 'Trait Dough bills', votes: 'Daily Poll votes', arcade: 'Arcade games' };
const BANNER_PATTERNS_ = ['peaks', 'stripes', 'stars', 'waves', 'chevrons'];
const BANNER_ICONS_ = ['🏔️', '🦅', '🐺', '🦁', '🐻', '🔥', '⚡', '⭐', '🌲', '🧭', '🛡️', '🌊'];

function ensureRallyTabs_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  [['Helping Hands', HANDS_HEADERS_], ['Squad Goals', GOAL_HEADERS_], ['Squad Banners', BANNER_HEADERS_]].forEach(function (t) {
    if (ss.getSheetByName(t[0])) return;
    const sh = ss.insertSheet(t[0]);
    sh.getRange(1, 1, 1, t[1].length).setValues([t[1]]).setFontWeight('bold').setBackground('#490e6f').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  });
}

function settingOn_(key, def) { return String(getSetting_(key) || def || 'OFF').trim().toUpperCase() === 'ON'; }

// Points a squad earned between two days, without bonuses or game points.
function earnedPoints_(events, from, to) {
  const out = { BLACK: 0, GOLD: 0, GREY: 0, PURPLE: 0 };
  events.forEach(function (e) {
    const d = String(e.date || '').slice(0, 10), t = String(e.type || '');
    if (d < from || d > to || BONUS_TYPES_.indexOf(t) > -1 || GAME_TYPES_.indexOf(t) > -1) return;
    SQUAD_KEYS_.forEach(function (s) { out[s] += Math.max(0, Number(e.pts[s]) || 0); });
  });
  return out;
}

// ── Helping Hands ──
function handsMult_() { const m = settingNum_('HELPING_HANDS_MULT', 1.25); return m > 1 && m <= 2 ? m : 1.25; }

/** Which squads have Helping Hands for the week starting mon. write = record it on the tab (daily job). */
function handsFor_(mon, events, write) {
  const rows = sheetRows_('Helping Hands').filter(function (r) { return asYmd_(r['Week']).slice(0, 10) === mon; });
  if (rows.length) return rows.map(function (r) { return squadOf_(r['Squad']); }).filter(function (s) { return SQUAD_KEYS_.indexOf(s) > -1; });
  if (!settingOn_('HELPING_HANDS', 'ON') || mon > todayYmd_()) return [];
  const last = earnedPoints_(events || readEventsTab_(), ymdShift_(mon, -7), ymdShift_(mon, -1));
  const vals = SQUAD_KEYS_.map(function (s) { return last[s]; });
  const low = Math.min.apply(null, vals), high = Math.max.apply(null, vals);
  const squads = high <= 0 || low === high ? [] : SQUAD_KEYS_.filter(function (s) { return last[s] === low; });
  if (write) {
    const sh = getSheet_('Helping Hands');
    if (!squads.length) sh.appendRow([mon, 'NONE', '', '', '', '']);
    squads.forEach(function (s) { sh.appendRow([mon, s, last[s], '', '', '']); });
  }
  return squads;
}

/** { squads, mult, week } for this week (empty squads when off). */
function helpingHandsNow_() {
  if (!settingOn_('HELPING_HANDS', 'ON')) return { squads: [], mult: handsMult_(), week: '' };
  ensureRallyTabs_();
  const mon = mondayOf_(todayYmd_());
  return { squads: handsFor_(mon), mult: handsMult_(), week: mon };
}

function payHelpingHands_() {
  if (!settingOn_('HELPING_HANDS', 'ON')) return 0;
  ensureRallyTabs_();
  const mon = mondayOf_(todayYmd_()), events = readEventsTab_(), sh = getSheet_('Helping Hands'), mult = handsMult_();
  handsFor_(mon, events, true); // record this week's squads (first run of the week)
  let paid = 0;
  sheetRows_('Helping Hands').forEach(function (r) {
    const week = asYmd_(r['Week']).slice(0, 10), s = squadOf_(r['Squad']);
    if (!week || week >= mon || String(r['Paid'] || '').trim() || SQUAD_KEYS_.indexOf(s) === -1) return;
    const earned = earnedPoints_(events, week, ymdShift_(week, 6))[s], bonus = Math.round(earned * (mult - 1));
    const pts = { BLACK: '', GOLD: '', GREY: '', PURPLE: '' };
    pts[s] = bonus;
    if (bonus > 0) getSheet_('events').appendRow(['', 'Helping Hands: ' + titleSquad_(s) + ' (' + mult + 'x)', 'helping_hands', ymdShift_(week, 6),
      earned + ' points the week of ' + week, '', pts.BLACK, pts.GOLD, pts.GREY, pts.PURPLE]);
    sh.getRange(r._row, 4, 1, 3).setValues([[earned, bonus, todayYmd_()]]);
    paid++;
  });
  return paid;
}

// ── Squad Goals ──
function goalValue_(measure, squad, from, to, events) {
  const inside = function (d) { d = String(d || '').slice(0, 10); return d >= from && d <= to; };
  if (measure === 'votes') return sheetRows_('Poll Votes').filter(function (v) { return squadOf_(v['Squad']) === squad && inside(asYmd_(v['Date'] || v['At'])); }).length;
  if (measure === 'arcade') return sheetRows_('Arcade Plays').filter(function (p) { return squadOf_(p['Squad']) === squad && inside(p['When']); }).length;
  events = events || readEventsTab_();
  if (measure === 'points') return earnedPoints_(events, from, to)[squad];
  const type = measure === 'service' ? 'service' : 'trait_dough';
  return events.filter(function (e) { return String(e.type) === type && inside(e.date) && (Number(e.pts[squad]) || 0) > 0; })
    .reduce(function (n, e) { return n + (measure === 'service' ? Number(e.pts[squad]) || 0 : 1); }, 0);
}

/** This week's goal for a squad, with progress, plus what a new goal must beat. */
function squadGoalFor_(squad) {
  if (!settingOn_('SQUAD_GOALS', 'ON')) return null;
  ensureRallyTabs_();
  const today = todayYmd_(), mon = mondayOf_(today), events = readEventsTab_();
  const row = sheetRows_('Squad Goals').filter(function (r) { return asYmd_(r['Week']).slice(0, 10) === mon && squadOf_(r['Squad']) === squad; }).pop();
  const beat = {};
  Object.keys(GOAL_MEASURES_).forEach(function (m) { beat[m] = goalValue_(m, squad, ymdShift_(mon, -7), ymdShift_(mon, -1), events); });
  const dow = new Date(today + 'T12:00:00Z').getUTCDay(), open = dow >= 1 && dow <= 3;
  const out = { week: mon, open: open, measures: GOAL_MEASURES_, beat: beat, bonus: settingNum_('SQUAD_GOAL_BONUS', 15), goal: null };
  if (row) {
    const m = String(row['Measure']);
    out.goal = { measure: m, label: GOAL_MEASURES_[m] || m, target: Number(row['Target']) || 0, note: String(row['Note'] || ''), by: String(row['Set By'] || ''),
      progress: goalValue_(m, squad, mon, ymdShift_(mon, 6), events) };
  }
  return out;
}

function setSquadGoal_(squad, by, measure, target, note) {
  if (!settingOn_('SQUAD_GOALS', 'ON')) throw new Error('Squad goals are off right now.');
  if (!GOAL_MEASURES_[measure]) throw new Error('Pick what counts.');
  const g = squadGoalFor_(squad);
  if (!g.open) throw new Error('Goals are set Monday to Wednesday. Set next week\'s on Monday.');
  target = Math.floor(Number(target) || 0);
  const min = Math.max(1, g.beat[measure] + 1);
  if (target < min) throw new Error('A goal has to beat last week: at least ' + min + ' ' + GOAL_MEASURES_[measure] + '.');
  if (target > 100000) throw new Error('That number is too big.');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = getSheet_('Squad Goals');
    const old = sheetRows_('Squad Goals').filter(function (r) { return asYmd_(r['Week']).slice(0, 10) === g.week && squadOf_(r['Squad']) === squad; }).pop();
    const row = [g.week, squad, measure, target, String(note || '').replace(/\s+/g, ' ').trim().slice(0, 120), by,
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'), '', '', ''];
    if (old) sh.getRange(old._row, 1, 1, row.length).setValues([row]); else sh.appendRow(row);
  } finally { lock.releaseLock(); }
  return squadGoalFor_(squad);
}

function paySquadGoals_() {
  if (!settingOn_('SQUAD_GOALS', 'ON')) return 0;
  ensureRallyTabs_();
  const mon = mondayOf_(todayYmd_()), events = readEventsTab_(), sh = getSheet_('Squad Goals'), bonus = settingNum_('SQUAD_GOAL_BONUS', 15);
  let n = 0;
  sheetRows_('Squad Goals').forEach(function (r) {
    const week = asYmd_(r['Week']).slice(0, 10), s = squadOf_(r['Squad']), m = String(r['Measure']);
    if (!week || week >= mon || String(r['Paid'] || '').trim() || SQUAD_KEYS_.indexOf(s) === -1 || !GOAL_MEASURES_[m]) return;
    const got = goalValue_(m, s, week, ymdShift_(week, 6), events), hit = got >= (Number(r['Target']) || 0);
    if (hit && bonus > 0) {
      const pts = { BLACK: '', GOLD: '', GREY: '', PURPLE: '' }; pts[s] = bonus;
      getSheet_('events').appendRow(['', 'Squad Goal: ' + titleSquad_(s) + ' hit ' + r['Target'] + ' ' + GOAL_MEASURES_[m], 'squad_goal', ymdShift_(week, 6), String(r['Note'] || ''), '', pts.BLACK, pts.GOLD, pts.GREY, pts.PURPLE]);
    }
    sh.getRange(r._row, 8, 1, 3).setValues([[got + (hit ? ' (hit)' : ' (missed)'), hit ? bonus : 0, todayYmd_()]]);
    n++;
  });
  return n;
}

// Who on My Climb may set goals and banners: Generals and captains on the MLC Roster.
function squadLeader_(id) {
  const r = sheetRows_('MLC Roster').filter(function (x) { return id800_(x['800#']) === id && /general|captain/i.test(String(x['Role'])); })[0];
  return r ? { id: id, name: String(r['Name'] || ''), squad: squadOf_(r['Squad']), role: String(r['Role']) } : null;
}

// ── Squad Banners ──
function bannerOf_(r) {
  const p = String(r['Pattern'] || '').toLowerCase(), icon = String(r['Icon'] || '');
  return { squad: squadOf_(r['Squad']), motto: String(r['Motto'] || '').trim(), pattern: BANNER_PATTERNS_.indexOf(p) > -1 ? p : 'peaks',
    icon: BANNER_ICONS_.indexOf(icon) > -1 ? icon : '🏔️' };
}

/** { BLACK: banner or null, ... } (newest approved). */
function squadBanners_() {
  const out = { BLACK: null, GOLD: null, GREY: null, PURPLE: null };
  sheetRows_('Squad Banners').forEach(function (r) { const s = squadOf_(r['Squad']); if (out[s] !== undefined && ticked_(r['Approved']) && String(r['Motto'] || '').trim()) out[s] = bannerOf_(r); });
  return out;
}

/** For the squad's leaders: the live banner, the one waiting, and the choices. */
function bannerFor_(squad) {
  ensureRallyTabs_();
  const rows = sheetRows_('Squad Banners').filter(function (r) { return squadOf_(r['Squad']) === squad; });
  const last = rows[rows.length - 1];
  return { live: squadBanners_()[squad], waiting: last && !ticked_(last['Approved']) ? bannerOf_(last) : null, patterns: BANNER_PATTERNS_, icons: BANNER_ICONS_ };
}

function saveBanner_(squad, by, motto, pattern, icon) {
  motto = String(motto || '').replace(/\s+/g, ' ').trim();
  if (motto.length < 3 || motto.length > 60) throw new Error('A motto is 3 to 60 characters.');
  if (BANNER_PATTERNS_.indexOf(pattern) === -1) throw new Error('Pick a pattern.');
  if (BANNER_ICONS_.indexOf(icon) === -1) throw new Error('Pick an icon.');
  ensureRallyTabs_();
  const sh = getSheet_('Squad Banners');
  const rows = sheetRows_('Squad Banners').filter(function (r) { return squadOf_(r['Squad']) === squad; });
  const last = rows[rows.length - 1];
  const row = [Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'), squad, motto, pattern, icon, by, false];
  // One waiting at a time: a new idea replaces the one staff haven't looked at yet.
  if (last && !ticked_(last['Approved'])) sh.getRange(last._row, 1, 1, row.length).setValues([row]); else sh.appendRow(row);
  return bannerFor_(squad);
}

// ── Calls from Generals HQ ──
function clashHqGoal(token, measure, target, note) { const me = hqSession_(token); setSquadGoal_(me.squad, me.name || me.id, measure, target, note); return clashGeneralsHq(token); }
function clashHqBanner(token, motto, pattern, icon) { const me = hqSession_(token); saveBanner_(me.squad, me.name || me.id, motto, pattern, icon); return clashGeneralsHq(token); }

// ── Calls from My Climb (captains and Generals) ──
function leaderFromMe_(asId, adminToken) {
  const id = meId_(asId, adminToken);
  if (id === SAMPLE_ID_) return { id: id, name: 'Sample captain', squad: sampleSquad_(asId) || 'GOLD', role: 'Captain', sample: true };
  return squadLeader_(id);
}
function clashMeGoal(measure, target, note, asId, adminToken) {
  return withSample_(asId, function () {
    const L = leaderFromMe_(asId, adminToken);
    if (!L) throw new Error('Only Generals and captains set squad goals.');
    if (!L.sample) setSquadGoal_(L.squad, L.name || L.id, measure, target, note);
    return clashMeExtras(asId, adminToken);
  });
}
function clashMeBanner(motto, pattern, icon, asId, adminToken) {
  return withSample_(asId, function () {
    const L = leaderFromMe_(asId, adminToken);
    if (!L) throw new Error('Only Generals and captains choose the banner.');
    if (!L.sample) saveBanner_(L.squad, L.name || L.id, motto, pattern, icon);
    return clashMeExtras(asId, adminToken);
  });
}

// file-stamp (tools/stamp-files.js): lets Admin › System tell whether this file is the newest one
var FILE_STAMP_54_Rally = '';

// file-stamp (tools/stamp-files.js): lets Admin › System tell whether this file is the newest one
var FILE_STAMP_54_Rally = 'deccc5da46';
