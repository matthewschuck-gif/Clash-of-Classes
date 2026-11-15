/**
 * Clash of Classes -- Apps Script backend
 * File 50: Weather Days and Expeditions. Both run from Sheet tabs; nothing to turn on.
 *
 * WEATHER DAYS tab: Show | Date | Through | Name | Counts | Multiplier | Message | Paid
 *   A day (or a few) when some points count extra, e.g. "Blizzard Day", Counts trait_dough, 2.
 *   Counts = event types (comma separated); blank = every regular point. Through blank = one day.
 *   While it runs it shows on the home page, My Climb, Generals HQ and in the morning announcement.
 *   The day after it ends, the extra points are added once as one Weather row (Paid fills in).
 *
 * EXPEDITIONS tab: Show | Expedition | What To Do | Counts | Measure | Goal | Start | End | Reward | Winner | Won On
 *   A race: the FIRST squad to reach the Goal wins the Reward (two squads that get there the same
 *   day both win). Counts = event types (blank = all regular points); Measure = Points or Entries.
 *   Everyone sees how close each squad is. The site fills in Winner + Won On and adds the reward.
 */

const WEATHER_HEADERS_ = ['Show', 'Date', 'Through', 'Name', 'Counts', 'Multiplier', 'Message', 'Paid'];
const EXPEDITION_HEADERS_ = ['Show', 'Expedition', 'What To Do', 'Counts', 'Measure', 'Goal', 'Start', 'End', 'Reward', 'Winner', 'Won On'];
// Bonus rows never count toward a weather day, an expedition or a challenge.
const EXTRA_BONUS_TYPES_ = ['weather', 'expedition', 'colors_day', 'arcade', 'challenge_bonus', 'duel_bonus', 'helping_hands', 'squad_goal'];

function ensureWeatherTabs_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  [['Weather Days', WEATHER_HEADERS_, [false, todayYmd_(), '', 'Blizzard Day (example: tick Show and set a date)', 'trait_dough', 2, 'Trait Dough counts double today!', '']],
   ['Expeditions', EXPEDITION_HEADERS_, [false, 'Example: Service Summit', 'First squad to 100 service points wins', 'service', 'Points', 100, todayYmd_(), '', 25, '', '']]].forEach(function (t) {
    if (ss.getSheetByName(t[0])) return;
    const sh = ss.insertSheet(t[0]);
    sh.getRange(1, 1, 1, t[1].length).setValues([t[1]]).setFontWeight('bold').setBackground('#490e6f').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    sh.appendRow(t[2]);
    try { sh.getRange(2, 1, 200, 1).insertCheckboxes(); } catch (e) { /* fine */ }
  });
}

function typesOf_(v) { return String(v || '').split(',').map(function (s) { return s.trim(); }).filter(String); }
function countsToward_(e, types) {
  const t = String(e.type || '');
  if (EXTRA_BONUS_TYPES_.indexOf(t) > -1) return false;
  return !types.length || types.indexOf(t) > -1;
}

// ── Weather ──
function weatherRows_() {
  return sheetRows_('Weather Days').filter(function (r) { return ticked_(r['Show']) && asYmd_(r['Date']) && String(r['Name'] || '').trim(); }).map(function (r) {
    const from = asYmd_(r['Date']), to = asYmd_(r['Through']) || from;
    return { row: r._row, name: String(r['Name']).trim(), from: from, to: to < from ? from : to, types: typesOf_(r['Counts']),
      mult: Math.max(1, Math.min(5, Number(r['Multiplier']) || 2)), message: String(r['Message'] || '').trim(), paid: String(r['Paid'] || '').trim() };
  });
}

/** Today's weather (or null) and what's coming in the next 7 days. */
function weatherNow_() {
  const today = todayYmd_(), soon = ymdShift_(today, 7), rows = weatherRows_();
  const now = rows.filter(function (w) { return w.from <= today && w.to >= today; })[0] || null;
  const next = rows.filter(function (w) { return w.from > today && w.from <= soon; }).sort(function (a, b) { return a.from.localeCompare(b.from); })[0] || null;
  const pub = function (w) { return w && { name: w.name, from: w.from, to: w.to, mult: w.mult, message: w.message || (w.types.length ? w.types.join(', ') + ' count ' + w.mult + 'x' : 'Every point counts ' + w.mult + 'x') }; };
  return { now: pub(now), next: pub(next) };
}

/** Adds the extra points for weather days that have ended. Returns how many rows it added. */
function payWeather_() {
  const today = todayYmd_(), done = weatherRows_().filter(function (w) { return !w.paid && w.to < today; });
  if (!done.length) return 0;
  const events = readEventsTab_(), sh = getSheet_('Weather Days'), paidCol = headerCol_(sh, 'Paid');
  done.forEach(function (w) {
    const base = { BLACK: 0, GOLD: 0, GREY: 0, PURPLE: 0 };
    events.forEach(function (e) { const d = String(e.date || '').slice(0, 10); if (d >= w.from && d <= w.to && countsToward_(e, w.types)) SQUAD_KEYS_.forEach(function (s) { base[s] += Number(e.pts[s]) || 0; }); });
    const add = {}; SQUAD_KEYS_.forEach(function (s) { add[s] = Math.max(0, Math.round(base[s] * (w.mult - 1))); });
    if (SQUAD_KEYS_.some(function (s) { return add[s]; }))
      getSheet_('events').appendRow(['', w.name + ' (' + w.mult + 'x)', 'weather', w.to, w.message, '', add.BLACK || '', add.GOLD || '', add.GREY || '', add.PURPLE || '']);
    if (paidCol) sh.getRange(w.row, paidCol).setValue(today);
  });
  return done.length;
}

// ── Expeditions ──
function expeditions_(events) {
  events = events || readEventsTab_();
  const today = todayYmd_();
  return sheetRows_('Expeditions').filter(function (r) { return ticked_(r['Show']) && String(r['Expedition'] || '').trim() && Number(r['Goal']) > 0; }).map(function (r) {
    const types = typesOf_(r['Counts']), entries = /^entr/i.test(String(r['Measure'] || '')), goal = Number(r['Goal']);
    const start = asYmd_(r['Start']), end = asYmd_(r['End']);
    const mine = events.filter(function (e) { const d = String(e.date || '').slice(0, 10); return d && (!start || d >= start) && (!end || d <= end) && countsToward_(e, types); })
      .sort(function (a, b) { return String(a.date).localeCompare(String(b.date)); });
    const progress = { BLACK: 0, GOLD: 0, GREY: 0, PURPLE: 0 };
    let winners = [], wonOn = '';
    for (let i = 0; i < mine.length; i++) {
      const e = mine[i], d = String(e.date).slice(0, 10);
      if (winners.length && d > wonOn) break; // everyone who got there that same day wins
      SQUAD_KEYS_.forEach(function (s) { const p = Number(e.pts[s]) || 0; if (p > 0) progress[s] += entries ? 1 : p; });
      SQUAD_KEYS_.forEach(function (s) { if (progress[s] >= goal && winners.indexOf(s) === -1) { winners.push(s); wonOn = d; } });
    }
    const set = String(r['Winner'] || '').toUpperCase().split(/[,\s]+/).map(squadOf_).filter(function (s) { return SQUAD_KEYS_.indexOf(s) > -1; });
    if (set.length) { winners = set; wonOn = asYmd_(r['Won On']) || wonOn; }
    return { row: r._row, title: String(r['Expedition']).trim(), what: String(r['What To Do'] || '').trim(), goal: goal, entries: entries, start: start, end: end,
      reward: Number(r['Reward']) || 0, progress: progress, winners: winners, wonOn: wonOn, recorded: !!set.length,
      open: !winners.length && (!start || start <= today) && (!end || end >= today) };
  });
}

function expeditionsPublic_() {
  return expeditions_().filter(function (x) { return x.open || (x.winners.length && x.wonOn >= ymdShift_(todayYmd_(), -14)); }).map(function (x) {
    return { title: x.title, what: x.what, goal: x.goal, unit: x.entries ? 'entries' : 'points', end: x.end, reward: x.reward, progress: x.progress, winners: x.winners, wonOn: x.wonOn, open: x.open };
  });
}
function clashExpeditions() { return JSON.stringify(expeditionsPublic_()); }

/** Records new winners and adds their reward once. Returns how many it recorded. */
function payExpeditions_() {
  const fresh = expeditions_().filter(function (x) { return x.winners.length && !x.recorded; });
  if (!fresh.length) return 0;
  const sh = getSheet_('Expeditions'), wc = headerCol_(sh, 'Winner'), dc = headerCol_(sh, 'Won On');
  fresh.forEach(function (x) {
    const pts = { BLACK: '', GOLD: '', GREY: '', PURPLE: '' };
    x.winners.forEach(function (s) { pts[s] = x.reward; });
    if (x.reward) getSheet_('events').appendRow(['', 'Expedition won: ' + x.title, 'expedition', x.wonOn || todayYmd_(), x.winners.map(titleSquad_).join(' + ') + ' reached ' + x.goal + ' first', '', pts.BLACK, pts.GOLD, pts.GREY, pts.PURPLE]);
    if (wc) sh.getRange(x.row, wc).setValue(x.winners.join(', '));
    if (dc) sh.getRange(x.row, dc).setValue(x.wonOn || todayYmd_());
  });
  return fresh.length;
}

// Once a day from the background tick: weather, expeditions, colors days, Family Pick.
function maybeDailyExtras_() {
  const props = PropertiesService.getScriptProperties(), today = todayYmd_();
  if (props.getProperty('extras-paid') === today) return;
  props.setProperty('extras-paid', today);
  try { payWeather_(); } catch (e) { noteError_('weather', e); }
  try { payExpeditions_(); } catch (e) { noteError_('expeditions', e); }
  try { payColorsDays_(); } catch (e) { noteError_('colors day', e); }
  try { payFamilyPick_(); } catch (e) { noteError_('family pick', e); }
  dailyLeagues_(); // 57_Leagues.gs: Helping Hands, squad goals, planned events, Rivalry Week, Team League
}

/** What the home page, My Climb and the TV show about weather, expeditions and colors days. */
function homeExtras_() {
  const out = {};
  try { out.weather = weatherNow_(); } catch (e) { out.weather = null; }
  try { out.expeditions = expeditionsPublic_(); } catch (e) { out.expeditions = []; }
  try { out.colors = colorsDayNext_(); } catch (e) { out.colors = null; }
  try { out.hall = hallYears_().length > 0; } catch (e) { out.hall = false; }
  try { out.hands = helpingHandsNow_(); } catch (e) { out.hands = null; } // 54_Rally.gs
  try { out.banners = squadBanners_(); } catch (e) { out.banners = null; }
  try { out.shoutouts = shoutoutsPublic_(6); } catch (e) { out.shoutouts = []; } // 56_Shoutouts.gs
  try { out.planned = plannedSoon_(); } catch (e) { out.planned = []; }
  try { out.rivalry = rivalryNow_(); } catch (e) { out.rivalry = null; } // 57_Leagues.gs
  return out;
}

/** Extra lines for the morning announcement (22_Week.gs). */
function announceExtras_() {
  const lines = [], x = homeExtras_(), today = todayYmd_();
  if (x.weather && x.weather.now) lines.push('Today is ' + x.weather.now.name + '! ' + x.weather.now.message);
  else if (x.weather && x.weather.next && x.weather.next.from === ymdShift_(today, 1)) lines.push('Tomorrow is ' + x.weather.next.name + '. ' + x.weather.next.message);
  (x.expeditions || []).forEach(function (e) {
    if (!e.open && e.wonOn >= ymdShift_(today, -3)) { lines.push(e.winners.map(titleSquad_).join(' and ') + ' won the ' + e.title + ' expedition' + (e.reward ? ' for ' + e.reward + ' points' : '') + '!'); return; }
    if (!e.open) return;
    const lead = SQUAD_KEYS_.slice().sort(function (a, b) { return e.progress[b] - e.progress[a]; })[0];
    if (e.progress[lead] > 0) lines.push('Expedition ' + e.title + ': ' + titleSquad_(lead) + ' leads with ' + e.progress[lead] + ' of ' + e.goal + '.');
  });
  if (x.colors && (x.colors.date === today || x.colors.date === ymdShift_(today, 1))) lines.push((x.colors.date === today ? 'Today' : 'Tomorrow') + ' is ' + x.colors.name + '. Wear your squad colors!');
  if (x.hands && x.hands.squads.length && new Date(today + 'T12:00:00Z').getUTCDay() === 1) lines.push('Helping Hands this week: ' + x.hands.squads.map(titleSquad_).join(' and ') + '\'s points count ' + x.hands.mult + ' times. Every point matters!');
  (x.planned || []).forEach(function (p) { if (p.date === today || p.date === ymdShift_(today, 1)) lines.push((p.date === today ? 'Today' : 'Tomorrow') + ': ' + p.name + '. ' + p.what); });
  if (x.rivalry) (x.rivalry.live || []).forEach(function (r) { const a = r.squads[0], b = r.squads[1]; if (a && b) lines.push(r.title + ': ' + titleSquad_(a) + ' ' + r.scores[a] + ', ' + titleSquad_(b) + ' ' + r.scores[b] + '.'); });
  const so = (x.shoutouts || []).filter(function (s) { return s.date >= ymdShift_(today, -1); })[0];
  if (so) lines.push('Shout-out to ' + so.who + ' from ' + so.from + ': ' + so.message);
  try { if (today.slice(8) <= '03') { const ym = monthBefore_(today.slice(0, 7)), won = sheetRows_('Team League').filter(function (r) { return monthKey_(r['Month']) === ym && SQUAD_KEYS_.indexOf(squadOf_(r['Squad'])) > -1; }); if (won.length) lines.push('Team League winners for last month: ' + won.map(function (r) { return 'Team ' + r['Team'] + ' in ' + titleSquad_(squadOf_(r['Squad'])); }).join(', ') + '!'); } } catch (e) { /* none */ }
  try { const sp = spotlightPublic_(1)[0]; if (sp && sp.week >= mondayOf_(today)) lines.push('Squad Spotlight: ' + sp.who + ' from ' + titleSquad_(sp.squad) + '. ' + sp.story); } catch (e) { /* none */ }
  return lines;
}

// file-stamp (tools/stamp-files.js): lets Admin › System tell whether this file is the newest one
var FILE_STAMP_50_Weather = '7fb56ba6da';
