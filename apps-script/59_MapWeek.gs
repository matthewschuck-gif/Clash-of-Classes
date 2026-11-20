/**
 * Clash of Classes -- Apps Script backend
 * File 59: the map turns every Monday at 7:00 AM.
 *
 * Settings › MAP_CYCLE weekly (the default) | monthly (the old way).
 * Weekly: every area, whatever its rule, is judged on LAST WEEK (Monday to Sunday), and holders
 * change only once, in the Monday 7:00 turn. In between, the map shows this week's race but
 * nobody gains or loses an area. Rules still decide HOW each area is won (most points, highest /
 * lowest number, per 100 students, Must Border, Defense %); only the window changed.
 *   - Points, entries, Pulse votes (approved polls), check-ins, squad events: last week's.
 *   - The Cafeteria's LRRs + LDs: counted for last week from the Follow-Up sheets (27).
 *   - Numbers that only come monthly (attendance, failing grades, interventions, Pride Day from
 *     Squad Monthly): the newest month on file. Those areas change hands when new numbers arrive.
 * The turn: a time trigger Monday at 7 (made by Run setup), and the background tick catches it
 * up if the trigger is missing (any time Monday from 7:00). Once a week, never twice. The weekly
 * pay for holding areas (15_Game.gs) is figured before the turn, for the week that just ended.
 * Map Battles shows every change with the week it was judged on.
 */

function mapWeekly_() { return String(getSetting_('MAP_CYCLE') || 'weekly').trim().toLowerCase() !== 'monthly'; }

// [Monday, Sunday] of the week before the one ymd is in.
function lastWeekBounds_(ymd) {
  const mon = mondayOf_(ymd || todayYmd_()); // 41_Morning.gs
  return [ymdShift_(mon, -7), ymdShift_(mon, -1)];
}

// The newest Squad Monthly month (at or before ym) where every squad has measure n.
function newestMonthlyWith_(n, ym) {
  const months = {};
  sheetRows_(MONTHLY_TAB_).forEach(function (r) { const m = ymOf_(r['Month']); if (m && m <= ym) months[m] = true; });
  const list = Object.keys(months).sort().reverse();
  for (let i = 0; i < list.length; i++) {
    const monthly = squadMonthly_(list[i]);
    if (SQUAD_KEYS_.every(function (s) { return monthly[s] && monthly[s][n] !== undefined; })) return { ym: list[i], monthly: monthly };
  }
  return null;
}

// LRR + LD students per squad for one week, from the Follow-Up sheets. Kept a day per week.
function weeklyLrrCounts_(from, to) {
  if (!getSetting_('LRR_SHEET')) return null;
  const key = 'lrr-week-' + from, cache = CacheService.getScriptCache(), hit = cache.get(key);
  if (hit) return JSON.parse(hit);
  try {
    const c = followUpLrrCounts_(from.slice(0, 7), [from, to]).counts;
    cache.put(key, JSON.stringify(c), 21600);
    return c;
  } catch (e) { console.warn('Weekly LRR count: ' + e.message); return null; }
}

/** The Monday turn: judge last week and move holders. Returns a line. force = run again this week. */
function weeklyMapTurn_(force) {
  if (!mapWeekly_()) return 'MAP_CYCLE is monthly: the map changes as numbers come in.';
  const props = PropertiesService.getScriptProperties(), mon = mondayOf_(todayYmd_());
  if (!force && props.getProperty('map-turn-week') === mon) return 'The map already turned this week (' + mon + ').';
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const before = sheetRows_('Map Battles').length;
    mapZones_(null, true, { turn: true });
    props.setProperty('map-turn-week', mon);
    const moved = sheetRows_('Map Battles').slice(before);
    try { CacheService.getScriptCache().remove('game-tick'); } catch (e) { /* fine */ }
    return 'Map turned for the week of ' + lastWeekBounds_(todayYmd_())[0] + ': ' +
      (moved.length ? moved.map(function (b) { return b['Zone'] + ' → ' + titleSquad_(squadOf_(b['Taken By'])); }).join(', ') : 'no areas changed hands') + '.';
  } finally {
    lock.releaseLock();
  }
}

// The trigger's handler (Mondays at 7).
function weeklyMapTurn() { noteRun_('weeklyMapTurn'); return weeklyMapTurn_(false); }

// From backgroundTick: Monday from 7:00, if the trigger hasn't run yet this week.
function maybeWeeklyMapTurn_() {
  if (!mapWeekly_()) return;
  const now = new Date(), tz = Session.getScriptTimeZone();
  if (Utilities.formatDate(now, tz, 'u') !== '1' || Utilities.formatDate(now, tz, 'HH:mm') < '07:00') return;
  weeklyMapTurn_(false);
}

function installWeeklyMapTrigger_() {
  try {
    ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'weeklyMapTurn') ScriptApp.deleteTrigger(t); });
    ScriptApp.newTrigger('weeklyMapTurn').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(7).nearMinute(0).create();
  } catch (e) { console.warn('Weekly map trigger: ' + e.message); }
}

function menuWeeklyMapTurn_() {
  const ui = SpreadsheetApp.getUi();
  if (ui.alert('Turn the map now?', 'This judges last week and moves areas now, even if the map already turned this week.', ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;
  ui.alert('Map', weeklyMapTurn_(true), ui.ButtonSet.OK);
}

// file-stamp (tools/stamp-files.js): lets Admin › System tell whether this file is the newest one
var FILE_STAMP_59_MapWeek = 'ffda1ad076';
