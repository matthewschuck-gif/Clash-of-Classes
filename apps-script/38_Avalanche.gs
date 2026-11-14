/**
 * Clash of Classes -- Apps Script backend
 * File 38: Avalanche, the comeback for the squad in last place (like star power: build it, then
 * set it off).
 *   - Only the squad in LAST place can charge it, and only once per marking period (MARKING_PERIODS).
 *   - It charges as that squad takes part this marking period: every Daily Poll / Pulse vote and every
 *     confirmed squad-event check-in by its students fills the meter (AVALANCHE_CHARGE, 100).
 *   - When the meter is full, a General sets it off in Generals HQ (or staff in Admin). For
 *     AVALANCHE_DAYS (7) the squad's points from AVALANCHE_TYPES (polls, check-ins, squad events,
 *     Pride Day) count double: a matching "Avalanche bonus" row is added on the events tab (checked
 *     about every hour), so the points tab always shows exactly what was earned.
 *   - Still last when it's full? Set it off any time before the marking period ends. Moved out of
 *     last place? The charge waits; it can only be set off from last place.
 * AVALANCHE tab: Squad | Marking Period | Started | Ends | Activated By | Bonus Paid.
 * AVALANCHE ON/OFF turns the whole thing on or off.
 */

const AVALANCHE_HEADERS_ = ['Squad', 'Marking Period', 'Started', 'Ends', 'Activated By', 'Bonus Paid'];

function avalancheOn_() { return String(getSetting_('AVALANCHE') || 'ON').toUpperCase() === 'ON'; }
function avalancheTypes_() { return listOf_(getSetting_('AVALANCHE_TYPES') || 'daily_poll, squad_move, squad_make, squad_talk, pride_day'); }

function ensureAvalancheTab_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getSheetByName('Avalanche')) return;
  const sh = ss.insertSheet('Avalanche');
  sh.getRange(1, 1, 1, AVALANCHE_HEADERS_.length).setValues([AVALANCHE_HEADERS_]).setFontWeight('bold').setBackground('#490e6f').setFontColor('#ffffff');
  sh.setFrozenRows(1);
}

// How much a squad has charged this marking period: its poll votes + confirmed squad-event check-ins.
function avalancheCharge_(squad, mp) {
  if (!mp) return 0;
  const inMp = function (v) { const d = String(v || '').slice(0, 10); return d >= mp.start && d <= mp.end; };
  let n = sheetRows_('Poll Votes').filter(function (v) { return squadOf_(v['Squad']) === squad && inMp(v['Date'] || v['At']); }).length;
  n += sheetRows_('Event Check-ins').filter(function (c) { return squadOf_(c['Squad']) === squad && ticked_(c['Confirmed']) && inMp(c['Date']); }).length;
  return n;
}

/** Everything the pages need about a squad's Avalanche. */
function avalancheStatus_(squad, totals) {
  if (!avalancheOn_()) return null;
  const today = todayYmd_(), mp = mpFor_(today).now;
  totals = totals || squadTotals_();
  const order = SQUAD_KEYS_.slice().sort(function (a, b) { return totals[b] - totals[a]; });
  const last = order[order.length - 1] === squad && totals[order[order.length - 1]] < totals[order[order.length - 2]];
  const rows = sheetRows_('Avalanche').filter(function (r) { return squadOf_(r['Squad']) === squad; });
  const active = rows.filter(function (r) { const s = String(r['Started']).slice(0, 10), e = String(r['Ends']).slice(0, 10); return s <= today && today <= e; })[0];
  const used = mp ? rows.some(function (r) { return String(r['Marking Period']) === mp.name; }) : false;
  const need = Math.max(1, settingNum_('AVALANCHE_CHARGE', 100));
  const charge = avalancheCharge_(squad, mp);
  return {
    squad: squad, last: last, mp: mp ? mp.name : '', need: need, charge: Math.min(charge, need), full: charge >= need, used: used,
    ready: !!mp && last && charge >= need && !used && !active,
    active: active ? { start: String(active['Started']).slice(0, 10), end: String(active['Ends']).slice(0, 10) } : null,
    days: Math.max(1, settingNum_('AVALANCHE_DAYS', 7)),
  };
}

/** Sets off the Avalanche for a squad (checks every rule). by = who did it. Returns the status. */
function startAvalanche_(squad, by) {
  const st = avalancheStatus_(squad);
  if (!st) throw new Error('Avalanche is turned off (Settings › AVALANCHE).');
  if (st.active) throw new Error('The Avalanche is already rolling until ' + st.active.end + '.');
  if (st.used) throw new Error('Your squad already used its Avalanche this marking period.');
  if (!st.last) throw new Error('Only the squad in last place can set off an Avalanche.');
  if (!st.full) throw new Error('The meter isn\'t full yet: ' + st.charge + ' of ' + st.need + '. Every vote and check-in charges it.');
  const today = todayYmd_(), end = ymdShift_(today, st.days - 1);
  ensureAvalancheTab_();
  getSheet_('Avalanche').appendRow([squad, st.mp, today, end, by || '', 0]);
  return avalancheStatus_(squad);
}

/** Generals HQ: the button. */
function clashGeneralsAvalanche(token) {
  const me = hqSession_(token);
  startAvalanche_(me.squad, 'General: ' + (me.name || me.id));
  return clashGeneralsHq(token);
}

/** Admin: set it off for a squad (same rules). */
function clashAvalancheStart(token, squad) {
  requireAdmin_(token);
  return JSON.stringify(startAvalanche_(squadOf_(squad), 'Admin'));
}

// The double points: for each Avalanche, the squad's matching points in its window, minus what was
// already paid, go in as one "Avalanche bonus" row. Runs from the background tick (about hourly).
function payAvalanches_() {
  const sh = getSheetSafe_('Avalanche');
  if (!sh) return [];
  const rows = sheetRows_('Avalanche');
  if (!rows.length) return [];
  const types = avalancheTypes_(), today = todayYmd_(), events = readEventsTab_(), paidCol = headerCol_(sh, 'Bonus Paid'), lines = [];
  rows.forEach(function (r) {
    const s = squadOf_(r['Squad']), start = String(r['Started']).slice(0, 10), end = String(r['Ends']).slice(0, 10);
    if (SQUAD_KEYS_.indexOf(s) === -1 || !start || start > today) return;
    const earned = events.filter(function (e) { const d = String(e.date || '').slice(0, 10); return d >= start && d <= end && types.indexOf(String(e.type || '').toLowerCase()) > -1; })
      .reduce(function (n, e) { return n + (Number(e.pts[s]) || 0); }, 0);
    const owed = earned - (Number(r['Bonus Paid']) || 0);
    if (owed <= 0) return;
    const pts = { BLACK: '', GOLD: '', GREY: '', PURPLE: '' }; pts[s] = owed;
    getSheet_('events').appendRow(['', 'Avalanche bonus: ' + titleSquad_(s), 'avalanche', today, 'x2 on polls, check-ins and squad events, ' + start + ' to ' + end, '', pts.BLACK, pts.GOLD, pts.GREY, pts.PURPLE]);
    if (paidCol) sh.getRange(r._row, paidCol).setValue(earned);
    lines.push(titleSquad_(s) + ' +' + owed + ' (Avalanche)');
  });
  return lines;
}

function maybePayAvalanches_() {
  const props = PropertiesService.getScriptProperties(), hour = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH');
  if (props.getProperty('avalanche-paid') === hour) return;
  props.setProperty('avalanche-paid', hour);
  payAvalanches_();
}

/** The squad that can charge right now (last place) and any rolling Avalanche, for the TV and Morning Clash. */
function avalancheBoard_(totals) {
  if (!avalancheOn_()) return null;
  totals = totals || squadTotals_();
  const list = SQUAD_KEYS_.map(function (s) { return avalancheStatus_(s, totals); });
  return { active: list.filter(function (x) { return x && x.active; }), last: list.filter(function (x) { return x && x.last; })[0] || null };
}

// file-stamp (tools/stamp-files.js): lets Admin › System tell whether this file is the newest one
var FILE_STAMP_38_Avalanche = '14c15fcae8';
