/**
 * Clash of Classes -- Apps Script backend
 * File 8: Squad Duels -- a weekly head-to-head the Generals pick. The winner gets bonus points.
 *
 * Squad Duels tab (one row per duel, filled in by staff or from the Generals' pick):
 *   Show | Title | Start | End | Measure | Event Types | Squads | Bonus Points | Status | Results
 *   Measure: Points (sum of the squad's points), Entries (events where the squad scored) or
 *            Poll Votes (Daily Poll votes from the squad's students).
 *   Event Types: which Clash event types count (e.g. trait_dough); blank = every event.
 *   Squads: blank = all four; or two squads, e.g. "GOLD vs GREY".
 *   The bonuses that other features give (challenge, duel) never count toward a duel.
 *
 * Points (when DUEL_BONUSES is ON, checked hourly once the menu turns it on): the day after a
 * duel's End, the squad(s) with the highest score get Bonus Points as one "Squad Duel" event
 * (type duel_bonus). A tie means every tied squad wins. Never twice: Status becomes "Awarded".
 */

const DUEL_HEADERS_ = ['Show', 'Title', 'Start', 'End', 'Measure', 'Event Types', 'Squads', 'Bonus Points', 'Status', 'Results'];
const BONUS_TYPES_ = ['challenge_bonus', 'duel_bonus'];

function ensureDuelTab_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getSheetByName('Squad Duels')) return;
  const sh = ss.insertSheet('Squad Duels');
  const monday = new Date(); monday.setDate(monday.getDate() - ((monday.getDay() + 6) % 7));
  const friday = new Date(monday); friday.setDate(monday.getDate() + 4);
  const ymd = function (d) { return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd'); };
  sh.getRange(1, 1, 2, DUEL_HEADERS_.length).setValues([DUEL_HEADERS_,
    [false, 'Trait Dough Showdown', ymd(monday), ymd(friday), 'Entries', 'trait_dough', '', 25, '', '']]);
  sh.getRange(1, 1, 1, DUEL_HEADERS_.length).setFontWeight('bold').setBackground('#490e6f').setFontColor('#ffffff');
  sh.setFrozenRows(1);
}

function duelSquads_(v) {
  const list = String(v || '').toUpperCase().replace(/GRAY/g, 'GREY').split(/[^A-Z]+/).filter(function (s) { return SQUAD_KEYS_.indexOf(s) > -1; });
  return list.length >= 2 ? list : SQUAD_KEYS_.slice();
}

// [{ row, title, start, end, measure, squads, bonus, status, scores: {SQUAD: n}, leaders: [..], active, done, awarded }]
function duelProgress_() {
  const rows = sheetRows_('Squad Duels');
  if (!rows.length) return [];
  const today = todayYmd_();
  const events = readEventsTab_();
  let votes = null;
  return rows.filter(function (r) { return ticked_(r['Show']) && String(r['Title'] || '').trim(); }).map(function (r) {
    const start = asYmd_(r['Start']), end = asYmd_(r['End']);
    const measure = /poll/i.test(String(r['Measure'])) ? 'Poll Votes' : /^entr/i.test(String(r['Measure'])) ? 'Entries' : 'Points';
    const types = String(r['Event Types'] || '').split(',').map(function (s) { return s.trim(); }).filter(String);
    const squads = duelSquads_(r['Squads']);
    const scores = {};
    squads.forEach(function (s) { scores[s] = 0; });
    const inside = function (d) { d = String(d || '').slice(0, 10); return (!start || d >= start) && (!end || d <= end); };
    if (measure === 'Poll Votes') {
      votes = votes || sheetRows_('Poll Votes');
      votes.forEach(function (v) { const s = squadOf_(v['Squad']); if (scores[s] !== undefined && inside(asYmd_(v['Date']))) scores[s]++; });
    } else {
      events.forEach(function (e) {
        if (BONUS_TYPES_.indexOf(String(e.type)) > -1) return;
        if (types.length && types.indexOf(String(e.type)) === -1) return;
        if (!inside(e.date)) return;
        squads.forEach(function (s) { const p = Number(e.pts[s]) || 0; if (p) scores[s] += measure === 'Entries' ? 1 : p; });
      });
    }
    const top = Math.max.apply(null, squads.map(function (s) { return scores[s]; }));
    return {
      row: r._row, title: String(r['Title']).trim(), start: start, end: end, measure: measure, squads: squads,
      bonus: Number(r['Bonus Points']) || 0, status: String(r['Status'] || ''), scores: scores,
      leaders: top > 0 ? squads.filter(function (s) { return scores[s] === top; }) : [],
      active: (!start || today >= start) && (!end || today <= end), done: !!end && today > end,
      awarded: /^awarded/i.test(String(r['Status'] || '')),
    };
  });
}

// For the site: running duels and ones that ended in the last 7 days.
function clashDuels() {
  const weekAgo = Utilities.formatDate(new Date(Date.now() - 7 * 86400000), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return JSON.stringify(duelProgress_().filter(function (d) { return d.active || (d.done && d.end >= weekAgo); })
    .map(function (d) { delete d.row; return d; }));
}

/** After a duel's End: the winner(s) get the bonus once. Returns what was sent. */
function awardDuels() {
  if (String(getSetting_('DUEL_BONUSES')).toUpperCase() !== 'ON') return ['DUEL_BONUSES is OFF.'];
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const lines = [];
    const sheet = getSheetSafe_('Squad Duels');
    const statusCol = DUEL_HEADERS_.indexOf('Status') + 1;
    duelProgress_().forEach(function (d) {
      if (!d.done || d.awarded || !d.bonus) return;
      const results = d.squads.map(function (s) { return s + ' ' + d.scores[s]; }).join(', ');
      if (!d.leaders.length) {
        sheet.getRange(d.row, statusCol, 1, 2).setValues([['Awarded ' + todayYmd_() + ' (no one scored)', results]]);
        return;
      }
      const pts = { BLACK: '', GOLD: '', GREY: '', PURPLE: '' };
      d.leaders.forEach(function (s) { pts[s] = d.bonus; });
      getSheet_(TABS.EVENTS).appendRow(['', 'Squad Duel: ' + d.title, 'duel_bonus', d.end,
        'Winner: ' + d.leaders.join(' & ') + ' (' + results + ')', '', pts.BLACK, pts.GOLD, pts.GREY, pts.PURPLE]);
      sheet.getRange(d.row, statusCol, 1, 2).setValues([['Awarded ' + todayYmd_() + ': ' + d.leaders.join(' & ') + ' +' + d.bonus, results]]);
      lines.push(d.title + ': ' + d.leaders.join(' & ') + ' +' + d.bonus);
    });
    return lines;
  } finally {
    lock.releaseLock();
  }
}

function menuTurnOnDuels_() {
  setSettingValue_('DUEL_BONUSES', 'ON');
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'awardDuels') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('awardDuels').timeBased().everyHours(1).create();
  const lines = awardDuels();
  SpreadsheetApp.getUi().alert('Squad Duel bonuses are ON (checked every hour; each duel is scored the day after it ends).\n\n' + (lines.length ? lines.join('\n') : 'Nothing to award yet.'));
}

function menuTurnOffDuels_() {
  setSettingValue_('DUEL_BONUSES', 'OFF');
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'awardDuels') ScriptApp.deleteTrigger(t); });
  SpreadsheetApp.getUi().alert('Squad Duel bonuses are OFF. Duels still show on the site; no bonus is added.');
}
