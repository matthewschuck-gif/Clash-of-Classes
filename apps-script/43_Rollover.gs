/**
 * Clash of Classes -- Apps Script backend
 * File 43: the new school year. Menu › New school year: preview, then roll over.
 *
 * 1. A full copy of this Sheet is made first ("Clash of Classes Points — 2026-27 archive", in the same
 *    Drive folder) and its link is saved on Settings › LAST_ARCHIVE. Nothing is cleared unless the
 *    copy exists.
 * 2. The year's records are cleared (header rows kept): ROLLOVER_CLEAR_ below (points, votes, polls,
 *    Squad Monthly, freezes, Avalanche, Virtual Dough, service, sign-ups, the logs...). The saved site
 *    data drops finalized metric months, the trait check-offs and the activity log.
 * 3. Holders on the map are cleared, Mayhem checklists unticked, Poll Ideas made ready again.
 * 4. Settings: SEASON_END moves a year, BASH_REVEAL and FRESH_START are cleared, TOUR_VERSION goes up.
 *    MARKING_PERIODS is left for you to update with the new district dates (the self-check flags it).
 * Kept: every setup tab (rules, zones, Poll Ideas, traits, wheel, tour, Dough Teachers, Squad Themes,
 * Year Plan...), Settings, and the Students tab (the roster sync refreshes it).
 * The real run asks you to type NEW YEAR.
 */

const ROLLOVER_CLEAR_ = ['events', 'Daily Polls', 'Poll Votes', 'Poll Insights', 'You Said We Did', 'Squad Monthly', 'Zone Overrides', 'Map Battles',
  'Point Freezes', 'Frost Log', 'Avalanche', 'Virtual Dough', 'Service Submissions', 'Squad Spotlight', 'Event Check-ins', 'Sign-ups', 'Homeroom Points',
  'Nominations', 'General Letters', 'Shirt Designs', 'Shirt Votes', 'Shirt Orders', 'Squad Night', 'Boost Ledger', 'Student Tickets', 'Squad Hand',
  'Streak Banks', 'Redemptions', 'Usage', 'Admin Edits', 'Weekly Digest Log', 'Points Balance', 'Student Board', 'Squad Board', 'Arcade Plays', 'Family Pick Votes', 'Staff Shout-outs', 'Squad Goals', 'Helping Hands', 'Team League'];

function schoolYearLabel_(ymd) {
  const y = Number(ymd.slice(0, 4)), m = Number(ymd.slice(5, 7));
  const start = m >= 7 ? y : y - 1;
  return start + '-' + String(start + 1).slice(2);
}

/** What a rollover would clear: [{ tab, rows }] plus the archive name. Nothing changes. */
function rolloverPreview_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tabs = ROLLOVER_CLEAR_.map(function (t) { const sh = ss.getSheetByName(t); return sh ? { tab: t, rows: Math.max(0, sh.getLastRow() - 1) } : null; })
    .filter(function (x) { return x && x.rows; });
  return { archive: 'Clash of Classes Points — ' + schoolYearLabel_(todayYmd_()) + ' archive', tabs: tabs };
}

/** The rollover. confirm must be "NEW YEAR". Returns lines saying what happened. */
function rolloverYear_(confirm) {
  if (String(confirm || '').trim().toUpperCase() !== 'NEW YEAR') throw new Error('Type NEW YEAR to roll over.');
  const ss = SpreadsheetApp.getActiveSpreadsheet(), lines = [];
  const plan = rolloverPreview_();
  // 1. The archive copy, first. Nothing is cleared without it.
  const me = DriveApp.getFileById(ss.getId());
  const copy = me.makeCopy(plan.archive);
  if (!copy || !copy.getId()) throw new Error('The archive copy could not be made, so nothing was cleared.');
  setSettingValue_('LAST_ARCHIVE', copy.getUrl());
  lines.push('Archive: ' + plan.archive + ' (link on Settings › LAST_ARCHIVE)');
  // 1b. The Hall of Fame keeps this year's champion and awards (52_Legacy.gs).
  try { const n = hallCapture_(); lines.push(n ? 'Hall of Fame: added ' + n + ' rows for this year' : 'Hall of Fame: this year was already there'); } catch (e) { lines.push('Hall of Fame not updated: ' + e.message); }
  // 2. The year's records.
  plan.tabs.forEach(function (t) {
    const sh = ss.getSheetByName(t.tab);
    if (sh && sh.getLastRow() > 1) sh.deleteRows(2, sh.getLastRow() - 1);
    lines.push('Cleared ' + t.tab + ' (' + t.rows + ' rows)');
  });
  try {
    const res = loadAppData_('coc'), v = res.value || {};
    v.monthlyMetrics = []; v.traitsDone = []; v.activityLog = [];
    if (Array.isArray(v.events)) v.events = [];
    if (v.metrics) ['attendance', 'referral', 'grades'].forEach(function (k) { if (v.metrics[k]) SQUAD_KEYS_.forEach(function (s) { v.metrics[k][s] = 0; }); });
    if (v.metrics) v.metrics.month = '';
    saveAppData_('coc', v);
    lines.push('Site data: finalized months, trait check-offs and the activity log cleared');
  } catch (e) { lines.push('Site data not reset: ' + e.message); }
  // 3. The map, the checklists, the poll bank.
  const blank = function (tab, col, value) {
    const sh = ss.getSheetByName(tab), c = sh ? headerCol_(sh, col) : 0;
    if (sh && c && sh.getLastRow() > 1) sh.getRange(2, c, sh.getLastRow() - 1, 1).setValue(value);
  };
  blank('Map Zones', 'Holder', ''); blank('Map Zones', 'Since', ''); blank('Mayhem Checklist', 'Done', false); blank('Poll Ideas', 'Used On', ''); blank('Expeditions', 'Winner', ''); blank('Expeditions', 'Won On', '');
  blank('Lore Chapters', 'Unlocked', ''); blank('Engine Powers', 'Unlocked', ''); blank('Opponents', 'Status', ''); blank('Round Goals', 'Status', '');
  blank('Challenges', 'Awarded', ''); blank('Squad Duels', 'Status', '');
  lines.push('Map holders cleared, checklists unticked, Poll Ideas ready again');
  // 4. Settings.
  const end = seasonEnd_();
  if (end) setSettingValue_('SEASON_END', ymdShift_(end, 364));
  setSettingValue_('BASH_REVEAL', ''); setSettingValue_('FRESH_START', '');
  setSettingValue_('TOUR_VERSION', String((Number(getSetting_('TOUR_VERSION')) || 1) + 1));
  PropertiesService.getScriptProperties().deleteProperty('state-month');
  lines.push('Settings: SEASON_END moved a year (check the day), BASH_REVEAL and FRESH_START cleared. Update MARKING_PERIODS with the new district dates.');
  return lines;
}

function menuRolloverPreview_() {
  const p = rolloverPreview_();
  SpreadsheetApp.getUi().alert('New school year: what would happen',
    'First a full copy is saved as "' + p.archive + '".\n\nThen these are cleared (headers kept):\n' +
    (p.tabs.length ? p.tabs.map(function (t) { return '• ' + t.tab + ': ' + t.rows + ' rows'; }).join('\n') : '• nothing (they are empty)') +
    '\n\nKept: rules, zones, Poll Ideas, traits, wheel, tour, Dough Teachers, Squad Themes, Year Plan, Settings, Students.\nNothing has changed yet.', SpreadsheetApp.getUi().ButtonSet.OK);
}

function menuRollover_() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('Roll over to a new school year', 'This archives the Sheet and clears this year\'s points, polls and records. Type NEW YEAR to go ahead.', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  try { ui.alert('New school year', rolloverYear_(r.getResponseText()).join('\n'), ui.ButtonSet.OK); } catch (e) { ui.alert(e.message); }
}

// file-stamp (tools/stamp-files.js): lets Admin › System tell whether this file is the newest one
var FILE_STAMP_43_Rollover = 'f2f69f3f5e';
