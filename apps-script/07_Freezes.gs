/**
 * Clash of Classes -- Apps Script backend
 * File 7: Point Freezes. Points are never taken away; a squad can be frozen from earning more.
 *
 * Point Freezes tab (one row per freeze, filled in by staff):
 *   Squad | Start | End | Reason | Set By
 *   Squad: BLACK, GOLD, GREY or PURPLE. End blank = frozen until someone fills it in.
 *   Reason stays in the Sheet; students and Generals only see "frozen until <date>".
 *
 * While a squad is frozen, any event dated inside the freeze counts 0 for that squad, from
 * every project (Trait Dough, Mindset, Trail Journal, polls, admin...). The points it would
 * have earned are written into that event's note ("❄️ GOLD frozen: +5 not counted") so
 * nothing is lost from the record. Points the squad already had are never touched.
 * This runs every time the events tab is read (readEventsTab_), so it applies right away.
 */

const FREEZE_HEADERS_ = ['Squad', 'Start', 'End', 'Reason', 'Set By'];

function ensureFreezeTab_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getSheetByName('Point Freezes')) return;
  const sh = ss.insertSheet('Point Freezes');
  sh.getRange(1, 1, 1, FREEZE_HEADERS_.length).setValues([FREEZE_HEADERS_]).setFontWeight('bold').setBackground('#490e6f').setFontColor('#ffffff');
  sh.setFrozenRows(1);
  sh.getRange('A2:A200').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(SQUAD_KEYS_, true).setAllowInvalid(false).build());
}

// [{ squad, start, end }] with yyyy-MM-dd strings; end '' = no end yet.
function freezeWindows_() {
  return sheetRows_('Point Freezes').map(function (r) {
    return { squad: squadOf_(r['Squad']), start: String(r['Start'] || '').trim(), end: String(r['End'] || '').trim() };
  }).filter(function (f) { return SQUAD_KEYS_.indexOf(f.squad) > -1 && f.start; });
}

function isFrozenOn_(windows, squad, ymd) {
  return windows.some(function (f) { return f.squad === squad && ymd >= f.start && (!f.end || ymd <= f.end); });
}

// The freeze in force today for a squad, or null: { until: 'yyyy-MM-dd' or '' }.
function freezeToday_(squad) {
  const today = todayYmd_();
  const f = freezeWindows_().filter(function (w) { return w.squad === squad && today >= w.start && (!w.end || today <= w.end); })[0];
  return f ? { until: f.end } : null;
}

// Every squad frozen today: [{ squad, until }] (for the leaderboard and TV).
function activeFreezes_() {
  return SQUAD_KEYS_.map(function (s) { const f = freezeToday_(s); return f ? { squad: s, until: f.until } : null; })
    .filter(function (f) { return f; });
}

/**
 * Called by readEventsTab_ with the events it built and the tab's raw rows (and by clashSave
 * with sheet = null, to fix the events before they are written). Zeroes points a
 * frozen squad earned inside a freeze and notes them on that row. Changes `events` in place.
 */
function applyFreezes_(sheet, events, rowNumbers) {
  const windows = freezeWindows_();
  if (!windows.length) return;
  const headers = SCHEMA[TABS.EVENTS];
  const noteCol = headers.indexOf('note');
  events.forEach(function (e, i) {
    const ymd = String(e.date || '').slice(0, 10);
    if (!ymd) return;
    const frozen = SQUAD_KEYS_.filter(function (s) { return (Number(e.pts[s]) || 0) > 0 && isFrozenOn_(windows, s, ymd); });
    if (!frozen.length) return;
    const notes = frozen.map(function (s) { return '❄️ ' + s + ' frozen: +' + e.pts[s] + ' not counted'; });
    e.note = (e.note ? String(e.note) + ' · ' : '') + notes.join(' · ');
    frozen.forEach(function (s) {
      e.pts[s] = 0;
      if (sheet) sheet.getRange(rowNumbers[i], headers.indexOf(s) + 1).setValue(0);
    });
    if (sheet) sheet.getRange(rowNumbers[i], noteCol + 1).setValue(e.note);
  });
}
