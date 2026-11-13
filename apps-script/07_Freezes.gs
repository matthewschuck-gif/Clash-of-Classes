/**
 * Clash of Classes -- Apps Script backend
 * File 7: Point Freezes. Points are never taken away; a squad can be frozen from earning more.
 *
 * Point Freezes tab (one row per freeze, filled in by staff):
 *   Squad | Start | End | Reason | Set By | Kind | Thaw Goals | Thaw By | Staff Sign-off | Thawed | Second Chance
 *   Squad: BLACK, GOLD, GREY or PURPLE. End blank = frozen until it thaws (or someone fills it in).
 *   Kind: Peaks (blank = Peaks) or Fun. How each thaws is in 14_Thaw.gs.
 *   Reason stays in the Sheet; students and Generals see the kind, the thaw goals and progress.
 *
 * While a squad is frozen, any event dated inside the freeze counts 0 for that squad, from
 * every project (Trait Dough, Mindset, Trail Journal, polls, admin...). The points it would
 * have earned are written into that event's note ("❄️ GOLD frozen: +5 not counted") so
 * nothing is lost from the record. Points the squad already had are never touched.
 * This runs every time the events tab is read (readEventsTab_), so it applies right away.
 */

const FREEZE_HEADERS_ = ['Squad', 'Start', 'End', 'Reason', 'Set By', 'Kind', 'Thaw Goals', 'Thaw By', 'Staff Sign-off', 'Thawed', 'Second Chance'];

function ensureFreezeTab_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('Point Freezes');
  if (!sh) {
    sh = ss.insertSheet('Point Freezes');
    sh.getRange(1, 1, 1, FREEZE_HEADERS_.length).setValues([FREEZE_HEADERS_]).setFontWeight('bold').setBackground('#490e6f').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    sh.getRange('A2:A200').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(SQUAD_KEYS_, true).setAllowInvalid(false).build());
  } else {
    // A tab made before thaw goals existed: add the missing columns at the end.
    const have = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0].map(function (h) { return String(h).trim(); });
    const missing = FREEZE_HEADERS_.filter(function (h) { return have.indexOf(h) === -1; });
    if (missing.length) sh.getRange(1, have.filter(String).length + 1, 1, missing.length).setValues([missing]).setFontWeight('bold').setBackground('#490e6f').setFontColor('#ffffff');
  }
  const col = function (name) { return sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(function (h) { return String(h).trim(); }).indexOf(name) + 1; };
  sh.getRange(2, col('Kind'), 199, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Peaks', 'Fun'], true).build());
  sh.getRange(2, col('Staff Sign-off'), 199, 1).insertCheckboxes();
}

// [{ squad, start, end, kind, goals, thawBy, signoff, thawed, secondChance, row }] with
// yyyy-MM-dd strings; end '' = no end yet; kind 'PEAKS' (blank counts as Peaks) or 'FUN'.
// Settings › FRESH_START (14_Thaw.gs): a freeze that began before it stops the day before, even
// before the Sheet row is updated, so the fresh-start day's points always count. raw = the rows as typed.
function freezeWindows_(raw) {
  const fresh = raw ? '' : String(getSetting_('FRESH_START') || '').trim().slice(0, 10);
  const cut = /^\d{4}-\d{2}-\d{2}$/.test(fresh) && todayYmd_() >= fresh ? fresh : '';
  return sheetRows_('Point Freezes').map(function (r) {
    return { squad: squadOf_(r['Squad']), start: String(r['Start'] || '').trim(), end: String(r['End'] || '').trim(),
      kind: /^fun/i.test(String(r['Kind'] || '').trim()) ? 'FUN' : 'PEAKS', goals: String(r['Thaw Goals'] || ''),
      thawBy: String(r['Thaw By'] || '').trim(), signoff: ticked_(r['Staff Sign-off']), thawed: String(r['Thawed'] || '').trim(),
      secondChance: String(r['Second Chance'] || '').trim(), row: r._row };
  }).filter(function (f) { return SQUAD_KEYS_.indexOf(f.squad) > -1 && f.start; }).map(function (f) {
    if (cut && f.start < cut && (!f.end || f.end >= cut)) f.end = ymdShift_(cut, -1);
    return f;
  });
}

function isFrozenOn_(windows, squad, ymd) {
  return windows.some(function (f) { return f.squad === squad && ymd >= f.start && (!f.end || ymd <= f.end); });
}

// The freeze in force today for a squad, or null: { until: 'yyyy-MM-dd' or '' }.
function freezeToday_(squad) {
  const today = todayYmd_();
  const f = freezeWindows_().filter(function (w) { return w.squad === squad && today >= w.start && (!w.end || today <= w.end); })[0];
  return f ? { until: f.end, kind: f.kind === 'FUN' ? 'Fun' : 'Peaks', thawBy: f.thawBy } : null;
}

// Every squad frozen today: [{ squad, until, kind }] (for the leaderboard and TV).
function activeFreezes_() {
  return SQUAD_KEYS_.map(function (s) { const f = freezeToday_(s); return f ? { squad: s, until: f.until, kind: f.kind } : null; })
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
