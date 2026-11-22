/**
 * Clash of Classes -- Apps Script backend
 * File 52: what lasts past one year, and the look of the seasons.
 *
 * HALL OF FAME tab: Show | Year | Award | Who | Squad | Detail. The new-year rollover adds the year's
 *   Clash Bash champion and every award before it clears anything (or Menu › Add this year to the
 *   Hall of Fame, any time after the reveal). Public page: clashofclasses.org/?page=hall. Students
 *   show as first name + last initial. Untick Show to hide a row; add old years by hand.
 *
 * 8TH GRADE SEND-OFF: SENDOFF_WRITE ON lets 8th graders write one piece of advice on My Climb
 *   (SEND-OFF ADVICE tab: When | Year | 800# | Squad | Advice | Show). Staff tick Show on the ones to
 *   share. SENDOFF_SHOW ON shows the ticked ones to 7th graders on My Climb ("from an 8th grader,
 *   Gold Squad"), never with a name. The tab is kept at rollover, so the fall's 7th graders see it.
 *
 * SEASON LOOK (SEASON_LOOK): AUTO = winter Dec to Feb, spring Mar to Jun, fall Sep to Nov; or WINTER,
 *   SPRING, FALL, OFF. The map and the Arcade games change their look. Nothing about the rules changes.
 */

const HALL_HEADERS_ = ['Show', 'Year', 'Award', 'Who', 'Squad', 'Detail'];
const SENDOFF_HEADERS_ = ['When', 'Year', '800#', 'Squad', 'Advice', 'Show'];

function ensureLegacyTabs_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  [['Hall of Fame', HALL_HEADERS_], ['Send-off Advice', SENDOFF_HEADERS_]].forEach(function (t) {
    if (ss.getSheetByName(t[0])) return;
    const sh = ss.insertSheet(t[0]);
    sh.getRange(1, 1, 1, t[1].length).setValues([t[1]]).setFontWeight('bold').setBackground('#490e6f').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  });
}

// ── Hall of Fame ──
/** Adds this season's champion and awards. Returns how many rows (0 if the year is already there). */
function hallCapture_() {
  ensureLegacyTabs_();
  const win = seasonWindow_(), year = schoolYearLabel_(win.to);
  if (sheetRows_('Hall of Fame').some(function (r) { return String(r['Year']) === year && !/^Summit Cup/.test(String(r['Award'] || '')); })) return 0; // monthly Summit Cup rows (63) don't count
  const totals = squadTotals_(), order = SQUAD_KEYS_.slice().sort(function (a, b) { return totals[b] - totals[a]; });
  const rows = [[true, year, 'Clash Bash Champion', titleSquad_(order[0]) + ' Squad', order[0], (totals[order[0]] || 0).toLocaleString() + ' points']];
  (yearAwards_(win) || []).forEach(function (a) { rows.push([true, year, a.title, shortName_(a.who), a.squad || '', a.detail || '']); });
  // Each squad's motto and the Rivalry Week trophies (54_Rally.gs, 57_Leagues.gs).
  try { const b = squadBanners_(); SQUAD_KEYS_.forEach(function (s) { if (b[s]) rows.push([true, year, titleSquad_(s) + ' Squad motto', b[s].icon + ' \u201c' + b[s].motto + '\u201d', s, '']); }); } catch (e) { /* none */ }
  try { const t = rivalryTrophies_(); SQUAD_KEYS_.forEach(function (s) { if (t[s]) rows.push([true, year, 'Rivalry Week trophies', titleSquad_(s) + ' Squad', s, t[s] + ' win' + (t[s] === 1 ? '' : 's')]); }); } catch (e) { /* none */ }
  const sh = getSheet_('Hall of Fame');
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, HALL_HEADERS_.length).setValues(rows);
  return rows.length;
}

function hallYears_() {
  const by = {};
  sheetRows_('Hall of Fame').filter(function (r) { return ticked_(r['Show']) && String(r['Year'] || '').trim() && String(r['Award'] || '').trim(); }).forEach(function (r) {
    const y = String(r['Year']).trim();
    const keep = /motto|trophies|team league/i.test(String(r['Award'])) || /squad$/i.test(String(r['Who']).trim());
    (by[y] = by[y] || []).push({ award: String(r['Award']).trim(), who: keep ? String(r['Who']).trim() : shortName_(r['Who']), squad: squadOf_(r['Squad']), detail: String(r['Detail'] || '').trim() });
  });
  return Object.keys(by).sort().reverse().map(function (y) {
    const champ = by[y].filter(function (a) { return /champion/i.test(a.award) && SQUAD_KEYS_.indexOf(a.squad) > -1; })[0];
    return { year: y, champion: champ ? champ.squad : '', awards: by[y] };
  });
}

function clashHall() { return JSON.stringify(hallYears_()); }

function serveHall_() {
  return HtmlService.createHtmlOutputFromFile('Hall').setTitle('Hall of Fame · Clash of Classes')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function menuHallCapture_() {
  const ui = SpreadsheetApp.getUi();
  if (ui.alert('Add this year to the Hall of Fame?', 'Adds the Clash Bash champion and the awards from the reveal to the Hall of Fame tab (once per year).', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  const n = hallCapture_();
  ui.alert(n ? 'Added ' + n + ' rows. Untick Show on any you want hidden.' : 'This year is already in the Hall of Fame.');
}

// ── 8th grade send-off ──
function sendoffOn_(key) { return String(getSetting_(key) || 'OFF').toUpperCase() === 'ON'; }

function studentGrade_(id) {
  const b = boardRows_().filter(function (r) { return id800_(r['800#']) === id; })[0];
  if (b && String(b['Grade'] || '').trim()) return String(b['Grade']).replace(/\D/g, '');
  const s = sheetRows_('Students').filter(function (r) { return id800_(r['800#']) === id; })[0];
  return s ? String(s['Grade'] || '').replace(/\D/g, '') : '';
}

/** My Climb: { write, mine, show: [{ advice, squad }] }. */
function clashSendoff(asId, adminToken) {
  return withSample_(asId, function () {
    const id = meId_(asId, adminToken), grade = id === SAMPLE_ID_ ? (adminToken ? '8' : '') : studentGrade_(id), preview = !!adminToken;
    const mine = sheetRows_('Send-off Advice').filter(function (r) { return id800_(r['800#']) === id; }).pop();
    const show = (sendoffOn_('SENDOFF_SHOW') && (grade === '7' || preview))
      ? sheetRows_('Send-off Advice').filter(function (r) { return ticked_(r['Show']) && String(r['Advice'] || '').trim(); })
        .map(function (r) { return { advice: String(r['Advice']).trim(), squad: squadOf_(r['Squad']) }; })
      : [];
    // A different handful each day.
    const start = show.length ? (Number(todayYmd_().replace(/\D/g, '')) % show.length) : 0;
    let weather = null; try { weather = weatherNow_().now; } catch (e) { weather = null; }
    return JSON.stringify({ weather: weather, write: sendoffOn_('SENDOFF_WRITE') && grade === '8', mine: mine ? String(mine['Advice'] || '') : '',
      show: show.slice(start).concat(show.slice(0, start)).slice(0, 4) });
  });
}

function clashSendoffSave(text, asId, adminToken) {
  return withSample_(asId, function () {
    const id = meId_(asId, adminToken);
    if (!sendoffOn_('SENDOFF_WRITE')) throw new Error('Send-off advice is closed right now.');
    if (id === SAMPLE_ID_) return clashSendoff(asId, adminToken);
    if (studentGrade_(id) !== '8') throw new Error('Send-off advice is for 8th graders.');
    text = String(text || '').replace(/\s+/g, ' ').trim().slice(0, 280);
    if (text.length < 15) throw new Error('Write a sentence or two of advice.');
    ensureLegacyTabs_();
    const sh = getSheet_('Send-off Advice'), b = boardRows_().filter(function (r) { return id800_(r['800#']) === id; })[0] || {};
    const old = sheetRows_('Send-off Advice').filter(function (r) { return id800_(r['800#']) === id; }).pop();
    const row = [Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'), schoolYearLabel_(todayYmd_()), id, squadOf_(b['Squad']), text, false];
    if (old) sh.getRange(old._row, 1, 1, row.length).setValues([row]); // edited advice needs a fresh look from staff
    else sh.appendRow(row);
    return clashSendoff(asId, adminToken);
  });
}

// ── Season look ──
function seasonLook_() {
  const v = String(getSetting_('SEASON_LOOK') || 'AUTO').toUpperCase();
  if (v === 'OFF') return '';
  if (['WINTER', 'SPRING', 'FALL'].indexOf(v) > -1) return v.toLowerCase();
  const m = Number(todayYmd_().slice(5, 7));
  return m === 12 || m <= 2 ? 'winter' : m <= 6 ? 'spring' : m >= 9 ? 'fall' : '';
}

// file-stamp (tools/stamp-files.js): lets Admin › System tell whether this file is the newest one
var FILE_STAMP_52_Legacy = '41642e292f';
