/**
 * Clash of Classes -- Apps Script backend
 * File 28: Frost checks. Squads freeze on set lines, warn first, staff confirm, thaw together.
 *
 * FREEZE LINES tab (staff edit any time): one row per marker, the same line for every squad.
 *   Show | Marker | Measures | Freezes If | Line | Per 100 Students | Kind | What Students See
 *   Measures: stat: <Squad Monthly columns, comma-separated, added> (lrr_ld | failing |
 *             interventions_teacher, interventions_admin | shout_outs ...) | attendance (Squad Monthly,
 *             else the metrics tab) | core f (% of students with a core F = 100 - the metrics tab's
 *             grades) | points (every point that month)
 *   Freezes If: Above or Below the Line. Kind: Peaks (serious) or Fun (push your luck).
 * THE CHECK (on the 1st, for the month before, from FROST_START on; Settings FROST_CHECK):
 *   no marker crossed = Safe; one = Frost warning; two or more, or the same marker as the month
 *   before = Freeze proposed. A squad already frozen is logged Still frozen. Every result is a
 *   row on the FROST LOG tab. Nothing freezes until staff tick Approve on a Freeze proposed row;
 *   then a Point Freezes row is written (07_Freezes.gs) with its thaw goals and the squad's new
 *   points stop counting.
 * THAW GOALS (all together, never one student):
 *   Peaks: marker: <each marker crossed> (back on the safe side for a whole month, checked on the
 *          1st) ; everyone: FROST_THAW_EVERYONE % ; points: FROST_THAW_POINTS ; a squad reflection
 *          or service act led by the Generals (staff sign-off).
 *   Fun:   points: FROST_FUN_POINTS, Thaw By the end of next month (14_Thaw.gs tiers).
 *   Not thawed by the next 1st = still frozen with the same goals.
 * WHAT STUDENTS SEE: squad totals against each line (no names): the frost meter on My Squad, My
 *   Climb and the TV, and the explainer (?page=frost: squad meeting slides + read-aloud).
 */

const FREEZE_LINE_HEADERS_ = ['Show', 'Marker', 'Measures', 'Freezes If', 'Line', 'Per 100 Students', 'Kind', 'What Students See'];
const FREEZE_LINES_SEED_ = [
  [true, 'LRR + LD students', 'stat: lrr_ld', 'Above', 8, true, 'Peaks', 'Students with an LRR or a Lunch Detention'],
  [true, 'Interventions', 'stat: interventions_teacher, interventions_admin', 'Above', 15, true, 'Peaks', 'Teacher and admin interventions'],
  [true, 'Attendance', 'attendance', 'Below', 92, false, 'Peaks', 'Attendance rate'],
  [true, 'Failing grades', 'stat: failing', 'Above', 15, true, 'Peaks', 'Failing grades'],
  [true, 'Points floor', 'points', 'Below', 50, true, 'Fun', 'Points earned'],
];
const FROST_LOG_HEADERS_ = ['Month', 'Squad', 'Status', 'Markers Crossed', 'Readings', 'Kind', 'Approve', 'Applied'];

function ensureFrostTabs_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  [['Freeze Lines', FREEZE_LINE_HEADERS_, FREEZE_LINES_SEED_], ['Frost Log', FROST_LOG_HEADERS_, null]].forEach(function (t) {
    if (ss.getSheetByName(t[0])) return;
    const sh = ss.insertSheet(t[0]);
    sh.getRange(1, 1, 1, t[1].length).setValues([t[1]]).setFontWeight('bold').setBackground('#490e6f').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    if (t[2]) sh.getRange(2, 1, t[2].length, t[1].length).setValues(t[2]);
    try {
      if (t[0] === 'Frost Log') sh.getRange(2, FROST_LOG_HEADERS_.indexOf('Approve') + 1, 400, 1).insertCheckboxes();
      if (t[0] === 'Freeze Lines') {
        sh.getRange(2, 4, 50, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Above', 'Below'], true).build());
        sh.getRange(2, 7, 50, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Peaks', 'Fun'], true).build());
      }
    } catch (e) { console.warn('Frost tab checkboxes: ' + e.message); }
  });
}

function freezeLines_() {
  return sheetRows_('Freeze Lines').filter(function (r) { return ticked_(r['Show']) && String(r['Marker'] || '').trim() && String(r['Line']).trim() !== ''; }).map(function (r) {
    return { marker: String(r['Marker']).trim(), measures: String(r['Measures'] || '').trim().toLowerCase(), below: /^below/i.test(String(r['Freezes If'] || '')),
      line: Number(r['Line']) || 0, per100: ticked_(r['Per 100 Students']), kind: /^fun/i.test(String(r['Kind'] || '')) ? 'Fun' : 'Peaks',
      label: String(r['What Students See'] || r['Marker']).trim() };
  });
}

function monthWindow_(month, today) {
  const y = Number(month.slice(0, 4)), m = Number(month.slice(5, 7));
  const last = month + '-' + ('0' + new Date(Date.UTC(y, m, 0)).getUTCDate()).slice(-2);
  return { from: month + '-01', to: today && today < last ? today : last, days: Number(last.slice(8)) };
}
function prevMonth_(month) { return lastMonthWindow_(month + '-15').month; }

// The calendar month before today's: { from, to, month: 'yyyy-MM' }.
function lastMonthWindow_(today) {
  const y = Number(today.slice(0, 4)), m = Number(today.slice(5, 7));
  const py = m === 1 ? y - 1 : y, pm = m === 1 ? 12 : m - 1;
  const month = py + '-' + (pm < 10 ? '0' : '') + pm;
  const days = new Date(Date.UTC(py, pm, 0)).getUTCDate();
  return { from: month + '-01', to: month + '-' + days, month: month };
}

// Older names for the Squad Monthly columns.
const STAT_ALIASES_ = { 'lrr_+_ld_students': 'lrr_ld', 'lrr_+_ld': 'lrr_ld', 'interventions': 'interventions_teacher, interventions_admin', 'failing_grades': 'failing' };

// One number per squad for a month: Squad Monthly columns (added together), or attendance / grades /
// referral from the metrics tab when Squad Monthly has no such column that month. null = no number.
function schoolStat_(name, month) {
  month = month || lastMonthWindow_(todayYmd_()).month;
  const out = { BLACK: null, GOLD: null, GREY: null, PURPLE: null };
  let keys = [];
  listOf_(name).forEach(function (k) { k = measureKey_(k); keys = keys.concat(STAT_ALIASES_[k] ? listOf_(STAT_ALIASES_[k]) : [k]); });
  if (!keys.length) return out;
  const monthly = squadMonthly_(month);
  const has = function (k) { return SQUAD_KEYS_.some(function (s) { return monthly[s] && monthly[s][k] !== undefined; }); };
  if (keys.every(has)) {
    SQUAD_KEYS_.forEach(function (s) {
      const vals = keys.map(function (k) { return monthly[s] ? monthly[s][k] : undefined; });
      out[s] = vals.some(function (v) { return v === undefined; }) ? null : vals.reduce(function (a, b) { return a + b; }, 0);
    });
    return out;
  }
  if (keys.length === 1 && ['attendance', 'grades', 'referral'].indexOf(keys[0]) > -1) {
    const m = readMetricsTab_();
    if (m && m[keys[0]]) SQUAD_KEYS_.forEach(function (s) { const v = Number(m[keys[0]][s]); if (v > 0) out[s] = v; });
  }
  return out;
}

// { SQUAD: number|null } for one line and month. pace = project a month still running to its end.
function lineReading_(L, month, events, sizes, today, pace) {
  const out = { BLACK: null, GOLD: null, GREY: null, PURPLE: null };
  const win = monthWindow_(month, today);
  let raw = null, count = false;
  if (L.measures.indexOf('stat:') === 0) { raw = schoolStat_(L.measures.slice(5).trim(), month); count = true; }
  else if (L.measures === 'attendance') raw = schoolStat_('attendance', month);
  else if (L.measures === 'core f') {
    const g = schoolStat_('grades'); raw = {};
    SQUAD_KEYS_.forEach(function (s) { raw[s] = g[s] == null ? null : Math.round(10 * (100 - g[s])) / 10; });
  } else if (L.measures.indexOf('points') === 0) {
    const spec = L.measures.replace(/^points\s*:?\s*/, '');
    raw = pointsBy_(events, win.from, win.to, spec ? eventMatcher_(spec) : null).points; count = true;
  }
  if (!raw) return out;
  const elapsed = Number(win.to.slice(8)) || win.days;
  SQUAD_KEYS_.forEach(function (s) {
    let v = raw[s];
    if (v == null) return;
    v = Number(v) || 0;
    if (count && pace && elapsed < win.days) v = v * win.days / Math.max(1, elapsed);
    if (L.per100 && !sizes[s]) return; // no students on the roster yet: nothing to measure
    if (L.per100) v = 100 * v / sizes[s];
    out[s] = Math.round(10 * v) / 10;
  });
  return out;
}

function crossed_(L, v) { return v != null && (L.below ? v < L.line : v > L.line); }

// Every squad's readings for a month: { SQUAD: [{ marker, label, value, line, below, crossed, kind }] }.
function frostReadings_(month, events, opts) {
  opts = opts || {};
  events = events || readEventsTab_();
  const sizes = squadSizes_(month), lines = freezeLines_(), out = {};
  SQUAD_KEYS_.forEach(function (s) { out[s] = []; });
  lines.forEach(function (L) {
    const r = lineReading_(L, month, events, sizes, opts.today, opts.pace);
    SQUAD_KEYS_.forEach(function (s) {
      out[s].push({ marker: L.marker, label: L.label, value: r[s], line: L.line, below: L.below, per100: L.per100, kind: L.kind, crossed: crossed_(L, r[s]) });
    });
  });
  return out;
}

function frostStart_() { const v = String(getSetting_('FROST_START') || '2027-01').trim().slice(0, 7); return /^\d{4}-\d{2}$/.test(v) ? v : '2027-01'; }

/** The check for one month (default last month). Writes Frost Log rows once. Returns lines. */
function runFrostCheck_(month, force) {
  ensureFrostTabs_();
  const today = todayYmd_();
  month = month || lastMonthWindow_(today).month;
  if (month < frostStart_()) return ['Frost check: ' + month + ' is before FROST_START (' + frostStart_() + '), nothing checked.'];
  const log = sheetRows_('Frost Log');
  const has = function (m, s) { return log.filter(function (r) { return String(r['Month']).slice(0, 7) === m && squadOf_(r['Squad']) === s; })[0]; };
  const events = readEventsTab_();
  const readings = frostReadings_(month, events);
  const prev = prevMonth_(month), sheet = getSheet_('Frost Log'), lines = [];
  SQUAD_KEYS_.forEach(function (s) {
    if (has(month, s) && !force) return;
    const list = readings[s], hit = list.filter(function (x) { return x.crossed; });
    const before = prev >= frostStart_() && has(prev, s) ? String(has(prev, s)['Markers Crossed'] || '').split(';').map(function (x) { return x.trim(); }).filter(String) : [];
    const repeat = hit.some(function (x) { return before.indexOf(x.marker) > -1; });
    const frozen = !!freezeDetail_(s, events);
    const status = !hit.length ? 'Safe' : frozen ? 'Still frozen' : hit.length >= 2 || repeat ? 'Freeze proposed' : 'Frost warning';
    const kind = !hit.length ? '' : hit.some(function (x) { return x.kind === 'Peaks'; }) ? 'Peaks' : 'Fun';
    const text = list.map(function (x) { return x.marker + ' ' + (x.value == null ? '—' : x.value) + (x.crossed ? ' (over the line ' + x.line + ')' : ''); }).join(' · ');
    sheet.appendRow([month, s, status, hit.map(function (x) { return x.marker; }).join('; '), text, kind, false, '']);
    lines.push(s + ' ' + month + ': ' + status + (hit.length ? ' (' + hit.map(function (x) { return x.marker; }).join(', ') + (repeat ? ', again' : '') + ')' : ''));
  });
  return lines;
}

// Thaw goals for a confirmed freeze, all communal.
function frostThawGoals_(kind, markers) {
  if (kind === 'Fun') return 'points: ' + (Number(getSetting_('FROST_FUN_POINTS')) || 100);
  return markers.map(function (m) { return 'marker: ' + m; }).concat([
    'everyone: ' + (Number(getSetting_('FROST_THAW_EVERYONE')) || 70) + '%',
    'points: ' + (Number(getSetting_('FROST_THAW_POINTS')) || 150),
    'A squad reflection or service act led by the Generals']).join('; ');
}

/** Freeze proposed + Approve ticked → a Point Freezes row. Returns lines. */
function applyApprovedFrost_() {
  const rows = sheetRows_('Frost Log').filter(function (r) { return String(r['Status']) === 'Freeze proposed' && ticked_(r['Approve']) && !String(r['Applied'] || '').trim(); });
  if (!rows.length) return [];
  ensureFreezeTab_();
  const today = todayYmd_(), log = getSheet_('Frost Log'), fz = getSheet_('Point Freezes'), lines = [];
  const appliedCol = FROST_LOG_HEADERS_.indexOf('Applied') + 1;
  rows.forEach(function (r) {
    const s = squadOf_(r['Squad']), kind = String(r['Kind']) === 'Fun' ? 'Fun' : 'Peaks';
    if (freezeDetail_(s)) { log.getRange(r._row, appliedCol).setValue('Already frozen ' + today); return; }
    const markers = String(r['Markers Crossed'] || '').split(';').map(function (x) { return x.trim(); }).filter(String);
    const nextEnd = monthWindow_(today.slice(0, 7)).days;
    const thawBy = kind === 'Fun' ? today.slice(0, 7) + '-' + ('0' + nextEnd).slice(-2) : '';
    const obj = { 'Squad': s, 'Start': today, 'End': '', 'Reason': 'Frost check ' + String(r['Month']).slice(0, 7) + ': ' + markers.join(', '), 'Set By': 'Frost check (approved)',
      'Kind': kind, 'Thaw Goals': frostThawGoals_(kind, markers), 'Thaw By': thawBy, 'Staff Sign-off': false, 'Thawed': '', 'Second Chance': '' };
    const head = fz.getRange(1, 1, 1, fz.getLastColumn()).getValues()[0].map(function (h) { return String(h).trim(); });
    fz.appendRow(head.map(function (h) { return obj[h] === undefined ? '' : obj[h]; }));
    log.getRange(r._row, appliedCol).setValue('Frozen ' + today);
    lines.push(s + ': ' + kind + ' Freeze started (' + markers.join(', ') + ')');
  });
  CacheService.getScriptCache().remove('thaw-check');
  return lines;
}

// "marker: X" thaw goal: X was on the safe side for a whole month that ended after the freeze began.
function markerSafe_(squad, marker, start) {
  const lm = lastMonthWindow_(todayYmd_()).month;
  if (lm < String(start).slice(0, 7)) return 0;
  const L = freezeLines_().filter(function (x) { return x.marker.toLowerCase() === String(marker).trim().toLowerCase(); })[0];
  if (!L) return 1; // the line was removed: nothing left to fix
  const v = lineReading_(L, lm, readEventsTab_(), squadSizes_(lm))[squad];
  return v != null && !crossed_(L, v) ? 1 : 0;
}

// From the background tick: the monthly check (once a month, after the 1st) and approvals.
function maybeFrostCheck_() {
  if (String(getSetting_('FROST_CHECK') || 'ON').toUpperCase() === 'OFF') return;
  const props = PropertiesService.getScriptProperties(), month = lastMonthWindow_(todayYmd_()).month;
  if (props.getProperty('frost-checked') !== month) { props.setProperty('frost-checked', month); runFrostCheck_(month); }
  applyApprovedFrost_();
}

function menuFrostCheck_() {
  const lines = runFrostCheck_(null).concat(applyApprovedFrost_());
  SpreadsheetApp.getUi().alert('Frost check', (lines.length ? lines.join('\n') : 'Already checked this month (see the Frost Log tab).') +
    '\n\nA Freeze proposed row freezes nothing until you tick Approve on the Frost Log tab.', SpreadsheetApp.getUi().ButtonSet.OK);
}

// ── WHAT STUDENTS SEE ───────────────────────────────────────────────
function frostUrl_() { try { return ScriptApp.getService().getUrl() + '?page=frost'; } catch (e) { return ''; } }

// One squad's frost meter: this month so far (counts on pace for the month) and the last check.
function frostFor_(squad, events) {
  if (String(getSetting_('FROST_CHECK') || 'ON').toUpperCase() === 'OFF') return null;
  const today = todayYmd_(), month = today.slice(0, 7);
  let r;
  try { r = frostReadings_(month, events, { today: today, pace: true })[squad]; } catch (e) { return null; }
  if (!r || !r.length) return null;
  const last = sheetRows_('Frost Log').filter(function (x) { return squadOf_(x['Squad']) === squad; }).pop();
  const start = frostStart_();
  return { month: month, live: month >= start, start: start, firstCheck: nextMonthFirst_(start),
    markers: r.map(function (x) { return { label: x.label, value: x.value, line: x.line, below: x.below, per100: x.per100, over: x.crossed }; }),
    over: r.filter(function (x) { return x.crossed; }).length,
    last: last ? { month: String(last['Month']).slice(0, 7), status: String(last['Status'] === 'Freeze proposed' ? (String(last['Applied'] || '') ? 'Frozen' : 'Frost warning') : last['Status']) } : null,
    url: frostUrl_() };
}
function nextMonthFirst_(month) { const y = Number(month.slice(0, 4)), m = Number(month.slice(5, 7)); return (m === 12 ? y + 1 : y) + '-' + ('0' + (m === 12 ? 1 : m + 1)).slice(-2) + '-01'; }

function frostMeter_(events) {
  events = events || readEventsTab_();
  const out = {};
  SQUAD_KEYS_.forEach(function (s) { out[s] = frostFor_(s, events); });
  return out;
}

/** The explainer page's data: the lines, the thaw goals, every squad's meter (squad totals only). */
function clashFrostExplainer() {
  const lines = freezeLines_();
  return JSON.stringify({
    lines: lines.map(function (L) { return { marker: L.marker, label: L.label, below: L.below, line: L.line, per100: L.per100, kind: L.kind, pct: L.measures === 'attendance' || L.measures === 'core f' }; }),
    thaw: { everyone: Number(getSetting_('FROST_THAW_EVERYONE')) || 70, points: Number(getSetting_('FROST_THAW_POINTS')) || 150, fun: Number(getSetting_('FROST_FUN_POINTS')) || 100 },
    start: frostStart_(), firstCheck: nextMonthFirst_(frostStart_()), meter: frostMeter_(), script: frostAnnouncement_(),
  });
}

function frostAnnouncement_() {
  const lines = freezeLines_();
  const say = lines.map(function (L) { return (L.below ? 'at least ' : 'no more than ') + L.line + (L.per100 ? ' per 100 students' : L.measures === 'attendance' || L.measures === 'core f' ? '%' : '') + ' for ' + L.label.toLowerCase(); });
  return 'Mountaineers, here is how Freeze and Thaw works in Clash of Classes. Every month, each squad is checked on the same lines: ' + say.join('; ') + '. ' +
    'Cross one line and your squad gets a frost warning. Cross two, or the same one two months in a row, and your squad can freeze: you keep every point you have, ' +
    'but new points stop counting. There is only one way out, and it is together. A frozen squad thaws when it gets back under the line, when most of the squad pitches in, ' +
    'when it earns points together, and when the Generals lead a squad reflection or service act. No one thaws alone. Watch your squad\'s frost meter on My Squad and My Climb at clashofclasses.org.';
}

function serveFrostPage_() {
  return HtmlService.createHtmlOutputFromFile('FrostPage').setTitle('Freeze & Thaw · Clash of Classes')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function menuFrostAnnouncement_() { SpreadsheetApp.getUi().alert('Freeze & Thaw (read aloud)', frostAnnouncement_(), SpreadsheetApp.getUi().ButtonSet.OK); }
function menuOpenFrostPage_() {
  SpreadsheetApp.getUi().alert('Freeze & Thaw slides', 'For squad meetings (arrow keys or click to move; F for full screen):\n\n' + frostUrl_() +
    '\n\nAdd &view=tv for the frost meter alone (for a TV).', SpreadsheetApp.getUi().ButtonSet.OK);
}
