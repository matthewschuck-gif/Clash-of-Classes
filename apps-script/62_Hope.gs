/**
 * Clash of Classes -- Apps Script backend
 * File 62: My Hope, the hope tracker on My Climb (the Hope tab on ?page=me).
 *
 * Hope = a goal, ways to get there (pathways) and the belief you can (agency). Each student:
 *   - writes ONE goal and TWO pathways (two different ways to reach it), and can change them;
 *   - checks in once a week: "How sure are you that you can move toward your goal this week?" 1-5;
 *   - answers the six-question Children's Hope Scale once each marking period (1 = None of the
 *     time ... 6 = All of the time). Agency = questions 1, 3, 5; Pathways = 2, 4, 6; Total 6-36.
 * Marking periods come from Settings › QUARTER_STARTS (the same dates as the quarter recap);
 * blank = Aug-Oct MP1, Nov-Jan MP2, Feb-Mar MP3, Apr-Jul MP4.
 *
 * Tabs (made on first use and by Run setup): Hope Goals, Hope Checks, Hope Scale, and Hope Scale
 * Items (the six questions, editable: check the wording against your copy of the scale).
 * Climb-er reads Hope Scale and Hope Checks (its Sources tab) so staff see the scores on the
 * student's page. Students only ever see their own answers. HOPE = OFF hides the tab.
 * No squad points: this is personal, not a contest.
 */

const HOPE_GOAL_HEADERS_ = ['800#', 'Squad', 'Goal', 'Pathway 1', 'Pathway 2', 'Set', 'Updated'];
const HOPE_CHECK_HEADERS_ = ['800#', 'Squad', 'Week Of', 'Agency', 'At'];
const HOPE_SCALE_HEADERS_ = ['800#', 'Squad', 'Period', 'Q1', 'Q2', 'Q3', 'Q4', 'Q5', 'Q6', 'Agency', 'Pathways', 'Total', 'At'];
const HOPE_ITEM_HEADERS_ = ['#', 'Question', 'Kind'];
const HOPE_ITEMS_ = [
  [1, 'I think I am doing pretty well.', 'Agency'],
  [2, 'I can think of many ways to get the things in life that are most important to me.', 'Pathways'],
  [3, 'I am doing just as well as other kids my age.', 'Agency'],
  [4, 'When I have a problem, I can come up with lots of ways to solve it.', 'Pathways'],
  [5, 'I think the things I have done in the past will help me in the future.', 'Agency'],
  [6, 'Even when others want to quit, I know that I can find ways to solve the problem.', 'Pathways'],
];
const HOPE_ANSWERS_ = ['None of the time', 'A little of the time', 'Some of the time', 'A lot of the time', 'Most of the time', 'All of the time'];
const HOPE_AGENCY_WORDS_ = ['Not sure at all', 'A little sure', 'Kind of sure', 'Pretty sure', 'Totally sure'];

function hopeOn_() { return String(getSetting_('HOPE') || 'ON').toUpperCase() !== 'OFF'; }

function ensureHopeTabs_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  [['Hope Goals', HOPE_GOAL_HEADERS_, null], ['Hope Checks', HOPE_CHECK_HEADERS_, null], ['Hope Scale', HOPE_SCALE_HEADERS_, null],
   ['Hope Scale Items', HOPE_ITEM_HEADERS_, HOPE_ITEMS_]].forEach(function (t) {
    if (ss.getSheetByName(t[0])) return;
    const sh = ss.insertSheet(t[0]);
    sh.getRange(1, 1, 1, t[1].length).setValues([t[1]]).setFontWeight('bold').setBackground('#490e6f').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    if (t[2]) sh.getRange(2, 1, t[2].length, t[1].length).setValues(t[2]);
  });
}

function hopeItems_() {
  const rows = sheetRows_('Hope Scale Items').filter(function (r) { return String(r['Question'] || '').trim(); });
  const src = rows.length === 6 ? rows.map(function (r, i) { return [i + 1, String(r['Question']).trim(), String(r['Kind'] || '').trim()]; }) : HOPE_ITEMS_;
  return src.map(function (r) { return { n: r[0], q: r[1], kind: /path/i.test(r[2]) ? 'Pathways' : 'Agency' }; });
}

// The marking period a date falls in: { name: 'MP2', start: 'yyyy-mm-dd' }.
function hopePeriod_(ymd) {
  ymd = ymd || todayYmd_();
  const starts = String(getSetting_('QUARTER_STARTS') || '').split(',').map(function (s) { return s.trim(); })
    .filter(function (s) { return /^\d{4}-\d{2}-\d{2}$/.test(s); }).sort();
  if (starts.length) {
    let i = -1;
    starts.forEach(function (s, k) { if (s <= ymd) i = k; });
    if (i > -1) return { name: 'MP' + (i + 1), start: starts[i] };
    return { name: 'MP1', start: starts[0] };
  }
  const y = Number(ymd.slice(0, 4)), m = Number(ymd.slice(5, 7));
  const sy = m >= 8 ? y : y - 1; // the school year starts in August
  const cuts = [[sy + '-08-01', 'MP1'], [sy + '-11-01', 'MP2'], [(sy + 1) + '-02-01', 'MP3'], [(sy + 1) + '-04-01', 'MP4']];
  let p = cuts[0];
  cuts.forEach(function (c) { if (c[0] <= ymd) p = c; });
  return { name: p[1], start: p[0] };
}

// Monday of this week, yyyy-mm-dd.
function hopeWeekOf_(d) {
  const ymd = Utilities.formatDate(d ? new Date(d) : new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const noon = new Date(ymd + 'T12:00:00Z'), back = (noon.getUTCDay() + 6) % 7; // days since Monday
  return new Date(noon.getTime() - back * 86400000).toISOString().slice(0, 10);
}

function hopeYmd_(v) {
  if (Object.prototype.toString.call(v) === '[object Date]') return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return String(v || '').slice(0, 10);
}

// Never a formula from a student's typing.
function hopeText_(v, max) {
  const t = String(v == null ? '' : v).replace(/\s+/g, ' ').trim().slice(0, max || 200);
  return /^[=+\-@]/.test(t) ? '\'' + t : t;
}

/** Everything the Hope tab shows for one student (only their own rows). */
function hopeFor_(id) {
  if (!hopeOn_()) return null;
  const mine = function (tab) { return sheetRows_(tab).filter(function (r) { return id800_(r['800#']) === id; }); };
  const g = mine('Hope Goals')[0] || {};
  const week = hopeWeekOf_(), period = hopePeriod_();
  const checks = mine('Hope Checks').map(function (r) { return { week: hopeYmd_(r['Week Of']), agency: Number(r['Agency']) || 0 }; })
    .filter(function (c) { return c.agency; }).sort(function (a, b) { return a.week < b.week ? -1 : 1; });
  const scales = mine('Hope Scale').map(function (r) {
    return { period: String(r['Period'] || ''), agency: Number(r['Agency']) || 0, pathways: Number(r['Pathways']) || 0, total: Number(r['Total']) || 0, at: hopeYmd_(r['At']) };
  });
  return {
    goal: { goal: String(g['Goal'] || ''), p1: String(g['Pathway 1'] || ''), p2: String(g['Pathway 2'] || ''), set: hopeYmd_(g['Set']) },
    week: week, thisWeek: (checks.filter(function (c) { return c.week === week; })[0] || {}).agency || 0,
    checks: checks.slice(-10), agencyWords: HOPE_AGENCY_WORDS_,
    period: period.name, scaleDone: scales.some(function (s) { return s.period === period.name; }),
    scales: scales, items: hopeItems_(), answers: HOPE_ANSWERS_,
  };
}

function hopeSaveRow_(tab, headers, id, match, row) {
  ensureHopeTabs_();
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const sheet = getSheet_(tab);
    const r = sheetRows_(tab).filter(function (x) { return id800_(x['800#']) === id && match(x); })[0];
    if (r) sheet.getRange(r._row, 1, 1, headers.length).setValues([row]); else sheet.appendRow(row);
  } finally { lock.releaseLock(); }
}

function hopeWho_(asId, adminToken) {
  if (!hopeOn_()) throw new Error('My Hope is turned off.');
  const id = meId_(asId, adminToken); // 16_Me.gs: the signed-in student, or staff preview
  noSampleSave_(id);
  const row = boardRows_().filter(function (r) { return id800_(r['800#']) === id; })[0] || {};
  return { id: id, squad: squadOf_(row['Squad']) || '' };
}

/** Goal + two pathways. */
function clashMeHopeGoal(goal, p1, p2, asId, adminToken) {
  const w = hopeWho_(asId, adminToken);
  const g = hopeText_(goal, 200), a = hopeText_(p1, 200), b = hopeText_(p2, 200);
  if (g.length < 3) throw new Error('Write your goal first.');
  if (a.length < 3 || b.length < 3) throw new Error('Write two different ways to get there.');
  if (a.toLowerCase() === b.toLowerCase()) throw new Error('Make your two pathways different: if one gets blocked, the other still works.');
  const old = sheetRows_('Hope Goals').filter(function (x) { return id800_(x['800#']) === w.id; })[0];
  hopeSaveRow_('Hope Goals', HOPE_GOAL_HEADERS_, w.id, function () { return true; },
    [w.id, w.squad, g, a, b, old && old['Set'] ? hopeYmd_(old['Set']) : todayYmd_(), todayYmd_()]);
  return JSON.stringify(hopeFor_(w.id));
}

/** This week's agency check, 1-5 (once a week; answering again changes it). */
function clashMeHopeCheck(agency, asId, adminToken) {
  const w = hopeWho_(asId, adminToken);
  const n = Math.round(Number(agency));
  if (!(n >= 1 && n <= 5)) throw new Error('Pick 1 to 5.');
  const week = hopeWeekOf_();
  hopeSaveRow_('Hope Checks', HOPE_CHECK_HEADERS_, w.id, function (x) { return hopeYmd_(x['Week Of']) === week; },
    [w.id, w.squad, week, n, new Date()]);
  return JSON.stringify(hopeFor_(w.id));
}

/** The six-question Hope Scale for this marking period. answers: six numbers 1-6, in order. */
function clashMeHopeScale(answers, asId, adminToken) {
  const w = hopeWho_(asId, adminToken);
  const a = (Array.isArray(answers) ? answers : []).map(function (x) { return Math.round(Number(x)); });
  if (a.length !== 6 || a.some(function (x) { return !(x >= 1 && x <= 6); })) throw new Error('Answer all six.');
  const items = hopeItems_();
  let agency = 0, pathways = 0;
  items.forEach(function (it, i) { if (it.kind === 'Pathways') pathways += a[i]; else agency += a[i]; });
  const period = hopePeriod_().name;
  hopeSaveRow_('Hope Scale', HOPE_SCALE_HEADERS_, w.id, function (x) { return String(x['Period']) === period; },
    [w.id, w.squad, period].concat(a).concat([agency, pathways, agency + pathways, new Date()]));
  return JSON.stringify(hopeFor_(w.id));
}

// file-stamp (tools/stamp-files.js): lets Admin › System tell whether this file is the newest one
var FILE_STAMP_62_Hope = 'b553963a3d';
