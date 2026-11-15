/**
 * Clash of Classes -- Apps Script backend
 * File 34: Admin › 📋 Sheet. Edit the Sheet's setup tabs from the admin panel, no spreadsheet needed.
 *
 * EVERY tab of the Sheet opens here. SHEET_TABS_ gives the setup tabs a group, a hint and dropdowns;
 * any other tab shows under Other. Tabs with students' 800#s or names sit in their own group,
 * "Student data (private)". Left out: app_data (the site's saved pages; the other admin tabs edit it)
 * and the guide tabs Run setup rewrites (How To Use This Sheet, How The Map Works, How My Climb Works).
 *   clashSheetTabs(token)            the tabs that exist, with row counts and a hint
 *   clashSheetLoad(token, tab)       headers + rows (dates as yyyy-mm-dd, checkboxes as true/false)
 *   clashSheetSave(token, tab, json) { edits: [{ row, before, values }], adds: [values], deletes: [{ row, before }] }
 *     A row is only changed or deleted if it still holds what the page loaded ("before"), so an edit
 *     made in the Sheet meanwhile is never overwritten; those rows come back as conflicts.
 *   clashSheetAction(token, name)    the menu jobs that go with a tab (next Pulse, schedule ideas,
 *                                    You Said results, Poll Insights, points balance, privacy check)
 * Admin only (the panel's sign-in). Every save is noted on the Admin Edits tab (who/when/what tab).
 */

const SHEET_TABS_ = [
  // [Tab, Group, Hint, { column: [choices] }]
  ['Daily Polls', 'Polls & Pulse', 'One row per poll. Kind: This or That, Trivia (type the Answer), Pulse (no right answer; Through = last day open) or a Check-in (Pride Day Check-in, Wheel Check-in). Show must be ticked.',
    { 'Kind': ['This or That', 'Trivia', 'Pulse', 'Pride Day Check-in', 'Wheel Check-in'] }],
  ['Poll Ideas', 'Polls & Pulse', 'The question bank. Untick Use to skip one; clear Used On to run it again. Repeat "monthly" brings it back every 4 weeks.', { 'Kind': ['Pulse', 'School Pulse', 'This or That', 'Trivia', 'Pride Day Check-in', 'Wheel Check-in'], 'Repeat': ['', 'monthly'] }],
  ['You Said We Did', 'Polls & Pulse', 'Closed Pulse results land here. Type What We Did (and Who), tick Show: students see it on the Pulse page and the TV.', {}],
  ['Squad Monthly', 'Map', 'One row per squad per month (Month like 2026-12). Climb-er fills most columns; LRR + LD comes from the Follow-Up sheets. Type to fill a gap.', { 'Squad': ['BLACK', 'GOLD', 'GREY', 'PURPLE'] }],
  ['Zone Overrides', 'Map', 'Replace one squad\'s number for an area for a month. Newest row wins; a blank Score clears it. (Admin › Map does this too.)', { 'Squad': ['BLACK', 'GOLD', 'GREY', 'PURPLE'] }],
  ['Conquest Rules', 'Map', 'The ways to win an area. Admin › Map edits these with a picker.', { 'Method': ['most points', 'most entries', 'staff pick', 'highest number', 'lowest number'] }],
  ['Map Zones', 'Map', 'Every area: rule, Weekly Points, Multiplier, Locked. Admin › Map is easier for position and size.', {}],
  ['Freeze Lines', 'Freeze & Thaw', 'The same lines for every squad. One crossed = frost warning; two, or the same one two months running = freeze proposed.', { 'Freezes If': ['Above', 'Below'], 'Kind': ['Peaks', 'Fun'] }],
  ['Frost Log', 'Freeze & Thaw', 'The monthly check. Tick Approve on a "Freeze proposed" row to start that freeze.', {}],
  ['Point Freezes', 'Freeze & Thaw', 'Freezes now running (Squad, Start, End). Thaw goals are on the row.', { 'Squad': ['BLACK', 'GOLD', 'GREY', 'PURPLE'] }],
  ['Challenges', 'Events', 'Automatic challenges: progress counts from the events tab.', {}],
  ['Squad Duels', 'Events', 'Two squads, one measure, a start and end date.', {}],
  ['Sign-up Events', 'Events', 'Events students sign up for. Kind (Move, Make, Talk) + Date opens check-in that day.', { 'Kind': ['', 'Move', 'Make', 'Talk'] }],
  ['Year Plan', 'Events', 'The year, month by month. Shows on the site and the staff Today page.', {}],
  ['School Calendar', 'Events', 'Days with no school (one row per day).', {}],
  ['Homeroom Forms', 'Events', 'Spirit-day forms that score homerooms.', {}],
  ['Homeroom Points', 'Events', 'Homeroom results. Tick Approved to count a row.', {}],
  ['Event Pages', 'Pages', 'Mayhem, Kick Off and Champs page text.', {}],
  ['Event Rows', 'Pages', 'Schedule, events and links on those pages.', {}],
  ['Lore Chapters', 'Pages', 'Squad story chapters and when they unlock.', { 'Squad': ['BLACK', 'GOLD', 'GREY', 'PURPLE'] }],
  ['Squad Themes', 'Pages', 'Each squad\'s story and theme.', {}],
  ['Mindset Links', 'Pages', 'Mindset lesson links by month.', {}],
  ['Tutorial Steps', 'Pages', 'Cliff\'s tour, step by step.', {}],
  ['Opponents', 'Game', 'The Opponent each month and the goal to beat it.', {}],
  ['Round Goals', 'Game', 'Round goals and their prizes.', {}],
  ['Engine Powers', 'Game', 'Powers squads unlock.', {}],
  ['Chance Cards', 'Game', 'The chance card deck.', {}],
  ['Evolution Boosts', 'Game', 'My Climb evolution boosts.', {}],
  ['Perks', 'Game', 'My Climb perks.', {}],
  ['Settings', 'Settings', 'Every setting with its note. Change Value only.', {}],
  ['LRG Traits', 'Pages', 'The 9 LRG traits on the Trait page, the home page and the TV. One row per skill; the first row of a month has the Trait and its statement.', {}],
  ['Wheel Slices', 'Pages', 'The Spin Wheel, slice by slice in order. Kind pts = points; special = a named slice (NONE, LOYALTY). Loyalty = pays only that color.', { 'Color': ['BLACK', 'GOLD', 'GREY', 'PURPLE'], 'Kind': ['pts', 'special'] }],
  ['events', 'Points', 'Every point given (one row per event). Admin › Points adds rows the easy way; edit here to fix one.', {}],
  ['metrics', 'Points', 'Attendance, referral and grades per squad (the Metrics tab edits these too).', {}],
  ['Squad Board', 'Points', 'Written by Climb-er every hour: an edit here is replaced on its next sync.', {}],
  ['Map Battles', 'Map', 'Every time an area changed hands.', {}],
  ['Destination Tickets', 'Game', 'The destination tickets students draw.', {}],
  ['Mayhem Checklist', 'Events', 'Each squad\'s Mayhem checklist.', {}],
  ['Shirt Designs', 'Events', 'Shirt designs; tick Approved to put one up for the vote.', {}],
  ['Poll Insights', 'Polls & Pulse', 'Answers by squad and grade (rebuilt by Update Poll Insights).', {}],
  ['Points Balance', 'Polls & Pulse', 'Rebuilt by the points balance check.', {}],
  ['Squad Spotlight', 'Pages', 'Suggested every Monday (one story per squad). Tick Show to put it on the Morning Clash, the TV and the home page; edit the story if you like.', { 'Squad': ['BLACK', 'GOLD', 'GREY', 'PURPLE'] }],
  ['Avalanche', 'Game', 'Every Avalanche set off: squad, marking period, the week it ran, the bonus paid.', { 'Squad': ['BLACK', 'GOLD', 'GREY', 'PURPLE'] }],
  ['Family Challenges', 'Families', 'The weekly family challenge (FAMILY_CHALLENGE ON): the newest shown row whose Week Of has started.', {}],
  ['Family Pulse', 'Families', 'A question for families (FAMILY_PULSE ON): Choices like Yes | No; From / Through dates (blank = open).', {}],
  ['Plinko Slots', 'Game', 'The Plinko prizes, left to right (one more slot than PLINKO_ROWS). Kind: points, mystery (Value like 5|10|25|50) or nothing. The middle slots come up most; the edges are rare.', { 'Kind': ['points', 'mystery', 'nothing'] }],
  ['Arcade Plays', 'Game', 'Every Arcade game on My Climb. Points add up once a day as one Arcade row per squad (Paid fills in). Delete a row to undo a play.', { 'Game': ['dash', 'catch', 'stack'] }],
  ['Plinko Drops', 'Game', 'Drops squads have earned (a row each). Add one here or with + Give a drop on the board; Used fills in when it drops.', { 'Squad': ['BLACK', 'GOLD', 'GREY', 'PURPLE'] }],
  ['Dough Teachers', 'Points', 'Teachers who can give Virtual Trait Dough: tick Active, their squad, and their password (e.g. their room number).', { 'Squad': ['BLACK', 'GOLD', 'GREY', 'PURPLE'] }],
];
const SHEET_HIDDEN_ = ['app_data', 'How To Use This Sheet', 'How The Map Works', 'How My Climb Works'];
const SHEET_PRIVATE_ = ['Students', 'Student Board', 'Poll Votes', 'Event Check-ins', 'Redemptions', 'Shirt Votes', 'Shirt Orders', 'Squad Night', 'Sign-ups',
  'Nominations', 'General Letters', 'Virtual Dough', 'Service Submissions', 'Family Shout-outs', 'Family Pulse Answers', 'Family Check-ins', 'MLC Roster', 'Boost Ledger', 'Student Tickets', 'Avatars', 'Color Squads', 'Homeroom Points', 'Frost Log', 'Admin Edits', 'Usage', 'Weekly Digest Log', 'Squad Hand', 'Streak Banks'];
const SHEET_EDIT_LOG_ = 'Admin Edits';

function sheetTabInfo_(name) {
  name = String(name || '');
  if (!name || SHEET_HIDDEN_.indexOf(name) > -1 || !SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name)) throw new Error('That tab is not editable here.');
  const t = SHEET_TABS_.filter(function (x) { return x[0] === name; })[0];
  if (t) return { tab: t[0], group: SHEET_PRIVATE_.indexOf(name) > -1 ? 'Student data (private)' : t[1], hint: t[2], choices: t[3] || {} };
  return { tab: name, group: SHEET_PRIVATE_.indexOf(name) > -1 ? 'Student data (private)' : 'Other', hint: SHEET_PRIVATE_.indexOf(name) > -1 ? 'Has students\' 800#s or names: staff only. Most rows here are written by the site or Climb-er.' : '', choices: {} };
}

// A cell as the page shows it: dates yyyy-mm-dd (or with the time when it has one), the rest as-is.
function sheetCellOut_(v, tz) {
  if (Object.prototype.toString.call(v) === '[object Date]') {
    const t = Utilities.formatDate(v, tz, 'HH:mm');
    return Utilities.formatDate(v, tz, 'yyyy-MM-dd') + (t === '00:00' ? '' : ' ' + t);
  }
  return v === null || v === undefined ? '' : v;
}

// What the page sends back, as a cell: true/false stay booleans, plain numbers become numbers.
function sheetCellIn_(v) {
  if (v === true || v === false) return v;
  const s = String(v == null ? '' : v);
  if (/^-?\d+(\.\d+)?$/.test(s.trim()) && s.trim().length < 12 && !/^0\d/.test(s.trim())) return Number(s.trim());
  return s;
}

function sheetSame_(a, b) { return String(a == null ? '' : a) === String(b == null ? '' : b); }

function clashSheetTabs(token) {
  requireAdmin_(token);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const order = ['Polls & Pulse', 'Map', 'Freeze & Thaw', 'Points', 'Events', 'Families', 'Pages', 'Game', 'Settings', 'Other', 'Student data (private)'];
  const known = SHEET_TABS_.map(function (t) { return t[0]; });
  const names = known.filter(function (n) { return ss.getSheetByName(n); })
    .concat(ss.getSheets().map(function (s) { return s.getName(); }).filter(function (n) { return known.indexOf(n) === -1 && SHEET_HIDDEN_.indexOf(n) === -1; }));
  const list = names.map(function (n) { const i = sheetTabInfo_(n); return { tab: n, group: i.group, hint: i.hint, rows: Math.max(0, ss.getSheetByName(n).getLastRow() - 1) }; });
  list.sort(function (a, b) { return order.indexOf(a.group) - order.indexOf(b.group); });
  return JSON.stringify(list);
}

function clashSheetLoad(token, tab) {
  requireAdmin_(token);
  const info = sheetTabInfo_(tab);
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(info.tab);
  if (!sh) throw new Error('There is no ' + info.tab + ' tab yet (Run setup makes it).');
  const tz = Session.getScriptTimeZone();
  const lastCol = sh.getLastColumn(), lastRow = sh.getLastRow();
  if (!lastCol) return JSON.stringify({ tab: info.tab, hint: info.hint, headers: [], rows: [], choices: info.choices, bools: [] });
  const vals = sh.getRange(1, 1, Math.max(1, lastRow), lastCol).getValues();
  const headers = vals[0].map(function (h) { return String(h).trim(); });
  const rows = [];
  for (let i = 1; i < vals.length; i++) {
    if (vals[i].every(function (v) { return v === '' || v === null; })) continue;
    rows.push({ row: i + 1, values: vals[i].map(function (v) { return sheetCellOut_(v, tz); }) });
  }
  // Checkbox columns: every filled cell is true/false.
  const bools = headers.map(function (h, c) {
    const filled = rows.map(function (r) { return r.values[c]; }).filter(function (v) { return v !== ''; });
    return filled.length > 0 && filled.every(function (v) { return v === true || v === false; });
  });
  // Daily Polls: what each row does on the site right now.
  if (info.tab === 'Daily Polls') {
    const byRow = {};
    pollsWithVotes_().forEach(function (p) { byRow[p.row] = p; });
    const today = todayYmd_();
    rows.forEach(function (r) {
      const p = byRow[r.row];
      r.status = !p ? '' : !p.show ? 'Hidden (Show is off)' : !p.question ? 'No question' : p.choices.length < (p.kind === 'Check-in' ? 1 : 2) ? 'Needs 2+ choices (A | B)'
        : !/^\d{4}-\d{2}-\d{2}$/.test(p.date) ? 'Needs a Date' : pollOpenOn_(p, today) ? 'LIVE today' : p.date > today ? 'Opens ' + p.date + ' (shown as coming up)' : 'Closed ' + p.through;
    });
  }
  return JSON.stringify({ tab: info.tab, hint: info.hint, headers: headers, rows: rows, choices: info.choices, bools: bools });
}

/** Saves the page's changes. Returns { edited, added, deleted, conflicts: [row numbers] }. */
function clashSheetSave(token, tab, json) {
  requireAdmin_(token);
  const info = sheetTabInfo_(tab);
  const p = JSON.parse(json || '{}');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(info.tab);
    if (!sh) throw new Error('There is no ' + info.tab + ' tab.');
    const tz = Session.getScriptTimeZone();
    const width = sh.getLastColumn();
    const head = sh.getRange(1, 1, 1, width).getValues()[0].map(function (h) { return String(h).trim(); });
    if (p.headers && p.headers.join('|') !== head.join('|')) throw new Error('The columns on the ' + info.tab + ' tab changed since you opened it. Reload the tab and try again.');
    const now = function (row) { return sh.getRange(row, 1, 1, width).getValues()[0].map(function (v) { return sheetCellOut_(v, tz); }); };
    const matches = function (row, before) { const cur = now(row); return (before || []).every(function (v, i) { return sheetSame_(cur[i], v); }); };
    const fit = function (values) { const out = []; for (let i = 0; i < width; i++) out.push(sheetCellIn_((values || [])[i])); return out; };
    const out = { edited: 0, added: 0, deleted: 0, conflicts: [] };
    (p.edits || []).forEach(function (e) {
      const row = Number(e.row);
      if (!(row >= 2) || row > sh.getLastRow()) { out.conflicts.push(row); return; }
      if (!matches(row, e.before)) { out.conflicts.push(row); return; }
      sh.getRange(row, 1, 1, width).setValues([fit(e.values)]);
      out.edited++;
    });
    // Deletes bottom-up so the row numbers above stay right.
    (p.deletes || []).slice().sort(function (a, b) { return Number(b.row) - Number(a.row); }).forEach(function (d) {
      const row = Number(d.row);
      if (!(row >= 2) || row > sh.getLastRow() || !matches(row, d.before)) { out.conflicts.push(row); return; }
      sh.deleteRow(row);
      out.deleted++;
    });
    (p.adds || []).forEach(function (values) {
      if (!(values || []).some(function (v) { return String(v == null ? '' : v).trim() !== '' && v !== false; })) return;
      sh.appendRow(fit(values));
      out.added++;
    });
    if (info.tab === 'Settings') { try { SETTINGS_MEMO_ = null; } catch (e) { /* no memo */ } }
    try {
      let log = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_EDIT_LOG_);
      if (!log) { log = SpreadsheetApp.getActiveSpreadsheet().insertSheet(SHEET_EDIT_LOG_); log.appendRow(['When', 'Tab', 'Edited', 'Added', 'Deleted', 'Skipped (changed in the Sheet)']); }
      if (out.edited || out.added || out.deleted || out.conflicts.length) log.appendRow([new Date(), info.tab, out.edited, out.added, out.deleted, out.conflicts.join(', ')]);
    } catch (e) { console.warn('Admin Edits log: ' + e.message); }
    return JSON.stringify(out);
  } finally { lock.releaseLock(); }
}

// The menu jobs, from the panel. Each returns a line to show.
const SHEET_ACTIONS_ = {
  nextPulse: ['Daily Polls', 'Add the next Pulse question', function () { const r = addNextPulse_(); return r ? 'Added "' + r.question + '", open ' + r.start + ' through ' + r.through + '.' : 'No Pulse idea is ready on Poll Ideas (Use ticked, Used On blank).'; }],
  schedule: ['Daily Polls', 'Schedule poll ideas (next 5 school days)', function () { const l = schedulePollIdeas_(5); return l.length ? 'Scheduled:\n' + l.join('\n') : 'Nothing scheduled: those days already have polls, or no ticked idea is ready.'; }],
  announce: ['Daily Polls', 'Today\'s poll announcement', function () { return pollAnnouncement_(); }],
  insights: ['Poll Ideas', 'Update Poll Insights', function () { return 'Poll Insights updated: ' + writePollInsights_() + ' answer rows.'; }],
  youSaid: ['You Said We Did', 'Add new Pulse results', function () { const n = syncYouSaid_(); return n ? n + ' new Pulse result' + (n === 1 ? '' : 's') + ' added.' : 'No new Pulse results.'; }],
  followUp: ['Squad Monthly', 'Count LRRs + LDs for the Cafeteria now', function () { return sendFollowUpStats_(); }],
  frost: ['Frost Log', 'Run the frost check (last month) + apply approvals', function () { const l = runFrostCheck_(null).concat(applyApprovedFrost_()); return l.length ? l.join('\n') : 'Already checked this month.'; }],
  balance: ['Settings', 'Points balance check', function () { const r = writePointsBalance_(); return r.flags.length ? r.flags.join('\n') : 'Nothing lopsided.'; }],
  privacy: ['Settings', 'Privacy check (every public page)', function () { const bad = privacyScan_().filter(function (x) { return x.findings.length; }); return bad.length ? bad.map(function (x) { return x.surface + ': ' + x.findings.join(', '); }).join('\n') : 'No 800# and no full student name on any public page.'; }],
};

function clashSheetActions(token) {
  requireAdmin_(token);
  return JSON.stringify(Object.keys(SHEET_ACTIONS_).map(function (k) { return { id: k, tab: SHEET_ACTIONS_[k][0], label: SHEET_ACTIONS_[k][1] }; }));
}

function clashSheetAction(token, name) {
  requireAdmin_(token);
  const a = SHEET_ACTIONS_[name];
  if (!a) throw new Error('Unknown action.');
  return JSON.stringify({ message: String(a[2]() || 'Done.') });
}

// file-stamp (tools/stamp-files.js): lets Admin › System tell whether this file is the newest one
var FILE_STAMP_34_SheetEditor = '85ad849d11';
