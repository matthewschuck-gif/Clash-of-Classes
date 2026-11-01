// Offline checks for the Clash of Classes Apps Script backend: fake Google services, real .gs code.
//   node tests/clash.test.js
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const assert = require('assert');

const dir = path.join(__dirname, '..', 'apps-script');
let passed = 0;
function test(name, fn) { fn(); passed++; console.log('ok - ' + name); }

function makeSheet(name, rows) {
  const sheet = {
    rows: rows.map(r => r.slice()),
    getName: () => name,
    getLastRow() { return this.rows.length; },
    getLastColumn() { return this.rows.reduce((m, r) => Math.max(m, r.length), 0); },
    getDataRange() { const s = this; return { getValues: () => s.rows.map(r => r.slice()) }; },
    getRange(r, c, nr, nc) {
      const s = this; nr = nr || 1; nc = nc || 1;
      let proxy;
      const api = {
        getValues: () => { const out = []; for (let i = 0; i < nr; i++) { const row = s.rows[r - 1 + i] || []; out.push(Array.from({ length: nc }, (_, j) => row[c - 1 + j] == null ? '' : row[c - 1 + j])); } return out; },
        setValues(vals) { vals.forEach((row, i) => { s.rows[r - 1 + i] = s.rows[r - 1 + i] || []; row.forEach((v, j) => { s.rows[r - 1 + i][c - 1 + j] = v; }); }); return proxy; },
        setValue(v) { return api.setValues([[v]]); },
        clearContent() { for (let i = 0; i < nr; i++) { const row = s.rows[r - 1 + i]; if (row) for (let j = 0; j < nc; j++) row[c - 1 + j] = ''; } return proxy; },
      };
      proxy = new Proxy(api, { get: (t, p) => p in t ? t[p] : () => proxy });
      return proxy;
    },
    appendRow(row) { this.rows.push(row.slice()); },
    deleteRow(n) { this.rows.splice(n - 1, 1); },
    setFrozenRows() {}, setColumnWidth() {},
  };
  return sheet;
}

function world(opts) {
  opts = opts || {};
  const blob = Object.assign({ password: 'oldpw', poll: { active: true, question: 'Best squad?', options: ['A', 'B'], votes: { 0: 2 } }, metrics: { month: 'Oct' }, nextId: 103 }, opts.blob || {});
  const json = JSON.stringify(blob);
  const sheets = [
    makeSheet('app_data', [['key', 'chunk_index', 'value_chunk', 'updated_at'], ['coc', 0, json, '2026-10-01']]),
    makeSheet('events', [['id', 'name', 'type', 'date', 'note', 'multiplier', 'BLACK', 'GOLD', 'GREY', 'PURPLE'],
      [101, 'Kickoff', 'event', '2026-09-01', '', '', 10, 0, 0, 0],
      [102, 'Bill', 'trait_dough', '2026-10-01', '', '', 0, 5, 0, 0]]),
    makeSheet('metrics', [['squad', 'attendance', 'referral', 'grades'], ['BLACK', 90, 80, 70], ['GOLD', 91, 81, 71], ['GREY', 92, 82, 72], ['PURPLE', 93, 83, 73]]),
  ];
  const book = {
    getSheets: () => sheets,
    getSheetByName: n => sheets.find(s => s.getName() === n) || null,
    insertSheet(n) { const s = makeSheet(n, []); sheets.push(s); return s; },
    getId: () => 'CLASH', deleteSheet() {}, setActiveSheet() {}, moveActiveSheet() {},
  };
  const tj = { getSheetByName: n => n === 'reflections' ? makeSheet(n, [['student_name', 'insight_level', 'created_at'], ['Ann', 'High', new Date('2026-10-02T12:00:00Z')]])
    : n === 'panel_sessions' ? makeSheet(n, [['clash_squad', 'status'], ['GOLD', 'complete']]) : null };
  const cache = {};
  const props = {};
  const ctx = {
    console, JSON, Math, Date, Object, Array, String, Number,
    SpreadsheetApp: { getActiveSpreadsheet: () => book, openById: id => { if (id !== '1rinTrtlmJKQrvtn5GVs-ngmaQRvkl8Nt39yOBICgggA') throw new Error('no ' + id); return tj; }, getUi: () => ({ alert() {} }),
      newDataValidation: () => { const b = new Proxy({}, { get: () => () => b }); return b; } },
    DriveApp: { getFileById: () => ({ getLastUpdated: () => new Date('2026-10-31T10:00:00Z') }) },
    CacheService: { getScriptCache: () => ({ get: k => cache[k] || null, put: (k, v) => { cache[k] = v; } }) },
    PropertiesService: { getScriptProperties: () => ({ getProperty: k => props[k] || null, setProperty: (k, v) => { props[k] = v; } }) },
    LockService: { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) },
    Utilities: {
      getUuid: (() => { let n = 0; return () => 'uuid-' + (++n); })(),
      computeDigest: (alg, s) => Array.from(require('crypto').createHash('sha256').update(s).digest()).map(b => b > 127 ? b - 256 : b),
      DigestAlgorithm: { SHA_256: 'sha256' }, Charset: { UTF_8: 'utf8' },
      formatDate: d => d.toISOString().slice(0, 10),
    },
    Session: { getScriptTimeZone: () => 'America/New_York' },
    HtmlService: {}, ContentService: {},
  };
  vm.createContext(ctx);
  const src = fs.readdirSync(dir).filter(f => f.endsWith('.gs')).sort().map(f => fs.readFileSync(path.join(dir, f), 'utf8')).join('\n');
  vm.runInContext(src, ctx);
  ctx.book = book; ctx.props = props;
  return ctx;
}

const events = w => w.book.getSheetByName('events').rows.slice(1);
const J = x => JSON.parse(JSON.stringify(x)); // values from the sandbox → plain data
const login = (w, pw) => JSON.parse(w.clashLogin(pw)).token;

test('load leaves the password out', () => {
  const w = world();
  const v = JSON.parse(w.clashLoad()).value;
  assert.strictEqual(v.password, undefined);
  assert.strictEqual(v.events.length, 2);
  assert.deepStrictEqual(v.metrics.attendance, { BLACK: 90, GOLD: 91, GREY: 92, PURPLE: 93 });
});

test('sign-in: old site password works once and is then stored only as a hash', () => {
  const w = world();
  assert.throws(() => w.clashLogin('nope'), /Incorrect/);
  assert.ok(login(w, 'oldpw'));
  assert.ok(w.props.ADMIN_PW_HASH && w.props.ADMIN_PW_HASH !== 'oldpw');
  assert.ok(!JSON.stringify(w.props).includes('oldpw'), 'never stored as plain text');
  const t = login(w, 'oldpw');
  w.clashSetPassword(t, 'newpass');
  assert.throws(() => w.clashLogin('oldpw'), /Incorrect/);
  assert.throws(() => w.clashLogin('ephrata2026'), /Incorrect/, 'built-in password no longer works once one is set');
  assert.ok(login(w, 'newpass'));
});

test('too many wrong passwords locks sign-in for a while', () => {
  const w = world();
  for (let i = 0; i < 10; i++) assert.throws(() => w.clashLogin('x' + i));
  assert.throws(() => w.clashLogin('oldpw'), /Too many/);
});

test('save needs an admin session', () => {
  const w = world();
  assert.throws(() => w.clashSave('', '{}', '{}'), /sign in/);
  assert.throws(() => w.clashSave('made-up', '{}', '{}'), /sign in/);
});

test('save keeps points another project added after the page loaded, and drops ones the admin deleted', () => {
  const w = world();
  const t = login(w, 'oldpw');
  const D = JSON.parse(w.clashLoad()).value;
  const sync = { seen: D.events.map(e => e.id + '|' + e.name), metrics: { attendance: D.metrics.attendance, referral: D.metrics.referral, grades: D.metrics.grades } };
  // Meanwhile Mindset appends a row (blank id) to the events tab.
  w.book.getSheetByName('events').appendRow(['', 'Mindset Summit', 'mountaineer_mindset', '2026-10-31', '', '', 0, 0, 3, 0]);
  // The admin deletes Kickoff and adds an event with the next id (103, the id the new row will also get).
  D.events = D.events.filter(e => e.name !== 'Kickoff');
  D.events.push({ id: 103, name: 'Pep rally', type: 'event', date: '2026-10-31', note: '', pts: { BLACK: 0, GOLD: 0, GREY: 0, PURPLE: 50 } });
  const res = JSON.parse(w.clashSave(t, JSON.stringify(D), JSON.stringify(sync)));
  assert.strictEqual(res.keptFromOthers, 1);
  const rows = events(w);
  assert.deepStrictEqual(rows.map(r => r[1]).sort(), ['Bill', 'Mindset Summit', 'Pep rally']);
  const ids = rows.map(r => r[0]);
  assert.strictEqual(new Set(ids).size, ids.length, 'no duplicate ids');
  assert.strictEqual(rows.find(r => r[1] === 'Mindset Summit')[8], 3);
});

test('save leaves the metrics tab alone unless the admin changed the numbers', () => {
  const w = world();
  const t = login(w, 'oldpw');
  const D = JSON.parse(w.clashLoad()).value;
  const sync = { seen: D.events.map(e => e.id + '|' + e.name), metrics: JSON.parse(JSON.stringify({ attendance: D.metrics.attendance, referral: D.metrics.referral, grades: D.metrics.grades })) };
  // The hourly squad sync writes new numbers after the page loaded.
  w.book.getSheetByName('metrics').rows[1][1] = 99;
  w.clashSave(t, JSON.stringify(D), JSON.stringify(sync));
  assert.strictEqual(w.book.getSheetByName('metrics').rows[1][1], 99, 'sync numbers kept');
  D.metrics.grades.GOLD = 55;
  w.clashSave(t, JSON.stringify(D), JSON.stringify(sync));
  assert.strictEqual(w.book.getSheetByName('metrics').rows[2][3], 55, 'admin edit written');
});

test('save never puts the password back into the site data', () => {
  const w = world();
  const t = login(w, 'oldpw');
  const D = JSON.parse(w.clashLoad()).value;
  D.password = 'sneaky';
  w.clashSave(t, JSON.stringify(D), '{}');
  assert.ok(!w.book.getSheetByName('app_data').rows.slice(1).map(r => r[2]).join('').includes('sneaky'));
});

test('poll vote: no sign-in, moves a changed vote, closed poll refuses', () => {
  const w = world();
  assert.deepStrictEqual(JSON.parse(w.clashVote('Best squad?', 1, -1)), { 0: 2, 1: 1 });
  assert.deepStrictEqual(JSON.parse(w.clashVote('Best squad?', 0, 1)), { 0: 3, 1: 0 });
  assert.throws(() => w.clashVote('Old question', 0, -1), /closed/);
  assert.throws(() => w.clashVote('Best squad?', 5, -1), /Unknown/);
  assert.strictEqual(events(w).length, 2, 'events untouched');
});

test('old API: load has no password and save is gone', () => {
  const w = world();
  w.props.APP_TOKEN = 'tok';
  let out;
  w.respondOut_ = o => { out = o; };
  w.handleRequest_({ parameter: { action: 'load', token: 'tok' } }, 'GET');
  assert.strictEqual(out.ok, true);
  assert.strictEqual(out.data.value.password, undefined);
  w.handleRequest_({ parameter: { payload: JSON.stringify({ action: 'save', token: 'tok', value: {} }) } }, 'POST');
  assert.strictEqual(out.ok, false);
  assert.ok(/Unknown action/.test(out.error));
});

test('Trail Journal stats read counts only, no names', () => {
  const w = world();
  w.ensureSettings_();
  const s = JSON.parse(w.clashTrailJournalStats());
  assert.deepStrictEqual(s.reflections, [{ insight_level: 'High', created_at: '2026-10-02T12:00:00.000Z' }]);
  assert.deepStrictEqual(s.panels, [{ clash_squad: 'GOLD', status: 'complete' }]);
  assert.ok(!JSON.stringify(s).includes('Ann'));
});

test('Settings tab: defaults added once, edits win, pasted links become IDs', () => {
  const w = world();
  w.ensureSettings_();
  const st = w.book.getSheetByName('Settings');
  assert.ok(st.rows.length > 3);
  st.rows.find(r => r[0] === 'POLL_SECONDS')[1] = '10';
  st.rows.find(r => r[0] === 'TJ_SHEET_ID')[1] = 'https://docs.google.com/spreadsheets/d/1rinTrtlmJKQrvtn5GVs-ngmaQRvkl8Nt39yOBICgggA/edit#gid=0';
  w.ensureSettings_();
  assert.strictEqual(w.getSetting_('POLL_SECONDS'), '10');
  assert.strictEqual(w.getSetting_('TJ_SHEET_ID'), '1rinTrtlmJKQrvtn5GVs-ngmaQRvkl8Nt39yOBICgggA');
});

function addStudentTabs(w) {
  const board = w.book.insertSheet('Student Board');
  board.rows.push(['800#', 'Name', 'Squad', 'Grade', 'Trait Dough Bills', 'Bills This Week', 'Trait Dough Wins', 'Mindset Summits', 'Trail Journal Growth', 'STAR Growth Points', 'Shout-outs', 'Streak (weeks)', 'Badges', 'Last Contribution', 'Updated'],
    ['8001234', 'Ann L.', 'Gold', '7', '12', '2', '1', '1', '', '4', '2', '3', '🍞 First Trait Dough, 🔥 3-Week Streak', '2026-10-30', '2026-10-31 19:00'],
    ['8005678', 'Bo K.', 'Gray', '8', '0', '0', '0', '0', '', '0', '0', '0', '', '', '2026-10-31 19:00']);
  const ch = w.book.insertSheet('Challenges');
  ch.rows.push(['Show', 'Title', 'Description', 'Event Types', 'Measure', 'Goal', 'Start', 'End', 'Bonus Points', 'Awarded'],
    [true, 'Dough Drive', 'Bills!', 'trait_dough', 'Points', 5, '2026-09-01', '2099-12-31', 50, ''],
    [true, 'Show up', '', '', 'Entries', 2, '2026-09-01', '2099-12-31', 0, ''],
    [false, 'Hidden', '', '', 'Points', 1, '', '', 10, '']);
}

test('My Squad: one student by 800#, positive counts, squad standing, their challenges', () => {
  const w = world();
  addStudentTabs(w);
  const r = JSON.parse(w.clashMyPage('800-1234'));
  assert.strictEqual(r.me.name, 'Ann L.');
  assert.strictEqual(r.me.bills, 12);
  assert.strictEqual(r.me.journal, null, 'Trail Journal hidden when blank');
  assert.deepStrictEqual(r.me.badges, ['🍞 First Trait Dough', '🔥 3-Week Streak']);
  // Events: Black 10, Gold 5 → Gold 2nd, 5 behind Black.
  assert.deepStrictEqual([r.standing.place, r.standing.ahead, r.standing.behindBy], [2, 'BLACK', 5]);
  assert.deepStrictEqual(r.challenges.map(c => c.title), ['Dough Drive', 'Show up']);
  assert.ok(r.allBadges.length >= 5);
  assert.strictEqual(JSON.parse(w.clashMyPage('8005678')).me.squad, 'GREY', 'Gray → GREY');
  assert.throws(() => w.clashMyPage('8009999'), /isn't on the list/);
  assert.throws(() => w.clashMyPage('12'), /800 number/);
  assert.ok(!w.clashMyPage('8001234').includes('8005678'), 'never sends another student');
});

test('My Squad look-ups are capped', () => {
  const w = world();
  addStudentTabs(w);
  w.ensureSettings_();
  w.setSettingValue_('MY_SQUAD_LOOKUPS_PER_10MIN', '3');
  for (let i = 0; i < 3; i++) w.clashMyPage('8001234');
  assert.throws(() => w.clashMyPage('8001234'), /busy/);
});

test('Challenges: progress from the events tab, bonuses once per squad when ON', () => {
  const w = world();
  addStudentTabs(w);
  const list = JSON.parse(w.clashChallenges());
  const dough = list.find(c => c.title === 'Dough Drive');
  assert.deepStrictEqual(dough.progress, { BLACK: 0, GOLD: 5, GREY: 0, PURPLE: 0 });
  assert.deepStrictEqual(list.find(c => c.title === 'Show up').progress, { BLACK: 1, GOLD: 1, GREY: 0, PURPLE: 0 });
  assert.ok(!list.some(c => c.title === 'Hidden'));
  w.ensureSettings_();
  assert.deepStrictEqual(J(w.checkChallengeBonuses()), ['CHALLENGE_BONUSES is OFF.']);
  w.setSettingValue_('CHALLENGE_BONUSES', 'ON');
  assert.deepStrictEqual(J(w.checkChallengeBonuses()), ['GOLD +50 for Dough Drive']);
  const bonus = events(w).filter(r => r[2] === 'challenge_bonus');
  assert.strictEqual(bonus.length, 1);
  assert.strictEqual(bonus[0][7], 50);
  assert.ok(/^GOLD \d{4}-\d{2}-\d{2}$/.test(w.book.getSheetByName('Challenges').rows[1][9]));
  assert.deepStrictEqual(J(w.checkChallengeBonuses()), [], 'never twice');
  assert.strictEqual(JSON.parse(w.clashChallenges()).find(c => c.title === 'Show up').progress.GOLD, 1, 'bonuses do not count toward challenges');
});

test('TV feed: latest points with labels; badge leaders only when TV_SHOW_NAMES is ON', () => {
  const w = world();
  addStudentTabs(w);
  w.ensureSettings_();
  let f = JSON.parse(w.clashTvFeed());
  assert.strictEqual(f.latest[0].name, 'Bill');
  assert.strictEqual(f.latest[0].label, '💵 Trait Dough');
  assert.deepStrictEqual(f.stars.map(s => s.name), ['Ann L.']);
  w.setSettingValue_('TV_SHOW_NAMES', 'OFF');
  f = JSON.parse(w.clashTvFeed());
  assert.deepStrictEqual(f.stars, []);
});

test('Generals HQ: General + code only, own squad only, everyone placed before doubling, checklist, service', () => {
  const w = world();
  addStudentTabs(w); // Ann L. (Gold), Bo K. (Grey) on the Student Board
  w.ensureSettings_();
  w.ensureGeneralsTabs_();
  const board = w.book.getSheetByName('Student Board');
  board.rows.push(['8007777', 'Gia P.', 'Gold', '8', '0', '0', '0', '0', '', '0', '0', '0', '', '', '']);
  w.book.getSheetByName('MLC Roster').rows.push(['8007777', 'Gia P.', 'Gold', 'General', ''], ['8001234', 'Ann L.', 'Gold', 'Hype Captain', ''],
    ['8005678', 'Bo K.', 'Grey', 'Vanguard', '']);
  w.book.getSheetByName('Sign-up Events').rows.push(['Dodgeball', '2026-11-20', 2, true], ['Lip Sync', '', '', false]);
  assert.throws(() => w.clashGeneralsSignIn('8007777', 'x'), /isn't open yet/);
  w.setSettingValue_('GENERALS_CODE', 'Summit');
  assert.throws(() => w.clashGeneralsSignIn('8001234', 'summit'), /don't match/, 'a Captain is not a General');
  const t = JSON.parse(w.clashGeneralsSignIn('800-7777', 'summit')).token;
  let hq = JSON.parse(w.clashGeneralsHq(t));
  assert.strictEqual(hq.squad, 'GOLD');
  assert.deepStrictEqual(hq.roster.map(s => s.name).sort(), ['Ann L.', 'Gia P.'], 'Gold only');
  assert.strictEqual(hq.notPlaced, 2);
  assert.strictEqual(hq.theme.theme, 'Hope');
  assert.deepStrictEqual(hq.leaders.map(l => l.role).sort(), ['General', 'Hype Captain']);
  hq = JSON.parse(w.clashGeneralsSignUp(t, 'Dodgeball', '8001234', false));
  assert.strictEqual(hq.notPlaced, 1);
  assert.deepStrictEqual(hq.roster[0].name, 'Gia P.', 'not-placed students are listed first');
  assert.throws(() => w.clashGeneralsSignUp(t, 'Dodgeball', '8005678', false), /isn't on your squad/);
  assert.throws(() => w.clashGeneralsSignUp(t, 'Lip Sync', '8007777', false), /closed/);
  w.clashGeneralsSignUp(t, 'Dodgeball', '8007777', false);
  w.book.getSheetByName('Student Board').rows.push(['8008888', 'Hal M.', 'Gold', '7', '0', '0', '0', '0', '', '0', '0', '0', '', '', '']);
  assert.throws(() => w.clashGeneralsSignUp(t, 'Dodgeball', '8008888', false), /full/);
  hq = JSON.parse(w.clashGeneralsSignUp(t, 'Dodgeball', '8001234', true));
  assert.strictEqual(hq.events[0].taken, 1, 'removed');
  const item = hq.mayhem[0];
  assert.strictEqual(hq.mayhem.length, 12);
  hq = JSON.parse(w.clashGeneralsChecklist(t, item.row, true));
  assert.strictEqual(hq.mayhem[0].done, true);
  const grey = JSON.parse(w.clashGeneralsHq(t)).mayhem; // still Gold's list
  assert.ok(grey.every(m => m.row !== undefined));
  w.book.getSheetByName('Squad Themes').rows[2][4] = 10; // Gold service goal
  hq = JSON.parse(w.clashGeneralsService(t, 4));
  assert.deepStrictEqual([hq.theme.goal, hq.theme.progress], [10, 4]);
  // My Squad shows public roles and the theme; Vanguard stays private.
  const ann = JSON.parse(w.clashMyPage('8001234'));
  assert.deepStrictEqual(ann.roles, ['Hype Captain']);
  assert.strictEqual(ann.theme.theme, 'Hope');
  assert.deepStrictEqual(JSON.parse(w.clashMyPage('8005678')).roles, [], 'Vanguard is not shown');
  // Nominations: own squad only, a reason required; letters saved once each and replaced on save.
  assert.throws(() => w.clashGeneralsNominate(t, 'Trailback Panel (Vanguard)', '8001234', 'short'), /why/);
  assert.throws(() => w.clashGeneralsNominate(t, 'Trailback Panel (Vanguard)', '8005678', 'Calm, fair and keeps things private.'), /your squad/);
  hq = JSON.parse(w.clashGeneralsNominate(t, 'Trailback Panel (Vanguard)', '8001234', 'Calm, fair and keeps things private.'));
  assert.deepStrictEqual(hq.nominations.map(n => n.nominee), ['Ann L.']);
  w.clashGeneralsLetter(t, 'To future Generals', 'Start early.');
  hq = JSON.parse(w.clashGeneralsLetter(t, 'To future Generals', 'Start early. Trust your Captains.'));
  assert.strictEqual(hq.letters[0].text, 'Start early. Trust your Captains.');
  assert.strictEqual(w.book.getSheetByName('General Letters').rows.length, 2, 'one row per letter, replaced on save');
  assert.throws(() => w.clashGeneralsHq('made-up'), /sign in/);
});

test('Daily Polls: one vote per 800#, results after voting, points for voters and winners once', () => {
  const w = world();
  addStudentTabs(w); // Ann L. 8001234 Gold, Bo K. 8005678 Grey
  w.ensureSettings_();
  w.ensurePollTabs_();
  const day = n => w.Utilities.formatDate(new Date(Date.now() + n * 86400000));
  const polls = w.book.getSheetByName('Daily Polls');
  polls.rows.splice(1); // drop the two samples
  polls.rows.push([true, day(0), 'This or That', 'Coke or Pepsi?', 'Coke | Pepsi', '', 1, 1, '', ''],
    [true, day(-1), 'Trivia', 'How many squads?', '3 | 4', '4', 1, 2, '', ''],
    [false, day(0), 'Trivia', 'Hidden?', 'a | b', 'a', 1, 1, '', '']);
  let d = JSON.parse(w.clashDailyPolls(''));
  assert.deepStrictEqual(d.today.map(p => p.question), ['Coke or Pepsi?']);
  assert.strictEqual(d.today[0].counts, null, 'no results before you vote');
  assert.strictEqual(d.last.question, 'How many squads?');
  assert.throws(() => w.clashPollVote('2', '8009999', 'Coke'), /isn't on the list/);
  assert.throws(() => w.clashPollVote('2', '8001234', 'Sprite'), /choices/);
  assert.throws(() => w.clashPollVote('3', '8001234', '4'), /closed/, 'yesterday is closed');
  d = JSON.parse(w.clashPollVote('2', '800-1234', 'Coke'));
  assert.strictEqual(d.today[0].voted, 'Coke');
  assert.deepStrictEqual(J(d.today[0].counts), { Coke: 1, Pepsi: 0 });
  assert.throws(() => w.clashPollVote('2', '8001234', 'Pepsi'), /already voted/);
  w.clashPollVote('2', '8005678', 'Coke');
  // Score yesterday's trivia: put two votes on it (as if cast yesterday).
  const votes = w.book.getSheetByName('Poll Votes');
  votes.rows.push(['3', day(-1), '8001234', 'GOLD', '4', ''], ['3', day(-1), '8005678', 'GREY', '3', '']);
  assert.deepStrictEqual(J(w.awardDailyPolls()), ['DAILY_POLL_POINTS is OFF.']);
  w.setSettingValue_('DAILY_POLL_POINTS', 'ON');
  assert.deepStrictEqual(J(w.awardDailyPolls()), ['How many squads?: BLACK +0, GOLD +3, GREY +1, PURPLE +0'], 'today\'s poll waits until tomorrow');
  const ev = events(w).filter(r => r[2] === 'daily_poll');
  assert.strictEqual(ev.length, 1);
  assert.deepStrictEqual([ev[0][7], ev[0][8]], [3, 1]);
  assert.ok(/^Awarded/.test(polls.rows[2][8]));
  assert.deepStrictEqual(J(w.awardDailyPolls()), [], 'never twice');
  const last = JSON.parse(w.clashDailyPolls('8001234')).last;
  assert.deepStrictEqual([last.answer, last.voted, J(last.winners)], ['4', '4', ['4']]);
  assert.ok(/Today's This or That: Coke or Pepsi\? Vote on/.test(w.pollAnnouncement_()), w.pollAnnouncement_());
  assert.ok(/The answer: 4/.test(w.pollAnnouncement_()));
});

test('point freeze: a frozen squad keeps its points but earns nothing new, from any project', () => {
  const w = world();
  w.ensureFreezeTab_();
  const fz = w.book.getSheetByName('Point Freezes');
  fz.rows.push(['GOLD', '2026-10-01', '2026-10-15', 'Unsportsmanlike conduct', 'Mr. S']);
  const ev = w.book.getSheetByName('events');
  // Another project adds a row during the freeze that also gives BLACK points.
  ev.rows.push([103, 'Poll', 'daily_poll', '2026-10-05', 'votes', '', 2, 4, 0, 0], [104, 'Later', 'trait_dough', '2026-10-20', '', '', 0, 7, 0, 0]);
  const totals = J(w.squadTotals_());
  assert.deepStrictEqual([totals.BLACK, totals.GOLD], [12, 7], 'GOLD: 5 + 4 frozen out, the 10-20 points count; BLACK unaffected');
  const row = ev.rows.find(r => r[0] === 103);
  assert.deepStrictEqual([row[6], row[7]], [2, 0]);
  assert.ok(/❄️ GOLD frozen: \+4 not counted/.test(row[4]) && /^votes/.test(row[4]), 'the lost-out points are noted on the row');
  assert.deepStrictEqual(J(w.squadTotals_()), totals, 'running again changes nothing');
  assert.strictEqual(w.freezeToday_('GOLD'), null, 'not frozen today (freeze ended)');
  fz.rows.push(['GREY', '2026-01-01', '', 'Open-ended', '']);
  assert.deepStrictEqual(J(w.standingFor_('GREY', totals).frozen), { until: '' });
  // The leaderboard and TV get today's freezes; an admin save never stores them, and an event
  // the admin adds for GREY during the freeze is saved as 0 for GREY.
  const loaded = JSON.parse(w.clashLoad()).value;
  assert.deepStrictEqual(loaded.freezes, [{ squad: 'GREY', until: '' }]);
  const t = login(w, 'oldpw');
  loaded.events.push({ id: 200, name: 'Admin add', type: 'event', date: '2026-10-25', note: '', pts: { BLACK: 1, GOLD: 0, GREY: 3, PURPLE: 0 } });
  w.clashSave(t, JSON.stringify(loaded), JSON.stringify({ seen: loaded.events.map(e => e.id + '|' + e.name) }));
  const saved = ev.rows.find(r => r[0] === 200);
  assert.deepStrictEqual([saved[6], saved[8]], [1, 0]);
  assert.ok(/❄️ GREY frozen: \+3/.test(saved[4]));
  assert.ok(!w.book.getSheetByName('app_data').rows.some(r => String(r[2]).includes('freezes')), 'freezes are not saved in the blob');
});

test('lore: Squad Themes has room for each squad\'s story; nothing shows until Show Lore is ticked', () => {
  const w = world();
  w.ensureGeneralsTabs_();
  const themes = w.book.getSheetByName('Squad Themes');
  assert.deepStrictEqual(themes.rows[0].slice(6), ['Trait', 'Lore Title', 'Lore', 'Lore Image', 'Show Lore']);
  const gold = themes.rows.find(r => r[0] === 'GOLD');
  gold[6] = 'Hope'; gold[7] = 'The Lantern Keepers'; gold[8] = 'Long ago…'; gold[9] = 'javascript:alert(1)';
  assert.strictEqual(w.squadTheme_('GOLD').lore, null, 'a draft stays hidden');
  assert.strictEqual(w.squadTheme_('GOLD').trait, '');
  gold[10] = true;
  assert.deepStrictEqual(J(w.squadTheme_('GOLD').lore), { title: 'The Lantern Keepers', text: 'Long ago…', image: '' }, 'only https images');
  assert.strictEqual(w.squadTheme_('GOLD').trait, 'Hope');
  // An older sheet without the lore columns gets them added at the end.
  const w2 = world();
  w2.book.insertSheet('Squad Themes').rows.push(['Squad', 'Theme', 'What It Means', 'Service Project', 'Service Goal', 'Service Progress'], ['GOLD', 'Hope', '', '', '', '']);
  w2.ensureGeneralsTabs_();
  assert.deepStrictEqual(w2.book.getSheetByName('Squad Themes').rows[0].slice(6), ['Trait', 'Lore Title', 'Lore', 'Lore Image', 'Show Lore']);
});

test('Squad Duels: scores by measure, a tie means both win, bonus once the day after it ends, only when ON', () => {
  const w = world();
  w.ensureSettings_();
  w.ensureDuelTab_();
  w.ensurePollTabs_();
  const day = n => w.Utilities.formatDate(new Date(Date.now() + n * 86400000));
  const duels = w.book.getSheetByName('Squad Duels');
  duels.rows.push([true, 'Dough Duel', day(-7), day(-1), 'Entries', 'trait_dough', 'GOLD vs GREY', 25, '', ''],
    [true, 'Vote-off', day(-7), day(-1), 'Poll Votes', '', '', 10, '', ''],
    [true, 'Running', day(-1), day(5), 'Points', '', '', 5, '', '']);
  const ev = w.book.getSheetByName('events');
  ev.rows.push([110, 'b1', 'trait_dough', day(-3), '', '', 0, 1, 0, 0], [111, 'b2', 'trait_dough', day(-2), '', '', 0, 0, 3, 0],
    [112, 'old bonus', 'challenge_bonus', day(-2), '', '', 0, 0, 50, 0]);
  w.book.getSheetByName('Poll Votes').rows.push(['2', day(-3), '8001', 'PURPLE', 'x', ''], ['2', day(-3), '8002', 'PURPLE', 'y', ''], ['2', day(-20), '8003', 'GOLD', 'x', '']);
  const list = JSON.parse(w.clashDuels());
  const dough = list.find(d => d.title === 'Dough Duel');
  assert.deepStrictEqual([J(dough.squads), J(dough.scores), J(dough.leaders), dough.done], [['GOLD', 'GREY'], { GOLD: 1, GREY: 1 }, ['GOLD', 'GREY'], true], 'entries, not points; bonuses ignored; a tie');
  assert.deepStrictEqual(J(list.find(d => d.title === 'Vote-off').leaders), ['PURPLE']);
  assert.ok(list.find(d => d.title === 'Running').active);
  assert.deepStrictEqual(J(w.awardDuels()), ['DUEL_BONUSES is OFF.']);
  w.setSettingValue_('DUEL_BONUSES', 'ON');
  assert.deepStrictEqual(J(w.awardDuels()), ['Dough Duel: GOLD & GREY +25', 'Vote-off: PURPLE +10'], 'the running duel waits');
  const bonus = ev.rows.filter(r => r[2] === 'duel_bonus');
  assert.deepStrictEqual(bonus.map(r => [r[7], r[8], r[9]]), [[25, 25, ''], ['', '', 10]]);
  assert.deepStrictEqual(J(w.awardDuels()), [], 'never twice');
  assert.ok(/^Awarded .*GOLD & GREY \+25/.test(duels.rows.find(r => r[1] === 'Dough Duel')[8]));
  assert.ok(JSON.parse(w.clashTvFeed()).latest.some(e => e.label === '⚔️ Squad Duel'), 'labelled on the TV even with the old EVENT_LABELS');
});

test('Lore chapters: unlock by points / challenge / date per squad, recorded once; drafts stay hidden', () => {
  const w = world();
  addStudentTabs(w); // challenge "Dough Drive": goal 5 trait_dough points (GOLD has 5 from the base events)
  w.ensureLoreTab_();
  const lore = w.book.getSheetByName('Lore Chapters');
  lore.rows.splice(1); // drop the two draft examples
  lore.rows.push([true, 'ALL', 1, 'The Climb Begins', 'Long ago…', 'https://img/1.png', '', ''],
    [true, 'ALL', 2, 'The Ridge', 'Onward.', '', 'points: 8', ''],
    [true, 'GOLD', 3, 'Gold Light', 'Shine.', '', 'challenge: Dough Drive', ''],
    [true, 'ALL', 4, 'Future', 'Later.', '', 'date: 2099-01-01', ''],
    [false, 'ALL', 5, 'Draft', 'secret', '', '', '']);
  const gold = J(w.loreFor_('GOLD')), grey = J(w.loreFor_('GREY'));
  assert.deepStrictEqual(gold.chapters.map(c => c.title), ['The Climb Begins', 'Gold Light'], 'GOLD: 5 points < 8, challenge done');
  assert.deepStrictEqual(gold.next, { chapter: 2, hint: 'Unlocks at 8 squad points' });
  assert.deepStrictEqual(grey.chapters.map(c => c.title), ['The Climb Begins']);
  assert.ok(!JSON.stringify(gold).includes('secret'), 'unticked drafts never reach the page');
  assert.ok(/GOLD \d{4}-\d{2}-\d{2}/.test(lore.rows[3][7]), 'unlock date recorded');
  w.book.getSheetByName('events').rows.push([120, 'x', 'event', '2026-10-02', '', '', 0, 3, 0, 0]);
  assert.deepStrictEqual(J(w.loreFor_('GOLD')).chapters.map(c => c.chapter), [1, 2, 3], 'reaching 8 points unlocks chapter 2');
  assert.deepStrictEqual(J(w.recentLoreUnlocks_()).map(u => u.squad + u.chapter).sort(), ['BLACK2', 'GOLD2', 'GOLD3'], 'TV: unlocks with a rule (BLACK has 10 points), not chapter 1');
  const me = JSON.parse(w.clashMyPage('8001234'));
  assert.deepStrictEqual(me.lore.chapters.map(c => c.chapter), [1, 2, 3]);
});

test('Trail Passport: stamps from the Student Board, poll votes and sign-ups; complete when all are done', () => {
  const w = world();
  addStudentTabs(w);
  w.ensureSettings_();
  w.ensurePollTabs_();
  w.ensureGeneralsTabs_();
  w.setSettingValue_('PASSPORT_STAMPS', 'bills=🍞 Trait Dough; star=📈 STAR; polls=🗳️ Polls; events=🎟️ Event');
  w.setSettingValue_('PASSPORT_POLL_VOTES', '2');
  let p = JSON.parse(w.clashMyPage('8001234')).passport;
  assert.deepStrictEqual(p.stamps.map(s => s.done), [true, true, false, false]);
  assert.strictEqual(p.complete, false);
  w.book.getSheetByName('Poll Votes').rows.push(['1', '2026-10-01', '8001234', 'GOLD', 'a', ''], ['2', '2026-10-02', '8001234', 'GOLD', 'b', '']);
  w.book.getSheetByName('Sign-ups').rows.push(['Dodgeball', '8001234', 'Ann L.', 'GOLD', 'G', '']);
  p = JSON.parse(w.clashMyPage('8001234')).passport;
  assert.strictEqual(p.complete, true);
  assert.deepStrictEqual(JSON.parse(w.clashMyPage('8005678')).passport.stamps.map(s => s.done), [false, false, false, false]);
});

test('Event pages: Mayhem/Kick Off fields and rows from the Sheet; blanks fall back; only https links; Mayhem Week scoreboard', () => {
  const w = world();
  w.ensureEventPageTabs_();
  let d = JSON.parse(w.clashEventPages());
  assert.strictEqual(d.mayhem.fields['Date'], 'Friday, March 19, 2027');
  assert.ok(!('Location' in d.mayhem.fields), 'a blank Value keeps the Edit Page text');
  assert.ok(/15NP5-GEWObOM2HTOm46fznCmqlS/.test(d.mayhem.fields['Planning Doc']));
  assert.deepStrictEqual([d.mayhem.week.length, d.mayhem.schedule.length, d.mayhem.events.length], [5, 15, 8]);
  assert.strictEqual(d.kickoff.fields['Date'], 'Friday, September 25, 2026');
  assert.strictEqual(d.mayhem.scoreboard, null, 'no Mayhem Week points yet');
  const rows = w.book.getSheetByName('Event Rows');
  rows.rows.push([true, 'kickoff', 'Links', 1, 'Bad', '', '', 'javascript:alert(1)'], [false, 'kickoff', 'Events', 1, 'Hidden', '', '', ''],
    [true, 'kickoff', 'Events', 2, 'Squad Relay', 'Rules', '10 pts', '']);
  w.book.getSheetByName('Event Pages').rows.push(['kickoff', 'Planning Doc', 'http://not-secure.example', '']);
  const ev = w.book.getSheetByName('events');
  const now = new Date(), sy = now.getMonth() >= 7 ? now.getFullYear() : now.getFullYear() - 1; // this school year starts Aug 1 of sy
  ev.rows.push([130, 'Paper Chain Challenge', 'mayhem_week', sy + '-08-15', '', '', 3, 5, 0, 1], [131, 'Trivia', 'mayhem_week', sy + '-08-16', '', '', 0, 0, 4, 0],
    [132, 'Old', 'mayhem_week', sy + '-03-15', '', '', 9, 9, 9, 9]);
  d = JSON.parse(w.clashEventPages());
  assert.deepStrictEqual(d.kickoff.links, [], 'non-https links are dropped');
  assert.strictEqual(d.kickoff.fields['Planning Doc'], '');
  assert.deepStrictEqual(d.kickoff.events.map(e => e.heading), ['Squad Relay']);
  assert.deepStrictEqual(J(d.mayhem.scoreboard.columns), ['Paper Chain Challenge', 'Trivia'], 'only this school year');
  assert.deepStrictEqual(J(d.mayhem.scoreboard.scores['Paper Chain Challenge']), { BLACK: 3, GOLD: 5, GREY: 0, PURPLE: 1 });
  assert.ok(JSON.parse(w.clashTvFeed()).latest.some(e => e.label === '🏔️ Mayhem Week'));
});

test('Mayhem Week: 3-event limit, no two at once, self sign-up only when ON, captain votes, shirt points, Squad Night door', () => {
  const w = world();
  addStudentTabs(w); // Ann L. 8001234 Gold, Bo K. 8005678 Grey
  w.ensureSettings_();
  w.ensureGeneralsTabs_();
  w.ensureMayhemWeekTabs_();
  const evTab = w.book.getSheetByName('Sign-up Events');
  assert.deepStrictEqual(evTab.rows[0].slice(4), ['Slot', 'Type', 'Self Sign-up']);
  evTab.rows.push(['Mario Kart', '2027-03-15', 4, true, 'Lunch', 'Gaming', true], ['Squad Bowl', '2027-03-15', 4, true, 'Lunch', 'Brain', true],
    ['Chess', '2027-03-16', 2, true, 'Core Extension', 'Brain', true], ['Spoons', '2027-03-17', 4, true, 'Lunch', 'Table', true],
    ['Tug-o-war', '2027-03-19', 15, true, 'Squad Night', 'Active', false], ['Old event', '2026-10-01', 0, false, '', '', true]);
  assert.throws(() => w.clashStudentSignUp('8001234', 'Mario Kart'), /Generals/, 'OFF by default');
  w.setSettingValue_('SELF_SIGNUPS', 'ON');
  let su = JSON.parse(w.clashStudentSignUp('8001234', 'Mario Kart'));
  assert.strictEqual(su.count, 1);
  assert.throws(() => w.clashStudentSignUp('8001234', 'Squad Bowl'), /same time/);
  JSON.parse(w.clashStudentSignUp('8001234', 'Chess'));
  su = JSON.parse(w.clashStudentSignUp('8001234', 'Spoons'));
  assert.strictEqual(su.count, 3);
  w.setSettingValue_('SELF_SIGNUPS', 'ON');
  assert.throws(() => w.clashStudentSignUp('8001234', 'Tug-o-war'), /isn't open for sign-ups/, 'not a self sign-up event');
  w.book.getSheetByName('Sign-ups').rows.push(['Old event', '8001234', 'Ann L.', 'GOLD', 'G', '']); // closed events don't count
  const ms = JSON.parse(w.clashMyPage('8001234')).signups;
  assert.deepStrictEqual([ms.count, ms.limit, ms.events.filter(e => e.mine).map(e => e.event)], [3, 3, ['Mario Kart', 'Chess', 'Spoons']]);
  // Generals hit the same limit.
  w.book.getSheetByName('MLC Roster').rows.push(['8007777', 'Gia P.', 'Gold', 'General', ''], ['8008888', 'Hal Q.', 'Gold', 'General', '']);
  w.book.getSheetByName('Student Board').rows.push(['8007777', 'Gia P.', 'Gold', '8', '0', '0', '0', '0', '', '0', '0', '0', '', '', '']);
  w.setSettingValue_('GENERALS_CODE', 'summit');
  const t = JSON.parse(w.clashGeneralsSignIn('8007777', 'summit')).token;
  assert.throws(() => w.clashGeneralsSignUp(t, 'Tug-o-war', '8001234', false), /3 events/);
  JSON.parse(w.clashStudentSignUp('8001234', 'Spoons', true)); // drop one
  JSON.parse(w.clashGeneralsSignUp(t, 'Tug-o-war', '8001234', false));
  // Captain votes: one per General per student, tallied for the squad.
  JSON.parse(w.clashGeneralsNominate(t, 'Squad Captain (7th grade)', '8001234', 'Always helps out at lunch.'));
  assert.throws(() => w.clashGeneralsNominate(t, 'Squad Captain (7th grade)', '8001234', 'Again, really helpful.'), /already nominated/);
  const t2 = JSON.parse(w.clashGeneralsSignIn('8008888', 'summit')).token;
  const hq = JSON.parse(w.clashGeneralsNominate(t2, 'Squad Captain (7th grade)', '8001234', 'Organized and kind.'));
  assert.deepStrictEqual(hq.captainTally.map(x => [x.nominee, x.votes]), [['Ann L.', 2]]);
  // Shirts: paid ones send points once, squad filled from the Student Board.
  const shirts = w.book.getSheetByName('Shirt Orders');
  shirts.rows.push(['2027-02-01', '8001234', 'Ann L.', '', 'M', true, '', ''], ['2027-02-01', '8005678', 'Bo K.', 'GREY', 'L', false, '', '']);
  assert.deepStrictEqual(J(w.sendShirtPoints()), ['1 shirts: BLACK +0, GOLD +2, GREY +0, PURPLE +0']);
  assert.deepStrictEqual(J(w.sendShirtPoints()), ['No new paid shirts.']);
  // Door: competitor free, unpaid stranger pays, nobody twice; attendance points once.
  const admin = login(w, 'oldpw');
  assert.throws(() => w.clashDoorCheck('', '8001234'), /sign in/);
  assert.strictEqual(JSON.parse(w.clashDoorCheck(admin, '8001234')).student.entry, 'Free (competitor)');
  assert.strictEqual(JSON.parse(w.clashDoorCheck(admin, '8005678')).student.entry, 'Pay $5');
  assert.throws(() => w.clashDoorAdmit(admin, '8005678', false), /paid first/);
  w.clashDoorAdmit(admin, '8005678', true);
  const r = JSON.parse(w.clashDoorAdmit(admin, '8001234', false));
  assert.deepStrictEqual(J(r.counts), { BLACK: 0, GOLD: 1, GREY: 1, PURPLE: 0 });
  assert.throws(() => w.clashDoorAdmit(admin, '8001234', false), /already came in/);
  assert.deepStrictEqual(J(w.sendSquadNightPoints()), ['BLACK 0 → +0, GOLD 1 → +1, GREY 1 → +1, PURPLE 0 → +0']);
  assert.ok(/already sent/.test(w.sendSquadNightPoints()[0]));
});

test('page files: scripts parse, no old token or Supabase key left', () => {
  for (const f of ['App.html', 'Widgets.html', 'Index.html', 'Student.html', 'Generals.html']) {
    const h = fs.readFileSync(path.join(dir, f), 'utf8');
    assert.ok(!/togetherweclimb4444HOPE|supabase\.co|TJ_ANON|ephrata2026/.test(h), f + ' has no secrets');
    [...h.matchAll(/<script[^>]*>([\s\S]*?)<\/script>/g)].forEach(m => new vm.Script(m[1].replace(/<\?!=[\s\S]*?\?>/g, '{}')));
  }
});

console.log('\n' + passed + ' tests passed');
