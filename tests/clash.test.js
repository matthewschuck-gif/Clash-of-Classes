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
    setFrozenRows() {}, setColumnWidth() {}, clear() { this.rows = []; },
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
    CacheService: { getScriptCache: () => ({ get: k => cache[k] || null, put: (k, v) => { cache[k] = v; }, remove: k => { delete cache[k]; } }) },
    PropertiesService: { getScriptProperties: () => ({ getProperty: k => props[k] || null, setProperty: (k, v) => { props[k] = v; } }) },
    LockService: { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) },
    Utilities: {
      getUuid: (() => { let n = 0; return () => 'uuid-' + (++n); })(),
      computeDigest: (alg, s) => Array.from(require('crypto').createHash('sha256').update(s).digest()).map(b => b > 127 ? b - 256 : b),
      DigestAlgorithm: { SHA_256: 'sha256' }, Charset: { UTF_8: 'utf8' },
      formatDate: (d, tz, f) => { // UTC stands in for the school's time zone in tests
        const p = n => String(n).padStart(2, '0'), ymd = d.toISOString().slice(0, 10), hm = p(d.getUTCHours()) + ':' + p(d.getUTCMinutes());
        return f === 'HH:mm' ? hm : f === 'u' ? String(((d.getUTCDay() + 6) % 7) + 1) : /HH/.test(f || '') ? ymd + ' ' + hm : ymd;
      },
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
const SQUADS = ['BLACK', 'GOLD', 'GREY', 'PURPLE'];
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

test('Pulse: open several days, live results right after a vote, no winner, a point per voter, by-grade results, live screen without 800#s, ideas bank', () => {
  const w = world();
  addStudentTabs(w); // Ann L. 8001234 Gold 7th, Bo K. 8005678 Grey 8th
  w.ensureSettings_(); w.ensurePollTabs_();
  const day = n => w.Utilities.formatDate(new Date(Date.now() + n * 86400000));
  const polls = w.book.getSheetByName('Daily Polls');
  assert.ok(polls.rows[0].includes('Through') && w.book.getSheetByName('Poll Votes').rows[0].includes('Grade'));
  assert.ok(w.book.getSheetByName('Poll Ideas').rows.length > 30, 'the bank of school questions is seeded');
  polls.rows.splice(1);
  polls.rows.push([true, day(-1), 'Pulse', 'Best time for extra help?', 'Before school | Lunch | After school', '', 1, 5, '', '', day(1)]);
  let d = JSON.parse(w.clashDailyPolls(''));
  assert.strictEqual(d.today.length, 1, 'opened yesterday, open through tomorrow');
  assert.strictEqual(d.today[0].kind, 'Pulse');
  d = JSON.parse(w.clashPollVote('2', '8001234', 'Lunch'));
  assert.deepStrictEqual(J(d.today[0].counts), { 'Before school': 0, Lunch: 1, 'After school': 0 }, 'results the moment you vote');
  w.clashPollVote('2', '8005678', 'Lunch');
  const vote = w.book.getSheetByName('Poll Votes').rows[1];
  assert.strictEqual(vote[1], w.todayYmd_(), 'the vote is dated the day it was cast');
  assert.strictEqual(vote[6], '7', 'grade comes from the Student Board');
  // The live screen: counts and squad turnout, never an 800#.
  const live = JSON.parse(w.clashPulseLive());
  assert.strictEqual(live.live[0].total, 2);
  assert.deepStrictEqual(J(live.live[0].turnout.GOLD), { n: 1, pct: 100 });
  assert.ok(!/8001234|Ann/.test(w.clashPulseLive()));
  assert.ok(/This week's Pulse question: Best time for extra help\?.*through/.test(w.pollAnnouncement_()), w.pollAnnouncement_());
  // Not scored while open; then a point per voter and no winner points (Win Points is ignored).
  w.setSettingValue_('DAILY_POLL_POINTS', 'ON');
  assert.deepStrictEqual(J(w.awardDailyPolls()), []);
  polls.rows[1][10] = day(-1);
  assert.throws(() => w.clashPollVote('2', '8001234', 'Lunch'), /closed/);
  assert.deepStrictEqual(J(w.awardDailyPolls()), ['Best time for extra help?: BLACK +0, GOLD +1, GREY +1, PURPLE +0']);
  assert.ok(/by grade: 7th \(Before school 0, Lunch 1, After school 0\); 8th/.test(polls.rows[1][9]), polls.rows[1][9]);
  const ev = events(w).find(r => r[2] === 'daily_poll');
  assert.ok(/^Pulse: /.test(ev[1]) && !/winner/.test(ev[4]));
  assert.ok(/Our last Pulse question/.test(w.pollAnnouncement_()));
  assert.strictEqual(JSON.parse(w.clashPulseLive()).recent[0].closed, true);
  // The next idea goes on Daily Polls as a Pulse, open PULSE_DAYS school days, and is marked used.
  const r = J(w.addNextPulse_());
  assert.ok(r && r.start > w.todayYmd_() && r.through >= r.start);
  const added = polls.rows[polls.rows.length - 1];
  assert.deepStrictEqual([added[2], added[3], added[10]], ['Pulse', r.question, r.through]);
  const ideas = w.book.getSheetByName('Poll Ideas').rows, qCol = ideas[0].indexOf('Question'), usedCol = ideas[0].indexOf('Used On');
  assert.strictEqual(ideas.find(x => x[qCol] === r.question)[usedCol], r.start, 'the idea is marked used (Used On)');
  assert.notStrictEqual(J(w.addNextPulse_()).question, r.question, 'the next call takes the next idea');
  // An older Daily Polls tab gets Through, and old one-day polls still work.
  const w2 = world(); w2.ensureSettings_();
  w2.book.insertSheet('Daily Polls').rows = [['Show', 'Date', 'Kind', 'Question', 'Choices', 'Answer', 'Participation Points', 'Win Points', 'Status', 'Results'], [true, day(0), 'This or That', 'A or B?', 'A | B', '', 1, 1, '', '']];
  w2.ensurePollTabs_();
  assert.strictEqual(w2.book.getSheetByName('Daily Polls').rows[0][10], 'Through');
  assert.strictEqual(JSON.parse(w2.clashDailyPolls('')).today.length, 1);
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
  const gf = J(w.standingFor_('GREY', totals).frozen);
  assert.deepStrictEqual([gf.until, gf.kind, gf.ready], ['', 'Peaks', false], 'blank Kind = Peaks; no sign-off yet, so not ready');
  // The leaderboard and TV get today's freezes; an admin save never stores them, and an event
  // the admin adds for GREY during the freeze is saved as 0 for GREY.
  const loaded = JSON.parse(w.clashLoad()).value;
  assert.deepStrictEqual(loaded.freezes, [{ squad: 'GREY', until: '', kind: 'Peaks' }]);
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
  assert.deepStrictEqual(themes.rows[0].slice(6), ['Trait', 'Lore Title', 'Lore', 'Lore Image', 'Show Lore', 'Mascot', 'Mascot Name', 'Mascot Image']);
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
  assert.deepStrictEqual(w2.book.getSheetByName('Squad Themes').rows[0].slice(6), ['Trait', 'Lore Title', 'Lore', 'Lore Image', 'Show Lore', 'Mascot', 'Mascot Name', 'Mascot Image']);
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
  assert.deepStrictEqual([d.mayhem.week.length, d.mayhem.schedule.length, d.mayhem.events.length, d.mayhem.night.length], [5, 15, 8, 8]);
  assert.ok(/Friday, March 19/.test(d.mayhem.fields['Squad Night']) && /\$5/.test(d.mayhem.fields['Squad Night Entry']));
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
  assert.strictEqual(evTab.rows.length - 1, 25, 'an empty tab starts with the Mayhem Week chart');
  assert.ok(evTab.rows.slice(1).every(r => r[3] === false), 'all closed until sign-ups open');
  assert.strictEqual(evTab.rows.filter(r => r[4] === 'Field Events').length, 5, 'Thursday field events share one slot');
  evTab.rows.splice(1);
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

test('Shirt designs: Generals upload to a private folder, staff approve, one changeable vote per student, own squad only', () => {
  const w = world();
  addStudentTabs(w); // Ann L. 8001234 Gold, Bo K. 8005678 Grey
  w.ensureSettings_();
  w.ensureGeneralsTabs_();
  w.ensureShirtDesignTabs_();
  const files = {};
  let folders = 0;
  w.DriveApp = {
    createFolder: () => ({ getId: () => 'FOLDER' + (++folders) + 'xxxxxxxxxxxxxxxxxxxxxxxxx', createFile: blob => { const id = 'F' + Object.keys(files).length + 'xxxxxxxxxxxxxxxxxxxxxxxxx'; files[id] = blob; return { getId: () => id }; } }),
    getFolderById: id => ({ createFile: blob => { const fid = 'F' + Object.keys(files).length + 'xxxxxxxxxxxxxxxxxxxxxxxxx'; files[fid] = blob; return { getId: () => fid }; } }),
    getFileById: id => ({ getBlob: () => files[id], getLastUpdated: () => new Date() }),
  };
  w.Utilities.base64Decode = s => Array.from(Buffer.from(s, 'base64'));
  w.Utilities.base64Encode = b => Buffer.from(b).toString('base64');
  w.Utilities.newBlob = (bytes, type, name) => ({ getBytes: () => bytes, getContentType: () => type, name });
  w.book.getSheetByName('MLC Roster').rows.push(['8007777', 'Gia P.', 'Gold', 'General', '']);
  w.book.getSheetByName('Student Board').rows.push(['8007777', 'Gia P.', 'Gold', '8', '0', '0', '0', '0', '', '0', '0', '0', '', '', '']);
  w.setSettingValue_('GENERALS_CODE', 'summit');
  const t = JSON.parse(w.clashGeneralsSignIn('8007777', 'summit')).token;
  const png = 'data:image/png;base64,' + Buffer.from('fake-png').toString('base64');
  assert.throws(() => w.clashGeneralsUploadDesign(t, 'X', 'data:text/html;base64,PGI+'), /PNG or JPG/);
  let hq = JSON.parse(w.clashGeneralsUploadDesign(t, 'Summit Lightning', png));
  hq = JSON.parse(w.clashGeneralsUploadDesign(t, 'Gold Rush', png));
  assert.deepStrictEqual(hq.designs.map(d => [d.title, d.approved]), [['Summit Lightning', false], ['Gold Rush', false]]);
  assert.strictEqual(folders, 1, 'one private folder, made on the first upload');
  const designs = w.book.getSheetByName('Shirt Designs');
  const id1 = designs.rows[1][0], id2 = designs.rows[2][0];
  assert.strictEqual(JSON.parse(w.clashShirtDesigns('8001234')).designs.length, 0, 'nothing shows before approval');
  assert.throws(() => w.clashShirtImage(id1, ''), /isn't approved/);
  assert.ok(/^data:image\/png;base64,/.test(w.clashShirtImage(id1, t)), 'the squad\'s General can see a pending design');
  designs.rows[1][6] = true; designs.rows[2][6] = true;
  assert.throws(() => w.clashShirtVote('8001234', id1), /closed/);
  w.setSettingValue_('SHIRT_VOTING', 'ON');
  let s = JSON.parse(w.clashShirtVote('8001234', id1));
  assert.deepStrictEqual([s.voted, s.designs.map(d => d.votes)], [id1, [1, 0]]);
  s = JSON.parse(w.clashShirtVote('8001234', id2));
  assert.deepStrictEqual([s.voted, s.designs.map(d => d.votes)], [id2, [0, 1]], 'a changed vote moves, not doubles');
  assert.throws(() => w.clashShirtVote('8005678', id1), /own squad/);
  assert.deepStrictEqual(JSON.parse(w.clashShirtDesigns('8005678')).designs, [], 'other squads don\'t see Gold\'s designs');
  assert.strictEqual(designs.rows[2][7], 1, 'the Votes column stays current for staff');
});

test('Privacy: the public site and TV never show a student name on behavior rows; positive rows show "First L."; admin saves keep the real names', () => {
  const w = world();
  w.ensureSettings_();
  const ev = w.book.getSheetByName('events');
  ev.rows.push([140, 'Trail Journal Growth — Jamie Doe', 'trail_journal', '2026-10-20', 'High insight reflection, staff-approved', '', 0, 1, 0, 0],
    [141, 'Trait Dough turn-in: Alex Sam Roe', 'trait_dough', '2026-10-21', '2 bill(s), team M', '', 0, 0, 2, 0],
    [142, 'Mountaineer Mindset Summit: Pat Lee', 'mountaineer_mindset', '2026-10-22', 'Oct Grit (Applying), HR 12, approved by staff', '', 1, 0, 0, 0]);
  const pub = JSON.parse(w.clashLoad()).value.events;
  const byId = id => pub.find(e => e.id === id);
  assert.deepStrictEqual([byId(140).name, byId(140).note], ['Trail Journal', ''], 'behavior row: label only');
  assert.strictEqual(byId(141).name, 'Trait Dough turn-in: Alex R.');
  assert.strictEqual(byId(142).name, 'Mountaineer Mindset Summit: Pat L.');
  assert.ok(!JSON.stringify(pub).includes('Jamie') && !JSON.stringify(pub).includes('Doe'), 'no trace of the name anywhere');
  w.setSettingValue_('TV_SHOW_NAMES', 'OFF');
  assert.strictEqual(JSON.parse(w.clashLoad()).value.events.find(e => e.id === 141).name, 'Trait Dough turn-in');
  assert.ok(!JSON.stringify(JSON.parse(w.clashTvFeed())).includes('Jamie'), 'TV feed too');
  // Admin sees the full rows.
  const t = login(w, 'oldpw');
  assert.strictEqual(JSON.parse(w.clashLoad(t)).value.events.find(e => e.id === 140).name, 'Trail Journal Growth — Jamie Doe');
  // A page that loaded publicly, then signed in and saved: real names kept, nothing duplicated.
  const D = JSON.parse(w.clashLoad()).value;
  const sync = { seen: D.events.map(e => e.id + '|' + e.name) };
  D.events.push({ id: 150, name: 'New admin event', type: 'event', date: '2026-10-25', note: '', pts: { BLACK: 5, GOLD: 0, GREY: 0, PURPLE: 0 } });
  w.clashSave(t, JSON.stringify(D), JSON.stringify(sync));
  const rows = ev.rows.slice(1);
  assert.strictEqual(rows.filter(r => r[0] === 140).length, 1, 'no duplicate row');
  assert.strictEqual(rows.find(r => r[0] === 140)[1], 'Trail Journal Growth — Jamie Doe');
  assert.strictEqual(rows.find(r => r[0] === 140)[4], 'High insight reflection, staff-approved', 'note kept');
  assert.strictEqual(rows.find(r => r[0] === 141)[1], 'Trait Dough turn-in: Alex Sam Roe');
  assert.ok(rows.some(r => r[1] === 'New admin event'));
});

test('school calendar: days off on an editable tab, synced from a district iCal feed, school-day counts', () => {
  const w = world();
  w.ensureSchoolCalendarTab_();
  const tab = w.book.getSheetByName('School Calendar');
  assert.ok(tab.rows.length > 25 && tab.rows.some(r => r[0] === '2026-11-26'), 'seeded with this year\'s days off');
  tab.rows.push(['2027-02-02', 'Snow day', '']);
  let cal = J(w.schoolCalendar_());
  assert.ok(cal.noSchool.includes('2027-02-02') && cal.noSchool.includes('2026-11-26'));
  // District feed: a snow day and a half day; only the day off is added, once.
  const ics = 'BEGIN:VCALENDAR\r\nBEGIN:VEVENT\r\nDTSTART;VALUE=DATE:20270203\r\nDTEND;VALUE=DATE:20270204\r\nSUMMARY:Emergency Closing - Snow\r\nEND:VEVENT\r\n' +
    'BEGIN:VEVENT\r\nDTSTART;VALUE=DATE:20270205\r\nSUMMARY:Half Day - Early Dismissal\r\nEND:VEVENT\r\nBEGIN:VEVENT\r\nDTSTART;VALUE=DATE:20261126\r\nSUMMARY:Thanksgiving Holiday\r\nEND:VEVENT\r\nEND:VCALENDAR';
  w.UrlFetchApp = { fetch: () => ({ getContentText: () => ics }) };
  w.setSettingValue_('SCHOOL_CALENDAR', 'https://district.example/cal.ics');
  w.setSettingValue_('SCHOOL_START', '2026-08-31'); w.setSettingValue_('LAST_DAY', '2027-06-09');
  const added = J(w.syncSchoolCalendar_());
  assert.deepStrictEqual(added.map(a => a[0]), ['2027-02-03'], 'snow day added; half day skipped; Thanksgiving already there');
  assert.deepStrictEqual(J(w.syncSchoolCalendar_()), [], 'running again adds nothing');
  cal = J(w.schoolCalendar_());
  assert.deepStrictEqual([cal.start, cal.end], ['2026-08-31', '2027-06-09']);
  assert.ok(JSON.parse(w.clashLoad()).value.schoolCalendar.noSchool.includes('2027-02-03'), 'the page gets the list');
  // The page's counts: Mon Nov 23 - Fri Nov 27 is break, so from Sat Nov 21 the next school day is Tue Dec 1.
  const app = fs.readFileSync(path.join(dir, 'App.html'), 'utf8');
  const pick = name => app.slice(app.indexOf('function ' + name), app.indexOf('\n}', app.indexOf('function ' + name)) + 2);
  const page = { D: { schoolCalendar: cal }, SCHOOL_HOLIDAYS_2627: new Set(), Date, String, Set, Array };
  vm.createContext(page);
  vm.runInContext(['noSchoolDays', 'localYmd', 'countSchoolDays'].map(pick).join('\n'), page);
  assert.strictEqual(page.countSchoolDays('2026-11-22', '2026-12-01'), 1, 'only Tue Dec 1 (Mon Nov 30 is still break)');
  assert.strictEqual(page.countSchoolDays('2026-11-02', '2026-11-06'), 5, 'a full week');
  assert.strictEqual(page.countSchoolDays('2027-02-01', '2027-02-05'), 3, 'snow days off');
});

test('thaw: a Peaks Freeze needs every group goal and staff sign-off; a Fun Freeze cashes out by tier, melts at Thaw By', () => {
  const w = world();
  const today = w.todayYmd_(), ago = n => w.ymdShift_(today, -n);
  w.ensureFreezeTab_();
  const fz = w.book.getSheetByName('Point Freezes');
  assert.deepStrictEqual(fz.rows[0].slice(5), ['Kind', 'Thaw Goals', 'Thaw By', 'Staff Sign-off', 'Thawed', 'Second Chance']);
  // Peaks: GOLD must earn 10 points together, 50% must pitch in, and have an advisor reflection.
  fz.rows.push(['GOLD', ago(3), '', 'Peaks expectations as a squad', 'Mr. S', 'Peaks', 'points: 10; everyone: 50%; Squad reflection with advisor', '', false, '', '']);
  const ev = w.book.getSheetByName('events');
  ev.rows.push([201, 'Trait Dough: Ann L.', 'trait_dough', ago(1), '', '', 0, 12, 0, 0]);
  // Two GOLD students, both pitched in since the freeze began (made-up names).
  w.book.insertSheet('Student Board').rows = [['800#', 'Name', 'Squad', 'Grade', 'Trait Dough Bills', 'Bills This Week', 'Trait Dough Wins', 'Mindset Summits', 'Trail Journal Growth', 'STAR Growth Points', 'Shout-outs', 'Streak (weeks)', 'Badges', 'Last Contribution', 'Updated'],
    ['8001111', 'Ann L.', 'Gold', '7', '', '', '', '', '', '', '', '', '', ago(1), ''], ['8002222', 'Ed P.', 'Gold', '8', '', '', '', '', '', '', '', '', '', ago(2), '']];
  let d = J(w.freezeDetail_('GOLD'));
  assert.strictEqual(d.kind, 'Peaks');
  assert.deepStrictEqual(d.goals.map(g => g.met), [true, true, false], '12 frozen points, everyone pitched in, staff goal waits on sign-off');
  assert.strictEqual(d.ready, false);
  w.checkThaws_();
  assert.ok(w.freezeToday_('GOLD'), 'still frozen without the sign-off');
  fz.rows[1][8] = true;
  assert.strictEqual(J(w.freezeDetail_('GOLD')).ready, true);
  w.checkThaws_();
  assert.strictEqual(w.freezeToday_('GOLD'), null, 'thawed together');
  assert.ok(/^Thawed together/.test(fz.rows[1][9]));
  assert.ok(!ev.rows.some(r => r[2] === 'freeze_thaw'), 'a Peaks thaw gives nothing back');
  // Fun: GREY's goal is 5 points; 11 frozen out = past twice the goal = 50% back.
  fz.rows.push(['GREY', ago(2), '', 'Freeze card from the Wheel', 'Mr. S', 'Fun', 'points: 5', w.ymdShift_(today, 3), false, '', '']);
  ev.rows.push([202, 'Poll', 'daily_poll', ago(1), '', '', 0, 0, 11, 0]);
  d = J(w.freezeDetail_('GREY'));
  assert.deepStrictEqual([d.kind, d.pct, d.iced, d.cash], ['Fun', 50, 11, 6]);
  assert.deepStrictEqual(d.tiers.map(t => t.reached), [true, true, false]);
  // A General cashes out: +6 back as a freeze_thaw event, and the freeze ends.
  w.CacheService.getScriptCache().put('hq-tok', JSON.stringify({ id: '8005678', name: 'Bo K.', squad: 'GREY' }));
  w.clashGeneralsHq = () => '{}';
  w.clashGeneralsCashOut('tok');
  const back = ev.rows.find(r => r[2] === 'freeze_thaw');
  assert.ok(back && back[8] === 6 && back[3] === today, 'GREY gets 6 points back today');
  assert.strictEqual(w.freezeToday_('GREY'), null);
  assert.strictEqual(J(w.squadTotals_()).GREY >= 6, true, 'the points back count (not frozen)');
  // Fun that melts: PURPLE's Thaw By passed with nothing reached.
  fz.rows.push(['PURPLE', ago(5), '', 'Freeze card', 'Mr. S', 'Fun', 'points: 50', ago(1), false, '', '']);
  w.checkThaws_();
  assert.ok(/^Melted/.test(fz.rows[3][9]) && w.freezeToday_('PURPLE') === null, 'melted at Thaw By, nothing back');
  // Second chance moves Thaw By once.
  fz.rows.push(['BLACK', ago(1), '', 'Freeze card', 'Mr. S', 'Fun', 'points: 50', today, false, '', '']);
  w.CacheService.getScriptCache().put('hq-tok2', JSON.stringify({ id: '8000001', name: 'Cy R.', squad: 'BLACK' }));
  w.clashGeneralsSecondChance('tok2');
  assert.ok(fz.rows[4][7] > today && /^Used/.test(fz.rows[4][10]));
  assert.throws(() => w.clashGeneralsSecondChance('tok2'), /already used/);
});

test('freeze update: votes thaw goal, January fresh start (Peaks ends, Fun cashes out, the day itself counts, once), student release, meeting slides', () => {
  const w = world();
  w.ensureSettings_(); w.ensureFreezeTab_(); w.ensurePollTabs_();
  const today = w.todayYmd_(), ago = n => w.ymdShift_(today, -n);
  const fz = w.book.getSheetByName('Point Freezes'), ev = w.book.getSheetByName('events'), votes = w.book.getSheetByName('Poll Votes');
  w.book.insertSheet('Student Board').rows = [['800#', 'Squad', 'Last Contribution'], ['8001111', 'Gold', ''], ['8002222', 'Purple', '']];
  // Votes goal: poll and Pulse votes since the freeze began.
  fz.rows.push(['GOLD', ago(5), '', 'Peaks', 'Mr. S', 'Peaks', 'votes: 3; points: 1', '', false, '', '']);
  votes.rows.push([2, ago(9), '8001111', 'GOLD', 'A', ''], [2, ago(2), '8001111', 'GOLD', 'A', ''], [3, ago(1), '8001111', 'GOLD', 'B', ''], [3, ago(1), '8002222', 'PURPLE', 'B', '']);
  let d = J(w.freezeDetail_('GOLD'));
  assert.deepStrictEqual(d.goals[0], { text: 'Cast 3 poll and Pulse votes', type: 'votes', have: 2, need: 3, met: false }, 'only GOLD\'s votes since the freeze began');
  // PURPLE: a Fun Freeze with 12 points frozen (goal 5: past twice = 50%).
  fz.rows.push(['PURPLE', ago(4), '', 'Wheel', 'Mr. S', 'Fun', 'points: 5', w.ymdShift_(today, 9), false, '', '']);
  ev.rows.push([601, 'Spirit', 'event', ago(2), '', '', 0, 0, 0, 12]);
  w.readEventsTab_(); // freezes zero the row and note it
  // A freeze set on/after the fresh start is untouched.
  fz.rows.push(['GREY', today, '', 'new one', 'Mr. S', 'Peaks', '', '', false, '', '']);
  // Fresh start today: points earned today count even before the check runs.
  w.setSettingValue_('FRESH_START', today);
  ev.rows.push([602, 'Fresh day', 'event', today, '', '', 0, 4, 0, 0]);
  assert.strictEqual(w.freezeToday_('GOLD'), null, 'the window stops the day before the fresh start');
  assert.strictEqual(J(w.readEventsTab_()).find(e => e.id === 602).pts.GOLD, 4, 'fresh-start day points count for GOLD');
  const lines = J(w.checkThaws_());
  assert.ok(lines.some(l => /^GOLD: Fresh start/.test(l)) && lines.some(l => /^PURPLE: Fresh start .*cashed out 50% \(\+6\)/.test(l)), lines.join(' | '));
  assert.strictEqual(fz.rows[1][2], ago(1));
  assert.ok(/^Fresh start/.test(fz.rows[1][9]));
  const back = ev.rows.filter(r => r[2] === 'freeze_thaw');
  assert.strictEqual(back.length, 1);
  assert.strictEqual(back[0][9], 6, 'PURPLE gets 50% of 12 back');
  assert.ok(w.freezeToday_('GREY'), 'a freeze set on the fresh-start day stays');
  assert.deepStrictEqual(J(w.checkThaws_()).filter(l => /Fresh/.test(l)), [], 'once only');
  // Before the date nothing happens.
  w.setSettingValue_('FRESH_START', w.ymdShift_(today, 30));
  assert.deepStrictEqual(J(w.applyFreshStart_(null)), []);
  // The student release (read aloud) says it plainly.
  const rel = w.freezeRelease_();
  assert.ok(/no squad ever loses points/.test(rel) && /Peaks Freeze/.test(rel) && /Fun Freeze/.test(rel) && /poll and Pulse vote/.test(rel) && /starts fresh/.test(rel), rel);
  // Meeting slides data: the fresh-start slides show near FRESH_START; 2x zones listed.
  w.setSettingValue_('FRESH_START', w.ymdShift_(today, 7));
  w.ensureGameTabs_();
  const m = J(w.meetingData_('GOLD'));
  assert.ok(m.fresh && m.fresh.date === w.ymdShift_(today, 7));
  assert.ok(m.doubles.some(z => z.zone === 'Gym 2' && z.mult === 2) && m.doubles.some(z => z.zone === 'Gym 1') && !m.doubles.some(z => z.zone === '70s Hall' || z.zone === 'Media Center'));
});

test('game layer: map claims, wing bonus, chance cards + Flip Three, Hot Streak, round goals, monster, tickets, My Climb sign-in', () => {
  const w = world();
  const today = w.todayYmd_(), ago = n => w.ymdShift_(today, -n);
  w.ensureSettings_(); w.ensureGameTabs_(); w.ensureMeTabs_(); w.ensurePollTabs_();
  w.ScriptApp = { getProjectTriggers: () => [], newTrigger: () => ({ timeBased: () => ({ everyHours: () => ({ create() {} }) }) }) };
  // This test plays the rolling-window rules (the set before the monthly rules, which have their own test).
  w.writeRowsByHeader_(w.book.getSheetByName('Conquest Rules'), [
    { Show: true, Rule: 'Full Climb', Method: 'most points', Counts: '*', 'Window Days': 14 },
    { Show: true, Rule: 'Game Day', Method: 'most points', Counts: 'mayhem, gaming, field, duel_bonus', 'Window Days': 21 },
    { Show: true, Rule: 'Poll Turnout', Method: 'most points', Counts: 'daily_poll', 'Window Days': 14 },
    { Show: true, Rule: 'Trait Dough Drive', Method: 'most points', Counts: 'trait_dough', 'Window Days': 14 },
    { Show: true, Rule: 'Growth Grind', Method: 'most points', Counts: 'star_growth, mountaineer_mindset', 'Window Days': 30 },
    { Show: true, Rule: 'Shout-out Storm', Method: 'most entries', Counts: 'shout, trait_dough', 'Window Days': 14 },
    { Show: true, Rule: 'Lunch Battles', Method: 'most points', Counts: 'lunch, mayhem_week', 'Window Days': 21 },
    { Show: true, Rule: 'Siege', Method: 'most points', Counts: '*', 'Window Days': 14, 'Must Border': true, 'Defense %': 25, 'Far Attack %': 75 },
    { Show: true, Rule: 'Staff Pick', Method: 'staff pick' }]);
  const oldRule = { 'Gym 2': 'Game Day', 'Gym 1': 'Game Day', Cafeteria: 'Full Climb', 'Music Wing': 'Poll Turnout', Auditorium: 'Shout-out Storm', '50s Hall': 'Trait Dough Drive',
    '60s Hall': 'Siege', 'Media Center': 'Growth Grind', 'Front Office': 'Staff Pick', '70s Hall': 'Lunch Battles', '80s Hall': 'Siege', '90s Hall': 'Trait Dough Drive' };
  w.writeRowsByHeader_(w.book.getSheetByName('Map Zones'), J(w.sheetRows_('Map Zones')).map(r => Object.assign(r, { 'Conquer By': oldRule[r.Zone], Multiplier: '' })));
  // Students (made-up names) with emails and attendance from Climb-er.
  w.book.insertSheet('Student Board').rows = [['800#', 'Name', 'Squad', 'Grade', 'Trait Dough Bills', 'Bills This Week', 'Trait Dough Wins', 'Mindset Summits', 'Trail Journal Growth', 'STAR Growth Points', 'Shout-outs', 'Ambassador Buddies', 'Streak (weeks)', 'Badges', 'Last Contribution', 'Updated', 'Attendance %', 'Email'],
    ['8001111', 'Ann L.', 'Gold', '7', '4', '1', '0', '1', '', '2', '0', '0', '2', '', ago(1), '', '96', 'ann.l@easdpa.org'],
    ['8002222', 'Ed P.', 'Gold', '8', '0', '0', '0', '0', '', '0', '1', '0', '0', '', '', '', '90', 'ed.p@easdpa.org'],
    ['8003333', 'Cy R.', 'Grey', '7', '1', '0', '0', '0', '', '0', '0', '0', '0', '', ago(3), '', '', '']];
  const ev = w.book.getSheetByName('events');
  // GOLD wins the poll zone and everything overall; GREY wins the Trait Dough zones (building map).
  ev.rows.push([301, 'Daily Poll', 'daily_poll', ago(2), '', '', 0, 9, 2, 0], [302, 'Trait Dough: Cy R.', 'trait_dough', ago(1), '', '', 0, 0, 5, 0]);
  let zones = J(w.mapZones_(null, true));
  const z = name => zones.find(x => x.zone === name);
  assert.strictEqual(z('Music Wing').holder, 'GOLD');
  assert.strictEqual(z('50s Hall').holder, 'GREY');
  assert.strictEqual(z('Cafeteria').holder, 'GOLD', 'King of the Mountain (the Cafeteria) = most points overall');
  // King of the hill: a tie doesn't take it away.
  ev.rows.push([303, 'Trait Dough: Ann L.', 'trait_dough', ago(1), '', '', 0, 5, 0, 0]);
  zones = J(w.mapZones_(null, true));
  assert.strictEqual(z('50s Hall').holder, 'GREY', 'tie: the holder keeps it');
  // A wing held whole pays again: give GOLD all of The Arts.
  const mz = w.book.getSheetByName('Map Zones');
  mz.rows.forEach(r => { if (r[2] === 'The Arts') r[10] = 'GOLD'; });
  w.setSettingValue_('GAME_AWARDS', 'ON');
  w.props['game-paid-week'] = 'old';
  w.gameTick();
  const terr = ev.rows.filter(r => r[2] === 'territory');
  const goldTerr = terr.find(r => r[7] > 0);
  // GOLD: Cafeteria (summit) 5x2 + Music Wing 3 + Auditorium 3 + 60s and 80s Halls 3+3 (Siege: a squad with
  // no zones lands anywhere) + The Arts again (3+3) = 28, + 1 per zone from Switchback Builders = 33.
  assert.strictEqual(goldTerr[7], 33, 'zones, the Summit doubled, the wing paid twice, +5 from the Switchback Builders engine power (5 zones)');
  w.gameTick();
  assert.strictEqual(ev.rows.filter(r => r[2] === 'territory').length, terr.length, 'paid once a week');
  // Chance cards: a General draws one a week.
  const cache = w.CacheService.getScriptCache();
  cache.put('hq-g1', JSON.stringify({ id: '8001111', name: 'Ann L.', squad: 'GOLD' }));
  w.clashGeneralsHq = () => '{}';
  const drew = JSON.parse(w.clashGeneralsDrawCard('g1')).drew;
  assert.ok(drew && drew.card);
  assert.throws(() => w.clashGeneralsDrawCard('g1'), /already drew/);
  // Flip Three: GREY flips three; a card that comes up twice busts (here every card is the same).
  const cards = w.book.getSheetByName('Chance Cards');
  cards.rows.forEach((r, i) => { if (i) r[0] = r[1] === 'Lucky Switchback'; });
  const hand = w.book.getSheetByName('Squad Hand');
  hand.rows.push(['GOLD', 'Flip Three', 'flip_three', 0, today, 'Held', '']);
  const flipped = JSON.parse(w.clashGeneralsPlayCard('g1', hand.rows.length, 'GREY')).flipped;
  assert.deepStrictEqual(flipped.map(f => f.status), ['Busted', 'Busted', 'Busted'], 'three of the same card: all bust, no points lost');
  assert.throws(() => w.clashGeneralsPlayCard('g1', hand.rows.length, 'GREY'), /isn't in your hand/);
  // Hot Streak: GOLD topped poll turnout three finished days in a row, then banks it.
  const polls = w.book.getSheetByName('Daily Polls'), votes = w.book.getSheetByName('Poll Votes');
  [3, 2, 1].forEach(n => {
    polls.rows.push([true, ago(n), 'This or That', 'Q' + n, 'A | B', '', 1, 1, '', '']);
    const id = polls.rows.length;
    votes.rows.push([id, ago(n), '8001111', 'GOLD', 'A', ''], [id, ago(n), '8002222', 'GOLD', 'B', ''], [id, ago(n), '8003333', 'GREY', 'A', '']);
  });
  const st = J(w.hotStreaks_());
  assert.strictEqual(st.GOLD.length, 3);
  assert.strictEqual(st.GREY.length, 0);
  w.clashGeneralsBankStreak('g1');
  assert.ok(ev.rows.some(r => r[2] === 'streak_bank' && r[7] === 6), '3 days banks 4 points, +2 from the Rope Team engine power (GOLD holds a whole wing)');
  assert.strictEqual(J(w.hotStreaks_()).GOLD.length, 0, 'the streak starts over after banking');
  // Round goal that ended yesterday: places get prizes once.
  const rg = w.book.getSheetByName('Round Goals');
  rg.rows.push([true, ago(10), ago(1), 'Most voters', 'voters', 8, 5, 3, '', '']);
  w.gameTick();
  const rgEv = ev.rows.find(r => r[2] === 'round_goal' && /Most voters/.test(r[1]));
  assert.ok(rgEv && rgEv[7] === 8 && rgEv[8] === 5, 'GOLD 1st, GREY 2nd');
  // The Opponent (was Monster of the Month): hidden until shown; defeated together.
  assert.strictEqual(w.monster_(), null, 'seeded monsters start hidden');
  const mon = w.book.getSheetByName('Opponents');
  mon.rows.push([true, ago(5), w.ymdShift_(today, 5), 'Test Beast', 'Grr', '', 'voters: 9', 2, '']);
  assert.ok(J(w.monster_()).defeated, '9 votes this week defeats it');
  w.gameTick();
  assert.strictEqual(ev.rows.filter(r => r[2] === 'monster').length, 1);
  w.setSettingValue_('OPPONENT_VISIBLE', 'OFF');
  assert.strictEqual(w.monster_(), null, 'one switch hides it everywhere');
  // My Climb is school-day only: 7:00 to 4:00, weekdays not on the School Calendar.
  w.ensureSchoolCalendarTab_();
  assert.strictEqual(w.climbOpen_(new Date('2026-11-02T10:00:00Z')), true, 'Monday 10:00');
  assert.strictEqual(w.climbOpen_(new Date('2026-11-02T06:30:00Z')), false, 'before 7');
  assert.strictEqual(w.climbOpen_(new Date('2026-11-02T16:00:00Z')), false, 'at 4 it closes');
  assert.strictEqual(w.climbOpen_(new Date('2026-11-07T12:00:00Z')), false, 'Saturday');
  assert.strictEqual(w.climbOpen_(new Date('2026-11-26T12:00:00Z')), false, 'Thanksgiving (School Calendar)');
  w.setSettingValue_('MY_CLIMB_HOURS', '00:00-23:59'); w.setSettingValue_('MY_CLIMB_SCHOOL_DAYS', 'OFF'); // the rest of this test runs any time
  // My Climb: Ann signs in with her school account.
  w.Session.getActiveUser = () => ({ getEmail: () => 'Ann.L@easdpa.org' });
  let me = JSON.parse(w.clashMe());
  assert.strictEqual(me.me.id, '8001111');
  const att = me.metrics.find(m => m.key === 'attendance');
  assert.deepStrictEqual([att.me, att.squadAvg], [96, 93], 'her attendance next to the GOLD average');
  assert.strictEqual(me.mascot.animal, 'Cottontail rabbit');
  w.Session.getActiveUser = () => ({ getEmail: () => 'someone@else.org' });
  assert.throws(() => w.clashMe(), /NOT_FOUND/);
  w.Session.getActiveUser = () => ({ getEmail: () => '' });
  assert.throws(() => w.clashMe(), /SIGN_IN/);
  // Staff preview needs an admin session.
  assert.throws(() => w.clashMe('8001111', 'nope'));
  const tok = login(w, 'oldpw');
  // Destination tickets: draw three, keep one; finishing it pays the squad once.
  const offer = JSON.parse(w.clashMeDrawTickets('8001111', tok)).offer;
  assert.strictEqual(offer.length, 3);
  const td = w.book.getSheetByName('Destination Tickets');
  td.rows.push([true, 'Poll Pal', 'Vote once.', 'polls: 1', 3]);
  cache.put('tickets-8001111', JSON.stringify(['Poll Pal']));
  me = JSON.parse(w.clashMeKeepTickets(['Poll Pal'], '8001111', tok));
  assert.strictEqual(me.tickets.find(t => t.ticket === 'Poll Pal').status, 'Holding');
  votes.rows.push([99, today, '8001111', 'GOLD', 'A', '']);
  me = JSON.parse(w.clashMe('8001111', tok));
  assert.strictEqual(me.tickets.find(t => t.ticket === 'Poll Pal').status, 'Completed');
  assert.ok(ev.rows.some(r => r[2] === 'destination_ticket' && r[7] === 3 && /Ann L\./.test(r[1])));
  JSON.parse(w.clashMe('8001111', tok));
  assert.strictEqual(ev.rows.filter(r => r[2] === 'destination_ticket').length, 1, 'paid once');
  // My climber: only allowed choices are saved; gear follows what she's done.
  assert.deepStrictEqual(JSON.parse(w.clashMeSaveLook({ fur: 2, eyes: 99, hat: 1 }, '8001111', tok)).look, { fur: 2, eyes: 0, mood: 0, hat: 1, accent: 0 });
  me = JSON.parse(w.clashMe('8001111', tok));
  assert.ok(me.gear.find(g => g.gear === 'boots').have && !me.gear.find(g => g.gear === 'lantern').have);
  // The public site never gets other students' numbers.
  assert.ok(!/8002222|ed\.p@/.test(w.clashGame()));
});

test('conquest: rules from the Sheet, Risk borders + defense edge, staff pick, entries, battle log, admin map editor, building swap', () => {
  const w = world();
  const today = w.todayYmd_(), ago = n => w.ymdShift_(today, -n);
  w.ensureSettings_(); w.ensureGameTabs_();
  const ev = w.book.getSheetByName('events'), mz = w.book.getSheetByName('Map Zones');
  // A small map: West | Middle | East in a row, plus Island (not touching) joined to West by Borders.
  const zoneRows = [
    { Show: true, Zone: 'West', Wing: 'Row', 'Conquer By': 'Open', 'Weekly Points': 3, X: 0, Y: 0, W: 2, H: 2 },
    { Show: true, Zone: 'Middle', Wing: 'Row', 'Conquer By': 'Siege', 'Weekly Points': 3, X: 2, Y: 0, W: 2, H: 2 },
    { Show: true, Zone: 'East', Wing: 'Row', 'Conquer By': 'Open', 'Weekly Points': 3, X: 4, Y: 0, W: 2, H: 2, Holder: 'PURPLE' },
    { Show: true, Zone: 'Island', 'Conquer By': 'Count', 'Weekly Points': 2, X: 8, Y: 5, W: 1, H: 1, Borders: 'west' },
    { Show: true, Zone: 'Office', 'Conquer By': 'Pick', X: 0, Y: 4, W: 1, H: 1, Holder: 'BLACK' },
  ];
  w.writeRowsByHeader_(mz, zoneRows);
  w.writeRowsByHeader_(w.book.getSheetByName('Conquest Rules'), [
    { Show: true, Rule: 'Open', Method: 'most points', Counts: 'trait_dough', 'Window Days': 7 },
    { Show: true, Rule: 'Siege', Method: 'most points', Counts: '*', 'Window Days': 7, 'Must Border': true, 'Defense %': 50 }, // no Far Attack: next door only
    { Show: true, Rule: 'Count', Method: 'most entries', Counts: 'shout' },
    { Show: true, Rule: 'Pick', Method: 'staff pick' },
  ]);
  // GOLD tops Trait Dough (West, East); GREY tops everything overall and holds nothing yet.
  ev.rows.push([401, 'Trait Dough', 'trait_dough', ago(1), '', '', 0, 6, 2, 0], [402, 'Kickoff', 'event', ago(1), '', '', 0, 0, 30, 0],
    [403, 'Shout-out', 'shout', ago(1), '', '', 0, 1, 0, 0], [404, 'Shout-out', 'shout', ago(2), '', '', 0, 1, 0, 0], [405, 'Shout-out', 'shout', ago(1), '', '', 0, 0, 9, 0]);
  let zones = J(w.mapZones_(null, true));
  const z = n => zones.find(x => x.zone === n);
  assert.strictEqual(z('West').holder, 'GOLD');
  assert.strictEqual(z('East').holder, 'GOLD', 'Trait Dough: GOLD 6 beats the holder PURPLE 0');
  assert.strictEqual(z('Middle').holder, 'GREY', 'a squad with no zones can land anywhere, even a Siege zone');
  assert.strictEqual(z('Island').holder, 'GOLD', 'most entries: GOLD 2 shout-outs beats GREY 1 (even though GREY earned 9 points)');
  assert.strictEqual(z('Office').holder, 'BLACK', 'staff pick never changes on its own');
  assert.deepStrictEqual(z('Island').neighbors, ['West'], 'Borders joins zones that do not touch');
  assert.deepStrictEqual(z('Middle').neighbors.slice().sort(), ['East', 'West']);
  assert.ok(!z('West').neighbors.includes('Office'), 'boxes that do not touch are not neighbors');
  const battles = w.book.getSheetByName('Map Battles').rows.slice(1);
  assert.ok(battles.some(b => b[1] === 'East' && b[2] === 'GOLD' && b[3] === 'PURPLE' && b[4] === 'Open'), 'captures are logged');
  // Siege: GOLD (next to Middle) now out-earns GREY, but needs 50% more than GREY to take it.
  ev.rows.push([406, 'Spirit', 'event', ago(1), '', '', 0, 34, 0, 0]); // GOLD 42 vs GREY 41: not enough (needs 61.5)
  zones = J(w.mapZones_(null, true));
  assert.strictEqual(z('Middle').holder, 'GREY', 'defense edge: 42 is not 50% more than 41');
  ev.rows.push([407, 'Spirit 2', 'event', ago(1), '', '', 0, 20, 0, 0]); // GOLD 62 > 61.5
  zones = J(w.mapZones_(null, true));
  assert.strictEqual(z('Middle').holder, 'GOLD', 'past the defense edge, a neighbor takes it');
  // Must Border: BLACK holds only the Office (not next to Middle) and can't attack it however much it earns.
  w.writeRowsByHeader_(mz, zoneRows.map(r => Object.assign({}, r, r.Zone === 'West' || r.Zone === 'East' ? { Holder: 'GOLD' } : {})));
  ev.rows.push([408, 'Black Out', 'event', ago(1), '', '', 500, 0, 0, 0]);
  zones = J(w.mapZones_(null, false));
  assert.strictEqual(z('Middle').holder, 'GOLD', 'BLACK leads but does not border Middle; GOLD is the best squad that does');
  // Far Attack: with it set, a squad from farther away can take a Siege zone, but needs the bigger lead.
  const sieges = w.book.getSheetByName('Conquest Rules');
  const farCol = sieges.rows[0].indexOf('Far Attack %');
  assert.ok(farCol > -1);
  sieges.rows.forEach(r => { if (r[1] === 'Siege') r[farCol] = 1000; });
  w.writeRowsByHeader_(mz, zoneRows.map(r => Object.assign({}, r, r.Zone === 'Middle' ? { Holder: 'GOLD' } : r.Zone === 'West' || r.Zone === 'East' ? { Holder: 'GOLD' } : {})));
  zones = J(w.mapZones_(null, false));
  assert.strictEqual(z('Middle').holder, 'GOLD', 'BLACK 500 from afar is not 1000% more than GOLD 62');
  sieges.rows.forEach(r => { if (r[1] === 'Siege') r[farCol] = 75; });
  zones = J(w.mapZones_(null, false));
  assert.strictEqual(z('Middle').holder, 'BLACK', 'BLACK 500 from afar is more than 75% over GOLD 62');
  // Both ways: GOLD (next door) takes it back with just the 50% edge, not the far bar.
  w.writeRowsByHeader_(mz, zoneRows.map(r => Object.assign({}, r, r.Zone === 'Middle' ? { Holder: 'BLACK' } : r.Zone === 'West' || r.Zone === 'East' ? { Holder: 'GOLD' } : {})));
  ev.rows.push([409, 'Comeback', 'event', ago(1), '', '', 0, 700, 0, 0]); // GOLD 762 > BLACK 500 x 1.5 = 750
  zones = J(w.mapZones_(null, false));
  assert.strictEqual(z('Middle').holder, 'GOLD', 'next door only needs the Defense edge');
  // Open zone: a far squad must beat the best squad next door by the far bar.
  w.writeRowsByHeader_(mz, zoneRows.map(r => Object.assign({}, r, r.Zone === 'West' || r.Zone === 'East' ? { Holder: 'GOLD' } : {})));
  zones = J(w.mapZones_(null, false));
  assert.strictEqual(z('Middle').holder, 'GOLD', 'BLACK 500 from afar does not beat GOLD 762 next door');
  // The page payload: grid fits the zones, rules explained, recent battles.
  const pay = J(w.mapPayload_(null));
  assert.deepStrictEqual(pay.grid, { cols: 12, rows: 8 });
  assert.ok(pay.rules.find(r => r.rule === 'Siege' && r.border && r.defense === 50 && r.far === 75 && r.zones.includes('Middle')));
  assert.ok(pay.battles.length >= 1 && pay.battles[0].zone);
  assert.strictEqual(pay.style, 'puzzle');
  // Admin map editor: needs a signed-in admin, round-trips zones and rules, a staff-set holder starts today.
  assert.throws(() => w.clashMapLoad('nope'));
  assert.throws(() => w.clashMapSave('nope', '{}'));
  w.CacheService.getScriptCache().put('clash-admin-T', '1');
  const loaded = JSON.parse(w.clashMapLoad('T'));
  assert.strictEqual(loaded.zones.length, 5);
  assert.ok(loaded.building.zones.length >= 10 && loaded.methods.includes('staff pick'));
  const west = loaded.zones.find(x => x.zone === 'West');
  west.x = 1; west.zone = 'West Wing'; loaded.zones.find(x => x.zone === 'Office').holder = 'GREY';
  loaded.rules.push({ show: true, rule: 'New Idea', text: 'Most poll points', method: 'most points', counts: 'daily_poll', days: 10, border: false, defense: 0 });
  w.clashMapSave('T', JSON.stringify(loaded));
  const again = JSON.parse(w.clashMapLoad('T'));
  assert.strictEqual(again.zones.find(x => x.zone === 'West Wing').x, 1);
  assert.strictEqual(again.zones.find(x => x.zone === 'Office').holder, 'GREY');
  assert.strictEqual(again.zones.find(x => x.zone === 'Office').since, today, 'a staff-set holder starts today');
  assert.ok(again.rules.find(r => r.rule === 'New Idea' && r.days === 10));
  loaded.rules.push({ rule: 'new idea' });
  assert.throws(() => w.clashMapSave('T', JSON.stringify(loaded)), /Two rules/);
  // An older Map Zones tab (no Conquer By / Borders) gets the columns; the building swap keeps old rules.
  const w2 = world(); w2.ensureSettings_();
  const old = w2.book.insertSheet('Map Zones');
  old.rows = [['Show', 'Zone', 'Wing', 'Summit', 'Claimed By', 'Weekly Points', 'X', 'Y', 'W', 'H', 'Holder', 'Since'], [true, 'The Summit', 'Summit', true, '*', 5, 5, 0, 2, 2, 'GOLD', '2026-11-03']];
  w2.ensureGameTabs_();
  assert.deepStrictEqual(old.rows[0].slice(-4), ['Conquer By', 'Borders', 'Multiplier', 'Locked']);
  assert.strictEqual(J(w2.mapZones_(null, false))[0].holder, 'GOLD', 'old zones keep working with no Conquer By');
  w2.book.getSheetByName('Conquest Rules').rows.push([true, 'My Rule', '', 'most points', '*', 7, false, 0]);
  w2.applyBuildingMap_();
  assert.ok(old.rows.slice(1).some(r => r[1] === 'Cafeteria'));
  assert.ok(!old.rows.slice(1).some(r => r[1] === 'The Summit'));
  assert.ok(w2.book.getSheetByName('Conquest Rules').rows.some(r => r[1] === 'My Rule'));
  // The staff guide tab comes with it, headings without their marker.
  const guide = w2.book.getSheetByName('How The Map Works').rows.map(r => r[0]);
  assert.ok(guide.includes('WHAT OWNING PAYS (THE MULTIPLIERS)') && guide.some(t => /WING_MULTIPLIER/.test(t)) && guide.some(t => /Far Attack %/.test(t)));
  assert.ok(!guide.some(t => /^## /.test(t)));
  w2.setupClimbGuideTab_();
  assert.ok(w2.book.getSheetByName('How My Climb Works').rows.some(r => /EVOLUTION BOOSTS/.test(r[0])));
});

test('monthly zones: last calendar month, highest/lowest numbers from Squad Monthly (Climb-er), per 100, waits for every squad, Multiplier pays, staff challenge, menu', () => {
  const w = world();
  w.ensureSettings_(); w.ensureGameTabs_(); w.ensurePollTabs_();
  const lm = w.monthBounds_(-1), ym = lm[0].slice(0, 7), mid = lm[0].slice(0, 8) + '15', thisMonth = w.todayYmd_();
  const ev = w.book.getSheetByName('events'), mz = w.book.getSheetByName('Map Zones');
  // A fresh setup seeds the monthly set: every building zone points at its monthly rule.
  const seeded = J(w.sheetRows_('Map Zones'));
  assert.strictEqual(seeded.find(r => r.Zone === 'Front Office')['Conquer By'], 'Staff Challenge');
  assert.strictEqual(seeded.find(r => r.Zone === 'Gym 2').Multiplier, 2, 'Pride Day: a check-in, 2x');
  assert.strictEqual(seeded.find(r => r.Zone === 'Gym 1').Multiplier, 2, 'Move · Make · Talk check-ins: 2x');
  assert.strictEqual(seeded.find(r => r.Zone === '70s Hall').Multiplier, 1.5, 'grades are a school record: 1.5x');
  assert.strictEqual(seeded.find(r => r.Zone === 'Media Center').Multiplier, 1.5);
  assert.strictEqual(seeded.find(r => r.Zone === '80s Hall').Multiplier, 1, 'points: 1x');
  assert.strictEqual(seeded.find(r => r.Zone === 'Gym 1')['Conquer By'], 'Move · Make · Talk');
  // Students on the board: GOLD 2, the rest 1 each (Students on Squad Monthly wins when filled in).
  w.book.insertSheet('Student Board').rows = [['800#', 'Squad'], ['8000001', 'Gold'], ['8000002', 'Gold'], ['8000003', 'Grey'], ['8000004', 'Black'], ['8000005', 'Purple']];
  const sm = w.book.getSheetByName('Squad Monthly');
  assert.ok(sm && sm.rows[0].includes('lrr_ld') && sm.rows[0].includes('interventions_admin'));
  let zones = J(w.mapZones_(null, true));
  const z = n => zones.find(x => x.zone === n);
  ['Cafeteria', '70s Hall', '90s Hall', 'Gym 2', '50s Hall'].forEach(n => assert.strictEqual(z(n).holder, '', n + ' waits: no numbers for the month yet'));
  // Climb-er sends last month (Month as a date cell, the way Sheets stores "2026-12"). GOLD has the most
  // LRR/LDs (6) but twice the students, so per 100 GOLD (300) beats BLACK (400).
  const H = sm.rows[0];
  const row = (sq, o) => H.map(h => h === 'Month' ? lm[0] : h === 'Squad' ? sq : o[h] === undefined ? '' : o[h]);
  sm.rows.push(row('GOLD', { lrr_ld: 6, failing: 2, interventions_teacher: 1, interventions_admin: 1, attendance: '94.5%', pride_day: 70, shout_outs: 12 }),
    row('BLACK', { lrr_ld: 4, failing: 0, interventions_teacher: 3, interventions_admin: 0, attendance: 95, pride_day: 80, shout_outs: 9 }),
    row('GREY', { lrr_ld: 5, failing: 1, interventions_teacher: 0, interventions_admin: 0, attendance: 93, pride_day: 60, shout_outs: 20 }));
  zones = J(w.mapZones_(null, true));
  assert.strictEqual(z('Cafeteria').holder, '', 'PURPLE has no row yet: nothing moves');
  sm.rows.push(row('PURPLE', { lrr_ld: 5, failing: 3, interventions_teacher: 2, interventions_admin: 2, attendance: 92, pride_day: 75, shout_outs: 3 }));
  // An older month never counts.
  sm.rows.push(H.map(h => h === 'Month' ? '2020-01' : h === 'Squad' ? 'PURPLE' : h === 'lrr_ld' ? 0 : ''));
  zones = J(w.mapZones_(null, true));
  assert.strictEqual(z('Cafeteria').holder, 'GOLD', 'fewest LRR/LDs per 100: GOLD 300 < GREY/PURPLE 500, BLACK 400');
  assert.strictEqual(z('Cafeteria').score.GOLD, 300);
  assert.strictEqual(z('70s Hall').holder, 'BLACK', 'zero failing grades is the best there is');
  assert.strictEqual(z('90s Hall').holder, 'GREY', 'teacher + admin interventions added together: GREY 0');
  assert.strictEqual(z('50s Hall').holder, 'BLACK', 'highest attendance (95 beats "94.5%")');
  assert.strictEqual(z('Gym 2').holder, 'BLACK', 'Pride Day participation');
  assert.strictEqual(z('60s Hall').holder, 'GREY', 'most shout-outs');
  assert.strictEqual(z('Cafeteria').window, ym);
  // Events in last month count; this month's do not.
  // Clash squad events: Move / Make / Talk types (Gym 1). A plain "event" row never counts there.
  ev.rows.push([501, 'Relay race', 'squad_move', mid, '', '', 0, 3, 0, 5], [502, 'Poster build', 'squad_make', mid, '', '', 0, 0, 0, 4],
    [503, 'Debate', 'squad_talk', thisMonth, '', '', 90, 0, 0, 0], [508, 'Squad trivia', 'squad_talk', mid, '', '', 0, 4, 0, 0],
    [504, 'First Friday staff challenge winner', 'staff_challenge', thisMonth, '', '', 0, 0, 1, 0], [509, 'Old staff challenge', 'staff_challenge', mid, 0, 0, 0, 0, 0, 1],
    [505, 'Big day', 'event', mid, '', '', 0, 40, 0, 0], [506, 'Huge day this month', 'event', thisMonth, '', '', 900, 0, 0, 0],
    [507, 'Trait Dough', 'trait_dough', mid, '', '', 0, 0, 0, 8]);
  // Poll + Pulse votes and a wheel spin last month (Auditorium).
  const votes = w.book.getSheetByName('Poll Votes');
  votes.rows.push([2, mid, '8000003', 'GREY', 'A', ''], [2, mid, '8000001', 'GOLD', 'A', ''], [3, mid, '8000003', 'GREY', 'B', ''], [3, thisMonth, '8000004', 'BLACK', 'B', ''], [3, thisMonth, '8000005', 'BLACK', 'B', ''], [3, thisMonth, '8000004', 'BLACK', 'B', '']);
  zones = J(w.mapZones_(null, true));
  assert.strictEqual(z('Gym 1').holder, 'PURPLE', 'Move · Make · Talk: PURPLE 9 beats GOLD 7 last month (BLACK\'s 90 is this month; GOLD\'s 40 "event" is not a squad event)');
  assert.strictEqual(z('Front Office').holder, 'GREY', 'this month\'s First Friday staff challenge winner takes the Front Office right away (PURPLE\'s win was last month)');
  assert.strictEqual(z('Music Wing').holder, 'GOLD', 'Full Climb = last calendar month only (BLACK\'s 900 is this month)');
  assert.strictEqual(z('Auditorium').holder, 'GREY', 'poll votes last month: GREY 2, GOLD 1 (BLACK voted this month)');
  assert.strictEqual(z('80s Hall').holder, 'PURPLE', 'Trait Dough Drive');
  // The payload says it is last month, per 100, lowest.
  const pay = J(w.mapPayload_(null));
  assert.ok(pay.rules.find(r => r.rule === 'Clean Cafeteria' && r.month && r.per100 && r.low && r.window === ym));
  assert.ok(pay.rules.find(r => r.rule === 'Staff Challenge' && r.thisMonth && !r.month));
  // Pay: weekly points x Multiplier. GREY holds 60s (3 x1), 90s (3 x1.5), Front Office (2), Auditorium (3 x2) = 15.5, rounded.
  w.ScriptApp = { getProjectTriggers: () => [], newTrigger: () => ({ timeBased: () => ({ everyHours: () => ({ create() {} }) }) }) };
  w.writeRowsByHeader_(w.book.getSheetByName('Engine Powers'), []);
  w.setSettingValue_('GAME_AWARDS', 'ON'); w.props['game-paid-week'] = 'old';
  w.gameTick();
  const grey = ev.rows.filter(r => r[2] === 'territory' && r[8] > 0)[0];
  assert.strictEqual(grey[8], 16, 'GREY: 60s 3 + 90s 4.5 + Front Office 2 + Auditorium 6 = 15.5 -> 16');
  // Admin map editor keeps the Multiplier, Per 100 and "last month".
  w.CacheService.getScriptCache().put('clash-admin-T', '1');
  const loaded = JSON.parse(w.clashMapLoad('T'));
  assert.ok(loaded.methods.includes('lowest number'));
  assert.strictEqual(loaded.zones.find(x => x.zone === 'Gym 2').mult, 2);
  assert.ok(loaded.rules.find(r => r.rule === 'Clean Cafeteria' && r.per100 && r.days === 'last month'));
  loaded.zones.find(x => x.zone === 'Gym 1').mult = 3;
  loaded.zones.find(x => x.zone === '50s Hall').mult = 1.25;
  w.clashMapSave('T', JSON.stringify(loaded));
  const again = JSON.parse(w.clashMapLoad('T'));
  assert.strictEqual(again.zones.find(x => x.zone === 'Gym 1').mult, 3);
  assert.strictEqual(again.zones.find(x => x.zone === '50s Hall').mult, 1.25, 'any multiplier, not just whole numbers');
  assert.ok(again.rules.find(r => r.rule === 'Passing Grades' && r.per100 && r.days === 'last month' && r.method === 'lowest number'));
  // Pride Day before Climb-er sends pride_day: Pride Day rows (Admin › Add Event, Type Pride Day) decide it.
  const w3 = world(); w3.ensureSettings_(); w3.ensureGameTabs_();
  w3.book.getSheetByName('events').rows.push([701, 'Pride Day', 'pride_day', mid, '', '', 0, 1, 0, 0], [702, 'Pride Day 2', 'pride_day', mid, '', '', 0, 1, 1, 0]);
  assert.strictEqual(J(w3.mapZones_(null, false)).find(x => x.zone === 'Gym 2').holder, 'GOLD', 'GOLD 2 Pride Day rows, GREY 1');
  // The menu on an older map: rules move over, a custom rule stays, January can start everyone fresh.
  const w2 = world(); w2.ensureSettings_(); w2.ensureGameTabs_();
  const cr2 = w2.book.getSheetByName('Conquest Rules');
  w2.writeRowsByHeader_(cr2, [{ Show: true, Rule: 'Siege', Method: 'most points', Counts: '*', 'Window Days': 7 }, { Show: true, Rule: 'Mine', Method: 'most points', Counts: 'lunch' }]);
  const mz2 = w2.book.getSheetByName('Map Zones');
  w2.writeRowsByHeader_(mz2, J(w2.sheetRows_('Map Zones')).map(r => Object.assign(r, { 'Conquer By': 'Siege', Multiplier: '', Holder: 'GOLD', Since: '2026-11-01' })));
  let res = J(w2.applyMonthlyRules_(false));
  assert.strictEqual(res.zones, 12);
  let z2 = J(w2.sheetRows_('Map Zones'));
  assert.strictEqual(z2.find(r => r.Zone === '90s Hall')['Conquer By'], 'Fewest Interventions');
  assert.strictEqual(z2.find(r => r.Zone === 'Auditorium').Multiplier, 2);
  assert.strictEqual(z2.find(r => r.Zone === '90s Hall').Holder, 'GOLD', 'No = holders kept');
  assert.ok(J(w2.sheetRows_('Conquest Rules')).some(r => r.Rule === 'Mine'), 'a rule of your own stays');
  assert.strictEqual(J(w2.sheetRows_('Conquest Rules')).find(r => r.Rule === 'Siege')['Window Days'], 14, 'a same-named rule is rewritten');
  w2.applyMonthlyRules_(true);
  assert.ok(J(w2.sheetRows_('Map Zones')).every(r => r.Holder === ''), 'Yes = a clean start');
  assert.ok(w2.book.getSheetByName('How The Map Works').rows.some(r => /SQUAD MONTHLY/.test(r[0])));
  assert.strictEqual(w.ymOf_('December 2026'), '2026-12');
  assert.strictEqual(w.ymOf_('2026-1'), '2026-01');
});

test('squad event check-ins: students check in, Generals confirm / remove / add, only confirmed count; Gym 1 = participants + Move/Make/Talk points', () => {
  const w = world();
  w.ensureSettings_(); w.ensureGeneralsTabs_(); w.ensureMayhemWeekTabs_(); w.ensureCheckinTabs_(); w.ensureGameTabs_();
  const today = w.todayYmd_(), lm = w.monthBounds_(-1), mid = lm[0].slice(0, 8) + '15';
  const evs = w.book.getSheetByName('Sign-up Events');
  assert.ok(evs.rows[0].includes('Kind'), 'Sign-up Events gains Kind');
  w.writeRowsByHeader_(evs, [
    { Event: 'Dodgeball', Date: today, Open: true, Kind: 'Move' }, { Event: 'Poster Lab', Date: w.ymdShift_(today, -5), Open: true, Kind: 'make' },
    { Event: 'No kind', Date: today, Open: true }, { Event: 'Debate', Date: mid, Kind: 'Talk' }]);
  w.book.insertSheet('Student Board').rows = [['800#', 'Name', 'Squad', 'Grade'], ['8001111', 'Ann L.', 'Gold', '7'], ['8002222', 'Ed P.', 'Gold', '8'], ['8003333', 'Cy R.', 'Grey', '7']];
  let d = JSON.parse(w.clashCheckins(''));
  assert.deepStrictEqual(d.events.map(e => e.event), ['Dodgeball'], 'only a Kind event on its day (CHECKIN_DAYS 1: Poster Lab 5 days ago is closed)');
  assert.throws(() => w.clashEventCheckIn('Poster Lab', '8001111'), /isn't open/);
  assert.throws(() => w.clashEventCheckIn('Dodgeball', '8009999'), /isn't on the list/);
  d = JSON.parse(w.clashEventCheckIn('Dodgeball', '800-1111'));
  assert.strictEqual(d.events[0].status, 'waiting');
  w.clashEventCheckIn('Dodgeball', '8001111'); // twice = once
  w.clashEventCheckIn('Dodgeball', '8002222'); w.clashEventCheckIn('Dodgeball', '8003333');
  const ci = w.book.getSheetByName('Event Check-ins');
  assert.strictEqual(ci.rows.length, 4);
  assert.deepStrictEqual(J(w.eventParticipants_(today, today)), { BLACK: 0, GOLD: 0, GREY: 0, PURPLE: 0 }, 'nothing counts until a General confirms');
  // GOLD's General: sees only GOLD's check-ins, confirms Ann, removes Ed, can't touch GREY's.
  const cache = w.CacheService.getScriptCache();
  cache.put('hq-g', JSON.stringify({ id: '8001111', name: 'Ann L.', squad: 'GOLD' }));
  w.clashGeneralsHq = () => '{}';
  const hq = J(w.hqCheckins_('GOLD'));
  assert.strictEqual(hq.rows.length, 2); assert.strictEqual(hq.waiting, 2);
  assert.deepStrictEqual(hq.events.map(e => e.event).sort(), ['Debate', 'Dodgeball', 'Poster Lab'].filter(e => e !== 'Debate' || mid >= w.ymdShift_(today, -45)).sort());
  const row = n => hq.rows.find(r => r.name === n).row;
  const greyRow = J(w.hqCheckins_('GREY')).rows[0].row;
  assert.throws(() => w.clashGeneralsConfirmCheckin('g', greyRow, true), /isn't on your squad/);
  w.clashGeneralsConfirmCheckin('g', row('Ed P.'), false); // removed: row deleted, so look Ann up again
  w.clashGeneralsConfirmCheckin('g', J(w.hqCheckins_('GOLD')).rows.find(r => r.name === 'Ann L.').row, true);
  assert.strictEqual(J(w.hqCheckins_('GOLD')).rows.length, 1);
  assert.deepStrictEqual(J(w.eventParticipants_(today, today)), { BLACK: 0, GOLD: 1, GREY: 0, PURPLE: 0 });
  assert.strictEqual(JSON.parse(w.clashCheckins('8001111')).events[0].status, 'confirmed');
  // A General adds a member who forgot, confirmed straight away; not someone from another squad.
  assert.throws(() => w.clashGeneralsAddCheckin('g', 'Poster Lab', '8003333'), /isn't on your squad/);
  assert.throws(() => w.clashGeneralsAddCheckin('g', 'No kind', '8002222'), /Move, Make or Talk/);
  w.clashGeneralsAddCheckin('g', 'Poster Lab', '8002222');
  assert.strictEqual(J(w.hqCheckins_('GOLD')).rows.filter(r => r.confirmed).length, 2);
  // Gym 1 last month: GOLD 2 confirmed at Debate + 3 Talk points = 5; PURPLE 4 Move points; GREY 1 unconfirmed = 0.
  ci.rows.push(['Debate', mid, 'Talk', '8001111', 'Ann L.', 'GOLD', '', true, 'Ann L.'], ['Debate', mid, 'Talk', '8002222', 'Ed P.', 'GOLD', '', true, 'Ann L.'],
    ['Debate', mid, 'Talk', '8003333', 'Cy R.', 'GREY', '', false, '']);
  const ev = w.book.getSheetByName('events');
  ev.rows.push([801, 'Debate', 'squad_talk', mid, '', '', 0, 3, 0, 0], [802, 'Relay', 'squad_move', mid, '', '', 0, 0, 0, 4], [803, 'Kickball', 'event', mid, '', '', 0, 0, 50, 0]);
  const gym1 = J(w.mapZones_(null, false)).find(z => z.zone === 'Gym 1');
  assert.deepStrictEqual(gym1.score, { BLACK: 0, GOLD: 5, GREY: 0, PURPLE: 4 }, 'a plain "event" row is not a squad event');
  assert.strictEqual(gym1.holder, 'GOLD');
  // The HQ payload carries the check-ins for the Generals tab.
  assert.ok(/hqCheckins_/.test(fs.readFileSync(path.join(dir, '05_Generals.gs'), 'utf8')));
});

test('staff overrides: Lock holder keeps a zone; Zone Overrides replace a squad\'s number (newest wins, blank clears); Map editor round-trip', () => {
  const w = world();
  w.ensureSettings_(); w.ensureGameTabs_();
  const lm = w.monthBounds_(-1), ym = lm[0].slice(0, 7), mid = lm[0].slice(0, 8) + '15';
  const ev = w.book.getSheetByName('events');
  ev.rows.push([901, 'Big day', 'event', mid, '', '', 0, 50, 10, 0]); // Music Wing (Full Climb): GOLD leads
  let z = n => J(w.mapZones_(null, true)).find(x => x.zone === n);
  assert.strictEqual(z('Music Wing').holder, 'GOLD');
  // Override: GREY's number for last month becomes 80, so GREY takes it.
  const ov = w.book.getSheetByName('Zone Overrides');
  assert.deepStrictEqual(ov.rows[0], ['Month', 'Zone', 'Squad', 'Score', 'Note', 'Set By', 'Set At']);
  ov.rows.push([ym, 'Music Wing', 'GREY', 80, 'missed spirit points', 'Mr. S', '']);
  let mw = z('Music Wing');
  assert.strictEqual(mw.holder, 'GREY');
  assert.deepStrictEqual([mw.score.GREY, mw.computed.GREY, J(mw.overridden)], [80, 10, ['GREY']]);
  // A newer blank row clears it; an override for another month does nothing.
  ov.rows.push([ym, 'music wing', 'Grey', '', '', 'Mr. S', ''], ['2020-01', 'Music Wing', 'GOLD', 999, '', '', '']);
  mw = z('Music Wing');
  assert.strictEqual(mw.score.GREY, 10);
  assert.strictEqual(mw.holder, 'GOLD', 'override cleared: the data decides again (GOLD 55 beats GREY 10)');
  // Lock: staff hand Music Wing to PURPLE and lock it; the data can't move it.
  const mz = w.book.getSheetByName('Map Zones');
  w.writeRowsByHeader_(mz, J(w.sheetRows_('Map Zones')).map(r => Object.assign(r, r.Zone === 'Music Wing' ? { Holder: 'PURPLE', Locked: true } : {})));
  assert.strictEqual(z('Music Wing').holder, 'PURPLE', 'locked');
  // Map editor: shows each zone's numbers, saves Lock and overrides, re-judges right away.
  w.CacheService.getScriptCache().put('clash-admin-T', '1');
  const loaded = JSON.parse(w.clashMapLoad('T'));
  const mwz = loaded.zones.find(x => x.zone === 'Music Wing');
  assert.strictEqual(mwz.locked, true);
  assert.deepStrictEqual([loaded.scores['Music Wing'].month, loaded.scores['Music Wing'].computed.GOLD], [ym, 55]);
  mwz.locked = false;
  loaded.overrides = [{ zone: 'Music Wing', month: ym, squad: 'BLACK', value: '200' }];
  const r = JSON.parse(w.clashMapSave('T', JSON.stringify(loaded)));
  assert.strictEqual(r.overrides, 1);
  assert.strictEqual(J(w.sheetRows_('Map Zones')).find(x => x.Zone === 'Music Wing').Holder, 'BLACK', 'unlocked + BLACK override 200: re-judged on save');
  assert.strictEqual(JSON.parse(w.clashMapLoad('T')).scores['Music Wing'].overrides.BLACK, 200);
  loaded.overrides = [{ zone: 'Music Wing', month: ym, squad: 'BLACK', value: 'lots' }];
  assert.throws(() => w.clashMapSave('T', JSON.stringify(loaded)), /must be a number/);
});

test('evolution: LRG point stages (4), three versions, a boost per evolution (different columns), weekly boost points, perks, gear overlays', () => {
  const w = world();
  w.ensureSettings_(); w.ensureGameTabs_(); w.ensureMeTabs_(); w.ensurePollTabs_();
  const tok = login(w, 'oldpw');
  // Made-up student. LRG points = bills 4 + summit 5 + STAR 2x1 + streak 2x2 = 15: stage 1.
  const board = w.book.insertSheet('Student Board');
  board.rows = [['800#', 'Name', 'Squad', 'Grade', 'Trait Dough Bills', 'Bills This Week', 'Trait Dough Wins', 'Mindset Summits', 'Trail Journal Growth', 'STAR Growth Points', 'Shout-outs', 'Ambassador Buddies', 'Streak (weeks)', 'Badges', 'Last Contribution', 'Updated', 'Attendance %', 'Email'],
    ['8001111', 'Ann L.', 'Gold', '7', '4', '0', '0', '1', '', '2', '0', '0', '2', '', '', '', '96', 'ann.l@easdpa.org']];
  const set = (col, v) => { board.rows[1][board.rows[0].indexOf(col)] = v; };
  let c = JSON.parse(w.clashMe('8001111', tok)).creature;
  assert.strictEqual(c.xp, 15); assert.strictEqual(c.stage, 1);
  assert.deepStrictEqual(c.next, { stage: 2, name: 'Ridge', need: 15, at: 30 });
  assert.strictEqual(c.art.length, 4);
  assert.ok(/creatures\/gold-1\.jpg$/.test(c.art[0].v[1]) && /gold-1f\.jpg$/.test(c.art[0].v[2]) && /gold-1w\.jpg$/.test(c.art[0].v[3]));
  assert.deepStrictEqual(c.versions, { 1: 'Version 1', 2: 'Version 2', 3: 'Version 3' });
  assert.deepStrictEqual(c.choices, {}, 'no boosts before stage 2');
  assert.throws(() => w.clashMePickBoost(2, 'Trait Dough', '8001111', tok), /hasn't reached/);
  // Stage 2 at 30: a x1.25 boost. Columns not on the Student Board are not offered.
  set('Trait Dough Bills', '20');
  c = JSON.parse(w.clashMe('8001111', tok)).creature;
  assert.strictEqual(c.stage, 2);
  const offered = c.choices[2].map(b => b.boost);
  assert.ok(offered.includes('Trait Dough') && offered.includes('Poll Turnout') && offered.includes('Attendance'));
  assert.ok(!offered.includes('Behavior') && !offered.includes('Student of the Month'), 'no Student Board column yet = not offered');
  assert.strictEqual(c.choices[2].find(b => b.boost === 'Trait Dough').mult, 1.25);
  c = JSON.parse(w.clashMePickBoost(2, 'Trait Dough', '8001111', tok)).creature;
  assert.deepStrictEqual(c.boosts.map(b => [b.stage, b.boost, b.mult]), [[2, 'Trait Dough', 1.25]]);
  assert.throws(() => w.clashMePickBoost(2, 'Attendance', '8001111', tok), /already picked/);
  // Stage 3 (High Camp) at 60: a different boost at x1.5.
  set('Trait Dough Bills', '50');
  c = JSON.parse(w.clashMe('8001111', tok)).creature;
  assert.strictEqual(c.stage, 3); assert.strictEqual(c.next.name, 'Summit');
  assert.ok(!c.choices[3].some(b => b.boost === 'Trait Dough'));
  assert.strictEqual(c.choices[3].find(b => b.boost === 'Attendance').mult, 1.5);
  w.clashMePickBoost(3, 'Attendance', '8001111', tok);
  // Stage 4 (Summit) at 100: a third, x2.
  set('Trait Dough Bills', '90');
  c = JSON.parse(w.clashMe('8001111', tok)).creature;
  assert.strictEqual(c.stage, 4); assert.strictEqual(c.next, null);
  assert.ok(!c.choices[4].some(b => b.boost === 'Trait Dough' || b.boost === 'Attendance'));
  c = JSON.parse(w.clashMePickBoost(4, 'Mindset', '8001111', tok)).creature;
  assert.deepStrictEqual(c.boosts.map(b => b.mult), [1.25, 1.5, 2]);
  // Weekly payout: Trait Dough 70 new bills x 0.25 = 17.5, Attendance 96% pays 4 x 0.5 = 2, Mindset no new summits: 19.5 -> 20 for GOLD.
  const ev = w.book.getSheetByName('events');
  w.props['boost-paid-week'] = 'old';
  w.payBoosts_();
  const row = ev.rows.find(r => r[2] === 'evolution_boost');
  assert.ok(row, 'one Evolution boosts row');
  assert.strictEqual(row[7], 20); assert.ok(!/Ann/.test(row.join(' ')), 'no student names on the points row');
  w.payBoosts_();
  assert.strictEqual(ev.rows.filter(r => r[2] === 'evolution_boost').length, 1, 'once a week');
  w.props['boost-paid-week'] = 'old2';
  w.payBoosts_();
  assert.strictEqual(ev.rows.filter(r => r[2] === 'evolution_boost')[1][7], 2, 'next week: attendance pays again, Trait Dough only on new bills');
  assert.strictEqual(J(w.pointsBy_(w.readEventsTab_(), '2000-01-01', '2999-12-31', e => e.type === 'evolution_boost')).points.GOLD, 0, 'boost points never count toward map claims');
  // Version 2 or 3, and show an earlier stage.
  c = JSON.parse(w.clashMeSaveCreature({ version: 3, show: 1 }, '8001111', tok)).creature;
  assert.strictEqual(c.version, 3); assert.strictEqual(c.show, 1);
  // Gear overlays in the squad's color: wear up to 3 earned pieces.
  const g = n => c.gear.find(x => x.gear === n);
  assert.ok(/gear\/gold-boots\.jpg$/.test(g('boots').image));
  ['boots', 'pack', 'rope'].forEach(n => { c = JSON.parse(w.clashMeSaveCreature({ gear: n }, '8001111', tok)).creature; });
  assert.deepStrictEqual(c.gear.filter(x => x.worn).map(x => x.gear), ['boots', 'pack', 'rope']);
  assert.throws(() => w.clashMeSaveCreature({ gear: 'compass' }, '8001111', tok), /Take one off/);
  assert.throws(() => w.clashMeSaveCreature({ gear: 'lantern' }, '8001111', tok), /haven't earned/);
  c = JSON.parse(w.clashMeSaveCreature({ gear: 'pack' }, '8001111', tok)).creature;
  assert.deepStrictEqual(c.gear.filter(x => x.worn).map(x => x.gear), ['boots', 'rope'], 'tap again to take it off');
  // Perks: equip an unlocked one, tap again to take it off; locked ones refuse.
  c = JSON.parse(w.clashMeSaveCreature({ perk: 'Gold Rim' }, '8001111', tok)).creature;
  assert.strictEqual(c.perks.equipped.frame, 'Gold Rim');
  c = JSON.parse(w.clashMeSaveCreature({ perk: 'Gold Rim' }, '8001111', tok)).creature;
  assert.strictEqual(c.perks.equipped.frame, undefined);
  assert.throws(() => w.clashMeSaveCreature({ perk: 'Iron Streak' }, '8001111', tok), /haven't unlocked/);
  assert.ok(c.perks.list.find(p => p.perk === 'Peak Legend').open);
  // The old look saver keeps the creature's choices; an old 'f' look shows Version 2.
  w.clashMeSaveLook({ fur: 1 }, '8001111', tok);
  assert.strictEqual(JSON.parse(w.clashMe('8001111', tok)).creature.version, 3);
  assert.strictEqual(w.versionOf_({ form: 'f' }), 2);
  // Cliff is on every page.
  assert.ok(/cliff\.png$/.test(JSON.parse(w.clashMe('8001111', tok)).cliff.image));
});

test("Cliff's tour: steps per page in order from the Tutorial Steps tab; Opponent art and defeated art", () => {
  const w = world();
  w.ensureSettings_(); w.ensureGameTabs_(); w.ensureTourTab_();
  let t = JSON.parse(w.clashTour('site'));
  assert.ok(t.steps.length >= 5 && t.auto && t.version === '1' && /cliff\.png$/.test(t.cliff.image));
  assert.strictEqual(t.steps[0].title, "Hi, I'm Cliff!");
  assert.ok(t.steps.some(s => s.spot === 'map rules'));
  const me = J(w.tourFor_('me'));
  assert.ok(me.steps.some(s => s.go === 'tickets' && s.spot === 'tickets'));
  // Staff edit the tab: order, hide a step, turn auto off.
  const tab = w.book.getSheetByName('Tutorial Steps');
  tab.rows.find(r => r[1] === 'site' && r[2] === 2)[0] = false;
  tab.rows.push([true, 'site', 0, 'home', '', 'First!', 'Added by staff.']);
  w.setSettingValue_('TOUR_AUTO', 'OFF');
  t = JSON.parse(w.clashTour('site'));
  assert.strictEqual(t.steps[0].title, 'First!'); assert.ok(!t.steps.some(s => s.title === 'Every point counts')); assert.strictEqual(t.auto, false);
  assert.strictEqual(JSON.parse(w.clashTour('anything')).page, 'site');
  // Opponents: built-in art by name, a defeated version once beaten; a row's Defeated Image wins.
  assert.ok(/monster-tardy-troll\.jpg$/.test(w.opponentArt_('The Tardy Troll', false)));
  assert.ok(/monster-litter-beast-defeated\.jpg$/.test(w.opponentArt_('The Hallway Litter Beast', true)));
  assert.ok(/monster-shady-defeated\.jpg$/.test(w.opponentArt_('Shady', true)));
  assert.strictEqual(w.opponentArt_('Someone New', true), '');
  const opp = w.book.getSheetByName('Opponents');
  assert.ok(opp.rows[0].includes('Defeated Image'));
  const row = opp.rows.find(r => r[3] === 'The Tardy Troll');
  const col = h => opp.rows[0].indexOf(h);
  row[col('Show')] = true; row[col('Status')] = 'Defeated 2026-11-05';
  let m = J(w.monster_([]));
  assert.ok(m.defeated && /monster-tardy-troll-defeated\.jpg$/.test(m.image));
  row[col('Defeated Image')] = 'https://example.org/troll-beaten.png';
  m = J(w.monster_([]));
  assert.strictEqual(m.image, 'https://example.org/troll-beaten.png');
});

test('speed + leaderboards: one Settings read per request, cached public data, program apps, squad boards, Top Climbers only when on, Trail Journal LRG', () => {
  const w = world();
  w.ensureSettings_(); w.ensureGameTabs_(); w.ensureMeTabs_();
  // Settings are read once, and a write clears the memo.
  const st = w.book.getSheetByName('Settings');
  let reads = 0; const orig = st.getDataRange.bind(st);
  st.getDataRange = () => { reads++; return orig(); };
  w.SETTINGS_MEMO_ = null;
  for (let i = 0; i < 50; i++) w.getSetting_('TV_SHOW_NAMES');
  assert.ok(reads <= 1, 'Settings tab read ' + reads + ' times for 50 lookups');
  w.setSettingValue_('POLL_SECONDS', '9');
  assert.strictEqual(w.getSetting_('POLL_SECONDS'), '9', 'a write is seen right away');
  // The cache splits big values and puts them back together.
  const store = {};
  const c = { put: (k, v) => { store[k] = v; }, remove: (k) => { delete store[k]; }, putAll: (o) => Object.assign(store, o),
    getAll: (ks) => { const o = {}; ks.forEach(k => { if (store[k] != null) o[k] = store[k]; }); return o; }, get: (k) => store[k] == null ? null : store[k] };
  w.CacheService.getScriptCache = () => c;
  const big = 'x'.repeat(250000) + 'end';
  w.cachePutBig_('t', big, 60);
  assert.strictEqual(store.t, '3'); assert.strictEqual(w.cacheGetBig_('t'), big);
  delete store['t:1']; assert.strictEqual(w.cacheGetBig_('t'), null, 'a missing piece = a miss, never half the data');
  // The public data: cached, and it carries the program apps.
  const first = JSON.parse(w.cachedPublicLoad_());
  assert.ok(/AKfycbz1P6NE/.test(first.value.apps.traitdough) && /AKfycbzt9HZ0/.test(first.value.apps.sotm) && /AKfycbwuQCz/.test(first.value.apps.mindset));
  assert.strictEqual(first.value.apps.embed, true);
  assert.ok(Object.keys(store).some(k => /^pub-load-/.test(k)), 'stored in the cache');
  assert.strictEqual(w.clashLoad(''), w.cachedPublicLoad_(), 'visitors get the cached copy');
  // Leaderboards: squads only until names are turned on.
  const board = w.book.insertSheet('Student Board');
  board.rows = [['800#', 'Name', 'Squad', 'Grade', 'Trait Dough Bills', 'Mindset Summits', 'STAR Growth Points', 'Shout-outs', 'Ambassador Buddies', 'Streak (weeks)', 'Trail Journal Growth', 'Last Contribution'],
    ['8001111', 'Ann L.', 'Gold', '7', '40', '1', '2', '0', '0', '2', '3', w.todayYmd_()],
    ['8002222', 'Ed Parker', 'Grey', '7', '5', '0', '0', '1', '0', '0', '', ''],
    ['8003333', 'Cy R.', 'Gold', '8', '1', '0', '0', '0', '0', '0', '', '']];
  let d = J(w.leadersData_());
  const b = k => d.boards.find(x => x.key === k);
  assert.strictEqual(b('bills').scores.GOLD, 41); assert.strictEqual(b('bills').scores.GREY, 5);
  assert.strictEqual(b('journal').scores.GOLD, 3);
  assert.strictEqual(b('pitched').scores.GOLD, 50, '1 of 2 Gold students pitched in lately');
  assert.ok(!d.names && !d.top, 'no names unless LEADERBOARD_NAMES is ON');
  assert.ok(!JSON.stringify(d).includes('Ann'), 'squad boards carry no student names');
  w.setSettingValue_('LEADERBOARD_NAMES', 'ON');
  d = J(w.leadersData_());
  assert.deepStrictEqual(d.top['7'].map(p => p.name), ['Ann L.', 'Ed P.'], 'ranked by LRG points, First L. only');
  assert.ok(d.top['7'][0].lrg > d.top['7'][1].lrg);
  // Trail Journal growth counts toward LRG points (2 each).
  assert.strictEqual(w.climbXp_({ journal: 3 }), 6);
});

test('the week: marking periods + beat-our-own-MP bonus, This Week data + up next, the digest email (once a week, no names), announcement, usage counts, health panel, meeting slides', () => {
  const w = world();
  w.ensureSettings_(); w.ensureGameTabs_(); w.ensureMeTabs_(); w.ensurePollTabs_(); w.ensureWeekTabs_(); w.ensureGeneralsTabs_(); w.ensureFreezeTab_();
  const today = w.todayYmd_(), ago = n => w.ymdShift_(today, -n), ahead = n => w.ymdShift_(today, n);
  const sent = [];
  w.MailApp = { sendEmail: o => sent.push(o), getRemainingDailyQuota: () => 90 };
  w.Session.getEffectiveUser = () => ({ getEmail: () => 'staff@easdpa.org' });
  w.ScriptApp = { getProjectTriggers: () => [{ getHandlerFunction: () => 'backgroundTick' }], deleteTrigger() {}, newTrigger: () => ({ timeBased: () => ({ everyMinutes: () => ({ create() {} }), onWeekDay: () => ({ atHour: () => ({ create() {} }) }) }) }), WeekDay: { FRIDAY: 5 } };
  // The district's marking periods are in Settings.
  const mps = J(w.markingPeriods_());
  assert.deepStrictEqual(mps.map(p => p.name), ['MP1', 'MP2', 'MP3', 'MP4']);
  assert.deepStrictEqual([mps[0].start, mps[0].end, mps[3].end], ['2026-08-31', '2026-10-30', '2027-06-09']);
  // Beat our own MP: a squad that out-earns its last marking period gets the bonus, once.
  w.setSettingValue_('MARKING_PERIODS', 'A=' + ago(40) + '..' + ago(21) + '; B=' + ago(20) + '..' + ago(1) + '; C=' + today + '..' + ahead(30));
  const ev = w.book.getSheetByName('events');
  ev.rows.push([501, 'Spirit', 'event', ago(30), '', '', 10, 10, 10, 10], [502, 'Trivia', 'event', ago(10), '', '', 0, 25, 5, 0]);
  w.payBeatMp_(w.readEventsTab_());
  const bonus = ev.rows.filter(r => r[2] === 'mp_bonus');
  assert.strictEqual(bonus.length, 1); assert.strictEqual(bonus[0][7], 50, 'GOLD 25 > 10'); assert.strictEqual(bonus[0][6], '', 'BLACK 0 < 10');
  w.payBeatMp_(w.readEventsTab_());
  assert.strictEqual(ev.rows.filter(r => r[2] === 'mp_bonus').length, 1, 'once per marking period');
  const mt = J(w.meetingData_('GOLD'));
  assert.strictEqual(mt.mp.name, 'C'); assert.strictEqual(mt.mp.prevName, 'B'); assert.strictEqual(mt.mp.prevPoints, 25);
  // This Week: the week's points, highlights, and next week's dated items.
  ev.rows.push([503, 'Kickoff Assembly', 'event', today, '', '', 5, 0, 0, 40]);
  const nextMon = w.ymdShift_(w.weekStart_(today), 7);
  w.book.getSheetByName('Sign-up Events').rows.push(['Trivia Bowl', w.ymdShift_(nextMon, 1), 4, true]);
  w.book.getSheetByName('Round Goals').rows.push([true, ago(5), w.ymdShift_(nextMon, 3), 'Most poll voters', 'voters', 8, 5, 3, '', '']);
  const d = J(w.weekData_());
  assert.strictEqual(d.week.PURPLE, 40); assert.strictEqual(d.highlights[0].name, 'Kickoff Assembly');
  assert.ok(d.next.some(n => n.what === 'Trivia Bowl') && d.next.some(n => /Round goal ends/.test(n.what)), JSON.stringify(d.next));
  // The digest email: HTML, the sections, nobody's name, once a week.
  assert.strictEqual(w.weeklyDigest(false), 'No recipients: add emails to Settings › DIGEST_EMAILS.');
  w.setSettingValue_('DIGEST_EMAILS', 'a@easdpa.org, b@easdpa.org, nope');
  assert.strictEqual(w.weeklyDigest(false), 'Sent to 2.');
  assert.strictEqual(sent.length, 1);
  assert.strictEqual(sent[0].to, 'a@easdpa.org,b@easdpa.org');
  ['This Week in Clash', 'Standings', 'Highlights', 'Up next', 'Trivia Bowl', 'Open Clash of Classes'].forEach(t => assert.ok(sent[0].htmlBody.includes(t), t));
  assert.strictEqual(w.weeklyDigest(false), 'Already sent this week.');
  assert.strictEqual(sent.length, 1);
  assert.ok(w.book.getSheetByName('Weekly Digest Log').rows.length === 2, 'logged');
  // The morning announcement.
  const ann = w.announcementScript_();
  assert.ok(/^Good morning, Mountaineers!/.test(ann) && / leads Clash of Classes/.test(ann), ann);
  // Usage: anonymous counts per page, unknown pages ignored, written to the Usage tab.
  const store = {};
  const c = { get: k => store[k] == null ? null : store[k], put: (k, v) => { store[k] = v; }, remove: k => { delete store[k]; }, putAll: o => Object.assign(store, o),
    getAll: ks => { const o = {}; ks.forEach(k => { if (store[k] != null) o[k] = store[k]; }); return o; } };
  w.CacheService.getScriptCache = () => c;
  w.clashTrack('home'); w.clashTrack('home'); w.clashTrack('leaders'); w.clashTrack('<script>');
  w.flushUsage_();
  const usage = w.book.getSheetByName('Usage').rows.slice(1);
  assert.deepStrictEqual(usage.map(r => [r[1], r[2]]).sort(), [['home', 2], ['leaders', 1]]);
  w.clashTrack('home'); w.flushUsage_();
  assert.strictEqual(w.book.getSheetByName('Usage').rows.find(r => r[1] === 'home')[2], 3, 'today\'s row is updated, not duplicated');
  // Health: admins only; shows jobs, missing settings and usage.
  assert.throws(() => w.clashHealth('nope'));
  c.put('clash-admin-T', '1');
  w.noteRun_('backgroundTick');
  const h = JSON.parse(w.clashHealth('T'));
  assert.ok(h.jobs.find(j => j.job === 'backgroundTick').ok);
  assert.ok(h.missing.some(m => /STUDENT_LINK/.test(m)));
  assert.ok(h.usage.find(u => u.page === 'home').opens === 3);
});

test('roster: the Students tab copies the district roster, keeps squads, and lets students sign in by 800#@gomounts.org', () => {
  const w = world();
  w.ensureSettings_(); w.ensureGameTabs_(); w.ensureMeTabs_();
  w.setSettingValue_('ROSTER_SHEET', 'https://docs.google.com/spreadsheets/d/ROSTER_SAMPLE_ID_0000000000/edit');
  // Made-up students in the district roster's layout.
  const roster = makeSheet('Roster', [['800 Number', 'Email Address', 'First Name', 'Last Name', 'Grade Level'],
    ['8090001', '8090001@gomounts.org', 'Ada', 'Lark', '7'], ['8090002', '', 'Ben', 'Moss', '8'], ['8090003', '8090003@gomounts.org', 'Cy', 'Pine', '8']]);
  w.SpreadsheetApp.openById = id => { assert.strictEqual(id, 'ROSTER_SAMPLE_ID_0000000000'); return { getSheets: () => [roster] }; };
  w.book.insertSheet('Student Board').rows = [['800#', 'Name', 'Squad', 'Grade', 'Trait Dough Bills'], ['8090003', 'Cy P.', 'Grey', '8', '4']];
  let r = J(w.importRoster_());
  assert.deepStrictEqual([r.students, r.squads, r.missingSquad], [3, 1, 2], 'Cy gets GREY from the Student Board');
  const tab = w.book.getSheetByName('Students');
  assert.strictEqual(tab.rows[2][4], '8090002@gomounts.org', 'a missing email is built from the 800#');
  tab.rows[1][5] = 'GOLD'; // staff type Ada's squad
  r = J(w.importRoster_());
  assert.strictEqual(tab.rows[1][5], 'GOLD', 'typed squads survive a re-import');
  // Ada isn't on the Student Board yet: My Squad still works from the Students tab.
  const page = JSON.parse(w.clashMyPage('8090001'));
  assert.deepStrictEqual([page.me.name, page.me.squad, page.me.bills], ['Ada L.', 'GOLD', 0]);
  assert.throws(() => w.clashMyPage('8090002'), /squad isn't set yet/);
  // My Climb sign-in: the school email is the 800#.
  w.setSettingValue_('MY_CLIMB_HOURS', '00:00-23:59'); w.setSettingValue_('MY_CLIMB_SCHOOL_DAYS', 'OFF');
  w.Session.getActiveUser = () => ({ getEmail: () => '8090001@gomounts.org' });
  assert.strictEqual(w.signedInId_(), '8090001');
  assert.strictEqual(JSON.parse(w.clashMe()).me.squad, 'GOLD');
});

test('page files: scripts parse, no old token or Supabase key left', () => {
  for (const f of ['App.html', 'Widgets.html', 'Index.html', 'Student.html', 'Generals.html', 'Pop.html']) {
    const h = fs.readFileSync(path.join(dir, f), 'utf8');
    assert.ok(!/togetherweclimb4444HOPE|supabase\.co|TJ_ANON|ephrata2026/.test(h), f + ' has no secrets');
    [...h.matchAll(/<script[^>]*>([\s\S]*?)<\/script>/g)].forEach(m => new vm.Script(m[1].replace(/<\?!=[\s\S]*?\?>/g, '{}')));
  }
});

test('from the drive: Mindset Links tab fills the Mindset page (site links win), Trait Dough leaderboard, Generals import onto MLC Roster', () => {
  const w = world();
  w.ensureSettings_(); w.ensureGeneralsTabs_();
  // Mindset links: the seeded tab fills every month and grade.
  w.ensureMindsetLinksTab_();
  let v = { links: { mindset: [] } };
  w.mergeMindsetLinks_(v);
  assert.strictEqual(v.links.mindset.length, 27, '9 months x 7th, 8th, Extension');
  assert.ok(v.links.mindset.every(l => /^https:\/\//.test(l.url) && l.fromSheet));
  assert.ok(/AKfycbwuQCz/.test(v.links.mindset.find(l => l.key === '8th-January').url), '8th grade = the Mindset app');
  assert.ok(v.links.mindset.find(l => l.key === 'ext-February').url.includes('1yVoCQfyx4V4'));
  // A link saved on the site wins for that month; a stale copy of a tab link is replaced by the tab.
  const tab = w.book.getSheetByName('Mindset Links');
  tab.rows[1][2] = 'https://docs.google.com/document/d/NEW8THSEPT/edit';
  v = { links: { mindset: [{ key: '7th-September', url: 'https://example.org/mine', label: 'x' }, { key: '8th-September', url: 'https://old', fromSheet: true }] } };
  w.mergeMindsetLinks_(v);
  const k = key => v.links.mindset.filter(l => l.key === key);
  assert.deepStrictEqual(k('7th-September').map(l => l.url), ['https://example.org/mine']);
  assert.deepStrictEqual(k('8th-September').map(l => l.url), ['https://docs.google.com/document/d/NEW8THSEPT/edit']);
  // The public load carries them.
  assert.strictEqual(JSON.parse(JSON.stringify(w.publicLoad_())).value.links.mindset.length, 27);
  // Trait Dough leaderboard: squads only until names are on.
  const board = w.book.insertSheet('Student Board');
  board.rows = [['800#', 'Name', 'Squad', 'Grade', 'Trait Dough Bills', 'Bills This Week', 'Trait Dough Wins'],
    ['8001111', 'Ann L.', 'Gold', '7', '10', '2', '1'], ['8002222', 'Cy R.', 'Gold', '8', '0', '0', '0'], ['8003333', 'Ed Parker', 'Grey', '7', '4', '4', '0']];
  let d = J(w.traitDoughBoard_());
  const b = key => d.boards.find(x => x.key === key).scores;
  assert.strictEqual(b('year').GOLD, 10); assert.strictEqual(b('week').GREY, 4); assert.strictEqual(b('wins').GOLD, 1);
  assert.strictEqual(b('per').GOLD, 5); assert.strictEqual(b('share').GOLD, 50); assert.strictEqual(d.total, 14);
  assert.ok(!d.top && !JSON.stringify(d).includes('Ann'), 'no names unless LEADERBOARD_NAMES is ON');
  w.setSettingValue_('LEADERBOARD_NAMES', 'ON');
  d = J(w.traitDoughBoard_());
  assert.deepStrictEqual(d.top['7'].map(p => p.name + ' ' + p.bills), ['Ann L. 10', 'Ed P. 4']);
  assert.deepStrictEqual(d.top['8'], [], 'zero bills are left off');
  // Generals: names from the Squad Leadership doc, 800# from the Students tab.
  const para = t => ({ getType: () => 'PARAGRAPH', asText: () => ({ getText: () => t }) });
  const li = t => ({ getType: () => 'LIST_ITEM', asText: () => ({ getText: () => t }) });
  const cell = kids => ({ getText: () => kids.map(x => x.asText().getText()).join('\n'), getNumChildren: () => kids.length, getChild: i => kids[i] });
  const row = cells => ({ getNumCells: () => cells.length, getCell: i => cells[i] });
  const heads = ['Black Squad', 'Gold Squad', 'Grey Squad', 'Purple Squad'].map(h => cell([para(h)]));
  const table = { getNumRows: () => 2, getRow: i => i === 0 ? row(heads) : row([
    cell([para('Ann Lopez'), li('Thomas'), li('*Band*'), para('Bo Kim'), li('Trettin')]), cell([para('Cy Ross'), li('Martin')]), cell([]), cell([para('Dee Fox'), li('Segura')])]) };
  w.DocumentApp = { openById: () => ({ getBody: () => ({ getTables: () => [table] }) }), ElementType: { LIST_ITEM: 'LIST_ITEM' } };
  const found = J(w.generalsFromDoc_('DOC'));
  assert.deepStrictEqual(found.map(g => g.name + '/' + g.squad + '/' + g.homeroom), ['Ann Lopez/BLACK/Thomas', 'Bo Kim/BLACK/Trettin', 'Cy Ross/GOLD/Martin', 'Dee Fox/PURPLE/Segura']);
  assert.deepStrictEqual(found[0].notes, ['Band']);
  const st = w.book.insertSheet('Students');
  st.rows = [['800#', 'First Name', 'Last Name', 'Grade', 'Email', 'Squad'], ['8009001', 'Ann', 'Lopez', '8', '', 'BLACK'], ['8009002', 'Cy', 'Ross', '8', '', 'GOLD']];
  const mlc = w.book.getSheetByName('MLC Roster');
  mlc.rows.push(['8009002', 'Cy Ross', 'GOLD', 'General', '']);
  const r = J(w.importGenerals_());
  assert.strictEqual(r.found, 4); assert.strictEqual(r.added, 3, 'Cy was already on MLC Roster');
  assert.deepStrictEqual(r.missing, ['Bo Kim (Black)', 'Dee Fox (Purple)']);
  assert.deepStrictEqual(mlc.rows.slice(1).map(x => x[0] + ' ' + x[1] + ' ' + x[3]), ['8009002 Cy Ross General', '8009001 Ann Lopez General', ' Bo Kim General', ' Dee Fox General']);
  assert.strictEqual(J(w.importGenerals_()).added, 0, 'running it again adds nobody twice');
});

test('Year Plan: seeded month by month from the Monthly Themes docs, hidden rows left off, links checked, next week\'s items in This Week', () => {
  const w = world();
  w.ensureSettings_(); w.ensureGeneralsTabs_(); w.ensureGameTabs_(); w.ensureMeTabs_(); w.ensurePollTabs_(); w.ensureWeekTabs_();
  w.ensureYearPlanTab_();
  const tab = w.book.getSheetByName('Year Plan');
  assert.ok(tab.rows.length > 60, 'the whole year is seeded');
  assert.ok(tab.rows.slice(1).every(r => r.length === 9), 'every seed row fills all 9 columns');
  const d = JSON.parse(w.clashYearPlan());
  assert.deepStrictEqual(d.months.map(m => m.month), ['September', 'October', 'November', 'December', 'January', 'February', 'March', 'April', 'May']);
  const sept = d.months[0].items;
  assert.ok(sept.some(i => i.kind === 'Mindset Lesson' && i.who === '8th' && i.links.length === 2));
  assert.ok(d.months.every(m => m.items.every(i => i.links.every(u => /^https?:\/\//.test(u)))));
  // Staff add a row as the year goes, hide another, and give one a bad link.
  const today = w.todayYmd_(), nextMon = w.ymdShift_(w.weekStart_(today), 7);
  tab.rows.push([true, '', w.ymdShift_(nextMon, 2), 'Color Squad Activity', 'Snowball Relay', 'All', 'javascript:alert(1)', '', '']);
  tab.rows[1][0] = false;
  w.CacheService.getScriptCache = () => ({ get: () => null, put() {}, remove() {}, putAll() {}, getAll: () => ({}) });
  const items = J(w.yearPlanItems_());
  const relay = items.find(i => i.title === 'Snowball Relay');
  assert.ok(relay && relay.month, 'month filled in from the date');
  assert.deepStrictEqual(relay.links, [], 'only http(s) links');
  assert.ok(!items.some(i => i.title === tab.rows[1][4]), 'unticked Show hides the row');
  w.book.getSheetByName('Sign-up Events') || w.ensureGeneralsTabs_();
  const week = J(w.weekData_());
  assert.ok(week.next.some(n => n.what === 'Snowball Relay' && n.kind === 'Color Squad Activity'), JSON.stringify(week.next));
});

test('Homerooms: form responses become homeroom points (10 x share), waiting for Approved, flagged oddities, once only, standings by room', () => {
  const w = world();
  w.ensureSettings_(); w.ensureHomeroomTabs_();
  const forms = w.book.getSheetByName('Homeroom Forms');
  assert.strictEqual(forms.rows.length, 4, 'three forms seeded');
  const head = ['Timestamp', 'What is your room number? (write the 2-digit number)', 'How many students are wearing a hat? Include yourself', 'How many total students are in your homeroom today? Include yourself in the total!'];
  const resp = [head, [new Date('2026-10-29T12:10:00Z'), '71', '8', '23'], [new Date('2026-10-29T12:11:00Z'), '84', '20', '22'],
    [new Date('2026-10-29T12:12:00Z'), '75', '22', '9'], [new Date('2026-10-29T12:30:00Z'), '71', '9', '23'], ['', '', '', '']];
  const read = () => resp;
  let r = w.importHomeroomResponses_(read);
  const pts = w.book.getSheetByName('Homeroom Points').rows;
  assert.strictEqual(r.added, 4 * 3, 'each seeded form reads the same fake sheet here');
  const first = pts.slice(1).filter(x => x[2] === 'Red Ribbon Week Hat Day');
  assert.deepStrictEqual(first.map(x => [x[3], x[6]]), [['71', 3], ['84', 9], ['75', 10], ['71', 4]]);
  assert.ok(first.every(x => x[0] === false), 'waits for Approved');
  assert.ok(/More took part/.test(first[2][7]) && /Second entry/.test(first[3][7]));
  assert.strictEqual(w.importHomeroomResponses_(read).added, 0, 'never imported twice');
  assert.strictEqual(J(w.homeroomBoard_()).rooms.length, 0, 'nothing counts until approved');
  pts.forEach((x, i) => { if (i && x[2] === 'Red Ribbon Week Hat Day' && !x[7]) x[0] = true; });
  const b = J(w.homeroomBoard_());
  assert.deepStrictEqual(b.rooms.map(x => x.room + ':' + x.total), ['84:9', '71:3']);
  assert.strictEqual(b.latest.event, 'Red Ribbon Week Hat Day');
  assert.strictEqual(w.hrPoints_(5, 0), 0);
});

test('every requireAdmin(…) button target is on the page script export list (Edit Generals was not)', () => {
  const app = fs.readFileSync(path.join(dir, 'App.html'), 'utf8');
  const markup = ['Index.html', 'App.html'].map(f => fs.readFileSync(path.join(dir, f), 'utf8')).join('\n');
  const start = app.indexOf('const _fns = {'), list = app.slice(start, app.indexOf('Object.entries(_fns)', start));
  const targets = [...new Set([...markup.matchAll(/requireAdmin\(([A-Za-z_]\w*)\)/g)].map(m => m[1]))].filter(n => n !== 'fn');
  const missing = targets.filter(n => !new RegExp('[{,\\s]' + n + '[,}\\s]').test(list));
  assert.deepStrictEqual(missing, []);
  assert.ok(/[{,\s]saveGenerals[,}\s]/.test(list));
});

test('My Climb sample student: admin sees a made-up student on any squad (no 800#), every tab has data, nothing is written, saving is refused', () => {
  const w = world();
  w.ensureSettings_(); w.ensureGeneralsTabs_(); w.ensureGameTabs_(); w.ensureMeTabs_(); w.ensurePollTabs_(); w.ensureWeekTabs_();
  addStudentTabs(w);
  w.CacheService.getScriptCache().put('clash-admin-T', '1');
  assert.throws(() => w.clashMe('sample-GOLD', 'nope'), /sign in as admin/i, 'admin only');
  const before = w.book.getSheetByName('Student Board').rows.length;
  ['BLACK', 'GOLD', 'GREY', 'PURPLE'].forEach(sq => {
    const d = JSON.parse(w.clashMe('sample-' + sq, 'T'));
    assert.strictEqual(d.sample, true);
    assert.strictEqual(d.me.squad, sq);
    assert.strictEqual(d.me.name, 'Sample S.');
    assert.strictEqual(d.metrics.find(m => m.key === 'bills').me, 14);
    assert.ok(d.creature && d.creature.xp > 0 && d.creature.stage === 2, 'creature with LRG points, partway up: ' + (d.creature && d.creature.xp));
    assert.ok(Array.isArray(d.tickets) && d.game, 'tickets and game tabs have data');
  });
  assert.strictEqual(w.book.getSheetByName('Student Board').rows.length, before, 'no sample row written');
  assert.ok(!w.book.getSheetByName('Student Board').rows.some(r => r[0] === '8000000'));
  assert.throws(() => w.clashMeSaveCreature({ version: 2 }, 'sample-GOLD', 'T'), /sample student/);
  assert.throws(() => w.clashMePickBoost(2, 'Trait Dough', 'sample-GOLD', 'T'), /sample student/);
  assert.throws(() => w.clashMeSaveLook({}, 'sample-GOLD', 'T'), /sample student/);
  const offer = JSON.parse(w.clashMeDrawTickets('sample-GOLD', 'T'));
  assert.ok(offer.offer.length >= 1, 'can see a ticket draw');
  assert.throws(() => w.clashMeKeepTickets([offer.offer[0].ticket], 'sample-GOLD', 'T'), /sample student/);
  // The real lookup still works afterwards (the Student Board read is back to normal).
  assert.strictEqual(w.boardRows_().length, before - 1);
});

test('location data: Pride Day check-ins (per 100) when Squad Monthly has no pride_day, poll kinds, ideas, insights, live board', () => {
  const w = world();
  const today = w.todayYmd_();
  w.ensureSettings_(); w.ensureGameTabs_(); w.ensurePollTabs_(); w.ensurePollIdeaTabs_();
  w.book.insertSheet('Student Board').rows = [['800#', 'Name', 'Squad', 'Grade'],
    ['8001111', 'Ann L.', 'Gold', '7'], ['8002222', 'Ed P.', 'Gold', '8'], ['8003333', 'Cy R.', 'Grey', '7'], ['8004444', 'Bo K.', 'Black', '8'], ['8005555', 'Di M.', 'Purple', '7']];
  assert.deepStrictEqual(J(w.lastMonthWindow_('2027-01-12')), { from: '2026-12-01', to: '2026-12-31', month: '2026-12' });
  assert.deepStrictEqual(J(w.lastMonthWindow_('2027-03-02')), { from: '2027-02-01', to: '2027-02-28', month: '2027-02' });
  assert.deepStrictEqual(J(w.monthBounds_(-1)), [w.lastMonthWindow_(today).from, w.lastMonthWindow_(today).to], 'the engine and Frost agree on last month');
  // Pride Day: a one-button check-in poll last month. GREY's 1 of 1 student beats GOLD's 1 of 2.
  const lm = w.lastMonthWindow_(today), d = w.ymdShift_(lm.from, 2);
  const polls = w.book.getSheetByName('Daily Polls'), votes = w.book.getSheetByName('Poll Votes');
  polls.rows.push([true, d, 'Pride Day Check-in', 'Wearing Mounts gear?', '', '', 1, 0, '', '']);
  const pid = polls.rows.length;
  votes.rows.push([pid, d, '8001111', 'GOLD', 'I\'m in! ✋', ''], [pid, d, '8003333', 'GREY', 'I\'m in! ✋', '']);
  // A School Pulse poll last month (counts for the Auditorium, not for Pride Day).
  polls.rows.push([true, d, 'School Pulse', 'How do you get to school?', 'Bus | Car', '', 1, 0, '', '']);
  const qid = polls.rows.length;
  votes.rows.push([qid, d, '8001111', 'GOLD', 'Bus', ''], [qid, d, '8002222', 'GOLD', 'Car', ''], [qid, d, '8004444', 'BLACK', 'Bus', '']);
  let zones = J(w.mapZones_(null, true));
  const z = n => zones.find(x => x.zone === n);
  assert.strictEqual(z('Gym 2').holder, 'GREY', 'Pride Day: 100 per 100 students beats 50');
  assert.deepStrictEqual([z('Gym 2').score.GOLD, z('Gym 2').score.GREY, z('Gym 2').score.PURPLE], [50, 100, 0]);
  assert.strictEqual(z('Auditorium').holder, 'GOLD', 'Wheel & Poll Turnout counts every vote and check-in: GOLD 3');
  // Once Climb-er (or staff) fills pride_day on Squad Monthly, that number decides instead.
  const sm = w.book.getSheetByName('Squad Monthly'), H = sm.rows[0];
  [['GOLD', 90], ['GREY', 40], ['BLACK', 30], ['PURPLE', 20]].forEach(([sq, v]) => sm.rows.push(H.map(h => h === 'Month' ? lm.month : h === 'Squad' ? sq : h === 'pride_day' ? v : '')));
  zones = J(w.mapZones_(null, true));
  assert.strictEqual(z('Gym 2').holder, 'GOLD', 'Squad Monthly pride_day wins over the check-ins');
  // Poll kinds: a check-in shows with one choice; School Pulse has no winner, only participation.
  polls.rows.push([true, today, 'Wheel Check-in', 'Spin today!', '', '', 1, 0, '', ''], [true, today, 'School Pulse', 'Sleep?', 'A | B', '', 1, 1, '', ''],
    [true, today, 'Trivia', '2+2?', '3 | 4', '4', 1, 1, '', '']);
  const dp = JSON.parse(w.clashDailyPolls(''));
  assert.deepStrictEqual(dp.today.map(p => p.kind), ['Check-in', 'Pulse', 'Trivia'], 'a School Pulse Kind is a Pulse');
  assert.deepStrictEqual(J(dp.today[0].choices), ['I\'m in! ✋']);
  w.clashPollVote(String(polls.rows.length - 2), '8001111', 'I\'m in! ✋');
  w.clashPollVote(String(polls.rows.length - 1), '8001111', 'A');
  w.clashPollVote(String(polls.rows.length), '8001111', '4');
  const pulse = w.pollsWithVotes_().find(p => p.question === 'How do you get to school?');
  assert.strictEqual(w.pollWinners_(pulse).length, 0, 'Pulse: nobody wins');
  assert.ok(/Pulse question/.test(w.pollAnnouncement_()) && /Check in/.test(w.pollAnnouncement_()));
  // The live board: counts live, but Trivia stays hidden until tomorrow; turnout per 100 students.
  const board = JSON.parse(w.clashPollBoard());
  assert.ok(board.live && board.polls.length === 3);
  assert.strictEqual(board.polls[2].counts, null, 'Trivia counts hidden today');
  assert.strictEqual(board.polls[1].counts.A, 1);
  assert.strictEqual(board.polls[0].turnout.GOLD, 50);
  assert.ok(!JSON.stringify(board).includes('8001111'), 'no 800#s on the board');
  // Poll Insights: answers by squad and grade, counts only.
  w.writePollInsights_();
  const ins = w.book.getSheetByName('Poll Insights').rows;
  const h = ins[0], bus = ins.find(r => r[h.indexOf('Question')] === 'How do you get to school?' && r[h.indexOf('Answer')] === 'Bus');
  assert.strictEqual(bus[h.indexOf('Share %')], 67, 'Bus is 2 of 3 answers');
  assert.strictEqual(bus[h.indexOf('Total')], 2);
  assert.strictEqual(bus[h.indexOf('Grade 7')], 1);
  assert.strictEqual(bus[h.indexOf('Grade 8')], 1);
  assert.strictEqual(bus[h.indexOf('BLACK')], 1);
  assert.ok(!JSON.stringify(ins).includes('8001111'));
  // Poll Ideas: a bank of questions; scheduling fills open school days and marks them used.
  const ideas = w.book.getSheetByName('Poll Ideas').rows;
  const ih = ideas[0];
  assert.ok(ideas.length > 80 && ideas.some(r => /Pride Day/.test(r[ih.indexOf('Kind')])));
  assert.ok(new Set(ideas.slice(1).map(r => r[ih.indexOf('Theme')])).size >= 12, 'many themes');
  assert.ok(ideas.slice(1).every(r => r[ih.indexOf('How To Use It')]), 'every idea says how to use it');
  const before = polls.rows.length;
  const lines = w.schedulePollIdeas_(3);
  assert.strictEqual(lines.length, 3);
  assert.strictEqual(polls.rows.length, before + 3);
  const added = polls.rows.slice(before).map(r => new Date(r[1] + 'T12:00:00Z').getUTCDay());
  assert.ok(added.every(dow => dow > 0 && dow < 6), 'school days only');
  assert.strictEqual(ideas.slice(1).filter(r => r[ih.indexOf('Used On')]).length, 3, 'Used On filled in');
  const themes = polls.rows.slice(before).map(r => w.sheetRows_('Poll Ideas').find(x => x.Question === r[3]).Theme);
  assert.ok(themes[0] !== themes[1] && themes[1] !== themes[2], 'themes rotate day to day');
  assert.ok(!polls.rows.slice(before).some(r => /check-in/i.test(r[2])), 'check-ins are not auto-scheduled');
  assert.strictEqual(w.schedulePollIdeas_(3).length, 0, 'those days are taken now');
  // Monthly trend questions come back after 28 days; others never repeat.
  const trend = w.sheetRows_('Poll Ideas').find(r => /month/.test(r.Repeat)), once = w.sheetRows_('Poll Ideas').find(r => !r.Repeat && r.Kind === 'School Pulse');
  const uc = ih.indexOf('Used On');
  ideas[trend._row - 1][uc] = w.ymdShift_(today, -40); ideas[once._row - 1][uc] = w.ymdShift_(today, -40);
  const ready = w.readyIdeas_(today).map(r => r.Question);
  assert.ok(ready.includes(trend.Question) && !ready.includes(once.Question));
  // An older Poll Ideas tab gets the new columns and the new questions, and keeps staff edits.
  const w4 = world(); w4.ensureSettings_();
  const oldIdeas = w4.book.insertSheet('Poll Ideas');
  oldIdeas.rows = [['Use', 'Kind', 'Question', 'Choices', 'Why We Ask', 'Used On'], [false, 'School Pulse', 'Do you feel like you belong at Ephrata Middle School?', 'Yes | No', 'mine', '2026-11-01']];
  w4.ensurePollIdeaTabs_();
  assert.ok(['Theme', 'How To Use It', 'Repeat'].every(h => oldIdeas.rows[0].includes(h)));
  assert.strictEqual(oldIdeas.rows.filter(r => r[2] === 'Do you feel like you belong at Ephrata Middle School?').length, 1, 'no duplicate');
  assert.strictEqual(oldIdeas.rows[1][3], 'Yes | No', 'staff edits kept');
  assert.ok(oldIdeas.rows.length > 80);

});

test('Cafeteria counts LRRs + LDs from the Follow-Up sheets; every area\'s multiplier can be set any time', () => {
  const w = world();
  const today = w.todayYmd_(), lm = w.lastMonthWindow_(today), d = w.ymdShift_(lm.from, 3);
  const [y, m, dd] = d.split('-');
  const us = Number(m) + '/' + Number(dd) + '/' + y + ' 12:05:00';
  w.ensureSettings_(); w.ensureGameTabs_();
  w.book.insertSheet('Students').rows = [['800#', 'First Name', 'Last Name', 'Grade', 'Email', 'Squad'],
    ['8001111', 'Ann', 'Lee', '7', '', 'Gold'], ['8002222', 'Ed', 'Park', '8', '', 'Gold'], ['8003333', 'Cy', 'Ro', '7', '', 'Grey'], ['8004444', 'Bo', 'Kim', '8', '', 'Black'], ['8005555', 'Di', 'May', '7', '', 'Purple']];
  // The LRR responses (Form tab) and the Master Log, like EMS Follow-Up reads them.
  const sheet = (name, rows, form) => { const s = { getName: () => name, getFormUrl: () => form ? 'form' : '', getLastRow: () => rows.length,
    getDataRange: () => ({ getValues: () => rows.map(r => r.slice()) }) }; return s; };
  const lrr = sheet('Form Responses 1', [['Timestamp', 'Email Address', 'First Name', 'Last Name'],
    [us, '8001111@gomounts.org', 'Ann', 'Lee'], [us, '8001111@gomounts.org', 'Ann', 'Lee'], // the same student twice = 1
    [us, 'teacher@easdpa.org', 'Cy', 'Ro'],                                                 // matched by name
    ['1/2/2020 9:00:00', '8004444@gomounts.org', 'Bo', 'Kim'],                               // another month
    [us, 'x@easdpa.org', 'Nobody', 'Here']], true);
  const log = sheet('Master Log', [['Entry Date', 'Student Name', 'Student 800#', 'Consequence Type', 'Consequence Date', 'Status'],
    [d, 'Kim, Bo', '', 'Admin Lunch Detention', d, ''], [d, 'Ed Park', '8002222', 'Teacher Lunch Detention', d, 'Cancelled'],
    [d, 'Ed Park', '8002222', 'Time in Office', d, '']]);
  w.SpreadsheetApp.openById = id => {
    if (id === '1g8VpAeyE8md-1us58uHZxtRAKKFCTCKHhuwC8ZOpzyc') return { getSheets: () => [sheet('Notes', [['x']]), lrr], getSheetByName: () => null };
    if (id === '1IVFnkQQvSvNSY_OZPMutPzkTqmJ1j5XYRoT9YvNy-v4') return { getSheets: () => [log], getSheetByName: n => n === 'Master Log' ? log : null };
    throw new Error('no ' + id);
  };
  const c = J(w.followUpLrrCounts_());
  assert.deepStrictEqual(c.counts, { BLACK: 1, GOLD: 1, GREY: 1, PURPLE: 0 }, 'one per student; cancelled LDs and other consequences skipped');
  assert.strictEqual(c.lrr, 4); assert.strictEqual(c.ld, 1); assert.strictEqual(c.unmatched, 1);
  assert.ok(/lrr_ld: BLACK 1, GOLD 1, GREY 1, PURPLE 0/.test(w.sendFollowUpStats_()));
  const sm = w.book.getSheetByName('Squad Monthly').rows, H = sm[0], lc = H.indexOf('lrr_ld');
  const rowOf = sq => sm.find(r => w.ymOf_(r[0]) === lm.month && r[1] === sq);
  assert.ok(SQUADS.every(sq => rowOf(sq)), 'a Squad Monthly row per squad for last month');
  assert.deepStrictEqual(SQUADS.map(sq => rowOf(sq)[lc]), [1, 1, 1, 0]);
  w.sendFollowUpStats_();
  assert.strictEqual(sm.filter(r => w.ymOf_(r[0]) === lm.month).length, 4, 'updates its own cells, no duplicate rows');
  let zones = J(w.mapZones_(null, true));
  assert.strictEqual(zones.find(z => z.zone === 'Cafeteria').holder, 'PURPLE', 'fewest LRR + LD students per 100 (Students tab sizes)');
  rowOf('GOLD')[lc] = 0; // staff typed a correction
  assert.ok(/Kept the number staff typed for GOLD/.test(w.sendFollowUpStats_()));
  assert.strictEqual(rowOf('GOLD')[lc], 0);
  zones = J(w.mapZones_(null, true));
  assert.strictEqual(zones.find(z => z.zone === 'Cafeteria').holder, 'PURPLE', 'GOLD ties at 0: the holder keeps it');
  // Multipliers: an area's own Multiplier beats its rule's; blank goes back to the rule.
  const mz = w.book.getSheetByName('Map Zones'), h = mz.rows[0], mcol = h.indexOf('Multiplier');
  assert.ok(mcol > -1, 'Map Zones has a Multiplier column');
  assert.strictEqual(zones.find(z => z.zone === 'Cafeteria').mult, 1.5);
  mz.rows.forEach(r => { if (r[1] === 'Cafeteria') r[mcol] = 4; if (r[1] === 'Music Wing') r[mcol] = 0.5; });
  zones = J(w.mapZones_(null, false));
  assert.strictEqual(zones.find(z => z.zone === 'Cafeteria').mult, 4);
  assert.strictEqual(zones.find(z => z.zone === 'Music Wing').mult, 0.5);
  assert.strictEqual(zones.find(z => z.zone === 'Auditorium').mult, 2, 'check-ins x2');
  // The weekly pay uses it: PURPLE holds the Cafeteria (Summit): 5 x 4 x 2 = 40.
  w.setSettingValue_('GAME_AWARDS', 'ON'); w.props['game-paid-week'] = 'old';
  w.ScriptApp = { getProjectTriggers: () => [], newTrigger: () => ({ timeBased: () => ({ everyHours: () => ({ create() {} }) }) }) };
  w.gameTick();
  const ev = w.book.getSheetByName('events').rows, terr = ev.filter(r => r[2] === 'territory').find(r => r[9] > 0);
  assert.ok(terr && terr[9] >= 40, 'the Cafeteria pays 5 x 4 x 2');
  // The Map editor round-trips it; blank clears it.
  w.CacheService.getScriptCache().put('clash-admin-T', '1');
  const loaded = JSON.parse(w.clashMapLoad('T'));
  assert.strictEqual(loaded.zones.find(z => z.zone === 'Cafeteria').mult, 4);
  loaded.zones.find(z => z.zone === 'Cafeteria').mult = '2.5';
  loaded.zones.find(z => z.zone === 'Music Wing').mult = '';
  w.clashMapSave('T', JSON.stringify(loaded));
  zones = J(w.mapZones_(null, false));
  assert.strictEqual(zones.find(z => z.zone === 'Cafeteria').mult, 2.5);
  assert.strictEqual(zones.find(z => z.zone === 'Music Wing').mult, 1, 'blank = 1');
});

test('frost check: same lines for every squad, a warning first, a repeat or two lines = freeze proposed, staff approve, thaw together, meter + explainer', () => {
  const w = world();
  const today = w.todayYmd_(), lm = w.lastMonthWindow_(today).month, prev = w.prevMonth_(lm), inLm = lm + '-05';
  w.ensureSettings_(); w.ensureGameTabs_(); w.ensureFrostTabs_();
  assert.strictEqual(w.book.getSheetByName('Freeze Lines').rows.length, 6, 'five starting lines');
  w.book.insertSheet('Student Board').rows = [['800#', 'Name', 'Squad', 'Grade', 'Last Contribution'],
    ['8001111', 'Ann L.', 'Gold', '7', today], ['8002222', 'Ed P.', 'Gold', '8', ''], ['8003333', 'Cy R.', 'Grey', '7', ''], ['8004444', 'Bo K.', 'Black', '8', '']];
  // Lines for the test (metrics tab: attendance BLACK 90, GOLD 91, GREY 92, PURPLE 93).
  w.writeRowsByHeader_(w.book.getSheetByName('Freeze Lines'), [
    { Show: true, Marker: 'LRR + LD students', Measures: 'stat: LRR + LD Students', 'Freezes If': 'Above', Line: 50, 'Per 100 Students': true, Kind: 'Peaks', 'What Students See': 'Students with an LRR or LD' },
    { Show: true, Marker: 'Attendance', Measures: 'attendance', 'Freezes If': 'Below', Line: 92, 'Per 100 Students': false, Kind: 'Peaks', 'What Students See': 'Attendance rate' },
    { Show: true, Marker: 'Points floor', Measures: 'points', 'Freezes If': 'Below', Line: 1, 'Per 100 Students': true, Kind: 'Fun', 'What Students See': 'Points earned' }]);
  // BLACK's LRR + LD number on Squad Monthly (the older 'stat: LRR + LD Students' name still reads lrr_ld).
  w.ensureConquestTabs_();
  const smH = w.book.getSheetByName('Squad Monthly').rows[0];
  w.book.getSheetByName('Squad Monthly').rows.push(smH.map(h => h === 'Month' ? lm : h === 'Squad' ? 'BLACK' : h === 'lrr_ld' ? 1 : ''));
  assert.deepStrictEqual(J(w.schoolStat_('lrr_ld', lm)), { BLACK: 1, GOLD: null, GREY: null, PURPLE: null });
  assert.deepStrictEqual(J(w.schoolStat_('LRR + LD Students', lm)), { BLACK: 1, GOLD: null, GREY: null, PURPLE: null }, 'old name, same column');
  w.book.getSheetByName('events').rows.splice(1); // only this test's points
  w.book.getSheetByName('events').rows.push([501, 'Spirit', 'event', inLm, '', '', 0, 10, 10, 0], [502, 'Spirit', 'event', prev + '-05', '', '', 0, 10, 10, 0]);
  // Before FROST_START nothing is checked.
  w.setSettingValue_('FROST_START', '2099-01');
  assert.ok(/before FROST_START/.test(w.runFrostCheck_(lm)[0]));
  w.setSettingValue_('FROST_START', prev);
  // Month 1: GOLD crosses attendance only = a frost warning. BLACK crosses attendance + points = freeze proposed.
  w.runFrostCheck_(prev);
  let log = w.sheetRows_('Frost Log');
  const st = (m, s) => (log.find(r => String(r.Month).slice(0, 7) === m && r.Squad === s) || {}).Status;
  assert.strictEqual(st(prev, 'GOLD'), 'Frost warning');
  assert.strictEqual(st(prev, 'GREY'), 'Safe');
  assert.strictEqual(st(prev, 'BLACK'), 'Freeze proposed');
  assert.strictEqual(st(prev, 'PURPLE'), 'Safe', 'no students on the roster: per-100 lines have nothing to measure');
  // Month 2: GOLD crosses attendance again = freeze proposed (the same line two months running).
  w.runFrostCheck_(lm);
  log = w.sheetRows_('Frost Log');
  assert.strictEqual(st(lm, 'GOLD'), 'Freeze proposed');
  assert.strictEqual(log.find(r => String(r.Month).slice(0, 7) === lm && r.Squad === 'GOLD').Kind, 'Peaks');
  assert.strictEqual(w.runFrostCheck_(lm).length, 0, 'each month is checked once');
  // Nothing freezes until staff approve.
  assert.strictEqual(w.freezeDetail_('GOLD'), null);
  assert.strictEqual(w.applyApprovedFrost_().length, 0);
  const fl = w.book.getSheetByName('Frost Log'), ap = fl.rows[0].indexOf('Approve');
  fl.rows.forEach(r => { if (String(r[0]).slice(0, 7) === lm && r[1] === 'GOLD') r[ap] = true; });
  assert.ok(/GOLD: Peaks Freeze started/.test(w.applyApprovedFrost_()[0]));
  assert.strictEqual(w.applyApprovedFrost_().length, 0, 'applied once');
  const fz = J(w.freezeDetail_('GOLD'));
  assert.strictEqual(fz.kind, 'Peaks');
  assert.deepStrictEqual(fz.goals.map(g => g.type), ['marker', 'everyone', 'points', 'staff'], 'fix the line, everyone pitches in, earn points, a reflection');
  assert.strictEqual(fz.goals[0].met, false, 'the line is checked on the 1st after a whole month');
  assert.strictEqual(fz.ready, false);
  assert.strictEqual(w.freezeDetail_('BLACK'), null, 'a proposal nobody approved never freezes');
  // Fix the line: a whole month on the safe side after the freeze began.
  assert.strictEqual(w.markerSafe_('GREY', 'Attendance', '2020-01-01'), 1);
  assert.strictEqual(w.markerSafe_('GOLD', 'Attendance', '2020-01-01'), 0);
  // A squad already frozen is logged Still frozen, not proposed again.
  const next = w.runFrostCheck_(lm, true).find(l => /^GOLD/.test(l));
  assert.ok(/Still frozen/.test(next));
  // What students see: squad totals vs each line, no names, no 800#s.
  const meter = J(w.frostMeter_());
  assert.strictEqual(meter.GOLD.markers.length, 3);
  assert.strictEqual(meter.GOLD.live, true);
  assert.ok(meter.GOLD.markers.find(m => m.label === 'Attendance rate').over);
  const feed = w.clashTvFeed();
  assert.ok(JSON.parse(feed).frost.BLACK, 'the TV gets the meter');
  const ex = JSON.parse(w.clashFrostExplainer());
  assert.strictEqual(ex.lines.length, 3);
  assert.ok(ex.lines.find(l => l.marker === 'Attendance').pct);
  assert.ok(/together/.test(ex.script) && /No one thaws alone/.test(ex.script));
  assert.ok(!/8001111|Ann L/.test(JSON.stringify(ex) + feed), 'no names or 800#s');
  assert.ok(w.thawGoals_('marker: Interventions')[0].type === 'marker');
});

test('insights: points balance flags, You said we did, State of the Squads email (once a month), privacy check of every public page, role guides from Settings', () => {
  const w = world();
  w.ensureSettings_(); w.ensureGameTabs_(); w.ensurePollTabs_(); w.ensureFrostTabs_(); w.ensureYouSaidTab_(); w.ensureWeekTabs_();
  const today = w.todayYmd_(), ago = n => w.ymdShift_(today, -n);
  w.book.insertSheet('Students').rows = [['800#', 'First Name', 'Last Name', 'Grade', 'Email', 'Squad'],
    ['8001111', 'Ann', 'Lee', '7', '', 'Gold'], ['8002222', 'Ed', 'Park', '8', '', 'Grey'], ['8003333', 'Cy', 'Ro', '7', '', 'Black'], ['8004444', 'Di', 'May', '8', '', 'Purple']];
  w.book.insertSheet('Student Board').rows = [['800#', 'Name', 'Squad', 'Grade'], ['8001111', 'Ann L.', 'Gold', '7'], ['8002222', 'Ed P.', 'Grey', '8']];
  // ── (2) Points balance: Trait Dough is most of everything, GOLD's lead is all Trait Dough, only GREY did duels.
  const ev = w.book.getSheetByName('events');
  ev.rows.splice(1);
  ev.rows.push([1, 'Trait Dough', 'trait_dough', ago(3), '', '', 10, 400, 10, 10], [2, 'Duel', 'duel_bonus', ago(4), '', '', 0, 0, 60, 0],
    [3, 'Poll', 'daily_poll', ago(5), '', '', 20, 20, 20, 20], [4, 'Old', 'trait_dough', ago(90), '', '', 0, 999, 0, 0],
    [5, 'TJ', 'trail_journal', ago(2), '', '', 0, 0, 0, 900]);
  const b = J(w.pointsBalance_());
  assert.strictEqual(b.total, 570, 'last 30 days only; private behavior rows never count');
  assert.ok(b.flags.some(l => /Trait Dough is 75\.4% of every point/.test(l)), b.flags.join(' | '));
  assert.ok(b.flags.some(l => /Only Grey earned .*Squad Duel/.test(l)));
  assert.ok(b.flags.some(l => /Gold's lead over Grey .* is mostly .*Trait Dough/.test(l)));
  assert.ok(b.flags.some(l => /One row gave Gold 400 points/.test(l)));
  w.writePointsBalance_();
  assert.strictEqual(w.book.getSheetByName('Points Balance').rows[1][0], 'trait_dough');
  // ── (3) You said, we did: a closed Pulse with enough answers gets a row; students see it once staff answer and tick Show.
  const polls = w.book.getSheetByName('Daily Polls'), votes = w.book.getSheetByName('Poll Votes');
  polls.rows.splice(1);
  polls.rows.push([true, ago(6), 'Pulse', 'What would make lunch better?', 'More time | Quieter', '', 1, '', '', '', ago(4)]);
  for (let i = 0; i < 12; i++) votes.rows.push([2, ago(5), '80099' + String(i).padStart(2, '0'), 'GOLD', i < 9 ? 'More time' : 'Quieter', '', '7']);
  w.setSettingValue_('YOU_SAID_MIN', 10);
  assert.strictEqual(w.syncYouSaid_(), 1);
  assert.strictEqual(w.syncYouSaid_(), 0, 'once per poll');
  const ys = w.book.getSheetByName('You Said We Did'), H = ys.rows[0];
  const row = ys.rows[1];
  assert.strictEqual(row[H.indexOf('What You Said')], 'Most said "More time" (75%), then "Quieter" (25%) · 12 answers');
  assert.deepStrictEqual(J(w.youSaidPublic_()), [], 'nothing public until staff answer');
  row[H.indexOf('What We Did')] = 'Lunch bell moved 5 minutes later'; row[H.indexOf('Who')] = 'the office';
  assert.deepStrictEqual(J(w.youSaidPublic_()), [], 'and tick Show');
  row[H.indexOf('Show')] = true;
  assert.strictEqual(JSON.parse(w.clashYouSaid())[0].did, 'Lunch bell moved 5 minutes later');
  assert.strictEqual(JSON.parse(w.clashPulseLive()).youSaid.length, 1, 'on the Pulse page');
  assert.strictEqual(JSON.parse(w.clashTvFeed()).youSaid.length, 1, 'and the TV');
  // ── (6) Privacy: no public page shows an 800# or a full name; a planted one is found and masked.
  let scan = J(w.privacyScan_());
  assert.ok(scan.length >= 15);
  assert.deepStrictEqual(scan.filter(x => x.findings.length), [], JSON.stringify(scan.filter(x => x.findings.length)));
  ev.rows.push([6, 'Shout-out for Ann Lee 8001111', 'shout_out', ago(1), '', '', 0, 1, 0, 0]);
  scan = J(w.privacyScan_());
  const hit = scan.filter(x => x.findings.length);
  assert.deepStrictEqual(hit.map(x => x.surface), ['Site data', 'TV feed', 'This Week'], 'a full name typed into an event name is caught');
  assert.deepStrictEqual(hit[0].findings, ['800# …111', 'full name A. L.'], 'and masked in the report');
  assert.ok(!JSON.stringify(hit).includes('8001111') && !/ann lee/i.test(JSON.stringify(hit)));
  assert.ok(scan.every(x => !x.note), 'every public page ran');
  ev.rows.pop();
  // ── (5) State of the Squads: last month's email, once a month, to STATE_EMAILS (or DIGEST_EMAILS).
  const sent = [];
  w.MailApp = { sendEmail: o => sent.push(o) };
  assert.ok(/No recipients/.test(w.stateOfSquadsEmail(false)));
  w.setSettingValue_('DIGEST_EMAILS', 'office@easdpa.org');
  const d = J(w.stateOfSquadsData_());
  assert.strictEqual(d.month, w.lastMonthWindow_(today).month);
  assert.ok(d.zones.length === 12 && Array.isArray(d.flags));
  assert.strictEqual(w.stateOfSquadsEmail(false), 'Sent to 1.');
  assert.strictEqual(w.stateOfSquadsEmail(false), 'Already sent for ' + d.month + '.');
  assert.ok(/State of the Squads/.test(sent[0].subject) && /The map/.test(sent[0].htmlBody) && /Squad Monthly numbers/.test(sent[0].htmlBody));
  assert.ok(!/8001111|Ann Lee/.test(sent[0].htmlBody), 'counts only');
  // ── (7) Role guides: written from Settings, nothing secret.
  w.setSettingValue_('GENERALS_CODE', 'SECRET-CODE-42');
  w.setSettingValue_('PULSE_DAYS', '4');
  const g = JSON.parse(w.clashGuides());
  assert.deepStrictEqual(g.roles.map(r => r.id), ['students', 'generals', 'staff', 'admin']);
  const text = JSON.stringify(g);
  assert.ok(/open 4 school days/.test(text), 'a Setting changes the guide');
  assert.ok(/Gym 2: .*x2\./.test(text) && /Cafeteria \(Summit\)/.test(text), 'each area with its multiplier');
  assert.ok(!/SECRET-CODE-42/.test(text) && !/8001111/.test(text) && !/1g8VpAeyE8md/.test(text), 'no code, 800# or sheet ID');
});

test('Redemption: a Trailback Summit on the private Redemptions tab gives the student a Summit Climber badge and 25 LRG points; the feed labels it', () => {
  const w = world();
  w.ensureSettings_();
  w.book.insertSheet('Student Board').rows = [['800#', 'Name', 'Squad', 'Grade', 'Badges'], ['8001111', 'Ann L.', 'Gold', '7', '🍞 First Trait Dough']];
  let me = JSON.parse(w.clashMyPage('8001111')).me;
  assert.strictEqual(me.redemption, 0);
  assert.ok(!me.badges.includes('🏔️ Summit Climber'));
  w.book.insertSheet('Redemptions').rows = [['800#', 'Squad', 'Date', 'Points', 'Case'], ['8001111', 'GOLD', '2027-01-20', 75, 'TB-1']];
  me = JSON.parse(w.clashMyPage('8001111')).me;
  assert.strictEqual(me.redemption, 1);
  assert.deepStrictEqual(J(me.badges), ['🍞 First Trait Dough', '🏔️ Summit Climber']);
  assert.strictEqual(w.climbXp_({ redemption: 1 }), 25);
  w.setSettingValue_('REDEMPTION_LRG', '40');
  assert.strictEqual(w.climbXp_({ redemption: 1 }), 40);
  w.book.getSheetByName('events').rows.push([501, 'Redemption: a Gold climber reached the Summit', 'redemption', w.todayYmd_(), 'Trailback Summit', '', '', 75, '', '']);
  const feed = JSON.parse(w.clashTvFeed());
  assert.ok(feed.latest.some(e => e.label === '🏔️ Redemption' && /Gold climber/.test(e.name)));
});

console.log('\n' + passed + ' tests passed');
