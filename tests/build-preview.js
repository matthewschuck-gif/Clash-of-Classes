// Local preview of the Apps Script site: renders Index.html with its includes, and runs the
// real .gs server code in the browser against fake Google services with sample data.
//   node tests/build-preview.js   → http://localhost:3041   (admin password in the preview: preview)
const fs = require('fs');
const path = require('path');
const http = require('http');
const dir = path.join(__dirname, '..', 'apps-script');
const read = f => fs.readFileSync(path.join(dir, f), 'utf8');

function build(page) {
  // ?page=me → My Climb (Me.html), signed in as the sample student Ann L.
  let html = (page === 'me' ? read('Me.html').replace(/<? if (shield) { ?>|<? } ?>/g, '').replace('<?= shield ?>', '/shield.png') : read('Index.html'))
    .replace(/<\?!= include\('([A-Za-z]+)'\); \?>/g, (m, f) => read(f + '.html'))
    .replace('<?!= config ?>', JSON.stringify({ pollMs: 5000 }));
  const gs = fs.readdirSync(dir).filter(f => f.endsWith('.gs')).sort().map(read).join('\n');
  const fake = `<script>
(function () {
  function makeSheet(name, rows) {
    var s = { rows: rows, getName: function () { return name; },
      getLastRow: function () { return this.rows.length; },
      getLastColumn: function () { return this.rows.reduce(function (m, r) { return Math.max(m, r.length); }, 0); },
      getDataRange: function () { var t = this; return { getValues: function () { return t.rows.map(function (r) { return r.slice(); }); } }; },
      getRange: function (r, c, nr, nc) { var t = this; nr = nr || 1; nc = nc || 1; var p; var api = {
        getValues: function () { var o = []; for (var i = 0; i < nr; i++) { var row = t.rows[r - 1 + i] || []; var l = []; for (var j = 0; j < nc; j++) l.push(row[c - 1 + j] == null ? '' : row[c - 1 + j]); o.push(l); } return o; },
        setValues: function (v) { v.forEach(function (row, i) { t.rows[r - 1 + i] = t.rows[r - 1 + i] || []; row.forEach(function (x, j) { t.rows[r - 1 + i][c - 1 + j] = x; }); }); return p; },
        setValue: function (v) { return api.setValues([[v]]); },
        clearContent: function () { for (var i = 0; i < nr; i++) { var row = t.rows[r - 1 + i]; if (row) for (var j = 0; j < nc; j++) row[c - 1 + j] = ''; } return p; } };
        p = new Proxy(api, { get: function (o, k) { return k in o ? o[k] : function () { return p; }; } }); return p; },
      appendRow: function (row) { this.rows.push(row.slice()); }, deleteRow: function (n) { this.rows.splice(n - 1, 1); },
      setFrozenRows: function () {}, setColumnWidth: function () {} };
    return s;
  }
  var blob = { poll: { active: true, question: 'Which squad wins November?', options: ['Black', 'Gold', 'Grey', 'Purple'], votes: { 1: 4 } }, metrics: { month: 'November 2026' }, nextId: 104 };
  var sheets = [
    makeSheet('app_data', [['key', 'chunk_index', 'value_chunk', 'updated_at'], ['coc', 0, JSON.stringify(blob), '2026-10-31']]),
    makeSheet('events', [['id', 'name', 'type', 'date', 'note', 'multiplier', 'BLACK', 'GOLD', 'GREY', 'PURPLE'],
      [101, 'Kickoff Assembly', 'event', '2026-09-02', '', '', 120, 100, 90, 110],
      [102, 'Trait Dough: week 6', 'trait_dough', '2026-10-30', '', '', 12, 18, 9, 14],
      [103, 'STAR Growth: Cycle 2', 'star_growth', '2026-10-31', '', '', 40, 52, 47, 38]]),
    makeSheet('metrics', [['squad', 'attendance', 'referral', 'grades'], ['BLACK', 94, 88, 81], ['GOLD', 95, 90, 84], ['GREY', 93, 86, 80], ['PURPLE', 96, 91, 79]]),
    makeSheet('Student Board', [['800#', 'Name', 'Squad', 'Grade', 'Trait Dough Bills', 'Bills This Week', 'Trait Dough Wins', 'Mindset Summits', 'Trail Journal Growth', 'STAR Growth Points', 'Shout-outs', 'Streak (weeks)', 'Badges', 'Last Contribution', 'Updated', 'Attendance %', 'Email'],
      ['8001234', 'Ann L.', 'Gold', '7', '12', '2', '1', '1', '', '4', '2', '3', '🍞 First Trait Dough, 🍞 Trait Dough Ten, 🏆 Trait Dough Winner, ⛰️ Mindset Summit, 📈 STAR Grower, 🙌 Shout-out, 🔥 3-Week Streak', '2026-10-30', '2026-10-31 19:00', '96.5', 'ann.l@example.org'],
      ['8005678', 'Bo K.', 'Grey', '8', '3', '0', '0', '0', '', '2', '0', '0', '🍞 First Trait Dough, 📈 STAR Grower', '2026-10-20', '2026-10-31 19:00', '91', 'bo.k@example.org']]),
    makeSheet('Challenges', [['Show', 'Title', 'Description', 'Event Types', 'Measure', 'Goal', 'Start', 'End', 'Bonus Points', 'Awarded'],
      [true, 'Trait Dough Week', 'Turn in Trait Dough bills all week!', 'trait_dough', 'Points', 25, '2026-10-26', '2099-12-31', 50, ''],
      [true, 'Growth Push', 'Every point from any project counts.', '', 'Points', 200, '2026-09-01', '2099-12-31', 0, '']]),
  ];
  var book = { getSheets: function () { return sheets; }, getSheetByName: function (n) { return sheets.filter(function (s) { return s.getName() === n; })[0] || null; },
    insertSheet: function (n) { var s = makeSheet(n, []); sheets.push(s); return s; }, getId: function () { return 'CLASH'; } };
  var cache = {}, props = {}, updated = Date.now();
  window.SpreadsheetApp = { getActiveSpreadsheet: function () { return book; },
    newDataValidation: function () { var b = new Proxy({}, { get: function () { return function () { return b; }; } }); return b; }, openById: function () { return { getSheetByName: function (n) {
    return n === 'reflections' ? makeSheet(n, [['insight_level', 'created_at'], ['High', new Date()], ['Medium', new Date()]]) : n === 'panel_sessions' ? makeSheet(n, [['clash_squad', 'status'], ['GOLD', 'complete']]) : null; } }; } };
  window.DriveApp = { getFileById: function () { return { getLastUpdated: function () { return new Date(updated); } }; } };
  window.CacheService = { getScriptCache: function () { return { get: function (k) { return cache[k] || null; }, put: function (k, v) { cache[k] = v; }, remove: function (k) { delete cache[k]; } }; } };
  window.PropertiesService = { getScriptProperties: function () { return { getProperty: function (k) { return props[k] || null; }, setProperty: function (k, v) { props[k] = v; } }; } };
  window.LockService = { getScriptLock: function () { return { waitLock: function () {}, releaseLock: function () {} }; } };
  window.Utilities = { getUuid: function () { return 'u' + Math.random().toString(36).slice(2); },
    computeDigest: function (a, s) { return s.split('').map(function (ch) { return ch.charCodeAt(0) % 128; }); }, // preview only
    DigestAlgorithm: {}, Charset: {}, formatDate: function (d) { return d.toISOString().slice(0, 10); } };
  window.Session = { getScriptTimeZone: function () { return 'America/New_York'; }, getActiveUser: function () { return { getEmail: function () { return 'ann.l@example.org'; } }; } };
  window.HtmlService = {}; window.ContentService = {};
  window.__previewTouch = function () { updated = Date.now(); };
  window.__previewSheets = sheets;
})();
</script>
<script>${gs.replace(/<\/script>/g, '<\\/script>')}
setPassword_('preview');
ensureSettings_(); setSettingValue_('GENERALS_CODE', 'preview'); ensureGeneralsTabs_();
getSheet_('MLC Roster').appendRow(['8001234', 'Ann L.', 'Gold', 'General', '']);
getSheet_('Sign-up Events').appendRow(['Dodgeball', '2026-11-20', 8, true]);
getSheet_('Sign-up Events').appendRow(['Lip Sync Battle', '2026-12-04', 4, true]);
getSheet_('Student Board').appendRow(['8009999', 'Gus T.', 'Gold', '7', '1', '0', '0', '0', '', '0', '0', '0', '', '', '', '89', 'gus.t@example.org']);
ensurePollTabs_(); getSheet_('Daily Polls').rows.slice(1).forEach(function (r) { r[0] = true; });
ensureFreezeTab_(); getSheet_('Point Freezes').appendRow(['GREY', '2026-11-01', '2099-11-07', 'preview', '', 'Peaks', 'points: 30; everyone: 60%; Squad reflection with your advisor', '', false, '', '']);
getSheet_('Point Freezes').appendRow(['GOLD', '2026-11-01', '', 'preview', '', 'Fun', 'points: 4', '2099-11-20', false, '', '']);
ensureDuelTab_(); getSheet_('Squad Duels').appendRow([true, 'Gold vs Grey Dough Duel', '2026-10-26', '2099-12-31', 'Points', 'trait_dough', 'GOLD vs GREY', 25, '', '']);
ensureLoreTab_(); getSheet_('Lore Chapters').rows.splice(1);
getSheet_('Lore Chapters').appendRow([true, 'ALL', 1, 'The Lantern at the Base', 'Long before the first bell, four lanterns were lit at the foot of the mountain…', '', '', '']);
getSheet_('Lore Chapters').appendRow([true, 'ALL', 2, 'The Ridge Path', 'The second chapter.', '', 'points: 300', '']);
getSheet_('Lore Chapters').appendRow([true, 'ALL', 3, 'Above the Clouds', 'The third chapter.', '', 'points: 5000', '']);
ensureEventPageTabs_(); ensureMayhemWeekTabs_(); setSettingValue_('SELF_SIGNUPS', 'ON');
getSheet_('Sign-up Events').appendRow(['Mario Kart', '2027-03-15', 4, true, 'Lunch', 'Gaming', true]);
getSheet_('Sign-up Events').appendRow(['Squad Bowl (trivia)', '2027-03-16', 4, true, 'Core Extension', 'Brain', true]);
getSheet_('Sign-up Events').appendRow(['Tug-o-war', '2027-03-19', 15, true, 'Squad Night', 'Active', false]);
getSheet_('events').appendRow([104, 'Paper Chain Challenge', 'mayhem_week', '2026-10-30', '', '', 3, 5, 1, 2]);
getSheet_('events').appendRow([105, 'Penny Tower Race', 'mayhem_week', '2026-10-31', '', '', 2, 0, 4, 1]);
ensureGameTabs_(); ensureMeTabs_();
setSettingValue_('MY_CLIMB_HOURS', '00:00-23:59'); setSettingValue_('MY_CLIMB_SCHOOL_DAYS', 'OFF'); // preview only: always open
getSheet_('Opponents').rows.forEach(function (r, i) { if (i && r[3] === 'Shady') r[0] = true; });
(function () { var t = todayYmd_(); getSheet_('events').appendRow([106, 'Daily Poll', 'daily_poll', t, '', '', 2, 6, 3, 1]); getSheet_('events').appendRow([107, 'Mindset summits', 'mountaineer_mindset', t, '', '', 4, 1, 2, 6]); getSheet_('events').appendRow([108, 'Trait Dough: Ann L.', 'trait_dough', t, '', '', 0, 9, 3, 2]); })();
getSheet_('Squad Hand').appendRow(['GOLD', 'Flip Three', 'flip_three', 0, todayYmd_(), 'Held', '']);
</script>
<script>
// google.script.run → the .gs functions above, asynchronously, like the real thing.
window.google = { script: { get run() {
  var ok = function () {}, fail = function (e) { console.error(e); };
  var runner = new Proxy({}, { get: function (t, name) {
    if (name === 'withSuccessHandler') return function (f) { ok = f; return runner; };
    if (name === 'withFailureHandler') return function (f) { fail = f; return runner; };
    return function () { var args = arguments; setTimeout(function () {
      try { var r = window[name].apply(null, args); if (/^clash(Save|Vote|SetPassword)$/.test(name)) window.__previewTouch(); ok(r); } catch (e) { fail(e); } }, 30); };
  } });
  return runner;
} } };
</script>`;
  return html.replace('<head>', () => '<head>\n' + fake); // a function, so "$&" in the server code stays as written
}

const out = path.join(__dirname, 'preview.html');
fs.writeFileSync(out, build());
http.createServer((req, res) => {
  if (req.url === '/shield.png') { res.writeHead(200, { 'Content-Type': 'image/png' }); return res.end(fs.readFileSync(path.join(__dirname, '..', 'assets', 'shield-2026-27.png'))); }
  fs.writeFileSync(out, build(/page=me/.test(req.url) ? 'me' : ''));
  res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
  res.end(fs.readFileSync(out));
}).listen(Number(process.env.PORT) || 3041, () => console.log('Preview on http://localhost:' + (Number(process.env.PORT) || 3041)));
