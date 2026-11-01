// Local preview of the Apps Script site: renders Index.html with its includes, and runs the
// real .gs server code in the browser against fake Google services with sample data.
//   node tests/build-preview.js   → http://localhost:3041   (admin password in the preview: preview)
const fs = require('fs');
const path = require('path');
const http = require('http');
const dir = path.join(__dirname, '..', 'apps-script');
const read = f => fs.readFileSync(path.join(dir, f), 'utf8');

function build() {
  let html = read('Index.html')
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
    makeSheet('Student Board', [['800#', 'Name', 'Squad', 'Grade', 'Trait Dough Bills', 'Bills This Week', 'Trait Dough Wins', 'Mindset Summits', 'Trail Journal Growth', 'STAR Growth Points', 'Shout-outs', 'Streak (weeks)', 'Badges', 'Last Contribution', 'Updated'],
      ['8001234', 'Ann L.', 'Gold', '7', '12', '2', '1', '1', '', '4', '2', '3', '🍞 First Trait Dough, 🍞 Trait Dough Ten, 🏆 Trait Dough Winner, ⛰️ Mindset Summit, 📈 STAR Grower, 🙌 Shout-out, 🔥 3-Week Streak', '2026-10-30', '2026-10-31 19:00'],
      ['8005678', 'Bo K.', 'Grey', '8', '3', '0', '0', '0', '', '2', '0', '0', '🍞 First Trait Dough, 📈 STAR Grower', '2026-10-20', '2026-10-31 19:00']]),
    makeSheet('Challenges', [['Show', 'Title', 'Description', 'Event Types', 'Measure', 'Goal', 'Start', 'End', 'Bonus Points', 'Awarded'],
      [true, 'Trait Dough Week', 'Turn in Trait Dough bills all week!', 'trait_dough', 'Points', 25, '2026-10-26', '2099-12-31', 50, ''],
      [true, 'Growth Push', 'Every point from any project counts.', '', 'Points', 200, '2026-09-01', '2099-12-31', 0, '']]),
  ];
  var book = { getSheets: function () { return sheets; }, getSheetByName: function (n) { return sheets.filter(function (s) { return s.getName() === n; })[0] || null; },
    insertSheet: function (n) { var s = makeSheet(n, []); sheets.push(s); return s; }, getId: function () { return 'CLASH'; } };
  var cache = {}, props = {}, updated = Date.now();
  window.SpreadsheetApp = { getActiveSpreadsheet: function () { return book; }, openById: function () { return { getSheetByName: function (n) {
    return n === 'reflections' ? makeSheet(n, [['insight_level', 'created_at'], ['High', new Date()], ['Medium', new Date()]]) : n === 'panel_sessions' ? makeSheet(n, [['clash_squad', 'status'], ['GOLD', 'complete']]) : null; } }; } };
  window.DriveApp = { getFileById: function () { return { getLastUpdated: function () { return new Date(updated); } }; } };
  window.CacheService = { getScriptCache: function () { return { get: function (k) { return cache[k] || null; }, put: function (k, v) { cache[k] = v; } }; } };
  window.PropertiesService = { getScriptProperties: function () { return { getProperty: function (k) { return props[k] || null; }, setProperty: function (k, v) { props[k] = v; } }; } };
  window.LockService = { getScriptLock: function () { return { waitLock: function () {}, releaseLock: function () {} }; } };
  window.Utilities = { getUuid: function () { return 'u' + Math.random().toString(36).slice(2); },
    computeDigest: function (a, s) { return s.split('').map(function (ch) { return ch.charCodeAt(0) % 128; }); }, // preview only
    DigestAlgorithm: {}, Charset: {}, formatDate: function (d) { return d.toISOString().slice(0, 10); } };
  window.Session = { getScriptTimeZone: function () { return 'America/New_York'; } };
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
getSheet_('Student Board').appendRow(['8009999', 'Gus T.', 'Gold', '7', '1', '0', '0', '0', '', '0', '0', '0', '', '', '']);
ensurePollTabs_(); getSheet_('Daily Polls').rows.slice(1).forEach(function (r) { r[0] = true; });
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
  return html.replace('<head>', '<head>\n' + fake);
}

const out = path.join(__dirname, 'preview.html');
fs.writeFileSync(out, build());
http.createServer((req, res) => {
  fs.writeFileSync(out, build());
  res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
  res.end(fs.readFileSync(out));
}).listen(3041, () => console.log('Preview on http://localhost:3041'));
