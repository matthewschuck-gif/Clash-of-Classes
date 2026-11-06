/**
 * Clash of Classes -- Apps Script backend
 * File 23: From the shared Clash drive -- Mindset lesson links and the Squad Generals.
 *
 * MINDSET LINKS tab: one row per month -- Month | 7th Grade | 8th Grade | Extension. Paste any
 * Drive / Docs / Slides link into a cell; a blank cell shows "No link yet" on the site. Seeded from
 * the shared drive's Monthly Themes docs (26-27): 7th grade uses the one Mindset Lessons doc all
 * year, 8th grade says "Mindset app" (= Settings › MINDSET_APP, which opens the month's self-paced
 * lesson), and Extension is the universal Choice Board (February and March have their own Extension
 * Activity docs). Any cell can be "Mindset app" or a link.
 * A link typed into Admin › Links › Mindset still wins over the tab for that month and grade.
 */

const MINDSET_LINK_HEADERS_ = ['Month', '7th Grade', '8th Grade', 'Extension', 'Notes'];
const MINDSET_LINK_COLS_ = [['7th', '7th Grade'], ['8th', '8th Grade'], ['ext', 'Extension']];

function mindsetLinkSeed_() {
  const doc = function (id) { return 'https://docs.google.com/document/d/' + id + '/edit?usp=sharing'; };
  const g7 = doc('1Xijj6wg-cHSjspOWHDGibhlafllkq9zA8v15vj3QWc8'); // 7th Grade Mindset Lessons (all months)
  const board = doc('1IWaNKNHbGdmnM-5bESYd4-vTRi_StiuQr3uZ2tqETeY'); // Universal Choice Board (print the month's page)
  const app = MINDSET_APP_WORD_; // 8th grade: the Mindset app's self-paced lesson for the month
  return [
    ['September', g7, app, board, ''],
    ['October', g7, app, board, ''],
    ['November', g7, app, board, ''],
    ['December', g7, app, board, ''],
    ['January', g7, app, board, ''],
    ['February', g7, app, doc('1yVoCQfyx4V4rPGeCZk-K0Zmz618sSrYJNYuLYQY2aAg'), ''],
    ['March', g7, app, doc('1oqBWenKjNKPq8zn7FYdU5HU92nxuGlSVnnhwXWIMIr8'), ''],
    ['April', g7, app, board, ''],
    ['May', g7, app, board, ''],
  ];
}
const MINDSET_APP_WORD_ = 'Mindset app';

function ensureMindsetLinksTab_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getSheetByName('Mindset Links')) return;
  const sh = ss.insertSheet('Mindset Links');
  sh.getRange(1, 1, 1, MINDSET_LINK_HEADERS_.length).setValues([MINDSET_LINK_HEADERS_]).setFontWeight('bold').setBackground('#490e6f').setFontColor('#ffffff');
  const seed = mindsetLinkSeed_();
  sh.getRange(2, 1, seed.length, MINDSET_LINK_HEADERS_.length).setValues(seed);
  sh.setFrozenRows(1);
}

// [{key:'7th-September', label, url}, …] from the Mindset Links tab.
function mindsetLinksFromTab_() {
  const out = [], app = String(getSetting_('MINDSET_APP') || '').trim();
  sheetRows_('Mindset Links').forEach(function (r) {
    const month = String(r['Month'] || '').trim();
    if (!month) return;
    MINDSET_LINK_COLS_.forEach(function (c) {
      let url = String(r[c[1]] || '').trim();
      if (url.toLowerCase() === MINDSET_APP_WORD_.toLowerCase()) url = app;
      if (/^https?:\/\//i.test(url)) out.push({ key: c[0] + '-' + month, label: month + ' — ' + c[1], url: url, fromSheet: true });
    });
  });
  return out;
}

// The site's saved Mindset links plus the tab's, for any month and grade the site has no link for.
function mergeMindsetLinks_(value) {
  if (!value.links || typeof value.links !== 'object') value.links = {};
  // Copies of the tab's links that came back in a site save are dropped, so the tab stays in charge.
  const saved = (Array.isArray(value.links.mindset) ? value.links.mindset : []).filter(function (l) { return l && !l.fromSheet; });
  const have = {};
  saved.forEach(function (l) { if (l && l.key && l.url) have[l.key] = true; });
  let tab = [];
  try { tab = mindsetLinksFromTab_(); } catch (e) { /* the page shows "No link yet" */ }
  value.links.mindset = saved.concat(tab.filter(function (l) { return !have[l.key]; }));
}

// ── SQUAD GENERALS (MLC Roster) ──────────────────────────────────────
// Menu › Load Squad Generals from the Squad Leadership doc. The doc (Settings › GENERALS_DOC) is
// one table: a header row of squads (Black Squad | Gold Squad | …) and a cell per squad listing
// each General as a line with their name, then bullets (first bullet = homeroom teacher, the rest
// = activities). Each name is matched to the Students tab for the 800# Generals HQ sign-in needs.
// Rows already on MLC Roster are kept; a General is added once (by 800#, or by name if not found).
function generalsFromDoc_(docId) {
  const body = DocumentApp.openById(docId).getBody();
  const out = [];
  body.getTables().forEach(function (table) {
    if (table.getNumRows() < 2) return;
    const head = table.getRow(0);
    const squads = [];
    for (let c = 0; c < head.getNumCells(); c++) squads.push(squadOf_(String(head.getCell(c).getText()).replace(/squad/i, '').trim()));
    for (let r = 1; r < table.getNumRows(); r++) {
      const row = table.getRow(r);
      for (let c = 0; c < row.getNumCells() && c < squads.length; c++) {
        if (SQUAD_KEYS_.indexOf(squads[c]) === -1) continue;
        const cell = row.getCell(c);
        let cur = null;
        for (let i = 0; i < cell.getNumChildren(); i++) {
          const el = cell.getChild(i);
          const text = String(el.asText ? el.asText().getText() : '').replace(/\*/g, '').trim();
          if (!text) continue;
          if (el.getType() === DocumentApp.ElementType.LIST_ITEM) {
            if (!cur) continue;
            if (!cur.homeroom) cur.homeroom = text; else cur.notes.push(text);
          } else {
            cur = { name: text, squad: squads[c], homeroom: '', notes: [] };
            out.push(cur);
          }
        }
      }
    }
  });
  return out;
}

function importGenerals_() {
  const docId = extractSheetId_(String(getSetting_('GENERALS_DOC') || ''));
  if (!docId) throw new Error('Add the Squad Leadership doc link to Settings › GENERALS_DOC first.');
  const found = generalsFromDoc_(docId);
  if (!found.length) throw new Error('No Generals found in that doc. It needs a table with a squad in each column header.');
  ensureGeneralsTabs_();
  const students = sheetRows_(ROSTER_TAB_);
  const key = function (s) { return String(s || '').toLowerCase().replace(/[^a-z]/g, ''); };
  const byName = {};
  students.forEach(function (s) { byName[key(String(s['First Name'] || '') + String(s['Last Name'] || ''))] = s; });
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('MLC Roster');
  const have = sheetRows_('MLC Roster');
  const added = [], missing = [];
  found.forEach(function (g) {
    const s = byName[key(g.name)];
    const id = s ? id800_(s['800#']) : '';
    const already = have.some(function (r) { return /general/i.test(String(r['Role'])) && (id ? id800_(r['800#']) === id : key(r['Name']) === key(g.name)); });
    if (already) return;
    sh.appendRow([id, g.name, g.squad, 'General', '']);
    added.push(g.name);
    if (!id) missing.push(g.name + ' (' + g.squad.charAt(0) + g.squad.slice(1).toLowerCase() + ')');
  });
  return { found: found.length, added: added.length, missing: missing };
}

function menuImportGenerals_() {
  const ui = SpreadsheetApp.getUi();
  try {
    const r = importGenerals_();
    ui.alert('Squad Generals: ' + r.found + ' in the doc, ' + r.added + ' added to MLC Roster.' +
      (r.missing.length ? '\n\nNo 800# found on the Students tab for:\n' + r.missing.join('\n') +
        '\n\nType their 800# into MLC Roster (column A) so they can sign in to Generals HQ.' : ''));
  } catch (e) { ui.alert(e.message); }
}
