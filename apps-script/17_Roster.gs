/**
 * Clash of Classes -- Apps Script backend
 * File 17: the student roster, kept in this Sheet (tab Students).
 *
 *   800# | First Name | Last Name | Grade | Email | Squad
 *
 * Clash of Classes › Import student roster copies 800#, email, first and last name and grade
 * from the district roster Sheet (Settings › ROSTER_SHEET, its link or ID). It runs inside your
 * Google account; the list is never sent anywhere else. Squad is kept for every 800# already on
 * the tab, and filled from Climb-er's Student Board when that has the student. Type any squad
 * that's missing (BLACK, GOLD, GREY or PURPLE).
 *
 * My Squad and My Climb use this tab for anyone the Student Board doesn't have yet, so a
 * student can sign in even before Climb-er's hourly sync has run. Students sign in to My Climb
 * with their school account; their email is their 800# + STUDENT_EMAIL_PATTERN's domain
 * ({800}@gomounts.org), so the email itself says who they are.
 */

const ROSTER_TAB_ = 'Students';
const ROSTER_HEADERS_ = ['800#', 'First Name', 'Last Name', 'Grade', 'Email', 'Squad'];

function ensureRosterTab_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(ROSTER_TAB_);
  if (sh) return sh;
  sh = ss.insertSheet(ROSTER_TAB_);
  sh.getRange(1, 1, 1, ROSTER_HEADERS_.length).setValues([ROSTER_HEADERS_]).setFontWeight('bold').setBackground('#490e6f').setFontColor('#ffffff');
  sh.setFrozenRows(1);
  sh.getRange('A:A').setNumberFormat('@'); // keep 800#s as typed
  sh.getRange('F2:F1000').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(SQUAD_KEYS_, true).setAllowInvalid(true).build());
  return sh;
}

function sheetIdFrom_(v) {
  const s = String(v || '').trim();
  const m = s.match(/\/d\/([a-zA-Z0-9_-]{20,})/);
  return m ? m[1] : s;
}

// Copies the district roster in. Returns { students, squads, missingSquad }.
function importRoster_() {
  // A blank ROSTER_SHEET row (made by an older setup) falls back to the built-in roster link.
  const id = sheetIdFrom_(getSetting_('ROSTER_SHEET')) || sheetIdFrom_((DEFAULT_SETTINGS.filter(function (d) { return d[0] === 'ROSTER_SHEET'; })[0] || [])[1]);
  if (!id) throw new Error('Put the roster Sheet\'s link in Settings › ROSTER_SHEET first.');
  const src = SpreadsheetApp.openById(id).getSheets()[0].getDataRange().getValues();
  const h = src[0].map(function (x) { return String(x).trim().toLowerCase(); });
  const col = function (words) { for (let i = 0; i < h.length; i++) if (words.some(function (w) { return h[i].indexOf(w) > -1; })) return i; return -1; };
  const cId = col(['800', 'student number', 'student id']), cEmail = col(['email']), cFirst = col(['first']), cLast = col(['last']), cGrade = col(['grade']), cSquad = col(['squad']);
  if (cId === -1) throw new Error('The roster Sheet needs an "800 Number" column.');
  const sh = ensureRosterTab_();
  // Squads we already know: this tab first, then Climb-er's Student Board.
  const known = {};
  sheetRows_(getSetting_('STUDENT_BOARD_TAB') || 'Student Board').forEach(function (r) { const s = squadOf_(r['Squad']); if (SQUAD_KEYS_.indexOf(s) > -1) known[id800_(r['800#'])] = s; });
  sheetRows_(ROSTER_TAB_).forEach(function (r) { const s = squadOf_(r['Squad']); if (SQUAD_KEYS_.indexOf(s) > -1) known[id800_(r['800#'])] = s; });
  const pattern = studentEmailPattern_();
  const rows = src.slice(1).filter(function (r) { return id800_(r[cId]); }).map(function (r) {
    const sid = id800_(r[cId]);
    const squad = cSquad > -1 && SQUAD_KEYS_.indexOf(squadOf_(r[cSquad])) > -1 ? squadOf_(r[cSquad]) : (known[sid] || '');
    const email = cEmail > -1 && String(r[cEmail]).trim() ? String(r[cEmail]).trim().toLowerCase() : (pattern ? pattern.replace('{800}', sid) : '');
    return [sid, cFirst > -1 ? String(r[cFirst]).trim() : '', cLast > -1 ? String(r[cLast]).trim() : '', cGrade > -1 ? String(r[cGrade]).trim() : '', email, squad];
  });
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, ROSTER_HEADERS_.length).clearContent();
  if (rows.length) sh.getRange(2, 1, rows.length, ROSTER_HEADERS_.length).setValues(rows);
  return { students: rows.length, squads: rows.filter(function (r) { return r[5]; }).length, missingSquad: rows.filter(function (r) { return !r[5]; }).length };
}

function menuImportRoster_() {
  const ui = SpreadsheetApp.getUi();
  // No roster link yet: ask for it once and remember it on Settings.
  if (!sheetIdFrom_(getSetting_('ROSTER_SHEET')) && !sheetIdFrom_((DEFAULT_SETTINGS.filter(function (d) { return d[0] === 'ROSTER_SHEET'; })[0] || [])[1])) {
    const ask = ui.prompt('Student roster', 'Paste the link to the roster Sheet (800 Number, Email Address, First Name, Last Name, Grade Level):', ui.ButtonSet.OK_CANCEL);
    if (ask.getSelectedButton() !== ui.Button.OK || !ask.getResponseText().trim()) return;
    setSettingValue_('ROSTER_SHEET', ask.getResponseText().trim());
  }
  try {
    const r = importRoster_();
    ui.alert('Students tab: ' + r.students + ' students copied.\n' + r.squads + ' have a squad.' +
      (r.missingSquad ? '\n' + r.missingSquad + ' still need one: type BLACK, GOLD, GREY or PURPLE in the Squad column (or let Climb-er\'s sync fill it).' : ''));
  } catch (e) { ui.alert(e.message); }
}

function studentEmailPattern_() {
  return String(getSetting_('STUDENT_EMAIL_PATTERN') || '{800}@gomounts.org').trim().toLowerCase();
}

// One student from the Students tab, shaped like a Student Board row (all counts 0), or null.
function rosterStudent_(sid) {
  const r = sheetRows_(ROSTER_TAB_).filter(function (x) { return id800_(x['800#']) === sid; })[0];
  if (!r) return null;
  const last = String(r['Last Name'] || '').trim();
  return { '800#': sid, 'Name': String(r['First Name'] || '').trim() + (last ? ' ' + last.charAt(0) + '.' : ''), 'Squad': squadOf_(r['Squad']),
    'Grade': String(r['Grade'] || ''), 'Email': String(r['Email'] || '').trim().toLowerCase() };
}

// Run from the Apps Script editor (pick importStudentRoster in the function list, then Run):
// makes the Students tab and fills it, with no menu needed. The result shows in the Execution log.
function importStudentRoster() {
  const r = importRoster_();
  console.log('Students tab: ' + r.students + ' students copied, ' + r.squads + ' with a squad, ' + r.missingSquad + ' still need one.');
  return r;
}
