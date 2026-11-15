/**
 * Clash of Classes -- Apps Script backend
 * File 53: two more family pieces, each with its own switch (Admin › ✅ Service, OFF to start).
 *
 * FAMILY_COLORS  Squad Colors Day (SQUAD COLORS DAYS tab: Date | Name | BLACK | GOLD | GREY | PURPLE |
 *   Points | Paid). The next one shows on the family update (so families can plan), the home page
 *   and the morning announcement. After the day, type how many wore their colors (or the percent) under
 *   each squad. The squad with the most gets Points (COLORS_DAY_POINTS when blank); ties share.
 * FAMILY_PICK  Family Pick of the Month (FAMILY PICK tab: Show | Month | Question | Choices | Winner).
 *   Families vote once (with their student's name and grade, like the Pulse) on a small school-wide
 *   reward. Families see the counts; no names. Each vote adds FAMILY_PICK_POINTS to the student's
 *   squad (paid once a day). When the month ends the site writes the Winner. Votes: FAMILY PICK VOTES.
 */

const COLORS_HEADERS_ = ['Date', 'Name', 'BLACK', 'GOLD', 'GREY', 'PURPLE', 'Points', 'Paid'];
const PICK_HEADERS_ = ['Show', 'Month', 'Question', 'Choices', 'Winner'];
const PICK_VOTE_HEADERS_ = ['Month', 'Student', 'Grade', 'Squad', '800#', 'Choice', 'At', 'Paid'];

function ensureFamilyMoreTabs_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  [['Squad Colors Days', COLORS_HEADERS_, ['', 'Example: Squad Colors Day (set a date)', '', '', '', '', '', '']],
   ['Family Pick', PICK_HEADERS_, [false, todayYmd_().slice(0, 7), 'Which reward should the whole school get this month?', 'Extra recess | Music at lunch | Hat day', '']],
   ['Family Pick Votes', PICK_VOTE_HEADERS_, null]].forEach(function (t) {
    if (ss.getSheetByName(t[0])) return;
    const sh = ss.insertSheet(t[0]);
    sh.getRange(1, 1, 1, t[1].length).setValues([t[1]]).setFontWeight('bold').setBackground('#490e6f').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    if (t[2]) sh.appendRow(t[2]);
  });
}

// ── Squad Colors Day ──
function colorsDayNext_() {
  if (!familyOn_('FAMILY_COLORS')) return null;
  const today = todayYmd_();
  const r = sheetRows_('Squad Colors Days').map(function (x) { return { date: asYmd_(x['Date']).slice(0, 10), name: String(x['Name'] || '').trim() || 'Squad Colors Day' }; })
    .filter(function (x) { return x.date && x.date >= today && x.date <= ymdShift_(today, 30); }).sort(function (a, b) { return a.date.localeCompare(b.date); })[0];
  return r || null;
}

/** Pays colors days that are over and have counts. Returns how many it paid. */
function payColorsDays_() {
  const sh = getSheetSafe_('Squad Colors Days');
  if (!sh) return 0;
  const today = todayYmd_(), paidCol = headerCol_(sh, 'Paid');
  let n = 0;
  sheetRows_('Squad Colors Days').forEach(function (r) {
    const d = asYmd_(r['Date']).slice(0, 10);
    if (!d || d >= today || String(r['Paid'] || '').trim()) return;
    const c = {}; SQUAD_KEYS_.forEach(function (s) { c[s] = Number(String(r[s]).replace('%', '')) || 0; });
    const best = Math.max.apply(null, SQUAD_KEYS_.map(function (s) { return c[s]; }));
    if (!best) return; // counts not in yet
    const pts = Number(r['Points']) || settingNum_('COLORS_DAY_POINTS', 25), row = { BLACK: '', GOLD: '', GREY: '', PURPLE: '' };
    SQUAD_KEYS_.forEach(function (s) { if (c[s] === best) row[s] = pts; });
    const name = String(r['Name'] || '').trim() || 'Squad Colors Day';
    getSheet_('events').appendRow(['', name + ': most squad colors', 'colors_day', d, SQUAD_KEYS_.filter(function (s) { return row[s]; }).map(titleSquad_).join(' + ') + ' wore it best', '', row.BLACK, row.GOLD, row.GREY, row.PURPLE]);
    if (paidCol) sh.getRange(r._row, paidCol).setValue(today);
    n++;
  });
  return n;
}

// ── Family Pick of the Month ──
function familyPick_() {
  if (!familyOn_('FAMILY_PICK')) return null;
  const ym = todayYmd_().slice(0, 7);
  const r = sheetRows_('Family Pick').filter(function (x) {
    const m = monthKey_(x['Month']);
    return ticked_(x['Show']) && String(x['Question'] || '').trim() && pollChoices_(x['Choices']).length >= 2 && (!m || m === ym);
  })[0];
  const last = sheetRows_('Family Pick').filter(function (x) { return String(x['Winner'] || '').trim() && monthKey_(x['Month']) < ym; })
    .sort(function (a, b) { return monthKey_(b['Month']).localeCompare(monthKey_(a['Month'])); })[0];
  const out = { lastWinner: last ? String(last['Winner']).trim() : '' };
  if (!r) return out.lastWinner ? out : null;
  const month = monthKey_(r['Month']) || ym, choices = pollChoices_(r['Choices']), counts = {};
  choices.forEach(function (c) { counts[c] = 0; });
  const votes = sheetRows_('Family Pick Votes').filter(function (v) { return monthKey_(v['Month']) === month; });
  votes.forEach(function (v) { const c = String(v['Choice']); if (counts[c] !== undefined) counts[c]++; });
  out.month = month; out.question = String(r['Question']).trim(); out.choices = choices; out.counts = counts; out.total = votes.length;
  return out;
}

function monthKey_(v) {
  if (Object.prototype.toString.call(v) === '[object Date]') return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM');
  const m = String(v || '').trim().match(/^(\d{4})-(\d{1,2})/);
  return m ? m[1] + '-' + ('0' + m[2]).slice(-2) : '';
}

function clashFamilyPickVote(json) {
  const fp = familyPick_();
  if (!fp || !fp.question) throw new Error('There is no Family Pick open right now.');
  const p = JSON.parse(json || '{}');
  if (fp.choices.indexOf(String(p.choice)) === -1) throw new Error('Pick one.');
  const st = findStudentByName_(p.first, p.last, p.grade);
  if (!st) throw new Error('We couldn\'t find that student. Check the spelling of their first and last name and their grade.');
  familyLimit_('fam-pick-all', 300, 3600);
  ensureFamilyMoreTabs_();
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (sheetRows_('Family Pick Votes').some(function (v) { return monthKey_(v['Month']) === fp.month && id800_(v['800#']) === st.id; })) throw new Error('Your family already voted this month. Thank you!');
    getSheet_('Family Pick Votes').appendRow([fp.month, st.first + ' ' + st.last, st.grade, st.squad, st.id, String(p.choice), Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'), '']);
  } finally { lock.releaseLock(); }
  return JSON.stringify(familyPick_());
}

/** Daily: vote points (one row per squad) and last month's Winner. */
function payFamilyPick_() {
  const vs = getSheetSafe_('Family Pick Votes'), today = todayYmd_(), ym = today.slice(0, 7);
  let n = 0;
  if (vs) {
    const rows = sheetRows_('Family Pick Votes').filter(function (r) { return !String(r['Paid'] || '').trim() && SQUAD_KEYS_.indexOf(squadOf_(r['Squad'])) > -1; });
    const per = settingNum_('FAMILY_PICK_POINTS', 1);
    if (rows.length && per > 0) {
      const pts = { BLACK: 0, GOLD: 0, GREY: 0, PURPLE: 0 };
      rows.forEach(function (r) { pts[squadOf_(r['Squad'])] += per; });
      getSheet_('events').appendRow(['', 'Family Pick: ' + rows.length + ' famil' + (rows.length === 1 ? 'y' : 'ies') + ' voted', 'family', today, '', '', pts.BLACK || '', pts.GOLD || '', pts.GREY || '', pts.PURPLE || '']);
    }
    const c = headerCol_(vs, 'Paid');
    rows.forEach(function (r) { if (c) vs.getRange(r._row, c).setValue(today); });
    n = rows.length;
  }
  const ps = getSheetSafe_('Family Pick');
  if (ps) {
    const wc = headerCol_(ps, 'Winner'), votes = sheetRows_('Family Pick Votes');
    sheetRows_('Family Pick').forEach(function (r) {
      const m = monthKey_(r['Month']);
      if (!m || m >= ym || String(r['Winner'] || '').trim() || !wc) return;
      const counts = {};
      votes.filter(function (v) { return monthKey_(v['Month']) === m; }).forEach(function (v) { counts[String(v['Choice'])] = (counts[String(v['Choice'])] || 0) + 1; });
      const top = Object.keys(counts).sort(function (a, b) { return counts[b] - counts[a]; })[0];
      if (top) ps.getRange(r._row, wc).setValue(top);
    });
  }
  return n;
}

// file-stamp (tools/stamp-files.js): lets Admin › System tell whether this file is the newest one
var FILE_STAMP_53_FamilyMore = 'e1eb96a3d2';
