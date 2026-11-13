/**
 * Clash of Classes -- Apps Script backend
 * File 27: the Cafeteria's LRR count, from the same sheets the EMS Follow-Up dashboard reads.
 *
 * The Cafeteria (Clean Lunch) is won by the squad with the fewest students who had an LRR or a
 * Lunch Detention last month. EMS Follow-Up already counts these, from:
 *   LRR_SHEET          the Lunchtime Redirection Reflection responses (its Form tab)
 *   CONSEQUENCE_SHEET  the Accountability & Intervention Log, tab CONSEQUENCE_TAB (Master Log):
 *                      rows whose Consequence Type has an LD_TYPES word (Lunch Detention)
 * This reads the same two sheets, finds each student's squad (Students tab by 800# or name, then
 * the Student Board by 800#), and writes one Squad Stats row a month:
 *   Month | "LRR + LD Students" | BLACK | GOLD | GREY | PURPLE | Source "Follow-Up: ..." | Updated
 * Counts only; no names leave the source sheets. Each student counts once a month.
 * Runs once a day with the background tick (FOLLOWUP_STATS ON) and from the menu. A Squad Stats
 * row staff typed (any other Source) is never overwritten; Climb-er's row for this stat is.
 */

const FOLLOWUP_STAT_ = 'LRR + LD Students';

// 'yyyy-MM-dd' from a cell: a date, yyyy-mm-dd..., or m/d/yyyy ... ('' if none).
function ymdOfCell_(v) {
  if (Object.prototype.toString.call(v) === '[object Date]') return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const s = String(v || '').trim();
  let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2);
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})/);
  if (m) return (m[3].length === 2 ? '20' + m[3] : m[3]) + '-' + ('0' + m[1]).slice(-2) + '-' + ('0' + m[2]).slice(-2);
  return '';
}

function nameKeyOf_(first, last) {
  const f = String(first || '').trim().toLowerCase().split(/\s+/)[0] || '', l = String(last || '').trim().toLowerCase().replace(/[^a-z'-]/g, '');
  return f && l ? f + '|' + l : '';
}
function nameKeyOfFull_(full) {
  const s = String(full || '').trim();
  if (!s) return '';
  if (s.indexOf(',') > -1) { const p = s.split(','); return nameKeyOf_(p[1], p[0]); }
  const p = s.split(/\s+/);
  return p.length > 1 ? nameKeyOf_(p[0], p[p.length - 1]) : '';
}

// { byId: { 800#: SQUAD }, byName: { 'first|last': SQUAD } } (a name two students share is left out).
function squadLookup_() {
  const byId = {}, byName = {}, bad = {};
  sheetRows_('Students').forEach(function (r) {
    const s = squadOf_(r['Squad']), id = id800_(r['800#']);
    if (SQUAD_KEYS_.indexOf(s) === -1) return;
    if (id) byId[id] = s;
    const k = nameKeyOf_(r['First Name'], r['Last Name']);
    if (!k || bad[k]) return;
    if (byName[k] && byName[k] !== s) { bad[k] = 1; delete byName[k]; return; }
    byName[k] = s;
  });
  sheetRows_(getSetting_('STUDENT_BOARD_TAB') || 'Student Board').forEach(function (r) {
    const s = squadOf_(r['Squad']), id = id800_(r['800#']);
    if (id && !byId[id] && SQUAD_KEYS_.indexOf(s) > -1) byId[id] = s;
  });
  return { byId: byId, byName: byName };
}

function openRows_(idOrUrl, tabName) {
  const id = String(idOrUrl || '').match(/[-\w]{25,}/);
  if (!id) return [];
  const book = SpreadsheetApp.openById(id[0]);
  let sh = tabName ? book.getSheetByName(tabName) : null;
  if (!sh && !tabName) sh = book.getSheets().filter(function (s) { return s.getFormUrl && s.getFormUrl(); })[0] || book.getSheets()[0];
  if (!sh || sh.getLastRow() < 2) return [];
  const v = sh.getDataRange().getValues(), h = v[0].map(function (x) { return String(x).trim(); });
  return v.slice(1).map(function (r) { const o = {}; h.forEach(function (k, j) { o[k] = r[j]; }); return o; });
}

function pick_(row, re) { const k = Object.keys(row).filter(function (h) { return re.test(h); })[0]; return k ? row[k] : ''; }

/** { month, counts: {BLACK..}, lrr, ld, unmatched } for one month (default last month). */
function followUpLrrCounts_(month) {
  month = month || lastMonthWindow_(todayYmd_()).month;
  const look = squadLookup_(), seen = {}, counts = { BLACK: 0, GOLD: 0, GREY: 0, PURPLE: 0 };
  const out = { month: month, counts: counts, lrr: 0, ld: 0, unmatched: 0 };
  const add = function (id, nameKey, day, kind) {
    if (String(day).slice(0, 7) !== month) return;
    out[kind]++;
    const s = (id && look.byId[id]) || (nameKey && look.byName[nameKey]) || '';
    if (!s) { out.unmatched++; return; }
    const who = id ? 'id:' + id : 'n:' + nameKey;
    if (seen[who]) return;
    seen[who] = 1; counts[s]++;
  };
  openRows_(getSetting_('LRR_SHEET'), '').forEach(function (r) {
    const email = String(pick_(r, /e-?mail/i) || ''), m = email.match(/^(\d{5,})@/);
    const id = m ? m[1] : id800_(pick_(r, /800/));
    add(id, nameKeyOf_(pick_(r, /^first/i), pick_(r, /^last/i)), ymdOfCell_(r['Timestamp']), 'lrr');
  });
  const ldWords = listOf_(getSetting_('LD_TYPES') || 'lunch detention');
  const skip = listOf_(getSetting_('SKIP_CONSEQUENCE_STATUSES') || 'cancelled, canceled, rescheduled');
  openRows_(getSetting_('CONSEQUENCE_SHEET'), getSetting_('CONSEQUENCE_TAB') || 'Master Log').forEach(function (r) {
    const type = String(r['Consequence Type'] || '').toLowerCase();
    if (!ldWords.some(function (w) { return type.indexOf(w) > -1; })) return;
    if (skip.indexOf(String(r['Status'] || '').trim().toLowerCase()) > -1) return;
    add(id800_(r['Student 800#']), nameKeyOfFull_(r['Student Name']), ymdOfCell_(r['Consequence Date'] || r['Entry Date']), 'ld');
  });
  return out;
}

/** Writes the month's row onto Squad Stats. Returns one line saying what happened. */
function sendFollowUpStats_(month) {
  ensureConquestTabs_();
  const c = followUpLrrCounts_(month);
  const sh = getSheet_('Squad Stats');
  const rows = sheetRows_('Squad Stats');
  const mine = rows.filter(function (r) { return String(r['Month'] || '').slice(0, 7) === c.month && String(r['Stat'] || '').trim() === FOLLOWUP_STAT_; })[0];
  const source = 'Follow-Up: ' + c.lrr + ' LRRs + ' + c.ld + ' LDs' + (c.unmatched ? ', ' + c.unmatched + ' with no squad found' : '');
  if (mine && !/^(follow-up|climb-er)/i.test(String(mine['Source'] || ''))) return FOLLOWUP_STAT_ + ' ' + c.month + ': kept the row staff typed.';
  const obj = { 'Month': c.month, 'Stat': FOLLOWUP_STAT_, 'BLACK': c.counts.BLACK, 'GOLD': c.counts.GOLD, 'GREY': c.counts.GREY, 'PURPLE': c.counts.PURPLE,
    'Source': source, 'Updated': Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm') };
  const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(function (h) { return String(h).trim(); });
  const vals = head.map(function (h) { return obj[h] === undefined ? '' : obj[h]; });
  if (mine) sh.getRange(mine._row, 1, 1, vals.length).setValues([vals]); else sh.appendRow(vals);
  return FOLLOWUP_STAT_ + ' ' + c.month + ': ' + SQUAD_KEYS_.map(function (s) { return s + ' ' + c.counts[s]; }).join(', ') + ' (' + source + ').';
}

// Once a day from the background tick.
function maybeFollowUpStats_() {
  if (String(getSetting_('FOLLOWUP_STATS') || 'ON').toUpperCase() === 'OFF' || !getSetting_('LRR_SHEET')) return;
  const props = PropertiesService.getScriptProperties(), today = todayYmd_();
  if (props.getProperty('followup-stats') === today) return;
  props.setProperty('followup-stats', today);
  sendFollowUpStats_();
}

function menuFollowUpStats_() {
  const last = sendFollowUpStats_();
  const now = followUpLrrCounts_(todayYmd_().slice(0, 7));
  SpreadsheetApp.getUi().alert('Cafeteria (fewest LRR + LD students)', last + '\n\nSo far this month: ' +
    SQUAD_KEYS_.map(function (s) { return s + ' ' + now.counts[s]; }).join(', ') + ' (' + now.lrr + ' LRRs, ' + now.ld + ' LDs' +
    (now.unmatched ? ', ' + now.unmatched + ' with no squad found' : '') + ').', SpreadsheetApp.getUi().ButtonSet.OK);
}
