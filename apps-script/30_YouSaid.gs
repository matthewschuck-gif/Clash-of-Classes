/**
 * Clash of Classes -- Apps Script backend
 * File 30: You said, we did. Students see that answering a Pulse question changes something.
 *
 * YOU SAID WE DID tab (staff edit): Show | Poll | Closed | Question | What You Said | What We Did | Who
 *   - Once a day, every Pulse question that closed with at least YOU_SAID_MIN (10) answers gets a
 *     row by itself: the question and what students said most (counts only, never names).
 *   - Staff type What We Did (and Who, e.g. "the cafeteria team") and tick Show.
 *   - Shown rows (with What We Did filled in) appear on the Pulse page, the TV and the monthly
 *     State of the Squads email lists the ones still waiting for an answer.
 */

const YOU_SAID_HEADERS_ = ['Show', 'Poll', 'Closed', 'Question', 'What You Said', 'What We Did', 'Who'];

function ensureYouSaidTab_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getSheetByName('You Said We Did')) return;
  const sh = ss.insertSheet('You Said We Did');
  sh.getRange(1, 1, 1, YOU_SAID_HEADERS_.length).setValues([YOU_SAID_HEADERS_]).setFontWeight('bold').setBackground('#490e6f').setFontColor('#ffffff');
  sh.setFrozenRows(1);
  try { sh.getRange(2, 1, 300, 1).insertCheckboxes(); } catch (e) { console.warn('You Said checkboxes: ' + e.message); }
}

// 'Most said "Lunch" (41%), then "After school" (30%) · 212 answers'
function youSaidLine_(p) {
  const order = p.choices.slice().sort(function (a, b) { return (p.counts[b] || 0) - (p.counts[a] || 0); });
  const pct = function (c) { return p.total ? Math.round(100 * (p.counts[c] || 0) / p.total) : 0; };
  return 'Most said "' + order[0] + '" (' + pct(order[0]) + '%)' + (order[1] ? ', then "' + order[1] + '" (' + pct(order[1]) + '%)' : '') +
    ' · ' + p.total + (p.total === 1 ? ' answer' : ' answers');
}

/** Adds a row for each closed Pulse with enough answers. Returns how many were added. */
function syncYouSaid_() {
  ensureYouSaidTab_();
  const have = {};
  sheetRows_('You Said We Did').forEach(function (r) { have[String(r['Poll'])] = 1; });
  const min = settingNum_('YOU_SAID_MIN', 10), today = todayYmd_();
  const add = pollsWithVotes_().filter(function (p) {
    return p.kind === 'Pulse' && p.question && p.through < today && p.total >= min && !have[p.id];
  });
  const sh = getSheet_('You Said We Did');
  const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(function (h) { return String(h).trim(); });
  add.forEach(function (p) {
    const o = { 'Show': false, 'Poll': p.id, 'Closed': p.through, 'Question': p.question, 'What You Said': youSaidLine_(p), 'What We Did': '', 'Who': '' };
    sh.appendRow(head.map(function (h) { return o[h] === undefined ? '' : o[h]; }));
  });
  return add.length;
}

function maybeSyncYouSaid_() {
  const props = PropertiesService.getScriptProperties(), today = todayYmd_();
  if (props.getProperty('you-said') === today) return;
  props.setProperty('you-said', today);
  syncYouSaid_();
}

/** The rows students may see: Show ticked and What We Did filled in, newest first. */
function youSaidPublic_(limit) {
  return sheetRows_('You Said We Did').filter(function (r) { return ticked_(r['Show']) && String(r['What We Did'] || '').trim() && String(r['Question'] || '').trim(); })
    .map(function (r) {
      return { closed: String(r['Closed'] || '').slice(0, 10), question: String(r['Question']).trim(), said: String(r['What You Said'] || '').trim(),
        did: String(r['What We Did']).trim(), who: String(r['Who'] || '').trim() };
    })
    .sort(function (a, b) { return b.closed.localeCompare(a.closed); }).slice(0, limit || 6);
}

/** Rows still waiting for What We Did (for the monthly email). */
function youSaidWaiting_() {
  return sheetRows_('You Said We Did').filter(function (r) { return String(r['Question'] || '').trim() && !String(r['What We Did'] || '').trim(); })
    .map(function (r) { return { closed: String(r['Closed'] || '').slice(0, 10), question: String(r['Question']).trim(), said: String(r['What You Said'] || '').trim() }; });
}

function clashYouSaid() { return JSON.stringify(youSaidPublic_()); }

function menuSyncYouSaid_() {
  const n = syncYouSaid_(), waiting = youSaidWaiting_().length;
  SpreadsheetApp.getUi().alert('You said, we did', (n ? n + ' new Pulse result' + (n === 1 ? '' : 's') + ' added. ' : 'No new Pulse results. ') +
    waiting + ' waiting for What We Did on the You Said We Did tab. Type what changed (and Who), then tick Show: it appears on the Pulse page and the TV.',
    SpreadsheetApp.getUi().ButtonSet.OK);
}
