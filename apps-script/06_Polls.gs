/**
 * Clash of Classes -- Apps Script backend
 * File 6: Daily Polls for the morning announcements (This or That, Trivia, School Pulse, Check-in).
 *
 * Daily Polls tab (one row per poll, filled in by staff):
 *   Show | Date | Kind | Question | Choices | Answer | Participation Points | Win Points | Status | Results
 *   Kind: This or That (the choice with the most votes wins; a tie means every tied choice
 *         wins) or Trivia (Answer is the right choice). Choices separated by | (Coke | Pepsi).
 *         School Pulse: a question the school wants answered (how students get to school, sleep,
 *         homework help...). No winner; every answer earns the participation point. Staff see the
 *         answers by squad and grade on the Poll Insights tab.
 *         Check-in: one button (Choices may be blank). Pride Day Check-in, Wheel Check-in ... A
 *         Kind with "pride" or "wheel" in it counts for that map zone (18_Conquest.gs check-ins).
 *   Ideas for all of these are on the Poll Ideas tab (26_PollIdeas.gs).
 *   Date: the school day it runs (open all that day). Status and Results are filled in.
 * Poll Votes tab: one row per vote (Poll | 800# | Squad | Choice | At). One vote per student
 * per poll; the squad comes from the Student Board.
 *
 * Points (when DAILY_POLL_POINTS is ON, checked hourly once the menu turns it on): after a
 * poll's day, each squad gets Participation Points for every student who voted plus Win Points
 * for every student who picked the winner, as one "Daily Poll" event (type daily_poll). Never
 * twice: Status becomes "Awarded".
 */

const POLL_HEADERS_ = ['Show', 'Date', 'Kind', 'Question', 'Choices', 'Answer', 'Participation Points', 'Win Points', 'Status', 'Results'];
const VOTE_HEADERS_ = ['Poll', 'Date', '800#', 'Squad', 'Choice', 'At'];

function ensurePollTabs_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  [['Daily Polls', POLL_HEADERS_, [[false, today, 'This or That', 'Coke or Pepsi?', 'Coke | Pepsi', '', 1, 1, '', ''],
    [false, today, 'Trivia', 'How many squads are in Clash of Classes?', '3 | 4 | 5', '4', 1, 1, '', '']]],
   ['Poll Votes', VOTE_HEADERS_, null]].forEach(function (t) {
    if (ss.getSheetByName(t[0])) return;
    const sh = ss.insertSheet(t[0]);
    sh.getRange(1, 1, 1, t[1].length).setValues([t[1]]).setFontWeight('bold').setBackground('#490e6f').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    if (t[2]) sh.getRange(2, 1, t[2].length, t[1].length).setValues(t[2]);
  });
}

function pollChoices_(v) { return String(v || '').split('|').map(function (s) { return s.trim(); }).filter(String); }
const CHECK_IN_CHOICE_ = 'I\'m in! ✋';
// The four kinds the site knows. The Kind cell can say more ("Pride Day Check-in").
function pollKind_(v) {
  const k = String(v || '');
  if (/trivia/i.test(k)) return 'Trivia';
  if (/check|pride|wheel/i.test(k)) return 'Check-in';
  if (/pulse|survey|school/i.test(k)) return 'School Pulse';
  return 'This or That';
}
function todayYmd_() { return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'); }

// Every poll row, with its votes tallied.
function pollsWithVotes_() {
  const votes = sheetRows_('Poll Votes');
  return sheetRows_('Daily Polls').map(function (r) {
    const id = String(r._row);
    const mine = votes.filter(function (v) { return String(v['Poll']) === id; });
    const counts = {}, bySquad = {};
    const kind = pollKind_(r['Kind']);
    let choices = pollChoices_(r['Choices']);
    if (kind === 'Check-in' && !choices.length) choices = [CHECK_IN_CHOICE_];
    choices.forEach(function (c) { counts[c] = 0; });
    mine.forEach(function (v) {
      const c = String(v['Choice']);
      counts[c] = (counts[c] || 0) + 1;
      const s = squadOf_(v['Squad']);
      bySquad[s] = bySquad[s] || {};
      bySquad[s][c] = (bySquad[s][c] || 0) + 1;
    });
    return {
      id: id, row: r._row, show: ticked_(r['Show']), date: String(r['Date'] || ''), kind: kind, kindLabel: String(r['Kind'] || '').trim() || kind,
      question: String(r['Question'] || ''), choices: choices, answer: String(r['Answer'] || '').trim(),
      partPts: Number(r['Participation Points']) || 0, winPts: Number(r['Win Points']) || 0,
      status: String(r['Status'] || ''), counts: counts, bySquad: bySquad, total: mine.length, votes: mine,
    };
  });
}

function pollWinners_(p) {
  if (p.kind === 'Trivia') return p.answer ? [p.answer] : [];
  if (p.kind === 'School Pulse' || p.kind === 'Check-in') return []; // nobody wins: every answer counts the same
  const max = Math.max.apply(null, p.choices.map(function (c) { return p.counts[c] || 0; }).concat([0]));
  return max ? p.choices.filter(function (c) { return (p.counts[c] || 0) === max; }) : [];
}

// What a student sees (no 800#s, no other students): today's polls, results once they voted.
function publicPoll_(p, votedChoice) {
  const closed = p.date < todayYmd_();
  return {
    id: p.id, kind: p.kind, kindLabel: p.kindLabel, question: p.question, choices: p.choices, date: p.date,
    voted: votedChoice || '', closed: closed,
    counts: votedChoice || closed ? p.counts : null, total: p.total,
    answer: closed && p.kind === 'Trivia' ? p.answer : '', // trivia answer only after the day
    winners: closed ? pollWinners_(p) : [],
  };
}

/** Today's polls (and yesterday's results) for the site. 800# optional: marks what they voted. */
function clashDailyPolls(id) {
  const sid = id800_(id);
  const today = todayYmd_();
  const polls = pollsWithVotes_().filter(function (p) { return p.show && p.question && p.choices.length >= (p.kind === 'Check-in' ? 1 : 2); });
  const voteOf = function (p) {
    const v = sid ? p.votes.filter(function (x) { return id800_(x['800#']) === sid; })[0] : null;
    return v ? String(v['Choice']) : '';
  };
  const past = polls.filter(function (p) { return p.date < today; }).sort(function (a, b) { return b.date.localeCompare(a.date); });
  return JSON.stringify({
    today: polls.filter(function (p) { return p.date === today; }).map(function (p) { return publicPoll_(p, voteOf(p)); }),
    last: past.length ? publicPoll_(past[0], voteOf(past[0])) : null,
  });
}

/** One vote: 800# on the Student Board, today's poll, one vote per poll. */
function clashPollVote(pollId, id, choice) {
  const cache = CacheService.getScriptCache();
  const used = Number(cache.get('poll-votes')) || 0;
  if (used >= (Number(getSetting_('POLL_VOTES_PER_10MIN')) || 1500)) throw new Error('So many votes right now! Try again in a minute.');
  cache.put('poll-votes', String(used + 1), 600);
  const sid = id800_(id);
  if (sid.length < 5) throw new Error('Type your 800 number.');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const p = pollsWithVotes_().filter(function (x) { return x.id === String(pollId); })[0];
    if (!p || !p.show) throw new Error('That poll isn\'t open.');
    if (p.date !== todayYmd_()) throw new Error('That poll is closed.');
    if (p.choices.indexOf(String(choice)) === -1) throw new Error('Pick one of the choices.');
    if (p.votes.some(function (v) { return id800_(v['800#']) === sid; })) throw new Error('You already voted on this one!');
    const me = sheetRows_(getSetting_('STUDENT_BOARD_TAB') || 'Student Board').filter(function (r) { return id800_(r['800#']) === sid; })[0];
    if (!me) throw new Error('That 800 number isn\'t on the list. Check it, or ask your homeroom teacher.');
    getSheet_('Poll Votes').appendRow([p.id, p.date, sid, squadOf_(me['Squad']), String(choice),
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm')]);
  } finally {
    lock.releaseLock();
  }
  return clashDailyPolls(sid);
}

/** After a poll's day: one daily_poll event with each squad's points. Returns what was sent. */
function awardDailyPolls() {
  if (String(getSetting_('DAILY_POLL_POINTS')).toUpperCase() !== 'ON') return ['DAILY_POLL_POINTS is OFF.'];
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const lines = [];
    const sheet = getSheet_('Daily Polls');
    const statusCol = POLL_HEADERS_.indexOf('Status') + 1;
    pollsWithVotes_().forEach(function (p) {
      if (!p.show || /^awarded/i.test(p.status) || p.date >= todayYmd_() || !p.total) return;
      const winners = pollWinners_(p);
      const pts = { BLACK: 0, GOLD: 0, GREY: 0, PURPLE: 0 };
      p.votes.forEach(function (v) {
        const s = squadOf_(v['Squad']);
        if (pts[s] === undefined) return;
        pts[s] += p.partPts + (winners.indexOf(String(v['Choice'])) > -1 ? p.winPts : 0);
      });
      getSheet_(TABS.EVENTS).appendRow(['', 'Daily Poll: ' + p.question, 'daily_poll', p.date,
        p.total + ' votes; winner: ' + (winners.join(' / ') || 'none'), '', pts.BLACK || '', pts.GOLD || '', pts.GREY || '', pts.PURPLE || '']);
      const results = p.choices.map(function (c) { return c + ' ' + (p.counts[c] || 0); }).join(', ');
      sheet.getRange(p.row, statusCol, 1, 2).setValues([['Awarded ' + todayYmd_(), results + ' · ' + SQUAD_KEYS_.map(function (s) { return s + ' +' + pts[s]; }).join(' ')]]);
      lines.push(p.question + ': ' + SQUAD_KEYS_.map(function (s) { return s + ' +' + pts[s]; }).join(', '));
    });
    if (lines.length) try { writePollInsights_(); } catch (e) { console.warn('Poll Insights: ' + e.message); }
    return lines;
  } finally {
    lock.releaseLock();
  }
}

/** Read-aloud text: today's poll(s) and yesterday's result. Used by the menu and Generals HQ. */
function pollAnnouncement_() {
  const d = JSON.parse(clashDailyPolls(''));
  let text = '';
  if (d.last) {
    const w = d.last.winners;
    text += d.last.kind === 'Trivia'
      ? 'Yesterday\'s trivia: ' + d.last.question + ' The answer: ' + d.last.answer + '! '
      : 'Yesterday\'s This or That: ' + d.last.question + ' ' + (w.length ? 'The winner: ' + w.join(' and ') + '! ' : '');
    text += 'Every squad earned points for each student who voted. ';
  }
  d.today.forEach(function (p) {
    if (p.kind === 'Check-in') {
      text += p.question + ' Check in on clashofclasses.org with your 800 number. Every check-in earns your squad a point, and it counts toward the map! ';
      return;
    }
    text += (p.kind === 'Trivia' ? 'Today\'s trivia: ' + p.question + ' Is it ' + p.choices.join(', ') + '?'
      : p.kind === 'School Pulse' ? 'Today\'s School Pulse question: ' + p.question + ' Help us make Ephrata Middle better.'
      : 'Today\'s This or That: ' + p.question) +
      ' Vote on clashofclasses.org with your 800 number. Every vote earns your squad a point' +
      (p.kind === 'Trivia' ? ', and every right answer earns one more!' : p.kind === 'School Pulse' ? '!' : ', and voting with the winning side earns one more!') + ' ';
  });
  return text.trim() || 'No daily poll today.';
}

function menuPollAnnouncement_() { SpreadsheetApp.getUi().alert('Read this in the morning announcements', pollAnnouncement_(), SpreadsheetApp.getUi().ButtonSet.OK); }

function menuTurnOnPolls_() {
  setSettingValue_('DAILY_POLL_POINTS', 'ON');
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'awardDailyPolls') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('awardDailyPolls').timeBased().everyHours(1).create();
  const lines = awardDailyPolls();
  SpreadsheetApp.getUi().alert('Daily poll points are ON (checked every hour; each poll is scored the day after it runs).\n\n' + (lines.length ? lines.join('\n') : 'Nothing to score yet.'));
}

function menuTurnOffPolls_() {
  setSettingValue_('DAILY_POLL_POINTS', 'OFF');
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'awardDailyPolls') ScriptApp.deleteTrigger(t); });
  SpreadsheetApp.getUi().alert('Daily poll points are OFF. Students can still vote; no points are added.');
}
