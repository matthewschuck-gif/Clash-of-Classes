/**
 * Clash of Classes -- Apps Script backend
 * File 6: Daily Polls for the morning announcements (This or That, Trivia, Pulse, Check-in).
 *
 * Daily Polls tab (one row per poll, filled in by staff):
 *   Show | Date | Kind | Question | Choices | Answer | Participation Points | Win Points | Status | Results | Through
 *   Kind: This or That (the choice with the most votes wins; a tie means every tied choice
 *         wins), Trivia (Answer is the right choice), or Pulse (a question the school wants to
 *         learn from: no right answer, no winner, every voter earns Participation Points, and
 *         students see the live results the moment they vote). Choices separated by | (Coke | Pepsi).
 *   Date: the school day it opens. Through: the last day it is open (blank = just that day; give a
 *         Pulse 2-3 days so more students answer). Status and Results are filled in.
 *         Check-in: one button (Choices may be blank). Pride Day Check-in, Wheel Check-in ... A
 *         Kind with "pride" or "wheel" in it counts for that map zone (18_Conquest.gs check-ins).
 *         A Kind of "School Pulse" is the same as Pulse.
 * Poll Votes tab: one row per vote (Poll | Date | 800# | Squad | Choice | At | Grade). One vote per
 * student per poll; squad and grade come from the Student Board.
 * Poll Ideas tab: a bank of Pulse questions the school would benefit from asking (26_PollIdeas.gs).
 *   Menu › Add the next Pulse question puts the next unused one on Daily Polls.
 *
 * Results students and the TV see are counts only, never an 800# or a name. Pulse questions are
 * answered with an 800#, so staff can see who said what: ask only what a student would be fine
 * saying out loud. Anything sensitive belongs in the counselors' anonymous surveys, not here.
 *
 * Points (when DAILY_POLL_POINTS is ON, checked hourly once the menu turns it on): after a
 * poll closes, each squad gets Participation Points for every student who voted plus (This or
 * That and Trivia) Win Points for every student who picked the winner, as one "Daily Poll" event
 * (type daily_poll). Never twice: Status becomes "Awarded". Pulse results add a by-grade breakdown.
 */

const POLL_HEADERS_ = ['Show', 'Date', 'Kind', 'Question', 'Choices', 'Answer', 'Participation Points', 'Win Points', 'Status', 'Results', 'Through'];
const VOTE_HEADERS_ = ['Poll', 'Date', '800#', 'Squad', 'Choice', 'At', 'Grade'];

// The Pulse bank: short, low-stakes questions whose answers help the school plan. Nothing about
// home life, health, safety incidents or other students: those go to counselor-run anonymous surveys.
// [_, Category, Question, Choices, Why the school wants to know, Who uses it]. 26_PollIdeas.gs adds
// these to the Poll Ideas tab (its columns) with the rest of the ideas.
const PULSE_BANK_ = [
  ['', 'Belonging', 'How many adults at EMS would you go to if you needed help?', '0 | 1 | 2 | 3+', 'Every student should have at least one trusted adult. Shows how many feel they have none.', 'Counselors, admin'],
  ['', 'Belonging', 'Where at school do you feel most like you belong?', 'Classroom | Lunch | Squad events | Clubs / sports | Hallways | Not sure yet', 'Shows which spaces are working and which need attention.', 'Admin, Clash team'],
  ['', 'Belonging', 'How many people in your squad do you know by name?', '0-2 | 3-5 | 6-10 | 11+', 'Measures whether squads are building connection across grades.', 'Clash team, Generals'],
  ['', 'Belonging', 'Best way for a teacher to recognize you?', 'Say it to me | Note home | Shout-out | Trait Dough | Keep it private', 'Lets staff recognize students the way they actually like.', 'All staff'],
  ['', 'Learning', 'When you get stuck on homework, what do you do first?', 'Ask a friend | Look it up | Ask a teacher next day | Skip it | Ask family', 'Shows where help-seeking breaks down: plan tutoring and office hours.', 'Teachers, MTSS team'],
  ['', 'Learning', 'Best time for extra help?', 'Before school | Homeroom / WIN | Lunch | After school | Online', 'Schedules tutoring when students will come.', 'Teachers, admin'],
  ['', 'Learning', 'How do you study for a test?', 'Reread notes | Quiz myself | Study with friends | Practice problems | I don\'t study', 'Shows which study skills to teach in Mindset lessons.', 'Mindset team, teachers'],
  ['', 'Learning', 'How do you learn best?', 'Hands-on | Watching it done | Talking it out | Reading | Practice on my own', 'Helps teachers balance how they teach.', 'Teachers'],
  ['', 'Learning', 'Do you know your current grade in every class?', 'Yes, all | Most | A few | No', 'Shows who checks PowerSchool: target grade check-ins before failing grades pile up.', 'Teachers, counselors'],
  ['', 'Learning', 'How often do you check PowerSchool or Schoology?', 'Daily | Weekly | Only when told | Never', 'Same: a lead indicator for failing grades (the 70s Hall).', 'Teachers, counselors'],
  ['', 'Learning', 'How much homework time is right for one night?', '0-15 min | 15-30 min | 30-60 min | 60+ min', 'Gives teams a student view when planning homework loads.', 'Grade-level teams'],
  ['', 'Mindset', 'How do you feel when you get a hard problem?', 'Excited | Fine | Nervous | Want to quit', 'Tracks growth mindset across the year.', 'Mindset team'],
  ['', 'Mindset', 'Which Mountaineer trait are you best at right now?', 'Growth Mindset | Problem Solving | Communication | Collaboration | Responsibility', 'Shows which LRG traits stick and which to revisit.', 'Mindset team'],
  ['', 'Mindset', 'Which trait do you most want to grow this month?', 'Growth Mindset | Problem Solving | Communication | Collaboration | Responsibility', 'Students pick their own goal: plan the next Mindset lesson around it.', 'Mindset team'],
  ['', 'Attendance', 'What makes you most excited to come to school?', 'Friends | A class I like | Clash / squad events | Sports / clubs | Lunch', 'Builds on what already brings students in (the 50s Hall).', 'Admin, attendance team'],
  ['', 'Attendance', 'What makes mornings hardest?', 'Waking up | Getting a ride | Not ready for a test | Nothing, mornings are fine | Rather not say', 'Points the attendance team at fixable barriers.', 'Attendance team, counselors'],
  ['', 'Attendance', 'What time do you usually go to sleep on a school night?', 'Before 9 | 9-10 | 10-11 | After 11', 'Sleep drives attendance and focus: shapes family tips in the newsletter.', 'Counselors, nurse'],
  ['', 'Lunch & Building', 'What would make lunch better?', 'More time | Seating choice | Lunch activities | Quieter area | Different food', 'Direct student input for the Cafeteria (fewer LRRs and LDs).', 'Admin, cafeteria staff'],
  ['', 'Lunch & Building', 'Where do hallway problems happen most?', 'By lockers | Stairways | Near bathrooms | Cafeteria line | Not a problem', 'Places adults where they matter (areas only, no names).', 'Admin'],
  ['', 'Lunch & Building', 'How long does it take you to get to class?', 'Plenty of time | Just enough | I\'m often late | Depends on the class', 'Checks passing time against tardies.', 'Admin'],
  ['', 'Clash', 'Which squad event should come back?', 'Mayhem Week | Spin Wheel | Kick Off | Squad Duels | Trivia', 'Plans next semester\'s events around what students want.', 'Clash team, Generals'],
  ['', 'Clash', 'Best reward for a winning squad?', 'Extra recess | Dress-down day | Food | Music at lunch | A field trip', 'Picks prizes that actually motivate.', 'Clash team'],
  ['', 'Clash', 'How do you hear about Clash news first?', 'Announcements | Homeroom teacher | Friends | Website / TV | Generals', 'Shows which channel to put news in.', 'Clash team, Generals'],
  ['', 'Clash', 'How often do you check clashofclasses.org?', 'Daily | Weekly | Only for polls | Rarely', 'Measures site reach.', 'Clash team'],
  ['', 'Clash', 'What would get more of your squad to vote and pitch in?', 'Bigger rewards | Reminders | Friendly competition | Class time to do it | Fun questions', 'Ideas for turnout (Auditorium zone).', 'Generals, Clash team'],
  ['', 'Clash', 'Which squad meeting activity do you like best?', 'Games | Hearing scores | Planning events | Mindset talk | Free time with squad', 'Shapes January squad meetings.', 'Generals, Clash team'],
  ['', 'Activities', 'Which new club would you join?', 'Gaming / esports | Art | Coding / robotics | Cooking | Outdoors | Music', 'Gauges interest before starting a club.', 'Activities, admin'],
  ['', 'Activities', 'What stops you from joining clubs or sports?', 'Ride home | Time | Didn\'t know about it | Friends aren\'t in it | Not interested', 'Shows barriers to involvement (late bus, publicity).', 'Activities, admin'],
  ['', 'Activities', 'Would you come to a Squad Night?', 'Yes | Maybe | Only with a friend | No', 'Plans attendance and staffing for evening events.', 'Clash team'],
  ['', 'Future Ready', 'What job sounds most interesting right now?', 'Healthcare | Trades | Tech | Business | Arts | Education | Not sure', 'Plans career speakers and LRG lessons.', 'Counselors'],
  ['', 'Future Ready', 'Have you talked with an adult about high school course choices?', 'Yes | Not yet | Didn\'t know I should', '8th-grade scheduling support.', 'Counselors (8th grade)'],
  ['', 'Future Ready', 'Which skill do you most want to learn outside class?', 'Cooking | Money | Fixing things | Public speaking | First aid | Coding', 'Ideas for advisory and WIN time.', 'Admin, Mindset team'],
  ['', 'Tech', 'What device do you do homework on most?', 'School Chromebook | Phone | Home computer | Tablet | No device', 'Device and internet access at home.', 'Tech team, admin'],
  ['', 'Tech', 'How is your internet at home for schoolwork?', 'Great | Okay | Spotty | None', 'Flags who needs hotspots before online assignments.', 'Tech team, counselors'],
  ['', 'Communication', 'Best way to remind you about a deadline?', 'Schoology | Email | Said in class | Posted on the board | Family text', 'Helps teachers reach students where they look.', 'Teachers'],
  ['', 'Communication', 'Do you read the morning announcements slides?', 'Always | Sometimes | Never | Didn\'t know there were slides', 'Tests whether announcements land.', 'Admin, Clash team'],
  ['', 'Wellness', 'How much water do you drink at school?', 'None | 1 bottle | 2+ bottles | Not sure', 'Light check-in: refill stations and passes.', 'Nurse'],
  ['', 'Wellness', 'What helps you reset when you\'re stressed at school?', 'Music | Walk / move | Talk to someone | Quiet space | Drawing / writing', 'Stocks calm-down spaces with what students use.', 'Counselors'],
  ['', 'Kindness', 'How often do you see someone do something kind at school?', 'Every day | A few times a week | Rarely | Never', 'A Freudenfreude check: pairs with shout-out counts (60s Hall).', 'Clash team, counselors'],
  ['', 'Kindness', 'Best way to welcome a new student?', 'Show them around | Sit with them at lunch | Squad buddy | Add them to a group chat | Introduce them to a teacher', 'Builds the new-student buddy plan.', 'Counselors, Generals'],
];

function ensurePollTabs_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  [['Daily Polls', POLL_HEADERS_, [[false, today, 'This or That', 'Coke or Pepsi?', 'Coke | Pepsi', '', 1, 1, '', '', ''],
    [false, today, 'Trivia', 'How many squads are in Clash of Classes?', '3 | 4 | 5', '4', 1, 1, '', '', ''],
    [false, today, 'Pulse', 'Best time for extra help?', 'Before school | Homeroom / WIN | Lunch | After school | Online', '', 1, '', '', '', '']]],
   ['Poll Votes', VOTE_HEADERS_, null]].forEach(function (t) {
    const sh = ss.getSheetByName(t[0]);
    if (sh) {
      // An older tab: add the newer columns at the end (Through, Grade).
      const have = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0].map(function (h) { return String(h).trim(); });
      const missing = t[1].filter(function (h) { return have.indexOf(h) === -1; });
      if (missing.length && have.filter(String).length) sh.getRange(1, have.filter(String).length + 1, 1, missing.length).setValues([missing]).setFontWeight('bold').setBackground('#490e6f').setFontColor('#ffffff');
      return;
    }
    const n = ss.insertSheet(t[0]);
    n.getRange(1, 1, 1, t[1].length).setValues([t[1]]).setFontWeight('bold').setBackground('#490e6f').setFontColor('#ffffff');
    n.setFrozenRows(1);
    if (t[2]) n.getRange(2, 1, t[2].length, t[1].length).setValues(t[2]);
  });
  ensurePollIdeaTabs_(); // Poll Ideas + Poll Insights (26_PollIdeas.gs)
}

function pollChoices_(v) { return String(v || '').split('|').map(function (s) { return s.trim(); }).filter(String); }
const CHECK_IN_CHOICE_ = 'I\'m in! ✋';
// The four kinds the site knows. The Kind cell can say more ("Pride Day Check-in", "School Pulse").
function pollKind_(v) {
  const k = String(v || '');
  if (/trivia/i.test(k)) return 'Trivia';
  if (/check|pride|wheel/i.test(k)) return 'Check-in';
  if (/pulse|survey|school/i.test(k)) return 'Pulse';
  return 'This or That';
}
function todayYmd_() { return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'); }

// Every poll row, with its votes tallied.
function pollsWithVotes_() {
  const votes = sheetRows_('Poll Votes');
  return sheetRows_('Daily Polls').map(function (r) {
    const id = String(r._row);
    const mine = votes.filter(function (v) { return String(v['Poll']) === id; });
    const counts = {}, bySquad = {}, byGrade = {};
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
      const g = String(v['Grade'] || '').replace(/\D/g, '') || '?';
      byGrade[g] = byGrade[g] || {};
      byGrade[g][c] = (byGrade[g][c] || 0) + 1;
    });
    const date = String(r['Date'] || '').slice(0, 10), through = String(r['Through'] || '').slice(0, 10);
    return {
      id: id, row: r._row, show: ticked_(r['Show']), date: date, through: through > date ? through : date, kind: kind, kindLabel: String(r['Kind'] || '').trim() || kind,
      question: String(r['Question'] || ''), choices: choices, answer: String(r['Answer'] || '').trim(),
      partPts: Number(r['Participation Points']) || 0, winPts: Number(r['Win Points']) || 0,
      status: String(r['Status'] || ''), counts: counts, bySquad: bySquad, byGrade: byGrade, total: mine.length, votes: mine,
    };
  });
}

function pollOpenOn_(p, ymd) { return p.date <= ymd && ymd <= p.through; }

function pollWinners_(p) {
  if (p.kind === 'Pulse') return [];
  if (p.kind === 'Trivia') return p.answer ? [p.answer] : [];
  if (p.kind === 'Check-in') return []; // nobody wins: every check-in counts the same
  const max = Math.max.apply(null, p.choices.map(function (c) { return p.counts[c] || 0; }).concat([0]));
  return max ? p.choices.filter(function (c) { return (p.counts[c] || 0) === max; }) : [];
}

// What a student sees (no 800#s, no other students): today's polls, results once they voted.
function publicPoll_(p, votedChoice) {
  const closed = p.through < todayYmd_();
  return {
    id: p.id, kind: p.kind, kindLabel: p.kindLabel, question: p.question, choices: p.choices, date: p.date, through: p.through,
    voted: votedChoice || '', closed: closed,
    counts: votedChoice || closed ? p.counts : null, total: p.total,
    answer: closed && p.kind === 'Trivia' ? p.answer : '', // trivia answer only after the day
    winners: closed ? pollWinners_(p) : [],
  };
}

/** Today's polls (and the last closed one's results) for the site. 800# optional: marks what they voted. */
function clashDailyPolls(id) {
  const sid = id800_(id);
  const today = todayYmd_();
  const polls = pollsWithVotes_().filter(function (p) { return p.show && p.question && p.choices.length >= (p.kind === 'Check-in' ? 1 : 2); });
  const voteOf = function (p) {
    const v = sid ? p.votes.filter(function (x) { return id800_(x['800#']) === sid; })[0] : null;
    return v ? String(v['Choice']) : '';
  };
  const past = polls.filter(function (p) { return p.through < today; }).sort(function (a, b) { return b.through.localeCompare(a.through); });
  return JSON.stringify({
    today: polls.filter(function (p) { return pollOpenOn_(p, today); }).map(function (p) { return publicPoll_(p, voteOf(p)); }),
    last: past.length ? publicPoll_(past[0], voteOf(past[0])) : null,
  });
}

/**
 * The live Pulse screen (Menti-style), for the Pulse page and the TV: open Pulse questions with
 * their counts and how much of each squad has answered, plus the latest few closed ones.
 * Counts only: no 800#s, names or grades. Live results are the point, so nothing waits for a vote.
 */
function clashPulseLive() {
  const today = todayYmd_();
  const sizes = { BLACK: 0, GOLD: 0, GREY: 0, PURPLE: 0 };
  sheetRows_(getSetting_('STUDENT_BOARD_TAB') || 'Student Board').forEach(function (r) { const s = squadOf_(r['Squad']); if (sizes[s] !== undefined) sizes[s]++; });
  const shape = function (p) {
    const turnout = {};
    SQUAD_KEYS_.forEach(function (s) {
      const n = Object.keys(p.bySquad[s] || {}).reduce(function (t, c) { return t + p.bySquad[s][c]; }, 0);
      turnout[s] = { n: n, pct: sizes[s] ? Math.round(100 * n / sizes[s]) : 0 };
    });
    return { id: p.id, question: p.question, choices: p.choices, counts: p.counts, total: p.total, through: p.through, closed: p.through < today, turnout: turnout };
  };
  const pulses = pollsWithVotes_().filter(function (p) { return p.show && p.kind === 'Pulse' && p.question && p.choices.length >= 2; });
  return JSON.stringify({
    live: pulses.filter(function (p) { return pollOpenOn_(p, today); }).map(shape),
    recent: pulses.filter(function (p) { return p.through < today; }).sort(function (a, b) { return b.through.localeCompare(a.through); }).slice(0, 3).map(shape),
    youSaid: youSaidPublic_(4), // 30_YouSaid.gs: what changed because of a Pulse
  });
}

/** One vote: 800# on the Student Board, a poll open today, one vote per poll. */
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
    if (!pollOpenOn_(p, todayYmd_())) throw new Error('That poll is closed.');
    if (p.choices.indexOf(String(choice)) === -1) throw new Error('Pick one of the choices.');
    if (p.votes.some(function (v) { return id800_(v['800#']) === sid; })) throw new Error('You already voted on this one!');
    const me = sheetRows_(getSetting_('STUDENT_BOARD_TAB') || 'Student Board').filter(function (r) { return id800_(r['800#']) === sid; })[0];
    if (!me) throw new Error('That 800 number isn\'t on the list. Check it, or ask your homeroom teacher.');
    // Written by header so an older Poll Votes tab (no Grade column yet) still lines up.
    const sh = getSheet_('Poll Votes');
    const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(function (h) { return String(h).trim(); });
    const vals = { 'Poll': p.id, 'Date': todayYmd_(), '800#': sid, 'Squad': squadOf_(me['Squad']), 'Choice': String(choice),
      'At': Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'), 'Grade': String(me['Grade'] || '') };
    sh.appendRow(head.map(function (h) { return vals[h] === undefined ? '' : vals[h]; }));
  } finally {
    lock.releaseLock();
  }
  return clashDailyPolls(sid);
}

/** After a poll closes: one daily_poll event with each squad's points. Returns what was sent. */
function awardDailyPolls() {
  if (String(getSetting_('DAILY_POLL_POINTS')).toUpperCase() !== 'ON') return ['DAILY_POLL_POINTS is OFF.'];
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const lines = [];
    const sheet = getSheet_('Daily Polls');
    const statusCol = POLL_HEADERS_.indexOf('Status') + 1;
    pollsWithVotes_().forEach(function (p) {
      if (!p.show || /^awarded/i.test(p.status) || p.through >= todayYmd_() || !p.total) return;
      const winners = pollWinners_(p);
      const pts = { BLACK: 0, GOLD: 0, GREY: 0, PURPLE: 0 };
      p.votes.forEach(function (v) {
        const s = squadOf_(v['Squad']);
        if (pts[s] === undefined) return;
        pts[s] += p.partPts + (winners.indexOf(String(v['Choice'])) > -1 ? p.winPts : 0);
      });
      getSheet_(TABS.EVENTS).appendRow(['', (p.kind === 'Pulse' ? 'Pulse: ' : 'Daily Poll: ') + p.question, 'daily_poll', p.through,
        p.total + ' votes' + (p.kind === 'Pulse' ? '' : '; winner: ' + (winners.join(' / ') || 'none')), '', pts.BLACK || '', pts.GOLD || '', pts.GREY || '', pts.PURPLE || '']);
      let results = p.choices.map(function (c) { return c + ' ' + (p.counts[c] || 0); }).join(', ');
      if (p.kind === 'Pulse') {
        results += ' · by grade: ' + Object.keys(p.byGrade).sort().map(function (g) {
          return (g === '?' ? 'no grade' : g + 'th') + ' (' + p.choices.map(function (c) { return c + ' ' + (p.byGrade[g][c] || 0); }).join(', ') + ')';
        }).join('; ');
      }
      sheet.getRange(p.row, statusCol, 1, 2).setValues([['Awarded ' + todayYmd_(), results + ' · ' + SQUAD_KEYS_.map(function (s) { return s + ' +' + pts[s]; }).join(' ')]]);
      lines.push(p.question + ': ' + SQUAD_KEYS_.map(function (s) { return s + ' +' + pts[s]; }).join(', '));
    });
    if (lines.length) try { writePollInsights_(); } catch (e) { console.warn('Poll Insights: ' + e.message); }
    return lines;
  } finally {
    lock.releaseLock();
  }
}

/** Read-aloud text: today's poll(s) and the last result. Used by the menu and Generals HQ. */
function pollAnnouncement_() {
  const d = JSON.parse(clashDailyPolls(''));
  let text = '';
  if (d.last) {
    const w = d.last.winners;
    if (d.last.kind === 'Trivia') text += 'Yesterday\'s trivia: ' + d.last.question + ' The answer: ' + d.last.answer + '! ';
    else if (d.last.kind === 'Pulse') {
      const top = d.last.choices.slice().sort(function (a, b) { return (d.last.counts[b] || 0) - (d.last.counts[a] || 0); })[0];
      text += 'Our last Pulse question: ' + d.last.question + ' ' + d.last.total + ' of you answered, and the top answer was ' + top + '. Thank you: staff use these answers to plan. ';
    } else text += 'Yesterday\'s This or That: ' + d.last.question + ' ' + (w.length ? 'The winner: ' + w.join(' and ') + '! ' : '');
    text += 'Every squad earned points for each student who voted. ';
  }
  d.today.forEach(function (p) {
    if (p.kind === 'Check-in') {
      text += p.question + ' Check in on clashofclasses.org with your 800 number. Every check-in earns your squad a point, and it counts toward the map! ';
      return;
    }
    if (p.kind === 'Pulse') {
      text += 'This week\'s Pulse question: ' + p.question + ' There\'s no right answer: the school wants to know what you think. Answer on clashofclasses.org with your 800 number ' +
        (p.through > p.date ? 'any day through ' + p.through.slice(5).replace('-', '/') : 'today') + ', and every answer earns your squad a point. ';
      return;
    }
    text += (p.kind === 'Trivia' ? 'Today\'s trivia: ' + p.question + ' Is it ' + p.choices.join(', ') + '?' : 'Today\'s This or That: ' + p.question) +
      ' Vote on clashofclasses.org with your 800 number. Every vote earns your squad a point' +
      (p.kind === 'Trivia' ? ', and every right answer earns one more!' : p.kind === 'Pulse' ? '!' : ', and voting with the winning side earns one more!') + ' ';
  });
  return text.trim() || 'No daily poll today.';
}

// ── POLL IDEAS → DAILY POLLS ────────────────────────────────────────
// The school day n school days after ymd (weekends and Days Off skipped).
function schoolDayAfter_(ymd, n) {
  let off = {};
  try { schoolCalendar_().noSchool.forEach(function (x) { off[x] = true; }); } catch (e) { off = {}; }
  let d = ymd, added = 0;
  while (added < n) {
    d = ymdShift_(d, 1);
    const dow = new Date(d + 'T12:00:00').getDay();
    if (dow !== 0 && dow !== 6 && !off[d]) added++;
  }
  return d;
}

/** Puts the next ready Pulse idea (Use ticked, not used yet or a monthly one that is due) on Daily Polls, open PULSE_DAYS school days. */
function addNextPulse_() {
  ensurePollTabs_(); ensurePollIdeaTabs_();
  const busy = pollsWithVotes_().filter(function (p) { return p.kind === 'Pulse' && p.show; }).map(function (p) { return p.through; }).sort().pop() || '';
  const start = schoolDayAfter_(busy > todayYmd_() ? busy : todayYmd_(), 1);
  const idea = readyIdeas_(start).filter(function (r) { return pollKind_(r['Kind']) === 'Pulse' && pollChoices_(r['Choices']).length >= 2; })[0];
  if (!idea) return null;
  const days = Math.max(1, Math.min(10, Number(getSetting_('PULSE_DAYS')) || 3));
  const through = days > 1 ? schoolDayAfter_(start, days - 1) : start;
  const sh = getSheet_('Daily Polls');
  const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(function (h) { return String(h).trim(); });
  const vals = { 'Show': true, 'Date': start, 'Kind': 'Pulse', 'Question': String(idea['Question']).trim(), 'Choices': String(idea['Choices']).trim(),
    'Participation Points': 1, 'Through': through };
  sh.appendRow(head.map(function (h) { return vals[h] === undefined ? '' : vals[h]; }));
  const usedCol = headerCol_(getSheet_('Poll Ideas'), 'Used On');
  if (usedCol) getSheet_('Poll Ideas').getRange(idea._row, usedCol).setValue(start);
  return { question: vals['Question'], start: start, through: through };
}

function menuNextPulse_() {
  const r = addNextPulse_();
  SpreadsheetApp.getUi().alert(r ? 'Added to Daily Polls: "' + r.question + '"\nOpen ' + r.start + ' through ' + r.through + '. Edit the row to change anything.'
    : 'No Pulse idea is ready on the Poll Ideas tab (Use ticked, Used On blank). Add rows there or clear a Used On date.');
}

function menuPollAnnouncement_() { SpreadsheetApp.getUi().alert('Read this in the morning announcements', pollAnnouncement_(), SpreadsheetApp.getUi().ButtonSet.OK); }

function menuTurnOnPolls_() {
  setSettingValue_('DAILY_POLL_POINTS', 'ON');
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'awardDailyPolls') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('awardDailyPolls').timeBased().everyHours(1).create();
  const lines = awardDailyPolls();
  SpreadsheetApp.getUi().alert('Daily poll points are ON (checked every hour; each poll is scored the day after it closes).\n\n' + (lines.length ? lines.join('\n') : 'Nothing to score yet.'));
}

function menuTurnOffPolls_() {
  setSettingValue_('DAILY_POLL_POINTS', 'OFF');
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'awardDailyPolls') ScriptApp.deleteTrigger(t); });
  SpreadsheetApp.getUi().alert('Daily poll points are OFF. Students can still vote; no points are added.');
}
