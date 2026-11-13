/**
 * Clash of Classes -- Apps Script backend
 * File 26: Poll Ideas, Poll Insights and the live Poll Board.
 *
 * POLL IDEAS tab: a bank of questions, ready to run. Most are School Pulse questions, things the
 * school learns from (how students get to school, sleep, homework help, what they would join).
 * Every answer comes with an 800#, so staff can see answers by squad and grade.
 *   Use | Kind | Question | Choices | Why We Ask | Used On
 *   Menu "Schedule poll ideas" puts the next ticked, unused ideas on Daily Polls, one per school
 *   day with no poll yet, and writes the date in Used On.
 * POLL INSIGHTS tab (staff only, menu or after each poll is scored): every poll's answers counted
 *   by squad and by grade. Counts only, never names; the 800#s stay on Poll Votes.
 * THE POLL BOARD (?page=pollboard): the projector view. Big question, live bars, the vote count
 *   and squad turnout per 100 students, refreshed every few seconds, like a Menti screen.
 *   Trivia hides the counts until the day is over so the room can't copy the answer.
 */

const IDEA_HEADERS_ = ['Use', 'Kind', 'Question', 'Choices', 'Why We Ask', 'Used On'];
const INSIGHT_HEADERS_ = ['Date', 'Kind', 'Question', 'Answer', 'Total', 'BLACK', 'GOLD', 'GREY', 'PURPLE', 'Grade 7', 'Grade 8', 'Other Grade'];

// [Kind, Question, Choices, Why We Ask]
const POLL_IDEAS_ = [
  ['Pride Day Check-in', 'It\'s Mountaineer Pride Day! Are you wearing purple, gold or Mounts gear?', 'I\'m repping! 💜💛', 'Pride Day turnout by squad (Gym 2 on the map). Check-ins count double on the map.'],
  ['Wheel Check-in', 'Spin Wheel today! Check in to put your squad in the spin.', 'Put us in! 🎡', 'Who joins the Wheel (the Auditorium on the map).'],
  ['School Pulse', 'How do you usually get to school?', 'Bus | Car | Walk or bike | Other', 'Arrival and dismissal planning; who is hurt by late buses.'],
  ['School Pulse', 'How much sleep did you get last night?', 'Under 6 hours | 6 to 7 | 8 to 9 | 10 or more', 'Sleep and attendance, tardies and focus.'],
  ['School Pulse', 'Did you eat breakfast this morning?', 'Yes | No', 'Whether to push breakfast in the classroom or grab-and-go.'],
  ['School Pulse', 'If we started a new club, which would you join?', 'Esports | Robotics | Art | Cooking | Something else', 'Which clubs to start.'],
  ['School Pulse', 'When is the best time for a squad event?', 'Before school | Lunch | PASS | After school', 'Scheduling squad events where students can come.'],
  ['School Pulse', 'Where do you go first when you need help with schoolwork?', 'A teacher | A friend | Family | Online | Nobody yet', 'Who needs a homework help connection.'],
  ['School Pulse', 'How do you hear about school events?', 'Announcements | Teachers | Friends | Posters | The website', 'Where to put announcements so students see them.'],
  ['School Pulse', 'Which reward would you work hardest for?', 'Extra recess | Food | Free time | A prize | A shout-out', 'Planning PBIS and squad rewards.'],
  ['School Pulse', 'How long did homework take you last night?', 'None | Under 1 hour | 1 to 2 hours | Over 2 hours', 'Homework load by grade.'],
  ['School Pulse', 'What is the hardest part of your school day?', 'Mornings | Class work | Lunch | Hallways | After school', 'Where to put adults and supports.'],
  ['School Pulse', 'Which subject do you feel most confident in?', 'Math | ELA | Science | Social Studies | Electives', 'Confidence next to grades and STAR.'],
  ['School Pulse', 'How do you study for a test?', 'Reread notes | Flashcards | Practice problems | Ask someone | I don\'t', 'Which study skills to teach in Core Extension.'],
  ['School Pulse', 'Do you read for fun outside school?', 'Most days | Sometimes | Rarely | Never', 'Reading habits next to STAR reading.'],
  ['School Pulse', 'What do you use for homework at home?', 'Chromebook | Phone | Home computer | Nothing', 'Device needs at home.'],
  ['School Pulse', 'Is your internet at home reliable?', 'Yes | Sometimes | No', 'Who may need a hotspot or time at school.'],
  ['School Pulse', 'What do you do after school most days?', 'Sports | Music or arts | Job or family care | Clubs | Nothing set', 'When students are free; who has family duties.'],
  ['School Pulse', 'What would make lunch better?', 'More time | Pick my seat | Activities | Quieter', 'Lunch changes students want.'],
  ['School Pulse', 'How do you like to get feedback on your work?', 'Written notes | Talking with my teacher | A rubric | Just the grade', 'Feedback practice for staff PD.'],
  ['School Pulse', 'How many days this week did you read for 20 minutes?', 'None | 1 or 2 | 3 or 4 | 5 or more', 'Reading goal tracking.'],
  ['School Pulse', 'Do you know your STAR goal?', 'Yes | Kind of | No', 'Whether goal setting is reaching students.'],
  ['School Pulse', 'When do you usually do homework?', 'Right after school | Evening | Late at night | At school', 'Homework timing and sleep.'],
  ['School Pulse', 'Would you come to a homework help session?', 'Before school | Lunch | After school | No thanks', 'Whether to run homework help, and when.'],
  ['School Pulse', 'What helps you calm down when you are stressed?', 'Music | Moving | Talking to someone | Gaming | Something else', 'Calm-down tools for classrooms and the Belay Line.'],
  ['School Pulse', 'How full is your battery today?', 'Full | Half | Low | Empty', 'A quick mood check by grade and squad.'],
  ['School Pulse', 'How do the hallways feel between classes?', 'Calm | Mostly fine | Too crowded | Not safe', 'Where hallway supervision is needed. Counts only on Poll Insights.'],
  ['School Pulse', 'Is there an adult in this building you could go to with a problem?', 'Yes | Not sure | No', 'Connection to an adult. Answers carry an 800#: decide who follows up before running it.'],
  ['School Pulse', 'How did you feel about this week at school?', 'Great | Good | OK | Rough', 'A weekly climate check.'],
  ['School Pulse', 'Which Clash event should come back?', 'Mayhem events | Trivia | Spirit days | Lunch battles | Squad Duels', 'Planning next year\'s Clash.'],
  ['This or That', 'Pizza or tacos?', 'Pizza | Tacos', 'Just for fun.'],
  ['This or That', 'Early bird or night owl?', 'Early bird | Night owl', 'Just for fun (and a little about sleep).'],
  ['Trivia', 'How many squads are in Clash of Classes?', '3 | 4 | 5', 'Just for fun.'],
];

function ensurePollIdeaTabs_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  [['Poll Ideas', IDEA_HEADERS_, POLL_IDEAS_.map(function (r) { return [true, r[0], r[1], r[2], r[3], '']; })],
   ['Poll Insights', INSIGHT_HEADERS_, null]].forEach(function (t) {
    if (ss.getSheetByName(t[0])) return;
    const sh = ss.insertSheet(t[0]);
    sh.getRange(1, 1, 1, t[1].length).setValues([t[1]]).setFontWeight('bold').setBackground('#490e6f').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    if (t[2]) sh.getRange(2, 1, t[2].length, t[1].length).setValues(t[2]);
  });
}

// The next school days (weekdays not on the School Calendar tab), starting tomorrow.
function nextSchoolDays_(n, from) {
  const off = {};
  schoolCalendar_().noSchool.forEach(function (d) { off[d] = true; });
  const out = [];
  let d = from || todayYmd_();
  for (let i = 0; i < 60 && out.length < n; i++) {
    d = ymdShift_(d, 1);
    const dow = new Date(d + 'T12:00:00Z').getUTCDay();
    if (dow > 0 && dow < 6 && !off[d]) out.push(d);
  }
  return out;
}

/** Puts unused ideas on Daily Polls, one per school day that has no poll. Returns what it did. */
function schedulePollIdeas_(days) {
  ensurePollTabs_(); ensurePollIdeaTabs_();
  const taken = {};
  sheetRows_('Daily Polls').forEach(function (p) { taken[String(p['Date'] || '').slice(0, 10)] = true; });
  const open = nextSchoolDays_(days || 5).filter(function (d) { return !taken[d]; });
  const ideas = sheetRows_('Poll Ideas').filter(function (r) { return ticked_(r['Use']) && String(r['Question'] || '').trim() && !String(r['Used On'] || '').trim(); });
  const polls = getSheet_('Daily Polls'), usedCol = IDEA_HEADERS_.indexOf('Used On') + 1, ideaSheet = getSheet_('Poll Ideas');
  const lines = [];
  open.forEach(function (d, i) {
    const idea = ideas[i];
    if (!idea) return;
    polls.appendRow([true, d, idea['Kind'] || 'School Pulse', idea['Question'], idea['Choices'] || '', '', 1, pollKind_(idea['Kind']) === 'This or That' || pollKind_(idea['Kind']) === 'Trivia' ? 1 : 0, '', '']);
    ideaSheet.getRange(idea._row, usedCol).setValue(d);
    lines.push(d + ': ' + idea['Question']);
  });
  return lines;
}

function menuSchedulePollIdeas_() {
  const lines = schedulePollIdeas_(5);
  SpreadsheetApp.getUi().alert(lines.length ? 'Put on Daily Polls (Trivia needs its Answer typed in):\n\n' + lines.join('\n')
    : 'Nothing scheduled: the next 5 school days already have polls, or every ticked idea is used. Add ideas on the Poll Ideas tab.');
}

/** Poll Insights: each poll's answers by squad and grade (counts only). Rewritten each time. */
function writePollInsights_() {
  ensurePollIdeaTabs_();
  const grade = {};
  sheetRows_(getSetting_('STUDENT_BOARD_TAB') || 'Student Board').forEach(function (r) { grade[id800_(r['800#'])] = String(r['Grade'] || '').replace(/\D/g, ''); });
  const rows = [];
  pollsWithVotes_().filter(function (p) { return p.total; }).sort(function (a, b) { return b.date.localeCompare(a.date); }).forEach(function (p) {
    p.choices.forEach(function (c) {
      const o = { Date: p.date, Kind: p.kindLabel, Question: p.question, Answer: c, Total: 0, BLACK: 0, GOLD: 0, GREY: 0, PURPLE: 0, 'Grade 7': 0, 'Grade 8': 0, 'Other Grade': 0 };
      p.votes.forEach(function (v) {
        if (String(v['Choice']) !== c) return;
        o.Total++;
        const s = squadOf_(v['Squad']); if (o[s] !== undefined) o[s]++;
        const g = grade[id800_(v['800#'])];
        o[g === '7' ? 'Grade 7' : g === '8' ? 'Grade 8' : 'Other Grade']++;
      });
      rows.push(o);
    });
  });
  writeRowsByHeader_(getSheet_('Poll Insights'), rows);
  return rows.length;
}

function menuPollInsights_() {
  const n = writePollInsights_();
  SpreadsheetApp.getUi().alert('Poll Insights is up to date: ' + n + ' answer rows, by squad and grade.');
}

// ── THE POLL BOARD (?page=pollboard) ────────────────────────────────
function servePollBoard_() {
  return HtmlService.createHtmlOutputFromFile('PollBoard')
    .setTitle('Live Poll · Clash of Classes')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Today's polls with live counts for the projector (no 800#s, no names). */
function clashPollBoard() {
  const today = todayYmd_();
  const size = squadSizes_();
  const all = pollsWithVotes_().filter(function (p) { return p.show && p.question && p.choices.length >= (p.kind === 'Check-in' ? 1 : 2); });
  let list = all.filter(function (p) { return p.date === today; });
  const live = list.length > 0;
  if (!live) list = all.filter(function (p) { return p.date < today; }).sort(function (a, b) { return b.date.localeCompare(a.date); }).slice(0, 1);
  return JSON.stringify({
    live: live, url: String(getSetting_('SITE_URL') || 'clashofclasses.org').replace(/^https?:\W+/i, ''),
    polls: list.map(function (p) {
      const hide = p.kind === 'Trivia' && p.date === today; // no copying the answer
      const turnout = {};
      SQUAD_KEYS_.forEach(function (s) {
        const n = Object.keys(p.bySquad[s] || {}).reduce(function (t, c) { return t + p.bySquad[s][c]; }, 0);
        turnout[s] = size[s] ? Math.round(1000 * n / size[s]) / 10 : n;
      });
      return { id: p.id, kind: p.kind, kindLabel: p.kindLabel, question: p.question, choices: p.choices, total: p.total, date: p.date,
        counts: hide ? null : p.counts, answer: p.date < today && p.kind === 'Trivia' ? p.answer : '', turnout: turnout, per100: SQUAD_KEYS_.some(function (s) { return size[s]; }) };
    }),
  });
}

function menuOpenPollBoard_() {
  const url = ScriptApp.getService().getUrl() + '?page=pollboard';
  SpreadsheetApp.getUi().alert('The live Poll Board', 'Open this on the projector or a TV (full screen):\n\n' + url +
    '\n\nIt shows today\'s polls with live bars and squad turnout, refreshing every 5 seconds.', SpreadsheetApp.getUi().ButtonSet.OK);
}
