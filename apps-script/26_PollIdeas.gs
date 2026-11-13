/**
 * Clash of Classes -- Apps Script backend
 * File 26: Poll Ideas, Poll Insights and the live Poll Board.
 *
 * POLL IDEAS tab: a bank of short questions that tell the school something real about the
 * student experience: belonging, safety, learning, homework, energy, mornings, lunch, interests,
 * home resources, student voice, motivation, Clash itself, goals and life online. Each one says
 * why we ask and how to use the answer. Every answer comes with an 800#, so staff can look at it
 * by squad and grade (and next to attendance, grades and behavior in Climb-er).
 *   Use | Theme | Kind | Question | Choices | Why We Ask | How To Use It | Repeat | Used On
 *   Repeat "monthly" = a trend question: it comes back about every 4 weeks so you can see change.
 *   Menu "Schedule poll ideas" puts the next ideas on Daily Polls, one per school day with no poll
 *   yet, rotating through the themes, and writes the date in Used On. New ideas from this file are
 *   added to an existing tab on Run setup (by question; nothing you edited is touched).
 * POLL INSIGHTS tab (staff only, menu or after each poll is scored): every poll's answers counted
 *   by squad and grade, with each answer's share. Counts only, never names; 800#s stay on Poll Votes.
 * THE POLL BOARD (?page=pollboard): the projector view, like a Menti screen.
 *   Trivia hides the counts until the day is over so the room can't copy the answer.
 * Keep it safe: these are not screening tools. A question flagged "Plan follow-up first" carries an
 * 800# with a sensitive answer: decide who looks and who follows up before you run it.
 */

const IDEA_HEADERS_ = ['Use', 'Theme', 'Kind', 'Question', 'Choices', 'Why We Ask', 'How To Use It', 'Repeat', 'Used On'];
const INSIGHT_HEADERS_ = ['Date', 'Theme', 'Kind', 'Question', 'Answer', 'Total', 'Share %', 'BLACK', 'GOLD', 'GREY', 'PURPLE', 'Grade 7', 'Grade 8', 'Other Grade'];
const FOLLOW_UP_FLAG_ = 'Plan follow-up first: ';

// [Theme, Kind, Question, Choices, Why We Ask, How To Use It, Repeat]
const POLL_IDEAS_ = [
  // ── Check-ins (they win map zones) ──
  ['Check-ins', 'Pride Day Check-in', 'It\'s Mountaineer Pride Day! Are you wearing purple, gold or Mounts gear?', 'I\'m repping! 💜💛', 'Pride Day turnout by squad.', 'Wins Gym 2 on the map (x2). Shout out the top squad on announcements.', ''],
  ['Check-ins', 'Wheel Check-in', 'Spin Wheel today! Check in to put your squad in the spin.', 'Put us in! 🎡', 'Who joins the Wheel.', 'Counts for the Auditorium on the map (x2).', ''],

  // ── Belonging & connection ──
  ['Belonging', 'School Pulse', 'Do you feel like you belong at Ephrata Middle School?', 'Yes | Most days | Not really | No', 'Belonging is the strongest predictor of attendance and behavior.', 'Track it monthly by grade and squad. A squad or grade sliding = target squad meetings and advisory there.', 'monthly'],
  ['Belonging', 'School Pulse', 'Is there an adult in this building you could go to with a problem?', 'Yes | Not sure | No', 'Connection to at least one adult protects kids.', FOLLOW_UP_FLAG_ + 'pull the "No" 800#s and quietly match each student with a mentor or check-in adult.', 'monthly'],
  ['Belonging', 'School Pulse', 'How many teachers know something about you outside of school?', 'None | 1 | 2 or 3 | 4 or more', 'Relationships, not just names.', 'Share by team at a staff meeting; challenge each team to move the "None" group.', ''],
  ['Belonging', 'School Pulse', 'Did an adult at school say hi to you by name today?', 'Yes | No | I don\'t remember', 'The simplest connection habit.', 'Run it before and after a staff "greet at the door" push to see if it changed.', 'monthly'],
  ['Belonging', 'School Pulse', 'Do you have a friend in most of your classes?', 'Yes | Some classes | Not really', 'Isolation shows up in class first.', 'Students with "Not really" are good candidates for clubs, Ambassadors and squad roles.', ''],
  ['Belonging', 'School Pulse', 'Who do you eat lunch with most days?', 'The same friends | Different people | Mostly alone | It depends', 'Lunch shows social connection.', 'If "Mostly alone" is high, start a lunch bunch or open-table day.', ''],
  ['Belonging', 'School Pulse', 'Do you feel proud to be a Mountaineer?', 'Very | A little | Not really', 'School pride and identity.', 'Pair with Pride Day turnout. Low pride in a grade = more student-led spirit events there.', 'monthly'],
  ['Belonging', 'School Pulse', 'Do you feel like your squad is a team?', 'Yes | Kind of | Not yet', 'Whether squads are building community, not just points.', 'Give Generals their squad\'s result; low = a squad bonding activity at the next meeting.', 'monthly'],

  // ── Safety & climate ──
  ['Safety & Climate', 'School Pulse', 'How do the hallways feel between classes?', 'Calm | Mostly fine | Too crowded | Not safe', 'Where supervision is needed.', 'Compare by grade; move hallway duty to where "Not safe" is highest.', 'monthly'],
  ['Safety & Climate', 'School Pulse', 'Where in school do you feel least comfortable?', 'Hallways | Bathrooms | Cafeteria | Bus | Locker room', 'The hot spots students see that adults may not.', 'Add adult presence or check-ins at the top answer; ask again in a month.', ''],
  ['Safety & Climate', 'School Pulse', 'If you saw someone being picked on, what would you most likely do?', 'Tell an adult | Help them | Stay out of it | Not sure', 'Upstander behavior.', 'Teach upstander moves (Belay Line) where "Stay out of it" is high; re-ask after.', ''],
  ['Safety & Climate', 'School Pulse', 'Do you know how to report something that worries you?', 'Yes | Kind of | No', 'Whether reporting tools are known.', 'Re-teach Lookout and Cairn in homeroom if "No" is high.', ''],
  ['Safety & Climate', 'School Pulse', 'How often do you see kindness at school?', 'Every day | Most days | Sometimes | Rarely', 'A positive climate measure, not just problems.', 'Celebrate on announcements when it climbs; push shout-outs when it drops.', 'monthly'],
  ['Safety & Climate', 'School Pulse', 'How often does drama from online follow you into school?', 'Never | Sometimes | Often', 'Online conflict that spills into the day.', 'High = digital citizenship lesson in Core Extension.', ''],
  ['Safety & Climate', 'School Pulse', 'How is the bus ride for you?', 'Good | OK | Hard | I don\'t ride', 'Buses are a common trouble spot.', 'Share "Hard" by bus route with transportation; add bus expectations talks.', ''],

  // ── Learning & engagement ──
  ['Learning', 'School Pulse', 'How interested were you in your classes this week?', 'Very | Somewhat | Not much | Not at all', 'Engagement over time.', 'Track monthly by grade; share at team meetings next to grades.', 'monthly'],
  ['Learning', 'School Pulse', 'What helps you learn best?', 'Hands-on | Working with others | Teacher explains | On my own | Videos', 'How students say they learn.', 'Share with staff for lesson planning and PD.', ''],
  ['Learning', 'School Pulse', 'When you don\'t understand something in class, what do you usually do?', 'Ask the teacher | Ask a friend | Look it up | Wait | Give up', 'Help-seeking habits.', 'If "Wait" or "Give up" is high, teach help-seeking scripts in Core Extension.', ''],
  ['Learning', 'School Pulse', 'How often do you feel lost in at least one class?', 'Never | Sometimes | Often | Every day', 'Hidden struggle before grades show it.', 'Cross-check the "Often" group with grades; offer Core Extension help.', 'monthly'],
  ['Learning', 'School Pulse', 'Which class do you feel most confident in?', 'Math | ELA | Science | Social Studies | Electives', 'Confidence by subject.', 'Compare with STAR and grades; low confidence + good grades = encouragement needed.', ''],
  ['Learning', 'School Pulse', 'Which class feels hardest right now?', 'Math | ELA | Science | Social Studies | Electives', 'Where to put tutoring.', 'Point homework help and Core Extension at the top answer by grade.', ''],
  ['Learning', 'School Pulse', 'Do you know your STAR goal?', 'Yes | Kind of | No', 'Whether goal setting reaches students.', 'If "No" is high before a STAR window, run goal-setting conferences.', ''],
  ['Learning', 'School Pulse', 'Do you know how you are doing in your classes right now?', 'Yes, all of them | Some of them | Not really', 'Grade awareness.', 'Teach students to check PowerSchool; follow up with the "Not really" group.', ''],
  ['Learning', 'School Pulse', 'How do you like to get feedback on your work?', 'Written notes | Talking with my teacher | A rubric | Just the grade', 'Feedback preferences.', 'Share with staff for feedback PD.', ''],

  // ── Study habits & homework ──
  ['Homework & Study', 'School Pulse', 'How long did homework take you last night?', 'None | Under 1 hour | 1 to 2 hours | Over 2 hours', 'Homework load by grade.', 'If one grade is high, coordinate test and project dates by team.', 'monthly'],
  ['Homework & Study', 'School Pulse', 'When do you usually do homework?', 'Right after school | Evening | Late at night | At school | I don\'t', 'Homework timing and sleep.', 'High "Late at night" next to tardies = talk about routines with families.', ''],
  ['Homework & Study', 'School Pulse', 'Where do you do homework?', 'Quiet spot at home | Noisy spot at home | Somewhere else | At school | I don\'t', 'Study space at home.', 'If many lack a quiet spot, open a before or after school study room.', ''],
  ['Homework & Study', 'School Pulse', 'How do you study for a test?', 'Reread notes | Flashcards | Practice problems | Ask someone | I don\'t', 'Study skills.', 'Teach the strongest methods (practice and flashcards) in Core Extension.', ''],
  ['Homework & Study', 'School Pulse', 'Would you come to a homework help session?', 'Before school | Lunch | After school | No thanks', 'Demand for homework help and when.', 'Schedule help at the most-picked time.', ''],
  ['Homework & Study', 'School Pulse', 'Do you use a planner or calendar for assignments?', 'Every day | Sometimes | Never', 'Organization habits.', 'Teach planning to the "Never" group, especially 7th grade.', ''],
  ['Homework & Study', 'School Pulse', 'What gets in the way of homework most?', 'Phone or games | Chores or family | Sports or activities | I don\'t get it | Nothing', 'Barriers to completion.', '"I don\'t get it" = more help sessions; "Phone" = study habits lesson.', ''],
  ['Homework & Study', 'School Pulse', 'How many days this week did you read for 20 minutes?', 'None | 1 or 2 | 3 or 4 | 5 or more', 'Reading habits.', 'Track next to STAR reading; celebrate squads that read most.', 'monthly'],
  ['Homework & Study', 'School Pulse', 'Do you read for fun outside school?', 'Most days | Sometimes | Rarely | Never', 'Reading for pleasure.', 'Share with the library; book talks and new titles where it is low.', ''],

  // ── Wellbeing & energy (not a screening tool) ──
  ['Wellbeing', 'School Pulse', 'How full is your battery today?', 'Full | Half | Low | Empty', 'A quick energy and mood check.', 'Watch the trend by grade, not single kids; a big "Empty" week = lighten the load or add a fun day.', 'monthly'],
  ['Wellbeing', 'School Pulse', 'How much sleep did you get last night?', 'Under 6 hours | 6 to 7 | 8 to 9 | 10 or more', 'Sleep affects focus, mood and attendance.', 'Share with families in the newsletter; compare with tardies.', 'monthly'],
  ['Wellbeing', 'School Pulse', 'Did you eat breakfast this morning?', 'Yes | No', 'Hungry kids can\'t focus.', 'High "No" = promote breakfast in the classroom or grab-and-go.', 'monthly'],
  ['Wellbeing', 'School Pulse', 'How did you feel about this week at school?', 'Great | Good | OK | Rough', 'A weekly climate check.', 'Run on Fridays; look for weeks and grades that dip.', 'monthly'],
  ['Wellbeing', 'School Pulse', 'What helps you calm down when you are stressed?', 'Music | Moving | Talking to someone | Gaming | Something else', 'Calm-down tools students use.', 'Stock calm corners and the Belay Line with what students pick.', ''],
  ['Wellbeing', 'School Pulse', 'How stressed do you feel about school right now?', 'Not at all | A little | A lot', 'School stress, especially around tests.', 'Run before PSSA and finals; plan calm activities if "A lot" is high.', 'monthly'],
  ['Wellbeing', 'School Pulse', 'How much time did you spend on a screen for fun yesterday?', 'Under 1 hour | 1 to 3 hours | 3 to 5 hours | Over 5 hours', 'Screen time and sleep.', 'Share with families; pair with the sleep question.', ''],
  ['Wellbeing', 'School Pulse', 'Do you get outside or move for at least an hour most days?', 'Yes | Some days | No', 'Activity outside school.', 'Share with PE and Gym Night planning.', ''],
  ['Wellbeing', 'School Pulse', 'Would you like to talk with a counselor about something?', 'No thanks | Maybe | Yes', 'An open door to counseling.', FOLLOW_UP_FLAG_ + 'counselors get the "Yes" and "Maybe" 800#s the same day.', ''],

  // ── Mornings & attendance ──
  ['Mornings & Attendance', 'School Pulse', 'How do you usually get to school?', 'Bus | Car | Walk or bike | Other', 'Arrival planning.', 'Compare with tardies by way of travel; share with transportation.', ''],
  ['Mornings & Attendance', 'School Pulse', 'What makes you late to school, if you are ever late?', 'I\'m never late | Waking up | Ride or bus | Family stuff | Don\'t want to come', 'Root causes of tardies.', 'Match fixes to causes: wake-up plans, bus calls, family outreach.', ''],
  ['Mornings & Attendance', 'School Pulse', 'What would make you want to be here every day?', 'Friends | Fun events | Classes I like | Clubs or sports | Rewards', 'What actually pulls students to school.', 'Use the top answer in attendance campaigns and squad rewards.', ''],
  ['Mornings & Attendance', 'School Pulse', 'What part of the day do you look forward to most?', 'Morning | A favorite class | Lunch | PASS or Core Extension | After school', 'Bright spots in the day.', 'Build on it: put events and shout-outs in the most-liked part.', ''],
  ['Mornings & Attendance', 'School Pulse', 'How was your morning before school today?', 'Smooth | Rushed | Stressful | I don\'t remember', 'Morning routines at home.', 'High "Stressful" = gentle start routines in homeroom.', 'monthly'],
  ['Mornings & Attendance', 'School Pulse', 'If you missed school recently, what was the main reason?', 'I haven\'t missed | Sick | Appointment | Family | Didn\'t want to come', 'Why students miss.', FOLLOW_UP_FLAG_ + '"Didn\'t want to come" 800#s go to the attendance team for a check-in.', ''],

  // ── Lunch, hallways & schedule ──
  ['Lunch & Schedule', 'School Pulse', 'What would make lunch better?', 'More time | Pick my seat | Activities | Quieter | Better food', 'Lunch changes students want.', 'Try the top answer for a month; ask again.', ''],
  ['Lunch & Schedule', 'School Pulse', 'Do you get enough time to eat at lunch?', 'Yes | Barely | No', 'Lunch line and timing.', 'Share with the cafeteria; adjust release order if "No" is high.', ''],
  ['Lunch & Schedule', 'School Pulse', 'How do you use PASS or Core Extension time?', 'Catch up on work | Get help | Read | Hang out | Not sure', 'Whether flex time is working.', 'Share with the MTSS team; tighten expectations if "Hang out" is high.', ''],
  ['Lunch & Schedule', 'School Pulse', 'When is the best time for a squad event?', 'Before school | Lunch | PASS | After school', 'When students can come.', 'Schedule squad events at the top answer.', ''],
  ['Lunch & Schedule', 'School Pulse', 'Do you have enough time between classes?', 'Yes | Most days | No', 'Passing time.', 'High "No" in one wing = look at that route or locker access.', ''],

  // ── Activities & interests ──
  ['Interests', 'School Pulse', 'If we started a new club, which would you join?', 'Esports | Robotics | Art | Cooking | Something else', 'Which clubs to start.', 'Start the top pick; ask "Something else" students for ideas.', ''],
  ['Interests', 'School Pulse', 'Are you in a club, team or activity at school?', 'Yes, one | Yes, more than one | Not now | Never', 'Activity involvement.', 'Invite the "Never" group personally to a club fair or squad role.', ''],
  ['Interests', 'School Pulse', 'What do you do after school most days?', 'Sports | Music or arts | Job or family care | Clubs | Nothing set', 'Free time and family duties.', 'Family care students may need homework flexibility; "Nothing set" = club invites.', ''],
  ['Interests', 'School Pulse', 'What keeps you from joining a club?', 'Nothing, I\'m in one | No ride | Not interested | Don\'t know about them | No time', 'Barriers to joining.', '"No ride" = ask about the activity bus; "Don\'t know" = club spotlight announcements.', ''],
  ['Interests', 'School Pulse', 'Which kind of event would you come to?', 'Game night | Sports night | Talent show | Movie night | Service project', 'What events draw students.', 'Plan Squad Night and family nights around the top picks.', ''],
  ['Interests', 'School Pulse', 'Which Clash event should come back?', 'Mayhem events | Trivia | Spirit days | Lunch battles | Squad Duels', 'Clash planning.', 'Bring back the top pick next marking period.', ''],

  // ── Home & resources ──
  ['Home & Resources', 'School Pulse', 'What do you use for homework at home?', 'Chromebook | Phone | Home computer | Nothing', 'Device needs at home.', 'Students with "Phone" or "Nothing" may need a take-home Chromebook.', ''],
  ['Home & Resources', 'School Pulse', 'Is your internet at home reliable?', 'Yes | Sometimes | No', 'Internet access.', FOLLOW_UP_FLAG_ + 'share "No" 800#s with tech for hotspots or school time.', ''],
  ['Home & Resources', 'School Pulse', 'Do you have the supplies you need for class?', 'Yes | Mostly | No', 'Supply gaps.', 'Stock a quiet supply closet; tell teachers where it is.', ''],
  ['Home & Resources', 'School Pulse', 'Who helps you with homework at home?', 'A parent or guardian | A sibling | Someone else | No one, I do it alone', 'Home support.', 'The "No one" group is a good fit for homework help.', ''],
  ['Home & Resources', 'School Pulse', 'How often do you talk with family about school?', 'Every day | A few times a week | Rarely | Never', 'Family engagement.', 'Share with the newsletter team; try "ask your kid" prompts.', ''],

  // ── Student voice & communication ──
  ['Student Voice', 'School Pulse', 'How do you hear about school events?', 'Announcements | Teachers | Friends | Posters | The website', 'Where messages land.', 'Put the big news where students actually look.', ''],
  ['Student Voice', 'School Pulse', 'Do you feel like students have a say in how school works?', 'Yes | A little | No', 'Student voice.', 'Low = invite students onto a planning group; share results back.', 'monthly'],
  ['Student Voice', 'School Pulse', 'Are the morning announcements helpful?', 'Yes | Sometimes | I don\'t listen', 'Whether announcements work.', 'Shorten or liven them up if "I don\'t listen" wins.', ''],
  ['Student Voice', 'School Pulse', 'If you could change one thing at school, what would it be?', 'Lunch | Schedule | Rules | Clubs and events | Classes', 'Where students want change.', 'Pick the top area for a student focus group.', ''],
  ['Student Voice', 'School Pulse', 'Do you understand the school rules and why we have them?', 'Yes | Most of them | Not really', 'Expectations clarity.', 'Re-teach the Mountaineer Peaks where "Not really" is high.', ''],
  ['Student Voice', 'School Pulse', 'Are school rules applied fairly?', 'Yes | Mostly | Not really | No', 'Fairness, a key to trust.', 'Look by grade; talk with the discipline team if it drops.', 'monthly'],

  // ── Recognition & motivation ──
  ['Recognition', 'School Pulse', 'Which reward would you work hardest for?', 'Extra recess | Food | Free time | A prize | A shout-out', 'What motivates students.', 'Set PBIS and squad rewards to the top picks by grade.', ''],
  ['Recognition', 'School Pulse', 'Has someone recognized something good you did this week?', 'Yes | No | Not sure', 'Recognition reach.', 'High "No" = push shout-outs and Trait Dough in that grade.', 'monthly'],
  ['Recognition', 'School Pulse', 'How do you like to be recognized?', 'In front of everyone | Privately | A note home | A shout-out online | Not at all', 'Recognition style.', 'Share with staff so praise fits the student.', ''],
  ['Recognition', 'School Pulse', 'Do you know how to earn Trait Dough?', 'Yes | Kind of | No', 'Program awareness.', 'Re-teach in homeroom if "No" is high.', ''],

  // ── Clash & squads ──
  ['Clash & Squads', 'School Pulse', 'How much do you care about your squad winning Clash?', 'A lot | Some | Not much | Not at all', 'Clash buy-in.', 'Low in a squad = Generals plan a hype moment; low overall = new event ideas.', 'monthly'],
  ['Clash & Squads', 'School Pulse', 'Do you know who your Squad Generals are?', 'Yes | One or two | No', 'Whether leaders are visible.', 'Introduce Generals at the next squad meeting where "No" is high.', ''],
  ['Clash & Squads', 'School Pulse', 'Do you check clashofclasses.org or My Climb?', 'Every day | Every week | Sometimes | Never', 'Site use.', 'Promote My Climb where "Never" is high.', ''],
  ['Clash & Squads', 'School Pulse', 'What should squad meetings have more of?', 'Games | Planning events | Talking about goals | Service | Celebrating', 'Better squad meetings.', 'Share with Generals before the next meeting.', ''],
  ['Clash & Squads', 'School Pulse', 'Does Clash make school more fun?', 'Yes | A little | No', 'Clash impact.', 'Share with the committee; pair with belonging.', 'monthly'],

  // ── Future & goals ──
  ['Future & Goals', 'School Pulse', 'Do you have a goal for this marking period?', 'Yes, and I wrote it down | Yes, in my head | Not yet', 'Goal setting.', 'Run a goal-setting lesson if "Not yet" is high.', ''],
  ['Future & Goals', 'School Pulse', 'What are you thinking about after high school?', 'College | Trade or tech school | Military | Work | Not sure yet', 'Future plans.', 'Plan career speakers and CTC visits around the answers.', ''],
  ['Future & Goals', 'School Pulse', 'Which career area interests you most?', 'Health | Tech | Building and trades | Arts | Business', 'Career interests.', 'Match Career Day speakers to the top picks.', ''],
  ['Future & Goals', 'School Pulse', 'How ready do you feel for high school?', 'Very | Somewhat | Not yet', 'Transition readiness (8th grade).', 'Use for 8th grade transition planning in the spring.', ''],
  ['Future & Goals', 'School Pulse', 'How do you feel about your grades right now?', 'Proud | OK | Worried', 'Academic confidence.', 'Pair with real grades: "Worried" with good grades needs encouragement, with low grades needs help.', 'monthly'],

  // ── Life online ──
  ['Life Online', 'School Pulse', 'Where do you spend the most time online?', 'Games | Video apps | Group chats | Social media | Not much online', 'Students\' online world.', 'Shape digital citizenship lessons around the top answer.', ''],
  ['Life Online', 'School Pulse', 'Do you know what to do if someone is mean to you online?', 'Yes | Kind of | No', 'Online safety skills.', 'Teach block, save, report, tell an adult where "No" is high.', ''],
  ['Life Online', 'School Pulse', 'Does your phone distract you from school work?', 'Never | Sometimes | A lot', 'Phone habits.', 'Share with staff and families when shaping phone routines.', ''],

  // ── Just for fun (a few keep the habit going) ──
  ['Just for Fun', 'This or That', 'Early bird or night owl?', 'Early bird | Night owl', 'A fun one that also hints at sleep habits.', 'Fun. Compare with the sleep question.', ''],
  ['Just for Fun', 'This or That', 'Pizza or tacos?', 'Pizza | Tacos', 'Just for fun.', 'Use for a squad celebration menu.', ''],
  ['Just for Fun', 'Trivia', 'How many squads are in Clash of Classes?', '3 | 4 | 5', 'Just for fun.', 'Type the Answer (4) on Daily Polls.', ''],
];

function ideaRow_(r) { return [true, r[0], r[1], r[2], r[3], r[4], r[5], r[6], '']; }

// Makes the tabs, adds newer columns to an older Poll Ideas tab, and adds any idea in this file
// whose question isn't on the tab yet (nothing staff edited is changed).
function ensurePollIdeaTabs_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  [['Poll Ideas', IDEA_HEADERS_], ['Poll Insights', INSIGHT_HEADERS_]].forEach(function (t) {
    let sh = ss.getSheetByName(t[0]);
    if (!sh) {
      sh = ss.insertSheet(t[0]);
      sh.getRange(1, 1, 1, t[1].length).setValues([t[1]]).setFontWeight('bold').setBackground('#490e6f').setFontColor('#ffffff');
      sh.setFrozenRows(1);
      return;
    }
    const head = sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn())).getValues()[0].map(function (h) { return String(h).trim(); });
    t[1].forEach(function (h) {
      if (head.indexOf(h) > -1) return;
      head.push(h);
      sh.getRange(1, head.length).setValue(h).setFontWeight('bold').setBackground('#490e6f').setFontColor('#ffffff');
    });
  });
  addNewPollIdeas_();
}

function addNewPollIdeas_() {
  const sh = getSheet_('Poll Ideas');
  const have = {};
  sheetRows_('Poll Ideas').forEach(function (r) { have[String(r['Question'] || '').trim().toLowerCase()] = 1; });
  const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(function (h) { return String(h).trim(); });
  const add = POLL_IDEAS_.filter(function (r) { return !have[r[2].toLowerCase()]; }).map(function (r) {
    const o = {}; IDEA_HEADERS_.forEach(function (h, i) { o[h] = ideaRow_(r)[i]; });
    return head.map(function (h) { return o[h] === undefined ? '' : o[h]; });
  });
  if (add.length) sh.getRange(sh.getLastRow() + 1, 1, add.length, head.length).setValues(add);
  return add.length;
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

// Ideas ready to run on a day: ticked, never used, or a monthly trend question last used 28+ days before.
function readyIdeas_(day) {
  return sheetRows_('Poll Ideas').filter(function (r) {
    if (!ticked_(r['Use']) || !String(r['Question'] || '').trim()) return false;
    const used = String(r['Used On'] || '').slice(0, 10);
    if (!used) return true;
    return /month/i.test(String(r['Repeat'] || '')) && used <= ymdShift_(day, -28);
  });
}

/** Puts ideas on Daily Polls, one per open school day, rotating themes (trend questions first when due). */
function schedulePollIdeas_(days) {
  ensurePollTabs_(); ensurePollIdeaTabs_();
  const taken = {};
  sheetRows_('Daily Polls').forEach(function (p) { taken[String(p['Date'] || '').slice(0, 10)] = true; });
  const open = nextSchoolDays_(days || 5).filter(function (d) { return !taken[d]; });
  const polls = getSheet_('Daily Polls'), ideaSheet = getSheet_('Poll Ideas'), usedCol = headerCol_(ideaSheet, 'Used On');
  const picked = {}, lines = [];
  let lastTheme = '';
  open.forEach(function (d) {
    const ready = readyIdeas_(d).filter(function (r) { return !picked[r._row] && !/check-in/i.test(String(r['Kind'] || '')); });
    if (!ready.length) return;
    // A due trend question first, then a theme different from yesterday's, else whatever is next.
    const due = ready.filter(function (r) { return String(r['Used On'] || '') && String(r['Theme'] || '') !== lastTheme; });
    const fresh = ready.filter(function (r) { return String(r['Theme'] || '') !== lastTheme; });
    const idea = due[0] || fresh[0] || ready[0];
    picked[idea._row] = 1; lastTheme = String(idea['Theme'] || '');
    const k = pollKind_(idea['Kind']);
    polls.appendRow([true, d, idea['Kind'] || 'School Pulse', idea['Question'], idea['Choices'] || '', '', 1, k === 'This or That' || k === 'Trivia' ? 1 : 0, '', '']);
    if (usedCol) ideaSheet.getRange(idea._row, usedCol).setValue(d);
    lines.push(d + ' (' + (idea['Theme'] || idea['Kind']) + '): ' + idea['Question'] + (String(idea['How To Use It'] || '').indexOf(FOLLOW_UP_FLAG_) === 0 ? '  ⚠ plan follow-up first' : ''));
  });
  return lines;
}

function menuSchedulePollIdeas_() {
  const lines = schedulePollIdeas_(5);
  SpreadsheetApp.getUi().alert(lines.length ? 'Put on Daily Polls (Trivia needs its Answer typed in):\n\n' + lines.join('\n') +
    '\n\nCheck-ins (Pride Day, the Wheel) are scheduled by hand on the day.'
    : 'Nothing scheduled: the next 5 school days already have polls, or every ticked idea is used. Add ideas on the Poll Ideas tab.');
}

/** Poll Insights: each poll's answers by squad and grade with each answer's share (counts only). */
function writePollInsights_() {
  ensurePollIdeaTabs_();
  const grade = {}, theme = {};
  sheetRows_(getSetting_('STUDENT_BOARD_TAB') || 'Student Board').forEach(function (r) { grade[id800_(r['800#'])] = String(r['Grade'] || '').replace(/\D/g, ''); });
  sheetRows_('Poll Ideas').forEach(function (r) { theme[String(r['Question'] || '').trim().toLowerCase()] = String(r['Theme'] || ''); });
  const rows = [];
  pollsWithVotes_().filter(function (p) { return p.total; }).sort(function (a, b) { return b.date.localeCompare(a.date); }).forEach(function (p) {
    p.choices.forEach(function (c) {
      const o = { Date: p.date, Theme: theme[p.question.trim().toLowerCase()] || '', Kind: p.kindLabel, Question: p.question, Answer: c, Total: 0, 'Share %': 0,
        BLACK: 0, GOLD: 0, GREY: 0, PURPLE: 0, 'Grade 7': 0, 'Grade 8': 0, 'Other Grade': 0 };
      p.votes.forEach(function (v) {
        if (String(v['Choice']) !== c) return;
        o.Total++;
        const s = squadOf_(v['Squad']); if (o[s] !== undefined) o[s]++;
        const g = grade[id800_(v['800#'])];
        o[g === '7' ? 'Grade 7' : g === '8' ? 'Grade 8' : 'Other Grade']++;
      });
      o['Share %'] = p.total ? Math.round(100 * o.Total / p.total) : 0;
      rows.push(o);
    });
  });
  writeRowsByHeader_(getSheet_('Poll Insights'), rows);
  return rows.length;
}

function menuPollInsights_() {
  const n = writePollInsights_();
  SpreadsheetApp.getUi().alert('Poll Insights is up to date: ' + n + ' answer rows, by theme, squad and grade, with each answer\'s share. ' +
    'Filter by Theme; trend questions (Repeat monthly) show once per date so you can see the change.');
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
