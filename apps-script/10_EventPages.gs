/**
 * Clash of Classes -- Apps Script backend
 * File 10: Event pages (Mountaineer Mayhem, Kick Off, Championships) edited in the Sheet.
 *
 * Event Pages tab: Page | Field | Value | Notes
 *   Page: mayhem, kickoff or champs. Fields: Subtitle, Date, Location, Format, Schedule Label,
 *   Intro, Winner (none = hide the banner), Winner Note, Winner Label, Planning Doc, Planning Doc Label, Audience Expectations,
 *   Participant Expectations. A blank Value keeps what the site's Edit Page panel has.
 *
 * Event Rows tab: Show | Page | Section | Order | Heading | Details | Points / Who | Link
 *   Section Schedule  Heading = time, Details = what happens
 *   Section Events    Heading = event name, Details = rules, Points / Who = scoring
 *   Section Week      Heading = day, Details = homeroom/spirit plan, Points / Who = lunch challenge
 *   Section Night     the Squad Night finals: Heading = event, Details = rules, Points / Who = scoring
 *   Section Links     Heading = label, Link = the address
 *   When a page has Schedule or Events rows, they replace the page's built-in list.
 *
 * Mayhem Week scoreboard: every event on the events tab with type mayhem_week, this school year,
 * one column per event name (e.g. "Paper Chain Challenge"), plus a total.
 */

const EVENT_PAGE_HEADERS_ = ['Page', 'Field', 'Value', 'Notes'];
const EVENT_ROW_HEADERS_ = ['Show', 'Page', 'Section', 'Order', 'Heading', 'Details', 'Points / Who', 'Link'];
const EVENT_PAGES_ = ['mayhem', 'kickoff', 'champs'];

const MAYHEM_WEEK_DOC_ = 'https://docs.google.com/document/d/15NP5-GEWObOM2HTOm46fznCmqlS-_t1T-yTYRxuziaQ/edit';
const DEFAULT_EVENT_FIELDS_ = [
  ['mayhem', 'Subtitle', 'Mayhem Week March 15–19 · Mountaineer Mayhem Friday, March 19', ''],
  ['mayhem', 'Date', 'Friday, March 19, 2027', 'From the Clash Committee agenda (October).'],
  ['mayhem', 'Location', '', 'Blank = what the site shows now (New Gymnasium). Last year: 7th in the New Gym, 8th in the Old Gym.'],
  ['mayhem', 'Format', 'Mayhem Week + Mayhem Friday + Closing Assembly', 'Shown in the third box (was "9 Competitions + Bonus").'],
  ['mayhem', 'Schedule Label', 'Friday, March 19 — last year\'s schedule until this year\'s is set', 'Update when the 2027 schedule is ready.'],
  ['mayhem', 'Intro', 'Mountaineer Mayhem is the biggest week of Clash of Classes. All week, squads battle at lunch in the Color Squad Finals Challenge: Paper Chain, Penny Tower, Trivia, Hula Hoop and Rock/Paper/Scissors. On Friday, wear your Color Squad shirt for the Closing Assembly and Mountaineer Mayhem.', ''],
  ['mayhem', 'Winner', '', 'This year\'s champion (e.g. Room 86). Blank = keep last year\'s banner; type none to hide the banner.'],
  ['mayhem', 'Winner Note', '', 'e.g. 🥈 Runner-Up: Room 71'],
  ['mayhem', 'Winner Label', 'Defending champion · Mountaineer Mayhem 2026', 'The small line above the winner. Change to "Mountaineer Mayhem 2027 Champion" once there is one.'],
  ['mayhem', 'Planning Doc', MAYHEM_WEEK_DOC_, 'The "📄 Planning doc" button. Anyone who opens the site sees the link; the doc\'s own sharing decides who can open it.'],
  ['mayhem', 'Planning Doc Label', 'Mayhem Week plan', ''],
  ['mayhem', 'Audience Expectations', '', 'Blank = keep what the site shows now.'],
  ['mayhem', 'Participant Expectations', '', 'Blank = keep what the site shows now.'],
  ['mayhem', 'Squad Night', 'Friday, March 19 · evening in the gym', 'When and where the Squad Night finals are. Blank hides the Squad Night section.'],
  ['mayhem', 'Squad Night Entry', 'Free for anyone competing in a Mayhem Week event or wearing a Mayhem shirt they bought. $5 for everyone else. Student Council concessions.', 'Shown under the Squad Night heading. The door check-in uses Settings › SQUAD_NIGHT_PRICE.'],
  ['kickoff', 'Subtitle', 'Clash of Classes Kick Off · Friday, September 25, 2026', ''],
  ['kickoff', 'Date', 'Friday, September 25, 2026', ''],
  ['kickoff', 'Location', '', ''],
  ['kickoff', 'Format', 'Squad competitions + rubric voting', ''],
  ['kickoff', 'Schedule Label', 'An earlier Kick Off schedule — add this year\'s on the Event Rows tab', 'Once you add Schedule rows for kickoff on Event Rows, they replace the cards below; then change this label.'],
  ['kickoff', 'Intro', 'The Kick Off launched this year\'s Clash of Classes. 7th graders had their Cornerstone day the same day, and plenty of them asked how they can get involved. The next all-school event is the Closing Assembly on Mountaineer Mayhem Friday, March 19.', ''],
  ['kickoff', 'Winner', '', ''],
  ['kickoff', 'Winner Note', '', ''],
  ['kickoff', 'Planning Doc', '', 'Optional: the Kick Off plan or slides.'],
  ['kickoff', 'Planning Doc Label', 'Kick Off plan', ''],
  ['champs', 'Planning Doc', '', ''],
  ['champs', 'Planning Doc Label', 'Championships plan', ''],
];

// [Show, Page, Section, Order, Heading, Details, Points / Who, Link]
const DEFAULT_EVENT_ROWS_ = [
  [true, 'mayhem', 'Week', 1, 'Monday, March 15 · Mind Games', 'Trivia Bowl · Spoons · Escape Room · Chess Tournament', 'Lunch: Paper Chain Challenge (3 per squad)', ''],
  [true, 'mayhem', 'Week', 2, 'Tuesday, March 16 · Skills & Stunts', 'Lip Sync / Glee Battle · Cupcake Wars · Badminton · Mario Kart', 'Lunch: Penny Tower Race (2 per squad)', ''],
  [true, 'mayhem', 'Week', 3, 'Wednesday, March 17 · Athletic Day', 'Volleyball · Smash Brothers · Basketball Knockout · Noodle Relay · Minute to Win It', 'Lunch: Trivia (2 per squad)', ''],
  [true, 'mayhem', 'Week', 4, 'Thursday, March 18 · Field Events', 'Skirmish / Dodgeball · Shuttle Race · Arm Wrestling · Standing Long Jump · Rope Climb (one field event each) · Tug of War', 'Lunch: Hula Hoop Contest (1 per squad)', ''],
  [true, 'mayhem', 'Week', 5, 'Friday, March 19 · Mayhem Day', 'Wear your Color Squad shirt · Closing Assembly · Mountaineer Mayhem · Squad Night finals that evening', 'Lunch: Rock / Paper / Scissors (3 per squad)', ''],
  [true, 'mayhem', 'Night', 1, 'Obstacle Course', 'Relay style. Every station must be touched; miss one and go back to it.', '6 per squad · 5 pts to the winner', ''],
  [true, 'mayhem', 'Night', 2, 'Squad Dance', 'Same song for every squad, 2.5 minutes, squad shirts on. Judged with a rubric.', '12–20 per squad · 0–12 pts judged', ''],
  [true, 'mayhem', 'Night', 3, 'Tricycle Relay', 'Relay across the gym floor, alternating ends.', '10 per squad · 2 pts to the winner', ''],
  [true, 'mayhem', 'Night', 4, 'Running Relay', 'Baton relay, 3 runners at each end, each runner goes twice.', '6 per squad · 5 pts to the winner', ''],
  [true, 'mayhem', 'Night', 5, 'Speed Pyramids', 'Best 2 of 3. The top person is up on knees, hands in the air, steady for 3 seconds.', '10 per squad · 5 pts to the winner', ''],
  [true, 'mayhem', 'Night', 6, 'Tire Pull', 'Best of 9, first to 5, 30-second limit. A hand on the tire to win.', '12 per squad · 5 pts to the winner', ''],
  [true, 'mayhem', 'Night', 7, 'Lucky Chance Spin', 'The Clash of Classes Spin Wheel closes the night for all four squads, with results added live.', 'Up to 50 pts', ''],
  [true, 'mayhem', 'Night', 8, 'Poster reveal and final standings', 'Squad posters judged (0–12), the participation award (10 pts to the squad with the most participants all week), the sportsmanship award (5 pts) and t-shirt sales (5 pts to the top squad).', '', ''],
  [true, 'mayhem', 'Links', 1, 'Mayhem Week plan', '', '', MAYHEM_WEEK_DOC_],
  [true, 'mayhem', 'Links', 2, 'Last year\'s Mayhem master doc (2026)', '', '', 'https://docs.google.com/spreadsheets/d/1M1USTU2PWcx0tdjTOWQyg3NW9a02JBCNbfpj11q_TsI/edit'],
  [true, 'mayhem', 'Links', 3, 'Last year\'s competition slides (2026)', '', '', 'https://docs.google.com/presentation/d/19ZyXKF3HE4JwCR3GHtG_33rAxTF_e4egPLfDjkcBCvM/edit'],
  [true, 'mayhem', 'Schedule', 1, '7:50 – 8:10', 'Arrival & Announcements — homeroom review of event rules; participant tags handed out', '', ''],
  [true, 'mayhem', 'Schedule', 2, '8:11 – 8:30', 'Directions & assignments for Mayhem — make team signs & chants', '', ''],
  [true, 'mayhem', 'Schedule', 3, '8:30 – 8:40', '7th Grade — RA Hallway to the New Gym, sit in your section | 8th Grade — Old Gym/Lobby, enter the New Gym from the Old Gym side', '', ''],
  [true, 'mayhem', 'Schedule', 4, '8:40 – 8:55', 'Welcome & Opening — live points display; students to the on-deck area', '', ''],
  [true, 'mayhem', 'Schedule', 5, '8:55 – 9:05', 'Event 1: Hungry Hungry Hippos', '', ''],
  [true, 'mayhem', 'Schedule', 6, '9:10 – 9:20', 'Event 2: Teacher Pop Darts', '', ''],
  [true, 'mayhem', 'Schedule', 7, '9:25 – 9:35', 'Event 3: Noodle Knockdown', '', ''],
  [true, 'mayhem', 'Schedule', 8, '9:35 – 9:45', 'Sort it Right (between events)', '', ''],
  [true, 'mayhem', 'Schedule', 9, '9:45 – 9:55', 'Event 4: Ping Pong Ball Pass', '', ''],
  [true, 'mayhem', 'Schedule', 10, '10:00 – 10:15', 'Event 5: Pack It Up (Hold It Relay)', '', ''],
  [true, 'mayhem', 'Schedule', 11, 'Between events', 'Chuck a Chicken — 1 rep per homeroom throws from the bleachers', '', ''],
  [true, 'mayhem', 'Schedule', 12, '10:25 – 10:35', 'Event 6: Tower Toss', '', ''],
  [true, 'mayhem', 'Schedule', 13, '10:40 – 10:50', 'Event 7: Basketball Knockout', '', ''],
  [true, 'mayhem', 'Schedule', 14, '10:50 – 10:55', 'Lucky Chance — spin to win up to 50 squad points', '', ''],
  [true, 'mayhem', 'Schedule', 15, '11:00', 'Dismiss — 7th Grade to lunch, 8th Grade to homeroom', '', ''],
  [true, 'mayhem', 'Events', 1, 'Hungry Hungry Hippos', 'Teams (MOUNTS format) compete around the center volleyball court with scooters, gator balls, buckets, cones and cardboard on a 4-minute timer. One grade level at a time; 6–8 participants per homeroom.', '1 point per ball collected', ''],
  [true, 'mayhem', 'Events', 2, 'Teacher Pop Darts', 'One teacher per team throws from the purple bleacher side toward the basketball key; the target faces the crowd. 2 attempts per team, one grade level at a time.', '10 / 8 / 6 / 4 / 2 / 1 per standing dart. Grade-level winner earns 5 bonus points.', ''],
  [true, 'mayhem', 'Events', 3, 'Noodle Knockdown', '1 contestant per homeroom balances a cup on a pool noodle and tries to knock opponents\' cups off — noodle only, no body contact. Win your group to advance; the 7th and 8th grade winners face off in the center circle.', 'Out in round 1 = 2 · round 2 = 5 · runner-up = 8 · winner = 12', ''],
  [true, 'mayhem', 'Events', 4, 'Ping Pong Ball Pass', '6 contestants per homeroom pass ping pong balls cup-to-cup into a bucket — pass, don\'t dump. A dropped ball restarts. 2-minute round, grade levels separate.', '1 point per ball in the bucket', ''],
  [true, 'mayhem', 'Events', 5, 'Pack It Up (Hold It Relay)', '4 contestants per homeroom relay items one at a time to a holder at half court. After time, the holder keeps everything secured for 10 more seconds — only those count. Top 6 teams score.', '2 points per item held', ''],
  [true, 'mayhem', 'Events', 6, 'Tower Toss', 'Relay-style with boxes on the volleyball sideline. 6 per squad (3 from 7th, 3 from 8th), 7 marked boxes per squad, 6-minute cap.', 'First to finish = 50 · finished in time = 25', ''],
  [true, 'mayhem', 'Events', 7, 'Basketball Knockout', '1 contestant per homeroom, shooting from the 3-point line. If the player behind you scores first, you\'re out. Each grade plays at its own end until one player remains.', '15 / 12 / 10 / 9 / 8 / 7 / 6 / 5 / 4 / 3 / 2 / 1 by finish', ''],
  [true, 'mayhem', 'Events', 8, 'Lucky Chance (bonus)', 'Near the end, each squad spins as a group.', 'Up to 50 squad points', ''],
];

function ensureEventPageTabs_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const make = function (name, headers, seed) {
    if (ss.getSheetByName(name)) return;
    const sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#490e6f').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    if (seed.length) sh.getRange(2, 1, seed.length, headers.length).setValues(seed);
  };
  make('Event Pages', EVENT_PAGE_HEADERS_, DEFAULT_EVENT_FIELDS_);
  make('Event Rows', EVENT_ROW_HEADERS_, DEFAULT_EVENT_ROWS_);
}

function safeLink_(v) {
  const s = String(v || '').trim();
  return /^https:\/\//i.test(s) ? s : '';
}

// Mayhem Week scoreboard: mayhem_week events since August 1 of this school year.
function mayhemWeekScores_() {
  const now = new Date();
  const year = now.getMonth() >= 7 ? now.getFullYear() : now.getFullYear() - 1;
  const since = year + '-08-01';
  const cols = [], scores = {};
  readEventsTab_().forEach(function (e) {
    if (String(e.type) !== 'mayhem_week' || String(e.date || '') < since) return;
    const name = String(e.name || '').trim() || 'Mayhem Week';
    if (cols.indexOf(name) === -1) { cols.push(name); scores[name] = { BLACK: 0, GOLD: 0, GREY: 0, PURPLE: 0 }; }
    SQUAD_KEYS_.forEach(function (s) { scores[name][s] += Number(e.pts[s]) || 0; });
  });
  return cols.length ? { columns: cols, scores: scores } : null;
}

/** Everything the three event pages need, from the two tabs. Public (no names). */
function clashEventPages() {
  const out = {};
  EVENT_PAGES_.forEach(function (p) { out[p] = { fields: {}, schedule: [], events: [], week: [], night: [], links: [] }; });
  sheetRows_('Event Pages').forEach(function (r) {
    const p = String(r['Page'] || '').trim().toLowerCase();
    const f = String(r['Field'] || '').trim();
    const v = String(r['Value'] == null ? '' : r['Value']).trim();
    if (!out[p] || !f || !v) return;
    out[p].fields[f] = /^Planning Doc$/i.test(f) ? safeLink_(v) : v;
  });
  const sections = { schedule: 'schedule', events: 'events', week: 'week', night: 'night', links: 'links' };
  sheetRows_('Event Rows').filter(function (r) { return ticked_(r['Show']); })
    .sort(function (a, b) { return (Number(a['Order']) || 0) - (Number(b['Order']) || 0); })
    .forEach(function (r) {
      const p = String(r['Page'] || '').trim().toLowerCase();
      const sec = sections[String(r['Section'] || '').trim().toLowerCase()];
      if (!out[p] || !sec) return;
      const row = { heading: String(r['Heading'] || '').trim(), details: String(r['Details'] || '').trim(),
        points: String(r['Points / Who'] || '').trim(), link: safeLink_(r['Link']) };
      if (sec === 'links' && !row.link) return;
      if (row.heading || row.details) out[p][sec].push(row);
    });
  out.mayhem.scoreboard = mayhemWeekScores_();
  return JSON.stringify(out);
}
