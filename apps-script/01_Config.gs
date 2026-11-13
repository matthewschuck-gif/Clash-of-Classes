/**
 * Clash of Classes -- Apps Script backend
 * File 1 of 2: Config + one-time spreadsheet provisioning
 *
 * SETUP STEPS (do these once, in order -- see MIGRATION_RUNBOOK.md for the full walkthrough):
 * 1. Create a new Google Sheet in your school Workspace Drive (any name, e.g. "Clash of
 *    Classes Data").
 * 2. Open it -> Extensions -> Apps Script. This creates a bound Apps Script project.
 * 3. Paste each 0N_*.gs file from this package into the Apps Script editor as its own file
 *    (File -> New -> Script file, name it to match, e.g. "01_Config").
 * 4. Project Settings -> Script Properties, add:
 *      APP_TOKEN  = <any random string you make up> (shared token, see 02_Router.gs)
 * 5. Run `setupSpreadsheet` once from the function dropdown. Approve the permission prompts
 *    (this includes a Drive scope now, used only to read the spreadsheet's last-modified time
 *    for change detection -- see getAppDataMeta_ in 02_Router.gs). This creates the `app_data`
 *    and `events` tabs with correct headers.
 * 6. Deploy ONCE as a Web App: Execute as "Me", Who has access "Anyone". Copy the resulting
 *    /exec URL -- it goes into index.html (see 02_Router.gs header for why only ONE
 *    deployment, never a domain-restricted one).
 *
 * ARCHITECTURE NOTE -- why this differs from a typical multi-table migration:
 * Clash of Classes doesn't use several relational tables like Trail Journal did. Almost the
 * entire site's state (squads, gallery, links, theme, spin wheel, admin password, etc.) is one
 * big JSON object, called `D` in the frontend, stored as a SINGLE row in a Postgres table
 * `app_data` (columns: key, value jsonb, updated_at) under key='coc'. The Sheet replacement
 * mirrors that table 1:1 in shape (see SCHEMA.app_data below) but adds one wrinkle Trail
 * Journal never needed: Google Sheets cells cap out at 50,000 characters, and `D` can exceed
 * that once gallery photos (uploaded as base64 data-URIs, see FileReader/readAsDataURL in
 * index.html) are added. So the JSON string is CHUNKED across multiple rows instead of living
 * in one cell -- see chunk_index below and saveAppData_/loadAppData_ in 02_Router.gs.
 *
 * ONE PIECE OF `D` GETS ITS OWN TAB, ON PURPOSE: `D.events` -- the squad point ledger -- is
 * pulled OUT of the JSON blob and given a real, human-editable tab (`events`, one row per
 * event, plain columns for each squad's points). That's what makes "just edit a cell in the
 * Sheet and the site updates" actually work for scores: nobody can usefully hand-edit a number
 * buried inside an opaque chunked JSON blob, but editing a plain number in a spreadsheet column
 * is exactly the point-of-Sheets-as-a-backend experience. See readEventsTab_/writeEventsTab_ in
 * 02_Router.gs for how this splits out of, and merges back into, `D` transparently -- the
 * frontend never needed to change for this, it still just sends/receives one `D` object.
 */

const TABS = {
  APP_DATA: 'app_data',
  EVENTS: 'events',
  METRICS: 'metrics',
  HELP: 'How To Use This Sheet',
};

// Mirrors the original Postgres `app_data` table's columns (key, value, updated_at), plus
// chunk_index -- the one addition needed to get around the Sheets cell-size limit (see note
// above). There is no separate "save_id" column because the original code already stores its
// own echo-suppression id (`_saveId`) INSIDE the `value` JSON itself, not as a separate
// Postgres column -- so it survives the migration for free, no schema change needed.
//
// `events` mirrors the shape of each object already pushed into D.events by the frontend
// (see addEvent() etc. in index.html: {id, name, type, date, note, pts:{BLACK,GOLD,GREY,
// PURPLE}, multiplier}) -- just flattened into plain columns instead of a nested pts object,
// which is what makes it directly editable in the Sheets UI.
const SCHEMA = {
  [TABS.APP_DATA]: ['key', 'chunk_index', 'value_chunk', 'updated_at'],
  [TABS.EVENTS]: ['id', 'name', 'type', 'date', 'note', 'multiplier', 'BLACK', 'GOLD', 'GREY', 'PURPLE'],
  // One row per squad. Mirrors D.metrics.{attendance,referral,grades}[SQUAD] -- the current,
  // in-progress month's percentages (see the monthly-metrics ranking design in index.html:
  // rankPoints_/finalizeMonth). Editing a percentage here updates the site's live preview the
  // same way editing it in Admin -> Metrics does; it does NOT award points by itself. Points
  // only lock in when an admin clicks "Finalize This Month" on the site, same as if the
  // numbers had been typed into the Admin form instead of this tab. The month LABEL (e.g.
  // "September 2026") is intentionally not a column here -- it stays admin-entered on the site
  // since it changes far less often than the percentages do.
  [TABS.METRICS]: ['squad', 'attendance', 'referral', 'grades'],
};

// Stay safely under Sheets' 50,000-char single-cell limit.
const CHUNK_SIZE = 40000;

function setupSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(SCHEMA).forEach(function (tabName) {
    let sheet = ss.getSheetByName(tabName);
    if (!sheet) sheet = ss.insertSheet(tabName);
    const headers = SCHEMA[tabName];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#490e6f').setFontColor('#ffffff');
  });
  const def = ss.getSheetByName('Sheet1');
  if (def && ss.getSheets().length > 1) ss.deleteSheet(def);
  ensureMetricsRows_();
  ensureSettings_();
  ensureChallenges_();
  ensureGeneralsTabs_();
  ensurePollTabs_();
  ensurePollIdeaTabs_(); // Poll Ideas + Poll Insights (26_PollIdeas.gs)
  ensureFreezeTab_();
  ensureDuelTab_();
  ensureLoreTab_();
  ensureEventPageTabs_();
  ensureMayhemWeekTabs_();
  ensureShirtDesignTabs_();
  ensureSchoolCalendarTab_();
  ensureGameTabs_();
  ensureMeTabs_();
  ensureRosterTab_();
  ensureTourTab_(); // 20_Tour.gs
  ensureWeekTabs_(); // 22_Week.gs
  ensureMindsetLinksTab_(); // 23_Drive.gs
  ensureYearPlanTab_(); // 24_YearPlan.gs
  ensureHomeroomTabs_(); // 25_Homerooms.gs
  installBackgroundTick_(); // 03_Web.gs: claims, thaws and a warm cache every 5 minutes
  setupHelpTab_();
  setupMapGuideTab_(); // 18_Conquest.gs
  setupClimbGuideTab_(); // 19_Evolve.gs
  try { SpreadsheetApp.getUi().alert('Done -- app_data, events, metrics, Settings and "How To Use This Sheet" tabs are set up. Clash of Classes has no separate PII tab (unlike Trail Journal), so there is no extra sharing/protection step required here. You can now add/edit rows directly in the events and metrics tabs and the site will pick them up within a few seconds. Staff directions live in the "How To Use This Sheet" tab -- start there.'); } catch (e) { console.log("Setup done (no pop-up when run from the editor).") } // the editor's Run button can't show pop-ups
}

// ── SETTINGS TAB (Key | Value | Notes) ──────────────────────────────
// Defaults live here; Run setup adds missing rows and never overwrites a value someone set.
// Read fresh on every call, so edits show on the next page load with no redeploy.
const SETTINGS_TAB = 'Settings';
const DEFAULT_SETTINGS = [
  ['SITE_TITLE', 'Clash of Classes — Ephrata Middle School', 'Browser tab title.'],
  ['SITE_URL', 'https://clashofclasses.org', 'The address students use (the menu\'s "Open the site" link).'],
  ['POLL_SECONDS', '5', 'How often an open page checks for new points (seconds, 3 or more). Higher = fewer requests.'],
  ['TJ_SHEET_ID', '1rinTrtlmJKQrvtn5GVs-ngmaQRvkl8Nt39yOBICgggA', 'Trail Journal Data Sheet (link or ID) for the Community Culture Stats panel. Only counts are read.'],
  ['STUDENT_BOARD_TAB', 'Student Board', 'Tab that Define a Climb-er fills every hour with each student\'s positive points and badges. The My Squad page reads it.'],
  ['MY_SQUAD_LOOKUPS_PER_10MIN', '300', 'Most My Squad look-ups the whole site allows in 10 minutes (stops anyone paging through every 800#).'],
  ['BADGE_LIST', '🍞 First Trait Dough; 🍞 Trait Dough Ten; 🏆 Trait Dough Winner; ⛰️ Mindset Summit; ⛰️ Triple Summit; 📈 STAR Grower; 🚀 STAR Climber; 🙌 Shout-out; 🔥 3-Week Streak; 🔥 6-Week Streak; 🤝 Ambassador', 'Every badge, in order, so My Squad can show the ones still to earn. Keep these names the same as BOARD_BADGES on the Climb-er Settings tab (that is where the rules live).'],
  ['CHALLENGE_BONUSES', 'OFF', 'ON = a squad that reaches a challenge goal on the Challenges tab gets its Bonus Points once, automatically (checked hourly). Turn on with Clash of Classes › Turn on challenge bonuses.'],
  ['TV_SHOW_NAMES', 'ON', 'ON = the hallway TV shows a "Badge leaders" slide with first name + last initial. OFF = no student names on the TV.'],
  ['GENERALS_CODE', '', 'The code Squad Generals type with their 800# to open Generals HQ (Role = General on the MLC Roster tab). Blank = Generals HQ is closed. Change it any time.'],
  ['DAILY_POLL_POINTS', 'OFF', 'ON = after each Daily Poll day, squads get points for every student who voted (plus more for the winning side or right trivia answer). Turn on with Clash of Classes › Turn on daily poll points.'],
  ['POLL_VOTES_PER_10MIN', '1500', 'Most Daily Poll votes the site takes in 10 minutes (stops someone flooding it).'],
  ['EVENT_LABELS', 'trait_dough=💵 Trait Dough, mountaineer_mindset=⛰️ Mindset, trail_journal=🏔️ Trail Journal, trailback_panel=🧭 Trailback Panel, star_growth=📈 STAR Growth, camp_mountaineer=⛺ Camp Mountaineer, daily_poll=🗳️ Daily Poll, challenge_bonus=🏁 Challenge, duel_bonus=⚔️ Squad Duel, mayhem_week=🏔️ Mayhem Week, shirt_sales=👕 Shirt Sales, squad_night=🌙 Squad Night, community_bonus=🎯 Bonus', 'Event type = how it is labelled on the TV feed and My Squad.'],
  ['PRIVATE_EVENT_TYPES', 'trail_journal, trailback_panel, camp_mountaineer', 'Event types that are about behavior or support: the public site and TV show only their label, never a student name or note. Staff signed in as admin still see the full rows.'],
  ['DUEL_BONUSES', 'OFF', 'ON = the day after each Squad Duel ends, the winning squad(s) get its Bonus Points (checked hourly). Turn on with Clash of Classes › Turn on Squad Duel bonuses.'],
  ['PASSPORT_STAMPS', 'bills=🍞 Turn in Trait Dough; mindset=⛰️ Reach a Mindset summit; star=📈 Grow on STAR; shoutouts=🙌 Get a shout-out; ambassador=🤝 Be a Camp buddy; polls=🗳️ Vote in Daily Polls; events=🎟️ Sign up for a squad event', 'Trail Passport stamps on My Squad: key=label; … Keys: bills, mindset, star, shoutouts, ambassador, polls, events. Remove one to drop that stamp. All stamps = 🧭 Trail Passport complete.'],
  ['SIGNUP_LIMIT', '3', 'Most events any student can be signed up for at a time, Generals and Captains included (counts only events that are Open on the Sign-up Events tab). 0 = no limit.'],
  ['SELF_SIGNUPS', 'OFF', 'ON = students can sign themselves up on My Squad for events with Self Sign-up ticked (Sign-up Events tab). OFF = only Generals sign students up.'],
  ['NOMINATE_FOR', 'Squad Captain (7th grade); Squad Captain (8th grade); Trailback Panel (Vanguard); Next year\'s Squad General', 'What Generals can nominate for in Generals HQ, separated by ;. Anything with "Captain" in it shows each squad\'s vote tally.'],
  ['SHIRT_VOTING', 'OFF', 'ON = students vote on My Squad for their squad\'s Mayhem shirt design (approved designs on the Shirt Designs tab). One vote each; they can change it until you turn this OFF.'],
  ['SHIRT_DESIGNS_PER_SQUAD', '6', 'Most designs one squad\'s Generals can upload.'],
  ['SHIRT_DESIGN_FOLDER_ID', '', 'The private Drive folder for uploaded designs. Blank = made automatically on the first upload.'],
  ['SHIRT_POINTS', '2', 'Squad points per paid Mayhem shirt (Shirt Orders tab). Sent with Clash of Classes › Send shirt sale points.'],
  ['SQUAD_NIGHT_PRICE', '$5', 'Entry for Squad Night if a student isn\'t competing and didn\'t buy a shirt.'],
  ['SQUAD_NIGHT_POINTS', '1', 'Squad points per student who comes to Squad Night. Sent with Clash of Classes › Send Squad Night attendance points.'],
  ['PASSPORT_POLL_VOTES', '5', 'Daily Poll votes a student needs for the 🗳️ passport stamp.'],
  ['FUN_THAW_TIERS', '1=25, 2=50, 3=100', 'Fun Freeze: reaching the thaw goal once gets back 25% of the points frozen out, twice 50%, three times 100% ("times the goal = percent back").'],
  ['FUN_SECOND_CHANCE_DAYS', '2', 'Fun Freeze: school days a squad\'s one second chance adds to Thaw By.'],
  ['GAME_AWARDS', 'OFF', 'ON = the game layer gives points: weekly territory, round goals, the Opponent, destination tickets (checked hourly). Turn on with Clash of Classes › Turn on game awards. OFF = the map, cards and goals still show; no bonus points.'],
  ['GAME_MAP', 'ON', 'OFF hides the school map (King of the Mountain, territories, wings).'],
  ['GAME_CARDS', 'ON', 'OFF hides Chance cards.'],
  ['GAME_STREAK', 'ON', 'OFF hides Hot Streaks.'],
  ['GAME_ROUNDS', 'ON', 'OFF hides round goals.'],
  ['GAME_ENGINE', 'ON', 'OFF hides engine powers.'],
  ['OPPONENT_VISIBLE', 'ON', 'OFF hides the Opponent everywhere (one click: Clash of Classes › Show or hide the Opponent). Each Opponent also needs Show ticked on the Opponents tab.'],
  ['CLAIM_DAYS', '14', 'School map: a zone goes to the squad with the most points from its events in the last this-many days.'],
  ['MAP_STYLE', 'puzzle', 'School map look: puzzle (the building cut into rough jigsaw pieces), clean (plain areas like a floor plan) or mountain (separate hand-cut blobs).'],
  ['WING_MULTIPLIER', '2', 'School map: holding every zone in a wing pays that wing this many times.'],
  ['CHANCE_DRAWS_PER_WEEK', '1', 'Chance cards each squad\'s Generals can draw a week (engine powers can add more).'],
  ['STREAK_BANK', '2=2, 3=4, 4=7, 5=11, 6=16', 'Hot Streak: days in a row = points when banked.'],
  ['TICKETS_HELD', '3', 'Destination tickets a student can hold at once.'],
  ['TICKET_CHOICES', '3', 'Destination tickets offered per draw (the student keeps at least one).'],
  ['AVATAR_GEAR', 'boots=bills>=1; pack=bills>=5; rope=mindset>=1; lantern=shoutouts>=1; compass=star>=1; scarf=streak>=3; badge=tickets>=3; flag=passport>=1', 'My Climb: gear a student\'s climber earns. gear=thing>=number; things: bills, mindset, shoutouts, star, streak, buddy, tickets, passport.'],
  ['LRG_POINTS', 'bills=1; mindset=5; shoutouts=2; star=1; streak=2; buddy=3; tickets=3; passport=5; journal=2', 'My Climb: LRG (Life Ready Graduate) points a student earns for each thing they do. thing=points; things: bills, mindset, shoutouts, star, streak, buddy, tickets, passport, journal (Trail Journal growth, once Define a Climb-er shows it). Their creature evolves at EVOLVE_LRG.'],
  ['EVOLVE_LRG', '2=30, 3=60, 4=100', 'My Climb: the LRG points a creature needs to evolve to stages 2, 3 and 4. Each evolution lets the student pick a boost (Evolution Boosts tab).'],
  ['EVOLVE_NAMES', 'Foothills, Ridge, High Camp, Summit', 'My Climb: names for the four creature stages.'],
  ['VERSION_NAMES', 'Version 1, Version 2, Version 3', 'My Climb: names for the three looks a student can pick (Version 3 is the winter coat).'],
  ['GEAR_WORN', '3', 'My Climb: how many pieces of earned gear a student can wear on their creature at once (1-4).'],
  ['CLIFF_NAME', 'Cliff', 'My Climb: the guide\'s name (the Mountaineer mascot).'],
  ['TOUR_AUTO', 'ON', 'Cliff\'s tour (Tutorial Steps tab) plays once on each device the first time someone opens the site or My Climb. OFF = only when they tap Ask Cliff.'],
  ['TOUR_VERSION', '1', 'Change this (2, 3…) after editing the Tutorial Steps tab to show everyone the new tour once.'],
  ['CLIFF_IMAGE', '', 'My Climb: an https link to a different picture of Cliff. Blank = the built-in one.'],
  ['MARKING_PERIODS', 'MP1=2026-08-31..2026-10-30; MP2=2026-11-02..2027-01-15; MP3=2027-01-18..2027-04-02; MP4=2027-04-05..2027-06-09', 'Marking periods, from the district calendar (subject to change for closings): name=first day..last day, separated by semicolons. Squad meeting slides and the beat-our-own-MP bonus use them.'],
  ['BEAT_MP_BONUS', '50', 'Points for every squad that out-earns its own previous marking period (paid once after each marking period ends, with the game awards). 0 = off.'],
  ['DIGEST', 'OFF', 'The weekly This Week in Clash email. Turn on with Clash of Classes › Turn on the weekly digest email.'],
  ['DIGEST_EMAILS', '', 'Who gets the weekly digest email: addresses separated by commas.'],
  ['DIGEST_DAY', 'FRIDAY', 'Which weekday the digest goes out (MONDAY to FRIDAY).'],
  ['DIGEST_HOUR', '14', 'About what hour the digest goes out (24-hour clock, 5 to 20).'],
  ['LEADERBOARD_NAMES', 'OFF', 'Leaderboards page: ON adds Top Climbers, the top students per grade by LRG points, shown as First L. OFF = squad leaderboards only.'],
  ['LEADERBOARD_SIZE', '10', 'Leaderboards page: how many Top Climbers per grade (3-25).'],
  ['PITCHED_IN_DAYS', '30', 'Leaderboards page: a student counts as pitching in if they did anything in this many days.'],
  ['TRAIT_DOUGH_APP', 'https://script.google.com/a/macros/easdpa.org/s/AKfycbz1P6NEegeJfY73L3GTA8vKB0oP5M2V4l8nQigjtaRapjTPZsz-DBmlkMVlRkTXe9wH/exec', 'The Trait Dough web app (bit.ly/ems-traitdough). Its info page shows inside the site\'s Trait Dough page.'],
  ['SOTM_APP', 'https://script.google.com/a/macros/easdpa.org/s/AKfycbzt9HZ0wCw7tLXy7ZBl2ZJkCQfM5xVZzpPqyxV-NbVZL0Hm_xQBXEQyFQjJuTRPrf11/exec', 'The Student of the Month web app (bit.ly/ems-sotm). Shows inside the site\'s Student of the Month page.'],
  ['MINDSET_APP', 'https://script.google.com/a/macros/easdpa.org/s/AKfycbwuQCzZfVvgX-k3czsKI74DKnt2qF2XJslElr3HHcwnJTyCcHI4e7NxegIPDpElcowL7w/exec', 'The Grade 8 Mountaineer Mindset web app (bit.ly/ems8-mindset). Shows inside the site\'s Mindset page.'],
  ['GENERALS_DOC', 'https://docs.google.com/document/d/1OeKBx8SOrfkG4zaHuj66LllR1nS6EINA8Dk4KYmM6Zo/edit', 'The Squad Leadership doc (shared Clash drive) that lists the Generals for each squad. Menu › Load Squad Generals copies them onto MLC Roster.'],
  ['GENERALS_HANDBOOK', 'https://docs.google.com/document/d/1DL1VXqEK9OhvJRBPztQvOwadA-UcWpDCg_ZUabP7304/edit', 'The Squad General one-pager / handbook. Generals HQ › My job links to it.'],
  ['FOLLOWUP_STATS', 'ON', 'Map (Cafeteria): ON = once a day, count last month: students per squad with an LRR or a Lunch Detention, from the same sheets EMS Follow-Up reads, onto Squad Stats.'],
  ['LRR_SHEET', '1g8VpAeyE8md-1us58uHZxtRAKKFCTCKHhuwC8ZOpzyc', 'The Lunchtime Redirection Reflection responses sheet (link or ID), the same one EMS Follow-Up reads.'],
  ['CONSEQUENCE_SHEET', '1IVFnkQQvSvNSY_OZPMutPzkTqmJ1j5XYRoT9YvNy-v4', 'The Accountability & Intervention Log (link or ID), for Lunch Detentions.'],
  ['CONSEQUENCE_TAB', 'Master Log', 'The tab of that sheet with one row per consequence.'],
  ['LD_TYPES', 'Lunch Detention', 'Consequence Types that count as a Lunch Detention (words, comma-separated).'],
  ['HOMEROOM_AUTO_APPROVE', 'OFF', 'Homeroom form responses (Homeroom Points tab): OFF = staff tick Approved before they count; ON = they count right away (flagged ones still wait).'],
  ['EMBED_APPS', 'ON', 'ON = those three apps show inside their Clash pages (signed in with a school account). OFF = an Open button only.'],
  ['ROSTER_SHEET', 'https://docs.google.com/spreadsheets/d/1SDmwtkP9RzMZ62WVn6awZW6-6gA_wPtc6altTwoAeME/edit', 'The district roster Sheet (link or ID) with 800 Number, Email Address, First Name, Last Name, Grade Level. Clash of Classes › Import student roster copies it to the Students tab.'],
  ['STUDENT_EMAIL_PATTERN', '{800}@gomounts.org', 'My Climb sign-in when the Student Board has no Email column: how a student email is built from the 800#, e.g. {800}@easdpa.org. Blank = use the Email column only.'],
  ['MY_CLIMB_HOURS', '07:00-16:00', 'My Climb is open to students only between these times (24-hour clock). Keeps it school-focused.'],
  ['MY_CLIMB_SCHOOL_DAYS', 'ON', 'ON = My Climb is open only on school days (weekdays not on the School Calendar tab). Staff previews are never locked.'],
  ['MASCOT_ART', 'ON', 'ON = squad pages, the TV and Shady use the rendered mascot art (repo assets folder). OFF = the drawn versions. A link in Squad Themes › Mascot Image always wins. Students always customize the drawn version.'],
  ['SHIELD_LOGO_URL', 'https://raw.githubusercontent.com/matthewschuck-gif/Clash-of-Classes/ready-to-go-live/assets/shield-2026-27.png', 'The 2026-27 shield logo shown on My Climb (any https picture link).'],
  ['STUDENT_LINK', '', 'The student link (the second deployment, Anyone within EASD). The site\'s My Climb button opens it with ?page=me.'],
  ['SCHOOL_START', '', 'First day of school (yyyy-mm-dd) for the home page day counts. Blank = the date set in the site\'s admin settings.'],
  ['LAST_DAY', '', 'Last day of school (yyyy-mm-dd). Blank = the date set in the site\'s admin settings.'],
  ['SCHOOL_CALENDAR', '', 'Optional: the district calendar, as its Google Calendar ID or its iCal (.ics) link. Clash of Classes › Sync days off from the district calendar adds its days off to the School Calendar tab.'],
  ['SCHOOL_CALENDAR_WORDS', 'no school, holiday, break, closed, closing, conference, records day, snow, vacation day, inservice, in-service, act 80, labor day, memorial day, thanksgiving', 'An all-day district calendar event counts as a day off when its title has one of these (half days and early dismissals never do).'],];

// Writes one Settings value (used by the menu switches).
function setSettingValue_(key, value) {
  ensureSettings_();
  SETTINGS_MEMO_ = null;
  const sheet = getSheet_(SETTINGS_TAB);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === key) { sheet.getRange(i + 1, 2).setValues([[value]]); return; }
  }
  sheet.appendRow([key, value, '']);
}

function ensureSettings_() {
  SETTINGS_MEMO_ = null;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SETTINGS_TAB);
  if (!sheet) {
    sheet = ss.insertSheet(SETTINGS_TAB);
    sheet.getRange(1, 1, 1, 3).setValues([['Key', 'Value', 'Notes']]).setFontWeight('bold').setBackground('#490e6f').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }
  const have = {};
  sheet.getDataRange().getValues().slice(1).forEach(function (r, i) {
    let key = String(r[0]).trim();
    // Renamed keys keep their value: MONSTER_VISIBLE is OPPONENT_VISIBLE now.
    if (key === 'MONSTER_VISIBLE') { key = 'OPPONENT_VISIBLE'; sheet.getRange(i + 2, 1).setValue(key); }
    have[key] = true;
  });
  const missing = DEFAULT_SETTINGS.filter(function (d) { return !have[d[0]]; });
  if (missing.length) sheet.getRange(sheet.getLastRow() + 1, 1, missing.length, 3).setValues(missing);
}

// The Settings tab is read ONCE per request and kept here: the public page asks for settings
// hundreds of times (twice per points row), and each sheet read is a slow round trip.
// Writes through setSettingValue_ / ensureSettings_ clear it.
let SETTINGS_MEMO_ = null;
function settingsMap_() {
  if (SETTINGS_MEMO_) return SETTINGS_MEMO_;
  const map = {};
  const sheet = getSheetSafe_(SETTINGS_TAB);
  if (sheet) sheet.getDataRange().getValues().slice(1).forEach(function (r) {
    const k = String(r[0]).trim();
    if (k && !(k in map)) map[k] = String(r[1] == null ? '' : r[1]).trim();
  });
  SETTINGS_MEMO_ = map;
  return map;
}

function getSetting_(key) {
  const map = settingsMap_();
  if (key in map) return key.slice(-3) === '_ID' ? extractSheetId_(map[key]) : map[key];
  const d = DEFAULT_SETTINGS.filter(function (x) { return x[0] === key; })[0];
  return d ? d[1] : '';
}

// A pasted Sheet link → just its ID.
function extractSheetId_(v) {
  const m = String(v || '').match(/\/d\/([a-zA-Z0-9_-]{20,})/);
  return m ? m[1] : String(v || '').trim();
}

// Pre-seeds the metrics tab with one row per squad (BLACK/GOLD/GREY/PURPLE) so staff just
// overwrite numbers instead of having to know which squads to add. Only adds rows for squads
// that aren't already present, so re-running setupSpreadsheet never overwrites percentages
// someone already entered.
function ensureMetricsRows_() {
  const sheet = getSheet_(TABS.METRICS);
  const squads = ['BLACK', 'GOLD', 'GREY', 'PURPLE'];
  const data = sheet.getDataRange().getValues();
  const existing = {};
  for (let i = 1; i < data.length; i++) {
    const sq = String(data[i][0] || '').trim().toUpperCase();
    if (sq) existing[sq] = true;
  }
  const missing = squads.filter(function (sq) { return !existing[sq]; });
  if (missing.length) {
    const rows = missing.map(function (sq) { return [sq, 0, 0, 0]; });
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 4).setValues(rows);
  }
}

// Human-readable directions for whoever is actually entering events/points day to day --
// deliberately plain-language, not aimed at whoever set up the Apps Script backend. Safe to
// re-run (setupSpreadsheet is idempotent): this rewrites the instructions text but never
// touches app_data or events rows.
function setupHelpTab_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(TABS.HELP);
  if (!sheet) sheet = ss.insertSheet(TABS.HELP);
  else sheet.clear();

  const rows = [
    ['Clash of Classes -- How To Use This Sheet'],
    [''],
    ['This spreadsheet is the live backend for clashofclasses.org. Two tabs matter for'],
    ['day-to-day use: "events" is where you add points for activities, and "metrics" is where'],
    ['you enter each month\'s attendance / referral / grades percentages. Changes here show up'],
    ['on the website automatically within about 5-10 seconds -- no need to tell anyone or'],
    ['refresh anything on your end.'],
    [''],
    ['The school map game (zones, ways to conquer, what owning a zone pays) has its own guide:'],
    ['the "How The Map Works" tab. Creatures, LRG points, boosts and perks: the "How My Climb Works" tab.'],
    [''],
    ['HOW TO ADD AN EVENT / GIVE POINTS'],
    ['1. Open the "events" tab (the tab bar is at the very bottom of this spreadsheet).'],
    ['2. Add a NEW ROW at the bottom of the data (do not insert rows in the middle).'],
    ['3. Fill in these columns for that row:'],
    ['     id       -- leave BLANK. The sheet fills this in automatically.'],
    ['     name     -- what the event/activity was called, e.g. "Spirit Week Kickoff"'],
    ['     type     -- a short category label, e.g. "spirit", "academic", "attendance"'],
    ['     date     -- the date of the event, format YYYY-MM-DD, e.g. 2026-09-15'],
    ['     note     -- optional, any extra context'],
    ['     multiplier -- optional, leave blank unless you specifically want to double/triple'],
    ['                   this event\'s points (e.g. enter 2 for double points)'],
    ['     BLACK / GOLD / GREY / PURPLE -- the number of points EACH squad earned for this'],
    ['                   specific event. Enter 0 for a squad that did not earn points here.'],
    ['4. That\'s it -- do not touch any other tab. The website will update on its own.'],
    [''],
    ['HOW TO FIX A MISTAKE'],
    ['- Wrong number? Just edit the number in that cell directly -- the site will pick up the'],
    ['  correction automatically, same as any other edit.'],
    ['- Need to remove an event entirely? Delete that whole row (right-click the row number ->'],
    ['  Delete row). Don\'t just clear the cells and leave a blank row behind.'],
    [''],
    ['HOW TO ENTER MONTHLY METRICS (ATTENDANCE / REFERRAL / GRADES)'],
    ['1. Open the "metrics" tab. There\'s already one row per squad (Black, Gold, Grey,'],
    ['   Purple) -- don\'t add or delete rows here, just edit the numbers.'],
    ['2. Enter each squad\'s percentage for that month in the attendance / referral / grades'],
    ['   columns (whole numbers or decimals both work, e.g. 94 or 94.5).'],
    ['3. That\'s it for entering the numbers -- the website\'s "Monthly Metrics" preview'],
    ['   updates automatically, same as the events tab.'],
    ['4. IMPORTANT: entering numbers here does NOT award points by itself. Points only lock'],
    ['   in when an admin goes to clashofclasses.org -> Admin -> Metrics tab and clicks'],
    ['   "Finalize This Month & Award Points." That is on purpose -- it means you can update'],
    ['   these numbers throughout the month as new reports come in without accidentally'],
    ['   awarding points early. The highest squad in each of the 3 metrics earns 250 points,'],
    ['   2nd place 125, 3rd place 50, and last place 0 (a tie splits the higher spot\'s points).'],
    ['5. The month LABEL (e.g. "September 2026") is entered on the website, not in this sheet'],
    ['   -- Admin -> Metrics tab -> Month Label.'],
    [''],
    ['HOW TO RESET/START A NEW SCHOOL YEAR -- IMPORTANT, READ THIS FIRST'],
    ['Do NOT reset the season by deleting all the rows in the "events" tab yourself. Because of'],
    ['how this sheet talks to the website, clearing the tab by hand will NOT actually reset the'],
    ['scores -- the website will just quietly restore the old numbers back into this tab within'],
    ['moments, and it will look like nothing happened (or like your edit got undone).'],
    [''],
    ['Instead: go to clashofclasses.org, log into Admin, and use the "Reset All Scores" button'],
    ['there. That is the only way that safely resets everything (it clears this tab AND the'],
    ['website\'s stored data together, in sync). If you\'re not sure how to get into Admin, ask'],
    ['whoever manages the site.'],
    [''],
    ['DO NOT TOUCH'],
    ['The "app_data" tab is not meant to be edited by hand -- it stores the rest of the site\'s'],
    ['settings (theme, gallery, admin password, etc.) as encoded data, not as plain readable'],
    ['numbers or text. Editing it directly can break the site. If something there needs to'],
    ['change, do it through the website\'s own Admin panel instead.'],
  ];

  sheet.getRange(1, 1, rows.length, 1).setValues(rows);
  sheet.getRange(1, 1).setFontWeight('bold').setFontSize(14).setFontColor('#490e6f');

  // Bold section headers by matching text rather than hardcoded row numbers, so this
  // doesn't silently drift out of sync if the copy above ever gets edited.
  const boldSections = ['HOW TO ADD AN EVENT / GIVE POINTS', 'HOW TO FIX A MISTAKE', 'HOW TO ENTER MONTHLY METRICS (ATTENDANCE / REFERRAL / GRADES)', 'DO NOT TOUCH'];
  const warnSections = ['HOW TO RESET/START A NEW SCHOOL YEAR -- IMPORTANT, READ THIS FIRST'];
  rows.forEach(function (r, i) {
    const text = r[0];
    if (boldSections.indexOf(text) !== -1) sheet.getRange(i + 1, 1).setFontWeight('bold');
    if (warnSections.indexOf(text) !== -1) sheet.getRange(i + 1, 1).setFontWeight('bold').setFontColor('#b91c1c');
  });

  sheet.setColumnWidth(1, 760);
  sheet.setFrozenRows(1);
  ss.setActiveSheet(sheet);
  ss.moveActiveSheet(1);
}

function getSheet_(tabName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(tabName);
  if (!sheet) throw new Error('Sheet tab not found: ' + tabName + ' -- run setupSpreadsheet() first.');
  return sheet;
}

// Same as getSheet_ but returns null instead of throwing when the tab doesn't exist yet.
// Used for tabs added after the initial setup (like METRICS) so that a deployment redeployed
// with newer router code doesn't immediately break every request just because the user
// hasn't re-run setupSpreadsheet() on their existing sheet yet -- it just behaves as if that
// feature's data isn't there yet, same as a brand new site.
function getSheetSafe_(tabName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(tabName) || null;
}

function getProp_(key) {
  const v = PropertiesService.getScriptProperties().getProperty(key);
  if (!v) throw new Error('Missing Script Property: ' + key + ' -- set it in Project Settings -> Script Properties.');
  return v;
}

function nowIso_() {
  return new Date().toISOString();
}
