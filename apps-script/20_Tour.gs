/**
 * Clash of Classes -- Apps Script backend
 * File 20: Cliff's tour. A pop-up walk-through where Cliff, the Mountaineer, spotlights each part
 * of the game on the page and explains it (CliffTour.html).
 *
 * TUTORIAL STEPS tab: Show | Page | Order | Go To | Spotlight | Title | Text
 *   Page      -- site (clashofclasses.org) or me (My Climb)
 *   Go To     -- the site page (home, squads, mysquad, tracker…) or My Climb tab (me, tickets,
 *                squad, map, game) to open first. Blank = stay where the tour is.
 *   Spotlight -- what to light up: a plain name (see SPOTLIGHTS in CliffTour.html: map, map rules,
 *                opponent, poll, score, my climb, creature, lrg bar, stages, versions, boosts,
 *                perks, gear, tickets, metrics, round goals, engine, season, spotlight, checkin, pulse,
 *                my squad, frost, avalanche, service…) or a CSS selector.
 *                Blank, or not on the page right now = Cliff talks in the middle of the screen.
 * The tour plays once on each device (TOUR_AUTO ON); "Ask Cliff" replays it any time. Change
 * TOUR_VERSION after editing the steps to show everyone the new tour once.
 */

const TOUR_HEADERS_ = ['Show', 'Page', 'Order', 'Go To', 'Spotlight', 'Title', 'Text'];

// The steps Run setup puts on the tab. Orders with a decimal slot newer steps between the first ones,
// so a tab staff already edited keeps its own order.
const TOUR_SEED_ = [
  [true, 'site', 1, 'home', '', 'Hi, I\'m Cliff!', 'I\'m the Mountaineer. Let me show you how your squad climbs to the top of Clash of Classes. It takes a minute.'],
  [true, 'site', 2, 'home', 'score', 'Every point counts', 'This is the race. Points from Trait Dough, events, polls and everything else show up here within seconds.'],
  [true, 'site', 2.5, 'home', 'season', 'The race to the Clash Bash', 'Points count until the last school day of May. In June we find out which squad wins the Clash Bash. Every point between now and then matters.'],
  [true, 'site', 3, 'home', 'my climb', 'Your own climb', 'Open My Climb with your school Google account to meet your creature, earn LRG points and evolve it.'],
  [true, 'site', 4, 'home', 'opponent', 'The Opponent', 'Sometimes the whole school teams up against an Opponent. Beat the goal together and every squad gets points.'],
  [true, 'site', 5, 'home', 'map', 'Conquer the school', 'Every area of the building is a territory, and each one is won a different way: Pride Day, attendance, passing grades, shout-outs, polls and more. Win last month and your squad holds it all month.'],
  [true, 'site', 6, 'home', 'map rules', 'How each zone is won', 'Tap here to see what wins each zone. Zones you win by checking in (Pride Day, squad events, polls) pay double every week.'],
  [true, 'site', 7, 'home', 'poll', 'The Daily Poll', 'Vote every morning at clashofclasses.org/vote on your Chromebook with your 800 number. Every vote is a point for your squad.'],
  [true, 'site', 7.2, 'home', 'checkin', 'Check in', 'On Pride Day, Wheel days and squad events (Move, Make, Talk), check in here with your 800 number. A General confirms squad events. Check-ins win zones on the map.'],
  [true, 'site', 7.4, 'pulse', 'pulse', 'The Pulse', 'Some questions are the school asking what you really think. There is no right answer, it stays open a few days, and you see the results the moment you vote. Watch for "You said, we did": what changed because you answered.'],
  [true, 'site', 7.6, 'home', 'spotlight', 'Squad Spotlight', 'Every week staff pick a story about a student who showed a trait or did service. It could be you.'],
  [true, 'site', 8, 'squads', 'squads', 'Four squads', 'Gold, Purple, Grey and Black. Every student belongs to one all year, and every point you earn helps them.'],
  [true, 'site', 8.2, 'mysquad', 'my squad', 'My Squad', 'Type your 800 number to see what you have added: Trait Dough, Mindset summits, badges, your squad\'s place, and the events you can join.'],
  [true, 'site', 8.4, 'mysquad', 'frost', 'Freeze and thaw', 'Points never go away, but a squad can freeze if it crosses a line (like attendance or behavior). The meter shows how close your squad is. A frozen squad thaws together, never one student alone.'],
  [true, 'site', 8.6, 'mysquad', 'avalanche', 'Avalanche!', 'If your squad is in last place, every vote and check-in charges the Avalanche meter. When it is full, a General sets it off: a week of double points. Once per marking period.'],
  [true, 'site', 8.8, 'mysquad', 'service', 'Service and kindness', 'Did service, volunteered or helped someone? Tell us what you did and what it meant, with a link or a photo if you have one. Staff approve it for squad points.'],
  [true, 'site', 8.9, 'home', '', 'Teachers give Trait Dough too', 'Teachers can hand out Virtual Trait Dough when they catch you showing a trait, and they can only give it to students in the other squads. Show your traits everywhere!'],
  [true, 'site', 9, 'home', '', 'That\'s Clash!', 'Tap Ask Cliff any time to see this again. Now go earn some points for your squad!'],
  [true, 'me', 1, 'me', '', 'Hi, I\'m Cliff!', 'Welcome to your climb. Let me show you around. It takes a minute.'],
  [true, 'me', 2, 'me', 'creature', 'This is your creature', 'It\'s your squad\'s animal, and it\'s all yours. It grows up as you earn LRG points.'],
  [true, 'me', 3, 'me', 'lrg bar', 'LRG points', 'LRG means Life Ready Graduate. Trait Dough, Mindset summits, shout-outs, STAR growth, streaks and tickets all earn them. Fill this bar to evolve.'],
  [true, 'me', 4, 'me', 'stages', 'Four forms', 'Foothills, Ridge, High Camp and Summit. Tap any form you\'ve reached to show it off.'],
  [true, 'me', 5, 'me', 'versions', 'Pick your version', 'Every form comes in three looks. Version 3 is the winter coat.'],
  [true, 'me', 6, 'me', 'boosts', 'Evolution boosts', 'Each time you evolve, you pick a boost that earns extra points for your squad every week: ×1.25, then ×1.5, then ×2.'],
  [true, 'me', 7, 'me', 'perks', 'Perks', 'Titles, frames and scenes you unlock as you climb. Tap one to wear it.'],
  [true, 'me', 8, 'me', 'gear', 'Gear', 'Earn gear in your squad\'s color and wear up to three pieces on your creature.'],
  [true, 'me', 9, 'tickets', 'tickets', 'Destination tickets', 'Draw three personal goals and keep at least one. Finish a ticket and your squad earns its points.'],
  [true, 'me', 10, 'map', 'map', 'Conquer the school', 'Squads take areas of the building, each won its own way (Pride Day, polls, attendance, grades and more), judged on last month. Tap "How to win each zone" to see every rule.'],
  [true, 'me', 11, 'game', '', 'Everything else', 'The Opponent, round goals, Hot Streaks and your squad\'s engine powers live in the Game tab.'],
  [true, 'me', 11.5, 'me', 'metrics', 'Your numbers', 'Your attendance and STAR next to your squad\'s average. Only you can see yours.'],
  [true, 'me', 12, 'me', '', 'That\'s your climb!', 'Tap Ask Cliff any time to see this again. Now go earn some LRG points!'],
];
// The first wording of steps that changed: a row still saying exactly this gets the new wording.
const TOUR_OLD_TEXT_ = {
  'site|Conquer the school': 'Every area of the building is a territory. Earn the right kind of points to take it, and hold it to earn bonus points every week.',
  'site|How each zone is won': 'Tap here to see what wins each zone. ⚔ Siege zones: a squad next door attacks with a small edge; from farther away you need a big lead.',
  'site|The Daily Poll': 'Vote every morning. Your squad gets points for turning out, and a Hot Streak for topping the vote days in a row.',
  'me|Conquer the school': 'Squads take areas of the building by earning the right kind of points. Tap "How to win each zone" to see every rule.',
};

function ensureTourTab_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('Tutorial Steps');
  if (!sh) {
    sh = ss.insertSheet('Tutorial Steps');
    sh.getRange(1, 1, 1, TOUR_HEADERS_.length).setValues([TOUR_HEADERS_]).setFontWeight('bold').setBackground('#490e6f').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    sh.getRange(2, 1, TOUR_SEED_.length, TOUR_HEADERS_.length).setValues(TOUR_SEED_);
    return 0;
  }
  return mergeTourSteps_(sh);
}

// Adds the newer steps an older tab doesn't have (matched by page + title) and refreshes steps that
// still have their first wording. Anything staff wrote is kept. New steps = TOUR_VERSION goes up by
// one, so everyone sees the new tour once. Returns how many steps changed.
function mergeTourSteps_(sh) {
  const rows = sheetRows_('Tutorial Steps'), have = {};
  rows.forEach(function (r) { have[String(r['Page'] || '').trim().toLowerCase() + '|' + String(r['Title'] || '').trim()] = r; });
  const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(function (h) { return String(h).trim(); });
  const textCol = head.indexOf('Text') + 1;
  let changed = 0;
  TOUR_SEED_.forEach(function (s) {
    const k = s[1] + '|' + s[5], r = have[k];
    if (!r) {
      const o = {}; TOUR_HEADERS_.forEach(function (h, i) { o[h] = s[i]; });
      sh.appendRow(head.map(function (h) { return o[h] === undefined ? '' : o[h]; }));
      changed++;
    } else if (textCol && TOUR_OLD_TEXT_[k] && String(r['Text'] || '').trim() === TOUR_OLD_TEXT_[k] && TOUR_OLD_TEXT_[k] !== s[6]) {
      sh.getRange(r._row, textCol).setValue(s[6]);
      changed++;
    }
  });
  if (changed) setSettingValue_('TOUR_VERSION', String((Number(getSetting_('TOUR_VERSION')) || 1) + 1));
  return changed;
}

// The steps for one page, in order, plus how the pop-up should behave.
function tourFor_(page) {
  page = String(page || '').trim().toLowerCase();
  const steps = sheetRows_('Tutorial Steps').filter(function (r) {
    return ticked_(r['Show']) && String(r['Page'] || '').trim().toLowerCase() === page && String(r['Title'] || r['Text'] || '').trim();
  }).sort(function (a, b) { return (Number(a['Order']) || 0) - (Number(b['Order']) || 0); }).map(function (r) {
    return { go: String(r['Go To'] || '').trim().toLowerCase(), spot: String(r['Spotlight'] || '').trim(), title: String(r['Title'] || '').trim(), text: String(r['Text'] || '') };
  });
  return { page: page, steps: steps, version: String(getSetting_('TOUR_VERSION') || '1'), auto: String(getSetting_('TOUR_AUTO') || 'ON').toUpperCase() !== 'OFF', cliff: cliffInfo_() };
}

function clashTour(page) { return JSON.stringify(tourFor_(page === 'me' ? 'me' : 'site')); }

// file-stamp (tools/stamp-files.js): lets Admin › System tell whether this file is the newest one
var FILE_STAMP_20_Tour = 'e2400bff5a';
