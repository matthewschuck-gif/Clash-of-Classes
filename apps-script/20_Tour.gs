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
 *                perks, gear, tickets, metrics, round goals, engine…) or a CSS selector.
 *                Blank, or not on the page right now = Cliff talks in the middle of the screen.
 * The tour plays once on each device (TOUR_AUTO ON); "Ask Cliff" replays it any time. Change
 * TOUR_VERSION after editing the steps to show everyone the new tour once.
 */

const TOUR_HEADERS_ = ['Show', 'Page', 'Order', 'Go To', 'Spotlight', 'Title', 'Text'];

function ensureTourTab_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getSheetByName('Tutorial Steps')) return;
  const rows = [
    [true, 'site', 1, 'home', '', 'Hi, I\'m Cliff!', 'I\'m the Mountaineer. Let me show you how your squad climbs to the top of Clash of Classes. It takes a minute.'],
    [true, 'site', 2, 'home', 'score', 'Every point counts', 'This is the race. Points from Trait Dough, events, polls and everything else show up here within seconds.'],
    [true, 'site', 3, 'home', 'my climb', 'Your own climb', 'Open My Climb with your school Google account to meet your creature, earn LRG points and evolve it.'],
    [true, 'site', 4, 'home', 'opponent', 'The Opponent', 'Sometimes the whole school teams up against an Opponent. Beat the goal together and every squad gets points.'],
    [true, 'site', 5, 'home', 'map', 'Conquer the school', 'Every area of the building is a territory. Earn the right kind of points to take it, and hold it to earn bonus points every week.'],
    [true, 'site', 6, 'home', 'map rules', 'How each zone is won', 'Tap here to see what wins each zone. ⚔ Siege zones: a squad next door attacks with a small edge; from farther away you need a big lead.'],
    [true, 'site', 7, 'home', 'poll', 'The Daily Poll', 'Vote every morning. Your squad gets points for turning out, and a Hot Streak for topping the vote days in a row.'],
    [true, 'site', 8, 'squads', 'squads', 'Four squads', 'Gold, Purple, Grey and Black. Every student belongs to one all year, and every point you earn helps them.'],
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
    [true, 'me', 10, 'map', 'map', 'Conquer the school', 'Squads take areas of the building by earning the right kind of points. Tap "How to win each zone" to see every rule.'],
    [true, 'me', 11, 'game', '', 'Everything else', 'The Opponent, round goals, Hot Streaks and your squad\'s engine powers live in the Game tab.'],
    [true, 'me', 12, 'me', '', 'That\'s your climb!', 'Tap Ask Cliff any time to see this again. Now go earn some LRG points!'],
  ];
  const sh = ss.insertSheet('Tutorial Steps');
  sh.getRange(1, 1, 1, TOUR_HEADERS_.length).setValues([TOUR_HEADERS_]).setFontWeight('bold').setBackground('#490e6f').setFontColor('#ffffff');
  sh.setFrozenRows(1);
  sh.getRange(2, 1, rows.length, TOUR_HEADERS_.length).setValues(rows);
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
