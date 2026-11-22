/**
 * Clash of Classes -- Apps Script backend
 * File 49: the Sheet's tabs, grouped, color coded and in priority order (Menu › Group and color the
 * tabs; Run setup does it too). Tabs a group doesn't list stay where they fall, just before the
 * system tabs at the end. Change TAB_GROUPS_ to change the order or the colors.
 *   1 Start here        red      the guide and Settings
 *   2 Every week        green    the tabs you work in most
 *   3 Points & map      purple   points, the map, freezes, Avalanche, Virtual Dough
 *   4 Students & families blue   rosters and family pieces
 *   5 Game & pages      orange   prizes, the Wheel, Plinko, traits, pages, the tour
 *   6 Written for you   gray     reports and logs the site writes (look, don't edit)
 *   7 System            dark     the site's saved data and raw votes (leave alone)
 */

const TAB_GROUPS_ = [
  ['Start here', '#cc0000', ['How To Use This Sheet', 'Settings']],
  ['Every week', '#34a853', ['Daily Polls', 'Poll Ideas', 'You Said We Did', 'Service Submissions', 'Family Shout-outs', 'Squad Spotlight', 'Weather Days', 'Expeditions', 'Squad Colors Days', 'Family Pick', 'Send-off Advice', 'Plinko Drops', 'Staff Shout-outs', 'Squad Banners', 'Planned Events', 'Squad Goals',
    'Frost Log', 'Sign-up Events', 'Year Plan', 'Family Challenges', 'Family Pulse']],
  ['Points & map', '#8e44ad', ['events', 'Squad Monthly', 'Zone Overrides', 'Map Zones', 'Conquest Rules', 'Map Battles', 'Point Freezes', 'Freeze Lines',
    'Avalanche', 'Virtual Dough', 'Dough Teachers', 'Challenges', 'Squad Duels', 'Helping Hands', 'Homeroom Points', 'Homeroom Forms']],
  ['Students & families', '#4a86e8', ['Students', 'Student Board', 'Squad Board', 'MLC Roster', 'Event Check-ins', 'Sign-ups', 'Nominations', 'General Letters',
    'Family Pulse Answers', 'Family Check-ins', 'Family Pick Votes', 'Team League', 'Team Teachers']],
  ['Game & pages', '#ff9900', ['Plinko Slots', 'Hall of Fame', 'Wheel Slices', 'LRG Traits', 'Opponents', 'Round Goals', 'Engine Powers', 'Chance Cards', 'Evolution Boosts', 'Perks',
    'Destination Tickets', 'Lore Chapters', 'Squad Themes', 'Event Pages', 'Event Rows', 'Mindset Links', 'Tutorial Steps', 'School Calendar', 'Mayhem Checklist', 'Shirt Designs']],
  ['Written for you', '#999999', ['Poll Insights', 'Points Balance', 'How The Map Works', 'How My Climb Works', 'Admin Edits', 'Weekly Digest Log', 'Usage',
    'Redemptions', 'Make It Right', 'Boost Ledger', 'Student Tickets', 'Squad Hand', 'Streak Banks', 'Arcade Plays', 'Pebble Cash-ins', 'Pebble Bonuses', 'Pebble Spends', 'Shirt Votes', 'Shirt Orders', 'Squad Night']],
  ['System', '#434343', ['Poll Votes', 'Avatars', 'Color Squads', 'metrics', 'app_data']],
];

/** Colors and reorders the tabs. Returns how many tabs it placed. */
function organizeTabs_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.moveActiveSheet) return 0;
  const listed = {}, order = [];
  TAB_GROUPS_.forEach(function (g) { g[2].forEach(function (name) { listed[name] = g; }); });
  const systemStart = TAB_GROUPS_.length - 2; // "Written for you" and "System" go after anything unlisted
  TAB_GROUPS_.slice(0, systemStart).forEach(function (g) { g[2].forEach(function (n) { const sh = ss.getSheetByName(n); if (sh) order.push([sh, g[1]]); }); });
  ss.getSheets().forEach(function (sh) { if (!listed[sh.getName()]) order.push([sh, null]); });
  TAB_GROUPS_.slice(systemStart).forEach(function (g) { g[2].forEach(function (n) { const sh = ss.getSheetByName(n); if (sh) order.push([sh, g[1]]); }); });
  const keep = ss.getActiveSheet();
  order.forEach(function (o, i) {
    try { if (o[1]) o[0].setTabColor(o[1]); ss.setActiveSheet(o[0]); ss.moveActiveSheet(i + 1); } catch (e) { console.warn('Tab order: ' + o[0].getName() + ': ' + e.message); }
  });
  try { ss.setActiveSheet(ss.getSheetByName('Settings') || keep); } catch (e) { /* fine */ }
  return order.length;
}

function menuOrganizeTabs_() {
  const n = organizeTabs_();
  SpreadsheetApp.getUi().alert('Tabs grouped', n + ' tabs, in this order:\n\n' + TAB_GROUPS_.map(function (g, i) { return (i + 1) + '. ' + g[0]; }).join('\n') +
    '\n\nThe colors match the groups. Tabs not in a group sit just before "Written for you".', SpreadsheetApp.getUi().ButtonSet.OK);
}

// file-stamp (tools/stamp-files.js): lets Admin › System tell whether this file is the newest one
var FILE_STAMP_49_TabOrder = '45791acfae';
