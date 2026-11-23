/**
 * File 66: Camp Mountaineer tie-ins.
 *
 * COMEBACK GOALS (COMEBACK_GOALS ON): Camp writes the Comeback Goals tab (800#, Goal, Set, Check By,
 * Status, Done, Key) when a student goes back to class. The goal shows only to that student on My Climb
 * (never on a board, the TV or a squad page). Done goals count as "comeback" for perks, so the Perks tab
 * rule "comeback: 1" unlocks Comeback Climber.
 *
 * STUDENT CREATURES (STUDENT_CREATURES ON): once a day the Student Creatures tab (800#, Squad, Stage,
 * Art URL, Pebbles) is rewritten so Camp's kiosk can greet a student with their own creature. Only the
 * Camp script reads it (the tab is in this private Sheet).
 *
 * Camp milestone Pebbles arrive on Pebble Bonuses as From "Camp milestone" (48_Arcade.gs counts them
 * in the student's pouch; Pebble watch shows them on their own line).
 */
const COMEBACK_HEADERS_ = ['800#', 'Goal', 'Set', 'Check By', 'Status', 'Done', 'Key'];
const CREATURE_HEADERS_ = ['800#', 'Squad', 'Stage', 'Art URL', 'Pebbles', 'Updated'];

function comebackOn_() { return String(getSetting_('COMEBACK_GOALS') || 'ON').toUpperCase() !== 'OFF'; }

// { open: [{ goal, set, checkBy }], done: n } for one student; null when off or no tab.
function comebackFor_(id) {
  if (!comebackOn_() || !getSheetSafe_('Comeback Goals')) return null;
  const out = { open: [], done: 0 };
  sheetRows_('Comeback Goals').forEach(function (r) {
    if (id800_(r['800#']) !== id) return;
    if (/^done/i.test(String(r['Status']))) out.done++;
    else if (/^open/i.test(String(r['Status']))) out.open.push({ goal: String(r['Goal'] || ''), set: String(r['Set'] || '').slice(0, 10), checkBy: String(r['Check By'] || '').slice(0, 10) });
  });
  return out.open.length || out.done ? out : null;
}
function comebackDone_(id) { try { const c = comebackFor_(id); return c ? c.done : 0; } catch (e) { return 0; } }

// Daily: every student's creature for Camp's kiosk welcome.
function maybeStudentCreatures_(force) {
  if (String(getSetting_('STUDENT_CREATURES') || 'ON').toUpperCase() === 'OFF') return 0;
  const props = PropertiesService.getScriptProperties(), today = todayYmd_();
  if (!force && props.getProperty('student-creatures') === today) return 0;
  props.setProperty('student-creatures', today);
  const names = stageNames_(), versions = {}, pouch = {};
  try { sheetRows_('Avatars').forEach(function (a) { let l = {}; try { l = JSON.parse(String(a['Look'] || '{}')); } catch (e) { l = {}; } versions[id800_(a['800#'])] = versionOf_(l); }); } catch (e) { /* no Avatars tab */ }
  if (pebblesOn_()) {
    const add = function (tab, col, sign) { try { sheetRows_(tab).forEach(function (r) { const id = id800_(r['800#']); pouch[id] = (pouch[id] || 0) + sign * (Number(r[col]) || 0); }); } catch (e) { /* tab not made yet */ } };
    add('Arcade Plays', 'Pebbles', 1); add('Pebble Bonuses', 'Pebbles', 1); add('Pebble Cash-ins', 'Pebbles', -1); add('Pebble Spends', 'Pebbles', -1);
  }
  const rows = [];
  boardRows_().forEach(function (r) {
    const id = id800_(r['800#']), s = squadOf_(r['Squad']);
    if (!id || SQUAD_KEYS_.indexOf(s) === -1) return;
    const st = stageFor_(climbXp_({ bills: boardNumber_(r['Trait Dough Bills']), mindset: boardNumber_(r['Mindset Summits']), star: boardNumber_(r['STAR Growth Points']),
      shoutouts: boardNumber_(r['Shout-outs']), buddy: boardNumber_(r['Ambassador Buddies']), streak: boardNumber_(r['Streak (weeks)']), journal: boardNumber_(r['Trail Journal Growth']) }));
    rows.push([id, s, names[st - 1], creatureArt_(s, st, versions[id] || 1), Math.max(0, pouch[id] || 0), today]);
  });
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('Student Creatures');
  if (!sh) { sh = ss.insertSheet('Student Creatures'); sh.getRange(1, 1, 1, CREATURE_HEADERS_.length).setValues([CREATURE_HEADERS_]).setFontWeight('bold'); sh.setFrozenRows(1); }
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, CREATURE_HEADERS_.length).clearContent();
  if (rows.length) sh.getRange(2, 1, rows.length, CREATURE_HEADERS_.length).setValues(rows);
  return rows.length;
}

// file-stamp (tools/stamp-files.js): lets Admin › System tell whether this file is the newest one
var FILE_STAMP_66_Comeback = 'bebb0a48eb';
