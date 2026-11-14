/**
 * Clash of Classes -- Apps Script backend
 * File 45: the setup self-check (Admin › 🩺 System, and the menu). Catches the setup slips that quietly
 * break something, and says how to fix each one:
 *   season end on a school day · marking periods cover today · Dough Teachers have a squad and a
 *   password, no two share one · Pulse ideas left · a poll in the next school days · the background
 *   timer running · Climb-er and Follow-Up numbers arrived for last month · email recipients set ·
 *   every student on the roster has a squad · the example Dough Teacher unticked.
 */

/** [{ ok, area, text, fix }] */
function selfCheck_() {
  const out = [], today = todayYmd_();
  const add = function (ok, area, text, fix) { out.push({ ok: !!ok, area: area, text: text, fix: ok ? '' : fix }); };
  let off = {};
  try { schoolCalendar_().noSchool.forEach(function (d) { off[d] = 1; }); } catch (e) { off = {}; }
  const schoolDay = function (ymd) { const dow = new Date(ymd + 'T12:00:00Z').getUTCDay(); return dow > 0 && dow < 6 && !off[ymd]; };
  // Season.
  const end = seasonEnd_();
  if (!end) add(false, 'Season', 'No SEASON_END: the Clash Bash standings never lock.', 'Settings › SEASON_END: the last school day of May (yyyy-mm-dd).');
  else add(schoolDay(end), 'Season', 'Points count through ' + end + (schoolDay(end) ? '.' : ', which is not a school day.'), 'Settings › SEASON_END: pick the last school day (not a weekend or a day off on the School Calendar tab).');
  const rev = String(getSetting_('BASH_REVEAL') || '').trim();
  if (rev) add(!end || rev > end, 'Season', 'The Clash Bash winner is revealed ' + rev + '.', 'Settings › BASH_REVEAL must be after SEASON_END.');
  // Marking periods (Avalanche and the beat-your-MP bonus need them).
  const mp = mpFor_(today).now, mps = markingPeriods_();
  add(mps.length && mp && today >= mp.start && today <= mp.end, 'Marking periods', mps.length ? (mp && today >= mp.start && today <= mp.end ? 'Today is in ' + mp.name + '.' : 'Today is not inside any marking period.') : 'No marking periods.', 'Settings › MARKING_PERIODS: MP1=yyyy-mm-dd..yyyy-mm-dd; MP2=... with this year\'s district dates.');
  // Virtual Trait Dough.
  const dough = sheetRows_('Dough Teachers');
  const active = dough.filter(function (r) { return ticked_(r['Active']) && String(r['Teacher'] || '').trim(); });
  if (!active.length) add(false, 'Virtual Trait Dough', 'No teacher can give Virtual Trait Dough yet.', 'Dough Teachers tab (or Admin › 📋 Sheet): tick Active, name, squad, password.');
  else {
    const bad = active.filter(function (r) { return SQUAD_KEYS_.indexOf(squadOf_(r['Squad'])) === -1 || !String(r['Password'] == null ? '' : r['Password']).trim(); });
    add(!bad.length, 'Virtual Trait Dough', active.length + ' teachers can give dough' + (bad.length ? '; ' + bad.length + ' missing a squad or a password (' + bad.map(function (r) { return r['Teacher']; }).join(', ') + ')' : '') + '.', 'Dough Teachers: give each one a squad (BLACK, GOLD, GREY, PURPLE) and a password.');
    const seen = {}, dup = [];
    active.forEach(function (r) { const p = String(r['Password']).trim().toLowerCase(); if (p && seen[p]) dup.push(r['Teacher'] + ' and ' + seen[p]); seen[p] = r['Teacher']; });
    add(!dup.length, 'Virtual Trait Dough', dup.length ? 'Two teachers share a password: ' + dup.join('; ') + '.' : 'Every password is different.', 'Dough Teachers: give each teacher their own password (e.g. their room number).');
    if (active.some(function (r) { return /example/i.test(String(r['Teacher'])); })) add(false, 'Virtual Trait Dough', 'The example teacher row is still active.', 'Dough Teachers: untick or replace "Example Teacher".');
  }
  // Polls.
  let ready = 0;
  try { ready = readyIdeas_(today).filter(function (r) { return pollKind_(r['Kind']) === 'Pulse'; }).length; } catch (e) { ready = 0; }
  add(ready >= 3, 'Polls', ready + ' Pulse questions ready on Poll Ideas.', 'Poll Ideas: add questions (Kind Pulse, Use ticked) or clear Used On on ones to repeat.');
  const soon = pollsWithVotes_().filter(function (p) { return p.show && p.through >= today && p.date <= ymdShift_(today, 7); }).length;
  add(soon > 0, 'Polls', soon ? soon + ' poll' + (soon === 1 ? '' : 's') + ' open now or in the next week.' : 'No poll open now or in the next week.', 'Menu › Schedule poll ideas, or Admin › 📋 Sheet › Daily Polls › Add the next Pulse question.');
  // The background timer.
  let tick = false;
  try { tick = ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'backgroundTick'; }); } catch (e) { tick = false; }
  const ran = PropertiesService.getScriptProperties().getProperty('ran-backgroundTick');
  const fresh = ran && Date.now() - Date.parse(ran) < 30 * 60000;
  add(tick && fresh, 'Background', tick ? (fresh ? 'The background timer ran in the last half hour.' : 'The background timer is set but hasn\'t run in 30 minutes.') : 'The background timer is not set, so the map, freezes, Avalanche and emails don\'t run.', 'Clash of Classes › Run setup (it sets the timer). If it still doesn\'t run, check Apps Script › Executions for errors.');
  // Last month's numbers.
  const lm = lastMonthWindow_(today).month, monthly = squadMonthly_(lm);
  const have = SQUAD_KEYS_.filter(function (s) { return monthly[s] && monthly[s].attendance !== undefined; }).length;
  add(have === 4, 'Squad Monthly', have === 4 ? 'Climb-er sent ' + monthLabel_(lm) + '\'s numbers for all four squads.' : 'Only ' + have + ' of 4 squads have ' + monthLabel_(lm) + ' attendance on Squad Monthly, so those areas of the map wait.', 'In Define a Climb-er: menu › Send last month\'s map numbers to Clash (and keep its squad sync on).');
  const lrr = SQUAD_KEYS_.filter(function (s) { return monthly[s] && monthly[s].lrr_ld !== undefined; }).length;
  add(lrr === 4, 'Squad Monthly', lrr === 4 ? 'LRR + LD counts are in for ' + monthLabel_(lm) + '.' : 'LRR + LD counts are missing for ' + monthLabel_(lm) + '.', 'Menu › Count LRRs + LDs for the Cafeteria now (Settings › LRR_SHEET and CONSEQUENCE_SHEET must be set).');
  // Email recipients.
  if (String(getSetting_('STATE_OF_SQUADS') || 'ON').toUpperCase() === 'ON') add(stateRecipients_().length > 0, 'Emails', stateRecipients_().length ? 'The State of the Squads goes to ' + stateRecipients_().length + '.' : 'No one gets the State of the Squads email.', 'Settings › STATE_EMAILS (or DIGEST_EMAILS): the admins\' emails.');
  // The roster.
  const students = sheetRows_('Students');
  const noSquad = students.filter(function (r) { return SQUAD_KEYS_.indexOf(squadOf_(r['Squad'])) === -1; }).length;
  add(students.length && !noSquad, 'Roster', students.length ? students.length + ' students' + (noSquad ? ', ' + noSquad + ' without a squad' : ', every one in a squad') + '.' : 'The Students tab is empty.', 'Clash of Classes › Run setup (the roster sync), or give the missing students a squad on the district roster.');
  return out;
}

function clashSelfCheck(token) { requireAdmin_(token); return JSON.stringify(selfCheck_()); }

function menuSelfCheck_() {
  const r = selfCheck_(), bad = r.filter(function (x) { return !x.ok; });
  SpreadsheetApp.getUi().alert('Setup check', (bad.length ? '⚠ ' + bad.length + ' to fix:\n\n' + bad.map(function (x) { return '• ' + x.area + ': ' + x.text + '\n   Fix: ' + x.fix; }).join('\n\n') : '✅ Everything checks out.') +
    '\n\nAll good:\n' + r.filter(function (x) { return x.ok; }).map(function (x) { return '✓ ' + x.area + ': ' + x.text; }).join('\n'), SpreadsheetApp.getUi().ButtonSet.OK);
}

// file-stamp (tools/stamp-files.js): lets Admin › System tell whether this file is the newest one
var FILE_STAMP_45_SelfCheck = '4eaab60f5d';
