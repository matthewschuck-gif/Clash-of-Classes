/**
 * Clash of Classes -- Apps Script backend
 * File 32: the privacy check. Runs every public page's data the way a student, a visitor or the TV
 * gets it and looks for anything that should never be there:
 *   - an 800# (any student number on the Students tab or the Student Board, or any 800xxxx)
 *   - a student's full name ("Ann Lee" or "Lee, Ann" from the Students tab). "Ann L." is fine:
 *     that is how positive rows and leaderboards name students on purpose.
 * The public pages: site data, TV feed, Daily Polls, Pulse, Poll Board, frost explainer, This Week,
 * leaderboards, Trait Dough board, Year Plan, Homerooms, Squad Duels, event pages, the map, mascots,
 * challenges and You Said We Did. (My Climb and Generals HQ need a sign-in and show a student their
 * own things; they are not public.)
 * Menu › Privacy check, and every month in the State of the Squads email. Findings are masked
 * (the last 3 digits, initials) so the report itself never repeats them.
 */

function publicSurfaces_() {
  return [
    ['Site data', function () { return clashLoad(); }],
    ['TV feed', clashTvFeed],
    ['Daily Polls', function () { return clashDailyPolls(''); }],
    ['Pulse', clashPulseLive],
    ['Poll Board', clashPollBoard],
    ['Freeze & Thaw explainer', clashFrostExplainer],
    ['This Week', clashWeek],
    ['Leaderboards', clashLeaders],
    ['Trait Dough board', clashTraitDough],
    ['Year Plan', clashYearPlan],
    ['Homerooms', clashHomerooms],
    ['Squad Duels', clashDuels],
    ['Event pages', clashEventPages],
    ['The map', clashGame],
    ['Mascots', clashMascots],
    ['Challenges', clashChallenges],
    ['You Said We Did', clashYouSaid],
  ];
}

function privacyLists_() {
  const ids = {}, names = [];
  const add = function (first, last) {
    first = String(first || '').trim().split(/\s+/)[0]; last = String(last || '').trim();
    if (first.length < 2 || last.length < 2) return;
    names.push({ full: (first + ' ' + last).toLowerCase(), back: (last + ', ' + first).toLowerCase(), mask: first.charAt(0) + '. ' + last.charAt(0) + '.' });
  };
  sheetRows_('Students').forEach(function (r) {
    const id = id800_(r['800#']);
    if (id) ids[id] = 1;
    add(r['First Name'], r['Last Name']);
  });
  sheetRows_(getSetting_('STUDENT_BOARD_TAB') || 'Student Board').forEach(function (r) { const id = id800_(r['800#']); if (id) ids[id] = 1; });
  return { ids: Object.keys(ids), names: names };
}

/** [{ surface, ok, findings: [masked lines] }] */
function privacyScan_() {
  const lists = privacyLists_();
  return publicSurfaces_().map(function (s) {
    let text = '';
    try { text = String(s[1]() || ''); } catch (e) { return { surface: s[0], ok: true, findings: [], note: 'could not run: ' + e.message }; }
    const low = text.toLowerCase(), found = {};
    lists.ids.forEach(function (id) { if (text.indexOf(id) > -1) found['800# …' + id.slice(-3)] = 1; });
    (text.match(/\b800\d{4,}\b/g) || []).forEach(function (id) { found['800# …' + id.slice(-3)] = 1; });
    lists.names.forEach(function (n) {
      if (low.indexOf(n.full) > -1 || low.indexOf(n.back) > -1) found['full name ' + n.mask] = 1;
    });
    const findings = Object.keys(found);
    return { surface: s[0], ok: !findings.length, findings: findings };
  });
}

function menuPrivacyCheck_() {
  const r = privacyScan_(), bad = r.filter(function (x) { return x.findings.length; });
  SpreadsheetApp.getUi().alert('Privacy check: ' + r.length + ' public pages',
    (bad.length ? '⚠ Found on a public page:\n' + bad.map(function (x) { return x.surface + ': ' + x.findings.join(', '); }).join('\n') +
      '\n\nFix that page data before more students see it (the State of the Squads email lists it too).'
      : '✅ No 800# and no full student name on any public page.') +
    (r.some(function (x) { return x.note; }) ? '\n\nNot checked: ' + r.filter(function (x) { return x.note; }).map(function (x) { return x.surface + ' (' + x.note + ')'; }).join('; ') : ''),
    SpreadsheetApp.getUi().ButtonSet.OK);
}
