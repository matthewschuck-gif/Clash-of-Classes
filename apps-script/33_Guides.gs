/**
 * Clash of Classes -- Apps Script backend
 * File 33: one-page role guides, written from the live Settings and map rules (?page=guides).
 *
 * Four guides, each fits on one printed page: Students, Squad Generals, Teachers and staff,
 * Coordinator and admins. Every number in them comes from the Sheet (poll points, Pulse days,
 * the sign-up limit, LRG points, evolution stages, each area's rule and multiplier, the freeze
 * lines...), so change a Setting and the guide changes with it. Nothing secret goes in: never the
 * Generals code, a password, a sheet ID or an 800#.
 * Menu › Open the role guides. The page has a Print button (one guide per page).
 */

function guideList_(spec) {
  return String(spec || '').split(/[;,]/).map(function (p) { return p.trim(); }).filter(String).map(function (p) {
    const i = p.indexOf('=');
    return i > 0 ? [p.slice(0, i).trim(), p.slice(i + 1).trim()] : [p, ''];
  });
}

/** The guides as data: { site, made, roles: [{ id, icon, title, lead, sections: [{ h, items: [text] }] }] } */
function guidesData_() {
  const S = function (k, def) { const v = getSetting_(k); return v === '' || v == null ? def : String(v); };
  const on = function (k, def) { return String(S(k, def)).toUpperCase() === 'ON'; };
  const site = S('SITE_URL', 'clashofclasses.org').replace(/^https?:\W+/i, '').replace(/\/$/, '');
  const pulseDays = S('PULSE_DAYS', '3'), checkinDays = S('CHECKIN_DAYS', '1'), limit = S('SIGNUP_LIMIT', '3');
  const lrgNames = { bills: 'Trait Dough bill', mindset: 'Mindset summit', shoutouts: 'shout-out', star: 'STAR growth point', streak: 'streak week', buddy: 'Camp buddy',
    tickets: 'destination ticket', passport: 'Trail Passport', journal: 'Trail Journal growth', redemption: 'Trailback Summit' };
  const lrg = guideList_(S('LRG_POINTS', '')).map(function (x) { return (lrgNames[x[0]] || x[0]) + ' = ' + x[1]; });
  lrg.push('Trailback Summit = ' + S('REDEMPTION_LRG', '25'));
  const stages = S('EVOLVE_NAMES', 'Foothills, Ridge, High Camp, Summit').split(',').map(function (s) { return s.trim(); });
  const evolve = guideList_(S('EVOLVE_LRG', '2=30, 3=60, 4=100')).map(function (x) { return (stages[Number(x[0]) - 1] || 'Stage ' + x[0]) + ' at ' + x[1] + ' LRG'; });
  let zones = [];
  try {
    const rules = conquestRules_();
    zones = sheetRows_('Map Zones').filter(function (r) { return ticked_(r['Show']) && String(r['Zone'] || '').trim(); }).map(function (r) {
      const rule = rules[String(r['Conquer By'] || '').trim().toLowerCase()] || {};
      const mult = r['Multiplier'] === '' || r['Multiplier'] == null ? 1 : Number(r['Multiplier']) || 0;
      return { zone: String(r['Zone']).trim(), rule: rule.rule || 'Most points', text: rule.text || '', mult: mult, summit: ticked_(r['Summit']), weekly: Number(r['Weekly Points']) || 0 };
    });
  } catch (e) { zones = []; }
  const zoneLine = function (z) { return z.zone + (z.summit ? ' (Summit)' : '') + ': ' + (z.text || z.rule) + ' Pays ' + z.weekly + ' a week x' + z.mult + '.'; };
  const doubles = zones.filter(function (z) { return z.mult >= 2; }).map(function (z) { return z.zone + ' (' + z.rule + ')'; });
  let lines = [];
  try { lines = freezeLines_(); } catch (e) { lines = []; }
  const lineText = lines.map(function (L) { return L.label + ': ' + (L.below ? 'below ' : 'above ') + L.line + (L.per100 ? ' per 100 students' : '') + ' (' + L.kind + ')'; });
  const pollPts = on('DAILY_POLL_POINTS', 'OFF') ? 'every vote earns your squad a point, and picking the winning side (or the right trivia answer) earns one more' : 'votes count toward the map (Auditorium) even before poll points are turned on';

  const roles = [
    { id: 'students', icon: '🧗', title: 'Students', lead: 'How you help your squad climb, and what stays private.', sections: [
      { h: 'Earn points for your squad', items: [
        'Daily Poll on ' + site + ': answer with your 800#; ' + pollPts + '.',
        'Pulse question: open ' + pulseDays + ' school days, no right answer, every answer is a squad point. Your answer helps the school decide things.',
        'Check-ins: Pride Day and the Wheel are one tap with your 800#. Squad events (Move, Make, Talk): check in on the day' + (Number(checkinDays) > 1 ? ' (or the next ' + (Number(checkinDays) - 1) + ')' : '') + ' and a General confirms it.',
        'Sign up for squad events: up to ' + limit + ' events, never two at the same time.',
        'Trait Dough, shout-outs, Mindset summits and STAR growth all add up for your squad.',
      ] },
      { h: 'My Climb (your creature)', items: [
        'Open My Climb from the site with your school Google account, ' + S('MY_CLIMB_HOURS', '07:00-16:00') + ' on school days.',
        'LRG points: ' + lrg.join(', ') + '.',
        'Your creature evolves: ' + evolve.join(', ') + '.',
      ] },
      { h: 'The school map', items: zones.length ? zones.map(zoneLine).concat(['Every area is judged on last calendar month (the Front Office on this month), so it changes hands on the 1st.']) : ['Areas of the building are won on their own data each month.'] },
      { h: 'Freeze & Thaw', items: ['Points are never taken away. A squad can freeze: its new points stop counting until it thaws.',
        'The same lines for every squad: ' + (lineText.join('; ') || 'set by staff') + '. One line crossed = a frost warning first.',
        'A squad thaws together: fix the line, everyone pitches in, earn points together, and a squad reflection. No one thaws alone.'] },
      { h: 'Private', items: ['Your 800# is only for checking in and voting. Results on the site and the TV are counts, never names.',
        'Positive things can show your first name and last initial. Behavior never shows on the site.'] },
    ] },
    { id: 'generals', icon: '🎖️', title: 'Squad Generals', lead: 'Your squad, every week. Generals HQ is on the site: sign in with your 800# and the code from your coordinator.', sections: [
      { h: 'Every week', items: [
        'Confirm squad event check-ins in HQ: only confirmed members count for the Gym 1 area.',
        'Place everyone in at least one event before anyone doubles up (up to ' + limit + ' each).',
        'Draw your chance card (' + S('CHANCE_DRAWS_PER_WEEK', '1') + ' a week) and decide on the Hot Streak: ride it or bank it (' + S('STREAK_BANK', '') + ').',
        'Hype the Daily Poll and the Pulse: turnout wins the Auditorium.',
      ] },
      { h: 'Where your squad can win', items: (doubles.length ? ['These areas pay double because they are won by showing up and checking in: ' + doubles.join(', ') + '.'] : [])
        .concat(zones.filter(function (z) { return z.mult < 2; }).map(zoneLine)) },
      { h: 'Freezes', items: ['A Peaks freeze thaws on group goals and a squad reflection you lead (staff sign off).',
        'A Fun freeze is push your luck: cash out in HQ at a tier (' + S('FUN_THAW_TIERS', '1=25, 2=50, 3=100') + ').',
        'Squad meeting slides (Freeze & Thaw) explain it in plain words: use them in January.'] },
      { h: 'Your job', items: ['Read the Generals handbook (link in HQ › My job). Be the first to cheer, never the first to blame.'] },
    ] },
    { id: 'staff', icon: '🍎', title: 'Teachers and staff', lead: 'Five minutes a week keeps squads going.', sections: [
      { h: 'Give points', items: ['Trait Dough: hand out bills for the traits you see. Turned-in bills are points for the squad.',
        'Shout-outs and Student of the Month: every one counts for the 60s Hall (shout-outs) and My Climb.',
        'Mindset lessons: summits count for the Media Center (Growth Grind).'] },
      { h: 'In homeroom', items: ['Remind students about the Daily Poll and the Pulse (' + pulseDays + ' days open). Put the live Poll Board on the projector if you can.',
        'Pride Day: students check in themselves on the site. That decides Gym 2.'] },
      { h: 'Squad events', items: ['Running a Move, Make or Talk event? Students check in; Generals confirm. Points from the event go in as its type (Move, Make, Talk).'] },
      { h: 'First Friday Coffee', items: ['Staff squads compete each month. The winner is entered as a Staff squad challenge winner and takes the Front Office right away.'] },
      { h: 'Pulse answers', items: ['Results by squad and grade are on the Poll Insights tab (counts only). If a question is marked "plan follow-up first", the coordinator decides who reads those answers.'] },
    ] },
    { id: 'admin', icon: '🗝️', title: 'Coordinator and admins', lead: 'What runs by itself, and what needs you.', sections: [
      { h: 'By itself', items: ['Climb-er fills the Squad Monthly tab every hour (Students, attendance, failing, interventions). LRR + LD comes from the Follow-Up sheets once a day.',
        'The map judges every area on last month on the 1st. Frost check on the 1st (from ' + S('FROST_START', '2027-01') + ').',
        'The weekly digest (' + (on('DIGEST', 'OFF') ? 'on, ' + S('DIGEST_DAY', 'FRIDAY') : 'off') + ') and the State of the Squads email (' + (on('STATE_OF_SQUADS', 'ON') ? 'on the ' + S('STATE_DAY', '2') + ' of each month' : 'off') + ').'] },
      { h: 'Needs you', items: ['Approve each proposed freeze on the Frost Log tab (nothing freezes without you).',
        'You Said We Did tab: type what changed after a Pulse and tick Show; students see it on the Pulse page and the TV.',
        'Admin › Map: lock an area, set a multiplier (any number), or override one squad\'s number for a month.',
        'Polls: menu Schedule poll ideas or Add the next Pulse question.'] },
      { h: 'Checks', items: ['Menu Points balance check: is one source deciding the race? (lines on Settings: BALANCE_*)',
        'Menu Privacy check: every public page, no 800# and no full student name.'] },
      { h: 'Multipliers now', items: zones.map(function (z) { return z.zone + ' x' + z.mult; }) },
    ] },
  ];
  return { site: site, made: todayYmd_(), roles: roles };
}

function clashGuides() { return JSON.stringify(guidesData_()); }

function serveGuides_() {
  return HtmlService.createHtmlOutputFromFile('Guides')
    .setTitle('Role guides · Clash of Classes')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function menuOpenGuides_() {
  const url = ScriptApp.getService().getUrl() + '?page=guides';
  SpreadsheetApp.getUi().alert('Role guides', 'One page each for Students, Squad Generals, Teachers and staff, and Coordinator and admins, written from today\'s Settings:\n\n' + url +
    '\n\nPrint button on the page (one guide per page). Add &role=students (or generals, staff, admin) for one guide.', SpreadsheetApp.getUi().ButtonSet.OK);
}
