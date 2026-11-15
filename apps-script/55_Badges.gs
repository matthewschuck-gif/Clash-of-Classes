/**
 * Clash of Classes -- Apps Script backend
 * File 55: Clash badges on My Climb (Settings › BADGES ON).
 *
 * Small milestones a student earns: Trait Dough, service, polls, the Arcade, a Weather Day, an
 * Expedition, Helping Hands, a squad goal, Rivalry Week, the Team League, a staff shout-out, the
 * 8th grade send-off. Badges are worth no points, so nobody gets an edge from them. Nothing is
 * stored: they're worked out from the tabs each time My Climb opens, so fixing a row fixes the badge.
 * Only the student sees their own badges.
 */

// "Did something" days for students: Trait Dough given to them, service sent, Arcade, poll votes.
// { 800#: { 'yyyy-MM-dd': 1 } }. Pass an id to read only that student.
function activityDays_(only) {
  const out = {};
  const mark = function (id, d) { id = id800_(id); d = String(d || '').slice(0, 10); if (!id || !/^\d{4}-\d{2}-\d{2}$/.test(d) || (only && id !== only)) return; (out[id] = out[id] || {})[d] = 1; };
  sheetRows_('Virtual Dough').forEach(function (r) { mark(r['800#'], r['When']); });
  sheetRows_('Service Submissions').forEach(function (r) { mark(r['800#'], r['Submitted']); });
  sheetRows_('Arcade Plays').forEach(function (r) { mark(r['800#'], r['When']); });
  sheetRows_('Poll Votes').forEach(function (r) { mark(r['800#'], asYmd_(r['Date'] || r['At'])); });
  return out;
}

function anyDayIn_(days, from, to) { return Object.keys(days || {}).some(function (d) { return d >= from && (!to || d <= to); }); }

const BADGE_DEFS_ = [
  ['dough1', '🍞', 'First Dough', 'Get your first Trait Dough bill'],
  ['dough10', '🥖', 'Baker\'s Dozen', 'Get 10 Trait Dough bills'],
  ['service1', '🤝', 'Helper', 'Have a service project approved'],
  ['service5', '🏅', 'Service Star', 'Have 5 service projects approved'],
  ['family', '🏡', 'Family Team', 'Finish a family challenge'],
  ['votes10', '🗳️', 'Voice', 'Vote in 10 Daily Polls'],
  ['votes50', '📣', 'Loud and Clear', 'Vote in 50 Daily Polls'],
  ['arcade1', '🕹️', 'Player One', 'Play an Arcade game'],
  ['arcadeTop', '🏆', 'High Score', 'Earn the most points one Arcade game can give'],
  ['arcade5', '🎮', 'Regular', 'Play the Arcade on 5 different days'],
  ['weather', '🌨️', 'Storm Chaser', 'Earn something on a Weather Day'],
  ['expedition', '🧭', 'Explorer', 'Pitch in while your squad wins an Expedition'],
  ['hands', '🫶', 'Helping Hand', 'Pitch in during your squad\'s Helping Hands week'],
  ['goal', '🎯', 'Goal Getter', 'Pitch in the week your squad hits its goal'],
  ['goalSet', '🗺️', 'Trail Boss', 'Set a squad goal (Generals and captains)'],
  ['rivalry', '⚔️', 'Rival', 'Pitch in while your squad wins a Rivalry Week'],
  ['team', '🚩', 'Team Banner', 'Be on the Team League group that wins the month'],
  ['shoutout', '⭐', 'Shout-out', 'Get a staff shout-out'],
  ['signup', '📝', 'In the Game', 'Sign up for a squad event'],
  ['sendoff', '✉️', 'Trail Guide', 'Write advice for next year\'s 7th graders (8th grade)'],
];

/** [{ id, icon, name, how, earned }] for one student. */
function badgesFor_(id, squad) {
  const days = activityDays_(id)[id] || {}, got = {};
  const mine = function (tab) { return sheetRows_(tab).filter(function (r) { return id800_(r['800#']) === id; }); };
  const bills = mine('Virtual Dough').length + boardRows_().filter(function (r) { return id800_(r['800#']) === id; }).reduce(function (n, r) { return n + (Number(r['Trait Dough Bills']) || 0); }, 0);
  got.dough1 = bills >= 1; got.dough10 = bills >= 10;
  const svc = mine('Service Submissions').filter(function (r) { return String(r['Status']) === 'Approved'; });
  const fam = svc.filter(function (r) { return String(r['Kind'] || '') === 'Family'; }).length;
  got.service1 = svc.length - fam >= 1; got.service5 = svc.length - fam >= 5; got.family = fam >= 1;
  const votes = mine('Poll Votes').length;
  got.votes10 = votes >= 10; got.votes50 = votes >= 50;
  const plays = mine('Arcade Plays'), playDays = {};
  plays.forEach(function (p) { playDays[String(p['When']).slice(0, 10)] = 1; });
  got.arcade1 = plays.length >= 1; got.arcade5 = Object.keys(playDays).length >= 5;
  const top = settingNum_('ARCADE_MAX_POINTS', 3);
  got.arcadeTop = plays.some(function (p) { return ARCADE_GAMES_[String(p['Game'])] && (Number(p['Points']) || 0) >= top; });
  try { got.weather = weatherRows_().some(function (w) { return anyDayIn_(days, w.from, w.to); }); } catch (e) { got.weather = false; }
  try { got.expedition = expeditions_().some(function (x) { return x.winners.indexOf(squad) > -1 && anyDayIn_(days, x.start || '', x.wonOn); }); } catch (e) { got.expedition = false; }
  try {
    got.hands = sheetRows_('Helping Hands').some(function (r) { const w = asYmd_(r['Week']).slice(0, 10); return squadOf_(r['Squad']) === squad && anyDayIn_(days, w, ymdShift_(w, 6)); });
    const goals = sheetRows_('Squad Goals');
    got.goal = goals.some(function (r) { const w = asYmd_(r['Week']).slice(0, 10); return squadOf_(r['Squad']) === squad && /hit/.test(String(r['Result'])) && anyDayIn_(days, w, ymdShift_(w, 6)); });
    const me = squadLeader_(id);
    got.goalSet = !!me && goals.some(function (r) { return String(r['Set By']) === (me.name || me.id); });
  } catch (e) { /* no tabs yet */ }
  try { got.rivalry = rivalryDuels_().some(function (d) { return d.awarded && d.leaders.indexOf(squad) > -1 && anyDayIn_(days, d.start, d.end); }); } catch (e) { got.rivalry = false; }
  try { const t = teamOf_()[id]; got.team = !!t && sheetRows_('Team League').some(function (r) { return squadOf_(r['Squad']) === squad && String(r['Team']).toUpperCase() === t; }); } catch (e) { got.team = false; }
  try { got.shoutout = mine('Staff Shout-outs').some(function (r) { return ticked_(r['Approved']); }); } catch (e) { got.shoutout = false; }
  got.signup = mine('Sign-ups').length >= 1;
  got.sendoff = mine('Send-off Advice').some(function (r) { return String(r['Advice'] || '').trim(); });
  return BADGE_DEFS_.map(function (b) { return { id: b[0], icon: b[1], name: b[2], how: b[3], earned: !!got[b[0]] }; });
}

// file-stamp (tools/stamp-files.js): lets Admin › System tell whether this file is the newest one
var FILE_STAMP_55_Badges = '';

// file-stamp (tools/stamp-files.js): lets Admin › System tell whether this file is the newest one
var FILE_STAMP_55_Badges = 'd5a7bab840';
