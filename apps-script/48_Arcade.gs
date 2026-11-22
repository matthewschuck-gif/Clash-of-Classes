/**
 * Clash of Classes -- Apps Script backend
 * File 48: the Arcade on My Climb (signed in with school Google, open MY_CLIMB_HOURS). Three short
 * games (Arcade.html): Summit Dash (jump the rocks), Dough Catch (catch Trait Dough, dodge snowballs)
 * and Peak Stack (stack the summit). A play costs a token; ARCADE_FREE_DAILY plays a day are free.
 *   Tokens (ARCADE_TOKENS): dough=1 per Trait Dough bill (paper or Virtual), service=2 per approved
 *   service, family=2 per approved family challenge, votes=5 (1 token per 5 poll / Pulse votes).
 *   ARCADE_DAILY plays a day in all. Each play earns the squad up to ARCADE_MAX_POINTS (score / the
 *   game's "per"), paid once a day as one Arcade row per squad.
 *   A score is only kept if the game ran long enough and the score is possible for that game.
 *   The week's best scores show first name + last initial only when LEADERBOARD_NAMES is ON.
 * ARCADE PLAYS tab: When | 800# | Name | Squad | Game | Score | Points | Free | Paid | Pebbles.
 * ARCADE OFF hides it.
 *
 * PEBBLES (ARCADE_PEBBLES ON, the default): instead of squad points, each game earns the student
 * Pebbles, the Arcade's own currency. ARCADE_PEBBLES_PER_POINT (100) Pebbles = 1 Clash point, and
 * only whole points move: the student taps Cash in on the Arcade, the Pebbles come out of their
 * pouch, and the points go to their squad in that night's Arcade row. Leftover Pebbles stay.
 *   Earning (ARCADE_PEBBLE_RATES, score for 1 Pebble): Summit Dash 120, Dough Catch 50, Peak Stack 2,
 *   so an average game earns about 6 and a strong one about 12. Every game that counts earns at
 *   least 1; at most ARCADE_PEBBLE_MAX (15). With 3 plays a day, a strong daily player fills a point
 *   in about a week, an average one in about two.
 *   PEBBLE CASH-INS tab: When | 800# | Squad | Pebbles | Points | Paid (its own tab, so a cash-in
 *   never counts as a play for badges, goals or the Team League).
 * ARCADE_PEBBLES OFF = the old way (each game pays up to ARCADE_MAX_POINTS squad points).
 */

const ARCADE_GAMES_ = {
  // max = most a real game can score, per = score for 1 squad point, minSec = shortest real game, rate = most score a second
  dash: { name: 'Summit Dash', blurb: 'Jump the rocks. The trail speeds up the farther you climb.', max: 8000, per: 400, minSec: 2, rate: 80 },
  catch: { name: 'Dough Catch', blurb: 'Catch Trait Dough, grab the gold, dodge the snowballs. 45 seconds.', max: 1500, per: 100, minSec: 40, rate: 40 },
  stack: { name: 'Peak Stack', blurb: 'Drop each block on the one below. Line it up perfectly to keep it wide.', max: 80, per: 6, minSec: 1, rate: 3 },
};
const ARCADE_HEADERS_ = ['When', '800#', 'Name', 'Squad', 'Game', 'Score', 'Points', 'Free', 'Paid', 'Pebbles'];
const CASHIN_HEADERS_ = ['When', '800#', 'Squad', 'Pebbles', 'Points', 'Paid'];
const PEBBLE_BONUS_HEADERS_ = ['When', '800#', 'Squad', 'Pebbles', 'From']; // Pebbles from outside the Arcade (Summit Party)

// CLASS TIME (ARCADE_CLASS_TIMES, e.g. "08:15-11:15, 12:33-14:05"): no Arcade games during instruction.
// Cash in still works. Blank = no class-time lock (My Climb's own hours still apply).
function arcadeClassTimes_() {
  return String(getSetting_('ARCADE_CLASS_TIMES') || '').split(',').map(function (p) {
    const m = p.trim().match(/^(\d{1,2}):(\d{2})\s*-\s*(\d{1,2}):(\d{2})$/);
    return m ? [('0' + m[1]).slice(-2) + ':' + m[2], ('0' + m[3]).slice(-2) + ':' + m[4]] : null;
  }).filter(Boolean);
}
function clock12_(hm) { const h = Number(hm.slice(0, 2)), m = hm.slice(3); return ((h + 11) % 12 + 1) + ':' + m; }
// { closed, until, text } right now.
function arcadeClassTime_(now) {
  const list = arcadeClassTimes_(), hm = Utilities.formatDate(now || new Date(), Session.getScriptTimeZone(), 'HH:mm');
  const hit = list.filter(function (w) { return hm >= w[0] && hm < w[1]; })[0];
  return { closed: !!hit, until: hit ? clock12_(hit[1]) : '', text: list.map(function (w) { return clock12_(w[0]) + '–' + clock12_(w[1]); }).join(' and ') };
}
// The "What are Pebbles?" card: the why is a Settings row (ARCADE_PEBBLE_WHY), the how is built from the settings.
function pebbleAbout_() {
  return String(getSetting_('ARCADE_PEBBLE_WHY') || '').trim() ||
    'Hikers mark the trail with cairns: stacks of small stones. One pebble does not look like much, but stack enough and everyone can see the way. Pebbles work the same way: lots of small games add up to something real for your squad.';
}

function pebblesOn_() { return String(getSetting_('ARCADE_PEBBLES') || 'ON').toUpperCase() !== 'OFF'; }
function pebblesPerPoint_() { return Math.max(1, settingNum_('ARCADE_PEBBLES_PER_POINT', 100)); }
function pebbleRates_() {
  const out = { dash: 120, catch: 50, stack: 2 };
  String(getSetting_('ARCADE_PEBBLE_RATES') || '').split(',').forEach(function (p) { const m = p.trim().match(/^(\w+)\s*=\s*([\d.]+)$/); if (m && out[m[1]] !== undefined && Number(m[2]) > 0) out[m[1]] = Number(m[2]); });
  return out;
}
// Pebbles for one kept game: score / the game's rate, at least 1, at most ARCADE_PEBBLE_MAX.
function pebblesFor_(game, score) {
  const rate = pebbleRates_()[game] || 100;
  return Math.max(1, Math.min(Math.max(1, settingNum_('ARCADE_PEBBLE_MAX', 15)), Math.floor(score / rate)));
}
// { balance, earned, cashed, per, max, rates, week } for one student.
function pebblePouch_(id) {
  const mon = mondayOf_(todayYmd_());
  let earned = 0, cashed = 0, week = 0;
  sheetRows_('Arcade Plays').forEach(function (p) { if (id800_(p['800#']) !== id) return; const n = Number(p['Pebbles']) || 0; earned += n; if (String(p['When']).slice(0, 10) >= mon) week += n; });
  sheetRows_('Pebble Bonuses').forEach(function (p) { if (id800_(p['800#']) !== id) return; const n = Number(p['Pebbles']) || 0; earned += n; if (String(p['When']).slice(0, 10) >= mon) week += n; });
  sheetRows_('Pebble Cash-ins').forEach(function (c) { if (id800_(c['800#']) === id) cashed += Number(c['Pebbles']) || 0; });
  return { on: pebblesOn_(), balance: Math.max(0, earned - cashed), earned: earned, cashed: cashed, week: week, per: pebblesPerPoint_(),
    max: Math.max(1, settingNum_('ARCADE_PEBBLE_MAX', 15)), rates: pebbleRates_() };
}

function arcadeOn_() { return String(getSetting_('ARCADE') || 'ON').toUpperCase() === 'ON'; }

function ensureArcadeTab_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  [['Arcade Plays', ARCADE_HEADERS_], ['Pebble Cash-ins', CASHIN_HEADERS_], ['Pebble Bonuses', PEBBLE_BONUS_HEADERS_]].forEach(function (t) {
    const have = ss.getSheetByName(t[0]);
    if (have) { // an older tab: add any missing header at the end (Pebbles)
      const h = have.getRange(1, 1, 1, Math.max(1, have.getLastColumn())).getValues()[0].map(function (x) { return String(x).trim(); });
      t[1].forEach(function (k) { if (h.indexOf(k) === -1) { h.push(k); have.getRange(1, h.length).setValue(k).setFontWeight('bold').setBackground('#490e6f').setFontColor('#ffffff'); } });
      return;
    }
    const sh = ss.insertSheet(t[0]);
    sh.getRange(1, 1, 1, t[1].length).setValues([t[1]]).setFontWeight('bold').setBackground('#490e6f').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  });
}
// Appends one row by header name (the tab's own column order).
function appendByHeader_(name, o) {
  const sh = getSheet_(name), h = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(function (x) { return String(x).trim(); });
  sh.appendRow(h.map(function (k) { return o[k] === undefined ? '' : o[k]; }));
}

function arcadeTokenRules_() {
  const out = { dough: 1, service: 2, family: 2, votes: 5 };
  String(getSetting_('ARCADE_TOKENS') || '').split(',').forEach(function (p) { const m = p.trim().match(/^(\w+)\s*=\s*([\d.]+)$/); if (m && out[m[1]] !== undefined) out[m[1]] = Number(m[2]); });
  return out;
}

/** { earned, spent, tokens, playsToday, freeLeft } for one student. */
function arcadeWallet_(id) {
  const r = arcadeTokenRules_(), today = todayYmd_();
  let bills = 0;
  boardRows_().forEach(function (b) { if (id800_(b['800#']) === id) bills += Number(b['Trait Dough Bills']) || 0; });
  bills += sheetRows_('Virtual Dough').filter(function (v) { return id800_(v['800#']) === id; }).length;
  const svc = sheetRows_('Service Submissions').filter(function (s) { return id800_(s['800#']) === id && String(s['Status']) === 'Approved'; });
  const service = svc.filter(function (s) { return String(s['Kind'] || 'Service') !== 'Family'; }).length, family = svc.length - service;
  const votes = sheetRows_('Poll Votes').filter(function (v) { return id800_(v['800#']) === id; }).length;
  const earned = Math.floor(bills * r.dough + service * r.service + family * r.family + (r.votes > 0 ? Math.floor(votes / r.votes) : 0));
  const plays = sheetRows_('Arcade Plays').filter(function (p) { return id800_(p['800#']) === id; });
  const spent = plays.filter(function (p) { return !ticked_(p['Free']); }).length;
  const todays = plays.filter(function (p) { return String(p['When']).slice(0, 10) === today; });
  const freeLeft = Math.max(0, settingNum_('ARCADE_FREE_DAILY', 1) - todays.filter(function (p) { return ticked_(p['Free']); }).length);
  return { earned: earned, spent: spent, tokens: Math.max(0, earned - spent), playsToday: todays.length, freeLeft: freeLeft, daily: Math.max(1, settingNum_('ARCADE_DAILY', 3)) };
}

function arcadeBoard_() {
  const mon = mondayOf_(todayYmd_()), names = String(getSetting_('LEADERBOARD_NAMES') || 'OFF').toUpperCase() === 'ON', out = {};
  const plays = sheetRows_('Arcade Plays').filter(function (p) { return String(p['When']).slice(0, 10) >= mon; });
  Object.keys(ARCADE_GAMES_).forEach(function (g) {
    out[g] = plays.filter(function (p) { return String(p['Game']) === g; }).sort(function (a, b) { return (Number(b['Score']) || 0) - (Number(a['Score']) || 0); }).slice(0, 5)
      .map(function (p) { return { who: names ? String(p['Name'] || '') : '', squad: squadOf_(p['Squad']), score: Number(p['Score']) || 0 }; });
  });
  return out;
}

/** The Arcade tab on My Climb. */
function clashArcade(asId, adminToken) {
  return withSample_(asId, function () {
    if (!arcadeOn_()) return JSON.stringify({ on: false });
    return arcadeView_(meId_(asId, adminToken));
  });
}
// The Arcade for one student (My Climb and the main site share it).
function arcadeView_(id) {
  {
    const mine = sheetRows_('Arcade Plays').filter(function (p) { return id800_(p['800#']) === id; }), best = {};
    mine.forEach(function (p) { const g = String(p['Game']); best[g] = Math.max(best[g] || 0, Number(p['Score']) || 0); });
    const pouch = pebblePouch_(id);
    pouch.why = pebbleAbout_();
    const hours = arcadeClassTime_();
    if (id === SAMPLE_ID_) hours.closed = false; // staff previews can always play
    return JSON.stringify({ on: true, wallet: arcadeWallet_(id), rules: arcadeTokenRules_(), maxPoints: settingNum_('ARCADE_MAX_POINTS', 3), pebbles: pouch, hours: hours,
      games: Object.keys(ARCADE_GAMES_).map(function (k) { const g = ARCADE_GAMES_[k]; return { key: k, name: g.name, blurb: g.blurb, per: g.per, pebble: pouch.rates[k], best: best[k] || 0 }; }),
      board: arcadeBoard_(), sample: id === SAMPLE_ID_, look: seasonLook_() });
  }
}

/** Starts a play: uses the free play first, then a token. Returns { play }. */
function clashArcadeStart(game, asId, adminToken) {
  return withSample_(asId, function () { return arcadeStart_(game, meId_(asId, adminToken)); });
}
function arcadeStart_(game, id) {
  {
    if (!arcadeOn_()) throw new Error('The Arcade is closed right now.');
    if (!ARCADE_GAMES_[game]) throw new Error('Pick a game.');
    const ct = arcadeClassTime_();
    if (ct.closed && id !== SAMPLE_ID_) throw new Error('It\'s class time, so the Arcade is resting. It opens again at ' + ct.until + '.');
    const w = arcadeWallet_(id);
    if (id !== SAMPLE_ID_) {
      if (w.playsToday >= w.daily) throw new Error('That\'s all ' + w.daily + ' plays for today. Come back tomorrow!');
      if (!w.freeLeft && !w.tokens) throw new Error('No tokens yet. Earn them with Trait Dough, service and poll votes.');
    }
    const play = Utilities.getUuid().replace(/-/g, '');
    CacheService.getScriptCache().put('arc-' + play, JSON.stringify({ id: id, game: game, at: Date.now(), free: w.freeLeft > 0 }), 1200);
    return JSON.stringify({ play: play });
  }
}

// ── The Arcade on the main site (ARCADE_ON_SITE): the same plays, tokens and points, by 800# ──
// Like the Pulse, a student types their 800# (it must be on the Student Board). Open the same
// hours as My Climb (MY_CLIMB_HOURS, school days). Lookups are capped like My Squad's.
function siteArcadeId_(id) {
  if (String(getSetting_('ARCADE_ON_SITE') || 'ON').toUpperCase() === 'OFF') throw new Error('The Arcade is on My Climb. Sign in there to play.');
  if (!arcadeOn_()) throw new Error('The Arcade is closed right now. Check back soon.');
  const cache = CacheService.getScriptCache(), used = Number(cache.get('arc-site-calls')) || 0;
  if (used >= (Number(getSetting_('POLL_VOTES_PER_10MIN')) || 1500)) throw new Error('The Arcade is busy right now! Try again in a minute.');
  cache.put('arc-site-calls', String(used + 1), 600);
  const sid = id800_(id);
  if (sid.length < 5) throw new Error('Type your 800 number.');
  if (!climbOpen_()) throw new Error('The Arcade is open ' + String(getSetting_('MY_CLIMB_HOURS') || '07:00-16:00').replace('-', ' to ') + ' on school days. See you on the trail!');
  const me = boardRows_().filter(function (r) { return id800_(r['800#']) === sid; })[0];
  if (!me) throw new Error('That 800 number isn\'t on the list. Check it, or ask your homeroom teacher.');
  return { id: sid, squad: squadOf_(me['Squad']) };
}
function withSquad_(text, squad) { const o = JSON.parse(text); o.squad = squad; return JSON.stringify(o); }
function clashSiteArcade(id) {
  const me = siteArcadeId_(id);
  return withSquad_(arcadeView_(me.id), me.squad);
}
function clashSiteArcadeStart(game, id) { return arcadeStart_(game, siteArcadeId_(id).id); }
function clashSiteArcadeFinish(play, score, id) {
  const me = siteArcadeId_(id);
  return withSquad_(arcadeFinish_(play, score, me.id), me.squad);
}

/** Ends a play with its score. Returns the Arcade tab with { result: { score, points, kept } }. */
function clashArcadeFinish(play, score, asId, adminToken) {
  return withSample_(asId, function () { return arcadeFinish_(play, score, meId_(asId, adminToken)); });
}
function arcadeFinish_(play, score, id) {
  {
    const cache = CacheService.getScriptCache();
    const raw = cache.get('arc-' + play);
    if (!raw) throw new Error('That game timed out. Start a new one.');
    const p = JSON.parse(raw), g = ARCADE_GAMES_[p.game];
    cache.remove('arc-' + play);
    if (p.id !== id) throw new Error('That game belongs to someone else.');
    score = Math.max(0, Math.floor(Number(score) || 0));
    const secs = (Date.now() - p.at) / 1000;
    const kept = secs >= g.minSec && score <= g.max && score <= (secs + 2) * g.rate;
    const peb = pebblesOn_();
    const points = kept && !peb ? Math.min(settingNum_('ARCADE_MAX_POINTS', 3), Math.floor(score / g.per)) : 0;
    const pebbles = kept && peb ? pebblesFor_(p.game, score) : 0;
    if (id !== SAMPLE_ID_) { // every play uses its token; a score that can't be real is kept as 0
      ensureArcadeTab_();
      const b = boardRows_().filter(function (r) { return id800_(r['800#']) === id; })[0] || {};
      appendByHeader_('Arcade Plays', { 'When': Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'), '800#': id, 'Name': String(b['Name'] || ''),
        'Squad': squadOf_(b['Squad']), 'Game': p.game, 'Score': kept ? score : 0, 'Points': points, 'Free': !!p.free, 'Paid': '', 'Pebbles': pebbles });
    }
    const out = JSON.parse(arcadeView_(id));
    out.result = { game: p.game, score: score, points: points, kept: kept, pebbles: pebbles };
    return JSON.stringify(out);
  }
}

/** Cash in: every full ARCADE_PEBBLES_PER_POINT Pebbles become 1 point for the student's squad (paid tonight). */
function clashArcadeCashIn(asId, adminToken) {
  return withSample_(asId, function () { return arcadeCashIn_(meId_(asId, adminToken)); });
}
function clashSiteArcadeCashIn(id) {
  const me = siteArcadeId_(id);
  return withSquad_(arcadeCashIn_(me.id), me.squad);
}
function arcadeCashIn_(id) {
  if (!arcadeOn_() || !pebblesOn_()) throw new Error('Cash in is closed right now.');
  if (id === SAMPLE_ID_) throw new Error('This is the sample student: Pebbles are pretend here, so nothing cashes in.');
  const lock = LockService.getScriptLock(); lock.waitLock(15000);
  let pts = 0;
  try {
    ensureArcadeTab_();
    const pouch = pebblePouch_(id);
    pts = Math.floor(pouch.balance / pouch.per);
    if (pts < 1) throw new Error('You need ' + pouch.per + ' Pebbles for 1 point. You have ' + pouch.balance + '.');
    const b = boardRows_().filter(function (r) { return id800_(r['800#']) === id; })[0] || {}, squad = squadOf_(b['Squad']);
    if (SQUAD_KEYS_.indexOf(squad) === -1) throw new Error('You are not on a squad yet. Ask your homeroom teacher.');
    appendByHeader_('Pebble Cash-ins', { 'When': Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'), '800#': id, 'Squad': squad, 'Pebbles': pts * pouch.per, 'Points': pts, 'Paid': '' });
  } finally { lock.releaseLock(); }
  const out = JSON.parse(arcadeView_(id));
  out.result = { cashed: pts };
  return JSON.stringify(out);
}

// Once a day: the day's arcade points (games, and Pebble cash-ins), one row per squad.
function payArcade_() {
  const pts = { BLACK: 0, GOLD: 0, GREY: 0, PURPLE: 0 }, today = todayYmd_(), marks = [];
  let games = 0, cash = 0;
  [['Arcade Plays', function () { games++; }], ['Pebble Cash-ins', function () { cash++; }]].forEach(function (t) {
    const sh = getSheetSafe_(t[0]);
    if (!sh) return;
    const rows = sheetRows_(t[0]).filter(function (r) { return !String(r['Paid'] || '').trim() && (Number(r['Points']) || 0) > 0 && SQUAD_KEYS_.indexOf(squadOf_(r['Squad'])) > -1; });
    rows.forEach(function (r) { pts[squadOf_(r['Squad'])] += Number(r['Points']) || 0; t[1](); });
    const c = headerCol_(sh, 'Paid');
    if (c) rows.forEach(function (r) { marks.push([sh, r._row, c]); });
  });
  if (!games && !cash) return 0;
  const name = 'Arcade: ' + [games ? games + ' game' + (games === 1 ? '' : 's') : '', cash ? cash + ' Pebble cash-in' + (cash === 1 ? '' : 's') : ''].filter(String).join(', ');
  getSheet_('events').appendRow(['', name, 'arcade', today, '', '', pts.BLACK || '', pts.GOLD || '', pts.GREY || '', pts.PURPLE || '']);
  marks.forEach(function (m) { m[0].getRange(m[1], m[2]).setValue(today); });
  return games + cash;
}

function maybePayArcade_() {
  const props = PropertiesService.getScriptProperties(), today = todayYmd_();
  if (props.getProperty('arcade-paid') === today) return;
  props.setProperty('arcade-paid', today);
  payArcade_();
}

// file-stamp (tools/stamp-files.js): lets Admin › System tell whether this file is the newest one
var FILE_STAMP_48_Arcade = 'adc1c1d193';
