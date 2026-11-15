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
 * ARCADE PLAYS tab: When | 800# | Name | Squad | Game | Score | Points | Free | Paid.
 * ARCADE OFF hides it.
 */

const ARCADE_GAMES_ = {
  // max = most a real game can score, per = score for 1 squad point, minSec = shortest real game, rate = most score a second
  dash: { name: 'Summit Dash', blurb: 'Jump the rocks. The trail speeds up the farther you climb.', max: 8000, per: 400, minSec: 2, rate: 80 },
  catch: { name: 'Dough Catch', blurb: 'Catch Trait Dough, grab the gold, dodge the snowballs. 45 seconds.', max: 1500, per: 100, minSec: 40, rate: 40 },
  stack: { name: 'Peak Stack', blurb: 'Drop each block on the one below. Line it up perfectly to keep it wide.', max: 80, per: 6, minSec: 1, rate: 3 },
};
const ARCADE_HEADERS_ = ['When', '800#', 'Name', 'Squad', 'Game', 'Score', 'Points', 'Free', 'Paid'];

function arcadeOn_() { return String(getSetting_('ARCADE') || 'ON').toUpperCase() === 'ON'; }

function ensureArcadeTab_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getSheetByName('Arcade Plays')) return;
  const sh = ss.insertSheet('Arcade Plays');
  sh.getRange(1, 1, 1, ARCADE_HEADERS_.length).setValues([ARCADE_HEADERS_]).setFontWeight('bold').setBackground('#490e6f').setFontColor('#ffffff');
  sh.setFrozenRows(1);
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
    const id = meId_(asId, adminToken);
    const mine = sheetRows_('Arcade Plays').filter(function (p) { return id800_(p['800#']) === id; }), best = {};
    mine.forEach(function (p) { const g = String(p['Game']); best[g] = Math.max(best[g] || 0, Number(p['Score']) || 0); });
    return JSON.stringify({ on: true, wallet: arcadeWallet_(id), rules: arcadeTokenRules_(), maxPoints: settingNum_('ARCADE_MAX_POINTS', 3),
      games: Object.keys(ARCADE_GAMES_).map(function (k) { const g = ARCADE_GAMES_[k]; return { key: k, name: g.name, blurb: g.blurb, per: g.per, best: best[k] || 0 }; }),
      board: arcadeBoard_(), sample: id === SAMPLE_ID_ });
  });
}

/** Starts a play: uses the free play first, then a token. Returns { play }. */
function clashArcadeStart(game, asId, adminToken) {
  return withSample_(asId, function () {
    if (!arcadeOn_()) throw new Error('The Arcade is closed right now.');
    if (!ARCADE_GAMES_[game]) throw new Error('Pick a game.');
    const id = meId_(asId, adminToken), w = arcadeWallet_(id);
    if (id !== SAMPLE_ID_) {
      if (w.playsToday >= w.daily) throw new Error('That\'s all ' + w.daily + ' plays for today. Come back tomorrow!');
      if (!w.freeLeft && !w.tokens) throw new Error('No tokens yet. Earn them with Trait Dough, service and poll votes.');
    }
    const play = Utilities.getUuid().replace(/-/g, '');
    CacheService.getScriptCache().put('arc-' + play, JSON.stringify({ id: id, game: game, at: Date.now(), free: w.freeLeft > 0 }), 1200);
    return JSON.stringify({ play: play });
  });
}

/** Ends a play with its score. Returns the Arcade tab with { result: { score, points, kept } }. */
function clashArcadeFinish(play, score, asId, adminToken) {
  return withSample_(asId, function () {
    const id = meId_(asId, adminToken), cache = CacheService.getScriptCache();
    const raw = cache.get('arc-' + play);
    if (!raw) throw new Error('That game timed out. Start a new one.');
    const p = JSON.parse(raw), g = ARCADE_GAMES_[p.game];
    cache.remove('arc-' + play);
    if (p.id !== id) throw new Error('That game belongs to someone else.');
    score = Math.max(0, Math.floor(Number(score) || 0));
    const secs = (Date.now() - p.at) / 1000;
    const kept = secs >= g.minSec && score <= g.max && score <= (secs + 2) * g.rate;
    const points = kept ? Math.min(settingNum_('ARCADE_MAX_POINTS', 3), Math.floor(score / g.per)) : 0;
    if (id !== SAMPLE_ID_) { // every play uses its token; a score that can't be real is kept as 0
      ensureArcadeTab_();
      const b = boardRows_().filter(function (r) { return id800_(r['800#']) === id; })[0] || {};
      getSheet_('Arcade Plays').appendRow([Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'), id, String(b['Name'] || ''), squadOf_(b['Squad']), p.game, kept ? score : 0, points, !!p.free, '']);
    }
    const out = JSON.parse(clashArcade(asId, adminToken));
    out.result = { game: p.game, score: score, points: points, kept: kept };
    return JSON.stringify(out);
  });
}

// Once a day: the day's arcade points, one row per squad.
function payArcade_() {
  const sh = getSheetSafe_('Arcade Plays');
  if (!sh) return 0;
  const rows = sheetRows_('Arcade Plays').filter(function (r) { return !String(r['Paid'] || '').trim() && (Number(r['Points']) || 0) > 0 && SQUAD_KEYS_.indexOf(squadOf_(r['Squad'])) > -1; });
  if (!rows.length) return 0;
  const pts = { BLACK: 0, GOLD: 0, GREY: 0, PURPLE: 0 }, today = todayYmd_();
  rows.forEach(function (r) { pts[squadOf_(r['Squad'])] += Number(r['Points']) || 0; });
  getSheet_('events').appendRow(['', 'Arcade: ' + rows.length + ' game' + (rows.length === 1 ? '' : 's'), 'arcade', today, '', '', pts.BLACK || '', pts.GOLD || '', pts.GREY || '', pts.PURPLE || '']);
  const c = headerCol_(sh, 'Paid');
  rows.forEach(function (r) { if (c) sh.getRange(r._row, c).setValue(today); });
  return rows.length;
}

function maybePayArcade_() {
  const props = PropertiesService.getScriptProperties(), today = todayYmd_();
  if (props.getProperty('arcade-paid') === today) return;
  props.setProperty('arcade-paid', today);
  payArcade_();
}

// file-stamp (tools/stamp-files.js): lets Admin › System tell whether this file is the newest one
var FILE_STAMP_48_Arcade = '69627447cd';
