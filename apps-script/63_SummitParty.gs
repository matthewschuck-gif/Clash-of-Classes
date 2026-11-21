/**
 * Clash of Classes -- Apps Script backend
 * File 63: Summit Party, the squad board game.
 *
 * The board shows on one big screen: clashofclasses.org/?page=summit (SummitParty.html). Each squad
 * plays from its own Chromebook: ?page=summit-play (SummitPlay.html), where a student types an
 * 800# (checked on the Student Board). On their squad's turn they roll; at a fork they pick a path.
 *
 * The game: four squads race around a mountain trail. Trail spaces give coins, rockslides take
 * them, cairns draw a trail card, springs add 2 to the next roll. One space holds the Summit Flag:
 * pass it (or land on it) with SUMMIT_PARTY_FLAG_COST coins and your squad plants a flag; then
 * the flag moves. After SUMMIT_PARTY_ROUNDS rounds the most flags wins (coins break ties).
 *
 * Squad points: SUMMIT_PARTY_POINTS by place (25,15,10,5) go to the Summit Party Points tab and
 * wait for a yes in Admin › System › Waiting for a yes, like every Clash point. 0 = no points.
 *
 * Staff run it from the screen: sign in (the admin password, or "Sign in with my school account"
 * on the staff link), then Start a game, Skip a turn (a squad not there), or End the game.
 *
 * The live game is kept in Script Properties (and the cache), not the Sheet, so a roll never
 * touches the Clash Sheet (whose edits reset the public site's cache). The Summit Games tab gets
 * one row when a game starts and its result when it ends.
 *
 * Switches: SUMMIT_PARTY (OFF until you turn it on: the pages say it is closed and every call is
 * refused), SUMMIT_PARTY_POINTS (0 = no squad points).
 */

const SP_GAME_HEADERS_ = ['Game', 'Status', 'Started', 'Ended', 'Rounds', 'Started By', 'Result'];
const SP_POINT_HEADERS_ = ['Game', 'Ended', 'Result', 'BLACK', 'GOLD', 'GREY', 'PURPLE', 'Approved', 'Sent'];
const SP_PROP_ = 'summit-party-game';
const SP_SPACES_ = { start: 'Base Camp', trail: 'Trail (+3 coins)', rock: 'Rockslide (-3 coins)', cairn: 'Cairn (draw a trail card)', spring: 'Spring (+2 on your next roll)' };

// The board: a loop of 20 spaces around the mountain, two forks with a shortcut each.
// x, y are percent of the board. next = where you can go from here (two = a fork).
function spBoard_() {
  const ring = ['start', 'trail', 'cairn', 'trail', 'trail', 'rock', 'trail', 'spring', 'trail', 'cairn',
    'trail', 'rock', 'trail', 'trail', 'cairn', 'trail', 'spring', 'trail', 'rock', 'trail'];
  const nodes = ring.map(function (type, i) {
    const a = Math.PI / 2 + (i / ring.length) * Math.PI * 2; // base camp at the bottom, going clockwise
    return { id: i, type: type, x: Math.round((50 + 42 * Math.cos(a)) * 10) / 10, y: Math.round((52 + 38 * Math.sin(a)) * 10) / 10, next: [(i + 1) % ring.length] };
  });
  // Fork 1 at space 4: Ridge Shortcut (20, 21) straight to 8, steep and rocky.
  nodes[4].next = [5, 20];
  nodes.push({ id: 20, type: 'rock', x: 31, y: 44, next: [21], path: 'Ridge Shortcut' });
  nodes.push({ id: 21, type: 'cairn', x: 37, y: 30, next: [8], path: 'Ridge Shortcut' });
  // Fork 2 at space 13: Lake Shortcut (22, 23) straight to 17.
  nodes[13].next = [14, 22];
  nodes.push({ id: 22, type: 'cairn', x: 66, y: 56, next: [23], path: 'Lake Shortcut' });
  nodes.push({ id: 23, type: 'rock', x: 60, y: 70, next: [17], path: 'Lake Shortcut' });
  return nodes;
}

const SP_CARDS_ = [
  { text: 'Trail mix stash! +5 coins.', fn: function (g, s) { g.squads[s].coins += 5; } },
  { text: 'Tailwind: +2 on your next roll.', fn: function (g, s) { g.squads[s].boost = 2; } },
  { text: 'Summit snacks for everyone: every squad +2 coins.', fn: function (g) { SQUAD_KEYS_.forEach(function (k) { if (g.squads[k]) g.squads[k].coins += 2; }); } },
  { text: 'Lost the map: -2 coins.', fn: function (g, s) { g.squads[s].coins = Math.max(0, g.squads[s].coins - 2); } },
  { text: 'Rope team: the leader shares 3 coins with you.', fn: function (g, s) {
    const lead = spPlaces_(g)[0];
    if (lead && lead !== s) { const n = Math.min(3, g.squads[lead].coins); g.squads[lead].coins -= n; g.squads[s].coins += n; }
    else g.squads[s].coins += 2;
  } },
];

function spOn_() { return String(getSetting_('SUMMIT_PARTY') || 'OFF').toUpperCase() === 'ON'; }
function spRequireOn_() { if (!spOn_()) throw new Error('Summit Party is closed right now.'); }
function spSetting_(k, d) { const v = Number(getSetting_(k)); return String(getSetting_(k) || '').trim() !== '' && !isNaN(v) ? v : d; }

function spLoad_() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get(SP_PROP_);
  if (hit) { try { return JSON.parse(hit); } catch (e) { /* read the property */ } }
  const raw = PropertiesService.getScriptProperties().getProperty(SP_PROP_);
  if (!raw) return null;
  try { cache.put(SP_PROP_, raw, 600); return JSON.parse(raw); } catch (e) { return null; }
}
function spSave_(g) {
  g.seq = (Number(g.seq) || 0) + 1;
  const raw = JSON.stringify(g);
  PropertiesService.getScriptProperties().setProperty(SP_PROP_, raw);
  CacheService.getScriptCache().put(SP_PROP_, raw, 600);
  return g;
}
function spWithLock_(fn) {
  const lock = LockService.getScriptLock(); lock.waitLock(15000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function spPlaces_(g) {
  return Object.keys(g.squads).sort(function (a, b) {
    return (g.squads[b].flags - g.squads[a].flags) || (g.squads[b].coins - g.squads[a].coins) || g.order.indexOf(a) - g.order.indexOf(b);
  });
}
function spTitle_(s) { return s.charAt(0) + s.slice(1).toLowerCase(); }
function spLog_(g, text) { g.log = [text].concat(g.log || []).slice(0, 10); g.last.events.push(text); }

function spMoveFlag_(g, board) {
  const spots = board.filter(function (n) { return n.type === 'trail' && n.id !== g.flagAt; });
  g.flagAt = spots[Math.floor(Math.random() * spots.length)].id;
}

// Walks `steps` spaces for squad s. Stops at a fork (pending choice) or the end.
function spWalk_(g, board, s, steps) {
  const me = g.squads[s];
  while (steps > 0) {
    const here = board[me.pos];
    if (here.next.length > 1) { g.pending = { squad: s, left: steps, at: here.id, options: here.next.map(function (id) { return { id: id, name: board[id].path || 'Main Trail' }; }) }; return; }
    me.pos = here.next[0];
    g.last.path.push(me.pos);
    steps--;
    if (me.pos === g.flagAt) spTryFlag_(g, board, s);
    if (me.pos === 0 && steps > 0) { me.coins += 2; spLog_(g, spTitle_(s) + ' passed Base Camp: +2 coins.'); }
  }
  spLand_(g, board, s);
}

function spTryFlag_(g, board, s) {
  const cost = spSetting_('SUMMIT_PARTY_FLAG_COST', 10), me = g.squads[s];
  if (me.coins >= cost) {
    me.coins -= cost; me.flags++;
    spLog_(g, '🚩 ' + spTitle_(s) + ' planted a Summit Flag!');
    g.last.flag = s;
    spMoveFlag_(g, board);
  } else spLog_(g, spTitle_(s) + ' reached the Summit Flag but needs ' + cost + ' coins.');
}

function spLand_(g, board, s) {
  const me = g.squads[s], n = board[me.pos];
  if (n.type === 'trail') { me.coins += 3; spLog_(g, spTitle_(s) + ' landed on a trail: +3 coins.'); }
  else if (n.type === 'rock') { const lost = Math.min(3, me.coins); me.coins -= lost; spLog_(g, spTitle_(s) + ' hit a rockslide: -' + lost + ' coins.'); }
  else if (n.type === 'spring') { me.boost = 2; spLog_(g, spTitle_(s) + ' found a spring: +2 on the next roll.'); }
  else if (n.type === 'cairn') { const c = SP_CARDS_[Math.floor(Math.random() * SP_CARDS_.length)]; c.fn(g, s); g.last.card = c.text; spLog_(g, spTitle_(s) + ' drew a trail card: ' + c.text); }
  spNextTurn_(g, board);
}

function spNextTurn_(g, board) {
  g.pending = null;
  g.turn++;
  if (g.turn >= g.order.length) { g.turn = 0; g.round++; }
  if (g.round > g.rounds) spFinish_(g, 'ended');
}

function spFinish_(g, status) {
  g.status = status;
  g.endedAt = new Date().toISOString();
  g.places = spPlaces_(g);
  const result = g.places.map(function (s, i) { return (i + 1) + '. ' + spTitle_(s) + ' (' + g.squads[s].flags + ' flags, ' + g.squads[s].coins + ' coins)'; }).join(' · ');
  try {
    const sh = spGamesTab_(), r = sheetRows_('Summit Games').filter(function (x) { return String(x['Game']) === g.id; })[0];
    if (r) sh.getRange(r._row, 2, 1, 6).setValues([[status === 'ended' ? 'Finished' : 'Stopped', g.startedAt.slice(0, 10), todayYmd_(), g.rounds, String(r['Started By'] || ''), result]]);
  } catch (e) { console.warn('Summit Games: ' + e); }
  if (status !== 'ended') return;
  const pts = String(getSetting_('SUMMIT_PARTY_POINTS') || '25,15,10,5').split(',').map(function (x) { return Math.max(0, Math.round(Number(x) || 0)); });
  if (!pts.some(Boolean)) return;
  const row = { BLACK: '', GOLD: '', GREY: '', PURPLE: '' };
  g.places.forEach(function (s, i) { row[s] = pts[i] || ''; });
  const sh = spPointsTab_();
  sh.appendRow([g.id, todayYmd_(), result, row.BLACK, row.GOLD, row.GREY, row.PURPLE, '', '']);
}

function spGamesTab_() { return getSheetSafe_('Summit Games') || spMakeTab_('Summit Games', SP_GAME_HEADERS_); }
function spPointsTab_() { return getSheetSafe_('Summit Party Points') || spMakeTab_('Summit Party Points', SP_POINT_HEADERS_); }
function spMakeTab_(name, headers) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().insertSheet(name);
  sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#490e6f').setFontColor('#ffffff');
  sh.setFrozenRows(1);
  return sh;
}
function ensureSummitPartyTabs_() { spGamesTab_(); spPointsTab_(); }

// What the screen and the Chromebooks see. sinceSeq = the last one the page has: { same: true } if nothing changed.
function spView_(g) {
  return { on: true, board: spBoard_(), spaces: SP_SPACES_, flagCost: spSetting_('SUMMIT_PARTY_FLAG_COST', 10), game: g };
}

/** Public: the live game for the screen and the controllers. */
function clashSummitState(sinceSeq) {
  if (!spOn_()) return JSON.stringify({ on: false });
  const g = spLoad_();
  if (g && Number(sinceSeq) && Number(sinceSeq) === g.seq) return JSON.stringify({ on: true, same: true, seq: g.seq });
  return JSON.stringify(spView_(g));
}

/** Staff: start a new game (an unfinished one is stopped first). */
function clashSummitStart(token, rounds) {
  const who = requireAdmin_(token);
  spRequireOn_();
  return spWithLock_(function () {
    const board = spBoard_(), old = spLoad_();
    if (old && old.status === 'playing') spFinish_(old, 'stopped');
    const order = SQUAD_KEYS_.slice().sort(function () { return Math.random() - 0.5; });
    const g = { id: 'SP-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmm'), status: 'playing', startedAt: new Date().toISOString(),
      rounds: Math.max(1, Math.min(30, Math.round(Number(rounds) || spSetting_('SUMMIT_PARTY_ROUNDS', 8)))), round: 1, turn: 0, order: order,
      squads: {}, flagAt: 0, pending: null, log: [], seq: 0, last: { squad: '', roll: 0, path: [], events: [] } };
    order.forEach(function (s) { g.squads[s] = { pos: 0, coins: 5, flags: 0, boost: 0 }; });
    spMoveFlag_(g, board);
    spLog_(g, 'A new Summit Party! ' + spTitle_(order[0]) + ' goes first.');
    spGamesTab_().appendRow([g.id, 'Playing', todayYmd_(), '', g.rounds, String(who && (who.name || who.email) || 'staff'), '']);
    return JSON.stringify(spView_(spSave_(g)));
  });
}

/** Staff: skip the squad whose turn it is (nobody at their Chromebook). */
function clashSummitSkip(token) {
  requireAdmin_(token);
  spRequireOn_();
  return spWithLock_(function () {
    const g = spLoad_();
    if (!g || g.status !== 'playing') throw new Error('No game is being played.');
    const s = g.order[g.turn];
    g.last = { squad: s, roll: 0, path: [], events: [] };
    spLog_(g, spTitle_(s) + ' skipped a turn.');
    spNextTurn_(g, spBoard_());
    return JSON.stringify(spView_(spSave_(g)));
  });
}

/** Staff: end the game now. finish = true scores it (points wait for approval); false just stops it. */
function clashSummitEnd(token, finish) {
  requireAdmin_(token);
  return spWithLock_(function () {
    const g = spLoad_();
    if (!g || g.status !== 'playing') throw new Error('No game is being played.');
    g.last = { squad: '', roll: 0, path: [], events: [] };
    spFinish_(g, finish ? 'ended' : 'stopped');
    spLog_(g, finish ? 'The game was ended early. Final results are in.' : 'The game was stopped. No points.');
    return JSON.stringify(spView_(spSave_(g)));
  });
}

// A student at a squad's Chromebook: their 800# must be on the Student Board.
function spPlayer_(id) {
  spRequireOn_();
  const cache = CacheService.getScriptCache(), used = Number(cache.get('sp-calls')) || 0;
  if (used >= 2000) throw new Error('Summit Party is busy. Try again in a minute.');
  cache.put('sp-calls', String(used + 1), 600);
  const sid = id800_(id);
  if (sid.length < 5) throw new Error('Type your 800 number.');
  const me = boardRows_().filter(function (r) { return id800_(r['800#']) === sid; })[0];
  if (!me) throw new Error('That 800 number isn\'t on the list. Check it, or ask your teacher.');
  const squad = squadOf_(me['Squad']);
  if (SQUAD_KEYS_.indexOf(squad) === -1) throw new Error('You are not on a squad yet. Ask your teacher.');
  return { id: sid, squad: squad };
}

/** Controller: who am I (the page remembers the 800# on that Chromebook). */
function clashSummitJoin(id) {
  const p = spPlayer_(id);
  return JSON.stringify({ squad: p.squad, state: JSON.parse(clashSummitState(0)) });
}

/** Controller: roll for your squad (only on its turn). */
function clashSummitRoll(id) {
  const p = spPlayer_(id);
  return spWithLock_(function () {
    const g = spLoad_(), board = spBoard_();
    if (!g || g.status !== 'playing') throw new Error('No game right now. Watch the big screen.');
    if (g.order[g.turn] !== p.squad) throw new Error('It is ' + spTitle_(g.order[g.turn]) + '\'s turn.');
    if (g.pending) throw new Error('Pick a path first.');
    const me = g.squads[p.squad], die = 1 + Math.floor(Math.random() * 6), roll = die + (me.boost || 0);
    g.last = { squad: p.squad, roll: roll, die: die, boost: me.boost || 0, path: [], events: [], from: me.pos };
    me.boost = 0;
    spLog_(g, spTitle_(p.squad) + ' rolled ' + die + (g.last.boost ? ' +' + g.last.boost : '') + '.');
    spWalk_(g, board, p.squad, roll);
    return JSON.stringify(spView_(spSave_(g)));
  });
}

/** Controller: at a fork, pick the path (one of the options' ids). */
function clashSummitChoose(id, nodeId) {
  const p = spPlayer_(id);
  return spWithLock_(function () {
    const g = spLoad_(), board = spBoard_();
    if (!g || g.status !== 'playing' || !g.pending) throw new Error('Nothing to choose right now.');
    if (g.pending.squad !== p.squad) throw new Error('It is ' + spTitle_(g.pending.squad) + '\'s choice.');
    const pick = Number(nodeId), opt = g.pending.options.filter(function (o) { return o.id === pick; })[0];
    if (!opt) throw new Error('Pick one of the paths.');
    const me = g.squads[p.squad], left = g.pending.left - 1;
    g.pending = null;
    me.pos = pick;
    g.last.path.push(pick);
    spLog_(g, spTitle_(p.squad) + ' took the ' + opt.name + '.');
    if (me.pos === g.flagAt) spTryFlag_(g, board, p.squad);
    if (left > 0) spWalk_(g, board, p.squad, left); else spLand_(g, board, p.squad);
    return JSON.stringify(spView_(spSave_(g)));
  });
}

// ── Approvals (Admin › System › Waiting for a yes) ──
function summitPointsWaiting_() {
  return getSheetSafe_('Summit Party Points') ? sheetRows_('Summit Party Points').filter(function (r) { return !ticked_(r['Approved']) && !declined_(r['Approved']); }) : [];
}
function decideSummitPoints_(row, yes) {
  const sh = spPointsTab_(), r = sheetRows_('Summit Party Points').filter(function (x) { return x._row === Number(row); })[0];
  if (!r) throw new Error('That Summit Party row moved. Refresh and try again.');
  sh.getRange(r._row, SP_POINT_HEADERS_.indexOf('Approved') + 1).setValue(yes ? true : 'Declined');
  if (!yes || String(r['Sent'] || '').trim()) return;
  if (String(getSetting_('SUMMIT_PARTY_POINTS') || '25,15,10,5').split(',').every(function (x) { return !(Number(x) > 0); })) return; // switched off since
  getSheet_(TABS.EVENTS).appendRow(['', 'Summit Party', 'summit_party', todayYmd_(), String(r['Game']) + ' (approved)', '',
    Number(r['BLACK']) || '', Number(r['GOLD']) || '', Number(r['GREY']) || '', Number(r['PURPLE']) || '']);
  sh.getRange(r._row, SP_POINT_HEADERS_.indexOf('Sent') + 1).setValue(todayYmd_());
}

function serveSummitParty_() {
  return HtmlService.createHtmlOutputFromFile('SummitParty').setTitle('Summit Party · Clash of Classes')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
function serveSummitPlay_() {
  return HtmlService.createHtmlOutputFromFile('SummitPlay').setTitle('Summit Party · My squad')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// file-stamp (tools/stamp-files.js): lets Admin › System tell whether this file is the newest one
var FILE_STAMP_63_SummitParty = '143732deb0';
