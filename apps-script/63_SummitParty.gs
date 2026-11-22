/**
 * Clash of Classes -- Apps Script backend
 * File 63: Summit Party, the squad party board game.
 *
 * HOW A GAME GOES (a party board game, in Clash's own art and names):
 *   - The board is on one big screen: clashofclasses.org/?page=summit (SummitParty.html). Each squad
 *     plays from one Chromebook: ?page=summit-play (SummitPlay.html), game code + an 800#.
 *   - On its turn a squad may use one item, then hits the dice block; its creature hops space by
 *     space. Blue spaces +3 coins, red -3, "?" = one of Cliff's trail events, item spaces hand out an
 *     item, Opponent spaces bring out Shady, the Tardy Troll or the Litter Beast, the Avalanche shakes
 *     up two squads. Passing the Camp Store lets a squad buy an item. Passing the Summit Flag with
 *     SUMMIT_PARTY_FLAG_COST coins (20) plants a flag, and the flag moves.
 *   - After every squad has moved, everyone plays a minigame: most rounds a 20-second game on all
 *     four Chromebooks at once (Tap Rally, Quick Draw, Trail Memory), scored automatically; every
 *     third round a live Move, Make or Talk challenge (Summit Live Challenges tab) and staff tap the
 *     winner. Minigame coins: 10 / 5 / 3 (live: winner 10, the rest 2).
 *   - The FINAL CLIMB (last 3 rounds): the squad in last place gets 10 coins, blue and red spaces count
 *     double. At the end, bonus flags: Minigame Champ, Coin Keeper (most coins at once), Trail Explorer
 *     (most "?" spaces). Most flags wins; coins break ties. Cliff narrates everything.
 *
 * GAME CODES: several games at once; each board shows its code. AUTO-ROLL after
 * SUMMIT_PARTY_AUTO_ROLL seconds (30) keeps a drop-in game moving. Any staff member on Staff Access
 * (any role) runs a game from the board.
 *
 * Squad points: SUMMIT_PARTY_POINTS by place (25,15,10,5) wait for a yes in Admin › System, like
 * every Clash point. The live game lives in Script Properties, so a roll never touches the Sheet.
 *
 * Switches: SUMMIT_PARTY (OFF until turned on), SUMMIT_PARTY_POINTS (0 = no points),
 * SUMMIT_PARTY_LIVE (live challenges every third round), SUMMIT_PARTY_ITEMS (items and the Camp
 * Store), SUMMIT_PARTY_BONUS (bonus flags), SUMMIT_PARTY_AUTO_ROLL (0 = off), SUMMIT_PARTY_MINIGAMES
 * (which Chromebook minigames are in the rotation).
 */

const SP_GAME_HEADERS_ = ['Game', 'Status', 'Started', 'Ended', 'Rounds', 'Started By', 'Result'];
const SP_POINT_HEADERS_ = ['Game', 'Ended', 'Result', 'BLACK', 'GOLD', 'GREY', 'PURPLE', 'Approved', 'Sent'];
const SP_LIVE_HEADERS_ = ['Show', 'Kind', 'Challenge', 'How to play'];
const SP_SPACES_ = { start: 'Base Camp (pass it: +3)', blue: 'Blue space (+3 coins)', red: 'Red space (-3 coins)', event: '? Trail event',
  item: 'Item space', villain: 'Opponent space', avalanche: 'Avalanche (shake-up)', shop: 'Camp Store (buy an item)', duel: 'Duel space (challenge a squad)' };
const SP_ITEMS_ = {
  double: { name: 'Double Dice', price: 10, text: 'Roll two dice this turn.' },
  compass: { name: 'Golden Compass', price: 8, text: 'Pick exactly how far you move (1 to 6).' },
  paws: { name: 'Sticky Paws', price: 12, text: 'Take up to 10 coins from the coin leader.' },
  zipline: { name: 'Zipline', price: 6, text: 'Swap places with a random squad.' },
};

// THE BOARD: a 32-space loop around the mountain with two forks (each a shortcut). The game uses
// id, type, next and path; x and y (percent) place the spaces for the chosen look:
//   trail  = the mountain island (16:10 board), round spaces along a winding trail
//   frost  = Frostpeak at Night: a snowy summit under the stars, the trail in a different shape
const SP_RING_ = ['start', 'blue', 'blue', 'event', 'blue', 'red', 'duel', 'item', 'blue', 'blue', 'villain', 'blue', 'event', 'shop', 'blue', 'red',
  'blue', 'avalanche', 'blue', 'item', 'blue', 'red', 'blue', 'event', 'duel', 'villain', 'blue', 'item', 'red', 'blue', 'event', 'red'];
const SP_SHORTCUTS_ = [
  { from: 8, to: 13, types: ['red', 'event', 'blue'], name: 'Ridge Shortcut' },
  { from: 20, to: 25, types: ['red', 'item'], name: 'Lake Shortcut' },
];
function spRingXY_(layout, i, n) {
  const t = Math.PI / 2 + (i / n) * Math.PI * 2; // Base Camp at the bottom, going clockwise up the left side
  if (layout === 'frost') { // a switchback trail that climbs high toward the summit at the top
    const fx = 50 + 41 * Math.cos(t) + 4 * Math.sin(2 * t), fy = 53 + 36 * Math.sin(t) + 6 * Math.cos(3 * t);
    return [Math.round(fx * 10) / 10, Math.round(fy * 10) / 10];
  }
  const x = 50 + 42 * Math.cos(t) + 3.5 * Math.sin(3 * t), y = 51 + 38 * Math.sin(t) + 3 * Math.cos(2 * t);
  return [Math.round(x * 10) / 10, Math.round(y * 10) / 10];
}
function spLayoutOf_(layout) {
  const ok = function (l) { return l === 'trail' || l === 'frost'; };
  if (ok(layout)) return layout;
  const d = String(getSetting_('SUMMIT_PARTY_BOARD') || '').toLowerCase();
  return ok(d) ? d : 'trail';
}
function spBoard_(layout) {
  layout = spLayoutOf_(layout);
  const n = SP_RING_.length;
  const nodes = SP_RING_.map(function (type, i) { const p = spRingXY_(layout, i, n); return { id: i, type: type, x: p[0], y: p[1], next: [(i + 1) % n] }; });
  SP_SHORTCUTS_.forEach(function (sc) {
    const a = nodes[sc.from], b = nodes[sc.to], first = nodes.length;
    sc.types.forEach(function (type, k) {
      const f = (k + 1) / (sc.types.length + 1);
      // A straight line between the two ends, pulled toward the middle of the board.
      let x = a.x + (b.x - a.x) * f, y = a.y + (b.y - a.y) * f;
      const pull = layout === 'frost' ? 0.36 : 0.3;
      x += (50 - x) * pull * Math.sin(Math.PI * f); y += (50 - y) * pull * Math.sin(Math.PI * f);
      nodes.push({ id: first + k, type: type, x: Math.round(x * 10) / 10, y: Math.round(y * 10) / 10, next: [k === sc.types.length - 1 ? sc.to : first + k + 1], path: sc.name });
    });
    a.next = [a.next[0], first];
  });
  return nodes;
}

// Art: each squad's evolved My Climb creature, the clear shield, Cliff, and the Opponents.
function spArt_() {
  const on = String(getSetting_('MASCOT_ART') || 'ON').toUpperCase() !== 'OFF', pieces = {};
  SQUAD_KEYS_.forEach(function (s) { pieces[s] = on ? ART_BASE_ + 'creatures/' + s.toLowerCase() + '-4.jpg' : ''; });
  const shield = String(getSetting_('SHIELD_LOGO_URL') || '').trim();
  // The board uses the clear-background shield unless Settings SHIELD_LOGO_URL names another picture.
  const custom = /^https:\/\//i.test(shield) && !/shield-2026-27\.png$/.test(shield);
  return { pieces: pieces, shield: custom ? shield : ART_BASE_ + 'shield-2026-27-clear.png', cliff: ART_BASE_ + 'cliff.png',
    villains: { shady: ART_BASE_ + 'monster-shady.jpg', troll: ART_BASE_ + 'monster-tardy-troll.jpg', litter: ART_BASE_ + 'monster-litter-beast.jpg' } };
}

// The live Move, Make, Talk challenges (from the December files). Editable on the Summit Live Challenges tab.
const SP_LIVE_SEED_ = [
  [true, 'Move', 'Boom, Clap, Snatch', 'Partners face each other with a marker between them. Follow the caller: boom, clap... snatch! Snatch too soon and you are out. Last squad standing wins.'],
  [true, 'Move', 'Four Corners', 'Everyone moves to a corner before "Freeze!". The caller picks a corner; anyone there sits down. Last squad with players standing wins.'],
  [true, 'Talk', 'Chain Gauntlet', 'One word each, linked to the word before it (snow, ball, park...). A pause or a repeat breaks the chain. Longest chain in 60 seconds wins.'],
  [true, 'Make', 'Gratitude Sprint', 'In 90 seconds, each squad writes as many real thank-you notes as it can to people in the building. Most notes wins.'],
  [true, 'Move', 'Silent Line-Up', 'Line up by birthday with no talking. Fastest correct line wins.'],
  [true, 'Talk', 'Two-Word Story', 'Tell a story two words at a time around the squad. The room votes for the best story.'],
];
function spLiveChallenges_() {
  const rows = getSheetSafe_('Summit Live Challenges') ? sheetRows_('Summit Live Challenges').filter(function (r) { return ticked_(r['Show']) && String(r['Challenge'] || '').trim(); }) : [];
  const list = rows.length ? rows.map(function (r) { return [String(r['Kind'] || 'Move'), String(r['Challenge']), String(r['How to play'] || '')]; })
    : SP_LIVE_SEED_.map(function (r) { return [r[1], r[2], r[3]]; });
  return list.map(function (r) { return { kind: r[0], name: r[1], how: r[2] }; });
}

// Chromebook minigames: 20 seconds, every squad at once, scored on the page. Five come from Ridgeline's arcade
// (Updraft, Summit Bounce, Ridge Runner, Switchback Slalom, Cairn Stack). Every squad gets the same course (mg.seed).
// SUMMIT_PARTY_MINIGAMES picks which ones are in the rotation (ALL, or a list of keys).
const SP_MINIGAMES_ = {
  tap: { name: 'Tap Rally', how: 'Tap the button as fast as you can for 20 seconds. Most taps wins.', better: 'high', max: 260, unit: 'taps', icon: 'ph-hand-tap' },
  react: { name: 'Quick Draw', how: 'Wait for the gold flash, then tap. Best of three. Fastest time wins.', better: 'low', max: 3000, unit: 's', icon: 'ph-lightning' },
  memory: { name: 'Trail Memory', how: 'Watch the trail signs light up, then tap them back in order. Longest trail wins.', better: 'high', max: 15, unit: 'signs', icon: 'ph-signpost' },
  updraft: { name: 'Updraft', how: 'Tap or press space to rise. Glide through the gaps in the rock gates and grab gold flags. Most gates wins.', better: 'high', max: 60, unit: 'gates', icon: 'ph-wind', from: 'Ridgeline' },
  bounce: { name: 'Summit Bounce', how: 'You bounce on your own. Steer onto the ledges with the arrow keys (or hold a side) and climb. Highest climb wins.', better: 'high', max: 1000, unit: 'm', icon: 'ph-arrow-fat-lines-up', from: 'Ridgeline' },
  runner: { name: 'Ridge Runner', how: 'Run the ridge. Grab only the GOLD summit flags (purple ones are look-alikes) and jump the rockfall. Most gold flags wins.', better: 'high', max: 40, unit: 'flags', icon: 'ph-person-simple-run', from: 'Ridgeline' },
  slalom: { name: 'Switchback Slalom', how: 'Ski down the mountain. Steer through the gold gates with the arrow keys (or hold a side). Most gates wins.', better: 'high', max: 60, unit: 'gates', icon: 'ph-snowflake', from: 'Ridgeline' },
  stack: { name: 'Cairn Stack', how: 'A stone slides back and forth. Tap or press space to drop it on the cairn. Only the part that lands stays. Tallest cairn wins.', better: 'high', max: 40, unit: 'stones', icon: 'ph-stack', from: 'Ridgeline' },
  catch: { name: 'Camp Snack Catch', how: 'Trail snacks fall from the pines. Steer your pack to catch them and dodge the rocks. Most snacks wins.', better: 'high', max: 150, unit: 'snacks', icon: 'ph-basket' },
  sketch: { name: 'Trail Sketch', how: 'Draw the word on your screen. The AI watches and guesses as you draw. The sooner it guesses right, the more points.', better: 'high', max: 200, unit: 'pts', icon: 'ph-paint-brush', secs: 45, ai: true },
  escape: { name: 'Avalanche Escape', how: 'Run from the avalanche! Jump the rocks (space, up, or Jump) and duck the low branches (down, or hold Duck). One crash ends your run. Farthest run wins.', better: 'high', max: 6000, unit: 'm', icon: 'ph-sneaker-move', secs: 60 },
  rockslide: { name: 'Rockslide', how: 'Stones fall. Slide and turn them to fill whole rows. Clear 2 or more rows at once to drop junk rows on the other squads. Most rows in 90 seconds wins.', better: 'high', max: 300, unit: 'rows', icon: 'ph-squares-four', secs: 90 },
  chase: { name: 'Trail Chase', how: 'On the big screen! One squad is the hiker: gobble the trail markers and dodge the ghosts. Every other squad steers a ghost to catch the hiker. Grab a lantern and the ghosts run!', better: 'high', max: 400, unit: 'markers', icon: 'ph-ghost', secs: 60, board: true },
  dodge: { name: 'Avalanche Dodge', how: 'Snowballs roll down the slope. Steer to dodge them and grab gold flags. A hit drops a flag. Most flags wins.', better: 'high', max: 60, unit: 'flags', icon: 'ph-mountains' },
};

// Everyone vs Shady: the average score every squad must reach together (the low-is-better game: at or under).
const SP_COOP_TARGET_ = { escape: 120, rockslide: 6, tap: 90, react: 450, memory: 5, updraft: 7, bounce: 110, runner: 5, slalom: 11, stack: 7, catch: 22, sketch: 100, dodge: 5 };
const SP_MODES_ = ['ffa', '2v2', 'ffa', '1v3', 'ffa', 'coop'];
function spSwitch_(k) { return String(getSetting_(k) || 'ON').toUpperCase() !== 'OFF'; }
/** Trail Sketch needs its switch on and a Claude key (Script Property CLAUDE_API_KEY, the same one Follow-Up uses). */
function spDrawOn_() { return spSwitch_('SUMMIT_PARTY_DRAW') && !!PropertiesService.getScriptProperties().getProperty('CLAUDE_API_KEY'); }

/** The minigames in the rotation (Settings › SUMMIT_PARTY_MINIGAMES: ALL, or keys like updraft, bounce, tap). */
function spMinigameKeys_() {
  const all = Object.keys(SP_MINIGAMES_).filter(function (k) { return (k !== 'sketch' || spDrawOn_()) && (k !== 'chase' || spSwitch_('SUMMIT_PARTY_CHASE')); }), raw = String(getSetting_('SUMMIT_PARTY_MINIGAMES') || 'ALL').trim().toLowerCase();
  if (!raw || raw === 'all') return all;
  const pick = raw.split(/[\s,]+/).filter(function (k) { return SP_MINIGAMES_[k]; });
  return pick.length ? all.filter(function (k) { return pick.indexOf(k) >= 0; }) : all;
}

function spOn_() { return String(getSetting_('SUMMIT_PARTY') || 'OFF').toUpperCase() === 'ON'; }
function spRequireOn_() { if (!spOn_()) throw new Error('Summit Party is closed right now.'); }
// A preview demo game (65_Preview.gs) runs even while SUMMIT_PARTY is OFF.
function spRequireOnFor_(code) { if (!spOn_() && !(spCode_(code) && (spLoad_(code) || {}).demo)) throw new Error('Summit Party is closed right now.'); }
function spSetting_(k, d) { const v = Number(getSetting_(k)); return String(getSetting_(k) || '').trim() !== '' && !isNaN(v) ? v : d; }

// ── GAMES BY CODE ──────────────────────────────────────────────────
// Each game has a 4-letter code (shown on its board). Several can run at once (classrooms, the
// cafeteria, an assembly). Each game lives in its own Script Property; 'sp-games' lists them.
const SP_LETTERS_ = 'ACDEFGHJKMNPQRTUVWXY'; // no B/I/L/O/S/Z: easy to read off a screen
const SP_MAX_GAMES_ = 20;
function spCode_(code) { return String(code || '').toUpperCase().replace(/[^A-Z]/g, '').slice(0, 4); }
function spKey_(code) { return 'sp-game-' + code; }
function spIndex_() { try { return JSON.parse(PropertiesService.getScriptProperties().getProperty('sp-games') || '[]'); } catch (e) { return []; } }
function spSetIndex_(list) { PropertiesService.getScriptProperties().setProperty('sp-games', JSON.stringify(list)); }

// Games older than 12 hours drop off the list (their result is already on the Summit Games tab).
function spTidy_() {
  const now = Date.now(), props = PropertiesService.getScriptProperties(), keep = [];
  spIndex_().forEach(function (x) {
    if (now - Number(x.started || 0) > (x.demo ? 2 : 12) * 3600000) { props.deleteProperty(spKey_(x.code)); CacheService.getScriptCache().remove(spKey_(x.code)); }
    else keep.push(x);
  });
  spSetIndex_(keep);
  return keep;
}

function spNewCode_(taken) {
  for (let i = 0; i < 50; i++) {
    let c = '';
    for (let k = 0; k < 4; k++) c += SP_LETTERS_.charAt(Math.floor(Math.random() * SP_LETTERS_.length));
    if (taken.indexOf(c) === -1) return c;
  }
  throw new Error('Could not make a game code. Try again.');
}

// The game for a code. No code: the one game being played, if there is exactly one (the cafeteria).
function spLoad_(code) {
  code = spCode_(code);
  if (!code) {
    const live = spIndex_().filter(function (x) { return x.status === 'playing' && !x.demo; });
    if (live.length !== 1) return null;
    code = live[0].code;
  }
  const cache = CacheService.getScriptCache(), hit = cache.get(spKey_(code));
  if (hit) { try { return JSON.parse(hit); } catch (e) { /* read the property */ } }
  const raw = PropertiesService.getScriptProperties().getProperty(spKey_(code));
  if (!raw) return null;
  try { cache.put(spKey_(code), raw, 600); return JSON.parse(raw); } catch (e) { return null; }
}
function spSave_(g) {
  g.seq = (Number(g.seq) || 0) + 1;
  const raw = JSON.stringify(g);
  PropertiesService.getScriptProperties().setProperty(spKey_(g.code), raw);
  CacheService.getScriptCache().put(spKey_(g.code), raw, 600);
  const list = spIndex_(), x = list.filter(function (y) { return y.code === g.code; })[0];
  if (x && x.status !== g.status) { x.status = g.status; spSetIndex_(list); }
  return g;
}
function spWithLock_(fn) {
  const lock = LockService.getScriptLock(); lock.waitLock(15000);
  try { return fn(); } finally { lock.releaseLock(); }
}
function spGame_(code) {
  const g = spLoad_(code);
  if (!g) throw new Error(spCode_(code) ? 'No game with the code ' + spCode_(code) + '. Check the big screen.' : 'Type the game code from the big screen.');
  return g;
}

// ── STAFF WHO RUN A GAME ───────────────────────────────────────────
// Any staff member on the Staff Access tab with Active ticked (any role, not only EDIT_ROLES) may
// start, skip and end games. Admin sessions (password or school account) work too.
function spStaff_(token) {
  const admin = adminSession_(token);
  if (admin) return admin;
  const raw = token && typeof token === 'string' && token.length < 200 ? CacheService.getScriptCache().get('sp-staff-' + token) : null;
  if (raw) {
    let who = null; try { who = JSON.parse(raw); } catch (e) { who = null; }
    if (who && spStaffRecord_(who.email)) return who;
  }
  throw new Error('Please sign in again (staff on the Clash Staff Access list).');
}
function spStaffRecord_(email) {
  email = String(email || '').toLowerCase();
  if (!email) return null;
  const rec = staffAccessMap_()[email];
  if (email === ownerEmail_()) return { email: email, name: (rec && rec.name) || 'Sheet owner', role: 'Admin' };
  if (!staffDomainOk_(email) || !rec || !rec.active) return null;
  return { email: email, name: rec.name || email, role: rec.role || 'Staff' };
}
/** Staff sign-in on the board with the school account (on the staff link). Any Staff Access role. */
function clashSummitSignIn() {
  const email = activeEmail_();
  if (!email) throw new Error('Open the board on the staff link, signed in with your school Google account. Or use the admin password.');
  const who = spStaffRecord_(email);
  if (!who) throw new Error(email + ' is not on the Clash Staff Access list. Ask Mr. Schuck to add you.');
  const token = Utilities.getUuid() + Utilities.getUuid().replace(/-/g, '');
  CacheService.getScriptCache().put('sp-staff-' + token, JSON.stringify(who), 6 * 3600);
  logAdminAction_({ name: who.name, email: who.email, role: who.role, via: 'google' }, 'Summit Party sign-in', '');
  return JSON.stringify({ token: token, name: who.name });
}

// ── THE GAME ───────────────────────────────────────────────────────
function spPlaces_(g) {
  return Object.keys(g.squads).sort(function (a, b) {
    return (g.squads[b].flags - g.squads[a].flags) || (g.squads[b].coins - g.squads[a].coins) || g.order.indexOf(a) - g.order.indexOf(b);
  });
}
function spTitle_(s) { return s.charAt(0) + s.slice(1).toLowerCase(); }
function spLog_(g, text) { g.log = [text].concat(g.log || []).slice(0, 12); if (g.last) g.last.events.push(text); }
function spHost_(g, text) { g.host = { text: text, at: Date.now() }; }
function spCoins_(g, s, n) {
  const me = g.squads[s];
  me.coins = Math.max(0, me.coins + n);
  if (me.coins > (me.peak || 0)) me.peak = me.coins;
}
function spOnFlag_() { return String(getSetting_('SUMMIT_PARTY_ITEMS') || 'ON').toUpperCase() !== 'OFF'; }
function spFlagCost_() { return spSetting_('SUMMIT_PARTY_FLAG_COST', 20); }

function spMoveFlag_(g, board) {
  const spots = board.filter(function (n) { return n.type === 'blue' && !n.path && n.id !== g.flagAt; });
  g.flagAt = spots[Math.floor(Math.random() * spots.length)].id;
}

// Walks `steps` spaces for squad s. Stops at a fork or at the Camp Store (a choice), or lands.
function spWalk_(g, board, s, steps) {
  const me = g.squads[s];
  while (steps > 0) {
    const here = board[me.pos];
    if (here.next.length > 1) {
      g.turnAt = Date.now();
      g.pending = { kind: 'fork', squad: s, left: steps, at: here.id, options: here.next.map(function (id) { return { id: id, name: board[id].path || 'Main Trail' }; }) };
      spHost_(g, spTitle_(s) + ', which way? The ' + board[here.next[1]].path + ' is shorter, but rougher.');
      return;
    }
    me.pos = here.next[0];
    me.moved = (me.moved || 0) + 1;
    g.last.path.push(me.pos);
    steps--;
    if (me.pos === g.flagAt) spTryFlag_(g, board, s);
    if (me.pos === 0 && steps > 0) { spCoins_(g, s, 3); spLog_(g, spTitle_(s) + ' passed Base Camp: +3 coins.'); }
    if (board[me.pos].type === 'shop' && spShopOpen_(g, s)) {
      g.turnAt = Date.now();
      g.pending = { kind: 'shop', squad: s, left: steps, at: me.pos, options: spShopList_(g, s) };
      spHost_(g, 'Welcome to the Camp Store, ' + spTitle_(s) + '! Want anything?');
      return;
    }
  }
  spLand_(g, board, s);
}

function spShopOpen_(g, s) {
  const me = g.squads[s];
  return spOnFlag_() && (me.items || []).length < 2 && Object.keys(SP_ITEMS_).some(function (k) { return SP_ITEMS_[k].price <= me.coins; });
}
function spShopList_(g, s) {
  return Object.keys(SP_ITEMS_).map(function (k) { return { id: k, name: SP_ITEMS_[k].name, price: SP_ITEMS_[k].price, text: SP_ITEMS_[k].text, afford: SP_ITEMS_[k].price <= g.squads[s].coins }; });
}

function spTryFlag_(g, board, s) {
  const cost = spFlagCost_(), me = g.squads[s];
  if (me.coins >= cost) {
    me.coins -= cost; me.flags++;
    spLog_(g, spTitle_(s) + ' planted a Summit Flag!');
    spHost_(g, 'A SUMMIT FLAG for ' + spTitle_(s) + '! Now where will the next one be?');
    g.last.flag = s;
    spMoveFlag_(g, board);
  } else {
    spLog_(g, spTitle_(s) + ' reached the Summit Flag but needs ' + cost + ' coins.');
    spHost_(g, 'So close! A Summit Flag costs ' + cost + ' coins, ' + spTitle_(s) + '.');
  }
}

const SP_EVENTS_ = [
  { text: 'A trail cache! +6 coins.', fn: function (g, board, s) { spCoins_(g, s, 6); } },
  { text: 'A tailwind carries you 3 spaces ahead.', fn: function (g, board, s) { const me = g.squads[s]; for (let i = 0; i < 3; i++) { me.pos = board[me.pos].next[0]; g.last.path.push(me.pos); } } },
  { text: 'The Summit Flag blew away to a new spot!', fn: function (g, board) { spMoveFlag_(g, board); } },
  { text: 'Trail mix for everyone: every squad +3 coins.', fn: function (g) { Object.keys(g.squads).forEach(function (k) { spCoins_(g, k, 3); }); } },
  { text: 'Switchback! You trade places with the squad in first.', fn: function (g, board, s) {
    const lead = spPlaces_(g)[0];
    if (lead === s) { spCoins_(g, s, 4); return; }
    const p = g.squads[s].pos; g.squads[s].pos = g.squads[lead].pos; g.squads[lead].pos = p; g.last.swap = [s, lead];
  } },
];
const SP_VILLAINS_ = [
  { key: 'troll', name: 'The Tardy Troll', text: 'charges a late fee: -6 coins.', fn: function (g, board, s) { spCoins_(g, s, -6); } },
  { key: 'litter', name: 'The Hallway Litter Beast', text: 'scatters 3 coins from every squad.', fn: function (g) { Object.keys(g.squads).forEach(function (k) { spCoins_(g, k, -3); }); } },
  { key: 'shady', name: 'Shady', text: 'swaps your coins with the squad that has the fewest!', fn: function (g, board, s) {
    const poor = Object.keys(g.squads).sort(function (a, b) { return g.squads[a].coins - g.squads[b].coins; })[0];
    if (poor === s) { spCoins_(g, s, -4); return; }
    const c = g.squads[s].coins; g.squads[s].coins = g.squads[poor].coins; spCoins_(g, poor, c - g.squads[poor].coins);
  } },
];

function spLand_(g, board, s) {
  const me = g.squads[s], n = board[me.pos], x = g.final ? 2 : 1;
  if (n.type === 'blue') { spCoins_(g, s, 3 * x); spLog_(g, spTitle_(s) + ' landed on blue: +' + 3 * x + ' coins.'); }
  else if (n.type === 'red') { spCoins_(g, s, -3 * x); spLog_(g, spTitle_(s) + ' landed on red: -' + 3 * x + ' coins.'); spHost_(g, 'Ouch, a red space! -' + 3 * x + ' coins.'); }
  else if (n.type === 'event') {
    const e = SP_EVENTS_[Math.floor(Math.random() * SP_EVENTS_.length)];
    me.events = (me.events || 0) + 1;
    e.fn(g, board, s);
    g.last.event = e.text;
    spLog_(g, spTitle_(s) + ': ' + e.text);
    spHost_(g, 'A ? space! ' + e.text);
  }
  else if (n.type === 'item') {
    if (spOnFlag_() && (me.items || []).length < 2) {
      const keys = Object.keys(SP_ITEMS_), k = keys[Math.floor(Math.random() * keys.length)];
      me.items = (me.items || []).concat([k]);
      g.last.item = k;
      spLog_(g, spTitle_(s) + ' found a ' + SP_ITEMS_[k].name + '.');
      spHost_(g, spTitle_(s) + ' found a ' + SP_ITEMS_[k].name + '! ' + SP_ITEMS_[k].text);
    } else { spCoins_(g, s, 3); spLog_(g, spTitle_(s) + ' found 3 coins.'); }
  }
  else if (n.type === 'villain') {
    const v = SP_VILLAINS_[Math.floor(Math.random() * SP_VILLAINS_.length)];
    v.fn(g, board, s);
    g.last.villain = v.key;
    spLog_(g, v.name + ' ' + v.text);
    spHost_(g, 'Uh oh. ' + v.name + ' ' + v.text);
  }
  else if (n.type === 'duel') {
    if (spSwitch_('SUMMIT_PARTY_DUELS')) {
      g.turnAt = Date.now();
      g.pending = { kind: 'duel', squad: s, options: Object.keys(g.squads).filter(function (k) { return k !== s; }).map(function (k) { return { id: k, coins: g.squads[k].coins }; }),
        bets: [5, 10, 20].filter(function (b) { return b <= me.coins; }) };
      spHost_(g, 'A DUEL space! ' + spTitle_(s) + ', pick a squad to challenge and how many coins to bet.');
      return;
    }
    spCoins_(g, s, 3 * x); spLog_(g, spTitle_(s) + ' landed on a duel space: +' + 3 * x + ' coins.');
  }
  else if (n.type === 'avalanche') {
    const others = Object.keys(g.squads).filter(function (k) { return k !== s; }), other = others[Math.floor(Math.random() * others.length)];
    const pick = Math.floor(Math.random() * 3);
    if (pick === 0) { const c = g.squads[s].coins; g.squads[s].coins = g.squads[other].coins; g.squads[other].coins = c; g.last.avalanche = spTitle_(s) + ' and ' + spTitle_(other) + ' swap coins!'; }
    else if (pick === 1) { const n2 = Math.min(10, g.squads[other].coins); spCoins_(g, other, -n2); spCoins_(g, s, n2); g.last.avalanche = spTitle_(other) + ' gives ' + spTitle_(s) + ' ' + n2 + ' coins!'; }
    else { const n3 = Math.min(10, g.squads[s].coins); spCoins_(g, s, -n3); spCoins_(g, other, n3); g.last.avalanche = spTitle_(s) + ' gives ' + spTitle_(other) + ' ' + n3 + ' coins!'; }
    spLog_(g, 'AVALANCHE! ' + g.last.avalanche);
    spHost_(g, 'AVALANCHE! ' + g.last.avalanche);
  }
  spNextTurn_(g, board);
}

// After the last squad of a round moves, everybody plays a minigame.
function spNextTurn_(g, board) {
  g.pending = null;
  g.turn++;
  g.turnAt = Date.now();
  if (g.turn >= g.order.length) { g.turn = 0; spStartMinigame_(g); }
}

function spStartMinigame_(g) {
  const live = !g.demo && String(getSetting_('SUMMIT_PARTY_LIVE') || 'ON').toUpperCase() !== 'OFF' && g.round % 3 === 0 && g.round < g.rounds + 1;
  g.phase = 'minigame';
  if (live) {
    const list = spLiveChallenges_(), c = list[(Math.floor(g.round / 3) - 1 + (g.seed || 0)) % list.length];
    g.mg = { round: g.round, live: true, kind: c.kind, name: c.name, how: c.how, scores: {} };
    spHost_(g, 'Minigame time! A LIVE ' + c.kind.toUpperCase() + ' challenge: ' + c.name + '. Staff, tap the winning squad.');
  } else {
    const keys = spMinigameKeys_(), k = keys[(g.round - 1 + (g.seed || 0)) % keys.length];
    const mode = k === 'chase' ? 'chase' : spSwitch_('SUMMIT_PARTY_TEAMS') ? SP_MODES_[(g.round - 1 + (g.seed || 0)) % SP_MODES_.length] : 'ffa';
    spNewMg_(g, k, mode, Object.keys(g.squads));
    spHost_(g, 'Minigame time! ' + spModeLine_(g) + g.mg.name + ': ' + g.mg.how);
  }
  spLog_(g, 'Minigame: ' + g.mg.name + '.');
}

// A Chromebook minigame for some squads. Practice ON = up to 25 seconds to practice, starting early once every squad taps Ready.
function spNewMg_(g, k, mode, squads, extra) {
  const m = SP_MINIGAMES_[k], now = Date.now(), wait = k === 'chase' ? 12000 : spSwitch_('SUMMIT_PARTY_PRACTICE') ? 25000 : 9000, secs = m.secs || 20;
  g.mgCount = (g.mgCount || 0) + 1;
  g.mg = { id: g.mgCount, round: g.round, live: false, game: k, name: m.name, how: m.how, mode: mode, squads: squads, secs: secs,
    startAt: now + wait, endAt: now + wait + secs * 1000, scores: {}, ready: {}, seed: Math.floor(Math.random() * 1e6) };
  if (mode === '2v2' || mode === '1v3') {
    const p = spPlaces_(g);
    g.mg.teams = mode === '2v2' ? [[p[0], p[3]], [p[1], p[2]]] : [[p[0]], [p[1], p[2], p[3]]];
  }
  if (mode === 'coop') g.mg.target = SP_COOP_TARGET_[k] || 1;
  if (k === 'sketch') spSketchSetup_(g.mg);
  if (k === 'chase') {
    g.chases = (g.chases || 0) + 1;
    const hiker = g.order[(g.chases - 1) % g.order.length], v = ['shady', 'troll', 'litter'];
    g.mg.hiker = hiker; g.mg.ghosts = {};
    g.order.filter(function (s) { return s !== hiker && squads.indexOf(s) >= 0; }).forEach(function (s, i) { g.mg.ghosts[s] = v[i % v.length]; });
  }
  if (extra) Object.keys(extra).forEach(function (x) { g.mg[x] = extra[x]; });
  if (g.demo) squads.forEach(function (s) { if (s !== g.demo.mine) g.mg.ready[s] = true; }); // the computer squads are always ready
  return g.mg;
}
// Games a duel or the Showdown can use (not the AI or big-screen ones).
function spSoloKeys_() {
  const k = spMinigameKeys_().filter(function (x) { return x !== 'sketch' && x !== 'chase'; });
  return k.length ? k : ['tap'];
}
function spModeLine_(g) {
  const mg = g.mg, t = function (list) { return list.map(spTitle_).join(' + '); };
  if (mg.mode === '2v2') return '2 vs 2! ' + t(mg.teams[0]) + ' against ' + t(mg.teams[1]) + '. ';
  if (mg.mode === '1v3') return '1 vs 3! ' + spTitle_(mg.teams[0][0]) + ' takes on everyone. ';
  if (mg.mode === 'coop') return 'Everyone vs Shady! Work together: every squad averages ' + mg.target + ' or Shady steals coins. ';
  if (mg.mode === 'duel') return 'DUEL! ' + t(mg.squads) + ' for ' + mg.bet + ' coins each. ';
  if (mg.mode === 'chase') return 'TRAIL CHASE! ' + spTitle_(mg.hiker) + ' is the hiker; ' + Object.keys(mg.ghosts).map(spTitle_).join(', ') + ' are the ghosts. ';
  if (mg.mode === 'showdown') return 'SHOWDOWN! ' + t(mg.squads) + ' are tied for first. ';
  return '';
}
// A squad's value: its score, or the worst score if it did not play.
function spVal_(mg, s) { const d = SP_MINIGAMES_[mg.game]; return mg.scores[s] != null ? mg.scores[s] : d.better === 'low' ? d.max : 0; }
function spBetter_(mg, a, b) { return SP_MINIGAMES_[mg.game].better === 'low' ? a < b : a > b; }
function spAvg_(mg, list) { return list.reduce(function (n, s) { return n + spVal_(mg, s); }, 0) / list.length; }

// Minigame coins by place; ties share the better place.
function spResolveMinigame_(g) {
  const mg = g.mg, squads = Object.keys(g.squads);
  if (mg.mode === 'duel') return spResolveDuel_(g);
  if (mg.mode === 'showdown') return spResolveShowdown_(g);
  let prizes = {};
  if (mg.live) {
    squads.forEach(function (s) { prizes[s] = mg.winner === s ? 10 : 2; });
    if (mg.winner) g.squads[mg.winner].mgWins = (g.squads[mg.winner].mgWins || 0) + 1;
  } else if (mg.mode === '2v2' || mg.mode === '1v3') {
    const a = spAvg_(mg, mg.teams[0]), b = spAvg_(mg, mg.teams[1]);
    const win = a === b ? -1 : spBetter_(mg, a, b) ? 0 : 1;
    const pay = mg.mode === '2v2' ? [10, 10] : [15, 5];
    squads.forEach(function (s) { prizes[s] = 0; });
    if (win < 0) squads.forEach(function (s) { prizes[s] = 3; });
    else mg.teams[win].forEach(function (s) { prizes[s] = pay[win]; g.squads[s].mgWins = (g.squads[s].mgWins || 0) + 1; });
    mg.winTeam = win;
  } else if (mg.mode === 'chase') {
    const r = mg.result || {}, h = mg.hiker, bonus = Math.floor((r.markers || 0) / 10);
    squads.forEach(function (s) { prizes[s] = 0; });
    if (mg.result) {
      mg.scores[h] = r.markers || 0;
      if (r.caught && mg.ghosts[r.by]) {
        Object.keys(mg.ghosts).forEach(function (s) { prizes[s] = 5; });
        prizes[r.by] = 10; prizes[h] = bonus;
        g.squads[r.by].mgWins = (g.squads[r.by].mgWins || 0) + 1;
      } else { prizes[h] = (r.cleared ? 15 : 10) + bonus; g.squads[h].mgWins = (g.squads[h].mgWins || 0) + 1; }
    }
  } else if (mg.mode === 'coop') {
    const avg = spAvg_(mg, squads), made = SP_MINIGAMES_[mg.game].better === 'low' ? avg <= mg.target : avg >= mg.target;
    squads.forEach(function (s) { prizes[s] = made ? 6 : -3; });
    mg.made = made; mg.avg = Math.round(avg);
  } else {
    const def = SP_MINIGAMES_[mg.game], has = squads.filter(function (s) { return mg.scores[s] != null; });
    const sorted = has.sort(function (a, b) { return def.better === 'low' ? mg.scores[a] - mg.scores[b] : mg.scores[b] - mg.scores[a]; });
    const pay = [10, 5, 3, 0];
    let place = 0;
    sorted.forEach(function (s, i) {
      if (i > 0 && mg.scores[s] !== mg.scores[sorted[i - 1]]) place = i;
      prizes[s] = pay[place] || 0;
      if (place === 0) g.squads[s].mgWins = (g.squads[s].mgWins || 0) + 1;
    });
    squads.forEach(function (s) { if (prizes[s] == null) prizes[s] = 0; });
  }
  squads.forEach(function (s) { spCoins_(g, s, prizes[s]); });
  mg.prizes = prizes;
  mg.done = true;
  mg.resolvedAt = Date.now();
  const best = Math.max.apply(null, squads.map(function (s) { return prizes[s]; })), top = squads.filter(function (s) { return best > 3 && prizes[s] === best; });
  if (mg.mode === 'coop') {
    spLog_(g, mg.name + ': ' + (mg.made ? 'everyone beat Shady! +6 coins each.' : 'Shady wins and steals 3 coins from each squad.'));
    spHost_(g, mg.made ? 'Teamwork! You beat Shady together. 6 coins for every squad!' : 'Shady wins this one and grabs 3 coins from every squad. Next time!');
  } else {
    spLog_(g, mg.name + ': ' + (top.length ? top.map(spTitle_).join(' and ') + ' won ' + best + ' coins.' : 'no winner.'));
    spHost_(g, top.length ? top.map(spTitle_).join(' and ') + ' win' + (top.length > 1 ? '' : 's') + ' ' + mg.name + '!' : 'No winner in ' + mg.name + '. Next round!');
  }
  // On to the next round, the Final Climb, or the end.
  g.round++;
  g.phase = 'turn';
  g.turnAt = Date.now() + 6000; // the results show for a moment before the auto-roll clock starts
  g.last = { squad: '', roll: 0, path: [], events: [] };
  if (g.round > g.rounds) { spBonus_(g); if (spStartShowdown_(g)) return; spFinish_(g, 'ended'); return; }
  if (!g.final && g.round >= g.rounds - 2 && g.rounds >= 4) {
    g.final = true;
    const last = spPlaces_(g).slice(-1)[0];
    spCoins_(g, last, 10);
    spLog_(g, 'The FINAL CLIMB! ' + spTitle_(last) + ' gets 10 catch-up coins. Blue and red spaces count double.');
    spHost_(g, 'It is the FINAL CLIMB! Three rounds left. ' + spTitle_(last) + ', here are 10 coins. Blue and red spaces count DOUBLE!');
    g.finalNews = Date.now();
  }
}

// ── Duels (SUMMIT_PARTY_DUELS): the squad on a duel space bets coins against a squad it picks. ──
function spBeginDuel_(g, s, opp, bet) {
  const keys = spSoloKeys_(), k = keys[((g.duels || 0) + (g.seed || 0)) % keys.length];
  g.duels = (g.duels || 0) + 1;
  const pot = Math.min(bet, g.squads[s].coins, g.squads[opp].coins);
  g.pending = null;
  g.phase = 'minigame';
  spNewMg_(g, k, 'duel', [s, opp], { bet: pot, challenger: s });
  spLog_(g, spTitle_(s) + ' challenged ' + spTitle_(opp) + ' to a duel for ' + pot + ' coins.');
  spHost_(g, spModeLine_(g) + g.mg.name + '!');
}
function spResolveDuel_(g) {
  const mg = g.mg, a = mg.squads[0], b = mg.squads[1], va = spVal_(mg, a), vb = spVal_(mg, b);
  const win = va === vb ? '' : spBetter_(mg, va, vb) ? a : b, lose = win === a ? b : a;
  if (win) {
    const take = Math.min(mg.bet, g.squads[lose].coins);
    spCoins_(g, lose, -take); spCoins_(g, win, mg.bet ? take : 5);
    mg.prizes = {}; mg.prizes[win] = mg.bet ? take : 5; mg.prizes[lose] = -take;
    spLog_(g, spTitle_(win) + ' won the duel' + (mg.bet ? ' and ' + take + ' coins from ' + spTitle_(lose) : ': +5 coins') + '.');
    spHost_(g, spTitle_(win) + ' wins the duel!' + (mg.bet ? ' ' + take + ' coins change hands.' : ''));
  } else { mg.prizes = {}; mg.prizes[a] = 0; mg.prizes[b] = 0; spLog_(g, 'The duel was a tie. Everyone keeps their coins.'); spHost_(g, 'A tie! Nobody loses coins.'); }
  mg.winner = win; mg.done = true; mg.resolvedAt = Date.now();
  g.duelDone = { id: mg.id, game: mg.game, name: mg.name, squads: mg.squads, scores: mg.scores, prizes: mg.prizes, winner: win, bet: mg.bet };
  g.phase = 'turn';
  spNextTurn_(g);
  if (g.phase === 'turn') g.turnAt = Date.now() + 6000;
}

// ── Showdown: an exact tie for first (flags and coins) at the end is settled by one more minigame. ──
function spStartShowdown_(g) {
  if (!spSwitch_('SUMMIT_PARTY_DUELS') || g.showdownDone) return false;
  const p = spPlaces_(g), top = g.squads[p[0]];
  const tied = p.filter(function (s) { return g.squads[s].flags === top.flags && g.squads[s].coins === top.coins; });
  if (tied.length < 2) return false;
  const keys = spSoloKeys_();
  g.phase = 'minigame';
  spNewMg_(g, keys[(g.seed || 0) % keys.length], 'showdown', tied);
  spLog_(g, 'Showdown: ' + tied.map(spTitle_).join(', ') + ' are tied for first.');
  spHost_(g, spModeLine_(g) + 'One more minigame decides it: ' + g.mg.name + '!');
  return true;
}
function spResolveShowdown_(g) {
  const mg = g.mg, order = mg.squads.slice().sort(function (a, b) { const va = spVal_(mg, a), vb = spVal_(mg, b); return va === vb ? mg.squads.indexOf(a) - mg.squads.indexOf(b) : spBetter_(mg, va, vb) ? -1 : 1; });
  const win = order[0];
  spCoins_(g, win, 1);
  mg.prizes = {}; mg.squads.forEach(function (s) { mg.prizes[s] = s === win ? 1 : 0; });
  mg.winner = win; mg.done = true; mg.resolvedAt = Date.now();
  g.showdownDone = true;
  spLog_(g, spTitle_(win) + ' won the Showdown!');
  spFinish_(g, 'ended');
}

// ── Trail Sketch: every squad draws the same word; Claude guesses from a short list of choices. ──
const SP_SKETCH_HEADERS_ = ['Show', 'Word'];
const SP_SKETCH_SEED_ = ['tent', 'compass', 'campfire', 'mountain', 'pine tree', 'backpack', 'canoe', 'snowflake', 'moon', 'fish', 'flag', 'boot', 'map', 'sun', 'cabin',
  'bear', 'owl', 'lantern', 'rope', 'cloud', 'river', 'bridge', 'star', 'deer', 'bicycle', 'mushroom', 'acorn', 'umbrella', 'kite', 'rainbow'].map(function (w) { return [true, w]; });
function spSketchWords_() {
  const rows = getSheetSafe_('Summit Sketch Words') ? sheetRows_('Summit Sketch Words').filter(function (r) { return ticked_(r['Show']) && String(r['Word'] || '').trim(); }) : [];
  const list = rows.length ? rows.map(function (r) { return String(r['Word']).trim().toLowerCase(); }) : SP_SKETCH_SEED_.map(function (r) { return r[1]; });
  return list.filter(function (w, i) { return list.indexOf(w) === i; });
}
function spSketchSetup_(mg) {
  const words = spSketchWords_(), pick = function (i) { return (mg.seed * 7 + i * 13) % words.length; };
  mg.word = words[mg.seed % words.length];
  const choices = [mg.word];
  for (let i = 0; choices.length < Math.min(10, words.length) && i < 200; i++) { const w = words[pick(i)]; if (choices.indexOf(w) < 0) choices.push(w); }
  mg.choices = choices.sort();
  mg.guesses = {};
}
// Only the drawing and the word list go to Claude: no names, no 800#s. The answer must be one of the choices.
function spAskClaude_(type, data, choices) {
  const key = PropertiesService.getScriptProperties().getProperty('CLAUDE_API_KEY');
  if (!key) throw new Error('Trail Sketch needs the Claude key (Script Property CLAUDE_API_KEY).');
  const res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { 'x-api-key': key, 'anthropic-version': '2023-06-01' },
    payload: JSON.stringify({
      model: String(getSetting_('SUMMIT_PARTY_DRAW_MODEL') || 'claude-haiku-5-5'), max_tokens: 20,
      messages: [{ role: 'user', content: [
        { type: 'image', source: { type: 'base64', media_type: 'image/' + type, data: data } },
        { type: 'text', text: 'This is a quick sketch from a school drawing game. Which ONE of these words does it show: ' + choices.join(', ') + '? Answer with only that word, written exactly as in the list. If the page is blank or you cannot tell yet, answer: unsure' }] }]
    })
  });
  if (res.getResponseCode() >= 300) { console.warn('Trail Sketch: Claude ' + res.getResponseCode() + ' ' + res.getContentText().slice(0, 200)); return 'unsure'; }
  const body = JSON.parse(res.getContentText() || '{}');
  const text = (body.content || []).filter(function (b) { return b.type === 'text'; }).map(function (b) { return b.text; }).join(' ').toLowerCase().replace(/[^a-z ]/g, ' ').replace(/\s+/g, ' ').trim();
  const exact = choices.filter(function (c) { return text === c; })[0];
  if (exact) return exact;
  const found = choices.filter(function (c) { return (' ' + text + ' ').indexOf(' ' + c + ' ') >= 0; }).sort(function (a, b) { return b.length - a.length; });
  return found.length === 1 ? found[0] : 'unsure';
}

// End of game: bonus flags (ties all get one).
function spBonus_(g) {
  if (String(getSetting_('SUMMIT_PARTY_BONUS') || 'ON').toUpperCase() === 'OFF') { g.bonus = []; return; }
  const squads = Object.keys(g.squads), out = [];
  [['Minigame Champ', 'mgWins'], ['Coin Keeper', 'peak'], ['Trail Explorer', 'events']].forEach(function (b) {
    const best = Math.max.apply(null, squads.map(function (s) { return g.squads[s][b[1]] || 0; }));
    if (best <= 0) return;
    const who = squads.filter(function (s) { return (g.squads[s][b[1]] || 0) === best; });
    who.forEach(function (s) { g.squads[s].flags++; });
    out.push({ name: b[0], squads: who, value: best });
    spLog_(g, 'Bonus flag, ' + b[0] + ': ' + who.map(spTitle_).join(' and ') + '.');
  });
  g.bonus = out;
}

function spFinish_(g, status) {
  g.status = status;
  g.phase = 'over';
  g.endedAt = new Date().toISOString();
  g.places = spPlaces_(g);
  const result = g.places.map(function (s, i) { return (i + 1) + '. ' + spTitle_(s) + ' (' + g.squads[s].flags + ' flags, ' + g.squads[s].coins + ' coins)'; }).join(' · ');
  if (status === 'ended') spHost_(g, 'And the winner of Summit Party is... ' + spTitle_(g.places[0]) + '!');
  if (g.demo) return; // a preview demo game: nothing on the Summit tabs, never any points
  try {
    const sh = spGamesTab_(), r = sheetRows_('Summit Games').filter(function (x) { return String(x['Game']) === g.id; })[0];
    if (r) sh.getRange(r._row, 2, 1, 6).setValues([[status === 'ended' ? 'Finished' : 'Stopped', g.startedAt.slice(0, 10), todayYmd_(), g.rounds, String(r['Started By'] || ''), result]]);
  } catch (e) { console.warn('Summit Games: ' + e); }
  if (status !== 'ended') return;
  const pts = String(getSetting_('SUMMIT_PARTY_POINTS') || '25,15,10,5').split(',').map(function (x) { return Math.max(0, Math.round(Number(x) || 0)); });
  if (!pts.some(Boolean)) return;
  const row = { BLACK: '', GOLD: '', GREY: '', PURPLE: '' };
  g.places.forEach(function (s, i) { row[s] = pts[i] || ''; });
  spPointsTab_().appendRow([g.id, todayYmd_(), result, row.BLACK, row.GOLD, row.GREY, row.PURPLE, '', '']);
}

function spGamesTab_() { return getSheetSafe_('Summit Games') || spMakeTab_('Summit Games', SP_GAME_HEADERS_); }
function spPointsTab_() { return getSheetSafe_('Summit Party Points') || spMakeTab_('Summit Party Points', SP_POINT_HEADERS_); }
function spMakeTab_(name, headers, rows) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().insertSheet(name);
  sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#490e6f').setFontColor('#ffffff');
  sh.setFrozenRows(1);
  if (rows && rows.length) sh.getRange(2, 1, rows.length, headers.length).setValues(rows);
  return sh;
}
function ensureSummitPartyTabs_() {
  spGamesTab_(); spPointsTab_();
  if (!getSheetSafe_('Summit Live Challenges')) spMakeTab_('Summit Live Challenges', SP_LIVE_HEADERS_, SP_LIVE_SEED_);
  if (!getSheetSafe_('Summit Sketch Words')) spMakeTab_('Summit Sketch Words', SP_SKETCH_HEADERS_, SP_SKETCH_SEED_);
}

// What the screen and the Chromebooks see.
function spView_(g) {
  const layout = spLayoutOf_(g && g.layout);
  return { on: true, board: spBoard_(layout), layout: layout, art: spArt_(), spaces: SP_SPACES_, items: SP_ITEMS_, minigames: SP_MINIGAMES_,
    flagCost: spFlagCost_(), autoRoll: spSetting_('SUMMIT_PARTY_AUTO_ROLL', 30),
    opts: { practice: spSwitch_('SUMMIT_PARTY_PRACTICE'), relay: spSwitch_('SUMMIT_PARTY_RELAY'), sound: spSwitch_('SUMMIT_PARTY_SOUND'), soundPlay: spSwitch_('SUMMIT_PARTY_SOUND_CHROMEBOOK') }, rounds: spSetting_('SUMMIT_PARTY_ROUNDS', 10), now: Date.now(), game: g };
}

// The clock: auto-roll (or auto-pick) for a squad that waits too long; minigames that run out of time.
// A preview demo game (g.demo = { mine }): the computer plays every other squad. Its turns go after
// 2.5 seconds, it is always ready, and its minigame scores come in near the end of the clock.
function spBot_(g, s) { return !!(g && g.demo && s && s !== g.demo.mine); }
function spBotScore_(k) {
  const d = SP_MINIGAMES_[k];
  if (d.better === 'low') return Math.min(d.max, Math.round(240 + Math.random() * 260));
  return Math.max(1, Math.round(d.max * (0.12 + Math.random() * 0.3)));
}
// Computer scores that are due now (fill = true writes them). Trail Chase is played on the big screen.
function spBotScores_(g, t, fill) {
  const mg = g.mg;
  if (!g.demo || !mg || mg.live || mg.done || mg.game === 'chase' || t < mg.endAt - 900) return false;
  const due = (mg.squads || Object.keys(g.squads)).filter(function (s) { return spBot_(g, s) && mg.scores[s] == null; });
  if (fill) due.forEach(function (s) { mg.scores[s] = spBotScore_(mg.game); });
  return due.length > 0;
}
function spAllScored_(g, t) {
  const mg = g.mg;
  return !!(g.demo && mg && !mg.live && mg.game !== 'chase' && t >= mg.endAt - 500 && (mg.squads || Object.keys(g.squads)).every(function (s) { return mg.scores[s] != null; }));
}
function spTickDue_(g, t, secs) {
  if (g.phase === 'minigame') return !!(g.mg && !g.mg.live && (t > g.mg.endAt + 8000 || spBotScores_(g, t, false) || spAllScored_(g, t)));
  if (g.phase !== 'turn') return false;
  const wait = spBot_(g, g.pending ? g.pending.squad : g.order[g.turn]) ? 2500 : secs > 0 ? secs * 1000 : 0;
  return wait > 0 && t - Number(g.turnAt || 0) >= wait;
}

function spTick_(g) {
  if (!g || g.status !== 'playing') return g;
  const secs = spSetting_('SUMMIT_PARTY_AUTO_ROLL', 30);
  if (!spTickDue_(g, Date.now(), secs)) return g;
  return spWithLock_(function () {
    const cur = spLoad_(g.code), board = spBoard_(cur && cur.layout), t = Date.now();
    if (!cur || cur.status !== 'playing' || !spTickDue_(cur, t, secs)) return cur;
    if (cur.phase === 'minigame') {
      spBotScores_(cur, t, true);
      if (t > cur.mg.endAt + 8000 || spAllScored_(cur, t)) spResolveMinigame_(cur);
      return spSave_(cur);
    }
    if (cur.pending) {
      const s = cur.pending.squad;
      if (cur.pending.kind === 'shop') { spLog_(cur, spTitle_(s) + ' walked past the Camp Store.'); spContinue_(cur, board, s); }
      else if (cur.pending.kind === 'duel') { const opp = cur.pending.options.slice().sort(function (a, b) { return b.coins - a.coins; })[0].id; spBeginDuel_(cur, s, opp, 5); }
      else { const opt = cur.pending.options[0]; spLog_(cur, spTitle_(s) + ' took the ' + opt.name + ' (auto).'); spTake_(cur, board, s, opt.id); }
    } else spRoll_(cur, board, cur.order[cur.turn], true);
    return spSave_(cur);
  });
}

function spRoll_(g, board, s, auto) {
  const me = g.squads[s], dice = me.dice || 1, fixed = me.fixed || 0;
  const dieA = 1 + Math.floor(Math.random() * 6), dieB = dice > 1 ? 1 + Math.floor(Math.random() * 6) : 0;
  const roll = fixed || (dieA + dieB);
  g.last = { squad: s, roll: roll, dice: fixed ? [fixed] : dieB ? [dieA, dieB] : [dieA], path: [], events: [], from: me.pos, auto: !!auto, fixed: !!fixed };
  me.dice = 0; me.fixed = 0;
  spLog_(g, spTitle_(s) + (fixed ? ' used the compass: ' + roll : ' rolled ' + roll) + (auto ? ' (auto-roll)' : '') + '.');
  spHost_(g, spTitle_(s) + ' rolled a ' + roll + '!');
  spWalk_(g, board, s, roll);
}
function spTake_(g, board, s, pick) {
  const me = g.squads[s], left = g.pending.left - 1;
  g.pending = null;
  me.pos = pick;
  me.moved = (me.moved || 0) + 1;
  g.last.path.push(pick);
  if (me.pos === g.flagAt) spTryFlag_(g, board, s);
  if (left > 0) spWalk_(g, board, s, left); else spLand_(g, board, s);
}
// After the Camp Store: keep walking the steps that were left.
function spContinue_(g, board, s) {
  const left = g.pending.left;
  g.pending = null;
  if (left > 0) spWalk_(g, board, s, left); else spLand_(g, board, s);
}

/** Public: one game for the screen and the controllers. sinceSeq = the page's last one ({ same: true } if nothing changed). */
function clashSummitState(code, sinceSeq) {
  let g = spLoad_(code);
  if (!spOn_() && !(g && g.demo && spCode_(code))) return JSON.stringify({ on: false });
  g = spTick_(g);
  if (g && Number(sinceSeq) && Number(sinceSeq) === g.seq) return JSON.stringify({ on: true, same: true, seq: g.seq, now: Date.now() });
  const v = spView_(g);
  if (!g) v.live = spIndex_().filter(function (x) { return x.status === 'playing'; }).length;
  return JSON.stringify(v);
}

/** Staff: start a new game; it gets its own code. */
function clashSummitStart(token, rounds, layout) {
  const who = spStaff_(token);
  spRequireOn_();
  return spCreate_(String(who.name || who.email || 'staff'), rounds, layout, null);
}
// A new game. demo = { mine: squad } for a preview demo game (65_Preview.gs): the computer plays the rest.
function spCreate_(by, rounds, layout, demo) {
  return spWithLock_(function () {
    if (!demo) ensureSummitPartyTabs_(); // first game: the Summit tabs, including the live challenges
    const list = spTidy_();
    if (!demo && list.filter(function (x) { return x.status === 'playing' && !x.demo; }).length >= SP_MAX_GAMES_) throw new Error('There are already ' + SP_MAX_GAMES_ + ' games going. End one first.');
    if (demo && list.filter(function (x) { return x.status === 'playing' && x.demo; }).length >= 10) throw new Error('Too many preview games right now. Try again in a little while.');
    const code = spNewCode_(list.map(function (x) { return x.code; }));
    const board = spBoard_(layout);
    const order = SQUAD_KEYS_.slice().sort(function () { return Math.random() - 0.5; });
    const g = { id: 'SP-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmm') + '-' + code, code: code, status: 'playing', phase: 'turn',
      startedAt: new Date().toISOString(), layout: spLayoutOf_(layout), seed: Math.floor(Math.random() * 6),
      rounds: Math.max(1, Math.min(30, Math.round(Number(rounds) || spSetting_('SUMMIT_PARTY_ROUNDS', 10)))), round: 1, turn: 0, order: order, turnAt: Date.now(),
      squads: {}, flagAt: 0, pending: null, mg: null, log: [], seq: 0, last: { squad: '', roll: 0, path: [], events: [] }, by: by };
    if (demo) { g.demo = { mine: demo.mine }; g.id = 'DEMO-' + code; }
    order.forEach(function (s) { g.squads[s] = { pos: 0, coins: 10, peak: 10, flags: 0, items: [], mgWins: 0, events: 0, moved: 0 }; });
    spMoveFlag_(g, board);
    spLog_(g, 'A new Summit Party! Game code ' + code + '. ' + spTitle_(order[0]) + ' goes first.');
    spHost_(g, 'Welcome to SUMMIT PARTY! I am Cliff, your host. Plant the most Summit Flags to win. ' + spTitle_(order[0]) + ', you are up first!' +
      (demo ? ' (Preview: you are ' + spTitle_(demo.mine) + '. The computer plays the others. No points.)' : ''));
    list.push(demo ? { code: code, status: 'playing', started: Date.now(), demo: true } : { code: code, status: 'playing', started: Date.now() });
    spSetIndex_(list);
    if (!demo) spGamesTab_().appendRow([g.id, 'Playing', todayYmd_(), '', g.rounds, g.by, '']);
    return JSON.stringify(spView_(spSave_(g)));
  });
}

/** Staff: skip the squad whose turn it is (nobody at their Chromebook), or end a live minigame with no winner. */
function clashSummitSkip(token, code) {
  spStaff_(token);
  spRequireOnFor_(code);
  return spWithLock_(function () {
    const g = spGame_(code), board = spBoard_(g.layout);
    if (g.status !== 'playing') throw new Error('That game is over.');
    if (g.phase === 'minigame') { spResolveMinigame_(g); return JSON.stringify(spView_(spSave_(g))); }
    const s = g.pending ? g.pending.squad : g.order[g.turn];
    g.last = { squad: s, roll: 0, path: [], events: [] };
    spLog_(g, spTitle_(s) + ' skipped a turn.');
    spNextTurn_(g, board);
    return JSON.stringify(spView_(spSave_(g)));
  });
}

/** Staff: the winner of a live Move, Make or Talk challenge. */
function clashSummitLiveWin(token, code, squad) {
  spStaff_(token);
  spRequireOnFor_(code);
  return spWithLock_(function () {
    const g = spGame_(code);
    if (g.status !== 'playing' || g.phase !== 'minigame' || !g.mg || !g.mg.live) throw new Error('No live challenge right now.');
    const s = String(squad || '').toUpperCase();
    if (!g.squads[s]) throw new Error('Pick a squad.');
    g.mg.winner = s;
    spResolveMinigame_(g);
    return JSON.stringify(spView_(spSave_(g)));
  });
}

/** Staff: end a game now. finish = true scores it (points wait for approval); false just stops it. */
function clashSummitEnd(token, code, finish) {
  spStaff_(token);
  return spWithLock_(function () {
    const g = spGame_(code);
    if (g.status !== 'playing') throw new Error('That game is over.');
    g.last = { squad: '', roll: 0, path: [], events: [] };
    if (finish) spBonus_(g);
    spFinish_(g, finish ? 'ended' : 'stopped');
    spLog_(g, finish ? 'The game was ended early. Final results are in.' : 'The game was stopped. No points.');
    return JSON.stringify(spView_(spSave_(g)));
  });
}

// A student at a squad's Chromebook: their 800# must be on the Student Board.
function spPlayer_(id, code) {
  const sample = sampleSquad_(id); // the previewed squad in a preview demo game (65_Preview.gs)
  if (sample) {
    const dg = spLoad_(code);
    if (!dg || !dg.demo || dg.demo.mine !== sample) throw new Error('Sample squads play only in a preview demo game.');
    return { id: SAMPLE_ID_, squad: sample };
  }
  spRequireOn_();
  const cache = CacheService.getScriptCache(), used = Number(cache.get('sp-calls')) || 0;
  if (used >= 6000) throw new Error('Summit Party is busy. Try again in a minute.');
  cache.put('sp-calls', String(used + 1), 600);
  const sid = id800_(id);
  if (sid.length < 5) throw new Error('Type your 800 number.');
  const me = boardRows_().filter(function (r) { return id800_(r['800#']) === sid; })[0];
  if (!me) throw new Error('That 800 number isn\'t on the list. Check it, or ask your teacher.');
  const squad = squadOf_(me['Squad']);
  if (SQUAD_KEYS_.indexOf(squad) === -1) throw new Error('You are not on a squad yet. Ask your teacher.');
  return { id: sid, squad: squad };
}

/** Controller: join a game (800# + the code on the big screen). */
function clashSummitJoin(id, code) {
  const p = spPlayer_(id, code), g = spGame_(code);
  return JSON.stringify({ squad: p.squad, code: g.code, state: JSON.parse(clashSummitState(g.code, 0)) });
}

function spMyTurn_(g, squad) {
  if (g.status !== 'playing') throw new Error('This game is over. Watch the big screen.');
  if (g.phase !== 'turn') throw new Error('It is minigame time!');
  if (g.order[g.turn] !== squad) throw new Error('It is ' + spTitle_(g.order[g.turn]) + '\'s turn.');
}

/** Controller: roll for your squad (only on its turn). */
function clashSummitRoll(id, code) {
  const p = spPlayer_(id, code);
  return spWithLock_(function () {
    const g = spGame_(code), board = spBoard_(g.layout);
    spMyTurn_(g, p.squad);
    if (g.pending) throw new Error(g.pending.kind === 'shop' ? 'Finish shopping first.' : 'Pick a path first.');
    spRoll_(g, board, p.squad, false);
    return JSON.stringify(spView_(spSave_(g)));
  });
}

/** Controller: use an item before rolling (one per turn). arg = the compass number. */
function clashSummitItem(id, code, item, arg) {
  const p = spPlayer_(id, code);
  if (!spOnFlag_()) throw new Error('Items are turned off.');
  return spWithLock_(function () {
    const g = spGame_(code), board = spBoard_(g.layout), me = g.squads[p.squad];
    spMyTurn_(g, p.squad);
    if (g.pending) throw new Error('Finish your move first.');
    const i = (me.items || []).indexOf(String(item));
    if (i === -1) throw new Error('You don\'t have that item.');
    if (me.usedOn === g.round) throw new Error('One item per turn.');
    me.items.splice(i, 1);
    me.usedOn = g.round;
    const name = SP_ITEMS_[item].name;
    g.last = { squad: p.squad, roll: 0, path: [], events: [], item: item };
    if (item === 'double') { me.dice = 2; spHost_(g, spTitle_(p.squad) + ' used Double Dice! Two dice this turn.'); }
    else if (item === 'compass') {
      const n = Math.round(Number(arg));
      if (!(n >= 1 && n <= 6)) { me.items.push(item); me.usedOn = 0; throw new Error('Pick 1 to 6.'); }
      me.fixed = n; spHost_(g, spTitle_(p.squad) + ' used the Golden Compass: exactly ' + n + '!');
    } else if (item === 'paws') {
      const lead = Object.keys(g.squads).filter(function (k) { return k !== p.squad; }).sort(function (a, b) { return g.squads[b].coins - g.squads[a].coins; })[0];
      const n = Math.min(10, g.squads[lead].coins);
      spCoins_(g, lead, -n); spCoins_(g, p.squad, n);
      spHost_(g, 'Sticky Paws! ' + spTitle_(p.squad) + ' grabbed ' + n + ' coins from ' + spTitle_(lead) + '!');
    } else if (item === 'zipline') {
      const others = Object.keys(g.squads).filter(function (k) { return k !== p.squad; }), o = others[Math.floor(Math.random() * others.length)];
      const pos = me.pos; me.pos = g.squads[o].pos; g.squads[o].pos = pos; g.last.swap = [p.squad, o];
      spHost_(g, 'Zipline! ' + spTitle_(p.squad) + ' and ' + spTitle_(o) + ' trade places!');
    }
    spLog_(g, spTitle_(p.squad) + ' used ' + name + '.');
    g.turnAt = Date.now();
    return JSON.stringify(spView_(spSave_(g)));
  });
}

/** Controller: at a fork pick the path (a node id); at the Camp Store pick an item id or 'skip'. */
function clashSummitChoose(id, code, choice) {
  const p = spPlayer_(id, code);
  return spWithLock_(function () {
    const g = spGame_(code), board = spBoard_(g.layout);
    if (g.status !== 'playing' || !g.pending) throw new Error('Nothing to choose right now.');
    if (g.pending.squad !== p.squad) throw new Error('It is ' + spTitle_(g.pending.squad) + '\'s choice.');
    if (g.pending.kind === 'duel') {
      const parts = String(choice).toUpperCase().split(':'), opp = parts[0], bet = Number(parts[1]) || 0;
      if (opp === p.squad || !g.squads[opp]) throw new Error('Pick a squad to challenge.');
      if ([0, 5, 10, 20].indexOf(bet) < 0 || bet > g.squads[p.squad].coins) throw new Error('Pick a bet you can pay.');
      spBeginDuel_(g, p.squad, opp, bet);
      return JSON.stringify(spView_(spSave_(g)));
    }
    if (g.pending.kind === 'shop') {
      const me = g.squads[p.squad], k = String(choice);
      if (k !== 'skip') {
        const it = SP_ITEMS_[k];
        if (!it) throw new Error('Pick an item.');
        if (me.coins < it.price) throw new Error('Not enough coins for that.');
        if ((me.items || []).length >= 2) throw new Error('Your pack is full (2 items).');
        me.coins -= it.price; me.items = (me.items || []).concat([k]);
        g.last.item = k;
        spLog_(g, spTitle_(p.squad) + ' bought a ' + it.name + '.');
        spHost_(g, 'Thanks for shopping! ' + spTitle_(p.squad) + ' got a ' + it.name + '.');
      } else spLog_(g, spTitle_(p.squad) + ' walked past the Camp Store.');
      spContinue_(g, board, p.squad);
      return JSON.stringify(spView_(spSave_(g)));
    }
    const pick = Number(choice), opt = g.pending.options.filter(function (o) { return o.id === pick; })[0];
    if (!opt) throw new Error('Pick one of the paths.');
    spLog_(g, spTitle_(p.squad) + ' took the ' + opt.name + '.');
    spTake_(g, board, p.squad, pick);
    return JSON.stringify(spView_(spSave_(g)));
  });
}

/** Controller: this squad's minigame score (the best one counts). */
function clashSummitScore(id, code, score) {
  const p = spPlayer_(id, code);
  return spWithLock_(function () {
    const g = spGame_(code), mg = g.mg;
    if (g.status !== 'playing' || g.phase !== 'minigame' || !mg || mg.live || mg.done) throw new Error('No minigame right now.');
    const now = Date.now();
    if ((mg.squads || Object.keys(g.squads)).indexOf(p.squad) < 0) throw new Error('Your squad sits this one out. Watch the big screen!');
    if (mg.game === 'sketch') throw new Error('Trail Sketch scores itself when the AI guesses your drawing.');
    if (mg.game === 'chase') throw new Error('Trail Chase is scored on the big screen.');
    if (now < mg.startAt - 1000 || now > mg.endAt + 8000) throw new Error('Time is up!');
    const def = SP_MINIGAMES_[mg.game];
    let v = Math.round(Number(score));
    if (!(v >= 0)) v = 0;
    if (def.better === 'low' && (v < 120 || v > def.max)) v = def.max; // a reaction faster than 0.12 s is not a reaction
    if (def.better === 'high') v = Math.min(v, def.max);
    const old = mg.scores[p.squad];
    if (old == null || (def.better === 'low' ? v < old : v > old)) mg.scores[p.squad] = v;
    if ((mg.squads || Object.keys(g.squads)).every(function (s) { return mg.scores[s] != null; }) && now >= mg.endAt - 500) spResolveMinigame_(g);
    return JSON.stringify(spView_(spSave_(g)));
  });
}

/** Controller: Trail Chase steering (U, D, L, R). Kept in the cache, not the game, so it is fast and needs no lock. */
function clashSummitSteer(id, code, dir) {
  const p = spPlayer_(id, code), g = spGame_(code), mg = g.mg, d = String(dir || '').toUpperCase();
  if (g.status !== 'playing' || g.phase !== 'minigame' || !mg || mg.game !== 'chase' || mg.done) throw new Error('No chase right now.');
  if (['U', 'D', 'L', 'R'].indexOf(d) < 0) throw new Error('Pick a direction.');
  if (mg.hiker !== p.squad && !mg.ghosts[p.squad]) throw new Error('Your squad sits this one out.');
  CacheService.getScriptCache().put('sp-steer-' + g.code + '-' + p.squad, d + ':' + Date.now(), 600);
  return JSON.stringify({ ok: true, dir: d });
}
/** Big screen: every squad's latest steering, with how long ago it came in (ms). */
function clashSummitSteers(code) {
  spRequireOnFor_(code);
  const g = spGame_(code), out = {}, c = CacheService.getScriptCache(), now = Date.now();
  Object.keys(g.squads).forEach(function (s) {
    const v = String(c.get('sp-steer-' + g.code + '-' + s) || '').split(':');
    if (v[0]) out[s] = { dir: v[0], ago: now - Number(v[1] || 0) };
  });
  return JSON.stringify(out);
}
/** Big screen (staff): how the chase ended. res = { caught, by (the ghost's squad), markers, cleared }. */
function clashSummitChaseEnd(token, code, res) {
  spStaff_(token);
  return spWithLock_(function () {
    const g = spGame_(code), mg = g.mg;
    if (g.status !== 'playing' || g.phase !== 'minigame' || !mg || mg.game !== 'chase' || mg.done) throw new Error('No chase right now.');
    const r = res || {}, by = String(r.by || '').toUpperCase();
    mg.result = { caught: !!r.caught && !!mg.ghosts[by], by: mg.ghosts[by] ? by : '', markers: Math.max(0, Math.min(400, Math.round(Number(r.markers) || 0))), cleared: !!r.cleared };
    spResolveMinigame_(g);
    return JSON.stringify(spView_(spSave_(g)));
  });
}
/** Controller: Rockslide. The running total of junk rows this squad has sent; every other squad takes them. */
function clashSummitJunk(id, code, total) {
  const p = spPlayer_(id, code);
  return spWithLock_(function () {
    const g = spGame_(code), mg = g.mg;
    if (g.status !== 'playing' || g.phase !== 'minigame' || !mg || mg.game !== 'rockslide' || mg.done) throw new Error('No Rockslide right now.');
    if ((mg.squads || Object.keys(g.squads)).indexOf(p.squad) < 0) throw new Error('Your squad sits this one out.');
    mg.junk = mg.junk || {};
    mg.junk[p.squad] = Math.max(mg.junk[p.squad] || 0, Math.min(500, Math.round(Number(total) || 0)));
    spSave_(g);
    return JSON.stringify({ junk: mg.junk });
  });
}

/** Controller: Ready after practice. Once every squad in the minigame is ready, it starts in 3.5 seconds. */
function clashSummitReady(id, code) {
  const p = spPlayer_(id, code);
  return spWithLock_(function () {
    const g = spGame_(code), mg = g.mg;
    if (g.status !== 'playing' || g.phase !== 'minigame' || !mg || mg.live || mg.done) throw new Error('No minigame right now.');
    mg.ready = mg.ready || {};
    mg.ready[p.squad] = true;
    const now = Date.now();
    if ((mg.squads || Object.keys(g.squads)).every(function (s) { return mg.ready[s]; }) && mg.startAt - now > 3500) {
      mg.startAt = now + 3500; mg.endAt = mg.startAt + (mg.secs || 20) * 1000;
    }
    return JSON.stringify(spView_(spSave_(g)));
  });
}

/** Controller: Trail Sketch. Sends the drawing (a small JPEG or PNG data URL); Claude guesses. At most one guess every 3 seconds per squad. */
function clashSummitGuess(id, code, img) {
  const p = spPlayer_(id, code), g0 = spGame_(code), mg0 = g0.mg;
  if (g0.status !== 'playing' || g0.phase !== 'minigame' || !mg0 || mg0.game !== 'sketch' || mg0.done) throw new Error('No drawing game right now.');
  if ((mg0.squads || Object.keys(g0.squads)).indexOf(p.squad) < 0) throw new Error('Your squad sits this one out.');
  if (Date.now() < mg0.startAt - 500 || Date.now() > mg0.endAt + 1500) throw new Error('Time is up!');
  if (mg0.scores[p.squad] != null) return JSON.stringify({ guess: mg0.word, right: true, score: mg0.scores[p.squad] });
  const m = /^data:image\/(png|jpeg);base64,([A-Za-z0-9+\/=]+)$/.exec(String(img || ''));
  if (!m || m[2].length > 300000) throw new Error('That drawing could not be read.');
  const cache = CacheService.getScriptCache(), ck = 'sp-guess-' + g0.code + '-' + p.squad;
  if (Number(cache.get(ck) || 0) > Date.now() - 2500) return JSON.stringify({ guess: '', wait: true });
  cache.put(ck, String(Date.now()), 10);
  if (!spSwitch_('SUMMIT_PARTY_DRAW')) throw new Error('Trail Sketch is turned off.');
  const guess = spAskClaude_(m[1], m[2], mg0.choices || []);
  return spWithLock_(function () {
    const g = spGame_(code), mg = g.mg;
    if (!mg || mg.id !== mg0.id || mg.done) return JSON.stringify({ guess: guess, right: false, late: true });
    mg.guesses = mg.guesses || {};
    mg.guesses[p.squad] = guess;
    const right = guess === mg.word && mg.scores[p.squad] == null && Date.now() <= mg.endAt + 1500;
    if (right) mg.scores[p.squad] = Math.max(20, Math.min(200, Math.round(20 + Math.max(0, mg.endAt - Date.now()) / 1000 * 4)));
    if ((mg.squads || Object.keys(g.squads)).every(function (s) { return mg.scores[s] != null; })) spResolveMinigame_(g);
    spSave_(g);
    return JSON.stringify({ guess: guess, right: right, score: mg.scores[p.squad] });
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
var FILE_STAMP_63_SummitParty = '065bc1c147';
