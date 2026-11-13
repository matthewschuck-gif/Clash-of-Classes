/**
 * Clash of Classes -- Apps Script backend
 * File 15: the game layer. Board-game mechanics on top of the points everyone already earns.
 * Every piece is a Sheet tab staff can edit, with a Show box on each row. Nothing here ever
 * takes points away; bonuses only go out once Clash of Classes › Turn on game awards is on.
 *
 * SCHOOL MAP (tab Map Zones) -- King of the Mountain, color sets and territory:
 *   Each zone is claimed by the squad with the most points from its events in the last
 *   CLAIM_DAYS days (Claimed By: event types or words in the event name; * = every point).
 *   The holder keeps it on a tie or a quiet stretch (king of the hill). The Summit zone's holder
 *   is King of the Mountain. Every week each holder earns the zone's Weekly Points (Summit
 *   x2), and holding every zone in a Wing (a color set) pays that wing again (WING_MULTIPLIER).
 *   How each zone is won (Conquer By), the Risk borders and the layout live in 18_Conquest.gs.
 *   The map is laid out like the building, by area only: never room numbers, doors or exits.
 * CHANCE CARDS (tabs Chance Cards, Squad Hand): Generals draw CHANCE_DRAWS_PER_WEEK cards a week
 *   in Generals HQ. Points cards count at once; the rest wait in the squad's hand:
 *   double_rent (next territory payout x2), freeze_pass (ends a Fun Freeze, never a Peaks
 *   Freeze), streak_shield (a Hot Streak survives one missed day), story (just for fun) and
 *   flip_three: a rival squad flips three cards; any card that comes up twice busts (those
 *   copies are lost, never points).
 * HOT STREAK (tab Streak Banks): days in a row a squad had the most Daily Poll voters. Ride it
 *   or bank it: Generals bank it in HQ for STREAK_BANK points; if another squad tops a day
 *   first, the streak cools off and nothing is lost.
 * ROUND GOALS (tab Round Goals): a goal for each stretch (a month), measured for every squad;
 *   1st / 2nd / 3rd get their points after End.
 * ENGINE POWERS (tab Engine Powers): lasting powers a squad unlocks (Unlock When), in three
 *   habitats: Foothills, Ridge, Summit. Effects: rent, draws, streak, tickets, poll.
 * THE OPPONENT (tab Opponents): the whole school against one opponent. Show box per row,
 *   and Clash of Classes › Show or hide the Opponent flips OPPONENT_VISIBLE. (It was "Monster of
 *   the Month"; an older Monsters tab is renamed, and the code still calls it monster inside.)
 */

const GAME_TYPES_ = ['territory', 'chance', 'streak_bank', 'round_goal', 'engine', 'monster', 'destination_ticket', 'freeze_thaw', 'evolution_boost', 'mp_bonus'];
const MAP_HEADERS_ = ['Show', 'Zone', 'Wing', 'Summit', 'Claimed By', 'Weekly Points', 'X', 'Y', 'W', 'H', 'Holder', 'Since', 'Conquer By', 'Borders'];
const CARD_HEADERS_ = ['Show', 'Card', 'Text', 'Effect', 'Value', 'Copies'];
const HAND_HEADERS_ = ['Squad', 'Card', 'Effect', 'Value', 'Drawn', 'Status', 'Note'];
const BANK_HEADERS_ = ['Squad', 'Through', 'Length', 'Points', 'At', 'By'];
const ROUND_HEADERS_ = ['Show', 'Start', 'End', 'Goal', 'Measure', '1st', '2nd', '3rd', 'Status', 'Results'];
const ENGINE_HEADERS_ = ['Show', 'Habitat', 'Power', 'Text', 'Unlock When', 'Effect', 'Value', 'Unlocked'];
const MONSTER_HEADERS_ = ['Show', 'Start', 'End', 'Opponent', 'Story', 'Image', 'Defeat Goal', 'Reward', 'Status', 'Defeated Image'];

function monthBounds_(offset) {
  const d = new Date(); d.setDate(1); d.setMonth(d.getMonth() + (offset || 0));
  const tz = Session.getScriptTimeZone();
  const start = Utilities.formatDate(d, tz, 'yyyy-MM-dd');
  const end = new Date(d.getFullYear(), d.getMonth() + 1, 0);
  return [start, Utilities.formatDate(end, tz, 'yyyy-MM-dd')];
}

function ensureGameTabs_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const m0 = monthBounds_(0), m1 = monthBounds_(1);
  // "Monster of the Month" is "the Opponent" now: rename an older tab and its name column.
  const oldTab = ss.getSheetByName('Monsters');
  if (oldTab && !ss.getSheetByName('Opponents')) {
    oldTab.setName('Opponents');
    const col = oldTab.getRange(1, 1, 1, oldTab.getLastColumn()).getValues()[0].map(String).indexOf('Monster');
    if (col > -1) oldTab.getRange(1, col + 1).setValue('Opponent');
  }
  // Older Opponents tabs get the Defeated Image column.
  const opp = ss.getSheetByName('Opponents');
  if (opp && opp.getLastColumn()) {
    const head = opp.getRange(1, 1, 1, opp.getLastColumn()).getValues()[0].map(function (h) { return String(h).trim(); });
    if (head.indexOf('Defeated Image') === -1) opp.getRange(1, head.length + 1).setValue('Defeated Image').setFontWeight('bold').setBackground('#490e6f').setFontColor('#ffffff');
  }
  [['Map Zones', MAP_HEADERS_, null],
   ['Chance Cards', CARD_HEADERS_, [
      [true, 'Trail Magic', 'Someone left snacks at the trailhead. +3', 'points', 3, 3],
      [true, 'Base Camp Supplies', 'Your squad restocks. +2', 'points', 2, 4],
      [true, 'Lucky Switchback', 'A shortcut appears. +1', 'points', 1, 4],
      [true, 'Summit Fever', 'Your next territory payout is doubled.', 'double_rent', 0, 2],
      [true, 'Second Wind', 'Ends a Fun Freeze on your squad (never a Peaks Freeze).', 'freeze_pass', 0, 1],
      [true, 'Flip Three', 'Make a rival squad flip three cards. A card that comes up twice busts.', 'flip_three', 0, 2],
      [true, 'Steady Footing', 'Your Hot Streak survives one missed day.', 'streak_shield', 0, 2],
      [true, 'Echo in the Canyon', 'Nothing happens, but it sounded cool.', 'story', 0, 2]]],
   ['Squad Hand', HAND_HEADERS_, null],
   ['Streak Banks', BANK_HEADERS_, null],
   ['Round Goals', ROUND_HEADERS_, [
      [true, m0[0], m0[1], 'Most Daily Poll voters', 'voters', 8, 5, 3, '', ''],
      [true, m1[0], m1[1], 'Most of the squad pitching in', 'pitched in', 8, 5, 3, '', '']]],
   ['Engine Powers', ENGINE_HEADERS_, [
      [true, 'Foothills', 'Trailhead Crew', 'Draw one extra Chance card each week.', 'points: 250', 'draws', 1, ''],
      [true, 'Foothills', 'Loud Echo', '+1 point each week for every 10 Daily Poll votes your squad casts.', 'streak: 3', 'poll', 1, ''],
      [true, 'Ridge', 'Switchback Builders', '+1 point for every zone you hold, every week.', 'zones: 2', 'rent', 1, ''],
      [true, 'Ridge', 'Storytellers', '+1 point for every destination ticket your squad completes.', 'tickets: 10', 'tickets', 1, ''],
      [true, 'Summit', 'Rope Team', '+2 points every time you bank a Hot Streak.', 'wings: 1', 'streak', 2, ''],
      [true, 'Summit', 'Summit Guides', 'Another +1 point for every zone you hold, every week.', 'round goals: 1', 'rent', 1, '']]],
   ['Opponents', MONSTER_HEADERS_, [
      // Shady: the small blue hairy gremlin of schadenfreude, back from a project that never left beta.
      [false, m0[0], m0[1], 'Shady', 'Shady is a small blue gremlin who feeds on schadenfreude: every laugh at someone else\'s stumble makes his fur stand taller. ' +
        'He can\'t stand Freudenfreude. Every Trait Dough bill for someone doing the right thing, every cheer for someone else\'s win, shrinks him.', '', 'entries: trait_dough 150', 5, '', ''],
      [false, m0[0], m0[1], 'The Tardy Troll', 'It lives under the bell and feeds on late arrivals. Every Daily Poll vote is a stone thrown at its bridge.', '', 'voters: 400', 5, '', ''],
      [false, m1[0], m1[1], 'The Hallway Litter Beast', 'It grows with every wrapper left behind. Every positive point the school earns shrinks it.', '', 'points: 1500', 5, '', '']]]
  ].forEach(function (t) {
    if (ss.getSheetByName(t[0])) return;
    const sh = ss.insertSheet(t[0]);
    sh.getRange(1, 1, 1, t[1].length).setValues([t[1]]).setFontWeight('bold').setBackground('#490e6f').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    if (t[2]) sh.getRange(2, 1, t[2].length, t[1].length).setValues(t[2]);
    if (t[0] === 'Map Zones') writeRowsByHeader_(sh, buildingZoneRows_()); // 18_Conquest.gs
  });
  ensureConquestTabs_();
}

function gameOn_(key) { return String(getSetting_(key) || 'ON').toUpperCase() !== 'OFF'; }
function awardsOn_() { return String(getSetting_('GAME_AWARDS') || 'OFF').toUpperCase() === 'ON'; }
function listOf_(v) { return String(v || '').split(',').map(function (s) { return s.trim().toLowerCase(); }).filter(String); }

// Positive points per squad from events in [from, to], counted by a matcher (game bonuses and
// private behavior rows never count toward claims or goals).
function pointsBy_(events, from, to, match) {
  const priv = listOf_(getSetting_('PRIVATE_EVENT_TYPES') || 'trail_journal, trailback_panel, camp_mountaineer');
  const out = { BLACK: 0, GOLD: 0, GREY: 0, PURPLE: 0 }, rows = { BLACK: 0, GOLD: 0, GREY: 0, PURPLE: 0 };
  events.forEach(function (e) {
    const d = String(e.date || '').slice(0, 10);
    const type = String(e.type || '').toLowerCase();
    if (!d || d < from || d > to || GAME_TYPES_.indexOf(type) > -1 || priv.indexOf(type) > -1) return;
    if (match && !match(e)) return;
    SQUAD_KEYS_.forEach(function (s) { const p = Number(e.pts[s]) || 0; if (p > 0) { out[s] += p; rows[s]++; } });
  });
  return { points: out, rows: rows };
}

function eventMatcher_(spec) {
  const words = listOf_(spec);
  if (!words.length || words.indexOf('*') > -1) return null;
  return function (e) {
    const type = String(e.type || '').toLowerCase(), name = String(e.name || '').toLowerCase();
    return words.some(function (w) { return type === w || name.indexOf(w) > -1; });
  };
}

// ── SCHOOL MAP ──────────────────────────────────────────────────────
// mapZones_ (who holds each zone) is in 18_Conquest.gs.

// Wings fully held by one squad: { wing: squad }.
function heldWings_(zones) {
  const wings = {};
  zones.forEach(function (z) { if (z.wing) (wings[z.wing] = wings[z.wing] || []).push(z.holder); });
  const out = {};
  Object.keys(wings).forEach(function (w) { const h = wings[w]; if (h.length > 1 && h[0] && h.every(function (x) { return x === h[0]; })) out[w] = h[0]; });
  return out;
}

// ── CHANCE CARDS ────────────────────────────────────────────────────
function chanceDeck_() {
  return sheetRows_('Chance Cards').filter(function (r) { return ticked_(r['Show']) && String(r['Card'] || '').trim(); }).map(function (r) {
    return { card: String(r['Card']).trim(), text: String(r['Text'] || ''), effect: String(r['Effect'] || 'story').trim().toLowerCase(),
      value: Number(r['Value']) || 0, copies: Math.max(0, Number(r['Copies']) || 1) };
  });
}

function drawFromDeck_(deck) {
  const total = deck.reduce(function (n, c) { return n + c.copies; }, 0);
  if (!total) throw new Error('The Chance Cards tab has no cards to draw.');
  let pick = Math.random() * total;
  for (let i = 0; i < deck.length; i++) { pick -= deck[i].copies; if (pick < 0) return deck[i]; }
  return deck[deck.length - 1];
}

function handFor_(squad) {
  return sheetRows_('Squad Hand').filter(function (r) { return squadOf_(r['Squad']) === squad; }).map(function (r) {
    return { row: r._row, card: String(r['Card'] || ''), effect: String(r['Effect'] || '').toLowerCase(), value: Number(r['Value']) || 0,
      drawn: String(r['Drawn'] || ''), status: String(r['Status'] || ''), note: String(r['Note'] || '') };
  });
}

function addPointsEvent_(name, type, note, squadPoints) {
  const pts = { BLACK: '', GOLD: '', GREY: '', PURPLE: '' };
  let any = false;
  Object.keys(squadPoints).forEach(function (s) { if (squadPoints[s] > 0) { pts[s] = squadPoints[s]; any = true; } });
  if (any) getSheet_(TABS.EVENTS).appendRow(['', name, type, todayYmd_(), note || '', '', pts.BLACK, pts.GOLD, pts.GREY, pts.PURPLE]);
  return any;
}

// One card into a squad's hand (points cards count at once). Returns the card with its outcome.
function giveCard_(squad, card, note) {
  const sheet = getSheet_('Squad Hand');
  let status = 'Held';
  if (card.effect === 'points') {
    status = card.value > 0 ? 'Counted' : 'Used';
    if (card.value > 0) addPointsEvent_('Chance: ' + card.card, 'chance', titleSquad_(squad) + ' drew it', (function () { const o = {}; o[squad] = card.value; return o; })());
  } else if (card.effect === 'story') status = 'Used';
  sheet.appendRow([squad, card.card, card.effect, card.value, todayYmd_(), status, note || '']);
  return { card: card.card, text: card.text, effect: card.effect, value: card.value, status: status };
}

function weekStart_(ymd) {
  const d = new Date((ymd || todayYmd_()) + 'T12:00:00');
  d.setDate(d.getDate() - ((d.getDay() + 6) % 7));
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function drawsLeft_(squad, engine) {
  const per = Math.max(0, Number(getSetting_('CHANCE_DRAWS_PER_WEEK')) || 1) + engineBonus_(engine, squad, 'draws');
  const wk = weekStart_();
  const used = handFor_(squad).filter(function (h) { return h.drawn >= wk && !/^flipped|^flip three/i.test(h.note); }).length;
  return Math.max(0, per - used);
}

function clashGeneralsDrawCard(token) {
  const me = hqSession_(token);
  if (!gameOn_('GAME_CARDS')) throw new Error('Chance cards are off.');
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    if (!drawsLeft_(me.squad, enginePowers_())) throw new Error('Your squad already drew its card this week.');
    const got = giveCard_(me.squad, drawFromDeck_(chanceDeck_()), 'Drawn by ' + me.name);
    const out = JSON.parse(clashGeneralsHq(token)); out.drew = got;
    return JSON.stringify(out);
  } finally { lock.releaseLock(); }
}

function clashGeneralsPlayCard(token, row, target) {
  const me = hqSession_(token);
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  let result = null;
  try {
    const card = handFor_(me.squad).filter(function (h) { return h.row === Number(row); })[0];
    if (!card || card.status !== 'Held') throw new Error('That card isn\'t in your hand.');
    const sheet = getSheet_('Squad Hand'), statusCol = HAND_HEADERS_.indexOf('Status') + 1, noteCol = HAND_HEADERS_.indexOf('Note') + 1;
    const today = todayYmd_();
    if (card.effect === 'double_rent' || card.effect === 'streak_shield') {
      sheet.getRange(card.row, statusCol).setValue('Armed');
    } else if (card.effect === 'freeze_pass') {
      const f = freezeWindows_().filter(function (w) { return w.squad === me.squad && !w.thawed && today >= w.start && (!w.end || today <= w.end); })[0];
      if (!f) throw new Error('Your squad isn\'t frozen.');
      if (f.kind !== 'FUN') throw new Error('Second Wind only works on a Fun Freeze. A Peaks Freeze thaws when your squad works together.');
      endFreeze_(f, 'Second Wind card ' + today + ' (' + me.name + ')');
      sheet.getRange(card.row, statusCol).setValue('Used ' + today);
    } else if (card.effect === 'flip_three') {
      const rival = squadOf_(target);
      if (SQUAD_KEYS_.indexOf(rival) === -1 || rival === me.squad) throw new Error('Pick a rival squad.');
      const deck = chanceDeck_(), three = [drawFromDeck_(deck), drawFromDeck_(deck), drawFromDeck_(deck)];
      const count = {}; three.forEach(function (c) { count[c.card] = (count[c.card] || 0) + 1; });
      result = three.map(function (c) {
        if (count[c.card] > 1) {
          getSheet_('Squad Hand').appendRow([rival, c.card, c.effect, c.value, today, 'Busted', 'Flipped by ' + titleSquad_(me.squad) + ' (came up twice)']);
          return { card: c.card, text: c.text, status: 'Busted' };
        }
        return giveCard_(rival, c, 'Flipped by ' + titleSquad_(me.squad));
      });
      sheet.getRange(card.row, statusCol, 1, 2).setValues([['Used ' + today, 'On ' + titleSquad_(rival)]]);
    } else {
      sheet.getRange(card.row, statusCol).setValue('Used ' + today);
    }
    if (card.effect !== 'flip_three') sheet.getRange(card.row, noteCol).setValue((card.note ? card.note + ' · ' : '') + 'Played by ' + me.name);
  } finally { lock.releaseLock(); }
  const out = JSON.parse(clashGeneralsHq(token)); out.flipped = result;
  return JSON.stringify(out);
}

// ── HOT STREAK ──────────────────────────────────────────────────────
// { squad: { length, through, shielded, last } } from finished Daily Polls since each squad's last bank.
function hotStreaks_() {
  const today = todayYmd_();
  const polls = pollsWithVotes_().filter(function (p) { return p.date && p.date < today && p.total > 0; })
    .sort(function (a, b) { return a.date.localeCompare(b.date); });
  const banks = sheetRows_('Streak Banks');
  const out = {};
  SQUAD_KEYS_.forEach(function (s) {
    const lastBank = banks.filter(function (b) { return squadOf_(b['Squad']) === s; }).map(function (b) { return String(b['Through'] || ''); }).sort().pop() || '';
    const shields = handFor_(s).filter(function (h) { return h.effect === 'streak_shield' && h.status === 'Armed'; }).length;
    let length = 0, through = '', gaps = 0;
    for (let i = polls.length - 1; i >= 0; i--) {
      const p = polls[i];
      if (p.date <= lastBank) break;
      const voters = {}; SQUAD_KEYS_.forEach(function (q) { voters[q] = Object.keys(p.bySquad[q] || {}).reduce(function (n, c) { return n + p.bySquad[q][c]; }, 0); });
      const best = Math.max.apply(null, SQUAD_KEYS_.map(function (q) { return voters[q]; }));
      const top = best > 0 && voters[s] === best;
      if (top) { length++; if (!through) through = p.date; continue; }
      if (length === 0 && i === polls.length - 1 && gaps < shields) { gaps++; continue; } // a shield covers the most recent miss
      if (length > 0 && gaps < shields) { gaps++; continue; }
      break;
    }
    out[s] = { length: length, through: through, shieldsUsed: gaps, shields: shields };
  });
  return out;
}

function streakTable_() {
  return String(getSetting_('STREAK_BANK') || '2=2, 3=4, 4=7, 5=11, 6=16').split(',').map(function (p) {
    const m = p.match(/(\d+)\s*=\s*(\d+)/); return m ? { days: Number(m[1]), pts: Number(m[2]) } : null;
  }).filter(Boolean).sort(function (a, b) { return a.days - b.days; });
}

function streakValue_(length, squad, engine) {
  const t = streakTable_().filter(function (x) { return length >= x.days; }).pop();
  return t ? t.pts + engineBonus_(engine, squad, 'streak') : 0;
}

function clashGeneralsBankStreak(token) {
  const me = hqSession_(token);
  if (!gameOn_('GAME_STREAK')) throw new Error('Hot Streaks are off.');
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const st = hotStreaks_()[me.squad];
    const pts = streakValue_(st.length, me.squad, enginePowers_());
    if (!pts) throw new Error('Your squad needs a longer streak before it can bank one.');
    getSheet_('Streak Banks').appendRow([me.squad, st.through, st.length, pts, todayYmd_(), me.name]);
    const o = {}; o[me.squad] = pts;
    addPointsEvent_('Hot Streak banked: ' + st.length + ' days', 'streak_bank', titleSquad_(me.squad) + ' topped Daily Poll turnout ' + st.length + ' days running', o);
    // Shields that carried the streak are spent.
    const sheet = getSheet_('Squad Hand');
    handFor_(me.squad).filter(function (h) { return h.effect === 'streak_shield' && h.status === 'Armed'; }).slice(0, st.shieldsUsed)
      .forEach(function (h) { sheet.getRange(h.row, HAND_HEADERS_.indexOf('Status') + 1).setValue('Used ' + todayYmd_()); });
  } finally { lock.releaseLock(); }
  return clashGeneralsHq(token);
}

// ── ROUND GOALS ─────────────────────────────────────────────────────
// A measure for every squad in [from, to]: points[: type], entries: type, voters, pitched in, zones, tickets.
function measureSquads_(measure, from, to, events) {
  const m = String(measure || '').trim().toLowerCase();
  const zero = { BLACK: 0, GOLD: 0, GREY: 0, PURPLE: 0 };
  let k = m, arg = '';
  const i = m.indexOf(':'); if (i > -1) { k = m.slice(0, i).trim(); arg = m.slice(i + 1).trim(); }
  if (k === 'points') return pointsBy_(events, from, to, arg ? eventMatcher_(arg) : null).points;
  if (k === 'entries') return pointsBy_(events, from, to, arg ? eventMatcher_(arg) : null).rows;
  if (k === 'voters') {
    const out = Object.assign({}, zero);
    sheetRows_('Poll Votes').forEach(function (v) { const d = String(v['Date'] || '').slice(0, 10); const s = squadOf_(v['Squad']); if (d >= from && d <= to && out[s] !== undefined) out[s]++; });
    return out;
  }
  if (k === 'pitched in') {
    const out = Object.assign({}, zero), size = Object.assign({}, zero);
    sheetRows_(getSetting_('STUDENT_BOARD_TAB') || 'Student Board').forEach(function (r) {
      const s = squadOf_(r['Squad']); if (size[s] === undefined) return;
      size[s]++; const d = String(r['Last Contribution'] || '').slice(0, 10); if (d >= from && d <= to) out[s]++;
    });
    SQUAD_KEYS_.forEach(function (s) { out[s] = size[s] ? Math.round(100 * out[s] / size[s]) : 0; });
    return out;
  }
  if (k === 'zones') { const out = Object.assign({}, zero); mapZones_(events).forEach(function (z) { if (z.holder) out[z.holder]++; }); return out; }
  if (k === 'tickets') {
    const out = Object.assign({}, zero);
    sheetRows_('Student Tickets').forEach(function (t) { const d = String(t['Completed'] || '').slice(0, 10); const s = squadOf_(t['Squad']); if (d && d >= from && d <= to && out[s] !== undefined) out[s]++; });
    return out;
  }
  return zero;
}

function places_(scores) {
  const vals = SQUAD_KEYS_.map(function (s) { return scores[s]; }).filter(function (v) { return v > 0; }).sort(function (a, b) { return b - a; });
  const uniq = vals.filter(function (v, i) { return vals.indexOf(v) === i; });
  const out = {};
  SQUAD_KEYS_.forEach(function (s) { const p = uniq.indexOf(scores[s]); out[s] = scores[s] > 0 && p > -1 && p < 3 ? p + 1 : 0; });
  return out;
}

function roundGoals_(events) {
  events = events || readEventsTab_();
  const today = todayYmd_();
  return sheetRows_('Round Goals').filter(function (r) { return ticked_(r['Show']) && String(r['Goal'] || '').trim() && r['Start']; }).map(function (r) {
    const from = String(r['Start']).slice(0, 10), to = String(r['End'] || '9999-12-31').slice(0, 10);
    const scores = measureSquads_(r['Measure'], from, to, events);
    return { row: r._row, goal: String(r['Goal']).trim(), measure: String(r['Measure'] || ''), start: from, end: to,
      prizes: [Number(r['1st']) || 0, Number(r['2nd']) || 0, Number(r['3rd']) || 0], scores: scores, places: places_(scores),
      status: String(r['Status'] || ''), active: today >= from && today <= to, done: today > to };
  });
}

// ── ENGINE POWERS ───────────────────────────────────────────────────
function engineReached_(rule, squad, ctx) {
  const m = String(rule || '').trim().toLowerCase().match(/^(points|zones|wings|round goals|streak|tickets|date)\s*:\s*(.+)$/);
  if (!String(rule || '').trim()) return true;
  if (!m) return false;
  const v = m[2].trim(), n = Number(v) || 0;
  if (m[1] === 'date') return todayYmd_() >= v;
  if (m[1] === 'points') { ctx.totals = ctx.totals || squadTotals_(); return ctx.totals[squad] >= n; }
  if (m[1] === 'zones') { ctx.zones = ctx.zones || mapZones_(ctx.events); return ctx.zones.filter(function (z) { return z.holder === squad; }).length >= n; }
  if (m[1] === 'wings') { ctx.zones = ctx.zones || mapZones_(ctx.events); const w = heldWings_(ctx.zones); return Object.keys(w).filter(function (k) { return w[k] === squad; }).length >= n; }
  if (m[1] === 'round goals') { ctx.rounds = ctx.rounds || sheetRows_('Round Goals'); return ctx.rounds.filter(function (r) { return new RegExp('\\b1st: [^;]*' + squad).test(String(r['Results'] || '')); }).length >= n; }
  if (m[1] === 'streak') return sheetRows_('Streak Banks').some(function (b) { return squadOf_(b['Squad']) === squad && Number(b['Length']) >= n; });
  if (m[1] === 'tickets') return sheetRows_('Student Tickets').filter(function (t) { return squadOf_(t['Squad']) === squad && t['Completed']; }).length >= n;
  return false;
}

// [{ habitat, power, text, effect, value, rule, unlocked: { SQUAD: date } }]; records new unlocks when write.
function enginePowers_(write, events) {
  const rows = sheetRows_('Engine Powers').filter(function (r) { return ticked_(r['Show']) && String(r['Power'] || '').trim(); });
  const sheet = write ? getSheetSafe_('Engine Powers') : null, col = ENGINE_HEADERS_.indexOf('Unlocked') + 1, today = todayYmd_();
  const ctx = { events: events };
  return rows.map(function (r) {
    const unlocked = parseUnlocked_(r['Unlocked']);
    if (write) {
      let changed = false;
      SQUAD_KEYS_.forEach(function (s) { if (!unlocked[s] && engineReached_(r['Unlock When'], s, ctx)) { unlocked[s] = today; changed = true; } });
      if (changed && sheet) sheet.getRange(r._row, col).setValue(Object.keys(unlocked).map(function (s) { return s + ' ' + unlocked[s]; }).join(', '));
    }
    return { habitat: String(r['Habitat'] || '').trim(), power: String(r['Power']).trim(), text: String(r['Text'] || ''),
      effect: String(r['Effect'] || '').trim().toLowerCase(), value: Number(r['Value']) || 0, rule: String(r['Unlock When'] || ''),
      hint: unlockHint_(r['Unlock When'] || ''), unlocked: unlocked };
  });
}

function engineBonus_(engine, squad, effect) {
  return (engine || []).filter(function (p) { return p.effect === effect && p.unlocked[squad]; }).reduce(function (n, p) { return n + p.value; }, 0);
}

// ── MONSTER OF THE MONTH ────────────────────────────────────────────
function monsterMeasure_(goal, from, to, events) {
  const g = String(goal || '').trim().toLowerCase();
  const m = g.match(/^(points|voters|entries|tickets|pitched in)\s*:\s*(?:([a-z_]+)\s+)?(\d+)\s*%?$/);
  if (!m) return null; // a goal staff decide: type Defeated in Status
  const need = Number(m[3]);
  const scores = measureSquads_(m[1] + (m[2] ? ': ' + m[2] : ''), from, to, events);
  let have = SQUAD_KEYS_.reduce(function (n, s) { return n + scores[s]; }, 0);
  if (m[1] === 'pitched in') have = Math.round(have / 4); // the school's share, averaged over the squads
  return { have: have, need: need, pct: m[1] === 'pitched in' };
}

// Built-in art for the seeded Opponents, matched by name, with a defeated version of each.
// Settings › MASCOT_ART = OFF turns it off; an https Image / Defeated Image on the row wins.
const OPPONENT_ART_ = [[/shady/i, 'monster-shady'], [/tardy|troll/i, 'monster-tardy-troll'], [/litter/i, 'monster-litter-beast']];
function opponentArt_(name, defeated) {
  if (String(getSetting_('MASCOT_ART') || 'ON').toUpperCase() === 'OFF') return '';
  const hit = OPPONENT_ART_.filter(function (a) { return a[0].test(String(name || '')); })[0];
  return hit ? ART_BASE_ + hit[1] + (defeated ? '-defeated' : '') + '.jpg' : ''; // ART_BASE_: 16_Me.gs
}

function monster_(events) {
  if (!gameOn_('OPPONENT_VISIBLE')) return null;
  const today = todayYmd_();
  const r = sheetRows_('Opponents').filter(function (x) {
    return ticked_(x['Show']) && String(x['Opponent'] || '').trim() &&
      (!x['Start'] || String(x['Start']).slice(0, 10) <= today) && (!x['End'] || String(x['End']).slice(0, 10) >= today || /^defeated/i.test(String(x['Status'])));
  })[0];
  if (!r) return null;
  const from = String(r['Start'] || '0000-01-01').slice(0, 10), to = String(r['End'] || '9999-12-31').slice(0, 10);
  const prog = monsterMeasure_(r['Defeat Goal'], from, to, events || readEventsTab_());
  const defeated = /^defeated/i.test(String(r['Status'] || '')) || (prog && prog.have >= prog.need);
  const https = function (v) { v = String(v || '').trim(); return /^https:\/\//i.test(v) ? v : ''; };
  // Defeated: the row's Defeated Image, else the built-in defeated art, else the normal picture.
  const image = defeated ? https(r['Defeated Image']) || opponentArt_(r['Opponent'], true) || https(r['Image'])
    : https(r['Image']) || opponentArt_(r['Opponent'], false);
  return { row: r._row, name: String(r['Opponent']).trim(), story: String(r['Story'] || ''), image: image,
    goal: String(r['Defeat Goal'] || ''), progress: prog, health: defeated ? 0 : prog ? Math.max(0, 1 - prog.have / Math.max(1, prog.need)) : 1,
    defeated: !!defeated, reward: Number(r['Reward']) || 0, end: to, status: String(r['Status'] || '') };
}

function menuToggleMonster_() {
  const on = !gameOn_('OPPONENT_VISIBLE');
  setSettingValue_('OPPONENT_VISIBLE', on ? 'ON' : 'OFF');
  SpreadsheetApp.getUi().alert('The Opponent is now ' + (on ? 'VISIBLE (rows with Show ticked, inside their dates).' : 'HIDDEN everywhere.'));
}

// ── EVERYTHING THE PAGES NEED ───────────────────────────────────────
function gameState_(events) {
  events = events || readEventsTab_();
  const zones = gameOn_('GAME_MAP') ? mapZones_(events) : [];
  const engine = gameOn_('GAME_ENGINE') ? enginePowers_(false, events) : [];
  return {
    map: mapPayload_(events, zones),
    streaks: gameOn_('GAME_STREAK') ? (function () { const s = hotStreaks_(); Object.keys(s).forEach(function (k) { s[k].bank = streakValue_(s[k].length, k, engine); }); return s; })() : null,
    rounds: gameOn_('GAME_ROUNDS') ? roundGoals_(events).filter(function (g) { return g.active || (g.done && !/awarded/i.test(g.status)); }).map(function (g) { delete g.row; return g; }) : [],
    engine: engine,
    monster: monster_(events),
  };
}

function clashGame() { return cachedGame_(); }

// The game data (map, streaks, goals, engine, the Opponent), cached until the Sheet changes (03_Web.gs).
function cachedGame_(fresh) {
  const key = 'game-' + dataStamp_();
  if (!fresh) { const hit = cacheGetBig_(key); if (hit) return hit; }
  const g = gameState_(); g.studentLink = String(getSetting_('STUDENT_LINK') || '');
  const text = JSON.stringify(g);
  cachePutBig_(key, text, 300);
  return text;
}

// ── THE TICK: claims, unlocks, and (with GAME_AWARDS ON) the weekly and end-of-round points ──
function gameTick() {
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  noteRun_('gameTick'); // 22_Week.gs
  try {
    ensureGameTabs_();
    const events = readEventsTab_(), lines = [];
    const zones = gameOn_('GAME_MAP') ? mapZones_(events, true) : [];
    const engine = gameOn_('GAME_ENGINE') ? enginePowers_(true, events) : [];
    if (!awardsOn_()) return ['GAME_AWARDS is OFF: claims and unlocks updated, no points sent.'];
    const props = PropertiesService.getScriptProperties();
    const wk = weekStart_();
    // Weekly territory + engine payout, once per week, for the week that just ended.
    if (props.getProperty('game-paid-week') !== wk) {
      const lastWk = ymdShift_(wk, -7), lastEnd = ymdShift_(wk, -1);
      const wings = heldWings_(zones), mult = Math.max(1, Number(getSetting_('WING_MULTIPLIER')) || 2);
      const votes = measureSquads_('voters', lastWk, lastEnd, events);
      const hand = sheetRows_('Squad Hand');
      SQUAD_KEYS_.forEach(function (s) {
        const mine = zones.filter(function (z) { return z.holder === s; });
        // Each zone pays its Weekly Points x its rule's Multiplier (how the data is captured) x 2 on the Summit.
        let pts = Math.round(mine.reduce(function (n, z) { return n + z.weekly * (z.mult || 1) * (z.summit ? 2 : 1); }, 0));
        pts += mine.length * engineBonus_(engine, s, 'rent');
        Object.keys(wings).forEach(function (w) { if (wings[w] === s) pts += zones.filter(function (z) { return z.wing === w; }).reduce(function (n, z) { return n + z.weekly; }, 0) * (mult - 1); });
        const armed = hand.filter(function (h) { return squadOf_(h['Squad']) === s && String(h['Effect']).toLowerCase() === 'double_rent' && String(h['Status']) === 'Armed'; })[0];
        if (armed && pts) { pts *= 2; getSheet_('Squad Hand').getRange(armed._row, HAND_HEADERS_.indexOf('Status') + 1).setValue('Used ' + todayYmd_()); }
        const pollPower = engineBonus_(engine, s, 'poll');
        if (pollPower) pts += Math.floor(votes[s] / 10) * pollPower;
        if (pts > 0) {
          const o = {}; o[s] = pts;
          const wingNames = Object.keys(wings).filter(function (w) { return wings[w] === s; });
          addPointsEvent_('Territory: ' + mine.length + ' zone' + (mine.length === 1 ? '' : 's') + (wingNames.length ? ' + ' + wingNames.join(', ') + ' wing' : ''), 'territory',
            'Week of ' + lastWk + (armed ? ' · Summit Fever x2' : ''), o);
          lines.push(s + ' +' + pts + ' territory');
        }
      });
      props.setProperty('game-paid-week', wk);
    }
    // Round goals that ended.
    const rsheet = getSheetSafe_('Round Goals');
    roundGoals_(events).forEach(function (g) {
      if (!g.done || /awarded/i.test(g.status)) return;
      const o = {}, res = [];
      SQUAD_KEYS_.forEach(function (s) { const p = g.places[s]; if (p) { o[s] = g.prizes[p - 1]; res.push(['1st', '2nd', '3rd'][p - 1] + ': ' + s); } });
      addPointsEvent_('Round goal: ' + g.goal, 'round_goal', res.join(', '), o);
      if (rsheet) rsheet.getRange(g.row, ROUND_HEADERS_.indexOf('Status') + 1, 1, 2).setValues([['Awarded ' + todayYmd_(), res.join('; ')]]);
      lines.push('Round goal ' + g.goal + ': ' + res.join(', '));
    });
    // Evolution boosts (19_Evolve.gs), once a week.
    Array.prototype.push.apply(lines, payBoosts_());
    // Beat our own marking period (22_Week.gs), once each marking period ends.
    Array.prototype.push.apply(lines, payBeatMp_(events));
    // The monster, once defeated.
    const mon = monster_(events);
    if (mon && mon.defeated && !/defeated/i.test(mon.status)) {
      const o = {}; SQUAD_KEYS_.forEach(function (s) { o[s] = mon.reward; });
      addPointsEvent_('Opponent defeated: ' + mon.name, 'monster', 'The whole school, together', o);
      getSheet_('Opponents').getRange(mon.row, MONSTER_HEADERS_.indexOf('Status') + 1).setValue('Defeated ' + todayYmd_());
      lines.push(mon.name + ' defeated');
    }
    return lines;
  } finally { lock.releaseLock(); }
}

// At most every 10 minutes, from page loads, so the map and unlocks stay current.
function maybeGameTick_() {
  const cache = CacheService.getScriptCache();
  if (cache.get('game-tick')) return;
  cache.put('game-tick', '1', 600);
  try { gameTick(); } catch (e) { console.warn('Game tick: ' + e.message); }
}

function menuTurnOnGame_() {
  setSettingValue_('GAME_AWARDS', 'ON');
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'gameTick') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('gameTick').timeBased().everyHours(1).create();
  const lines = gameTick();
  SpreadsheetApp.getUi().alert('Game awards are ON (checked every hour): weekly territory, round goals, the Opponent.\n\n' + (lines.length ? lines.join('\n') : 'Nothing to award yet.'));
}

function menuTurnOffGame_() {
  setSettingValue_('GAME_AWARDS', 'OFF');
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'gameTick') ScriptApp.deleteTrigger(t); });
  SpreadsheetApp.getUi().alert('Game awards are OFF. The map, cards and goals still show; no bonus points go out.');
}

// ── GENERALS HQ: the squad's game hand ──────────────────────────────
function hqGame_(squad) {
  try {
    const events = readEventsTab_(), engine = gameOn_('GAME_ENGINE') ? enginePowers_(false, events) : [];
    const st = gameOn_('GAME_STREAK') ? hotStreaks_()[squad] : null;
    return {
      cards: gameOn_('GAME_CARDS'),
      hand: gameOn_('GAME_CARDS') ? handFor_(squad).filter(function (h) { return h.status === 'Held' || h.status === 'Armed'; })
        .map(function (h) { const d = chanceDeck_().filter(function (c) { return c.card === h.card; })[0]; return { row: h.row, card: h.card, effect: h.effect, value: h.value, status: h.status, text: d ? d.text : '' }; }) : [],
      drawsLeft: gameOn_('GAME_CARDS') ? drawsLeft_(squad, engine) : 0,
      streak: st ? { length: st.length, bank: streakValue_(st.length, squad, engine), shields: st.shields } : null,
      engine: engine,
      map: mapPayload_(events),
      monster: monster_(events),
    };
  } catch (e) { return { error: e.message }; }
}
