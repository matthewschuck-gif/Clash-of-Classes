/**
 * Clash of Classes -- Apps Script backend
 * File 19: Evolution. Each student's creature (their squad's animal) grows up as they climb,
 * Pokémon-style, and every evolution lets them pick a boost that earns their squad extra points.
 *
 * LRG POINTS (Settings › LRG_POINTS): Life Ready Graduate points a student earns from what they already do, e.g.
 *   "bills=1; mindset=5; shoutouts=2; star=1; streak=2; buddy=3; tickets=3; passport=5".
 *   Things: bills, mindset, shoutouts, star, streak, buddy, tickets, passport. LRG points only go up.
 * STAGES (Settings › EVOLVE_LRG "2=30, 3=100", EVOLVE_NAMES "Foothills, Ridge, Summit"): stage 1
 *   from the start, stage 2 and 3 at those LRG points. Art: assets/creatures/{squad}-{1|2|3}[f].jpg, a
 *   male and a female version of every stage. A student picks either and can show any stage
 *   they've reached.
 * EVOLUTION BOOSTS (tab Evolution Boosts): reaching stage 2 lets a student pick one boost at its
 *   Ridge ×; stage 3 adds a second, different one at its Summit ×. They stack and stay.
 *   Each boost watches a Student Board column:
 *     count -- growth in the column since the last payout × Points Per = the student's base points
 *     level -- column at or above Threshold (or Yes) this week = Points Per
 *   Each week the boost adds base × (multiplier − 1) to the student's squad. One "Evolution
 *   boosts" row per squad, no names (Boost Ledger tab keeps who earned what, staff only).
 *   A boost whose column isn't on the Student Board yet isn't offered and pays nothing. Besides
 *   Student Board columns, a boost can count Poll Votes, Tickets Completed or Event Sign-ups
 *   (tracked in this Sheet).
 *   Paid with the other game awards (GAME_AWARDS ON), once a week. Freezes still apply.
 * PERKS (tab Perks): things students unlock and equip, one of each kind at a time:
 *   title  -- words under their name (Value blank = the Perk name)
 *   frame  -- the glow around their creature (Value: a color like #ffe100, squad, or aurora)
 *   scene  -- the background of their hero card (Value: night, sunrise, forest, aurora, or an https image)
 *   reward -- a real-world perk staff honor (Value: what to show staff); nothing to equip
 *   Unlock When: blank (everyone), lrg: N (LRG points), stage: N, boost: any, or bills / mindset / shoutouts /
 *   star / streak / buddy / tickets / passport: N.
 * CLIFF: the Mountaineer, the school mascot, guides every student on My Climb (assets/cliff.png;
 *   Settings › CLIFF_IMAGE overrides it).
 */

const BOOST_HEADERS_ = ['Show', 'Boost', 'What Students See', 'Student Board Column', 'Kind', 'Threshold', 'Points Per', 'Ridge ×', 'Summit ×'];
const LEDGER_HEADERS_ = ['800#', 'Squad', 'Boost', 'Stage', 'Multiplier', 'Picked', 'Baseline', 'Paid Through', 'Bonus Total'];
const PERK_HEADERS_ = ['Show', 'Perk', 'Kind', 'Value', 'What Students See', 'Unlock When'];
const PERK_KINDS_ = ['title', 'frame', 'scene', 'reward'];
// Counts this Sheet keeps itself, usable as a boost's Student Board Column.
const BUILT_IN_COUNTS_ = ['Poll Votes', 'Tickets Completed', 'Event Sign-ups'];

function ensureEvolveTabs_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  [['Evolution Boosts', BOOST_HEADERS_, [
      [true, 'Trait Dough', 'Every Trait Dough bill you turn in counts extra for your squad.', 'Trait Dough Bills', 'count', '', 1, 1.25, 2],
      [true, 'STAR Growth', 'Every STAR growth point you earn counts extra.', 'STAR Growth Points', 'count', '', 1, 1.25, 2],
      [true, 'Attendance', 'Each week your attendance is 95% or better, your squad gets extra.', 'Attendance %', 'level', 95, 4, 1.25, 2],
      [true, 'Mindset', 'Every Mountaineer Mindset summit counts extra.', 'Mindset Summits', 'count', '', 5, 1.25, 2],
      [true, 'Shout-outs', 'Every shout-out you earn counts extra.', 'Shout-outs', 'count', '', 2, 1.25, 2],
      [true, 'Behavior', 'Every week with no office referral counts extra.', 'Referral-Free Weeks', 'count', '', 3, 1.25, 2],
      [true, 'Grades', 'Each week you are passing every class, your squad gets extra.', 'Passing All Classes', 'level', 1, 4, 1.25, 2],
      [true, 'Student of the Month', 'Being named Student of the Month counts extra.', 'Student of the Month', 'count', '', 10, 1.25, 2],
      [true, 'Poll Turnout', 'Every Daily Poll vote counts extra.', 'Poll Votes', 'count', '', 1, 1.25, 2],
      [true, 'Ticket Hunter', 'Every destination ticket you finish counts extra.', 'Tickets Completed', 'count', '', 3, 1.25, 2],
      [true, 'Squad Events', 'Every squad event you sign up for counts extra.', 'Event Sign-ups', 'count', '', 2, 1.25, 2],
      [true, 'Steady Streak', 'Each week your streak is 3 weeks or longer, your squad gets extra.', 'Streak (weeks)', 'level', 3, 3, 1.25, 2],
      [true, 'Ambassador', 'Every new student you buddy up with counts extra.', 'Ambassador Buddies', 'count', '', 3, 1.25, 2]]],
   ['Perks', PERK_HEADERS_, [
      [true, 'Trailhead Rookie', 'title', '', 'Everyone starts here.', ''],
      [true, 'Dough Slinger', 'title', '', 'Turn in 10 Trait Dough bills.', 'bills: 10'],
      [true, 'Summit Seeker', 'title', '', 'Reach 2 Mindset summits.', 'mindset: 2'],
      [true, 'Hype Machine', 'title', '', 'Earn 5 shout-outs.', 'shoutouts: 5'],
      [true, 'Ticket Master', 'title', '', 'Finish 5 destination tickets.', 'tickets: 5'],
      [true, 'Iron Streak', 'title', '', 'Keep a 6-week streak.', 'streak: 6'],
      [true, 'Trail Blazer', 'title', '', 'Evolve to stage 2.', 'stage: 2'],
      [true, 'Peak Legend', 'title', '', 'Evolve to stage 3.', 'stage: 3'],
      [true, 'Squad Glow', 'frame', 'squad', 'Your squad color.', ''],
      [true, 'Gold Rim', 'frame', '#ffe100', 'Reach 60 LRG points.', 'lrg: 60'],
      [true, 'Frost', 'frame', '#9be7ff', 'Evolve to stage 2.', 'stage: 2'],
      [true, 'Aurora', 'frame', 'aurora', 'Evolve to stage 3.', 'stage: 3'],
      [true, 'Night Trail', 'scene', 'night', 'The trail at night.', ''],
      [true, 'Sunrise Ridge', 'scene', 'sunrise', 'Reach 40 LRG points.', 'lrg: 40'],
      [true, 'Pine Forest', 'scene', 'forest', 'Keep a 3-week streak.', 'streak: 3'],
      [true, 'Aurora Summit', 'scene', 'aurora', 'Evolve to stage 3.', 'stage: 3'],
      [false, 'Front of the Lunch Line', 'reward', 'Show this to a lunch duty teacher once this month.', 'A real perk for Summit climbers (staff decide).', 'stage: 3']]],
   ['Boost Ledger', LEDGER_HEADERS_, null]].forEach(function (t) {
    if (ss.getSheetByName(t[0])) return;
    const sh = ss.insertSheet(t[0]);
    sh.getRange(1, 1, 1, t[1].length).setValues([t[1]]).setFontWeight('bold').setBackground('#490e6f').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    if (t[2]) sh.getRange(2, 1, t[2].length, t[1].length).setValues(t[2]);
  });
}

function rulesList_(setting, fallback) {
  const out = {};
  String(getSetting_(setting) || fallback).split(/[;,]/).forEach(function (p) {
    const m = p.trim().match(/^([a-z0-9]+)\s*=\s*(\d+(?:\.\d+)?)$/i);
    if (m) out[m[1].toLowerCase()] = Number(m[2]);
  });
  return out;
}

function climbXp_(vals) {
  const w = rulesList_('LRG_POINTS', 'bills=1; mindset=5; shoutouts=2; star=1; streak=2; buddy=3; tickets=3; passport=5');
  return Math.round(Object.keys(w).reduce(function (n, k) { return n + (Number(vals[k]) || 0) * w[k]; }, 0));
}

function evolveAt_() {
  const t = rulesList_('EVOLVE_LRG', '2=30, 3=100');
  return { 2: t['2'] || 30, 3: Math.max(t['3'] || 100, (t['2'] || 30) + 1) };
}

function stageNames_() {
  const n = String(getSetting_('EVOLVE_NAMES') || 'Foothills, Ridge, Summit').split(',').map(function (s) { return s.trim(); });
  return [n[0] || 'Foothills', n[1] || 'Ridge', n[2] || 'Summit'];
}

function stageFor_(xp) { const at = evolveAt_(); return xp >= at[3] ? 3 : xp >= at[2] ? 2 : 1; }

function creatureArt_(squad, stage, form) {
  return ART_BASE_ + 'creatures/' + String(squad).toLowerCase() + '-' + stage + (form === 'f' ? 'f' : '') + '.jpg';
}

function boardHeader_() {
  const sh = getSheetSafe_(getSetting_('STUDENT_BOARD_TAB') || 'Student Board');
  if (!sh || !sh.getLastColumn()) return [];
  return sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(function (h) { return String(h).trim(); });
}

// The boosts on the tab; ready = its Student Board column exists.
function boostDefs_() {
  const head = boardHeader_();
  return sheetRows_('Evolution Boosts').filter(function (r) { return ticked_(r['Show']) && String(r['Boost'] || '').trim(); }).map(function (r) {
    const col = String(r['Student Board Column'] || '').trim();
    return { boost: String(r['Boost']).trim(), text: String(r['What Students See'] || ''), column: col,
      kind: String(r['Kind'] || 'count').trim().toLowerCase() === 'level' ? 'level' : 'count',
      threshold: Number(r['Threshold']) || 0, per: Number(r['Points Per']) || 0,
      mult: { 2: Number(r['Ridge ×']) || 0, 3: Number(r['Summit ×']) || 0 }, ready: !!col && (head.indexOf(col) > -1 || BUILT_IN_COUNTS_.indexOf(col) > -1) };
  });
}

function lookRow_(id) { return sheetRows_('Avatars').filter(function (a) { return id800_(a['800#']) === id; })[0]; }

function saveLook_(id, look) {
  const sheet = getSheet_('Avatars'), r = lookRow_(id);
  const row = [id, JSON.stringify(look), todayYmd_()];
  if (r) sheet.getRange(r._row, 1, 1, 3).setValues([row]); else sheet.appendRow(row);
}

// A student's number for a boost: a Student Board column, or a count this Sheet keeps.
function studentValue_(id, row, col) {
  if (col === 'Poll Votes') return sheetRows_('Poll Votes').filter(function (v) { return id800_(v['800#']) === id; }).length;
  if (col === 'Event Sign-ups') return sheetRows_('Sign-ups').filter(function (v) { return id800_(v['800#']) === id; }).length;
  if (col === 'Tickets Completed') return sheetRows_('Student Tickets').filter(function (t) { return id800_(t['800#']) === id && String(t['Status']) === 'Completed'; }).length;
  return boardNumber_(row[col]);
}

function boardNumber_(v) {
  const s = String(v == null ? '' : v).trim();
  if (/^(yes|y|true|✓|x)$/i.test(s)) return 1;
  return Number(s) || 0;
}

// Everything My Climb shows about the creature.
function creatureFor_(id, squad, vals, look) {
  look = look || {};
  const xp = climbXp_(vals), stage = stageFor_(xp), at = evolveAt_(), names = stageNames_();
  const form = look.form === 'f' ? 'f' : 'm';
  const show = Math.min(stage, Math.max(1, Number(look.show) || stage));
  const defs = boostDefs_();
  const ledger = sheetRows_('Boost Ledger').filter(function (b) { return id800_(b['800#']) === id; });
  const picked = [2, 3].map(function (s) {
    const b = ledger.filter(function (x) { return Number(x['Stage']) === s; })[0];
    if (!b) return null;
    const d = defs.filter(function (x) { return x.boost === String(b['Boost']); })[0] || {};
    return { stage: s, boost: String(b['Boost']), text: d.text || '', mult: Number(b['Multiplier']) || 0, bonus: Math.round(10 * (Number(b['Bonus Total']) || 0)) / 10 };
  }).filter(Boolean);
  const takenCols = picked.map(function (p) { const d = defs.filter(function (x) { return x.boost === p.boost; })[0]; return d ? d.column : ''; });
  const choices = {};
  [2, 3].forEach(function (s) {
    if (stage < s || picked.some(function (p) { return p.stage === s; })) return;
    choices[s] = defs.filter(function (d) { return d.ready && d.mult[s] > 0 && takenCols.indexOf(d.column) === -1; })
      .map(function (d) { return { boost: d.boost, text: d.text, mult: d.mult[s] }; });
  });
  return {
    xp: xp, stage: stage, show: show, form: form, names: names, at: at,
    next: stage < 3 ? { stage: stage + 1, name: names[stage], need: at[stage + 1] - xp, at: at[stage + 1] } : null,
    from: stage === 1 ? 0 : at[stage],
    art: [1, 2, 3].map(function (s) { return { stage: s, name: names[s - 1], m: creatureArt_(squad, s, 'm'), f: creatureArt_(squad, s, 'f'), open: s <= stage }; }),
    boosts: picked, choices: choices, perks: perksFor_(Object.assign({ lrg: xp, xp: xp, stage: stage, boost: picked.length }, vals), look),
  };
}

// ── PERKS ───────────────────────────────────────────────────────────
function perkUnlocked_(rule, vals) {
  const r = String(rule || '').trim().toLowerCase();
  if (!r) return true;
  const m = r.match(/^([a-z]+)\s*:\s*(any|\d+)$/);
  if (!m) return false; // a rule we can't read stays locked
  return m[2] === 'any' ? (Number(vals[m[1]]) || 0) > 0 : (Number(vals[m[1]]) || 0) >= Number(m[2]);
}

function perkDefs_() {
  return sheetRows_('Perks').filter(function (r) { return ticked_(r['Show']) && String(r['Perk'] || '').trim(); }).map(function (r) {
    const kind = String(r['Kind'] || '').trim().toLowerCase();
    return { perk: String(r['Perk']).trim(), kind: PERK_KINDS_.indexOf(kind) > -1 ? kind : 'title', value: String(r['Value'] || '').trim(),
      text: String(r['What Students See'] || ''), rule: String(r['Unlock When'] || '').trim() };
  });
}

// { list: [{ perk, kind, value, text, rule, open }], equipped: { title, frame, scene } }
function perksFor_(vals, look) {
  const list = perkDefs_().map(function (p) { return Object.assign(p, { open: perkUnlocked_(p.rule, vals) }); });
  const equipped = {};
  ['title', 'frame', 'scene'].forEach(function (k) {
    const want = (look && look[k]) || '';
    const ok = list.filter(function (p) { return p.kind === k && p.open && p.perk === want; })[0];
    if (ok) equipped[k] = ok.perk;
  });
  return { list: list, equipped: equipped };
}

function cliffInfo_() {
  const img = String(getSetting_('CLIFF_IMAGE') || '').trim();
  return { name: String(getSetting_('CLIFF_NAME') || 'Cliff').trim(), image: /^https:\/\//i.test(img) ? img : ART_BASE_ + 'cliff.png' };
}

// ── STUDENT ACTIONS ─────────────────────────────────────────────────
function climbVals_(id, page) {
  const me = page.me;
  return { bills: me.bills, mindset: me.mindset, shoutouts: me.shoutouts, star: me.star, streak: me.streak, buddy: me.ambassador,
    tickets: (page.tickets || []).filter(function (t) { return t.status === 'Completed'; }).length,
    passport: page.passport && page.passport.complete ? 1 : 0 };
}

// form: 'm' | 'f'; show: which reached stage to show; perk: equip (or take off) an unlocked perk.
function clashMeSaveCreature(choice, asId, adminToken) {
  const id = meId_(asId, adminToken);
  choice = choice || {};
  let perk = null;
  if (choice.perk) {
    perk = JSON.parse(clashMe(asId, adminToken)).creature.perks.list.filter(function (p) { return p.perk === choice.perk && p.kind !== 'reward'; })[0];
    if (!perk) throw new Error('That perk isn\'t available.');
    if (!perk.open) throw new Error('You haven\'t unlocked ' + perk.perk + ' yet.');
  }
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const r = lookRow_(id);
    let look = {}; try { look = r ? JSON.parse(String(r['Look'] || '{}')) : {}; } catch (e) { look = {}; }
    if (choice.form === 'm' || choice.form === 'f') look.form = choice.form;
    if (choice.show) look.show = Math.max(1, Math.min(3, Math.floor(Number(choice.show)) || 1));
    if (perk) look[perk.kind] = look[perk.kind] === perk.perk ? '' : perk.perk; // pick it again to take it off
    saveLook_(id, look);
  } finally { lock.releaseLock(); }
  return clashMe(asId, adminToken);
}

function clashMePickBoost(stage, boost, asId, adminToken) {
  const id = meId_(asId, adminToken);
  stage = Number(stage);
  const page = JSON.parse(clashMe(asId, adminToken));
  const c = page.creature;
  if (!c.choices[stage]) throw new Error(c.stage < stage ? 'Your creature hasn\'t reached that stage yet.' : 'You already picked that boost.');
  const pick = c.choices[stage].filter(function (x) { return x.boost === boost; })[0];
  if (!pick) throw new Error('That boost isn\'t one of your choices.');
  const def = boostDefs_().filter(function (d) { return d.boost === boost; })[0];
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const row = boardRows_().filter(function (r) { return id800_(r['800#']) === id; })[0] || {};
    const base = def.kind === 'count' ? studentValue_(id, row, def.column) : '';
    getSheet_('Boost Ledger').appendRow([id, page.me.squad, boost, stage, pick.mult, todayYmd_(), base, '', 0]);
  } finally { lock.releaseLock(); }
  return clashMe(asId, adminToken);
}

// ── THE WEEKLY PAYOUT (from gameTick, GAME_AWARDS ON) ───────────────
function payBoosts_() {
  const props = PropertiesService.getScriptProperties(), wk = weekStart_();
  if (props.getProperty('boost-paid-week') === wk) return [];
  const sheet = getSheetSafe_('Boost Ledger');
  props.setProperty('boost-paid-week', wk);
  if (!sheet) return [];
  const defs = {}; boostDefs_().forEach(function (d) { defs[d.boost] = d; });
  const board = {}; boardRows_().forEach(function (r) { board[id800_(r['800#'])] = r; });
  const head = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
  const col = function (h) { return head.indexOf(h) + 1; };
  const sum = { BLACK: 0, GOLD: 0, GREY: 0, PURPLE: 0 }, who = { BLACK: {}, GOLD: {}, GREY: {}, PURPLE: {} };
  const today = todayYmd_();
  sheetRows_('Boost Ledger').forEach(function (b) {
    const d = defs[String(b['Boost'])], row = board[id800_(b['800#'])];
    if (!d || !d.ready || !row) return;
    const squad = squadOf_(row['Squad']) || squadOf_(b['Squad']);
    if (sum[squad] === undefined) return;
    const cur = studentValue_(id800_(b['800#']), row, d.column);
    let base = 0;
    if (d.kind === 'count') {
      base = Math.max(0, cur - boardNumber_(b['Baseline'])) * d.per;
      sheet.getRange(b._row, col('Baseline')).setValue(cur);
    } else if (cur >= (d.threshold || 1)) base = d.per;
    const extra = base * Math.max(0, (Number(b['Multiplier']) || 1) - 1);
    sheet.getRange(b._row, col('Paid Through'), 1, 2).setValues([[today, Math.round(100 * ((Number(b['Bonus Total']) || 0) + extra)) / 100]]);
    if (extra > 0) { sum[squad] += extra; who[squad][id800_(b['800#'])] = 1; }
  });
  const lines = [];
  SQUAD_KEYS_.forEach(function (s) {
    const pts = Math.round(sum[s]);
    if (pts <= 0) return;
    const n = Object.keys(who[s]).length, o = {}; o[s] = pts;
    addPointsEvent_('Evolution boosts', 'evolution_boost', n + ' climber' + (n === 1 ? '' : 's') + ', week of ' + ymdShift_(wk, -7), o);
    lines.push(s + ' +' + pts + ' evolution boosts');
  });
  return lines;
}

// ── THE "HOW MY CLIMB WORKS" TAB ────────────────────────────────────
// Staff guide to creatures, LRG points, evolution boosts and perks. Rewritten every Run setup.
const CLIMB_GUIDE_TAB_ = 'How My Climb Works';

function setupClimbGuideTab_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CLIMB_GUIDE_TAB_);
  if (!sheet) sheet = ss.insertSheet(CLIMB_GUIDE_TAB_);
  else sheet.clear();
  const H = '## ';
  const rows = [
    'Clash of Classes -- How My Climb Works (creatures, LRG points, evolution boosts, perks)',
    '',
    'Every student has their own creature: their squad\'s animal (Gold cottontail rabbit, Purple ruffled grouse,',
    'Grey fisher, Black white-tailed deer), in a male or female version they pick. It evolves as they earn',
    'LRG (Life Ready Graduate) points. Cliff, the Mountaineer, guides them through it on My Climb.',
    '',
    H + 'LRG POINTS',
    'Settings -> LRG_POINTS sets what each thing a student does is worth, e.g.',
    '     bills=1; mindset=5; shoutouts=2; star=1; streak=2; buddy=3; tickets=3; passport=5',
    '     bills = Trait Dough bills, mindset = Mindset summits, shoutouts, star = STAR growth points,',
    '     streak = weeks in a row, buddy = Ambassador buddies, tickets = destination tickets finished,',
    '     passport = 1 when the Trail Passport is complete.',
    'The numbers come from the Student Board, so LRG points only go up during the year.',
    '',
    H + 'EVOLVING',
    'Settings -> EVOLVE_LRG (2=30, 3=100) sets the LRG points for stage 2 and stage 3.',
    'Settings -> EVOLVE_NAMES (Foothills, Ridge, Summit) names the three stages.',
    'Students can show any stage they have reached. The first time they open My Climb after evolving,',
    'they see their creature evolve.',
    '',
    H + 'EVOLUTION BOOSTS (THE EVOLUTION BOOSTS TAB)',
    'At stage 2 a student picks ONE boost at its Ridge x. At stage 3 they add a SECOND, different one at its',
    'Summit x. Boosts stack and stay for the year. Each row:',
    '     Show                 -- untick to stop offering it (boosts already picked keep paying).',
    '     Boost                -- the name students see.',
    '     What Students See    -- one sentence on the boost card.',
    '     Student Board Column -- what it watches: any Student Board column, or Poll Votes, Tickets',
    '                             Completed, Event Sign-ups (counted in this Sheet). A column that is not',
    '                             on the Student Board yet (e.g. Referral-Free Weeks) is simply not offered.',
    '     Kind                 -- count: growth in the column since last week. level: the column is at or',
    '                             above Threshold this week (or says Yes).',
    '     Threshold            -- level boosts only (e.g. 95 for attendance %).',
    '     Points Per           -- what one unit is worth to the squad (one bill, one summit, one good week).',
    '     Ridge x / Summit x   -- the multiplier at each stage. Blank = not offered at that stage.',
    'What it pays: each week, (growth x Points Per) x (multiplier - 1) goes to the student\'s squad.',
    'Example: a x1.25 Trait Dough boost and 8 new bills -> 8 x 1 x 0.25 = 2 extra points that week.',
    'All boosts in a squad are added up into ONE "Evolution boosts" row per squad on the events tab, with no',
    'student names. The Boost Ledger tab shows who earned what (staff only). Boost points go out with the',
    'other game awards (Clash of Classes -> Turn on game awards), once a week. Freezes still apply.',
    '',
    H + 'PERKS (THE PERKS TAB)',
    'Things students unlock and wear, one of each kind at a time:',
    '     title  -- words under their name. Value blank = the Perk name.',
    '     frame  -- the glow around their creature. Value: a color like #ffe100, squad, or aurora.',
    '     scene  -- the background of their card. Value: night, sunrise, forest, aurora, or an https image.',
    '     reward -- a real-world perk staff honor (Value = what the student shows staff). Ships turned off.',
    'Unlock When: blank (everyone), lrg: N, stage: N, boost: any, or bills / mindset / shoutouts / star /',
    'streak / buddy / tickets / passport: N. Add a row to add a perk; untick Show to hide one.',
    '',
    H + 'CLIFF',
    'Cliff is on every My Climb page with the most useful tip for that student. Settings -> CLIFF_NAME renames',
    'him; CLIFF_IMAGE (an https link) swaps his picture.',
  ];
  const values = rows.map(function (r) { return [r.indexOf(H) === 0 ? r.slice(H.length) : r]; });
  sheet.getRange(1, 1, values.length, 1).setValues(values);
  sheet.getRange(1, 1).setFontWeight('bold').setFontSize(14).setFontColor('#490e6f');
  rows.forEach(function (r, i) { if (r.indexOf(H) === 0) sheet.getRange(i + 1, 1).setFontWeight('bold').setFontColor('#490e6f'); });
  sheet.setColumnWidth(1, 900);
  sheet.setFrozenRows(1);
}
