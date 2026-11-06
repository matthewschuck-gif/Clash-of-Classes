/**
 * Clash of Classes -- Apps Script backend
 * File 21: Leaderboards (the site's Leaderboards page).
 *
 * SQUAD LEADERBOARDS (always): the four squads ranked on each thing students do -- Trait Dough
 * bills, Mindset summits, STAR growth, shout-outs, Ambassador buddies, Trail Journal growth (when
 * Define a Climb-er shows it), average LRG points, creatures evolved, who's pitching in (anything
 * in the last PITCHED_IN_DAYS) and points this month. Sums from the Student Board; no names.
 * TOP CLIMBERS (Settings › LEADERBOARD_NAMES, OFF to start): the top LEADERBOARD_SIZE students per
 * grade by LRG points, as "First L." with their squad and creature stage. Positive numbers only.
 * Cached until the Sheet changes, like the rest of the public data.
 */

const BOARD_CATS_ = [
  ['bills', 'Trait Dough bills', 'Trait Dough Bills', '\u{1F4B5}'],
  ['mindset', 'Mindset summits', 'Mindset Summits', '⛰️'],
  ['star', 'STAR growth points', 'STAR Growth Points', '\u{1F4C8}'],
  ['shoutouts', 'Shout-outs', 'Shout-outs', '\u{1F64C}'],
  ['buddy', 'Ambassador buddies', 'Ambassador Buddies', '\u{1F91D}'],
  ['journal', 'Trail Journal growth', 'Trail Journal Growth', '\u{1F3D4}️'],
];

function leadersData_() {
  const rows = boardRows_().filter(function (r) { return SQUAD_KEYS_.indexOf(squadOf_(r['Squad'])) > -1; });
  const today = todayYmd_(), since = ymdShift_(today, -Math.max(1, Number(getSetting_('PITCHED_IN_DAYS')) || 30));
  const zero = function () { return { BLACK: 0, GOLD: 0, GREY: 0, PURPLE: 0 }; };
  const size = zero(), lrg = zero(), evolved = zero(), pitched = zero(), cats = {};
  BOARD_CATS_.forEach(function (c) { cats[c[0]] = zero(); });
  const people = [];
  rows.forEach(function (r) {
    const s = squadOf_(r['Squad']);
    size[s]++;
    const vals = { bills: boardNumber_(r['Trait Dough Bills']), mindset: boardNumber_(r['Mindset Summits']), star: boardNumber_(r['STAR Growth Points']),
      shoutouts: boardNumber_(r['Shout-outs']), buddy: boardNumber_(r['Ambassador Buddies']), streak: boardNumber_(r['Streak (weeks)']),
      journal: boardNumber_(r['Trail Journal Growth']) };
    BOARD_CATS_.forEach(function (c) { cats[c[0]][s] += Math.max(0, vals[c[0]] || 0); });
    const xp = climbXp_(vals), stage = stageFor_(xp);
    lrg[s] += xp;
    if (stage >= 2) evolved[s]++;
    if (String(r['Last Contribution'] || '').slice(0, 10) >= since) pitched[s]++;
    people.push({ name: String(r['Name'] || '').trim(), squad: s, grade: String(r['Grade'] || '').trim(), lrg: xp, stage: stage });
  });
  const avg = zero(), share = zero();
  SQUAD_KEYS_.forEach(function (s) {
    avg[s] = size[s] ? Math.round(10 * lrg[s] / size[s]) / 10 : 0;
    share[s] = size[s] ? Math.round(100 * pitched[s] / size[s]) : 0;
  });
  const m = monthBounds_(0);
  const month = pointsBy_(readEventsTab_(), m[0], today, null).points;
  const boards = [
    { key: 'month', label: 'Points this month', icon: '\u{1F3C6}', scores: month },
    { key: 'lrg', label: 'Average LRG points', icon: '⭐', scores: avg, note: 'per student' },
    { key: 'pitched', label: 'Pitching in', icon: '\u{1F64B}', scores: share, unit: '%', note: 'did something in the last ' + (Number(getSetting_('PITCHED_IN_DAYS')) || 30) + ' days' },
    { key: 'evolved', label: 'Creatures evolved', icon: '\u{1F43E}', scores: evolved, note: 'reached stage 2 or higher' },
  ].concat(BOARD_CATS_.map(function (c) { return { key: c[0], label: c[1], icon: c[3], scores: cats[c[0]] }; }))
    .filter(function (b) { return SQUAD_KEYS_.some(function (s) { return b.scores[s] > 0; }) || b.key === 'month'; });
  const out = { boards: boards, names: String(getSetting_('LEADERBOARD_NAMES') || 'OFF').toUpperCase() === 'ON' };
  if (out.names) {
    const n = Math.max(3, Math.min(25, Number(getSetting_('LEADERBOARD_SIZE')) || 10)), names = stageNames_();
    const top = function (g) {
      return people.filter(function (p) { return p.grade === g && p.lrg > 0 && p.name; }).sort(function (a, b) { return b.lrg - a.lrg; }).slice(0, n)
        .map(function (p) { return { name: shortName_(p.name), squad: p.squad, lrg: p.lrg, stage: names[p.stage - 1] }; });
    };
    out.top = { '7': top('7'), '8': top('8') };
  }
  return out;
}

// "Ann L." stays as it is; "Ann Lopez" becomes "Ann L.".
function shortName_(name) {
  const parts = String(name).trim().split(/\s+/);
  return parts.length > 1 ? parts[0] + ' ' + parts[parts.length - 1].charAt(0) + '.' : parts[0];
}

function clashLeaders() {
  const key = 'leaders-' + dataStamp_();
  const hit = cacheGetBig_(key);
  if (hit) return hit;
  const text = JSON.stringify(leadersData_());
  cachePutBig_(key, text, 300);
  return text;
}

// ── TRAIT DOUGH LEADERBOARD (the Trait Dough page) ───────────────────
// Squads on Trait Dough bills turned in this year, this week, per student and drawing wins, from the
// Student Board. Top earners by grade ("First L.") only when Settings › LEADERBOARD_NAMES is ON.
function traitDoughBoard_() {
  const rows = boardRows_().filter(function (r) { return SQUAD_KEYS_.indexOf(squadOf_(r['Squad'])) > -1; });
  const zero = function () { return { BLACK: 0, GOLD: 0, GREY: 0, PURPLE: 0 }; };
  const year = zero(), week = zero(), wins = zero(), size = zero(), earners = zero(), people = [];
  rows.forEach(function (r) {
    const s = squadOf_(r['Squad']), b = Math.max(0, boardNumber_(r['Trait Dough Bills']) || 0);
    size[s]++;
    year[s] += b;
    week[s] += Math.max(0, boardNumber_(r['Bills This Week']) || 0);
    wins[s] += Math.max(0, boardNumber_(r['Trait Dough Wins']) || 0);
    if (b > 0) earners[s]++;
    people.push({ name: String(r['Name'] || '').trim(), squad: s, grade: String(r['Grade'] || '').trim(), bills: b });
  });
  const per = zero(), share = zero();
  SQUAD_KEYS_.forEach(function (s) {
    per[s] = size[s] ? Math.round(100 * year[s] / size[s]) / 100 : 0;
    share[s] = size[s] ? Math.round(100 * earners[s] / size[s]) : 0;
  });
  const out = {
    boards: [
      { key: 'year', label: 'Bills this year', icon: '\u{1F4B5}', scores: year },
      { key: 'week', label: 'Bills this week', icon: '\u{1F4C5}', scores: week },
      { key: 'per', label: 'Bills per student', icon: '⚖️', scores: per, note: 'fair across squad sizes' },
      { key: 'share', label: 'Students earning', icon: '\u{1F64B}', scores: share, unit: '%', note: 'have turned in at least one bill' },
      { key: 'wins', label: 'Friday drawing wins', icon: '\u{1F3A1}', scores: wins },
    ],
    total: SQUAD_KEYS_.reduce(function (a, s) { return a + year[s]; }, 0),
    names: String(getSetting_('LEADERBOARD_NAMES') || 'OFF').toUpperCase() === 'ON',
  };
  if (out.names) {
    const n = Math.max(3, Math.min(25, Number(getSetting_('LEADERBOARD_SIZE')) || 10));
    const top = function (g) {
      return people.filter(function (p) { return p.grade === g && p.bills > 0 && p.name; }).sort(function (a, b) { return b.bills - a.bills; }).slice(0, n)
        .map(function (p) { return { name: shortName_(p.name), squad: p.squad, bills: p.bills }; });
    };
    out.top = { '7': top('7'), '8': top('8') };
  }
  return out;
}

function clashTraitDough() {
  const key = 'traitdough-' + dataStamp_();
  const hit = cacheGetBig_(key);
  if (hit) return hit;
  const text = JSON.stringify(traitDoughBoard_());
  cachePutBig_(key, text, 300);
  return text;
}
