/**
 * Clash of Classes -- Apps Script backend
 * File 4: the student side -- My Squad page, automatic squad challenges, TV feed.
 *
 * MY SQUAD: a student types their 800# and sees their own row of the "Student Board" tab
 * (written hourly by Define a Climb-er's squad sync: Trait Dough, Mindset summits, STAR
 * growth, shout-outs, streak, badges -- positive things only) plus where their squad stands.
 * Only that one row is ever sent to the page. Lookups are capped (MY_SQUAD_LOOKUPS_PER_10MIN)
 * so nobody can page through every 800#.
 *
 * CHALLENGES tab (one row per challenge, edited in the Sheet):
 *   Show | Title | Description | Event Types | Measure | Goal | Start | End | Bonus Points | Awarded
 *   Event Types: Clash event types that count (e.g. trait_dough), blank = every event.
 *   Measure: Points (sum of the squad's points) or Entries (events where the squad scored).
 *   Progress fills in by itself from the events tab between Start and End.
 *   Bonus Points: when CHALLENGE_BONUSES is ON, a squad that reaches the Goal gets one
 *   "Challenge: <Title>" event (type challenge_bonus) once, noted in Awarded. Checked hourly
 *   by the trigger made from the menu.
 */

const SQUAD_KEYS_ = ['BLACK', 'GOLD', 'GREY', 'PURPLE'];
const CHALLENGE_HEADERS_ = ['Show', 'Title', 'Description', 'Event Types', 'Measure', 'Goal', 'Start', 'End', 'Bonus Points', 'Awarded'];

// ── SQUAD STANDINGS (same math as the site: events + finalized months) ──
function squadTotals_() {
  const totals = { BLACK: 0, GOLD: 0, GREY: 0, PURPLE: 0 };
  readEventsTab_().forEach(function (e) { SQUAD_KEYS_.forEach(function (s) { totals[s] += Number(e.pts[s]) || 0; }); });
  let months = [];
  try { months = (loadAppData_('coc').value || {}).monthlyMetrics || []; } catch (err) { /* no data yet */ }
  months.forEach(function (m) { SQUAD_KEYS_.forEach(function (s) { totals[s] += Number(m.pts && m.pts[s]) || 0; }); });
  return totals;
}

function standingFor_(squad, totals) {
  const order = SQUAD_KEYS_.slice().sort(function (a, b) { return totals[b] - totals[a]; });
  const place = order.indexOf(squad);
  if (place < 0) return null;
  const ahead = place > 0 ? order[place - 1] : '';
  return {
    squad: squad, place: place + 1, total: totals[squad],
    ahead: ahead, behindBy: ahead ? totals[ahead] - totals[squad] : 0,
    leadBy: place === 0 ? totals[squad] - totals[order[1]] : 0,
    frozen: freezeDetail_(squad), // the freeze and its thaw progress while the squad can't earn points (14_Thaw.gs)
  };
}

// ── MY SQUAD ────────────────────────────────────────────────────────
function clashMyPage(id) {
  const cache = CacheService.getScriptCache();
  const used = Number(cache.get('clash-lookups')) || 0;
  if (used >= (Number(getSetting_('MY_SQUAD_LOOKUPS_PER_10MIN')) || 300)) throw new Error('The site is busy. Try again in a few minutes.');
  cache.put('clash-lookups', String(used + 1), 600);

  const want = String(id || '').replace(/\D/g, '');
  if (want.length < 5) throw new Error('Type your 800 number.');
  const sheet = getSheetSafe_(getSetting_('STUDENT_BOARD_TAB') || 'Student Board');
  if (!sheet || sheet.getLastRow() < 2) throw new Error('My Squad isn\'t set up yet.');
  const data = sheet.getDataRange().getValues();
  const h = data[0].map(function (x) { return String(x).trim(); });
  const row = data.slice(1).filter(function (r) { return String(r[0]).replace(/\D/g, '') === want; })[0];
  if (!row) throw new Error('That 800 number isn\'t on the list. Check it, or ask your homeroom teacher.');
  const get = function (name) { const i = h.indexOf(name); return i > -1 ? row[i] : ''; };
  const num = function (name) { return Number(get(name)) || 0; };
  const squad = String(get('Squad')).toUpperCase().replace('GRAY', 'GREY');
  const me = {
    name: String(get('Name')), squad: squad,
    bills: num('Trait Dough Bills'), billsWeek: num('Bills This Week'), wins: num('Trait Dough Wins'),
    mindset: num('Mindset Summits'), journal: get('Trail Journal Growth') === '' ? null : num('Trail Journal Growth'),
    star: num('STAR Growth Points'), shoutouts: num('Shout-outs'), ambassador: num('Ambassador Buddies'), streak: num('Streak (weeks)'),
    badges: String(get('Badges')).split(',').map(function (s) { return s.trim(); }).filter(String),
    last: String(get('Last Contribution')), updated: String(get('Updated')),
  };
  const totals = squadTotals_();
  return JSON.stringify({
    me: me,
    standing: standingFor_(squad, totals),
    challenges: challengeProgress_().filter(function (c) { return c.active; }),
    allBadges: badgeNames_(),
    roles: publicRoles_()[want] || [], // Mountaineer Leadership Corps roles (MLC Roster tab)
    theme: squadTheme_(squad),
    lore: loreFor_(squad), // 09_Lore.gs: the chapters this squad has unlocked
    duels: JSON.parse(clashDuels()).filter(function (d) { return d.squads.indexOf(squad) > -1; }),
    passport: passportFor_(want, me),
    signups: mySignups_(want, squad), // 11_MayhemWeek.gs
  });
}

// ── TRAIL PASSPORT ──────────────────────────────────────────────────
// A stamp for each part of the program a student has tried (Settings › PASSPORT_STAMPS,
// "key=label; …"). Keys: bills, mindset, star, shoutouts, ambassador (from the Student Board),
// polls (Daily Poll votes, PASSPORT_POLL_VOTES of them) and events (a squad event sign-up).
// Positive things only: nothing a student gets from a consequence is ever a stamp.
function passportFor_(id, me) {
  const defs = String(getSetting_('PASSPORT_STAMPS') || '').split(';').map(function (p) {
    const i = p.indexOf('=');
    return i > 0 ? { key: p.slice(0, i).trim(), label: p.slice(i + 1).trim() } : null;
  }).filter(function (d) { return d && d.label; });
  if (!defs.length) return null;
  const needVotes = Number(getSetting_('PASSPORT_POLL_VOTES')) || 5;
  let votes = null, signedUp = null;
  const have = function (key) {
    if (key === 'polls') { votes = votes === null ? sheetRows_('Poll Votes').filter(function (v) { return id800_(v['800#']) === id; }).length : votes; return votes >= needVotes; }
    if (key === 'events') { signedUp = signedUp === null ? sheetRows_('Sign-ups').some(function (s) { return id800_(s['800#']) === id; }) : signedUp; return signedUp; }
    return (Number(me[key]) || 0) > 0;
  };
  const stamps = defs.map(function (d) { return { label: d.label, done: have(d.key) }; });
  return { stamps: stamps, complete: stamps.every(function (s) { return s.done; }) };
}

// Every badge name, so the page can show the ones still to earn (from the Settings tab).
function badgeNames_() {
  return String(getSetting_('BADGE_LIST') || '').split(';').map(function (s) { return s.trim(); }).filter(String);
}

// ── CHALLENGES ──────────────────────────────────────────────────────
function ensureChallenges_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Challenges');
  if (sheet) return sheet;
  sheet = ss.insertSheet('Challenges');
  const today = new Date();
  const monday = new Date(today); monday.setDate(today.getDate() - ((today.getDay() + 6) % 7));
  const friday = new Date(monday); friday.setDate(monday.getDate() + 4);
  const ymd = function (d) { return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd'); };
  sheet.getRange(1, 1, 2, CHALLENGE_HEADERS_.length).setValues([CHALLENGE_HEADERS_,
    [false, 'Trait Dough Week', 'Turn in Trait Dough bills all week!', 'trait_dough', 'Entries', 25, ymd(monday), ymd(friday), 50, '']]);
  sheet.getRange(1, 1, 1, CHALLENGE_HEADERS_.length).setFontWeight('bold').setBackground('#490e6f').setFontColor('#ffffff');
  sheet.setFrozenRows(1);
  return sheet;
}

function asYmd_(v) {
  if (Object.prototype.toString.call(v) === '[object Date]') return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return String(v || '').trim();
}

// [{ row, title, description, types, measure, goal, start, end, bonus, awarded: {GOLD: '2026-11-03'}, active, done,
//    progress: {BLACK: n, …} }]
function challengeProgress_() {
  const sheet = getSheetSafe_('Challenges');
  if (!sheet || sheet.getLastRow() < 2) return [];
  const data = sheet.getDataRange().getValues();
  const h = data[0].map(function (x) { return String(x).trim(); });
  const col = function (name) { return h.indexOf(name); };
  const events = readEventsTab_();
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const out = [];
  data.slice(1).forEach(function (r, i) {
    const show = r[col('Show')];
    if (!(show === true || /^(true|yes|y|on)$/i.test(String(show)))) return;
    const title = String(r[col('Title')] || '').trim();
    if (!title) return;
    const types = String(r[col('Event Types')] || '').split(',').map(function (s) { return s.trim(); }).filter(String);
    const entries = /^entr/i.test(String(r[col('Measure')] || ''));
    const start = asYmd_(r[col('Start')]), end = asYmd_(r[col('End')]);
    const progress = { BLACK: 0, GOLD: 0, GREY: 0, PURPLE: 0 };
    events.forEach(function (e) {
      if (types.length && types.indexOf(String(e.type)) === -1) return;
      if (BONUS_TYPES_.indexOf(String(e.type)) > -1) return; // bonuses never count toward a challenge
      const d = String(e.date || '');
      if ((start && d < start) || (end && d > end)) return;
      SQUAD_KEYS_.forEach(function (s) {
        const p = Number(e.pts[s]) || 0;
        if (p) progress[s] += entries ? 1 : p;
      });
    });
    const awarded = {};
    String(r[col('Awarded')] || '').split(',').forEach(function (part) {
      const m = part.trim().match(/^([A-Z]+)\s+(\S+)/);
      if (m) awarded[m[1]] = m[2];
    });
    out.push({
      row: i + 2, title: title, description: String(r[col('Description')] || ''), types: types,
      measure: entries ? 'Entries' : 'Points', goal: Number(r[col('Goal')]) || 0, start: start, end: end,
      bonus: Number(r[col('Bonus Points')]) || 0, awarded: awarded, progress: progress,
      active: (!start || today >= start) && (!end || today <= end),
      done: !!end && today > end,
    });
  });
  return out;
}

function clashChallenges() {
  return JSON.stringify(challengeProgress_().filter(function (c) { return c.active || c.done; }).slice(0, 8));
}

// Hourly (menu: Turn on challenge bonuses). Awards each squad that reached a goal, once.
function checkChallengeBonuses() {
  if (String(getSetting_('CHALLENGE_BONUSES')).toUpperCase() !== 'ON') return ['CHALLENGE_BONUSES is OFF.'];
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sheet = getSheetSafe_('Challenges');
    const lines = [];
    const awardedCol = CHALLENGE_HEADERS_.indexOf('Awarded') + 1;
    challengeProgress_().forEach(function (c) {
      if (!c.bonus || !c.goal) return;
      const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
      SQUAD_KEYS_.forEach(function (s) {
        if (c.awarded[s] || c.progress[s] < c.goal) return;
        const pts = { BLACK: '', GOLD: '', GREY: '', PURPLE: '' };
        pts[s] = c.bonus;
        getSheet_(TABS.EVENTS).appendRow(['', 'Challenge: ' + c.title, 'challenge_bonus', today,
          c.progress[s] + ' of ' + c.goal + ' ' + c.measure.toLowerCase(), '', pts.BLACK, pts.GOLD, pts.GREY, pts.PURPLE]);
        c.awarded[s] = today;
        lines.push(s + ' +' + c.bonus + ' for ' + c.title);
      });
      const text = Object.keys(c.awarded).map(function (s) { return s + ' ' + c.awarded[s]; }).join(', ');
      if (sheet) sheet.getRange(c.row, awardedCol).setValues([[text]]);
    });
    return lines;
  } finally {
    lock.releaseLock();
  }
}

function menuTurnOnChallenges_() {
  setSettingValue_('CHALLENGE_BONUSES', 'ON');
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'checkChallengeBonuses') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('checkChallengeBonuses').timeBased().everyHours(1).create();
  const lines = checkChallengeBonuses();
  SpreadsheetApp.getUi().alert('Challenge bonuses are ON (checked every hour).\n\n' + (lines.length ? 'Awarded now:\n' + lines.join('\n') : 'No squad has reached a goal yet.'));
}

function menuTurnOffChallenges_() {
  setSettingValue_('CHALLENGE_BONUSES', 'OFF');
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'checkChallengeBonuses') ScriptApp.deleteTrigger(t); });
  SpreadsheetApp.getUi().alert('Challenge bonuses are OFF. Progress still shows on the site.');
}

// ── TV FEED ─────────────────────────────────────────────────────────
// Latest points (newest first) and students with the most badges, for the hallway TV.
function clashTvFeed() {
  const labels = parsePairsSetting_('EVENT_LABELS');
  const latest = readEventsTab_().slice().sort(function (a, b) {
    return String(b.date).localeCompare(String(a.date)) || (Number(b.id) - Number(a.id));
  }).slice(0, 8).map(function (e) {
    return { name: publicEvent_(e).name, label: labels[e.type] || ({ duel_bonus: '⚔️ Squad Duel', mayhem_week: '🏔️ Mayhem Week', shirt_sales: '👕 Shirt Sales', squad_night: '🌙 Squad Night' })[e.type] || '', date: e.date, pts: e.pts };
  });
  const out = { latest: latest, stars: [], lore: recentLoreUnlocks_(), duels: JSON.parse(clashDuels()) };
  if (String(getSetting_('TV_SHOW_NAMES')).toUpperCase() === 'ON') {
    const sheet = getSheetSafe_(getSetting_('STUDENT_BOARD_TAB') || 'Student Board');
    if (sheet && sheet.getLastRow() > 1) {
      const data = sheet.getDataRange().getValues();
      const h = data[0].map(function (x) { return String(x).trim(); });
      const bi = h.indexOf('Badges'), ni = h.indexOf('Name'), si = h.indexOf('Squad');
      out.stars = data.slice(1).map(function (r) {
        const b = String(r[bi] || '').split(',').filter(function (x) { return x.trim(); });
        return { name: String(r[ni]), squad: String(r[si]).toUpperCase().replace('GRAY', 'GREY'), badges: b.length };
      }).filter(function (s) { return s.badges > 0; })
        .sort(function (a, b) { return b.badges - a.badges; }).slice(0, 10);
    }
  }
  return JSON.stringify(out);
}

// "a=Label, b=Label" → { a: 'Label' }
function parsePairsSetting_(key) {
  const out = {};
  String(getSetting_(key) || '').split(',').forEach(function (p) {
    const i = p.indexOf('=');
    if (i > 0) out[p.slice(0, i).trim()] = p.slice(i + 1).trim();
  });
  return out;
}
