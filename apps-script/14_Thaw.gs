/**
 * Clash of Classes -- Apps Script backend
 * File 14: Thawing a freeze. Two kinds, set in the Point Freezes tab's Kind column:
 *
 * PEAKS FREEZE (Kind blank or "Peaks"): the squad as a whole wasn't meeting Mountaineer Peaks
 *   expectations. It thaws only when the squad works together: EVERY goal in Thaw Goals is met
 *   AND staff tick Staff Sign-off. Points earned while frozen never come back.
 * FUN FREEZE (Kind "Fun"): a game. The squad pushes its luck: reaching the goal once gets back
 *   25% of the points frozen out, twice the goal 50%, three times 100% (FUN_THAW_TIERS).
 *   Generals cash out in Generals HQ whenever a tier is reached (the top tier cashes out by
 *   itself). If Thaw By passes first, the ice melts: the freeze ends and nothing comes back.
 *   Generals get one second chance per freeze: Thaw By moves FUN_SECOND_CHANCE_DAYS school days.
 *
 * Thaw Goals, separated by ; or new lines:
 *   points: 150      the squad earns 150 points while frozen (they still don't count)
 *   everyone: 70%    70% of the squad (Student Board) contributed something since the freeze began
 *   votes: 60        the squad casts 60 Daily Poll / Pulse votes since the freeze began (answers
 *                    the school learns from, so a freeze also grows the data)
 *   anything else    a goal staff check (a squad reflection, a service act); Staff Sign-off covers it
 *
 * The goals are checked every few minutes when the site loads (checkThaws_), and by
 * Clash of Classes › Check thaw goals now.
 *
 * FRESH START (Settings › FRESH_START, a date; Menu › Fresh start: end every freeze on a date):
 *   on that day every freeze that began before it ends. A Peaks Freeze simply ends (nothing comes
 *   back); a Fun Freeze cashes out whatever tier it reached. Use it for the January restart so every
 *   squad walks into its first meeting of the year unfrozen. Freezes set on or after it are untouched.
 */

function thawGoals_(text) {
  return String(text || '').split(/[;\n]+/).map(function (g) { return g.trim(); }).filter(String).map(function (g) {
    let m = g.match(/^points?\s*:\s*(\d+)/i);
    if (m) return { type: 'points', need: Number(m[1]), text: 'Earn ' + m[1] + ' points together' };
    m = g.match(/^everyone\s*:\s*(\d+)\s*%?/i);
    if (m) return { type: 'everyone', need: Number(m[1]), text: m[1] + '% of the squad pitches in' };
    m = g.match(/^votes?\s*:\s*(\d+)/i);
    if (m) return { type: 'votes', need: Number(m[1]), text: 'Cast ' + m[1] + ' poll and Pulse votes' };
    return { type: 'staff', need: 1, text: g };
  });
}

// Points the squad earned while frozen (the "❄️ GOLD frozen: +5 not counted" notes since Start).
function icedPoints_(squad, start, events) {
  const re = new RegExp('❄️ ' + squad + ' frozen: \\+(\\d+) not counted', 'g');
  let total = 0;
  (events || readEventsTab_()).forEach(function (e) {
    const ymd = String(e.date || '').slice(0, 10);
    if (ymd < start) return;
    let m;
    re.lastIndex = 0;
    while ((m = re.exec(String(e.note || '')))) total += Number(m[1]) || 0;
  });
  return total;
}

// Daily Poll / Pulse votes the squad cast on or after Start.
function squadVotesSince_(squad, start) {
  return sheetRows_('Poll Votes').filter(function (v) { return squadOf_(v['Squad']) === squad && String(v['Date'] || '').slice(0, 10) >= start; }).length;
}

// Share of the squad (Student Board rows) with a contribution on or after Start: { pct, n, of }.
function pitchedIn_(squad, start) {
  const rows = sheetRows_(getSetting_('STUDENT_BOARD_TAB') || 'Student Board').filter(function (r) { return squadOf_(r['Squad']) === squad; });
  const n = rows.filter(function (r) { return String(r['Last Contribution'] || '').slice(0, 10) >= start; }).length;
  return { pct: rows.length ? Math.round(100 * n / rows.length) : 0, n: n, of: rows.length };
}

function funTiers_() {
  const t = String(getSetting_('FUN_THAW_TIERS') || '1=25, 2=50, 3=100').split(',').map(function (p) {
    const m = p.match(/(\d+(?:\.\d+)?)\s*=\s*(\d+)/); return m ? { mult: Number(m[1]), pct: Number(m[2]) } : null;
  }).filter(Boolean).sort(function (a, b) { return a.mult - b.mult; });
  return t.length ? t : [{ mult: 1, pct: 25 }, { mult: 2, pct: 50 }, { mult: 3, pct: 100 }];
}

// The freeze in force today (or on asOf) for a squad, with its thaw progress (null when not frozen).
// What a page gets; the Reason never leaves the Sheet.
function freezeDetail_(squad, events, asOf) {
  const today = asOf || todayYmd_();
  const f = freezeWindows_().filter(function (w) { return w.squad === squad && today >= w.start && (!w.end || today <= w.end); })[0];
  if (!f) return null;
  const goals = thawGoals_(f.goals);
  let iced = null, pitched = null, votes = null;
  const measure = function (g) {
    if (g.type === 'points') { if (iced === null) iced = icedPoints_(squad, f.start, events); return iced; }
    if (g.type === 'votes') { if (votes === null) votes = squadVotesSince_(squad, f.start); return votes; }
    if (g.type === 'everyone') { if (!pitched) pitched = pitchedIn_(squad, f.start); return pitched.pct; }
    return f.signoff ? 1 : 0;
  };
  if (f.kind === 'FUN') {
    const goal = goals.filter(function (g) { return g.type !== 'staff'; })[0] || { type: 'points', need: 50, text: 'Earn 50 points together' };
    const have = measure(goal);
    if (iced === null) iced = icedPoints_(squad, f.start, events);
    const tiers = funTiers_().map(function (t) { return { mult: t.mult, pct: t.pct, need: Math.ceil(goal.need * t.mult), reached: have >= Math.ceil(goal.need * t.mult) }; });
    const top = tiers.filter(function (t) { return t.reached; }).pop();
    const pct = top ? top.pct : 0;
    return { kind: 'Fun', until: f.end, thawBy: f.thawBy, goal: { text: goal.text, type: goal.type, have: have, need: goal.need },
      tiers: tiers, pct: pct, iced: iced, cash: Math.round(iced * pct / 100), secondChanceLeft: !f.secondChance, everyone: pitched };
  }
  const list = (goals.length ? goals : []).map(function (g) {
    const have = measure(g);
    return { text: g.text, type: g.type, have: have, need: g.need, met: have >= g.need };
  });
  return { kind: 'Peaks', until: f.end, goals: list, signoff: !!f.signoff, everyone: pitched,
    ready: !!f.signoff && list.every(function (g) { return g.met; }) };
}

function ymdShift_(ymd, days) {
  const d = new Date(ymd + 'T12:00:00'); d.setDate(d.getDate() + days);
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

// Ends a freeze as of yesterday (today's points count again) and notes how it ended.
function endFreeze_(f, how) {
  const sh = getSheet_('Point Freezes');
  const h = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(function (x) { return String(x).trim(); });
  const y = ymdShift_(todayYmd_(), -1);
  sh.getRange(f.row, h.indexOf('End') + 1).setValue(y < f.start ? f.start : y);
  if (y < f.start) sh.getRange(f.row, h.indexOf('Start') + 1).setValue(ymdShift_(f.start, 1)); // started today: no frozen days at all
  sh.getRange(f.row, h.indexOf('Thawed') + 1).setValue(how);
}

function cashOut_(f, d, by) {
  const pts = { BLACK: '', GOLD: '', GREY: '', PURPLE: '' };
  pts[f.squad] = d.cash;
  if (d.cash > 0) {
    getSheet_(TABS.EVENTS).appendRow(['', 'Fun Freeze thaw: ' + titleSquad_(f.squad) + ' cashed out ' + d.pct + '%', 'freeze_thaw', todayYmd_(),
      'Got back ' + d.pct + '% of ' + d.iced + ' frozen points' + (by ? ' (' + by + ')' : ''), '', pts.BLACK, pts.GOLD, pts.GREY, pts.PURPLE]);
  }
  endFreeze_(f, 'Cashed out ' + d.pct + '% (+' + d.cash + ') ' + todayYmd_() + (by ? ' by ' + by : ''));
}

function titleSquad_(s) { return s.charAt(0) + s.slice(1).toLowerCase(); }

// FRESH_START reached: end every freeze that began before it (a Fun Freeze cashes out its tier).
// Ends them the day before, so the fresh-start day itself counts. Returns lines. Idempotent.
function applyFreshStart_(events) {
  const day = String(getSetting_('FRESH_START') || '').trim().slice(0, 10), today = todayYmd_(), lines = [];
  if (!/^\d{4}-\d{2}-\d{2}$/.test(day) || today < day) return lines;
  const eve = ymdShift_(day, -1);
  freezeWindows_(true).forEach(function (f) {
    if (f.thawed || f.start >= day || (f.end && f.end < day)) return;
    const sh = getSheet_('Point Freezes');
    const h = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(function (x) { return String(x).trim(); });
    let note = 'Fresh start ' + day;
    if (f.kind === 'FUN') {
      // Measured as of the fresh start: what it earned while frozen, and the tier that reached.
      const d = freezeDetailOn_(f, eve, events);
      if (d && d.cash > 0) {
        const pts = { BLACK: '', GOLD: '', GREY: '', PURPLE: '' }; pts[f.squad] = d.cash;
        getSheet_(TABS.EVENTS).appendRow(['', 'Fun Freeze thaw: ' + titleSquad_(f.squad) + ' cashed out ' + d.pct + '% (fresh start)', 'freeze_thaw', day,
          'Got back ' + d.pct + '% of ' + d.iced + ' frozen points at the fresh start', '', pts.BLACK, pts.GOLD, pts.GREY, pts.PURPLE]);
        note += ': cashed out ' + d.pct + '% (+' + d.cash + ')';
      }
    }
    sh.getRange(f.row, h.indexOf('End') + 1).setValue(eve < f.start ? f.start : eve);
    sh.getRange(f.row, h.indexOf('Thawed') + 1).setValue(note);
    lines.push(f.squad + ': ' + note);
  });
  return lines;
}

// A freeze's thaw detail as of a given day (for the fresh start, which may run after that day).
function freezeDetailOn_(f, ymd, events) { return freezeDetail_(f.squad, events, ymd); }

function menuFreshStart_() {
  const ui = SpreadsheetApp.getUi();
  // Suggest the first school day of January after today.
  const now = todayYmd_(), y = Number(now.slice(0, 4)) + (now.slice(5, 7) === '01' ? 0 : 1);
  const suggest = schoolDayAfter_(ymdShift_(y + '-01-01', -1), 1);
  const r = ui.prompt('Fresh start', 'On this day every freeze that began before it ends (a Fun Freeze cashes out the tier it reached). Type a date (yyyy-mm-dd), or "off" to clear it.\n\nSuggested: ' + suggest,
    ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const v = String(r.getResponseText() || '').trim() || suggest;
  if (/^off$/i.test(v)) { setSettingValue_('FRESH_START', ''); ui.alert('Fresh start cleared.'); return; }
  if (!/^\d{4}-\d{2}-\d{2}$/.test(v)) { ui.alert('Type the date as yyyy-mm-dd, e.g. ' + suggest + '.'); return; }
  setSettingValue_('FRESH_START', v);
  const lines = checkThaws_();
  ui.alert('Fresh start set for ' + v + '. ' + (v > todayYmd_() ? 'Freezes end by themselves that morning.' : 'Done:\n' + (lines.join('\n') || 'No freezes were open.')));
}

// Thaws what's ready, cashes out a Fun Freeze at the top tier, melts one past its Thaw By.
// Returns lines describing what happened.
function checkThaws_() {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const today = todayYmd_(), lines = [];
    const events = readEventsTab_();
    applyFreshStart_(events).forEach(function (l) { lines.push(l); });
    freezeWindows_().forEach(function (f) {
      if (f.thawed || today < f.start || (f.end && today > f.end)) return;
      const d = freezeDetail_(f.squad, events);
      if (!d) return;
      if (d.kind === 'Peaks' && d.ready) { endFreeze_(f, 'Thawed together ' + today); lines.push(f.squad + ': Peaks Freeze thawed'); }
      else if (d.kind === 'Fun' && d.pct >= 100) { cashOut_(f, d, 'top tier'); lines.push(f.squad + ': Fun Freeze cashed out at 100%'); }
      else if (d.kind === 'Fun' && f.thawBy && today > f.thawBy) { endFreeze_(f, 'Melted ' + today + ': Thaw By passed, nothing back'); lines.push(f.squad + ': Fun Freeze melted'); }
    });
    return lines;
  } finally {
    lock.releaseLock();
  }
}

// At most every 5 minutes, from page loads.
function maybeCheckThaws_() {
  const cache = CacheService.getScriptCache();
  if (cache.get('thaw-check')) return;
  cache.put('thaw-check', '1', 300);
  try { if (freezeWindows_().length) checkThaws_(); } catch (e) { console.warn('Thaw check: ' + e.message); }
}

function menuCheckThaws_() {
  const lines = checkThaws_();
  SpreadsheetApp.getUi().alert(lines.length ? lines.join('\n') : 'No freeze is ready to thaw yet.');
}

// ── Generals HQ (Fun Freeze only) ───────────────────────────────────
function clashGeneralsCashOut(token) {
  const me = hqSession_(token);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const today = todayYmd_();
    const f = freezeWindows_().filter(function (w) { return w.squad === me.squad && !w.thawed && today >= w.start && (!w.end || today <= w.end); })[0];
    const d = f && freezeDetail_(me.squad);
    if (!d || d.kind !== 'Fun') throw new Error('Only a Fun Freeze can be cashed out.');
    if (!d.pct) throw new Error('Reach the first goal before you can cash out.');
    cashOut_(f, d, me.name);
  } finally {
    lock.releaseLock();
  }
  return clashGeneralsHq(token);
}

function clashGeneralsSecondChance(token) {
  const me = hqSession_(token);
  const today = todayYmd_();
  const f = freezeWindows_().filter(function (w) { return w.squad === me.squad && !w.thawed && today >= w.start && (!w.end || today <= w.end); })[0];
  if (!f || f.kind !== 'FUN') throw new Error('Second chances are for a Fun Freeze.');
  if (f.secondChance) throw new Error('Your squad already used its second chance on this freeze.');
  if (!f.thawBy) throw new Error('This freeze has no Thaw By date to move.');
  const days = Math.max(1, Number(getSetting_('FUN_SECOND_CHANCE_DAYS')) || 2);
  let off = {};
  try { schoolCalendar_().noSchool.forEach(function (x) { off[x] = true; }); } catch (e) { off = {}; }
  let next = f.thawBy, added = 0;
  while (added < days) {
    next = ymdShift_(next, 1);
    const dow = new Date(next + 'T12:00:00').getDay();
    if (dow !== 0 && dow !== 6 && !off[next]) added++;
  }
  const sh = getSheet_('Point Freezes');
  const h = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(function (x) { return String(x).trim(); });
  sh.getRange(f.row, h.indexOf('Thaw By') + 1).setValue(next);
  sh.getRange(f.row, h.indexOf('Second Chance') + 1).setValue('Used ' + today + ' by ' + me.name + ': Thaw By moved to ' + next);
  return clashGeneralsHq(token);
}

// ── STUDENT RELEASE: how freezes work, in student words ─────────────
// Read aloud (menu) and shown on My Squad (Pop.html freezeExplainer) and in January squad meetings.
function freezeRelease_() {
  const fresh = String(getSetting_('FRESH_START') || '').trim().slice(0, 10);
  const when = /^\d{4}-\d{2}-\d{2}$/.test(fresh) && fresh >= todayYmd_() ? 'On ' + Number(fresh.slice(5, 7)) + '/' + Number(fresh.slice(8, 10)) + ', ' : 'In January, ';
  const tiers = funTiers_();
  return [
    'Clash of Classes news: here is how freezes work for the new year.',
    when + 'every squad starts fresh. Any freeze ends, and every squad walks into its first squad meeting of the year unfrozen.',
    'First, the big one: no squad ever loses points. A freeze only means new points stop counting for a little while.',
    'A Peaks Freeze happens when a squad as a whole is not meeting Mountaineer Peaks expectations. You thaw it together: ' +
      'your squad\'s goals are on your My Squad page, like getting most of the squad to pitch in or casting poll and Pulse votes, and then your advisor signs off.',
    'A Fun Freeze is a game. Hit your squad\'s goal once and you get back ' + tiers[0].pct + ' percent of your frozen points' +
      (tiers.length > 2 ? ', ' + tiers[1].mult + ' times the goal gets ' + tiers[1].pct + ' percent, and ' + tiers[tiers.length - 1].mult + ' times the goal gets them all back' : '') +
      '. Your Generals decide when to cash out, but if the Thaw By date passes first, the ice melts and nothing comes back.',
    'New this year: every poll and Pulse vote counts toward thawing a freeze, so your answers help your squad and help EMS learn what you think.',
    'Check My Squad on clashofclasses.org with your 800 number to see how your squad is doing. Let\'s climb!',
  ].join(' ');
}

function menuFreezeRelease_() { SpreadsheetApp.getUi().alert('Freeze & thaw: read this in the morning announcements', freezeRelease_(), SpreadsheetApp.getUi().ButtonSet.OK); }
