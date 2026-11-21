/**
 * Clash of Classes -- Apps Script backend
 * File 57: Rivalry Week, the Team League, the quarter recap, and the staff cards that go with them.
 *
 * RIVALRY WEEK (Settings › RIVALRY ON, RIVALRY_BONUS 30)
 *   Built on Squad Duels. Admin › System › Start Rivalry Week (or the menu) adds two head-to-head
 *   duels for this week, Monday to Friday: 1st vs 2nd and 3rd vs 4th ("close races") or 1st vs 4th
 *   and 2nd vs 3rd ("upsets"). Each winner gets RIVALRY_BONUS the day after it ends, even when
 *   DUEL_BONUSES is OFF, and a Rivalry Week trophy. The home page shows the live scoreboard and
 *   each squad's trophies this year; rollover puts the trophy count in the Hall of Fame.
 *
 * TEAM LEAGUE (Settings › TEAM_LEAGUE ON, TEAM_NAMES "M, O, U, N, T, S", TEAM_MIN 3)
 *   Every team (M, O, U, N, T, S) has a group in each squad. Type each student's team in the
 *   Students tab's Team column. Inside each squad, the team groups race on how many of their
 *   students did something this month (got Trait Dough, sent service, played the Arcade, voted):
 *   a percent, so small groups have a fair shot. Groups smaller than TEAM_MIN aren't ranked.
 *   On the 1st, last month's winner in each squad is written to the Team League tab (Month |
 *   Squad | Team | Active % | Recorded), goes in the announcement, and earns the Team Banner badge.
 *   Team Teachers tab (Team | BLACK | GOLD | GREY | PURPLE): each group's teacher, shown with it.
 *   No squad points: every group in a squad is on the same squad.
 *
 * QUARTER RECAP (clashofclasses.org/recap, public; Settings › QUARTER_STARTS)
 *   The quarter's winner, standings, best week, Expeditions, Rivalry Weeks, Team League winners,
 *   Spotlight students and staff shout-outs, big enough for the TV. ?q=2 shows quarter 2.
 */

const TEAM_LEAGUE_HEADERS_ = ['Month', 'Squad', 'Team', 'Active %', 'Recorded'];
const TEAM_TEACHER_HEADERS_ = ['Team', 'BLACK', 'GOLD', 'GREY', 'PURPLE'];

function ensureLeagueTabs_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  [['Team League', TEAM_LEAGUE_HEADERS_, null], ['Team Teachers', TEAM_TEACHER_HEADERS_, teamNames_().map(function (t) { return [t, '', '', '', '']; })]].forEach(function (t) {
    if (ss.getSheetByName(t[0])) return;
    const sh = ss.insertSheet(t[0]);
    sh.getRange(1, 1, 1, t[1].length).setValues([t[1]]).setFontWeight('bold').setBackground('#490e6f').setFontColor('#ffffff');
    if (t[2] && t[2].length) sh.getRange(2, 1, t[2].length, t[1].length).setValues(t[2]);
    sh.setFrozenRows(1);
  });
}

// ── Rivalry Week ──
function rivalryDuels_() { return duelProgress_().filter(function (d) { return /^rivalry week/i.test(d.title); }); }

/** Adds this week's two rivalry duels. mode: 'close' (1v2, 3v4) or 'upset' (1v4, 2v3). */
function startRivalryWeek_(mode) {
  if (!settingOn_('RIVALRY', 'ON')) throw new Error('Rivalry Week is off (Settings › RIVALRY).');
  ensureDuelTab_();
  const today = todayYmd_(), dow = new Date(today + 'T12:00:00Z').getUTCDay();
  const mon = dow === 0 || dow === 6 ? mondayOf_(ymdShift_(today, 7)) : mondayOf_(today), fri = ymdShift_(mon, 4);
  if (rivalryDuels_().some(function (d) { return d.start === mon; })) throw new Error('Rivalry Week is already set for the week of ' + mon + '.');
  const totals = squadTotals_(), o = SQUAD_KEYS_.slice().sort(function (a, b) { return totals[b] - totals[a]; });
  const pairs = mode === 'upset' ? [[o[0], o[3]], [o[1], o[2]]] : [[o[0], o[1]], [o[2], o[3]]];
  const bonus = settingNum_('RIVALRY_BONUS', 30), sh = getSheet_('Squad Duels');
  pairs.forEach(function (p) {
    sh.appendRow([true, 'Rivalry Week: ' + titleSquad_(p[0]) + ' vs ' + titleSquad_(p[1]), mon, fri, 'Points', '', p[0] + ' vs ' + p[1], bonus, '', '']);
  });
  return pairs.map(function (p) { return titleSquad_(p[0]) + ' vs ' + titleSquad_(p[1]); });
}

function clashStartRivalry(token, mode) { requireAdmin_(token); return JSON.stringify({ pairs: startRivalryWeek_(mode), week: mondayOf_(todayYmd_()) }); }
function menuStartRivalry_() {
  try { SpreadsheetApp.getUi().alert('Rivalry Week is set (Monday to Friday):\n' + startRivalryWeek_('close').join('\n') + '\n\nEach winner gets RIVALRY_BONUS the day after it ends.'); }
  catch (e) { SpreadsheetApp.getUi().alert(e.message); }
}

/** Rivalry trophies this school year: { BLACK: n, ... }. */
function rivalryTrophies_() {
  const year = schoolYearLabel_(todayYmd_()), out = { BLACK: 0, GOLD: 0, GREY: 0, PURPLE: 0 };
  rivalryDuels_().forEach(function (d) { if (d.awarded && d.end && schoolYearLabel_(d.end) === year) d.leaders.forEach(function (s) { out[s]++; }); });
  return out;
}

/** For the home page: { live: [{ title, squads, scores, end }], trophies }. */
function rivalryNow_() {
  if (!settingOn_('RIVALRY', 'ON')) return null;
  const live = rivalryDuels_().filter(function (d) { return d.active; }).map(function (d) { return { title: d.title, squads: d.squads, scores: d.scores, end: d.end }; });
  const trophies = rivalryTrophies_();
  return live.length || SQUAD_KEYS_.some(function (s) { return trophies[s]; }) ? { live: live, trophies: trophies } : null;
}

function awardRivalries_() {
  if (!settingOn_('RIVALRY', 'ON')) return [];
  return awardDuelRows_(function (d) { return /^rivalry week/i.test(d.title); });
}

// ── Team League ──
function teamNames_() {
  const l = String(getSetting_('TEAM_NAMES') || 'M, O, U, N, T, S').split(',').map(function (s) { return s.trim().toUpperCase(); }).filter(String);
  return l.length ? l : ['M', 'O', 'U', 'N', 'T', 'S'];
}
function teamKey_(v) {
  const t = String(v || '').trim().toUpperCase().replace(/^TEAM\s+/, ''), names = teamNames_();
  return names.indexOf(t) > -1 ? t : '';
}

/** { 800#: 'M' } from the Students tab (Team column), then the Student Board. */
function teamOf_() {
  const out = {};
  boardRows_().forEach(function (r) { const t = teamKey_(r['Team']); if (t) out[id800_(r['800#'])] = t; });
  sheetRows_('Students').forEach(function (r) { const t = teamKey_(r['Team']); if (t) out[id800_(r['800#'])] = t; });
  return out;
}

function teamTeachers_() {
  const out = {};
  sheetRows_('Team Teachers').forEach(function (r) { const t = teamKey_(r['Team']); if (t) SQUAD_KEYS_.forEach(function (s) { if (String(r[s] || '').trim()) out[s + '-' + t] = String(r[s]).trim(); }); });
  return out;
}

/** { from, to, squads: { BLACK: [{ team, teacher, size, active, pct, ranked }] best first }, leaders: { BLACK: ['M'] } }. */
function teamLeague_(from, to) {
  const teams = teamOf_(), days = activityDays_(), teachers = teamTeachers_(), min = Math.max(1, settingNum_('TEAM_MIN', 3));
  const squadById = {};
  boardRows_().forEach(function (r) { const s = squadOf_(r['Squad']); if (SQUAD_KEYS_.indexOf(s) > -1) squadById[id800_(r['800#'])] = s; });
  sheetRows_('Students').forEach(function (r) { const s = squadOf_(r['Squad']); if (SQUAD_KEYS_.indexOf(s) > -1) squadById[id800_(r['800#'])] = s; });
  const groups = {};
  Object.keys(teams).forEach(function (id) {
    const s = squadById[id]; if (!s) return;
    const k = s + '-' + teams[id], g = groups[k] = groups[k] || { squad: s, team: teams[id], size: 0, active: 0 };
    g.size++;
    if (anyDayIn_(days[id], from, to)) g.active++;
  });
  const out = { from: from, to: to, squads: {}, leaders: {}, any: Object.keys(groups).length > 0 };
  SQUAD_KEYS_.forEach(function (s) {
    const list = teamNames_().map(function (t) { return groups[s + '-' + t]; }).filter(Boolean).map(function (g) {
      return { team: g.team, teacher: teachers[s + '-' + g.team] || '', size: g.size, active: g.active, pct: g.size ? Math.round(100 * g.active / g.size) : 0, ranked: g.size >= min };
    }).sort(function (a, b) { return (b.ranked - a.ranked) || b.pct - a.pct || a.team.localeCompare(b.team); });
    out.squads[s] = list;
    const top = list.filter(function (g) { return g.ranked && g.active > 0; });
    out.leaders[s] = top.length ? top.filter(function (g) { return g.pct === top[0].pct; }).map(function (g) { return g.team; }) : [];
  });
  return out;
}

function calMonth_(ym) {
  const y = Number(ym.slice(0, 4)), m = Number(ym.slice(5, 7));
  const last = new Date(Date.UTC(y, m, 0)).getUTCDate();
  return { from: ym + '-01', to: ym + '-' + (last < 10 ? '0' : '') + last };
}
function monthBefore_(ym) { const y = Number(ym.slice(0, 4)), m = Number(ym.slice(5, 7)) - 1; return m < 1 ? (y - 1) + '-12' : y + '-' + (m < 10 ? '0' : '') + m; }

/** Daily: writes last month's winners once. Returns the rows added. */
function recordTeamLeague_() {
  if (!settingOn_('TEAM_LEAGUE', 'ON')) return 0;
  ensureLeagueTabs_();
  const ym = monthBefore_(todayYmd_().slice(0, 7));
  if (sheetRows_('Team League').some(function (r) { return monthKey_(r['Month']) === ym; })) return 0;
  const b = calMonth_(ym), L = teamLeague_(b.from, b.to);
  if (!L.any) return 0;
  const rows = [];
  SQUAD_KEYS_.forEach(function (s) { L.leaders[s].forEach(function (t) { const g = L.squads[s].filter(function (x) { return x.team === t; })[0]; rows.push([ym, s, t, g.pct, todayYmd_()]); }); });
  if (!rows.length) rows.push([ym, 'NONE', '', '', todayYmd_()]);
  const sh = getSheet_('Team League');
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, TEAM_LEAGUE_HEADERS_.length).setValues(rows);
  return rows.length;
}

/** One student's group this month: { team, teacher, pct, active, size, place, of, leaders } or null. */
function myTeam_(id, squad) {
  if (!settingOn_('TEAM_LEAGUE', 'ON')) return null;
  const t = teamOf_()[id];
  if (!t) return null;
  const b = calMonth_(todayYmd_().slice(0, 7)), L = teamLeague_(b.from, b.to), list = L.squads[squad] || [];
  const g = list.filter(function (x) { return x.team === t; })[0];
  if (!g) return null;
  const ranked = list.filter(function (x) { return x.ranked; });
  return { team: t, teacher: g.teacher, pct: g.pct, active: g.active, size: g.size, place: ranked.indexOf(g) + 1, of: ranked.length, leaders: L.leaders[squad] };
}

// ── Quarter recap ──
function recapWindow_(q) {
  const today = todayYmd_();
  const starts = String(getSetting_('QUARTER_STARTS') || '').split(',').map(function (s) { return s.trim(); }).filter(function (s) { return /^\d{4}-\d{2}-\d{2}$/.test(s); }).sort();
  if (!starts.length) return { from: ymdShift_(today, -62), to: today, label: 'The last 9 weeks', q: 0, quarters: 0 };
  let i = Number(q) > 0 ? Math.min(starts.length, Number(q)) - 1 : starts.filter(function (s) { return s <= today; }).length - 1;
  if (i < 0) i = 0;
  const end = starts[i + 1] ? ymdShift_(starts[i + 1], -1) : today;
  return { from: starts[i], to: end < today ? end : today, label: 'Quarter ' + (i + 1), q: i + 1, quarters: starts.length, done: end < today };
}

function recapData_(q) {
  const win = recapWindow_(q), events = readEventsTab_();
  const inside = function (d) { d = String(d || '').slice(0, 10); return d >= win.from && d <= win.to; };
  const totals = { BLACK: 0, GOLD: 0, GREY: 0, PURPLE: 0 };
  const priv = listOf_(getSetting_('PRIVATE_EVENT_TYPES') || 'trail_journal, trailback_panel, camp_mountaineer');
  let service = 0, dough = 0;
  events.forEach(function (e) {
    if (!inside(e.date)) return;
    SQUAD_KEYS_.forEach(function (s) { totals[s] += Number(e.pts[s]) || 0; });
    if (priv.indexOf(String(e.type)) > -1) return;
    if (String(e.type) === 'service') service++;
    if (String(e.type) === 'trait_dough') dough++;
  });
  const order = SQUAD_KEYS_.slice().sort(function (a, b) { return totals[b] - totals[a]; });
  let best = null;
  for (let mon = mondayOf_(win.from); mon <= win.to; mon = ymdShift_(mon, 7)) {
    const pts = { BLACK: 0, GOLD: 0, GREY: 0, PURPLE: 0 }, end = ymdShift_(mon, 6);
    events.forEach(function (e) { const d = String(e.date || '').slice(0, 10); if (d >= mon && d <= end && inside(d)) SQUAD_KEYS_.forEach(function (s) { pts[s] += Number(e.pts[s]) || 0; }); });
    SQUAD_KEYS_.forEach(function (s) { if (pts[s] > 0 && (!best || pts[s] > best.pts)) best = { squad: s, week: mon, pts: pts[s] }; });
  }
  const safe = function (fn, def) { try { return fn(); } catch (e) { return def; } };
  return {
    window: win, totals: totals, order: order, best: best,
    expeditions: safe(function () { return expeditions_(events).filter(function (x) { return x.winners.length && inside(x.wonOn); }).map(function (x) { return { title: x.title, winners: x.winners }; }); }, []),
    rivalries: safe(function () { return rivalryDuels_().filter(function (d) { return d.awarded && inside(d.end); }).map(function (d) { return { title: d.title, winners: d.leaders, scores: d.scores }; }); }, []),
    weather: safe(function () { return weatherRows_().filter(function (w) { return inside(w.from); }).map(function (w) { return w.name; }); }, []),
    team: safe(function () { return sheetRows_('Team League').filter(function (r) { const m = monthKey_(r['Month']); return m && SQUAD_KEYS_.indexOf(squadOf_(r['Squad'])) > -1 && m + '-28' >= win.from && m + '-01' <= win.to; })
      .map(function (r) { return { month: monthKey_(r['Month']), squad: squadOf_(r['Squad']), team: String(r['Team']), pct: Number(r['Active %']) || 0 }; }); }, []),
    spotlight: safe(function () { return spotlightPublic_(60).filter(function (p) { return inside(p.week); }).slice(0, 8).map(function (p) { return { who: p.who, squad: p.squad }; }); }, []),
    shoutouts: safe(function () { return sheetRows_('Staff Shout-outs').filter(function (r) { return ticked_(r['Approved']) && inside(r['When']); }).length; }, 0),
    hands: safe(function () { return sheetRows_('Helping Hands').filter(function (r) { return inside(asYmd_(r['Week'])) && SQUAD_KEYS_.indexOf(squadOf_(r['Squad'])) > -1; }).length; }, 0),
    goals: safe(function () { return sheetRows_('Squad Goals').filter(function (r) { return inside(asYmd_(r['Week'])) && /hit/.test(String(r['Result'])); }).length; }, 0),
    arcade: safe(function () { return sheetRows_('Arcade Plays').filter(function (p) { return inside(p['When']); }).length; }, 0),
    service: service, dough: dough,
    banners: safe(squadBanners_, {}),
  };
}

function clashRecap(q) {
  if (String(getSetting_('QUARTER_RECAP') || 'ON').toUpperCase() === 'OFF') return JSON.stringify({ off: true }); // the Tracker shows nothing
  return JSON.stringify(recapData_(q));
}

function serveRecap_() {
  return HtmlService.createHtmlOutputFromFile('Recap').setTitle('Quarter recap · Clash of Classes')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ── Staff cards (Admin › System) ──
/** The fairness check on the site: the Points Balance flags and the biggest sources. */
function clashFairness(token) {
  requireAdmin_(token);
  const r = pointsBalance_();
  return JSON.stringify({ from: r.from, to: r.to, days: r.days, flags: r.flags, types: r.types.slice(0, 8).map(function (t) { return { label: t.label, total: t.total, share: t.share, pts: t.pts, flag: t.flag }; }) });
}

function declined_(v) { return /^declined/i.test(String(v || '').trim()); }

/** Shout-outs and banners waiting for a yes. */
function clashApprovals(token) {
  requireAdmin_(token);
  ensureShoutoutTabs_(); ensureRallyTabs_();
  const waiting = function (r) { return !ticked_(r['Approved']) && !declined_(r['Approved']); };
  return JSON.stringify({
    shoutouts: sheetRows_('Staff Shout-outs').filter(waiting).map(function (r) { return { row: r._row, when: String(r['When']), from: String(r['From']), who: shortName_(r['Student']), squad: squadOf_(r['Squad']), message: String(r['Message']) }; }),
    banners: sheetRows_('Squad Banners').filter(waiting).map(function (r) { const b = bannerOf_(r); b.row = r._row; b.by = String(r['By'] || ''); return b; }),
    // Closed Pulse polls whose squad points wait for a yes (06_Polls.gs, POLL_APPROVAL).
    polls: pollPointsWaiting_().filter(function (r) { return !ticked_(r['Approved']); }).map(function (r) {
      const pts = {}; SQUAD_KEYS_.forEach(function (s) { pts[s] = Number(r[s]) || 0; });
      return { row: r._row, question: String(r['Question'] || ''), closed: asYmd_(r['Closed']).slice(0, 10), votes: Number(r['Votes']) || 0, pts: pts };
    }),
  });
}

function clashApprove(token, kind, row, yes) {
  requireAdmin_(token);
  if (kind === 'poll') { decidePollPoints_(row, yes); return clashApprovals(token); }
  const tab = kind === 'banner' ? 'Squad Banners' : 'Staff Shout-outs', headers = kind === 'banner' ? BANNER_HEADERS_ : SHOUT_HEADERS_;
  const r = sheetRows_(tab).filter(function (x) { return x._row === Number(row); })[0];
  if (!r) throw new Error('That row moved. Refresh and try again.');
  getSheet_(tab).getRange(r._row, headers.indexOf('Approved') + 1).setValue(yes ? true : 'Declined');
  return clashApprovals(token);
}

/** Team League grid for Admin › System: this month, every squad. */
function clashTeamLeague(token) {
  requireAdmin_(token);
  ensureLeagueTabs_();
  const b = calMonth_(todayYmd_().slice(0, 7)), L = teamLeague_(b.from, b.to);
  L.month = todayYmd_().slice(0, 7);
  L.teams = teamNames_();
  return JSON.stringify(L);
}

// ── My Climb and Generals HQ ──
function squadOfId_(id) {
  const b = boardRows_().filter(function (r) { return id800_(r['800#']) === id; })[0];
  if (b && SQUAD_KEYS_.indexOf(squadOf_(b['Squad'])) > -1) return squadOf_(b['Squad']);
  const s = sheetRows_('Students').filter(function (r) { return id800_(r['800#']) === id; })[0];
  return s ? squadOf_(s['Squad']) : '';
}

/** Everything new on My Climb: badges, shout-outs, Helping Hands, the banner, the team, and (for leaders) goals. */
function clashMeExtras(asId, adminToken) {
  return withSample_(asId, function () {
    const id = meId_(asId, adminToken), squad = id === SAMPLE_ID_ ? sampleSquad_(asId) : squadOfId_(id);
    const safe = function (fn, def) { try { return fn(); } catch (e) { return def; } };
    const L = safe(function () { return leaderFromMe_(asId, adminToken); }, null);
    const goal = safe(function () { return squadGoalFor_(squad); }, null);
    return JSON.stringify({
      squad: squad,
      badges: settingOn_('BADGES', 'ON') ? safe(function () { return badgesFor_(id, squad); }, []) : [],
      shoutouts: safe(function () { return myShoutouts_(id); }, []),
      hands: safe(helpingHandsNow_, null),
      banner: safe(function () { return squadBanners_()[squad]; }, null),
      team: safe(function () { return myTeam_(id, squad); }, null),
      goal: goal,
      leader: L ? { role: L.role, banner: safe(function () { return bannerFor_(squad); }, null) } : null,
    });
  });
}

/** Extra Generals HQ data (05_Generals.gs). */
function hqRally_(squad) {
  const safe = function (fn, def) { try { return fn(); } catch (e) { return def; } };
  const b = calMonth_(todayYmd_().slice(0, 7));
  return {
    goal: safe(function () { return squadGoalFor_(squad); }, null),
    banner: safe(function () { return bannerFor_(squad); }, null),
    hands: safe(helpingHandsNow_, null),
    team: settingOn_('TEAM_LEAGUE', 'ON') ? safe(function () { const L = teamLeague_(b.from, b.to); return { list: L.squads[squad], leaders: L.leaders[squad] }; }, null) : null,
  };
}

// Once a day (50_Weather.gs maybeDailyExtras_).
function dailyLeagues_() {
  try { payHelpingHands_(); } catch (e) { noteError_('helping hands', e); }
  try { paySquadGoals_(); } catch (e) { noteError_('squad goals', e); }
  try { postPlannedEvents_(); } catch (e) { noteError_('planned events', e); }
  try { awardRivalries_(); } catch (e) { noteError_('rivalry week', e); }
  try { recordTeamLeague_(); } catch (e) { noteError_('team league', e); }
}

// file-stamp (tools/stamp-files.js): lets Admin › System tell whether this file is the newest one
var FILE_STAMP_57_Leagues = '';

// file-stamp (tools/stamp-files.js): lets Admin › System tell whether this file is the newest one
var FILE_STAMP_57_Leagues = 'a2825d57c5';
