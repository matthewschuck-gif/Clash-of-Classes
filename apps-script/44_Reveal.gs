/**
 * Clash of Classes -- Apps Script backend
 * File 44: the end of the year. The Clash Bash reveal show and the year's awards.
 *   clashofclasses.org/?page=reveal (admin password): a click-through show for the TV or an
 *   assembly: each area of the school and the squad that held it longest, the awards, the Spotlight
 *   stories, then the final standings counted up from 4th place, then the Clash Bash winner.
 *   The same page prints award certificates (one per page).
 * Awards (season = first marking period to SEASON_END):
 *   Trait Dough champion (overall and one per squad), Service star, Most poll and Pulse votes,
 *   Most improved squad (last marking period vs the first), Best attendance month, Longest Hot
 *   Streak, Never froze, Most of the map held. Students are named first name + last initial.
 */

function seasonWindow_() {
  const mps = markingPeriods_(), today = todayYmd_();
  const start = String(getSetting_('SCHOOL_START') || '').trim().slice(0, 10) || (mps[0] ? mps[0].start : ymdShift_(today, -300));
  const end = seasonEnd_() && seasonEnd_() < today ? seasonEnd_() : today;
  return { from: start, to: end };
}

// Days each squad held each area this season, from the Map Battles log.
function areaHolds_(win) {
  const battles = sheetRows_('Map Battles').map(function (r) { return { date: String(r['Date']).slice(0, 10), zone: String(r['Zone'] || ''), to: squadOf_(r['Taken By']) }; })
    .filter(function (b) { return b.zone && b.date && b.date <= win.to; }).sort(function (a, b) { return a.date.localeCompare(b.date); });
  const zones = {};
  sheetRows_('Map Zones').forEach(function (r) { if (ticked_(r['Show']) && String(r['Zone'] || '').trim()) zones[String(r['Zone']).trim()] = { BLACK: 0, GOLD: 0, GREY: 0, PURPLE: 0 }; });
  const days = function (a, b) { return Math.max(0, Math.round((new Date(b + 'T12:00:00Z') - new Date(a + 'T12:00:00Z')) / 86400000)); };
  Object.keys(zones).forEach(function (z) {
    const list = battles.filter(function (b) { return b.zone === z; });
    list.forEach(function (b, i) {
      const from = b.date < win.from ? win.from : b.date, to = i + 1 < list.length ? list[i + 1].date : win.to;
      if (SQUAD_KEYS_.indexOf(b.to) > -1) zones[z][b.to] += days(from, to < win.from ? win.from : to);
    });
  });
  return Object.keys(zones).map(function (z) {
    const best = SQUAD_KEYS_.slice().sort(function (a, b) { return zones[z][b] - zones[z][a]; })[0];
    return { zone: z, squad: zones[z][best] ? best : '', days: zones[z][best], all: zones[z] };
  });
}

function boardNames_() {
  const out = {};
  sheetRows_(getSetting_('STUDENT_BOARD_TAB') || 'Student Board').forEach(function (r) { const id = id800_(r['800#']); if (id) out[id] = { name: String(r['Name'] || ''), squad: squadOf_(r['Squad']) }; });
  sheetRows_('Students').forEach(function (r) {
    const id = id800_(r['800#']);
    if (id && !out[id]) out[id] = { name: shortName_(String(r['First Name'] || '') + ' ' + String(r['Last Name'] || '')), squad: squadOf_(r['Squad']) };
  });
  return out;
}

/** The year's awards: [{ title, who, squad, detail }]. */
function yearAwards_(win) {
  win = win || seasonWindow_();
  const names = boardNames_(), awards = [], inWin = function (d) { d = String(d || '').slice(0, 10); return d >= win.from && d <= win.to; };
  const top = function (counts) { return Object.keys(counts).sort(function (a, b) { return counts[b] - counts[a]; })[0]; };
  // Trait Dough: paper bills (Student Board) + Virtual Trait Dough bills.
  const dough = {};
  sheetRows_(getSetting_('STUDENT_BOARD_TAB') || 'Student Board').forEach(function (r) { const id = id800_(r['800#']); if (id) dough[id] = (dough[id] || 0) + (Number(r['Trait Dough Bills']) || 0); });
  sheetRows_('Virtual Dough').forEach(function (r) { const id = id800_(r['800#']); if (id && inWin(r['When'])) dough[id] = (dough[id] || 0) + 1; });
  const dTop = top(dough);
  if (dTop && dough[dTop]) awards.push({ title: 'Trait Dough Champion', who: (names[dTop] || {}).name || 'A Mountaineer', squad: (names[dTop] || {}).squad || '', detail: dough[dTop] + ' Trait Dough bills this year' });
  SQUAD_KEYS_.forEach(function (s) {
    const mine = {}; Object.keys(dough).forEach(function (id) { if ((names[id] || {}).squad === s) mine[id] = dough[id]; });
    const t = top(mine);
    if (t && mine[t] && t !== dTop) awards.push({ title: titleSquad_(s) + ' Trait Dough Leader', who: names[t].name, squad: s, detail: mine[t] + ' bills' });
  });
  // Service.
  const svc = {};
  sheetRows_('Service Submissions').forEach(function (r) { const id = id800_(r['800#']); if (id && String(r['Status']) === 'Approved' && inWin(r['Decided At'])) svc[id] = (svc[id] || 0) + (Number(r['Points']) || 0); });
  const sTop = top(svc);
  if (sTop && svc[sTop]) awards.push({ title: 'Service Star', who: (names[sTop] || {}).name || 'A Mountaineer', squad: (names[sTop] || {}).squad || '', detail: svc[sTop] + ' service points for the squad' });
  // Poll and Pulse votes.
  const votes = {};
  sheetRows_('Poll Votes').forEach(function (v) { const id = id800_(v['800#']); if (id && inWin(v['Date'] || v['At'])) votes[id] = (votes[id] || 0) + 1; });
  const vTop = top(votes);
  if (vTop && votes[vTop]) awards.push({ title: 'Voice of the School', who: (names[vTop] || {}).name || 'A Mountaineer', squad: (names[vTop] || {}).squad || '', detail: votes[vTop] + ' poll and Pulse votes' });
  // Most improved squad: the last marking period of the season vs the first.
  const events = readEventsTab_(), mps = markingPeriods_().filter(function (m) { return m.start <= win.to && m.end >= win.from && m.start <= todayYmd_(); });
  if (mps.length >= 2) {
    const a = pointsBy_(events, mps[0].start, mps[0].end, null).points, b = pointsBy_(events, mps[mps.length - 1].start, mps[mps.length - 1].end, null).points;
    const gain = {}; SQUAD_KEYS_.forEach(function (s) { gain[s] = (b[s] || 0) - (a[s] || 0); });
    const g = top(gain);
    if (gain[g] > 0) awards.push({ title: 'Most Improved Squad', who: titleSquad_(g), squad: g, detail: '+' + gain[g] + ' points from ' + mps[0].name + ' to ' + mps[mps.length - 1].name });
  }
  // Best attendance month (Squad Monthly).
  let bestAtt = null;
  sheetRows_('Squad Monthly').forEach(function (r) {
    const v = Number(String(r['attendance']).replace('%', '')), ym = ymOf_(r['Month']);
    if (v > 0 && ym && ym + '-01' >= win.from.slice(0, 7) + '-01' && ym <= win.to.slice(0, 7) && (!bestAtt || v > bestAtt.v)) bestAtt = { v: v, s: squadOf_(r['Squad']), ym: ym };
  });
  if (bestAtt) awards.push({ title: 'Best Attendance Month', who: titleSquad_(bestAtt.s), squad: bestAtt.s, detail: bestAtt.v + '% in ' + monthLabel_(bestAtt.ym) });
  // Longest Hot Streak.
  let streak = null;
  sheetRows_('Streak Banks').forEach(function (r) { const n = Number(r['Length']) || 0, s = squadOf_(r['Squad']); if (n && (!streak || n > streak.n)) streak = { n: n, s: s }; });
  if (streak) awards.push({ title: 'Longest Hot Streak', who: titleSquad_(streak.s), squad: streak.s, detail: streak.n + ' days in a row' });
  // Never froze.
  const froze = {};
  sheetRows_('Point Freezes').forEach(function (r) { if (inWin(r['Start'])) froze[squadOf_(r['Squad'])] = 1; });
  SQUAD_KEYS_.filter(function (s) { return !froze[s]; }).forEach(function (s) { awards.push({ title: 'Never Froze', who: titleSquad_(s), squad: s, detail: 'No freeze all year' }); });
  // Most of the map.
  const held = { BLACK: 0, GOLD: 0, GREY: 0, PURPLE: 0 };
  areaHolds_(win).forEach(function (a) { SQUAD_KEYS_.forEach(function (s) { held[s] += a.all[s]; }); });
  const h = top(held);
  if (held[h]) awards.push({ title: 'King of the Map', who: titleSquad_(h), squad: h, detail: held[h] + ' area-days held' });
  return awards;
}

/** Everything the reveal show needs. Admin only (the winner can be a secret until the show). */
function clashReveal(token) {
  requireAdmin_(token);
  const win = seasonWindow_(), totals = squadTotals_();
  const order = SQUAD_KEYS_.slice().sort(function (a, b) { return totals[b] - totals[a]; });
  return JSON.stringify({
    season: seasonInfo_(totals), window: win, label: schoolYearLabel_(win.to),
    standings: order.slice().reverse().map(function (s) { return { squad: s, total: totals[s], place: order.indexOf(s) + 1 }; }),
    winner: order[0], areas: areaHolds_(win).filter(function (a) { return a.squad; }), awards: yearAwards_(win), spotlight: spotlightPublic_(6),
  });
}

function serveReveal_() {
  return HtmlService.createHtmlOutputFromFile('Reveal').setTitle('Clash Bash reveal · Clash of Classes')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function menuOpenReveal_() {
  const url = siteUrl_() + '/?page=reveal';
  SpreadsheetApp.getUi().alert('Clash Bash reveal', 'Open this on the TV or the assembly screen and sign in with the admin password:\n\n' + url +
    '\n\nClick (or press →) for each slide: every area of the school, the awards, the Spotlight, then the standings from 4th to 1st. The 🖨 Certificates button prints one certificate per award.', SpreadsheetApp.getUi().ButtonSet.OK);
}

// file-stamp (tools/stamp-files.js): lets Admin › System tell whether this file is the newest one
var FILE_STAMP_44_Reveal = '56e15fa84c';
