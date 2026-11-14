/**
 * Clash of Classes -- Apps Script backend
 * File 42: the monthly family update. clashofclasses.org/?page=family (&squad=gold for one squad;
 * &month=2026-12 for another month). A one-page, printable letter home about last month: each squad's
 * points and place, the areas of the school it holds, what students told the school in the Pulse and
 * what changed because of it, the Squad Spotlight, and what's coming up. Counts only; names are first
 * name + last initial, and only stories staff chose for the Spotlight. Link it in the family newsletter.
 */

function clashFamily(month, squad) {
  const today = todayYmd_();
  month = /^\d{4}-\d{2}$/.test(String(month || '')) ? month : lastMonthWindow_(today).month;
  squad = squadOf_(squad);
  const win = monthWindow_(month);
  const events = readEventsTab_();
  const pts = pointsBy_(events, win.from, win.to, null).points, totals = squadTotals_();
  const place = SQUAD_KEYS_.slice().sort(function (a, b) { return totals[b] - totals[a]; });
  let zones = [];
  try { zones = mapZones_(events, false).filter(function (z) { return z.holder; }).map(function (z) { return { zone: z.zone, holder: z.holder, rule: z.rule }; }); } catch (e) { zones = []; }
  const pulses = pollsWithVotes_().filter(function (p) { return p.kind === 'Pulse' && p.question && String(p.through).slice(0, 7) === month && p.total; })
    .map(function (p) { return { question: p.question, said: youSaidLine_(p) }; });
  let youSaid = [];
  try { youSaid = youSaidPublic_(6).filter(function (y) { return y.closed.slice(0, 7) <= month; }); } catch (e) { youSaid = []; }
  const next = monthWindow_(ymdShift_(win.to, 1).slice(0, 7));
  let coming = [];
  try {
    coming = sheetRows_('Year Plan').filter(function (r) { return ticked_(r['Show']) && r['Title'] && String(r['Date']).slice(0, 10) >= next.from && String(r['Date']).slice(0, 10) <= next.to; })
      .slice(0, 8).map(function (r) { return { date: String(r['Date']).slice(0, 10), title: String(r['Title']), kind: String(r['Kind'] || '') }; });
  } catch (e) { coming = []; }
  let spotlight = [];
  try { spotlight = spotlightPublic_(8).filter(function (s) { return s.week.slice(0, 7) === month || s.week.slice(0, 7) === next.from.slice(0, 7); }); } catch (e) { spotlight = []; }
  return JSON.stringify({
    month: month, label: monthLabel_(month), squad: SQUAD_KEYS_.indexOf(squad) > -1 ? squad : '', site: siteUrl_().replace(/^https?:\W+/i, ''),
    squads: SQUAD_KEYS_.map(function (s) { return { squad: s, month: pts[s] || 0, total: totals[s], place: place.indexOf(s) + 1, zones: zones.filter(function (z) { return z.holder === s; }).map(function (z) { return z.zone; }) }; }),
    pulses: pulses, youSaid: youSaid, spotlight: spotlight, coming: coming, season: seasonInfo_(totals),
  });
}

function serveFamily_() {
  return HtmlService.createHtmlOutputFromFile('Family').setTitle('Clash of Classes · Family update')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// file-stamp (tools/stamp-files.js): lets Admin › System tell whether this file is the newest one
var FILE_STAMP_42_Family = '9a5382c8d3';
