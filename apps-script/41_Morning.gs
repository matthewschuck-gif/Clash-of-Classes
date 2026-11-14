/**
 * Clash of Classes -- Apps Script backend
 * File 41: Morning Clash (homeroom) and the Squad Spotlight.
 *
 * MORNING CLASH: clashofclasses.org/?tv=morning (or ?page=morning). A short deck that plays by itself
 * (TV_SECONDS a slide) for homeroom: good morning + the announcement, today's poll with the vote link,
 * the standings (with the season and any Avalanche), this month's trait, and the Squad Spotlight.
 *
 * SQUAD SPOTLIGHT tab: Show | Week Of | Squad | Who | Story | From. Every Monday the site suggests one
 * story per squad from the week before (Virtual Trait Dough bills and approved service), unticked.
 * Staff tick Show (and edit the story if they like); shown rows go on the Morning Clash, the TV and
 * the home page. Names are first name + last initial only.
 */

const SPOTLIGHT_HEADERS_ = ['Show', 'Week Of', 'Squad', 'Who', 'Story', 'From'];

function ensureSpotlightTab_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getSheetByName('Squad Spotlight')) return;
  const sh = ss.insertSheet('Squad Spotlight');
  sh.getRange(1, 1, 1, SPOTLIGHT_HEADERS_.length).setValues([SPOTLIGHT_HEADERS_]).setFontWeight('bold').setBackground('#490e6f').setFontColor('#ffffff');
  sh.setFrozenRows(1);
  try { sh.getRange(2, 1, 300, 1).insertCheckboxes(); } catch (e) { /* fine */ }
}

function shortName_(full) { const p = String(full || '').trim().split(/\s+/); return p[0] ? p[0] + (p.length > 1 ? ' ' + p[p.length - 1].charAt(0) + '.' : '') : ''; }
function mondayOf_(ymd) { const d = new Date(ymd + 'T12:00:00Z'), dow = d.getUTCDay(); return ymdShift_(ymd, dow === 0 ? -6 : 1 - dow); }

/** Suggests this week's stories (one per squad, from last week). Returns how many were added. */
function suggestSpotlights_() {
  ensureSpotlightTab_();
  const week = mondayOf_(todayYmd_()), from = ymdShift_(week, -7), to = ymdShift_(week, -1);
  const have = sheetRows_('Squad Spotlight');
  if (have.some(function (r) { return String(r['Week Of']).slice(0, 10) === week; })) return 0;
  const used = {};
  have.forEach(function (r) { used[String(r['Story'] || '').trim()] = 1; });
  const pool = [];
  sheetRows_('Virtual Dough').forEach(function (r) {
    const d = String(r['When']).slice(0, 10);
    if (d >= from && d <= to) pool.push({ squad: squadOf_(r['Squad']), who: shortName_(r['Student']), story: String(r['What They Did'] || '').trim() + (r['Trait'] ? ' (' + r['Trait'] + ')' : ''), from: 'Virtual Trait Dough from ' + r['Teacher'] });
  });
  sheetRows_('Service Submissions').forEach(function (r) {
    const d = String(r['Decided At']).slice(0, 10);
    if (String(r['Status']) === 'Approved' && d >= from && d <= to) pool.push({ squad: squadOf_(r['Squad']), who: shortName_(r['Student']), story: String(r['What'] || '').trim(), from: 'Service' });
  });
  const sh = getSheet_('Squad Spotlight');
  let n = 0;
  SQUAD_KEYS_.forEach(function (s) {
    const pick = pool.filter(function (x) { return x.squad === s && x.story && !used[x.story]; }).sort(function (a, b) { return b.story.length - a.story.length; })[0];
    if (!pick) return;
    sh.appendRow([false, week, s, pick.who, pick.story, pick.from]);
    n++;
  });
  return n;
}

function maybeSuggestSpotlights_() {
  const props = PropertiesService.getScriptProperties(), week = mondayOf_(todayYmd_());
  if (props.getProperty('spotlight-week') === week) return;
  props.setProperty('spotlight-week', week);
  suggestSpotlights_();
}

/** Shown stories, newest week first. */
function spotlightPublic_(limit) {
  return sheetRows_('Squad Spotlight').filter(function (r) { return ticked_(r['Show']) && String(r['Story'] || '').trim(); })
    .map(function (r) { return { week: String(r['Week Of']).slice(0, 10), squad: squadOf_(r['Squad']), who: shortName_(r['Who']), story: String(r['Story']).trim() }; })
    .sort(function (a, b) { return b.week.localeCompare(a.week); }).slice(0, limit || 4);
}

function clashSpotlight() { return JSON.stringify(spotlightPublic_(4)); }

/** Everything the Morning Clash shows. */
function clashMorning() {
  const totals = squadTotals_(), today = todayYmd_();
  let announcement = '', board = {}, trait = null, avalanche = null;
  try { announcement = announcementScript_(); } catch (e) { announcement = ''; }
  try { board = JSON.parse(clashPollBoard()); } catch (e) { board = {}; }
  try {
    const month = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'][Number(today.slice(5, 7)) - 1];
    const t = (siteTraits_() || []).filter(function (x) { return x.month === month; })[0];
    if (t) trait = { month: t.month, trait: t.trait, desc: t.desc, skill: t.skills.length ? t.skills[(Number(today.slice(8)) || 0) % t.skills.length] : null };
  } catch (e) { trait = null; }
  try { avalanche = avalancheBoard_(totals); } catch (e) { avalanche = null; }
  return JSON.stringify({
    date: today, announcement: announcement, polls: board.polls || [], live: !!board.live, next: board.next || [], voteText: board.voteText || voteText_(),
    standings: SQUAD_KEYS_.slice().sort(function (a, b) { return totals[b] - totals[a]; }).map(function (s) { return { squad: s, total: totals[s] }; }),
    season: seasonInfo_(totals), avalanche: avalanche, trait: trait, spotlight: spotlightPublic_(4),
    seconds: Math.max(4, Math.min(60, settingNum_('TV_SECONDS', 9))) + 3,
  });
}

function serveMorning_() {
  return HtmlService.createHtmlOutputFromFile('Morning').setTitle('Morning Clash · Clash of Classes')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// file-stamp (tools/stamp-files.js): lets Admin › System tell whether this file is the newest one
var FILE_STAMP_41_Morning = '1704655998';
