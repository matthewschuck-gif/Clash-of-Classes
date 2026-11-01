/**
 * Clash of Classes -- Apps Script backend
 * File 9: Lore Chapters -- each squad's story, told in chapters it unlocks.
 *
 * Lore Chapters tab (one row per chapter; write drafts any time, nothing shows until Show):
 *   Show | Squad | Chapter | Title | Story | Image | Unlock When | Unlocked
 *   Squad: BLACK, GOLD, GREY, PURPLE, or ALL (the same chapter for every squad, each unlocks its own).
 *   Unlock When:
 *     blank                    unlocked from the start
 *     challenge: <title>       when the squad reaches that challenge's goal (Challenges tab)
 *     duel: <title>            when the squad wins that Squad Duel
 *     points: 500              when the squad's total reaches 500
 *     date: 2026-12-01         on that day
 *   Unlocked is filled in by itself ("GOLD 2026-11-03, GREY 2026-11-05"), the first time each
 *   squad's chapter is seen unlocked. Type a squad and date there to unlock one by hand.
 *
 * My Squad shows the squad's unlocked chapters and what unlocks the next one; the TV shows
 * chapters unlocked in the last 7 days; the Generals' meeting slides show the newest one.
 */

const LORE_HEADERS_ = ['Show', 'Squad', 'Chapter', 'Title', 'Story', 'Image', 'Unlock When', 'Unlocked'];

function ensureLoreTab_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getSheetByName('Lore Chapters')) return;
  const sh = ss.insertSheet('Lore Chapters');
  sh.getRange(1, 1, 3, LORE_HEADERS_.length).setValues([LORE_HEADERS_,
    [false, 'ALL', 1, 'Chapter 1 (draft)', 'Write the opening of the story here.', '', '', ''],
    [false, 'ALL', 2, 'Chapter 2 (draft)', 'This one unlocks when a squad reaches 500 points.', '', 'points: 500', '']]);
  sh.getRange(1, 1, 1, LORE_HEADERS_.length).setFontWeight('bold').setBackground('#490e6f').setFontColor('#ffffff');
  sh.setFrozenRows(1);
}

function parseUnlocked_(v) {
  const out = {};
  String(v || '').split(',').forEach(function (part) {
    const m = part.trim().toUpperCase().replace('GRAY', 'GREY').match(/^([A-Z]+)\s+(\d{4}-\d{2}-\d{2})/);
    if (m && SQUAD_KEYS_.indexOf(m[1]) > -1) out[m[1]] = m[2];
  });
  return out;
}

// Plain-words reason a chapter is still locked, for My Squad.
function unlockHint_(rule) {
  const m = String(rule || '').match(/^\s*(challenge|duel|points|date)\s*:\s*(.+)$/i);
  if (!m) return 'Unlocks soon';
  const k = m[1].toLowerCase(), v = m[2].trim();
  return k === 'challenge' ? 'Unlocks when your squad finishes the "' + v + '" challenge'
    : k === 'duel' ? 'Unlocks when your squad wins the "' + v + '" duel'
    : k === 'points' ? 'Unlocks at ' + Number(v).toLocaleString() + ' squad points'
    : 'Unlocks on ' + v;
}

/**
 * Every shown chapter with each squad's unlock date: [{ row, squads: [...], chapter, title, story,
 * image, rule, unlocked: {SQUAD: date} }]. Writes newly reached unlocks to the Unlocked column.
 */
function loreChapters_() {
  const rows = sheetRows_('Lore Chapters').filter(function (r) { return ticked_(r['Show']) && String(r['Title'] || r['Story'] || '').trim(); });
  if (!rows.length) return [];
  const today = todayYmd_();
  let totals = null, challenges = null, duels = null;
  const reached = function (rule, squad) {
    const m = String(rule || '').match(/^\s*(challenge|duel|points|date)\s*:\s*(.+)$/i);
    if (!String(rule || '').trim()) return true;
    if (!m) return false;
    const k = m[1].toLowerCase(), v = m[2].trim().toLowerCase();
    if (k === 'points') { totals = totals || squadTotals_(); return totals[squad] >= Number(v); }
    if (k === 'date') return today >= v;
    if (k === 'challenge') {
      challenges = challenges || challengeProgress_();
      return challenges.some(function (c) { return c.title.toLowerCase() === v && c.goal && c.progress[squad] >= c.goal; });
    }
    duels = duels || duelProgress_();
    return duels.some(function (d) { return d.title.toLowerCase() === v && d.done && d.leaders.indexOf(squad) > -1; });
  };
  const sheet = getSheetSafe_('Lore Chapters');
  const col = LORE_HEADERS_.indexOf('Unlocked') + 1;
  return rows.map(function (r) {
    const sq = squadOf_(r['Squad']);
    const squads = SQUAD_KEYS_.indexOf(sq) > -1 ? [sq] : SQUAD_KEYS_.slice();
    const unlocked = parseUnlocked_(r['Unlocked']);
    let changed = false;
    squads.forEach(function (s) { if (!unlocked[s] && reached(r['Unlock When'], s)) { unlocked[s] = today; changed = true; } });
    if (changed && sheet) sheet.getRange(r._row, col).setValue(Object.keys(unlocked).map(function (s) { return s + ' ' + unlocked[s]; }).join(', '));
    const image = String(r['Image'] || '').trim();
    return { row: r._row, squads: squads, chapter: Number(r['Chapter']) || 0, title: String(r['Title'] || '').trim(), story: String(r['Story'] || ''),
      image: /^https:\/\//i.test(image) ? image : '', rule: String(r['Unlock When'] || ''), unlocked: unlocked };
  }).sort(function (a, b) { return a.chapter - b.chapter; });
}

// One squad's story: { chapters: [unlocked…], next: { chapter, hint } or null }.
function loreFor_(squad) {
  const mine = loreChapters_().filter(function (c) { return c.squads.indexOf(squad) > -1; });
  const open = mine.filter(function (c) { return c.unlocked[squad]; });
  const locked = mine.filter(function (c) { return !c.unlocked[squad]; })[0];
  return {
    chapters: open.map(function (c) { return { chapter: c.chapter, title: c.title, story: c.story, image: c.image, unlocked: c.unlocked[squad] }; }),
    next: locked ? { chapter: locked.chapter, hint: unlockHint_(locked.rule) } : null,
  };
}

// For the TV: chapters unlocked in the last 7 days, newest first. Titles only, no story.
function recentLoreUnlocks_() {
  const weekAgo = Utilities.formatDate(new Date(Date.now() - 7 * 86400000), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const out = [];
  loreChapters_().forEach(function (c) {
    Object.keys(c.unlocked).forEach(function (s) {
      if (c.rule.trim() && c.unlocked[s] >= weekAgo) out.push({ squad: s, chapter: c.chapter, title: c.title, date: c.unlocked[s] });
    });
  });
  return out.sort(function (a, b) { return a.date < b.date ? 1 : -1; }).slice(0, 4);
}
