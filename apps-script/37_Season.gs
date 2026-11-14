/**
 * Clash of Classes -- Apps Script backend
 * File 37: the season. Squad points count toward the Clash Bash through SEASON_END (the last school
 * day of May); points dated later are still recorded but no longer change the standings. June is
 * the reveal: the site, the TV and the Morning Clash say the season is over and who won the Bash.
 *   SEASON_END   yyyy-mm-dd, the last day points count (blank = the season never ends)
 *   BASH_REVEAL  yyyy-mm-dd (optional): until this day the final standings show, but the winner
 *                is kept for the reveal ("The Clash Bash winner is revealed on ..."). Blank = right away.
 */

function seasonEnd_() { const v = String(getSetting_('SEASON_END') || '').trim().slice(0, 10); return /^\d{4}-\d{2}-\d{2}$/.test(v) ? v : ''; }

// Does a points row (or a finalized metrics month) count toward this season?
function inSeason_(ymd) { const end = seasonEnd_(); return !end || !ymd || String(ymd).slice(0, 10) <= end; }

/** { end, over, daysLeft (calendar), reveal, revealed, winner } for the page, the TV and the Morning Clash. */
function seasonInfo_(totals) {
  const end = seasonEnd_(), today = todayYmd_();
  if (!end) return { end: '', over: false };
  const reveal = String(getSetting_('BASH_REVEAL') || '').trim().slice(0, 10);
  const over = today > end;
  const out = { end: end, over: over, reveal: /^\d{4}-\d{2}-\d{2}$/.test(reveal) ? reveal : '' };
  out.daysLeft = over ? 0 : Math.round((new Date(end + 'T12:00:00Z') - new Date(today + 'T12:00:00Z')) / 86400000);
  out.revealed = over && (!out.reveal || today >= out.reveal);
  if (out.revealed) {
    totals = totals || squadTotals_();
    out.winner = SQUAD_KEYS_.slice().sort(function (a, b) { return totals[b] - totals[a]; })[0];
  }
  return out;
}

// file-stamp (tools/stamp-files.js): lets Admin › System tell whether this file is the newest one
var FILE_STAMP_37_Season = 'd68c5a193b';
