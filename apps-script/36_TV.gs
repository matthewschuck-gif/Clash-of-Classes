/**
 * Clash of Classes -- Apps Script backend
 * File 36: the TV. Three links, all live:
 *   clashofclasses.org/?tv=1            the full rotation (standings, podium, categories, metrics,
 *                                       challenges, latest points, summit, map, the Opponent, the
 *                                       Pulse, today's polls with live bars)
 *   clashofclasses.org/?tv=polls        polls only: today's polls, the Pulse and the standings
 *                                       (homeroom mornings)
 *   clashofclasses.org/?page=pollboard  the Poll Board: one poll at a time, big bars, squad turnout
 *                                       and the live standings along the bottom
 * Every poll screen shows the short link clashofclasses.org/vote (phones aren't allowed in the building,
 * so students type it on a Chromebook). It opens the site at today's poll (the vote/ page in the repo).
 * Settings (edit in Admin › Sheet › Settings):
 *   TV_SECONDS     seconds per slide (default 9)
 *   TV_SLIDES      which slides run: all, or a list of standings, podium, categories, metrics,
 *                  challenges, feed, summit, map, opponent, pulse, polls
 *   TV_POLL_SLIDES the slides for ?tv=polls (default polls, pulse, standings)
 */

const TV_SLIDE_NAMES_ = ['standings', 'podium', 'categories', 'metrics', 'challenges', 'feed', 'summit', 'map', 'opponent', 'pulse', 'polls'];

function tvSlideList_(spec, def) {
  const words = listOf_(spec || def);
  if (!words.length || words.indexOf('all') > -1) return TV_SLIDE_NAMES_.slice();
  const keep = TV_SLIDE_NAMES_.filter(function (n) { return words.indexOf(n) > -1; });
  return keep.length ? keep : TV_SLIDE_NAMES_.slice();
}

// The short vote link as students should type it: clashofclasses.org/vote
function voteText_() { return siteUrl_().replace(/^https?:\W+/i, '') + '/vote'; }

function siteUrl_() { return String(getSetting_('SITE_URL') || 'https:\/\/clashofclasses.org').trim().replace(/\/+$/, ''); }

/** What the page needs to run the TV (rides along with the site data). */
function tvSettings_() {
  return {
    seconds: Math.max(4, Math.min(60, settingNum_('TV_SECONDS', 9))),
    slides: tvSlideList_(getSetting_('TV_SLIDES'), 'all'),
    pollSlides: tvSlideList_(getSetting_('TV_POLL_SLIDES'), 'polls, pulse, standings'),
    siteUrl: siteUrl_(),
    voteUrl: siteUrl_() + '/vote',
    voteText: voteText_(),
  };
}

/** Standings for the TV and the Poll Board: total points and what each squad earned today. */
function tvStandings_() {
  const totals = squadTotals_(), today = todayYmd_();
  const gained = pointsBy_(readEventsTab_(), today, today, null).points;
  return SQUAD_KEYS_.slice().sort(function (a, b) { return totals[b] - totals[a]; })
    .map(function (s) { return { squad: s, total: totals[s], today: gained[s] || 0 }; });
}

// file-stamp (tools/stamp-files.js): lets Admin › System tell whether this file is the newest one
var FILE_STAMP_36_TV = 'a801d2f28e';
