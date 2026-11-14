/**
 * Clash of Classes -- Apps Script backend
 * File 29: the Points Balance check. Is one source of points deciding the race?
 *
 * Looks at the last BALANCE_DAYS (30) days of points by type (Trait Dough, polls, squad events,
 * STAR growth...) and writes the POINTS BALANCE tab (staff only):
 *   Type | Label | BLACK | GOLD | GREY | PURPLE | Total | Share % | Flag
 * Flags (every line is a Setting):
 *   - one type is more than BALANCE_MAX_SHARE % (40) of every point given
 *   - the leader's lead over second place comes mostly (BALANCE_LEAD_SHARE %, 60) from one type
 *   - only one squad earned a type that gave at least 50 points (do the others know about it?)
 *   - one row gave a squad more than BALANCE_BIG_ROW (250) points at once
 * Game bonuses (territory, boosts, cards...) are listed but never flagged as "one source".
 * Private behavior rows never count. Counts only, no names. Runs from the menu and goes into the
 * monthly State of the Squads email (31_StateOfSquads.gs).
 */

const BALANCE_HEADERS_ = ['Type', 'Label', 'BLACK', 'GOLD', 'GREY', 'PURPLE', 'Total', 'Share %', 'Flag'];

function settingNum_(key, def) {
  const t = String(getSetting_(key) == null ? '' : getSetting_(key)).trim();
  return t === '' || isNaN(Number(t)) ? def : Number(t);
}

/** { from, to, days, total, squads, types: [{ type, label, pts, total, share, flag }], flags: [lines], leader } */
function pointsBalance_(events, to, days) {
  events = events || readEventsTab_();
  to = to || todayYmd_();
  days = days || settingNum_('BALANCE_DAYS', 30);
  const from = ymdShift_(to, -days + 1);
  const priv = listOf_(getSetting_('PRIVATE_EVENT_TYPES') || 'trail_journal, trailback_panel, camp_mountaineer');
  const labels = parsePairsSetting_('EVENT_LABELS');
  const big = settingNum_('BALANCE_BIG_ROW', 250);
  const by = {}, squads = { BLACK: 0, GOLD: 0, GREY: 0, PURPLE: 0 }, flags = [];
  events.forEach(function (e) {
    const d = String(e.date || '').slice(0, 10);
    const type = String(e.type || '').trim().toLowerCase() || 'event';
    if (!d || d < from || d > to || priv.indexOf(type) > -1) return;
    const t = by[type] = by[type] || { type: type, label: labels[type] || type.replace(/_/g, ' '), pts: { BLACK: 0, GOLD: 0, GREY: 0, PURPLE: 0 }, total: 0 };
    SQUAD_KEYS_.forEach(function (s) {
      const p = Number(e.pts && e.pts[s]) || 0;
      if (p <= 0) return;
      t.pts[s] += p; t.total += p; squads[s] += p;
      if (p > big && GAME_TYPES_.indexOf(type) === -1) flags.push('One row gave ' + titleSquad_(s) + ' ' + p + ' points at once: "' + publicEvent_(e).name + '" (' + d + ').');
    });
  });
  const total = SQUAD_KEYS_.reduce(function (n, s) { return n + squads[s]; }, 0);
  const types = Object.keys(by).map(function (k) { return by[k]; }).sort(function (a, b) { return b.total - a.total; });
  const maxShare = settingNum_('BALANCE_MAX_SHARE', 40), leadShare = settingNum_('BALANCE_LEAD_SHARE', 60);
  types.forEach(function (t) {
    t.share = total ? Math.round(1000 * t.total / total) / 10 : 0;
    t.flag = '';
    if (GAME_TYPES_.indexOf(t.type) > -1) return;
    if (t.share > maxShare) {
      t.flag = 'One source: ' + t.share + '% of all points';
      flags.push(t.label + ' is ' + t.share + '% of every point in the last ' + days + ' days: one source is deciding the race.');
    }
    const earners = SQUAD_KEYS_.filter(function (s) { return t.pts[s] > 0; });
    if (earners.length === 1 && t.total >= 50) {
      t.flag = t.flag || 'Only ' + titleSquad_(earners[0]) + ' earned it';
      flags.push('Only ' + titleSquad_(earners[0]) + ' earned ' + t.label + ' (' + t.total + ' points). Do the other squads know about it?');
    }
  });
  const place = SQUAD_KEYS_.slice().sort(function (a, b) { return squads[b] - squads[a]; });
  const lead = squads[place[0]] - squads[place[1]];
  if (lead > 0) {
    const edge = types.filter(function (t) { return GAME_TYPES_.indexOf(t.type) === -1; })
      .map(function (t) { return { t: t, gap: t.pts[place[0]] - t.pts[place[1]] }; })
      .sort(function (a, b) { return b.gap - a.gap; })[0];
    if (edge && edge.gap > 0 && 100 * edge.gap / lead > leadShare) {
      flags.push(titleSquad_(place[0]) + '\'s lead over ' + titleSquad_(place[1]) + ' (' + lead + ' points) is mostly ' + edge.t.label + ' (' + edge.gap + ' of it).');
    }
  }
  return { from: from, to: to, days: days, total: total, squads: squads, types: types, flags: flags, leader: place[0] };
}

/** Writes the Points Balance tab. Returns the report. */
function writePointsBalance_() {
  const r = pointsBalance_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('Points Balance');
  if (!sh) {
    sh = ss.insertSheet('Points Balance');
    sh.getRange(1, 1, 1, BALANCE_HEADERS_.length).setValues([BALANCE_HEADERS_]).setFontWeight('bold').setBackground('#490e6f').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  }
  writeRowsByHeader_(sh, r.types.map(function (t) {
    return { 'Type': t.type, 'Label': t.label, 'BLACK': t.pts.BLACK, 'GOLD': t.pts.GOLD, 'GREY': t.pts.GREY, 'PURPLE': t.pts.PURPLE, 'Total': t.total, 'Share %': t.share, 'Flag': t.flag };
  }));
  return r;
}

function menuPointsBalance_() {
  const r = writePointsBalance_();
  SpreadsheetApp.getUi().alert('Points balance: ' + r.from + ' to ' + r.to,
    (r.flags.length ? '⚠ ' + r.flags.join('\n⚠ ') : 'Nothing looks lopsided: no one source decides the race.') +
    '\n\nEvery type, by squad, is on the Points Balance tab. Change the lines on Settings (BALANCE_*).', SpreadsheetApp.getUi().ButtonSet.OK);
}
