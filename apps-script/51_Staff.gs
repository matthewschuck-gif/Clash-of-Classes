/**
 * Clash of Classes -- Apps Script backend
 * File 51: helpers for staff and Generals.
 *   QUIET CLIMBERS   students with nothing in QUIET_DAYS days (no Trait Dough, service, Arcade or
 *                    contribution on the Student Board). Staff only: Admin › System and the Friday
 *                    wrap-up. Names are first name + last initial.
 *   CAPTAINS' TO-DO  three suggested rally moves on Generals HQ, from where the squad is behind.
 *   DOUGH REACH      on the Virtual Trait Dough page: which squads a teacher has given to this month,
 *                    and a few students in the squad they've reached least.
 *   FRIDAY WRAP-UP   a short staff email (WRAP_EMAILS, or STATE_EMAILS / DIGEST_EMAILS). It is never
 *                    sent by itself: preview and send it from Admin › System (or the menu).
 *   ANNOUNCEMENT     clashofclasses.org/?page=announce: today's morning announcement, big, to read
 *                    over the PA (Copy and Print buttons).
 */

function quietDays_() { return Math.max(7, settingNum_('QUIET_DAYS', 21)); }

/** { BLACK: [{ id, name, grade, since }], ... } */
function quietClimbers_() {
  const since = ymdShift_(todayYmd_(), -quietDays_()), active = {};
  const mark = function (id) { id = id800_(id); if (id) active[id] = 1; };
  sheetRows_('Virtual Dough').forEach(function (r) { if (String(r['When']).slice(0, 10) >= since) mark(r['800#']); });
  sheetRows_('Service Submissions').forEach(function (r) { if (String(r['Submitted']).slice(0, 10) >= since) mark(r['800#']); });
  sheetRows_('Arcade Plays').forEach(function (r) { if (String(r['When']).slice(0, 10) >= since) mark(r['800#']); });
  const out = { BLACK: [], GOLD: [], GREY: [], PURPLE: [] };
  boardRows_().forEach(function (r) {
    const id = id800_(r['800#']), s = squadOf_(r['Squad']), last = asYmd_(r['Last Contribution']).slice(0, 10);
    if (!id || !out[s] || active[id] || (last && last >= since) || (Number(r['Bills This Week']) || 0) > 0) return;
    out[s].push({ id: id, name: shortName_(r['Name']), grade: String(r['Grade'] || ''), since: last });
  });
  SQUAD_KEYS_.forEach(function (s) { out[s].sort(function (a, b) { return (a.since || '0').localeCompare(b.since || '0') || a.name.localeCompare(b.name); }); });
  return out;
}

// ── Captains' to-do (Generals HQ) ──
/** Up to three suggestions: [{ icon, text }]. Squad sizes: squadSizes_ (18_Conquest.gs). */
function captainTodo_(squad) {
  const out = [], today = todayYmd_(), mon = mondayOf_(today), size = squadSizes_(today.slice(0, 7));
  // Today's weather first: it's the easiest rally.
  try { const w = weatherNow_().now; if (w) out.push({ icon: '🌨️', text: 'It\'s ' + w.name + ' today: ' + w.message + ' Get the word out at lunch.' }); } catch (e) { /* none */ }
  // Poll turnout this week.
  try {
    const voters = { BLACK: {}, GOLD: {}, GREY: {}, PURPLE: {} };
    sheetRows_('Poll Votes').forEach(function (v) { const s = squadOf_(v['Squad']), d = String(v['Date'] || v['At']).slice(0, 10); if (voters[s] && d >= mon) voters[s][id800_(v['800#'])] = 1; });
    const pct = {}; SQUAD_KEYS_.forEach(function (s) { pct[s] = size[s] ? Math.round(100 * Object.keys(voters[s]).length / size[s]) : 0; });
    const best = SQUAD_KEYS_.slice().sort(function (a, b) { return pct[b] - pct[a]; })[0];
    if (best !== squad && pct[best] - pct[squad] >= 5) out.push({ icon: '🗳️', text: pct[squad] + '% of ' + titleSquad_(squad) + ' voted in the Daily Poll this week (' + titleSquad_(best) + ': ' + pct[best] + '%). Remind your homerooms in the morning.' });
  } catch (e) { /* no polls */ }
  // An expedition you could win.
  try {
    expeditionsPublic_().filter(function (x) { return x.open; }).forEach(function (x) {
      const lead = SQUAD_KEYS_.slice().sort(function (a, b) { return x.progress[b] - x.progress[a]; })[0], mine = x.progress[squad];
      out.push({ icon: '🧭', text: 'Expedition ' + x.title + ': ' + titleSquad_(squad) + ' has ' + mine + ' of ' + x.goal + (lead !== squad && x.progress[lead] > mine ? ' (' + titleSquad_(lead) + ' has ' + x.progress[lead] + ')' : ', in the lead') + '. ' + (x.what || '') });
    });
  } catch (e) { /* none */ }
  // Service this month.
  try {
    const ym = today.slice(0, 7);
    const svc = sheetRows_('Service Submissions').filter(function (r) { return squadOf_(r['Squad']) === squad && String(r['Submitted']).slice(0, 7) === ym; }).length;
    if (!svc) out.push({ icon: '🤝', text: 'Nobody in ' + titleSquad_(squad) + ' has sent in service this month. Ask a few members to log something they did on My Climb.' });
  } catch (e) { /* none */ }
  // Quiet members.
  try {
    const q = quietClimbers_()[squad].length;
    if (q >= 3) out.push({ icon: '👋', text: q + ' ' + titleSquad_(squad) + ' members haven\'t earned anything in ' + quietDays_() + ' days. Invite them to a sign-up or the Arcade.' });
  } catch (e) { /* none */ }
  return out.slice(0, 3);
}

// ── Dough reach (Virtual Trait Dough page) ──
function clashDoughReach(token) {
  const me = doughSession_(token), ym = todayYmd_().slice(0, 7), others = SQUAD_KEYS_.filter(function (s) { return s !== me.squad; });
  const mine = sheetRows_('Virtual Dough').filter(function (r) { return String(r['Teacher']) === me.name && String(r['When']).slice(0, 7) === ym; });
  const by = {}, given = {};
  others.forEach(function (s) { by[s] = 0; });
  mine.forEach(function (r) { const s = squadOf_(r['Squad']); if (by[s] !== undefined) by[s]++; given[id800_(r['800#'])] = 1; });
  const low = others.slice().sort(function (a, b) { return by[a] - by[b]; })[0];
  // A few students in that squad nobody has given Virtual Dough to this month, quiet ones first.
  const got = {};
  sheetRows_('Virtual Dough').forEach(function (r) { if (String(r['When']).slice(0, 7) === ym) got[id800_(r['800#'])] = 1; });
  const quiet = {}; (quietClimbers_()[low] || []).forEach(function (q) { quiet[q.id] = 1; });
  const pool = sheetRows_('Students').filter(function (r) { const id = id800_(r['800#']); return squadOf_(r['Squad']) === low && id && !got[id] && !given[id]; })
    .map(function (r) { return { id: id800_(r['800#']), name: String(r['First Name'] || '').trim() + ' ' + String(r['Last Name'] || '').trim(), grade: String(r['Grade'] || ''), squad: low, quiet: !!quiet[id800_(r['800#'])] }; })
    .sort(function (a, b) { return (b.quiet ? 1 : 0) - (a.quiet ? 1 : 0) || a.name.localeCompare(b.name); });
  // Rotate by day so the same names don't always come up.
  const start = pool.length ? (Number(todayYmd_().slice(8)) * 3) % pool.length : 0;
  const pick = pool.filter(function (p) { return p.quiet; }).slice(0, 2).concat(pool.slice(start).concat(pool.slice(0, start)).filter(function (p) { return !p.quiet; })).slice(0, 4);
  return JSON.stringify({ bySquad: by, students: Object.keys(given).length, total: mine.length, low: low, suggest: pick });
}

// ── Friday wrap-up ──
function wrapRecipients_() {
  const raw = String(getSetting_('WRAP_EMAILS') || '').trim();
  if (!raw) return stateRecipients_();
  return raw.split(/[,;\s]+/).map(function (s) { return s.trim().toLowerCase(); }).filter(function (s) { return /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(s); });
}

function wrapData_() {
  const today = todayYmd_(), mon = mondayOf_(today), events = readEventsTab_(), totals = squadTotals_();
  const week = { BLACK: 0, GOLD: 0, GREY: 0, PURPLE: 0 };
  events.forEach(function (e) { const d = String(e.date || '').slice(0, 10); if (d >= mon && d <= today) SQUAD_KEYS_.forEach(function (s) { week[s] += Number(e.pts[s]) || 0; }); });
  const top = events.filter(function (e) { const d = String(e.date || '').slice(0, 10); return d >= mon && d <= today; })
    .map(function (e) { return { name: publicEvent_(e).name, pts: SQUAD_KEYS_.reduce(function (n, s) { return n + (Number(e.pts[s]) || 0); }, 0) }; })
    .sort(function (a, b) { return b.pts - a.pts; }).slice(0, 5);
  const dough = sheetRows_('Virtual Dough').filter(function (r) { return String(r['When']).slice(0, 10) >= mon; });
  const teachers = {}; dough.forEach(function (r) { teachers[String(r['Teacher'])] = (teachers[String(r['Teacher'])] || 0) + 1; });
  const quiet = quietClimbers_();
  const nextWeek = [mondayOf_(ymdShift_(today, 7)), ymdShift_(mondayOf_(ymdShift_(today, 7)), 6)];
  const coming = [];
  try { weatherRows_().forEach(function (w) { if (w.from >= nextWeek[0] && w.from <= nextWeek[1]) coming.push(w.from + ': ' + w.name); }); } catch (e) { /* none */ }
  try { sheetRows_('Sign-up Events').forEach(function (r) { const d = asYmd_(r['Date']); if (d >= nextWeek[0] && d <= nextWeek[1]) coming.push(d + ': ' + String(r['Event'])); }); } catch (e) { /* none */ }
  try { const c = colorsDayNext_(); if (c && c.date <= nextWeek[1]) coming.push(c.date + ': ' + c.name); } catch (e) { /* none */ }
  try { expeditionsPublic_().forEach(function (x) { if (x.open && x.end && x.end <= nextWeek[1]) coming.push(x.end + ': the ' + x.title + ' expedition ends'); }); } catch (e) { /* none */ }
  return { mon: mon, today: today, week: week, totals: totals, top: top, dough: dough.length, teachers: teachers, quiet: quiet,
    spotlight: spotlightPublic_(8).filter(function (s) { return s.week >= mon; }), expeditions: (function () { try { return expeditionsPublic_(); } catch (e) { return []; } })(), coming: coming.sort() };
}

function wrapHtml_(d) {
  const esc = function (s) { return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); };
  const order = SQUAD_KEYS_.slice().sort(function (a, b) { return d.week[b] - d.week[a]; });
  const h2 = function (t) { return '<h3 style="margin:18px 0 6px;color:#490e6f">' + t + '</h3>'; };
  let html = '<div style="font-family:Arial,sans-serif;max-width:640px;font-size:15px;line-height:1.45">' +
    '<h2 style="margin:0;color:#490e6f">Clash Friday wrap-up</h2><div style="color:#666">Week of ' + esc(d.mon) + '</div>' +
    '<p><b>' + titleSquad_(order[0]) + '</b> won the week with ' + d.week[order[0]] + ' points.</p>' +
    '<table style="border-collapse:collapse"><tr><th align="left" style="padding:4px 12px 4px 0">Squad</th><th align="right" style="padding:4px 12px">This week</th><th align="right" style="padding:4px 0">Total</th></tr>' +
    order.map(function (s) { return '<tr><td style="padding:3px 12px 3px 0">' + titleSquad_(s) + '</td><td align="right" style="padding:3px 12px">' + d.week[s] + '</td><td align="right">' + (d.totals[s] || 0).toLocaleString() + '</td></tr>'; }).join('') + '</table>';
  if (d.top.length) html += h2('Biggest moments') + '<ul>' + d.top.map(function (t) { return '<li>' + esc(t.name) + ' (' + t.pts + ')</li>'; }).join('') + '</ul>';
  if (d.spotlight.length) html += h2('Squad Spotlight') + '<ul>' + d.spotlight.map(function (s) { return '<li><b>' + esc(s.who) + '</b> (' + titleSquad_(s.squad) + '): ' + esc(s.story) + '</li>'; }).join('') + '</ul>';
  const ex = d.expeditions.filter(function (x) { return x.open || x.winners.length; });
  if (ex.length) html += h2('Expeditions') + '<ul>' + ex.map(function (x) { return '<li><b>' + esc(x.title) + '</b>: ' + (x.winners.length ? 'won by ' + x.winners.map(titleSquad_).join(' + ') : SQUAD_KEYS_.map(function (s) { return titleSquad_(s) + ' ' + x.progress[s]; }).join(' · ') + ' of ' + x.goal) + '</li>'; }).join('') + '</ul>';
  const tnames = Object.keys(d.teachers);
  html += h2('Virtual Trait Dough') + '<p>' + d.dough + ' bill' + (d.dough === 1 ? '' : 's') + ' this week' + (tnames.length ? ' from ' + tnames.length + ' teacher' + (tnames.length === 1 ? '' : 's') + '. Thank you, ' + tnames.map(esc).join(', ') + '!' : '. Nobody gave Virtual Dough yet; the link is clashofclasses.org/dough.') + '</p>';
  const qn = SQUAD_KEYS_.reduce(function (n, s) { return n + d.quiet[s].length; }, 0);
  if (qn) html += h2('Quiet Climbers to catch next week') + '<p style="color:#666;margin-top:0">Nothing earned in ' + quietDays_() + ' days. Catch them doing something good.</p>' +
    SQUAD_KEYS_.filter(function (s) { return d.quiet[s].length; }).map(function (s) { return '<p style="margin:4px 0"><b>' + titleSquad_(s) + ':</b> ' + d.quiet[s].slice(0, 10).map(function (q) { return esc(q.name) + (q.grade ? ' (' + esc(q.grade) + ')' : ''); }).join(', ') + (d.quiet[s].length > 10 ? ' and ' + (d.quiet[s].length - 10) + ' more' : '') + '</p>'; }).join('');
  if (d.coming.length) html += h2('Coming up next week') + '<ul>' + d.coming.map(function (c) { return '<li>' + esc(c) + '</li>'; }).join('') + '</ul>';
  return html + '<p style="color:#888;font-size:12px;margin-top:20px">Clash of Classes · clashofclasses.org</p></div>';
}

/** Admin › System: the quiet list and this week's wrap-up (not sent). */
function clashWeekly(token) {
  requireAdmin_(token);
  const d = wrapData_(), props = PropertiesService.getScriptProperties();
  return JSON.stringify({ quiet: d.quiet, quietDays: quietDays_(), html: wrapHtml_(d), to: wrapRecipients_(), sent: props.getProperty('wrap-week') === d.mon, isFriday: new Date(d.today + 'T12:00:00').getDay() === 5 });
}

function clashWrapSend(token) {
  requireAdmin_(token);
  return sendWrap_();
}

function sendWrap_() {
  const d = wrapData_(), to = wrapRecipients_();
  if (!to.length) throw new Error('Add who should get it first: Settings › WRAP_EMAILS (or STATE_EMAILS / DIGEST_EMAILS).');
  const order = SQUAD_KEYS_.slice().sort(function (a, b) { return d.week[b] - d.week[a]; });
  MailApp.sendEmail({ to: to.join(','), subject: 'Clash Friday wrap-up · ' + titleSquad_(order[0]) + ' won the week', htmlBody: wrapHtml_(d), name: 'Clash of Classes' });
  PropertiesService.getScriptProperties().setProperty('wrap-week', d.mon);
  return 'Sent to ' + to.length + '.';
}

function menuPreviewWrap_() {
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(wrapHtml_(wrapData_())).setWidth(720).setHeight(640), 'Friday wrap-up (preview)');
}
function menuSendWrap_() {
  const ui = SpreadsheetApp.getUi(), to = wrapRecipients_();
  if (!to.length) { ui.alert('Add who should get it first: Settings › WRAP_EMAILS (or STATE_EMAILS / DIGEST_EMAILS).'); return; }
  if (ui.alert('Send this week\'s Friday wrap-up now?', 'To: ' + to.join(', '), ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  ui.alert(sendWrap_());
}

// ── The announcement page ──
function clashAnnouncePage() {
  return JSON.stringify({ date: todayYmd_(), text: announcementScript_() });
}
function serveAnnounce_() {
  return HtmlService.createHtmlOutputFromFile('Announce').setTitle('Morning announcement · Clash of Classes')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// file-stamp (tools/stamp-files.js): lets Admin › System tell whether this file is the newest one
var FILE_STAMP_51_Staff = '988bd1265a';
