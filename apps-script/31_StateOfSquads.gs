/**
 * Clash of Classes -- Apps Script backend
 * File 31: the State of the Squads, a monthly email to admins about last calendar month.
 *
 * What it holds (counts only, never a student's name or 800#):
 *   - points per squad last month and the standings
 *   - the map: who holds each area, on what rule and number, and what changed hands
 *   - the Squad Monthly numbers per squad (Students, attendance, LRR + LD, failing, interventions)
 *   - the frost check (Safe / Frost warning / Freeze proposed / Still frozen) and any freeze to approve
 *   - every Pulse that closed last month and what students said most, plus the You Said We Did
 *     rows still waiting for an answer
 *   - the points balance flags (29_Balance.gs) and the privacy check (32_Privacy.gs)
 * Sent once a month on STATE_DAY (2 = the 2nd, after Climb-er's numbers and the frost check are in)
 * to STATE_EMAILS (blank = DIGEST_EMAILS) when STATE_OF_SQUADS is ON. Menu: preview or send now.
 */

function stateRecipients_() {
  const raw = String(getSetting_('STATE_EMAILS') || '').trim() || String(getSetting_('DIGEST_EMAILS') || '');
  return raw.split(/[,;\s]+/).map(function (s) { return s.trim().toLowerCase(); }).filter(function (s) { return /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(s); });
}

/** Everything the email shows, for one month (default: last month). */
function stateOfSquadsData_(month) {
  const today = todayYmd_();
  month = month || lastMonthWindow_(today).month;
  const win = monthWindow_(month);
  const events = readEventsTab_();
  const points = pointsBy_(events, win.from, win.to, null).points;
  const place = SQUAD_KEYS_.slice().sort(function (a, b) { return points[b] - points[a]; });
  let zones = [];
  try {
    zones = mapZones_(events, false).map(function (z) {
      return { zone: z.zone, holder: z.holder, rule: z.rule || 'Most points', mult: z.mult, low: z.low, per100: z.per100, locked: z.locked,
        score: z.holder && z.score ? z.score[z.holder] : null };
    });
  } catch (e) { zones = []; }
  const battles = sheetRows_('Map Battles').filter(function (r) { return ymOf_(r['Date']) === month || String(r['Date'] || '').slice(0, 7) === month; })
    .map(function (r) { return { zone: String(r['Zone']), to: squadOf_(r['Taken By']), from: squadOf_(r['From']) }; });
  const monthly = squadMonthly_(month), sizes = squadSizes_(month, monthly);
  const frost = sheetRows_('Frost Log').filter(function (r) { return ymOf_(r['Month']) === month || String(r['Month'] || '').slice(0, 7) === month; })
    .map(function (r) { return { squad: squadOf_(r['Squad']), status: String(r['Status'] || ''), markers: String(r['Markers Crossed'] || ''), approved: ticked_(r['Approve']) }; });
  const pulses = pollsWithVotes_().filter(function (p) { return p.kind === 'Pulse' && p.question && String(p.through).slice(0, 7) === month && p.total; })
    .map(function (p) { return { question: p.question, said: youSaidLine_(p) }; });
  let waiting = [], balance = { flags: [] }, privacy = [];
  try { waiting = youSaidWaiting_(); } catch (e) { waiting = []; }
  try { balance = pointsBalance_(events, win.to, Number(win.to.slice(8))); } catch (e) { balance = { flags: [] }; }
  try { privacy = privacyScan_().filter(function (x) { return x.findings.length; }); } catch (e) { privacy = []; }
  return { month: month, points: points, place: place, zones: zones, battles: battles, monthly: monthly, sizes: sizes, frost: frost,
    pulses: pulses, waiting: waiting, flags: balance.flags || [], privacy: privacy, site: String(getSetting_('SITE_URL') || 'https:\/\/clashofclasses.org') };
}

function monthLabel_(ym) {
  const names = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
  return names[Number(ym.slice(5, 7)) - 1] + ' ' + ym.slice(0, 4);
}

// Email-safe HTML (tables and inline styles), the same look as the weekly digest.
function stateOfSquadsHtml_(d) {
  const C = { BLACK: '#3f4652', GOLD: '#b8860b', GREY: '#5f6f86', PURPLE: '#6d28d9' };
  const T = function (s) { return s ? titleSquad_(s) : 'open'; };
  const h = function (s) { return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); };
  const sq = function (s) { return s ? '<span style="font-weight:700;color:' + C[s] + '">' + T(s) + '</span>' : '<span style="color:#999">open</span>'; };
  const sec = function (title, body) { return '<tr><td style="padding:18px 24px 4px"><div style="font:700 17px Georgia,serif;color:#2a0f45;border-bottom:2px solid #f2e3ff;padding-bottom:6px">' + title + '</div>' + body + '</td></tr>'; };
  const p = function (t) { return '<p style="margin:8px 0;font:14px/1.5 Arial,sans-serif;color:#2b2b3a">' + t + '</p>'; };
  const ul = function (items) { return items.length ? '<ul style="margin:8px 0;padding-left:18px;font:14px/1.6 Arial,sans-serif;color:#2b2b3a">' + items.map(function (x) { return '<li>' + x + '</li>'; }).join('') + '</ul>' : ''; };
  const table = function (head, rows) {
    return '<table width="100%" cellpadding="5" style="border-collapse:collapse;font:13px Arial,sans-serif;margin-top:8px"><tr style="color:#6b5a85;font-size:12px">' +
      head.map(function (x, i) { return '<td' + (i ? ' align="right"' : '') + '>' + x + '</td>'; }).join('') + '</tr>' +
      rows.map(function (r) { return '<tr style="border-top:1px solid #eee">' + r.map(function (x, i) { return '<td' + (i ? ' align="right"' : '') + '>' + x + '</td>'; }).join('') + '</tr>'; }).join('') + '</table>';
  };
  const standings = table(['Place', 'Squad', 'Points last month'], d.place.map(function (s, i) { return [String(i + 1), sq(s), d.points[s].toLocaleString()]; }));
  const map = table(['Area', 'Rule', 'Holder', 'x Pay'], d.zones.map(function (z) {
    return [h(z.zone) + (z.locked ? ' 🔒' : ''), h(z.rule), sq(z.holder) + (z.score != null ? ' <span style="color:#888">(' + h(z.score) + (z.per100 ? ' /100' : '') + ')</span>' : ''), h(z.mult)];
  })) + p(d.battles.length ? 'Changed hands: ' + d.battles.map(function (b) { return sq(b.to) + ' took the ' + h(b.zone) + (b.from ? ' from ' + sq(b.from) : ''); }).join('; ') + '.' : 'No area changed hands.');
  const cols = [['Students', 'Students'], ['attendance', 'Attendance %'], ['lrr_ld', 'LRR + LD'], ['failing', 'Failing'], ['interventions_teacher', 'Teacher int.'], ['interventions_admin', 'Admin int.'], ['pride_day', 'Pride Day %'], ['shout_outs', 'Shout-outs']];
  const numbers = table(['Squad'].concat(cols.map(function (c) { return c[1]; })), SQUAD_KEYS_.map(function (s) {
    const m = d.monthly[s] || {};
    return [sq(s)].concat(cols.map(function (c) { const v = c[0] === 'Students' ? (m.students || d.sizes[s]) : m[c[0]]; return v === undefined || v === null || v === '' ? '<span style="color:#bbb">·</span>' : h(v); }));
  })) + p('<span style="color:#888;font-size:12px">· = no number yet. Climb-er fills these on its hourly sync; LRR + LD comes from the Follow-Up sheets. A highest/lowest area waits until every squad has a number.</span>');
  const frost = d.frost.length ? ul(d.frost.map(function (f) {
    return sq(f.squad) + ': ' + h(f.status) + (f.markers ? ' (' + h(f.markers) + ')' : '') + (f.status === 'Freeze proposed' && !f.approved ? ' <b style="color:#b45309">needs your Approve on the Frost Log tab</b>' : '');
  })) : p('No frost check for this month (it starts at FROST_START).');
  const pulse = (d.pulses.length ? ul(d.pulses.map(function (x) { return '<b>' + h(x.question) + '</b><br><span style="color:#555">' + h(x.said) + '</span>'; })) : p('No Pulse question closed this month.')) +
    (d.waiting.length ? p('<b>' + d.waiting.length + '</b> waiting for "What we did" on the You Said We Did tab: ' + d.waiting.slice(0, 5).map(function (w) { return h(w.question); }).join('; ') + (d.waiting.length > 5 ? '…' : '') + '. Students see it only once you answer.') : '');
  const balance = d.flags.length ? ul(d.flags.map(h)) : p('Nothing lopsided: no one source of points decided the month.');
  const privacy = d.privacy.length ? ul(d.privacy.map(function (x) { return '<b style="color:#b91c1c">' + h(x.surface) + '</b>: ' + h(x.findings.join('; ')); })) : p('Every public page checked: no 800# and no full student name.');
  return '<!doctype html><html><body style="margin:0;background:#f4f0fa"><table width="100%" cellpadding="0" style="background:#f4f0fa;padding:20px 0"><tr><td align="center">' +
    '<table width="640" cellpadding="0" style="max-width:640px;width:100%;background:#ffffff;border-radius:14px;overflow:hidden;border:1px solid #e6dcf5">' +
    '<tr><td style="background:#2a0f45;padding:22px 24px"><div style="font:700 12px Arial,sans-serif;letter-spacing:2px;color:#ffe100;text-transform:uppercase">State of the Squads</div>' +
    '<div style="font:700 24px Georgia,serif;color:#ffffff;margin-top:4px">' + monthLabel_(d.month) + '</div>' +
    '<div style="font:13px Arial,sans-serif;color:#d6c2ff;margin-top:4px">' + T(d.place[0]) + ' earned the most points: ' + d.points[d.place[0]].toLocaleString() + '.</div></td></tr>' +
    sec('Standings for the month', standings) + sec('The map', map) + sec('Squad Monthly numbers', numbers) + sec('Freeze &amp; Thaw', frost) +
    sec('What students told us (Pulse)', pulse) + sec('Points balance', balance) + sec('Privacy check', privacy) +
    '<tr><td style="padding:18px 24px 24px"><a href="' + h(d.site) + '" style="display:inline-block;background:#ffe100;color:#2a0f45;font:700 14px Arial,sans-serif;padding:10px 18px;border-radius:999px;text-decoration:none">Open Clash of Classes</a>' +
    '<p style="font:12px Arial,sans-serif;color:#999;margin:14px 0 0">Sent automatically on the ' + settingNum_('STATE_DAY', 2) + ' of each month by Clash of Classes. Recipients: Settings › STATE_EMAILS (blank = DIGEST_EMAILS). Counts only.</p></td></tr></table></td></tr></table></body></html>';
}

/** Sends last month's email (once a month unless force). Returns a line. */
function stateOfSquadsEmail(force) {
  noteRun_('stateOfSquads');
  const props = PropertiesService.getScriptProperties(), d = stateOfSquadsData_(), to = stateRecipients_();
  if (!force && props.getProperty('state-month') === d.month) return 'Already sent for ' + d.month + '.';
  if (!to.length) return 'No recipients: add emails to Settings › STATE_EMAILS (or DIGEST_EMAILS).';
  const subject = 'State of the Squads · ' + monthLabel_(d.month) + ' · ' + titleSquad_(d.place[0]) + ' led the month';
  try {
    MailApp.sendEmail({ to: to.join(','), subject: subject, htmlBody: stateOfSquadsHtml_(d), name: 'Clash of Classes' });
  } catch (e) { noteError_('stateOfSquads', e); throw e; }
  props.setProperty('state-month', d.month);
  return 'Sent to ' + to.length + '.';
}

// From the background tick: once a month, on or after STATE_DAY.
function maybeStateOfSquads_() {
  if (String(getSetting_('STATE_OF_SQUADS') || 'ON').toUpperCase() !== 'ON') return;
  if (Number(todayYmd_().slice(8)) < Math.max(1, settingNum_('STATE_DAY', 2))) return;
  if (!stateRecipients_().length) return;
  stateOfSquadsEmail(false);
}

function menuPreviewStateOfSquads_() {
  const html = HtmlService.createHtmlOutput(stateOfSquadsHtml_(stateOfSquadsData_())).setWidth(720).setHeight(640);
  SpreadsheetApp.getUi().showModalDialog(html, 'State of the Squads (preview, last month)');
}

function menuSendStateOfSquads_() {
  const ui = SpreadsheetApp.getUi(), to = stateRecipients_();
  if (!to.length) { ui.alert('Add who should get it first: Settings › STATE_EMAILS (or DIGEST_EMAILS).'); return; }
  if (ui.alert('Send the State of the Squads for last month now?', 'To: ' + to.join(', '), ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  ui.alert(stateOfSquadsEmail(true));
}
