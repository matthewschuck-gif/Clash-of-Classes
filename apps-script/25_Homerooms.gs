/**
 * Clash of Classes -- Apps Script backend
 * File 25: Homeroom points (the site's Homerooms page). Separate from squad points.
 *
 * Homeroom spirit days (Mounts Pride Day, Patriots Day, Red Ribbon Hat Day …) are Google Forms:
 * each homeroom sends its room number, how many took part and how many were there. A homeroom
 * earns round(10 × took part ÷ there), 0 to 10 -- the same rule the response sheets used.
 *
 * HOMEROOM FORMS tab: Show | Event | Responses Sheet. One row per form; paste the link to its
 *   responses Sheet. Columns are found by their questions (room number / how many … / total).
 *   A form used on several days (Mounts Pride Day) becomes one event per day.
 * HOMEROOM POINTS tab: Approved | Date | Event | Room | Took Part | Total | Points | Notes | Key.
 *   New responses land here every half hour (or Menu › Import homeroom form responses) with
 *   Approved unticked; staff tick Approved to count them (Settings › HOMEROOM_AUTO_APPROVE ON =
 *   count straight away). Fix a typo'd count right in the row. Notes flags odd entries (more
 *   took part than were there, a second entry from the same room that day).
 * No squad points come from these; squads keep their own events.
 */

const HR_FORM_HEADERS_ = ['Show', 'Event', 'Responses Sheet'];
const HR_POINT_HEADERS_ = ['Approved', 'Date', 'Event', 'Room', 'Took Part', 'Total', 'Points', 'Notes', 'Key'];
const HR_FORM_SEED_ = [
  [true, 'Patriots Day Red, White & Blue', 'https://docs.google.com/spreadsheets/d/1c95v-5HpvOnrVQNUYqh1sq1FQo4dltQbf7Arl-3eazM/edit'],
  [true, 'Mounts Pride Day', 'https://docs.google.com/spreadsheets/d/1zyoqCl_9mDG8GKQs8ygt1qC0dczJF3hcFS9otDUUO4M/edit'],
  [true, 'Red Ribbon Week Hat Day', 'https://docs.google.com/spreadsheets/d/11F5BAqfVRS4IiUIF-deZ79KEbNUaJ9n7QXn0eONwlDM/edit'],
];

function ensureHomeroomTabs_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  [['Homeroom Forms', HR_FORM_HEADERS_, HR_FORM_SEED_], ['Homeroom Points', HR_POINT_HEADERS_, null]].forEach(function (t) {
    if (ss.getSheetByName(t[0])) return;
    const sh = ss.insertSheet(t[0]);
    sh.getRange(1, 1, 1, t[1].length).setValues([t[1]]).setFontWeight('bold').setBackground('#490e6f').setFontColor('#ffffff');
    if (t[2]) { sh.getRange(2, 1, t[2].length, t[1].length).setValues(t[2]); sh.getRange(2, 1, t[2].length, 1).insertCheckboxes(); }
    sh.setFrozenRows(1);
  });
}

function hrPoints_(took, total) { return total > 0 ? Math.max(0, Math.min(10, Math.round(10 * took / total))) : 0; }

// The form's answer rows as {date, room, took, total}; columns found by their questions.
function hrResponses_(values) {
  if (!values || values.length < 2) return [];
  const h = values[0].map(function (x) { return String(x).toLowerCase(); });
  const find = function (test) { for (let j = 0; j < h.length; j++) if (test(h[j])) return j; return -1; };
  const cTime = find(function (x) { return x.indexOf('timestamp') > -1; });
  const cRoom = find(function (x) { return x.indexOf('room') > -1; });
  const cTotal = find(function (x) { return x.indexOf('total') > -1; });
  const cTook = find(function (x) { return x.indexOf('how many') > -1 && x.indexOf('total') === -1; });
  if (cTime < 0 || cRoom < 0 || cTotal < 0 || cTook < 0) return [];
  return values.slice(1).map(function (r) {
    const t = r[cTime], d = t instanceof Date ? t : new Date(String(t || ''));
    return { date: isNaN(d) ? '' : Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd'), stamp: isNaN(d) ? '' : d.toISOString(),
      room: String(r[cRoom] || '').replace(/\D/g, ''), took: Number(String(r[cTook]).replace(/[^\d.]/g, '')) || 0, total: Number(String(r[cTotal]).replace(/[^\d.]/g, '')) || 0 };
  }).filter(function (x) { return x.date && x.room; });
}

// Adds responses not yet on Homeroom Points. read(sheetId) returns a responses sheet's values.
function importHomeroomResponses_(read) {
  ensureHomeroomTabs_();
  read = read || function (id) { return SpreadsheetApp.openById(id).getSheets()[0].getDataRange().getValues(); };
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Homeroom Points');
  const have = {}, seenDay = {};
  sheetRows_('Homeroom Points').forEach(function (r) { have[String(r['Key'])] = true; seenDay[r['Event'] + '|' + r['Date'] + '|' + r['Room']] = true; });
  const auto = String(getSetting_('HOMEROOM_AUTO_APPROVE') || 'OFF').toUpperCase() === 'ON';
  const added = [], problems = [];
  sheetRows_('Homeroom Forms').forEach(function (f) {
    if (!(String(f['Show']).trim() === '' || ticked_(f['Show']))) return;
    const id = extractSheetId_(String(f['Responses Sheet'] || ''));
    if (!id) return;
    let values;
    try { values = read(id); } catch (e) { problems.push(f['Event'] + ': ' + e.message); return; }
    hrResponses_(values).forEach(function (x) {
      const key = id.slice(0, 10) + '|' + x.stamp + '|' + x.room;
      if (have[key]) return;
      have[key] = true;
      const day = f['Event'] + '|' + x.date + '|' + x.room, notes = [];
      if (x.took > x.total) notes.push('More took part than were there: check the counts');
      if (seenDay[day]) notes.push('Second entry from this room for this day');
      seenDay[day] = true;
      const row = [auto && !notes.length, x.date, String(f['Event']), x.room, x.took, x.total, hrPoints_(x.took, x.total), notes.join('; '), key];
      sh.appendRow(row);
      added.push(row);
    });
  });
  return { added: added.length, problems: problems };
}

function maybeImportHomerooms_() {
  const props = PropertiesService.getScriptProperties(), last = Number(props.getProperty('hr-import')) || 0;
  if (Date.now() - last < 30 * 60 * 1000) return;
  props.setProperty('hr-import', String(Date.now()));
  importHomeroomResponses_();
}

function menuImportHomerooms_() {
  const r = importHomeroomResponses_();
  SpreadsheetApp.getUi().alert(r.added + ' new homeroom response' + (r.added === 1 ? '' : 's') + ' added to Homeroom Points' +
    (String(getSetting_('HOMEROOM_AUTO_APPROVE') || 'OFF').toUpperCase() === 'ON' ? '.' : ': tick Approved to count them.') +
    (r.problems.length ? '\n\nCould not read: ' + r.problems.join('; ') : ''));
}

// The Homerooms page: approved points by room, this month, and the latest event.
function homeroomBoard_() {
  const rows = sheetRows_('Homeroom Points').filter(function (r) { return ticked_(r['Approved']) && String(r['Room']).trim(); });
  const m = monthBounds_(0), rooms = {}, events = {};
  rows.forEach(function (r) {
    const room = String(r['Room']).replace(/\D/g, ''), pts = Number(r['Points']) || 0, date = String(r['Date'] || '').slice(0, 10);
    const o = rooms[room] = rooms[room] || { room: room, total: 0, month: 0, events: 0 };
    o.total += pts; o.events++;
    if (date >= m[0]) o.month += pts;
    const ek = r['Event'] + '|' + date;
    (events[ek] = events[ek] || { event: String(r['Event']), date: date, rooms: [] }).rooms.push({ room: room, points: pts, took: Number(r['Took Part']) || 0, total: Number(r['Total']) || 0 });
  });
  const list = Object.keys(rooms).map(function (k) { return rooms[k]; }).sort(function (a, b) { return b.total - a.total || a.room.localeCompare(b.room); });
  const latest = Object.keys(events).map(function (k) { return events[k]; }).sort(function (a, b) { return b.date.localeCompare(a.date); })[0] || null;
  if (latest) latest.rooms.sort(function (a, b) { return b.points - a.points; });
  return { rooms: list, latest: latest, eventsCount: Object.keys(events).length };
}

function clashHomerooms() {
  const key = 'homerooms-' + dataStamp_();
  const hit = cacheGetBig_(key);
  if (hit) return hit;
  const text = JSON.stringify(homeroomBoard_());
  cachePutBig_(key, text, 300);
  return text;
}
