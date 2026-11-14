/**
 * Clash of Classes -- Apps Script backend
 * File 13: School Calendar. The home page's "Days of School", "Days Until" and "Days Left"
 * count only real school days. Which days have no school lives on a tab staff can edit.
 *
 * School Calendar tab (one row per day with no students):
 *   Date | Note | From
 *   Comes filled with this year's EASD days off. Add a row for a snow day or any change;
 *   delete a row if school is back in session. Half days are school days: don't list them.
 *
 * The district calendar (optional): Settings › SCHOOL_CALENDAR = the district's Google Calendar
 * ID, or its iCal (.ics) link. Clash of Classes › Sync days off from the district calendar
 * adds every all-day event whose title sounds like a day off (SCHOOL_CALENDAR_WORDS), once.
 * Clash of Classes › Turn on daily calendar sync does that every morning at 5.
 * Rows it adds say "district calendar" in From; rows you type are never changed.
 */

const SCHOOL_CAL_TAB_ = 'School Calendar';
const SCHOOL_CAL_HEADERS_ = ['Date', 'Note', 'From'];

// 2026-2027 EASD non-pupil days (the list the site used to have built in).
const SCHOOL_CAL_SEED_ = [
  ['2026-09-04', 'No school (pre-Labor Day)'], ['2026-09-07', 'Labor Day'],
  ['2026-10-12', 'Columbus Day / Indigenous Peoples Day'], ['2026-10-23', 'Parent-Teacher Conferences (no students)'],
  ['2026-11-23', 'Thanksgiving break'], ['2026-11-24', 'Thanksgiving break'], ['2026-11-25', 'Thanksgiving break'],
  ['2026-11-26', 'Thanksgiving'], ['2026-11-27', 'Thanksgiving break'], ['2026-11-30', 'Thanksgiving break'],
  ['2026-12-23', 'Winter break'], ['2026-12-24', 'Winter break'], ['2026-12-25', 'Winter break'], ['2026-12-28', 'Winter break'],
  ['2026-12-29', 'Winter break'], ['2026-12-30', 'Winter break'], ['2026-12-31', 'Winter break'], ['2027-01-01', 'New Year\'s Day'],
  ['2027-01-18', 'MLK Day'], ['2027-01-19', 'Records Day (no students)'], ['2027-02-15', 'Presidents Day'],
  ['2027-03-22', 'Parent-Teacher Conferences (no students)'], ['2027-03-23', 'Parent-Teacher Conferences (no students)'],
  ['2027-03-24', 'Spring break'], ['2027-03-25', 'Spring break'], ['2027-03-26', 'Spring break'],
  ['2027-03-29', 'Weather make-up / vacation day'], ['2027-05-28', 'Weather make-up / vacation day'], ['2027-05-31', 'Memorial Day']];

function ensureSchoolCalendarTab_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getSheetByName(SCHOOL_CAL_TAB_)) return;
  const sh = ss.insertSheet(SCHOOL_CAL_TAB_);
  sh.getRange(1, 1, 1, SCHOOL_CAL_HEADERS_.length).setValues([SCHOOL_CAL_HEADERS_]).setFontWeight('bold').setBackground('#490e6f').setFontColor('#ffffff');
  sh.setFrozenRows(1);
  sh.getRange(2, 1, SCHOOL_CAL_SEED_.length, 3).setValues(SCHOOL_CAL_SEED_.map(function (r) { return [r[0], r[1], 'built in']; }));
  sh.getRange('A2:A400').setNumberFormat('yyyy-mm-dd');
  sh.setColumnWidth(2, 320);
}

// What the page needs: { noSchool: ['yyyy-MM-dd', …], start, end } (start/end '' = the site's own).
function schoolCalendar_() {
  const days = {};
  sheetRows_(SCHOOL_CAL_TAB_).forEach(function (r) {
    const d = String(r['Date'] || '').trim();
    if (/^\d{4}-\d{2}-\d{2}$/.test(d)) days[d] = true;
  });
  const ymd = function (k) { const v = String(getSetting_(k) || '').trim(); return /^\d{4}-\d{2}-\d{2}$/.test(v) ? v : ''; };
  return { noSchool: Object.keys(days).sort(), start: ymd('SCHOOL_START'), end: ymd('LAST_DAY'), nextStart: ymd('NEXT_YEAR_START') };
}

// ── District calendar sync ──────────────────────────────────────────
// [{ date: 'yyyy-MM-dd', title }] for every day covered by an all-day event in the window.
function districtDaysOff_(fromYmd, toYmd) {
  const src = String(getSetting_('SCHOOL_CALENDAR') || '').trim();
  if (!src) throw new Error('Put the district calendar on Settings › SCHOOL_CALENDAR first (its Google Calendar ID or its .ics link).');
  const words = String(getSetting_('SCHOOL_CALENDAR_WORDS') || '').split(',').map(function (w) { return w.trim().toLowerCase(); }).filter(String);
  const skip = /half[ -]?day|early dismissal|early release|delay/i;
  const isOff = function (t) { t = String(t || '').toLowerCase(); return !skip.test(t) && words.some(function (w) { return t.indexOf(w) > -1; }); };
  const tz = Session.getScriptTimeZone(), out = [];
  const addRange = function (startDate, endExclusive, title) {
    for (let d = new Date(startDate); d < endExclusive; d.setDate(d.getDate() + 1)) {
      const k = Utilities.formatDate(d, tz, 'yyyy-MM-dd');
      if (k >= fromYmd && k <= toYmd) out.push({ date: k, title: title });
    }
  };
  if (/^https?:|^webcal:/i.test(src)) {
    calendarIcsEvents_(src.replace(/^webcal:/i, 'https:')).forEach(function (e) { if (e.allDay && isOff(e.title)) addRange(e.start, e.end, e.title); });
  } else {
    const id = (src.match(/[?&]src=([^&]+)/) || [])[1] ? decodeURIComponent(src.match(/[?&]src=([^&]+)/)[1]) : src;
    const cal = CalendarApp.getCalendarById(id);
    if (!cal) throw new Error('Couldn\'t open that calendar. In Google Calendar, add the district calendar to your own list first, then copy its Calendar ID from Settings and sharing › Integrate calendar.');
    cal.getEvents(new Date(fromYmd + 'T00:00:00'), new Date(toYmd + 'T23:59:59')).forEach(function (e) {
      if (e.isAllDayEvent() && isOff(e.getTitle())) addRange(e.getAllDayStartDate(), e.getAllDayEndDate(), e.getTitle());
    });
  }
  return out;
}

// All-day and timed events from an iCal feed: [{ title, start: Date, end: Date (exclusive), allDay }].
function calendarIcsEvents_(url) {
  const text = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true }).getContentText().replace(/\r?\n[ \t]/g, '');
  const day = function (v) { return new Date(+v.slice(0, 4), +v.slice(4, 6) - 1, +v.slice(6, 8)); };
  return text.split('BEGIN:VEVENT').slice(1).map(function (block) {
    // A property line looks like NAME;PARAMS:VALUE or NAME:VALUE.
    const lines = block.split(/[\r\n]+/);
    const get = function (k) {
      const line = lines.filter(function (l) { return l.indexOf(k + ':') === 0 || l.indexOf(k + ';') === 0; })[0];
      if (!line) return null;
      const colon = line.indexOf(':');
      return { params: line.slice(k.length, colon), value: line.slice(colon + 1).trim() };
    };
    const s = get('DTSTART'), e = get('DTEND'), t = get('SUMMARY');
    if (!s) return null;
    const allDay = /VALUE=DATE/.test(s.params) || /^\d{8}$/.test(s.value);
    const start = day(s.value), end = e ? day(e.value) : new Date(start.getTime() + 864e5);
    if (!allDay) return { title: t ? t.value : '', start: start, end: start, allDay: false };
    return { title: (t ? t.value : '').replace(/\\([,;\\])/g, '$1'), start: start, end: end > start ? end : new Date(start.getTime() + 864e5), allDay: true };
  }).filter(Boolean);
}

// Adds district days off that aren't on the tab yet. Returns { added: [...], skipped }.
function syncSchoolCalendar_() {
  ensureSchoolCalendarTab_();
  const cal = schoolCalendar_(), sh = getSheet_(SCHOOL_CAL_TAB_);
  const tz = Session.getScriptTimeZone(), today = todayYmd_();
  const from = cal.start || Utilities.formatDate(new Date(new Date().getFullYear() - (new Date().getMonth() < 6 ? 1 : 0), 7, 1), tz, 'yyyy-MM-dd');
  const to = cal.end || Utilities.formatDate(new Date(+from.slice(0, 4) + 1, 6, 1), tz, 'yyyy-MM-dd');
  const have = {}; cal.noSchool.forEach(function (d) { have[d] = true; });
  const added = [];
  districtDaysOff_(from, to).forEach(function (o) {
    const dow = new Date(o.date + 'T12:00:00').getDay();
    if (dow === 0 || dow === 6 || have[o.date]) return;
    have[o.date] = true;
    added.push([o.date, o.title, 'district calendar ' + today]);
  });
  if (added.length) sh.getRange(sh.getLastRow() + 1, 1, added.length, 3).setValues(added);
  return added;
}

function menuSyncSchoolCalendar_() {
  const ui = SpreadsheetApp.getUi();
  try {
    const added = syncSchoolCalendar_();
    ui.alert(added.length ? 'Added ' + added.length + ' day' + (added.length === 1 ? '' : 's') + ' off to the School Calendar tab:\n\n' +
      added.map(function (a) { return a[0] + '  ' + a[1]; }).join('\n') : 'The School Calendar tab already has every day off from the district calendar.');
  } catch (e) { ui.alert(e.message); }
}

function scheduledSchoolCalendarSync() { try { syncSchoolCalendar_(); } catch (e) { console.warn('School calendar sync: ' + e.message); } }

function menuTurnOnCalendarSync_() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'scheduledSchoolCalendarSync') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('scheduledSchoolCalendarSync').timeBased().everyDays(1).atHour(5).create();
  SpreadsheetApp.getUi().alert('Every morning at 5, new days off on the district calendar (Settings › SCHOOL_CALENDAR) are added to the School Calendar tab.');
}
