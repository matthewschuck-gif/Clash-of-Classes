/**
 * Clash of Classes -- Apps Script backend
 * File 24: Year Plan (the site's Year Plan page and This Week's "up next").
 *
 * YEAR PLAN tab: every squad event, Mindset lesson and material for the year, one row each --
 *   Show | Month | Date | Kind | Title | For | Link | Link 2 | Notes
 * Show: untick to hide a row (blank = shown). Date: blank = sometime that month. Kind: Mindset
 * Lesson, Kickoff, Homeroom Activity, Color Squad Activity, Mounts Pride Day, Student of the Month,
 * Assembly, Resource (any word works; these get icons). For: 7th, 8th, All or Staff. Links: any
 * Docs / Slides / Forms / Drive / Canva link. Add rows as the year goes; the page updates itself.
 * Seeded from the shared drive's 26-27 Monthly Themes docs. January to May still carried last
 * year's dates there, so those rows have no date yet (Notes says last year's) -- fill them in.
 */

const YEAR_PLAN_HEADERS_ = ['Show', 'Month', 'Date', 'Kind', 'Title', 'For', 'Link', 'Link 2', 'Notes'];
const SCHOOL_MONTHS_ = ['August', 'September', 'October', 'November', 'December', 'January', 'February', 'March', 'April', 'May', 'June'];

function yearPlanSeed_() {
  const d = function (id) { return 'https://docs.google.com/document/d/' + id + '/edit?usp=sharing'; };
  const s = function (id) { return 'https://docs.google.com/presentation/d/' + id + '/edit?usp=sharing'; };
  const v = function (id) { return 'https://drive.google.com/file/d/' + id + '/view'; };
  const f = function (id) { return 'https://docs.google.com/forms/d/e/' + id + '/viewform'; };
  const g7 = d('1Xijj6wg-cHSjspOWHDGibhlafllkq9zA8v15vj3QWc8'), board = d('1IWaNKNHbGdmnM-5bESYd4-vTRi_StiuQr3uZ2tqETeY');
  const squadBoard = d('1NVu1nDPbhGMK685BM5HanU3KEwjzeVjHwvU5duwD3RM'), info8 = d('1zTINwnyUIFt4Ur8Izzd9ZG4oYNhxApo7EFUN2zxHqXs');
  const pride = f('1FAIpQLScSlPXd7PP177T8l3V_hRjEjfvMmKGyKJBylKru1D0jIrOZQw'), sotm = f('1FAIpQLSe9UTY7Um1-AsrgMMmAt63yG2HEZNe1yeaJcmyk4G7ziPAeXg');
  const prideNote = 'Submit homeroom counts before 7th grade lunch (11:11); include yourself; cover absent teammates.';
  const sotmNote = 'Teams choose two students, prepare Mountaineer cards, submit names, deliver in homeroom.';
  const lastYear = function (when) { return 'Last year: ' + when + '. Set the 26-27 date.'; };
  const stations = 'Stations rotate by squad: Distance Challenge (Cafeteria), Trivia (Auditorium), Memory Minesweeper (Gym), Sync or Sink (Media Center).';
  const R = function (month, date, kind, title, who, link, link2, notes) { return [true, month, date || '', kind, title, who || 'All', link || '', link2 || '', notes || '']; };
  return [
    R('September', '', 'Resource', 'LRG student one-pager (print for LRG folders)', 'Staff', d('1O_h3UKXgOZOWo6Vq3vl6Luv_ohcQJ8U2Rw_eGX4aTn4')),
    R('September', '', 'Resource', 'Mountaineer Peaks resources (four Core Extension periods)', 'Staff', 'http://bit.ly/EMSpeaks'),
    R('September', '2026-09-11', 'Homeroom Activity', 'Patriots Day Red, White & Blue count', '7th', f('1FAIpQLSeL9F82VwzOHHJTf--WeAFvLgR_KoDlNSDd1vZRHg5e3iYOuQ'), '', 'Submit homeroom counts on the form.'),
    R('September', '2026-09-11', 'Homeroom Activity', 'Homeroom Logo Design (due 9/25)', '8th', s('1y16UoRK163Tj0JBeqmKj0jtZF2mKFsWJOKWrKWA7_LQ'), '', 'Sept 11 and 18. No points for this one.'),
    R('September', '2026-09-18', 'Mounts Pride Day', 'Mounts Pride Day', 'All', pride, '', prideNote),
    R('September', '2026-09-21', 'Mindset Lesson', 'Continual Learning & Growth Mindset: trait video + Mindset Lessons', '7th', g7, v('1wSdUeSgRmuCj2bSzV6lc_AVJiQn3WKAg')),
    R('September', '2026-09-21', 'Mindset Lesson', 'Trait slide + Choice Board (print September)', '8th', s('1EQv6RKuCzgew3Q9wCjSpwfTfpkhAb03ftDvokEdw948'), board, 'Schoology: copy the folder from the faculty course.'),
    R('September', '2026-09-25', 'Assembly', 'Kick-off Assembly (altered schedule)', 'All', '', '', 'Homeroom and Color Squad points.'),
    R('September', '2026-09-28', 'Mindset Lesson', 'Choice Board (print September)', '8th', board),
    R('October', '2026-10-02', 'Homeroom Activity', 'Halloween Costume Design', 'All', d('1T6eZ2eRpNKHUY8RV7oFu5Fp-svD7JKHpNnEHhFGK8lM'), 'https://docs.google.com/presentation/d/1mFEIGkJ6K-VRztU_X7ykUklugzkz0FO_/edit?usp=sharing', 'Link 2 = the drawing page.'),
    R('October', '2026-10-05', 'Mindset Lesson', 'Critical Thinking & Problem Solving: Mindset Lesson', '7th', g7),
    R('October', '2026-10-05', 'Mindset Lesson', 'Trait slide + Choice Board (print October)', '8th', s('1DQd-qVH0iJasxQ61Q8PRHsWjLZbxx-rLTZJGMOM116E'), board),
    R('October', '2026-10-08', 'Student of the Month', 'September Students of the Month celebrated', 'Staff', sotm, '', sotmNote),
    R('October', '2026-10-16', 'Color Squad Activity', 'Color Squad Choice Board', 'All', squadBoard),
    R('October', '2026-10-16', 'Mounts Pride Day', 'Mounts Pride Day', 'All', pride, '', prideNote),
    R('October', '2026-10-19', 'Mindset Lesson', 'Mindset Lesson', '7th', g7),
    R('October', '2026-10-19', 'Mindset Lesson', 'Choice Board (print October)', '8th', board),
    R('November', '2026-11-02', 'Mindset Lesson', 'Creativity & Innovation: Mindset Lesson', '7th', g7),
    R('November', '2026-11-02', 'Mindset Lesson', 'Mindset information (Nov 2 and 9)', '8th', info8),
    R('November', '2026-11-05', 'Student of the Month', 'October Students of the Month celebrated', 'Staff', sotm, '', sotmNote),
    R('November', '2026-11-13', 'Homeroom Activity', 'Trashketball', 'All', s('12HnMeP-eGz6khbYYZBHLjv4XSzlKEZs1CF6Gh8Uto94'), 'https://docs.google.com/forms/d/e/1FAIpQLSdndCaEZZRQNe7GTRqePvJ5xVpT-SKVMkvAM1UC_eSGRQNQUw/viewform', 'Link 2 = submit points.'),
    R('November', '2026-11-19', 'Mounts Pride Day', 'Mounts Pride Day', 'All', pride, '', prideNote),
    R('November', '2026-11-20', 'Color Squad Activity', 'Color Squad Choice Board (half day)', 'All', squadBoard),
    R('December', '2026-12-03', 'Student of the Month', 'November Students of the Month celebrated', 'Staff', sotm, '', sotmNote),
    R('December', '2026-12-04', 'Color Squad Activity', 'Color Squad Choice Board', 'All', squadBoard),
    R('December', '2026-12-07', 'Mindset Lesson', 'Honesty, Integrity & Responsibility: Mindset Lesson (Dec 7 and 14)', '7th', g7),
    R('December', '2026-12-07', 'Mindset Lesson', 'Mindset information (Dec 7 and 14)', '8th', info8),
    R('December', '2026-12-07', 'Homeroom Activity', 'Door Decorating (due Dec 18)', 'All', 'https://www.canva.com/design/DAGtX_akRME/OfW1r0_P2Wg3Ix9UaMr34A/edit', '', 'Pictures at the end of PASS Dec 18; students vote in Core Extension Dec 19.'),
    R('December', '2026-12-08', 'Kickoff', 'Monthly kickoff video', 'All', v('1iKC_THuolYHeRiFNOIYp2GkFZf052rMi')),
    R('December', '2026-12-18', 'Mounts Pride Day', 'Mounts Pride Day', 'All', pride, '', prideNote),
    R('January', '', 'Kickoff', 'Monthly kickoff video', 'All', v('1fOfyWHH9NCq9lBvUO0Fc7Q4stnkal7bH'), '', lastYear('Jan 5')),
    R('January', '', 'Mindset Lesson', 'Collaboration & Teamwork: Mindset Lesson', '7th', g7, '', lastYear('Jan 5')),
    R('January', '', 'Mindset Lesson', 'Mindset Lesson', '8th', d('17nDIfrp5li295cQLYppl4p7e4FC8bwX9AYmz-7ak_1w'), '', lastYear('Jan 5') + ' Check this is the 26-27 lesson.'),
    R('January', '', 'Student of the Month', 'December Students of the Month celebrated', 'Staff', sotm, '', lastYear('Jan 7')),
    R('January', '', 'Homeroom Activity', '24 Competition', 'All', d('1u1S5-Yd-B6beVFLcK3LoVAd6d_i4_7K1IrKb6A_848g'), '', lastYear('Jan 9')),
    R('January', '', 'Color Squad Activity', 'Color Squad stations', 'All', '', '', stations + ' ' + lastYear('Jan 23')),
    R('January', '', 'Mounts Pride Day', 'Mounts Pride Day', 'All', pride, '', lastYear('Jan 30')),
    R('February', '', 'Kickoff', 'Monthly kickoff video', 'All', v('1GM-axfrlqIUEJXZNhxt9LN3vvCEMrbnJ'), '', lastYear('Feb 2')),
    R('February', '', 'Mindset Lesson', 'Communication & Empathy: Mindset Lesson', '7th', g7, '', lastYear('Feb 2 and 9')),
    R('February', '', 'Mindset Lesson', 'Mindset Lesson + Extension Activity', '8th', d('13QBpHYjDAw2kpXAFIBvFqTv08fCFjcj2QWy3IbbezTo'), d('1yVoCQfyx4V4rPGeCZk-K0Zmz618sSrYJNYuLYQY2aAg'), lastYear('Feb 2 and 9') + ' Check this is the 26-27 lesson.'),
    R('February', '', 'Homeroom Activity', 'Super Bowl Challenge', 'All', d('1G3oT56hOVrMt6W6OqaeQSpuK7hcIRMJjJJZCFIZg7SI'), '', lastYear('Feb 6')),
    R('February', '', 'Color Squad Activity', 'Color Squad stations', 'All', '', '', stations + ' ' + lastYear('Feb 20')),
    R('February', '', 'Student of the Month', 'February Students of the Month celebrated', 'Staff', sotm, '', lastYear('Feb 27')),
    R('February', '', 'Mounts Pride Day', 'Mounts Pride Day', 'All', pride, '', lastYear('Feb 27')),
    R('March', '', 'Kickoff', 'Monthly kickoff video', 'All', v('1VJR3_s9KaAC6M9qgwl3A548wBKwU6AsM'), '', lastYear('Mar 2')),
    R('March', '', 'Mindset Lesson', 'Healthy Living, Wellness & Self-Awareness: Mindset Lesson', '7th', g7, '', lastYear('Mar 2 and 9')),
    R('March', '', 'Mindset Lesson', 'Mindset Plan + Extension Activity', '8th', d('1lTGwRVbUp1Q55OMXtrWv0h0RQaWZ8DQ3WxwD0uCWFPE'), d('1oqBWenKjNKPq8zn7FYdU5HU92nxuGlSVnnhwXWIMIr8'), lastYear('Mar 2 and 9') + ' Check this is the 26-27 lesson.'),
    R('March', '', 'Homeroom Activity', 'Mountaineer Mayhem: Clash Logo Design Challenge', 'All', s('1cRxb3_2bND2avG3VUMe9e-gj2yMD-auW05nTjdEjEe0'), '', lastYear('Mar 13')),
    R('March', '', 'Color Squad Activity', 'Color Squad stations (7th grade called down first)', 'All', '', '', stations + ' ' + lastYear('Mar 27')),
    R('March', '', 'Student of the Month', 'March Students of the Month celebrated', 'Staff', sotm, '', lastYear('Mar 26')),
    R('March', '', 'Mounts Pride Day', 'Mounts Pride Day', 'All', pride, '', lastYear('Mar 27')),
    R('April', '', 'Kickoff', 'Monthly kickoff video + Trait Tro video', 'All', v('1ItVvcmXtlwlsg5TUlkNbnPYim0NGQyKp'), 'https://www.youtube.com/watch?v=xnZggbD1Ses', lastYear('Apr 13')),
    R('April', '', 'Mindset Lesson', 'Resilience & Grit: Mindset Lesson', '7th', g7, '', lastYear('Apr 13')),
    R('April', '', 'Mindset Lesson', 'Mindset Lesson', '8th', d('1gGeA4krvDaouygSsGQ3YoEl1E59Og1Q7fGbOI6PPg_E'), '', lastYear('Apr 13') + ' Check this is the 26-27 lesson.'),
    R('April', '', 'Homeroom Activity', 'Family Feud Kahoot', 'All', s('1hnieU_nHVwwOohjDw1ty6toQcJsKpHB54voVx6_T2zI'), '', lastYear('Apr 10')),
    R('April', '', 'Mounts Pride Day', 'Mounts Pride Day', 'All', pride, '', lastYear('Apr 24')),
    R('May', '', 'Student of the Month', 'April Students of the Month celebrated', 'Staff', sotm, '', lastYear('May 1')),
    R('May', '', 'Kickoff', 'Monthly kickoff video', '7th', v('1BHc7r6s9G9tLtgBVtaQbcX9LPaUcS8I0'), '', lastYear('May 4')),
    R('May', '', 'Kickoff', 'Monthly kickoff video', '8th', v('1LAXm56UVJT05T_8hpYGlDLJ2ZO6zIbI9'), '', lastYear('May 4')),
    R('May', '', 'Mindset Lesson', 'Adaptability & Flexibility: Mindset Lesson', '7th', g7, '', lastYear('May 4')),
    R('May', '', 'Mindset Lesson', 'Mindset Lesson', '8th', d('10c9YeVSqTsSx6bmk9u8Au9KGAZMs1BuWDEr-m_lKKe8'), '', lastYear('May 4') + ' Check this is the 26-27 lesson.'),
    R('May', '', 'Assembly', 'Family Feud Assembly', 'All', '', '', lastYear('May 6')),
    R('May', '', 'Color Squad Activity', 'Color Squad activity', 'All', '', '', lastYear('May 15')),
    R('May', '', 'Mounts Pride Day', 'Mounts Pride Day', 'All', pride, '', lastYear('May 22')),
    R('May', '', 'Student of the Month', 'May Students of the Month celebrated', 'Staff', sotm, '', lastYear('May 29')),
  ];
}

function ensureYearPlanTab_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getSheetByName('Year Plan')) return;
  const sh = ss.insertSheet('Year Plan');
  sh.getRange(1, 1, 1, YEAR_PLAN_HEADERS_.length).setValues([YEAR_PLAN_HEADERS_]).setFontWeight('bold').setBackground('#490e6f').setFontColor('#ffffff');
  const seed = yearPlanSeed_();
  sh.getRange(2, 1, seed.length, YEAR_PLAN_HEADERS_.length).setValues(seed);
  sh.getRange(2, 1, seed.length, 1).insertCheckboxes();
  sh.setFrozenRows(1);
}

// The shown rows, cleaned up: [{month, date, kind, title, who, links:[…], notes}], in school order.
function yearPlanItems_() {
  const http = function (u) { u = String(u || '').trim(); return /^https?:\/\//i.test(u) ? u : ''; };
  return sheetRows_('Year Plan').filter(function (r) {
    return String(r['Title'] || '').trim() && (String(r['Show']).trim() === '' || ticked_(r['Show']));
  }).map(function (r) {
    const date = /^\d{4}-\d{2}-\d{2}/.test(String(r['Date'] || '')) ? String(r['Date']).slice(0, 10) : '';
    let month = String(r['Month'] || '').trim();
    if (!month && date) month = SCHOOL_MONTHS_.filter(function (m) { return new Date(date + 'T12:00:00').getMonth() === new Date(m + ' 1, 2000').getMonth(); })[0] || '';
    return { month: month, date: date, kind: String(r['Kind'] || '').trim(), title: String(r['Title']).trim(), who: String(r['For'] || 'All').trim() || 'All',
      links: [http(r['Link']), http(r['Link 2'])].filter(String), notes: String(r['Notes'] || '').trim() };
  }).sort(function (a, b) {
    const ma = SCHOOL_MONTHS_.indexOf(a.month), mb = SCHOOL_MONTHS_.indexOf(b.month);
    return (ma < 0 ? 99 : ma) - (mb < 0 ? 99 : mb) || (a.date || '9999').localeCompare(b.date || '9999');
  });
}

function clashYearPlan() {
  const key = 'yearplan-' + dataStamp_();
  const hit = cacheGetBig_(key);
  if (hit) return hit;
  const items = yearPlanItems_(), months = [];
  items.forEach(function (it) {
    let m = months.filter(function (x) { return x.month === it.month; })[0];
    if (!m) { m = { month: it.month, items: [] }; months.push(m); }
    m.items.push(it);
  });
  const today = todayYmd_();
  const text = JSON.stringify({ months: months, today: today, current: SCHOOL_MONTHS_.filter(function (m) { return new Date(today + 'T12:00:00').getMonth() === new Date(m + ' 1, 2000').getMonth(); })[0] || '' });
  cachePutBig_(key, text, 300);
  return text;
}
