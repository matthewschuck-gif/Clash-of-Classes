/**
 * Clash of Classes -- Apps Script backend
 * File 1 of 2: Config + one-time spreadsheet provisioning
 *
 * SETUP STEPS (do these once, in order -- see MIGRATION_RUNBOOK.md for the full walkthrough):
 * 1. Create a new Google Sheet in your school Workspace Drive (any name, e.g. "Clash of
 *    Classes Data").
 * 2. Open it -> Extensions -> Apps Script. This creates a bound Apps Script project.
 * 3. Paste each 0N_*.gs file from this package into the Apps Script editor as its own file
 *    (File -> New -> Script file, name it to match, e.g. "01_Config").
 * 4. Project Settings -> Script Properties, add:
 *      APP_TOKEN  = <any random string you make up> (shared token, see 02_Router.gs)
 * 5. Run `setupSpreadsheet` once from the function dropdown. Approve the permission prompts
 *    (this includes a Drive scope now, used only to read the spreadsheet's last-modified time
 *    for change detection -- see getAppDataMeta_ in 02_Router.gs). This creates the `app_data`
 *    and `events` tabs with correct headers.
 * 6. Deploy ONCE as a Web App: Execute as "Me", Who has access "Anyone". Copy the resulting
 *    /exec URL -- it goes into index.html (see 02_Router.gs header for why only ONE
 *    deployment, never a domain-restricted one).
 *
 * ARCHITECTURE NOTE -- why this differs from a typical multi-table migration:
 * Clash of Classes doesn't use several relational tables like Trail Journal did. Almost the
 * entire site's state (squads, gallery, links, theme, spin wheel, admin password, etc.) is one
 * big JSON object, called `D` in the frontend, stored as a SINGLE row in a Postgres table
 * `app_data` (columns: key, value jsonb, updated_at) under key='coc'. The Sheet replacement
 * mirrors that table 1:1 in shape (see SCHEMA.app_data below) but adds one wrinkle Trail
 * Journal never needed: Google Sheets cells cap out at 50,000 characters, and `D` can exceed
 * that once gallery photos (uploaded as base64 data-URIs, see FileReader/readAsDataURL in
 * index.html) are added. So the JSON string is CHUNKED across multiple rows instead of living
 * in one cell -- see chunk_index below and saveAppData_/loadAppData_ in 02_Router.gs.
 *
 * ONE PIECE OF `D` GETS ITS OWN TAB, ON PURPOSE: `D.events` -- the squad point ledger -- is
 * pulled OUT of the JSON blob and given a real, human-editable tab (`events`, one row per
 * event, plain columns for each squad's points). That's what makes "just edit a cell in the
 * Sheet and the site updates" actually work for scores: nobody can usefully hand-edit a number
 * buried inside an opaque chunked JSON blob, but editing a plain number in a spreadsheet column
 * is exactly the point-of-Sheets-as-a-backend experience. See readEventsTab_/writeEventsTab_ in
 * 02_Router.gs for how this splits out of, and merges back into, `D` transparently -- the
 * frontend never needed to change for this, it still just sends/receives one `D` object.
 */

const TABS = {
  APP_DATA: 'app_data',
  EVENTS: 'events',
  METRICS: 'metrics',
  HELP: 'How To Use This Sheet',
};

// Mirrors the original Postgres `app_data` table's columns (key, value, updated_at), plus
// chunk_index -- the one addition needed to get around the Sheets cell-size limit (see note
// above). There is no separate "save_id" column because the original code already stores its
// own echo-suppression id (`_saveId`) INSIDE the `value` JSON itself, not as a separate
// Postgres column -- so it survives the migration for free, no schema change needed.
//
// `events` mirrors the shape of each object already pushed into D.events by the frontend
// (see addEvent() etc. in index.html: {id, name, type, date, note, pts:{BLACK,GOLD,GREY,
// PURPLE}, multiplier}) -- just flattened into plain columns instead of a nested pts object,
// which is what makes it directly editable in the Sheets UI.
const SCHEMA = {
  [TABS.APP_DATA]: ['key', 'chunk_index', 'value_chunk', 'updated_at'],
  [TABS.EVENTS]: ['id', 'name', 'type', 'date', 'note', 'multiplier', 'BLACK', 'GOLD', 'GREY', 'PURPLE'],
  // One row per squad. Mirrors D.metrics.{attendance,referral,grades}[SQUAD] -- the current,
  // in-progress month's percentages (see the monthly-metrics ranking design in index.html:
  // rankPoints_/finalizeMonth). Editing a percentage here updates the site's live preview the
  // same way editing it in Admin -> Metrics does; it does NOT award points by itself. Points
  // only lock in when an admin clicks "Finalize This Month" on the site, same as if the
  // numbers had been typed into the Admin form instead of this tab. The month LABEL (e.g.
  // "September 2026") is intentionally not a column here -- it stays admin-entered on the site
  // since it changes far less often than the percentages do.
  [TABS.METRICS]: ['squad', 'attendance', 'referral', 'grades'],
};

// Stay safely under Sheets' 50,000-char single-cell limit.
const CHUNK_SIZE = 40000;

function setupSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(SCHEMA).forEach(function (tabName) {
    let sheet = ss.getSheetByName(tabName);
    if (!sheet) sheet = ss.insertSheet(tabName);
    const headers = SCHEMA[tabName];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#490e6f').setFontColor('#ffffff');
  });
  const def = ss.getSheetByName('Sheet1');
  if (def && ss.getSheets().length > 1) ss.deleteSheet(def);
  ensureMetricsRows_();
  setupHelpTab_();
  SpreadsheetApp.getUi().alert('Done -- app_data, events, metrics, and "How To Use This Sheet" tabs are set up. Clash of Classes has no separate PII tab (unlike Trail Journal), so there is no extra sharing/protection step required here. You can now add/edit rows directly in the events and metrics tabs and the site will pick them up within a few seconds. Staff directions live in the "How To Use This Sheet" tab -- start there.');
}

// Pre-seeds the metrics tab with one row per squad (BLACK/GOLD/GREY/PURPLE) so staff just
// overwrite numbers instead of having to know which squads to add. Only adds rows for squads
// that aren't already present, so re-running setupSpreadsheet never overwrites percentages
// someone already entered.
function ensureMetricsRows_() {
  const sheet = getSheet_(TABS.METRICS);
  const squads = ['BLACK', 'GOLD', 'GREY', 'PURPLE'];
  const data = sheet.getDataRange().getValues();
  const existing = {};
  for (let i = 1; i < data.length; i++) {
    const sq = String(data[i][0] || '').trim().toUpperCase();
    if (sq) existing[sq] = true;
  }
  const missing = squads.filter(function (sq) { return !existing[sq]; });
  if (missing.length) {
    const rows = missing.map(function (sq) { return [sq, 0, 0, 0]; });
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 4).setValues(rows);
  }
}

// Human-readable directions for whoever is actually entering events/points day to day --
// deliberately plain-language, not aimed at whoever set up the Apps Script backend. Safe to
// re-run (setupSpreadsheet is idempotent): this rewrites the instructions text but never
// touches app_data or events rows.
function setupHelpTab_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(TABS.HELP);
  if (!sheet) sheet = ss.insertSheet(TABS.HELP);
  else sheet.clear();

  const rows = [
    ['Clash of Classes -- How To Use This Sheet'],
    [''],
    ['This spreadsheet is the live backend for clashofclasses.org. Two tabs matter for'],
    ['day-to-day use: "events" is where you add points for activities, and "metrics" is where'],
    ['you enter each month\'s attendance / referral / grades percentages. Changes here show up'],
    ['on the website automatically within about 5-10 seconds -- no need to tell anyone or'],
    ['refresh anything on your end.'],
    [''],
    ['HOW TO ADD AN EVENT / GIVE POINTS'],
    ['1. Open the "events" tab (the tab bar is at the very bottom of this spreadsheet).'],
    ['2. Add a NEW ROW at the bottom of the data (do not insert rows in the middle).'],
    ['3. Fill in these columns for that row:'],
    ['     id       -- leave BLANK. The sheet fills this in automatically.'],
    ['     name     -- what the event/activity was called, e.g. "Spirit Week Kickoff"'],
    ['     type     -- a short category label, e.g. "spirit", "academic", "attendance"'],
    ['     date     -- the date of the event, format YYYY-MM-DD, e.g. 2026-09-15'],
    ['     note     -- optional, any extra context'],
    ['     multiplier -- optional, leave blank unless you specifically want to double/triple'],
    ['                   this event\'s points (e.g. enter 2 for double points)'],
    ['     BLACK / GOLD / GREY / PURPLE -- the number of points EACH squad earned for this'],
    ['                   specific event. Enter 0 for a squad that did not earn points here.'],
    ['4. That\'s it -- do not touch any other tab. The website will update on its own.'],
    [''],
    ['HOW TO FIX A MISTAKE'],
    ['- Wrong number? Just edit the number in that cell directly -- the site will pick up the'],
    ['  correction automatically, same as any other edit.'],
    ['- Need to remove an event entirely? Delete that whole row (right-click the row number ->'],
    ['  Delete row). Don\'t just clear the cells and leave a blank row behind.'],
    [''],
    ['HOW TO ENTER MONTHLY METRICS (ATTENDANCE / REFERRAL / GRADES)'],
    ['1. Open the "metrics" tab. There\'s already one row per squad (Black, Gold, Grey,'],
    ['   Purple) -- don\'t add or delete rows here, just edit the numbers.'],
    ['2. Enter each squad\'s percentage for that month in the attendance / referral / grades'],
    ['   columns (whole numbers or decimals both work, e.g. 94 or 94.5).'],
    ['3. That\'s it for entering the numbers -- the website\'s "Monthly Metrics" preview'],
    ['   updates automatically, same as the events tab.'],
    ['4. IMPORTANT: entering numbers here does NOT award points by itself. Points only lock'],
    ['   in when an admin goes to clashofclasses.org -> Admin -> Metrics tab and clicks'],
    ['   "Finalize This Month & Award Points." That is on purpose -- it means you can update'],
    ['   these numbers throughout the month as new reports come in without accidentally'],
    ['   awarding points early. The highest squad in each of the 3 metrics earns 250 points,'],
    ['   2nd place 125, 3rd place 50, and last place 0 (a tie splits the higher spot\'s points).'],
    ['5. The month LABEL (e.g. "September 2026") is entered on the website, not in this sheet'],
    ['   -- Admin -> Metrics tab -> Month Label.'],
    [''],
    ['HOW TO RESET/START A NEW SCHOOL YEAR -- IMPORTANT, READ THIS FIRST'],
    ['Do NOT reset the season by deleting all the rows in the "events" tab yourself. Because of'],
    ['how this sheet talks to the website, clearing the tab by hand will NOT actually reset the'],
    ['scores -- the website will just quietly restore the old numbers back into this tab within'],
    ['moments, and it will look like nothing happened (or like your edit got undone).'],
    [''],
    ['Instead: go to clashofclasses.org, log into Admin, and use the "Reset All Scores" button'],
    ['there. That is the only way that safely resets everything (it clears this tab AND the'],
    ['website\'s stored data together, in sync). If you\'re not sure how to get into Admin, ask'],
    ['whoever manages the site.'],
    [''],
    ['DO NOT TOUCH'],
    ['The "app_data" tab is not meant to be edited by hand -- it stores the rest of the site\'s'],
    ['settings (theme, gallery, admin password, etc.) as encoded data, not as plain readable'],
    ['numbers or text. Editing it directly can break the site. If something there needs to'],
    ['change, do it through the website\'s own Admin panel instead.'],
  ];

  sheet.getRange(1, 1, rows.length, 1).setValues(rows);
  sheet.getRange(1, 1).setFontWeight('bold').setFontSize(14).setFontColor('#490e6f');

  // Bold section headers by matching text rather than hardcoded row numbers, so this
  // doesn't silently drift out of sync if the copy above ever gets edited.
  const boldSections = ['HOW TO ADD AN EVENT / GIVE POINTS', 'HOW TO FIX A MISTAKE', 'HOW TO ENTER MONTHLY METRICS (ATTENDANCE / REFERRAL / GRADES)', 'DO NOT TOUCH'];
  const warnSections = ['HOW TO RESET/START A NEW SCHOOL YEAR -- IMPORTANT, READ THIS FIRST'];
  rows.forEach(function (r, i) {
    const text = r[0];
    if (boldSections.indexOf(text) !== -1) sheet.getRange(i + 1, 1).setFontWeight('bold');
    if (warnSections.indexOf(text) !== -1) sheet.getRange(i + 1, 1).setFontWeight('bold').setFontColor('#b91c1c');
  });

  sheet.setColumnWidth(1, 760);
  sheet.setFrozenRows(1);
  ss.setActiveSheet(sheet);
  ss.moveActiveSheet(1);
}

function getSheet_(tabName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(tabName);
  if (!sheet) throw new Error('Sheet tab not found: ' + tabName + ' -- run setupSpreadsheet() first.');
  return sheet;
}

// Same as getSheet_ but returns null instead of throwing when the tab doesn't exist yet.
// Used for tabs added after the initial setup (like METRICS) so that a deployment redeployed
// with newer router code doesn't immediately break every request just because the user
// hasn't re-run setupSpreadsheet() on their existing sheet yet -- it just behaves as if that
// feature's data isn't there yet, same as a brand new site.
function getSheetSafe_(tabName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(tabName) || null;
}

function getProp_(key) {
  const v = PropertiesService.getScriptProperties().getProperty(key);
  if (!v) throw new Error('Missing Script Property: ' + key + ' -- set it in Project Settings -> Script Properties.');
  return v;
}

function nowIso_() {
  return new Date().toISOString();
}
