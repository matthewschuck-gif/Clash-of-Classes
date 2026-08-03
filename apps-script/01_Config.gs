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
  SpreadsheetApp.getUi().alert('Done -- app_data and events tabs created with headers. Clash of Classes has no separate PII tab (unlike Trail Journal), so there is no extra sharing/protection step required here. You can now add/edit rows directly in the events tab and the site will pick them up within a few seconds.');
}

function getSheet_(tabName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(tabName);
  if (!sheet) throw new Error('Sheet tab not found: ' + tabName + ' -- run setupSpreadsheet() first.');
  return sheet;
}

function getProp_(key) {
  const v = PropertiesService.getScriptProperties().getProperty(key);
  if (!v) throw new Error('Missing Script Property: ' + key + ' -- set it in Project Settings -> Script Properties.');
  return v;
}

function nowIso_() {
  return new Date().toISOString();
}
