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
 * 5. Run `setupSpreadsheet` once from the function dropdown. Approve the permission prompts.
 *    This creates the single `app_data` tab with correct headers.
 * 6. Deploy ONCE as a Web App: Execute as "Me", Who has access "Anyone". Copy the resulting
 *    /exec URL -- it goes into index.html (see 02_Router.gs header for why only ONE
 *    deployment, never a domain-restricted one).
 *
 * ARCHITECTURE NOTE -- why this differs from a typical multi-table migration:
 * Clash of Classes doesn't use several relational tables like Trail Journal did. The entire
 * site's state (squads, events, gallery, links, theme, spin wheel, admin password, etc.) is
 * one big JSON object, called `D` in the frontend, stored as a SINGLE row in a Postgres table
 * `app_data` (columns: key, value jsonb, updated_at) under key='coc'. The Sheet replacement
 * mirrors that table 1:1 in shape (see SCHEMA below) but adds one wrinkle Trail Journal never
 * needed: Google Sheets cells cap out at 50,000 characters, and `D` can exceed that once
 * gallery photos (uploaded as base64 data-URIs, see FileReader/readAsDataURL in index.html)
 * are added. So the JSON string for a given key is CHUNKED across multiple rows instead of
 * living in one cell -- see chunk_index below and saveAppData_/loadAppData_ in 02_Router.gs.
 */

const TABS = {
  APP_DATA: 'app_data',
};

// Mirrors the original Postgres `app_data` table's columns (key, value, updated_at), plus
// chunk_index -- the one addition needed to get around the Sheets cell-size limit (see note
// above). There is no separate "save_id" column because the original code already stores its
// own echo-suppression id (`_saveId`) INSIDE the `value` JSON itself, not as a separate
// Postgres column -- so it survives the migration for free, no schema change needed.
const SCHEMA = {
  [TABS.APP_DATA]: ['key', 'chunk_index', 'value_chunk', 'updated_at'],
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
  SpreadsheetApp.getUi().alert('Done -- app_data tab created with headers. Clash of Classes has no separate PII tab (unlike Trail Journal), so there is no extra sharing/protection step required here.');
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
