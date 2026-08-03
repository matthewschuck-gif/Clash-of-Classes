/**
 * Clash of Classes -- Apps Script backend
 * File 2 of 2: doPost/doGet router + chunked read/write of the single app_data blob.
 *
 * DESIGN NOTE ON CORS (read before touching index.html):
 * Apps Script web apps don't let you set custom response headers, and they don't handle CORS
 * preflight (OPTIONS) requests. The fix: send requests with `Content-Type:
 * text/plain;charset=utf-8` instead of `application/json`. That keeps the browser from sending
 * a preflight at all, and Google's infrastructure serves script.google.com/exec responses
 * cross-origin without extra headers. The body is still just JSON.stringify(...) text; only
 * the header name changes. handleRequest_ below parses the body as JSON regardless of what
 * Content-Type was declared. Same pattern used for the Trail Journal migration.
 *
 * DESIGN NOTE ON SECURITY (unchanged from the live Supabase version, on purpose):
 * The current site protects writes with nothing but a Supabase anon key -- any browser can
 * already POST straight to the Postgres REST endpoint with that key and overwrite `app_data`;
 * the "admin password" that gates Settings/Admin in the UI is itself just a value stored
 * INSIDE the D blob (D.password) and only checked client-side in JS. It was never real
 * server-side access control. This migration ports that behavior exactly rather than quietly
 * hardening it -- every request here only needs the shared APP_TOKEN (same trust model as the
 * old anon key: visible in page source, filters out random scanners, not real authentication).
 * If you want server-side enforcement of the admin password on writes, that's a separate,
 * deliberate follow-up (see MIGRATION_RUNBOOK.md) -- it wasn't added silently here.
 *
 * DESIGN NOTE ON "LIVE UPDATING":
 * Supabase Realtime pushed changes over a WebSocket. Apps Script Web Apps have no push/
 * WebSocket equivalent -- there is no way around this, it's a hard platform limitation, not an
 * oversight. The frontend's `sb.channel(...).on(...).subscribe()` shim (in index.html) now
 * polls the lightweight `meta` action below every few seconds and only fetches the full
 * (potentially large) blob via `load` when something actually changed. This is "live" in
 * the sense that every open tab sees changes within a few seconds, but it is polling, not a
 * push -- see MIGRATION_RUNBOOK.md for the interval and quota trade-off.
 *
 * DESIGN NOTE ON THE `events` TAB (editable straight from Sheets, not just from the site):
 * `getAppDataMeta_` below uses the spreadsheet FILE's last-modified time (DriveApp), not a
 * single cell -- that's a deliberate choice, not an oversight. It means ANY edit anywhere in
 * the spreadsheet (including a teacher typing a new row, or fixing a point value, directly
 * into the `events` tab) is picked up by the same polling loop that already detects the site's
 * own saves, with no separate onEdit trigger needed. See readEventsTab_/writeEventsTab_.
 */

function doPost(e) {
  return handleRequest_(e, 'POST');
}

function doGet(e) {
  return handleRequest_(e, 'GET');
}

function handleRequest_(e, method) {
  try {
    let body;
    if (method === 'POST' && e.postData && e.postData.contents) {
      body = JSON.parse(e.postData.contents);
    } else {
      body = {
        action: e.parameter.action,
        token: e.parameter.token,
      };
    }

    checkAppToken_(body.token);

    let result;
    switch (body.action) {
      case 'load':
        result = loadAppData_('coc');
        break;
      case 'meta':
        result = getAppDataMeta_('coc');
        break;
      case 'save':
        result = saveAppData_('coc', body.value);
        break;
      default:
        throw new Error('Unknown action: ' + body.action);
    }

    return jsonOut_({ ok: true, data: result });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function checkAppToken_(token) {
  const expected = getProp_('APP_TOKEN');
  if (token !== expected) throw new Error('Invalid or missing token');
}

/**
 * Cheap change check for the polling loop: returns the whole SPREADSHEET FILE's last-modified
 * timestamp (one Drive metadata call, no reading/joining/parsing of any tab's actual rows).
 * This catches a save from the site (app_data tab) AND a manual edit to the events tab (or any
 * other tab) with the same single check -- see the design note above. Requires the Drive
 * scope, which you'll be asked to approve the first time you run setupSpreadsheet().
 */
function getAppDataMeta_(key) {
  const file = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId());
  return { updated_at: file.getLastUpdated().toISOString(), exists: true };
}

/**
 * Joins every chunk row for `key`, in chunk_index order, back into one JSON string and
 * parses it. Mirrors the old `select value from app_data where key = eq.<key>`. Then merges
 * in `events` from its own tab (see readEventsTab_) so the frontend sees one seamless `D`
 * object, exactly as before -- it never needs to know events live in a different tab now.
 */
function loadAppData_(key) {
  const sheet = getSheet_(TABS.APP_DATA);
  const data = sheet.getDataRange().getValues();
  const chunks = []; // {idx, chunk}
  let updatedAt = null;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) !== String(key)) continue;
    chunks.push({ idx: Number(data[i][1]), chunk: String(data[i][2]) });
    if (Number(data[i][1]) === 0) updatedAt = data[i][3] || null;
  }

  let value;
  if (chunks.length === 0) {
    value = null;
  } else {
    chunks.sort(function (a, b) { return a.idx - b.idx; });
    const json = chunks.map(function (c) { return c.chunk; }).join('');
    try {
      value = JSON.parse(json);
    } catch (e) {
      throw new Error('Stored app_data for key=' + key + ' is not valid JSON: ' + e.message);
    }
  }

  const events = readEventsTab_();
  if (value) {
    // One-time self-migration: if this blob still has events embedded from before the events
    // tab existed, and the events tab is empty, seed the tab from the blob instead of silently
    // discarding those events. After this runs once, the tab is the source of truth going
    // forward (see writeEventsTab_ in saveAppData_, which always fully replaces it).
    if (events.length === 0 && Array.isArray(value.events) && value.events.length > 0) {
      writeEventsTab_(value.events);
    } else {
      value.events = events;
    }
  }

  return { value: value, updated_at: updatedAt };
}

/**
 * Replaces every existing row for `key` with fresh chunk rows built from JSON.stringify(value).
 * Mirrors the old `upsert({key, value, updated_at})` -- delete-then-append is simplest and
 * correct here because a single logical row can grow or shrink its chunk count between saves
 * (e.g. someone removes a gallery photo), so an in-place partial overwrite could leave stale
 * trailing chunks behind if not handled carefully. Delete-then-append avoids that entirely.
 */
function saveAppData_(key, value) {
  if (value === undefined) throw new Error('save requires a value');
  const sheet = getSheet_(TABS.APP_DATA);
  const data = sheet.getDataRange().getValues();

  // Pull events out into their own tab (see writeEventsTab_) instead of chunking them into
  // the blob -- that's what keeps them plainly editable in the Sheets UI. Store the rest of
  // `value` (theme, gallery, links, spin config, admin password, etc.) as before.
  const events = Array.isArray(value.events) ? value.events : [];
  const rest = Object.assign({}, value);
  delete rest.events;
  writeEventsTab_(events);

  // Delete existing app_data rows for this key, bottom-up so row indices stay valid mid-loop.
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]) === String(key)) sheet.deleteRow(i + 1);
  }

  const json = JSON.stringify(rest);
  const now = nowIso_();
  const rows = [];
  for (let pos = 0, idx = 0; pos < json.length; pos += CHUNK_SIZE, idx++) {
    rows.push([key, idx, json.substring(pos, pos + CHUNK_SIZE), idx === 0 ? now : '']);
  }
  // Edge case: value serializes to an empty string (shouldn't normally happen since `D` is
  // always an object, but guard anyway) -- still write one row so meta/load can find it.
  if (rows.length === 0) rows.push([key, 0, '', now]);

  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 4).setValues(rows);
  return { saved: true, chunks: rows.length, events: events.length, updated_at: now };
}

/**
 * Reads every data row in the `events` tab back into the shape the frontend already expects
 * (see D.events in index.html: {id, name, type, date, note, pts:{BLACK,GOLD,GREY,PURPLE},
 * multiplier}). Backfills a missing `id` (blank because a teacher typed a new row by hand
 * rather than the site assigning one) and writes it back to the sheet so the id is stable on
 * the next read -- this is the only case where a read also writes.
 */
function readEventsTab_() {
  const sheet = getSheet_(TABS.EVENTS);
  const data = sheet.getDataRange().getValues();
  const headers = SCHEMA[TABS.EVENTS];
  const idCol = headers.indexOf('id');
  const events = [];
  let maxId = 100; // matches the frontend's own starting nextId

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    // Skip fully blank rows (e.g. stray formatting on an otherwise empty row).
    if (headers.every(function (h, c) { return row[c] === '' || row[c] === null; })) continue;

    let id = row[idCol];
    if (id === '' || id === null || id === undefined) {
      maxId += 1;
      id = maxId;
      sheet.getRange(i + 1, idCol + 1).setValue(id);
    } else {
      id = Number(id);
      if (id > maxId) maxId = id;
    }

    const evt = { id: id };
    headers.forEach(function (h, c) {
      if (h === 'id' || h === 'BLACK' || h === 'GOLD' || h === 'GREY' || h === 'PURPLE') return;
      let v = row[c];
      // Sheets silently converts a hand-typed date like "2026-08-15" into a real Date object
      // (cell formatting), not a string -- normalize back to the plain YYYY-MM-DD string the
      // frontend already expects everywhere else (new Date().toISOString().split('T')[0]).
      if (h === 'date' && Object.prototype.toString.call(v) === '[object Date]') {
        v = Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      }
      evt[h] = v;
    });
    evt.pts = {
      BLACK: Number(row[headers.indexOf('BLACK')]) || 0,
      GOLD: Number(row[headers.indexOf('GOLD')]) || 0,
      GREY: Number(row[headers.indexOf('GREY')]) || 0,
      PURPLE: Number(row[headers.indexOf('PURPLE')]) || 0,
    };
    if (evt.multiplier === '' || evt.multiplier === null) delete evt.multiplier;
    events.push(evt);
  }

  return events;
}

/**
 * Fully replaces the `events` tab's data rows from the given array. Same delete-then-append
 * strategy as saveAppData_, for the same reason: the number of events can shrink (an event
 * gets deleted from the admin UI) as easily as it can grow, so a partial in-place update could
 * leave stale trailing rows behind.
 */
function writeEventsTab_(events) {
  const sheet = getSheet_(TABS.EVENTS);
  const headers = SCHEMA[TABS.EVENTS];
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, headers.length).clearContent();
  if (!events || events.length === 0) return;

  const rows = events.map(function (evt) {
    const pts = evt.pts || {};
    return headers.map(function (h) {
      if (h === 'BLACK' || h === 'GOLD' || h === 'GREY' || h === 'PURPLE') return pts[h] || 0;
      const v = evt[h];
      return v === undefined || v === null ? '' : v;
    });
  });
  sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
}
