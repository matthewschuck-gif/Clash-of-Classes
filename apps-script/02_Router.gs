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
 * (potentially large) blob via `load` when `updated_at` actually changed. This is "live" in
 * the sense that every open tab sees changes within a few seconds, but it is polling, not a
 * push -- see MIGRATION_RUNBOOK.md for the interval and quota trade-off.
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
 * Cheap existence/change check: reads only the chunk_index=0 row's updated_at, never joins or
 * parses the full (possibly large) JSON blob. This is what the frontend's polling loop calls
 * every few seconds -- keep it fast.
 */
function getAppDataMeta_(key) {
  const sheet = getSheet_(TABS.APP_DATA);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(key) && Number(data[i][1]) === 0) {
      return { updated_at: data[i][3] || null, exists: true };
    }
  }
  return { updated_at: null, exists: false };
}

/**
 * Joins every chunk row for `key`, in chunk_index order, back into one JSON string and
 * parses it. Mirrors the old `select value from app_data where key = eq.<key>`.
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
  if (chunks.length === 0) return { value: null, updated_at: null };
  chunks.sort(function (a, b) { return a.idx - b.idx; });
  const json = chunks.map(function (c) { return c.chunk; }).join('');
  let value;
  try {
    value = JSON.parse(json);
  } catch (e) {
    throw new Error('Stored app_data for key=' + key + ' is not valid JSON: ' + e.message);
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

  // Delete existing rows for this key, bottom-up so row indices stay valid mid-loop.
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]) === String(key)) sheet.deleteRow(i + 1);
  }

  const json = JSON.stringify(value);
  const now = nowIso_();
  const rows = [];
  for (let pos = 0, idx = 0; pos < json.length; pos += CHUNK_SIZE, idx++) {
    rows.push([key, idx, json.substring(pos, pos + CHUNK_SIZE), idx === 0 ? now : '']);
  }
  // Edge case: value serializes to an empty string (shouldn't normally happen since `D` is
  // always an object, but guard anyway) -- still write one row so meta/load can find it.
  if (rows.length === 0) rows.push([key, 0, '', now]);

  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 4).setValues(rows);
  return { saved: true, chunks: rows.length, updated_at: now };
}
