# Clash of Classes → Google Sheets/Apps Script Migration Runbook

Same migration as Trail Journal, applied to this site. Everything is built and pushed to the
`migration/google-appscript` branch. This is the checklist for your side: setting up the actual
Google Sheet + Apps Script project (I don't have write access to your Workspace), pointing the
frontend at it, testing, and cutting over. Nothing here requires touching code.

## Where everything lives

- **`main`** — untouched, still the live Supabase-backed app exactly as it was.
- **`archive/supabase-original`** — a frozen copy of that same state, as an explicit rollback point.
- **`migration/google-appscript`** — the new Google-backed version. Both `.gs` backend files
  plus the rewired `index.html` are here, fully committed.

## How this migration differs from Trail Journal's

Trail Journal had 7 relational-ish Postgres tables (one row per reflection, per panel session,
etc.), so that migration built a generic table-by-table CRUD router. Clash of Classes only ever
had **one** Postgres table, `app_data`, holding a **single row** (`key='coc'`) whose `value`
column is the *entire* site state as one JSON blob — squads, events, gallery, links, theme,
spin wheel, admin password, everything. That's a simpler shape in some ways, but it introduces
one problem Trail Journal never had: gallery photos are uploaded as base64 data-URIs and stored
inline in that blob, so the JSON string can exceed a Google Sheet cell's 50,000-character limit.
The backend handles this by chunking the JSON across multiple rows instead of one cell (see
`CHUNK_SIZE` in `apps-script/01_Config.gs`) — invisible to you, but worth knowing it's there.

The other real difference: the live site used Supabase **Realtime** (a WebSocket) so every open
device saw changes within a second or two with zero effort. Apps Script Web Apps have no push/
WebSocket equivalent — this is a hard platform limitation, not something I can route around.
The rewired frontend now **polls** a lightweight `meta` endpoint every 5 seconds and only pulls
the full (potentially large) blob when something actually changed. In practice this means
everyone still sees updates within about 5 seconds instead of instantly — worth testing on your
end to confirm 5 seconds feels "live enough" for how this gets used (projector display, spin
wheel, live leaderboard, etc.). If not, `GAS_POLL_MS` near the top of the rewired `index.html`
is the one number to change — lower it for snappier updates, at the cost of more requests
against your Apps Script quota.

## Yes — editing the Sheet directly updates the site's points

One deliberate addition beyond a straight port: squad points (`D.events` in the original code)
get their own tab, **`events`**, with plain columns — `id, name, type, date, note, multiplier,
BLACK, GOLD, GREY, PURPLE` — instead of being buried inside the opaque JSON blob in `app_data`.
That's what makes hand-editing actually work: nobody can usefully type into an opaque chunked
JSON cell, but typing a number into a normal spreadsheet column is exactly the point.

What this means in practice:
- Add a new row to `events` (any squad's points, a name, a date) → it becomes a real scored
  event on the site within about 5 seconds, no admin login needed. Leave `id` blank — the
  backend assigns one automatically the first time it's read and writes it back into the row.
- Edit a points cell on an existing row → that event's contribution to the squad total updates.
- Delete a row → that event is gone from the site, same as deleting it from the admin UI.
- The site's own admin actions (Add Event, spin wheel, Squad Generals, etc.) still work exactly
  as before and write into this same tab — editing from the Sheet and editing from the site are
  two doors into the same room, not two separate systems to keep in sync.
- The polling `meta` check (`getAppDataMeta_` in `02_Router.gs`) now looks at the whole
  spreadsheet **file's** last-modified time (via `DriveApp`), not one cell — that's what lets a
  manual edit to `events` get picked up by the same 5-second poll as a site-initiated save, with
  no extra trigger to set up. This does mean `setupSpreadsheet()`'s permission prompt now
  includes a Drive scope; approve it like the others.
- Small, acceptable trade-off: every site-initiated save fully replaces all rows in `events` (to
  cleanly handle events being deleted, not just added). If someone is mid-edit in the `events`
  tab at the exact moment the site saves, that edit could be overwritten — the same kind of risk
  any shared spreadsheet has when two people touch it at once. Not expected to matter for normal
  classroom-pace use, just worth knowing.

Everything else — theme colors, gallery photos, links, spin wheel config, admin password, squad
names — stays in the chunked `app_data` blob, not a plain tab, because that data isn't the kind
a teacher would want to hand-edit as spreadsheet rows.

## Part 1 — One-time Google setup

1. In your school Workspace Drive, create a new Google Sheet (any name — e.g. "Clash of Classes Data").
2. Open it → **Extensions → Apps Script**.
3. Delete the default empty `Code.gs` file it creates.
4. Create 2 new script files (**File → New → Script file**) named exactly:
   `01_Config`, `02_Router`
   and paste in the matching content from `apps-script/` in the `migration/google-appscript` branch.
5. **Project Settings (gear icon) → Script Properties**, add:
   | Property | Value |
   |---|---|
   | `APP_TOKEN` | any random string you make up — this is the shared token the frontend sends on every request |

   That's the only Script Property this one needs — no AI key, no email key, no admin
   password (see the security note below for why).
6. In the function dropdown at the top, select `setupSpreadsheet`, click **Run**. Approve the
   permission prompts — this now includes a Drive scope (used only to check the spreadsheet
   file's last-modified time for the polling/live-update check, see below). This creates the
   `app_data` and `events` tabs with correct headers. Unlike Trail Journal's `incident_reports`
   tab, there's no separate PII tab here to lock down — nothing in either tab contains
   individual student names tied to disciplinary records, just the site's public content and
   squad point totals.
7. **Deploy once** (Deploy → New deployment → type "Web app"): Execute as **Me**, Who has access
   **Anyone**. Copy this single URL.

   **Why only one deployment:** exactly the same reason as Trail Journal — a domain-restricted
   ("Anyone within your domain") deployment routes cross-origin requests through an extra Google
   auth-check hop that never returns CORS headers, so `clashofclasses.org` could never read the
   response. Confirmed by testing on the Trail Journal migration, applies identically here.

   **Use the plain URL, not the domain-vanity one.** After deploying, Google may show you the
   URL in the form `https://script.google.com/a/macros/easdpa.org/s/DEPLOYMENT_ID/exec`. Do NOT
   use that one — for this district it returned a generic "unable to open the file" error for
   every request, even fully logged out, even with the deployment correctly set to "Anyone".
   Use the plain canonical form instead: `https://script.google.com/macros/s/DEPLOYMENT_ID/exec`
   (same deployment ID, just without the `/a/macros/easdpa.org` prefix) — confirmed working via
   a logged-out incognito test. This was NOT a district Workspace/IT policy blocking anonymous
   access (a wrong theory floated and then ruled out during this migration, based on Trail
   Journal's own history showing plain "Anyone" deployments work fine on this domain).

   **A much deeper issue past the URL format: no JavaScript-initiated request works against
   this deployment at all, only a real browser navigation.** Three things were tried, in order:

   1. `fetch()` (both GET and POST) — failed every time, with Google's generic "unable to open
      the file" error, even though the identical URL worked fine pasted directly into a browser
      address bar.
   2. A dynamically-inserted `<script src="...">` tag (classic JSONP) — also failed; the
      script's `onerror` fired outright, it couldn't even load.
   3. A hidden `<iframe>` navigated to the URL (for reads), and a hidden `<form target="an
      iframe">` submission (for writes), with the Apps Script response rendering as a tiny HTML
      page that calls `window.parent.postMessage(...)` to hand data back — **this is what
      actually works.** An iframe load is a genuine document navigation in its own browsing
      context, not a `fetch()`/script-tag resource load, so it's the one mechanism that behaves
      like the thing that was reliably succeeding all along.

   See `gasIframeCall_` in `index.html` (and the duplicated copies in the Freudenfreude Ticker
   and Growth Leaderboard widgets) and `respondOut_`/`handleRequest_` in `02_Router.gs` for how
   this works, including a `reqId` on every call — `window`'s `message` event listener is
   global to the whole page, not scoped to one iframe, so without a request id to match a
   response back to the specific call that asked for it, a reply meant for one in-flight call
   (say, the 5-second `meta` poll) could get delivered to a different concurrent call (say, a
   user clicking "Add Event" at the same moment) instead.

   **One more piece was needed even after switching to iframes: `HtmlService` output blocks
   cross-origin framing by default** (the `X-Frame-Options` equivalent), which showed up as the
   iframe requests failing with a 403. The fix is `.setXFrameOptionsMode(HtmlService.
   XFrameOptionsMode.ALLOWALL)` on the response in `respondOut_` — without it, the whole
   iframe/postMessage approach can't work no matter how correct everything else is.

   **And one more layer past the 403 fix: `postMessage` aimed at `window.parent` gets silently
   dropped.** After the 403 was fixed, calls still timed out, and devtools showed lines like
   `...-mae_html_user_bin_i18n_mae_html_user.js: dropping postMessage.. was from unexpected
   window`. That script is Google's own — `HtmlService`'s sandbox mode (mandatory since 2020,
   can't be disabled) wraps every served page in a Google-controlled bridge/wrapper frame, so
   from inside our injected `<script>`, `window.parent` is that wrapper, not the real host page,
   and the wrapper's own bridge script intercepts and discards postMessage calls that don't
   match its internal protocol. The fix is to target `window.top` instead (falling back to
   `window.parent` only if it differs) — `window.top` always resolves to the true outermost
   page (clashofclasses.org), skipping the wrapper frame entirely.

## Part 2 — Point the frontend at your URLs

In `index.html`, find these two lines (near the top of the big `<script type="module">` block,
just above where the old `SB_URL`/`SB_KEY` used to be):

```js
const GAS_URL='PASTE_YOUR_APPS_SCRIPT_WEB_APP_URL_HERE';
const GAS_TOKEN='PASTE_YOUR_APP_TOKEN_HERE';
```

Paste in your `/exec` URL from Part 1 step 7, and the `APP_TOKEN` value you chose in step 5.
That's the only edit needed anywhere in the file.

## Part 3 — Test before touching production

**Do this on a separate test copy of the site first**, not on clashofclasses.org directly.

Correction from an earlier draft of this runbook: do **not** enable GitHub Pages on the
`migration/google-appscript` branch to test it. This repo serves clashofclasses.org as a GitHub
Pages **custom domain** (see the `CNAME` file) — Pages only serves one branch at a time for the
whole site, so switching the Pages source branch would put the untested migration code on the
live domain immediately, not a safe side-by-side preview.

The actually-safe way: run it locally on your own computer, completely outside GitHub Pages.

1. Get the `migration/google-appscript` branch onto your computer:
   - No git needed: go to the repo on GitHub, switch the branch dropdown (top-left) to
     `migration/google-appscript`, then **Code → Download ZIP**. Unzip it.
   - Or, if you're comfortable with git: `git clone -b migration/google-appscript
     https://github.com/matthewschuck-gif/Clash-of-Classes.git coc-test`
2. `index.html` references `/styles.css`, `/manifest.json`, etc. with root-absolute paths, so
   just double-clicking `index.html` to open it as a `file://` URL will load with no styling —
   it needs to be served from a tiny local web server instead, which takes one command:
   - Open Terminal / Command Prompt in that folder.
   - Run: `python3 -m http.server 8000` (Mac/Linux) or `py -m http.server 8000` (Windows, if
     Python's installed). No Python? `npx serve` works too if Node is installed.
   - Open `http://localhost:8000/` in your browser.
3. This talks to your real Apps Script backend and real Google Sheet over the internet — it's
   only the frontend that's local. Nothing here touches clashofclasses.org or its Pages
   deployment at all.

Checklist:
- [ ] Site loads with existing data (or starts fresh with defaults if the Sheet is empty —
      confirm a first save populates the `app_data` tab).
- [ ] Add an event / award points from Admin → confirm the score updates on screen and a
      matching row appears in the `events` tab (not `app_data` — see the section above).
- [ ] Type a new row directly into the `events` tab (fill in a squad's points, leave `id`
      blank) → confirm it appears on the site within ~5 seconds AND that an id got written back
      into that row. Then edit one of its point values directly in the Sheet → confirm the
      squad total updates on the site too.
- [ ] Open the site in two tabs/devices at once: make a change in one, confirm the other
      reflects it within ~5 seconds (this is the polling replacing Realtime — see note above).
- [ ] Spin wheel: enter the admin password, spin, confirm result saves and broadcasts to the
      TV display view (`?tv=1`).
- [ ] Gallery: upload a photo, confirm it saves and displays (this is the base64/chunking path
      — the one most likely to surface a size-limit issue if something's off).
- [ ] Growth panel, Freudenfreude Ticker, and Growth Leaderboard widgets all still render COC's
      own point totals correctly.
- [ ] Admin Settings: change the admin password, confirm it takes effect (log out/in).
- [ ] Confirm Trail Journal itself still works normally — this migration doesn't touch it.

## Part 4 — Cut over

Once everything above checks out: merge `migration/google-appscript` into `main` (a normal PR
merge), turn GitHub Pages back to serving `main`, and clashofclasses.org goes live on the new
backend. `archive/supabase-original` stays as your rollback point indefinitely.

## Known follow-up items (not fixed here, flagged on purpose rather than silently changed)

- **Trail Journal cross-widget fetch**: the Growth panel widget on the home/culture page pulls
  live reflection/panel-session stats directly from Trail Journal's Supabase project
  (`TJ_URL`/`TJ_ANON` in `index.html`, left untouched). If Trail Journal has already cut over
  (or cuts over later) to its own Apps Script backend, that Supabase project will eventually
  auto-pause and be deleted, and this widget will start failing silently (it already has a
  try/catch that shows "Could not load data" rather than crashing the page, but the stats will
  stop updating). Let me know when Trail Journal's cutover happens and I'll repoint this widget
  to its new backend as a small follow-up.
- **Admin password isn't real server-side security, same as before**: the password checked in
  Admin/Settings lives inside the saved JSON blob itself and is only checked in the browser —
  true in the original Supabase version (anyone with the anon key could already write over it
  directly) and true here (anyone with `APP_TOKEN`, visible in page source, can write over it
  via `action=save`). This migration ported that behavior exactly rather than quietly changing
  it. If you want real write protection later, that needs a different, deliberate design
  (e.g., a separate server-side `ADMIN_PASSWORD` Script Property checked before `save` — happy
  to build that as a follow-up if you want it).
- **Polling interval is a judgment call, not a fixed requirement**: 5 seconds (`GAS_POLL_MS` in
  `index.html`) balances "feels live" against request volume. Adjust after Part 3 testing if it
  feels too slow or you'd rather trade a snappier feel for lower request volume.

## Housekeeping

- **GitHub token**: if you gave me one for this push, revoke it now (Settings → Developer
  settings → Personal access tokens).
- **Old Supabase project**: same advice as Trail Journal's runbook — free-tier Supabase
  projects auto-pause after about 7 days idle and are *permanently deleted* after about 90 days
  paused. If Clash of Classes' current data needs to survive as a long-term archive, either open
  the project every couple months, export it, or ask me to set up a scheduled reminder.
