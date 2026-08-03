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
   permission prompts. This creates the single `app_data` tab with correct headers. Unlike
   Trail Journal's `incident_reports` tab, there's no separate PII tab here to lock down —
   `app_data` doesn't contain individual student names tied to disciplinary records, just the
   site's public content and squad point totals.
7. **Deploy once** (Deploy → New deployment → type "Web app"): Execute as **Me**, Who has access
   **Anyone**. Copy this single URL.

   **Why only one deployment:** exactly the same reason as Trail Journal — a domain-restricted
   ("Anyone within your domain") deployment routes cross-origin requests through an extra Google
   auth-check hop that never returns CORS headers, so `clashofclasses.org` could never read the
   response. Confirmed by testing on the Trail Journal migration, applies identically here.

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
Easiest way: enable GitHub Pages for the `migration/google-appscript` branch temporarily
(Settings → Pages → Branch), which gives you a
`matthewschuck-gif.github.io/Clash-of-Classes/` URL to test against without affecting the live
domain.

Checklist:
- [ ] Site loads with existing data (or starts fresh with defaults if the Sheet is empty —
      confirm a first save populates the `app_data` tab).
- [ ] Add an event / award points from Admin → confirm the score updates on screen and a new
      set of chunk rows appears in `app_data` (old chunk rows for the same key get replaced,
      not accumulated — check row count doesn't grow forever after repeated saves).
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
