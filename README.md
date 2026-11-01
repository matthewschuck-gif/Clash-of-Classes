# Clash of Classes

Ephrata Middle School's squad competition site, clashofclasses.org.

- **The site and its data** run from one Apps Script project, bound to the **Clash of Classes
  Points** Google Sheet (`1wF01Limr7uO36DRGC6qh3XfYA7dBdXEfGRWuasmEf0w`).
- **clashofclasses.org** is a tiny GitHub Pages page that shows that web app full-screen, so the
  address students know stays the same. It stores nothing.

## What's where
| File | What it does |
|---|---|
| `apps-script/01_Config.gs` | Sets up the tabs (`app_data`, `events`, `metrics`, `Settings`, `Challenges`, how-to) and holds the Settings defaults |
| `apps-script/02_Router.gs` | Reads and writes the site data. `events` and `metrics` are plain, editable tabs. `?action=load` is the old read-only API |
| `apps-script/03_Web.gs` | Serves the page. Also handles admin sign-in (scrambled password), merging saves, poll votes, Trail Journal stats and the Sheet menu |
| `apps-script/04_Students.gs` | The My Squad page, automatic squad challenges and bonuses, and the TV feed |
| `apps-script/05_Generals.gs` | Generals HQ (800# + `GENERALS_CODE`): squad roster with who isn't placed yet, event sign-ups, Mayhem Week checklist, service project. Also MLC roles on My Squad (Vanguard stays private) |
| `apps-script/06_Polls.gs` | Daily Polls (This or That and Trivia) from the Daily Polls tab. Students vote with their 800#, and squads earn points for each voter plus each winner (DAILY_POLL_POINTS). Also the read-aloud announcement |
| `apps-script/Index.html` | The page markup |
| `apps-script/Styles.html`, `SiteCss.html` | Styles |
| `apps-script/App.html` | The main site script |
| `apps-script/Widgets.html` | Freudenfreude Ticker and Growth board |
| `apps-script/Student.html` | My Squad, challenge bars and the two new TV slides |
| `apps-script/Generals.html` | The Generals HQ page: roster and sign-ups, Mayhem checklist, hype script, full-screen squad meeting, service project |
| `tools/wrapper-index.html`, `tools/wrapper-sw.js` | Become `index.html` and `sw.js` at cutover. The new `sw.js` clears the old site's offline cache |

## Where the numbers come from
- **Points:** the `events` tab. Trait Dough, Mountaineer Mindset, Trail Journal and Define a
  Climb-er (STAR growth) each add rows there. Admins add rows on the site.
- **Monthly metrics:** the `metrics` tab, filled every hour by Define a Climb-er's squad sync
  from PowerSchool. Points lock in when an admin clicks **Finalize This Month**.
- **My Squad:** the `Student Board` tab, written every hour by Define a Climb-er.
  - Positive things only.
  - A student types their 800# and sees their own row and nothing else.
  - Trail Journal is hidden unless the Climb-er setting `BOARD_SHOW_JOURNAL` is ON.
- **Squad Board:** a summary tab, also written by Climb-er.
- **Challenges:** the `Challenges` tab. Progress fills in by itself from `events`.

## Settings tab (Clash sheet)
- `SITE_TITLE`
- `SITE_URL`
- `POLL_SECONDS`: how often open pages check for new points
- `TJ_SHEET_ID`
- `STUDENT_BOARD_TAB`
- `MY_SQUAD_LOOKUPS_PER_10MIN`
- `BADGE_LIST`
- `CHALLENGE_BONUSES`: OFF by default. Turn on from the menu.
- `TV_SHOW_NAMES`
- `EVENT_LABELS`

Badge **rules** live on the Climb-er Settings tab (`BOARD_BADGES`).

## Admin password
- Stored only as a salted hash in Script Properties.
- Set it from **Clash of Classes › Set admin password…**, or from Admin › Settings on the site.
- The first time someone signs in after the move, the old site password works once and is then
  saved scrambled.

## Tests
```bash
node tests/clash.test.js
```
```bash
node tests/build-preview.js
```
The first runs the server tests with fake Google services. The second serves a local preview at
http://localhost:3041; its admin password is `preview`.
