# Session recap (Nov 13, 2026): Clash of Classes winter update

## Where things stand
- **Branch:** `claude/bold-bell-ql60m1` (pushed). **PR:** https://github.com/matthewschuck-gif/Clash-of-Classes/pull/1 into `main`, not merged yet.
- **Commits:** `7bd0c5e` (monthly zones, Pulse, freeze reset) · `96e5f46` (Gym 1 check-ins, Media Center 2x, event types) · `84c08f2` (staff overrides).
- **Tests:** `node tests/clash.test.js`, 46 passing. Local preview: `node tests/build-preview.js` → http://localhost:3041, admin password `preview`.
- **Nothing is live yet.** Matt pastes the Apps Script files later. The admin password is already changed (done by Matt).
- **Staff playbook (private page):** https://claude.ai/artifact/V4FnHhDu65DfLiW7jKiwTX
- **Full change notes:** `docs/2026-11-13-update.md`

## Decisions made
- **Map zones:** each zone is won on the **last calendar month**.

  | Zone | Rule | Wins | Pays |
  |---|---|---|---|
  | Gym 2 (old gym) | Pride Day participation % (Climb-er `pride_day`; until then, Pride Day rows) | highest | 2x |
  | Gym 1 (new gym) | Move · Make · Talk: confirmed squad event participants + Move/Make/Talk event points | highest | 1x |
  | Cafeteria (Summit) | LRRs + lunch detentions per 100 students | fewest | 1x (Summit 2x) |
  | Music Wing | Full Climb: total points | most | 1x |
  | Auditorium | Poll / Pulse votes + wheel spins | highest | 2x |
  | 50s Hall | Attendance rate | highest | 1x |
  | 60s Hall | Shout-outs | highest | 2x |
  | Media Center | Growth Grind: STAR growth + Mindset summits | most | 2x |
  | Front Office | Monthly staff squad challenge winner (Type "Staff squad challenge winner") | most | 1x |
  | 70s Hall | Failing grades per 100 | fewest | 1x |
  | 80s Hall | Trait Dough earned | most | 2x |
  | 90s Hall | Teacher + admin interventions per 100 | fewest | 1x |

- **Multiplier rule:** zones where students create the data by doing something pay 2x. Zones that school systems or staff record pay 1x. Matt picked the Media Center as the extra 2x zone, because Mindset summits are student reflections.
- **Gym 1 participation:** students check in with their 800#, and Generals confirm in Generals HQ. Score = confirmed participants + event points.
- **Pride Day:** Climb-er sends `pride_day` per squad.
- **Overrides:** staff must be able to edit and override every piece.

## What was built (Apps Script files in `apps-script/`)
- **18_Conquest.gs:**
  - methods `highest number` / `lowest number`, Window Days `last month`, a **Per 100** column
  - a **Multiplier** column on Map Zones, and the monthly rules and zones as above
  - the **Squad Monthly** tab: Month, Squad, Students, attendance, pride_day, lrr_ld, failing, interventions_teacher, interventions_admin, shout_outs. Climb-er fills it. A highest/lowest zone waits until all four squads have a number.
  - built-in measures: `poll_votes`, `event_participants`, `squad_event_points`, attendance from finalized months
  - Menu: *Use the monthly zone rules (keeps the map)*
  - **Overrides:** Map Zones › **Locked** (holder never changes automatically), and the **Zone Overrides** tab (Month | Zone | Squad | Score | Note | Set By | Set At; newest row wins, a blank clears it). Admin › Map has a Lock holder tick and "Override [month]'s numbers" boxes that show the real data; saving re-checks holders right away.
  - the How The Map Works guide tab is rewritten
- **26_CheckIns.gs (new):**
  - Sign-up Events gets a **Kind** column (Move / Make / Talk)
  - the Event Check-ins tab; students check in that day plus `CHECKIN_DAYS` (1)
  - a Generals HQ **Check-ins** tab (confirm / remove / add a forgotten member); only confirmed check-ins count
- **06_Polls.gs, Pulse:**
  - a new poll Kind `Pulse`: no winner, a point per voter, live results right after voting
  - a Through column (multi-day), a Grade on each vote, a by-grade Results breakdown
  - `clashPulseLive()` feeds the **Pulse page** and **TV slide 9**: Menti-style bars + a squad turnout race
  - the **Poll Ideas** tab: 40 school-useful questions; Menu › *Add the next Pulse question* (`PULSE_DAYS` 3)
- **07_Freezes.gs / 14_Thaw.gs:**
  - a `votes: N` thaw goal
  - the `FRESH_START` setting + Menu › *Fresh start: end every freeze on a date…*: earlier freezes end the day before, a Fun Freeze cashes out its tier, and it runs once
  - Menu › *Freeze & thaw student release (read aloud)*
- **Pop.html / Student.html:** a "How freezes work" panel on My Squad; check-in cards on home and My Squad; the Pulse UI
- **Generals.html:** the Check-ins tab; January meeting slides (Fresh start, Freeze & thaw in 60 seconds, Our freeze, Zones that pay 2x, This week's Pulse)
- **Index.html / App.html:**
  - the Pulse page and nav entry, TV slide 9
  - Add Event and **Bulk Entry** Type picker: Squad event Move / Make / Talk, Staff squad challenge winner, Pride Day. Bulk Entry used to save only `event`.
- **MapEditor.html:** Multiplier, Per 100, "last month", Lock holder, override boxes
- **01_Config.gs:** settings `FRESH_START`, `PULSE_DAYS`, `CHECKIN_DAYS`; new EVENT_LABELS; setup calls `ensureCheckinTabs_`

## Go-live steps (when Matt pastes)
1. Paste the Apps Script files, run `setupSpreadsheet`, deploy a new version.
2. Menu › Use the monthly zone rules (Yes = clean January start).
3. Sign-up Events: give each squad event a Date and a Kind (Move / Make / Talk).
4. Climb-er owner: fill the Squad Monthly columns, including `pride_day` (the Climb-er repo was not in this session).
5. Menu › Fresh start… with the first January school day.
6. Menu › Add the next Pulse question; make sure daily poll points are ON.
7. Read the student release the first week back. Generals run the January meeting from Generals HQ › Squad meeting.

## Open / not done
- PR #1 is not merged.
- Nothing has run in the real Apps Script yet. Everything was tested offline and in the local preview.
- Climb-er still needs its Squad Monthly change.
- The go-live steps in the PR body and the playbook still include a "set the admin password" line, which isn't needed (already done).
