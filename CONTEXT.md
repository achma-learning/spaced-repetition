# Spaced Repetition (Medical SRS) — AI Context File
_Last synced: 2026-06-08 @ d0fa071_

## 1. What This Is (Plain English)
- **In one sentence:** A study planner for medical school that tells you *which lessons to review today* and (when run online) drops those reviews straight onto your Google Calendar.
- **Why it exists:** Cramming 90 lessons before an exam is impossible to track by hand. You log *when* you last studied a topic and *how well you remembered it* (0–5); the system does the math on when you should see it again — and, crucially, it **recovers gracefully** when you miss days instead of dumping everything into one panic pile (`ai-report.txt`, `srs-appscript.js:6-9`).
- **Who uses it:** Just the owner — a medical student (curriculum is FMPM / Faculté de Médecine et de Pharmacie, French-language, in `data/`). Treat it as a personal tool: low blast radius, but the owner's real study schedule depends on it.
- **Vibe:** Polished-ish personal tool that's been iterated many times (see the big `spaced repetition old/` graveyard). Not a software product — it's two files that work together: a spreadsheet and a script.

## 2. How To Run It
There is **no build, no install, no npm**. It runs in two ways, on purpose:

### Online (full features — Google Sheets + Calendar sync)
1. Upload **`srs-system.xlsx`** to Google Drive and open it as a Google Sheet.
2. In the sheet: **Extensions → Apps Script**, delete the stub, paste all of **`srs-appscript.js`**, save.
3. Run `setupTriggers()` once and grant the Google permissions it asks for.
4. Done. A **📚 SRS v2** menu appears in the sheet; reviews auto-sync to your Google Calendar every hour and whenever you edit. (`srs-appscript.js:440-462`, `ai-report.txt:6`)

### Offline (tracker only — Microsoft Excel)
- Just **open `srs-system.xlsx` directly in Excel**. All the planning columns (interval, next review, status, priority) and the Dashboard/Weekly tabs recalculate using plain spreadsheet formulas.
- **Trade-off:** no Google Calendar sync and no menu — Apps Script is Google-only and does not run in Excel (`README.md:2`). Offline = see your schedule; online = schedule + calendar.

### Daily use (both modes)
You only ever type in **two columns**: `Last Review` (date you studied — Ctrl+; for today) and `Mastery` (0 = forgot … 5 = nailed it). Everything else fills itself in.

- **Required env vars:** None. Auth is Google account consent at Apps Script setup time — no `.env`, no keys in the repo.

## 3. Tech Stack
- **Google Apps Script (V8 JavaScript runtime)** — `srs-appscript.js`. Single file, global functions, no modules, no dependencies. Uses `SpreadsheetApp`, `CalendarApp`, `ScriptApp`, `Utilities`.
- **Spreadsheet workbook** — `srs-system.xlsx`, authored/last-saved by **LibreOffice Calc** (`workbook.xml` `appName="Calc"`). Logic lives in cell formulas (`IF`, `TODAY`, `COUNTIF`, `COUNTIFS`, `COUNTA`) — all of which exist in *both* Excel and Google Sheets, which is what makes offline mode work.
- **What kind of project:** Not a typical codebase — it's a **spreadsheet + a companion Apps Script**, version-controlled. No package manager, no lockfile, no CI, no LICENSE (all absent as of sync).
- **External services:** Google Sheets, Google Calendar (the account's *default* calendar, `srs-appscript.js:74`). Online mode only.

## 4. Code Map (The Important Files Only)
- `srs-appscript.js` — **Open this first.** The whole online engine: calendar sync (`syncToCalendar`), overdue spreader (`smartReschedule`), `showDailyDigest`, triggers (`setupTriggers`/`onEditTrigger`), and the sheet menu (`onOpen`). Top-of-file comment block is the authoritative column map.
- `srs-system.xlsx` — The data + offline brain. 4 sheets:
  - **`lesson-database`** — the engine. 90 pre-loaded lessons (rows 3–92). Header is 2 rows; data starts row 3.
  - **`Dashboard`** — live counts: overdue/today/this-week, per-module progress, mastery distribution, exam countdown.
  - **`Weekly`** — next-7-days workload forecast with a 🟢/🟡/🟠/🔴 load rating.
  - **`How-To`** — one-screen cheat sheet.
- `data/` — source curriculum lists (FMPM S1–S10) the 90 lessons were drawn from. Reference only; nothing reads these at runtime.
- `ai-report.txt` — the original design write-up from when v2 was built. Good "why" context; note it calls the workbook `medical_srs_v2.xlsx` (now renamed `srs-system.xlsx`).
- `backups/simple system/` — **READ-ONLY.** An older, simpler variant (`appscript_simple.js` + `final_updated_srs_simple.xlsx`). See §6 — its column layout is *different* from the current one.
- `spaced repetition old/` — archive of past versions, screenshots, and experiments. History, not active code.

### Current column map (`lesson-database`, data from row 3)
`A` # · `B` Module · `C` Subject · `D` Topic · `E` **Last Review (you type)** · `F` **Mastery 0–5 (you type)** · `G` Interval (formula) · `H` Next Review (formula) · `I` Status (formula) · `J` Priority (formula) · `K` Synced (script) · `L` Event ID (script-managed, hidden) — `srs-appscript.js:31-44`.

### The core SRS rule (lives in TWO places — keep them identical)
Mastery → days until next review: **0→1, 1→3, 2→7, 3→14, 4→30, 5→60**.
- In the sheet (col G): `IF(F=0,1,IF(F=1,3,IF(F=2,7,IF(F=3,14,IF(F=4,30,60)))))`
- In the script: the array `[1, 3, 7, 14, 30, 60]` (`srs-appscript.js:370`).
- Next Review `H = E + G`. Mastery **5 = DONE** → dropped from the calendar (`srs-appscript.js:114,149-153`).

### Column legends (so the emoji aren't a mystery later)
- **Status (I):** `⚪ NEW` (never reviewed) · `✅ DONE` (mastery 5) · `🔴 OVERDUE` · `🟢 TODAY` · `🔵 TOMORROW` · `📅 THIS WEEK` · `⏳ SCHEDULED`.
- **Priority (J):** days until the review is due (negative = overdue). Sentinels: `999` = no date / new, `9999` = mastered (sorts to the bottom).
- **Synced (K):** `✅` synced to calendar · `✅ done` mastered & removed · blank = needs (re)sync.

### Online behavior knobs (`CONFIG`, `srs-appscript.js:26-66`)
All-day calendar events, title prefix `📚 `, color by mastery (0 red → 5 grape), 8-hour-before popup reminder. Tunables: `MAX_REVIEWS_PER_DAY: 15`, `SPREAD_OVERDUE_DAYS: 3`, `SYNC_FUTURE_DAYS: 90`, `AUTO_SYNC_HOURS: 1`. Sheet menu (`onOpen`): Daily Digest · Sync Now · Smart Reschedule · Setup/Disable Auto-Sync · Clear Sync Markers · Delete All SRS Events.

## 5. Rules For Editing This Code
- **Change the interval logic in BOTH files or not at all.** The mastery→interval mapping is duplicated: the `xlsx` col-G formula and the `srs-appscript.js` array (`:370`). Edit one without the other and online vs offline silently disagree. This is the #1 rule.
- **Keep sheet formulas cross-compatible.** Offline mode depends on it: only use functions present in *both* Excel and Google Sheets (`IF`, `TODAY`, `COUNTIF`, `COUNTIFS`, `COUNTA`, arithmetic). No Apps-Script logic in cells.
- **Don't move or reorder columns** without updating `CONFIG.COL` in `srs-appscript.js:31-44`. The script reads cells by hardcoded position (A–L), not by header name.
- **The user only edits columns E and F.** Don't add manual-input columns that collide with the formula/script-managed ones (G–L).
- **`backups/` is read-only — never modify, overwrite, or "tidy" anything inside it.** (Explicit owner rule.)
- **Apps Script style:** one file, global functions, no `require`/`import`/npm. Functions ending in `_` (e.g. `deleteEvent_`, `ensureEventIdColumn_`) are private helpers and are intentionally hidden from the script's run menu.
- `HEADER_ROWS = 2`; data starts at row 3 (`srs-appscript.js:28`). Respect it everywhere.

## 6. Fragile Bits & Landmines
- **The "fake Last Review" trick in `smartReschedule`.** To land an item on a target day it back-dates `Last Review = targetDate − interval` so the `H = E + G` formula resolves to the day it wants (`srs-appscript.js:368-374`). It re-derives the interval from the same `[1,3,7,14,30,60]` array — so changing that array breaks rescheduling in a non-obvious way.
- **`onEditTrigger` has a deliberate `Utilities.sleep(1500)`** before syncing (`srs-appscript.js:486-488`). It's waiting for the sheet's formulas (H/I) to recalc after your edit. Remove it and the sync reads stale dates. It also *only* fires on edits to columns 5 (E) and 6 (F) — by design.
- **Hidden Event-ID column L is auto-created** (`ensureEventIdColumn_`, `srs-appscript.js:267-277`) the first time online sync runs. The offline `.xlsx` may not have it — that's fine, sync adds it. Don't "clean up" the empty/hidden column.
- **`parseDate_` rejects any year <2025 or >2100** as a junk-data guard (`srs-appscript.js:259`). A date outside that range is silently treated as "no date" (which deletes its calendar event). Bump the upper bound before this matters.
- **The simple backup uses a DIFFERENT column layout** — `A=Subject, B=Topic, C=Status, D=Mastery, E=Last Review, F=Interval, G=Next Review, H=Synced` (no Module/Priority/Event-ID). Do **not** copy its `CONFIG` into the current script (`backups/simple system/appscript_simple.js:26-37`).
- **Dashboard/Weekly ranges are hardcoded to rows 3:92** (the 90 lessons). Add lesson #91+ and it won't be counted until you widen those `COUNTIF`/`COUNTIFS` ranges.
- **Module names are hardcoded in Dashboard formulas** (`Traumato`, `Rhumato`, `MPR`, `Anapath`, `Génétique`, `Immuno`, `Syst`). Rename a module in column B and its dashboard count silently drops to 0.
- **The workbook is a LibreOffice Calc save.** Re-saving through different apps can mangle the emoji in formulas/labels and conditional formatting. Edit cautiously and eyeball the Status/Dashboard emojis afterward.

## 7. Current State
- **Last shipped:** the v2 system — graceful missed-day recovery, a daily workload cap (15/day), Smart Reschedule, color-coded calendar events by mastery, and the one-click Daily Digest (`srs-appscript.js:6-13`, `ai-report.txt`).
- **Working on now:** treating **`srs-appscript.js` + `srs-system.xlsx`** as the one canonical pair, and supporting **both** modes — online via Google Sheets (with calendar sync) and offline directly in Microsoft Excel.
- **Next up:** _Not yet figured out_ — likely candidates are extending the lesson set from `data/` past row 92 and keeping the dual-mode formulas in lockstep.

## 8. Update Protocol (Verbatim)
> **For the AI Assistant:** When asked to "Update CONTEXT.md":
> 1. Re-run Phase 0 — check for new `GEMINI.md` / `CLAUDE.md` / `.github/` files.
> 2. Re-scan the tree, manifests, and `.github/workflows/` for drift.
> 3. Read our recent conversation for new decisions, fragile bits discovered, or shifted goals.
> 4. Refresh the `_Last synced_` line with today's date and current commit SHA.
> 5. Rewrite — do not append. One clean source of truth. Preserve still-true content, revise the rest.
> 6. Keep §1 and §2 in plain English. Keep the file under ~350 lines.
>
> **Project-specific musts:** Always treat `srs-system.xlsx` + `srs-appscript.js` as the two files that matter. Never modify anything under `backups/` (read-only). Keep the system working **both** online (Google Sheets) and offline (Excel) — so any interval/formula change must stay mirrored between the `.xlsx` and the `.js`.
