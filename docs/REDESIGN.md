# Medical SRS — Deep System Redesign

_Analysis date: 2026-06-10 · against commit `3500abd` · 1310 lessons, S1–S10, 44 module/semester pairs, 7 live progress records in S8._

This document is the full redesign blueprint requested for the FMPM Medical SRS. It is written against the actual code (`scripts/build_srs_xlsx.py`, `srs-appscript.js`, `srs-system.xlsx`) — every weakness cites the line that causes it, and every recommendation states its rationale, complexity, and expected impact on learning outcomes. Nothing here breaks the current system; the target architecture is a strict superset of the current one (the two-cell daily workflow, the generator, and Excel compatibility all survive).

**TL;DR — the five moves that matter most**

1. **Add a `Today` tab** that answers "what do I study now?" in one glance — ranked due queue + a small "new lessons" budget. This is the single highest-impact change and is pure formulas.
2. **Add an append-only `Review Log`.** The current model stores only the *last* review per lesson; everything longitudinal (retention rate, streaks, weak modules, forecasts, achievements) is impossible without history. One hidden sheet + 5 lines of Apps Script unlocks the entire analytics layer.
3. **Fix the sync hot path**: per-row sync on edit (not a full 1310-row pass), true batched writes, and actually use the `SYNC_FUTURE_DAYS` window that exists but is dead code today.
4. **Kill "Mastery 5 = done forever."** Forgetting does not stop at 30 days; switch G=5 to a 60-day maintenance cycle. The formula *already computes* the 60-day interval — the status formula and the script just discard it.
5. **Make the sheet the single source of scheduling truth.** Move overdue-rescheduling to a hidden Override column the script writes, instead of faking `Last Review`. This deletes the duplicated `[1,3,7,14,30,60]` array from the script and retires the #1 maintenance landmine ("change the rule in BOTH files").

---

## 1. Complete project analysis

### 1.1 What the system is

Three artifacts in lockstep:

| Artifact | Role | Notes |
|---|---|---|
| `srs-system.xlsx` | Data + offline brain | 5 sheets; all logic in cross-platform formulas |
| `srs-appscript.js` | Online engine | Calendar sync, smart reschedule, digest, triggers |
| `scripts/build_srs_xlsx.py` | Generator | Rebuilds the workbook from `data/curriculum_FMPM_S1-S10.txt`, preserving progress |

Curriculum scale (parsed from the source of truth): **1310 lessons** — S1:105, S2:110, S3:162, S4:152, S5:142, S6:97, S7:159, S8:101, S9:154, S10:128 — across **44 module/semester pairs**. Largest module: *S7 Maladies de l'enfant* (61 lessons). 1129 lessons have a Subject, 181 don't. No duplicate `(Semester, Module, Lesson)` keys, so the progress-preservation key is currently collision-free.

### 1.2 What already works well (do not regress)

- **The two-cell contract.** You only ever type `F = Last Review` and `G = Mastery`. This is genuinely excellent UX and the redesign keeps it as the *offline floor* — everything new layers on top.
- **Graceful failure recovery.** Overdue items don't collapse into a panic pile; `smartReschedule` spreads them. The *intent* is right (the mechanism corrupts data — see W6).
- **Reproducible generation with progress preservation** (`build_srs_xlsx.py:77-89`). The "edit the generator, not the binary" discipline is the right architecture for a 1310-row workbook.
- **Cross-platform formula discipline** — `IF/TODAY/COUNTIF(S)/INDEX/MATCH` everywhere, with `FILTER` used exactly once and degrading gracefully (`build_srs_xlsx.py:319-320`).
- **Macro/micro split**: Dashboard (bird's-eye) vs Module View (one module) is the correct two-altitude design.
- **Calendar events color-coded by mastery, batch reads, idempotent event update/delete** — the sync loop's *read* side is well built.

### 1.3 Data flow today

```
curriculum.txt ──(generator)──► lesson-database (A–M)
user types F,G ──► H interval ──► I next review ──► J status / K priority
                                  │
   onEdit/hourly trigger ──► syncToCalendar() ──► Google Calendar events
                                  │                (L synced, M event id)
Dashboard / Module View / Weekly ◄┴── COUNTIFS over lesson-database
```

One state row per lesson. No history. The Apps Script re-derives the interval ladder independently (`srs-appscript.js:375`), which is why CONTEXT.md's #1 rule is "change the rule in BOTH files or neither."

---

## 2. Critical weaknesses

Ranked by how much they hold the system back. Each has an ID used throughout the rest of the document.

### W1 — Single-state data model: no review history *(severity: critical, structural)*
Each review **overwrites** `F`/`G`. The system cannot answer: How many reviews did I do this week? What's my recall success rate? Which module do I fail most? What's my streak? Which lessons are "leeches" (failed 3+ times)? Every analytics ambition in the brief (retention trends, weak areas, consistency metrics, achievements) is blocked by this one gap. It also means `smartReschedule`'s fake-date trick (W6) *destroys the only record that exists*.

### W2 — "What should I study now?" has no first-class answer *(severity: critical, UX)*
The daily entry point is a 1312-row table. To find today's work you must sort/filter by Status or Priority manually; the Daily Digest is a modal `ui.alert` (`srs-appscript.js:438`) — unclickable, unscrollable, useless beyond ~20 items, and it vanishes when dismissed. The brief's primary goal ("never think about what to study") is currently met only halfway: the *data* knows, but no *surface* shows it.

### W3 — New-material sequencing doesn't exist *(severity: high, learning)*
1310 lessons start as `⚪ NEW` with `Priority 999`. The SRS schedules *reviews* but nothing answers "which new lesson should I start next, and how many today?" Without a new-lessons-per-day budget, the user either over-commits (intake spike → review avalanche in 3–7 days) or stalls. Anki solved this with the new-cards/day setting for exactly this reason.

### W4 — `onEditTrigger` runs a full-table sync per keystroke *(severity: high, performance)*
`srs-appscript.js:483-494`: every edit to F or G sleeps 1.5 s then runs `syncToCalendar()` over **all 1310 rows** plus calendar API calls. A 20-review study session triggers 20 full passes — minutes of trigger runtime, sluggish editing, and needless quota burn (Apps Script: 90 min/day total trigger runtime).

### W5 — "Batch" writes aren't batched; sync window is dead code *(severity: high, performance/scale)*
- `applyBatchUpdates_` (`srs-appscript.js:284-295`) does two `setValue` calls **per row** — the comment says "Group updates for efficiency" but nothing is grouped. (Side nit: the comments say "Column K/L"; the code correctly writes L/M.)
- `CONFIG.SYNC_FUTURE_DAYS: 90` (`srs-appscript.js:49`) is **never referenced**. Every dated, non-mastered lesson gets a calendar event no matter how far out. At steady state with hundreds of active lessons, that's hundreds of far-future events re-checked every hour, and a first full sync that can blow the 6-minute execution limit.

### W6 — `smartReschedule` falsifies `Last Review` and duplicates the ladder *(severity: high, integrity)*
`srs-appscript.js:374-379` back-dates `F` so that `F + H` lands on the target day, re-deriving the interval from a private `[1,3,7,14,30,60]` copy. Consequences: (a) the only historical fact the sheet stores becomes a lie; (b) the interval rule lives in two places and **drifts silently** if either changes — CONTEXT.md documents this as the #1 landmine; (c) any future analytics built on `F` inherit corrupted data.

### W7 — Mastery 5 = "done forever" is scientifically wrong *(severity: high, learning)*
Forgetting curves don't stop at 30 days (Ebbinghaus; Bahrick's permastore work shows even well-learned material needs occasional retrieval for years). For medicine specifically, S1 anatomy must survive until the externat/résidanat. The irony: the H formula **already computes 60 days for G=5** (`build_srs_xlsx.py:145` — the final `ELSE 60` branch), but the J status formula marks `✅ DONE` at `G>=5` (line 147) and the script drops the calendar event (`isMastered`, `srs-appscript.js:117,151`). The system half-implements maintenance and then discards it.

### W8 — Fixed intervals create lockstep review waves *(severity: medium, learning/workload)*
Every lesson graded `3` on the same day returns on exactly the same day (+14). Study 25 lessons in a pre-exam week and a 25-review wall arrives together two weeks later. `MAX_REVIEWS_PER_DAY: 15` exists (`srs-appscript.js:52`) but is only enforced inside the *manual* reschedule — nothing prevents waves at scheduling time, and no interval fuzz spreads them.

### W9 — Filtering doesn't propagate *(severity: medium, UX)*
Row-2 AutoFilter only affects the database tab. The Dashboard, Weekly, and digest always show *global* numbers. You cannot ask the system "show me my world as S3+S4+Cardiologie" — the brief's combined-filter requirement has no mechanism at all.

### W10 — Priority ignores fragility *(severity: medium, learning)*
`K = I − TODAY()` (days until due). A lesson 5 days overdue at mastery 4 outranks one 3 days overdue at mastery 0, yet the latter is far closer to total forgetting (low stability → steeper decay). Urgency should combine lateness with mastery.

### W11 — Progress preservation keyed on mutable text *(severity: medium, maintainability)*
The generator re-matches progress by `(Semester|Module|Lesson)` string equality (`build_srs_xlsx.py:83-87`). Fix one typo in a lesson title in `data/` and that lesson's history silently evaporates on regenerate. Documented in CONTEXT.md §6, but documentation is not a fix.

### W12 — Module View is coupled to Dashboard layout *(severity: low, structural)*
The module dropdown validates against `Dashboard!$A$46:$A$89` (`build_srs_xlsx.py:287`). Reordering the Dashboard breaks Module View. Reference data belongs on a hidden lists sheet, not inside a presentation surface.

### W13 — Assorted smaller issues
- `LAST = 2000` is hard-coded (`build_srs_xlsx.py:34`); should be derived from `len(rows) + headroom` so a curriculum that grows past 2000 can never silently fall off the formulas.
- `Utilities.sleep(1500)` where `SpreadsheetApp.flush()` is the correct primitive (`srs-appscript.js:491`).
- Estimated time is a flat `reviews × 10 min` (`build_srs_xlsx.py:344`, `srs-appscript.js:435`) — fine as default, wrong for anatomy vs. a stats lesson.
- `deleteAllSRSEvents_` only scans 2025-01-01 → now+1yr (`srs-appscript.js:541-543`); events beyond that window survive.
- Events live on the **default** personal calendar; SRS noise mixes with real life, and "delete all SRS events" is a scan instead of a calendar drop.

---

## 3. High-impact improvements (the big bets)

Each improvement below is specified in §4–§8; this section is the argument for *why*.

| ID | Improvement | Fixes | Learning impact | Effort |
|---|---|---|---|---|
| R1 | **`Today` tab — the mission queue** | W2, W3 | Removes the daily decision entirely; the brief's primary goal | Low (formulas only) |
| R2 | **`Review Log` + daily rollup** | W1 | Enables retention rate, streaks, weak-module analytics, honest forecasting | Low (1 sheet + small script) |
| R3 | **Per-row sync + true batching + live sync window** | W4, W5 | None directly; makes the system feel instant and scale to 3000+ | Low–Med |
| R4 | **Override column; ladder lives only in the sheet** | W6 | Honest history; reschedules stop corrupting data; kills dual-maintenance | Low |
| R5 | **G=5 → 60-day maintenance cycle** | W7 | Long-term retention for cumulative medical knowledge | Trivial |
| R6 | **Interval fuzz + due-date load awareness** | W8 | Flat, predictable daily workload | Trivial–Low |
| R7 | **Focus scope (Settings) + `InScope` column** | W9 | Every surface (Dashboard/Today/forecast) obeys semester/module filters | Med |
| R8 | **Exam mode (interval cap + revision campaign)** | — | The highest-stakes weeks of med school get a dedicated, principled behavior | Med |
| R9 | **Dashboard v2: mission → forecast → risk → progress** | W2, W10 | 10-second situational awareness | Med |
| R10 | **Stable lesson UIDs** | W11 | Regeneration becomes rename-safe; log gets a foreign key | Low |

---

## 4. Recommended architecture

### 4.1 Principles

1. **The sheet is the single source of scheduling truth.** Formulas compute `H`/`I`; the script *reads* them and writes only its own columns (`L`, `M`, `O`). The `[1,3,7,14,30,60]` array disappears from JS. The old rule "keep the rule identical in two places" becomes "the rule exists in one place."
2. **Excel is the floor, Sheets is the ceiling.** Offline Excel keeps: full scheduling, Today queue, Dashboard, filters. Online adds: calendar, auto-logging, one-tap review, campaigns. Every online-only feature must degrade to a no-op, not an error — same pattern as the existing `FILTER` fallback.
3. **Visible surface stays minimal.** All new columns are hidden. The user-facing contract remains: *type in F and G, look at Today.*
4. **Append, don't overwrite** for anything historical.

### 4.2 Target sheet map

| Tab | Status | Purpose |
|---|---|---|
| **`Today`** | NEW — the landing tab | Mission KPIs, ranked due queue, new-lessons budget, streak |
| `lesson-database` | extended (cols N–P, hidden) | The engine, unchanged contract |
| `Dashboard` | redesigned (§5) | Command center: mission → forecast → risk → progress |
| `Module View` | kept; dropdown re-pointed to `_lists` | Micro view |
| `Settings` | NEW | Focus scope, exam dates, daily budgets, toggles |
| `Review Log` | NEW (hidden, online-fed) | Append-only history: timestamp, UID, grade, interval |
| `_lists` | NEW (hidden) | Module/semester lists, named ranges, quotes |
| `Weekly` | retired — folded into Dashboard forecast strip | (keep until Dashboard v2 ships) |
| `How-To` | kept, rewritten for v3 | One screen, updated |

### 4.3 Target column map (`lesson-database`)

Columns A–M **unchanged** (zero migration risk). Three hidden additions:

| Col | Name | Owner | Content |
|---|---|---|---|
| `N` | InScope | formula | `1` if the row matches the Focus scope in Settings (§8.2), else `0` |
| `O` | Override | script | Explicit next-review date set by reschedule/campaign; blank normally |
| `P` | UID | generator | Stable lesson ID (§4.5) |

`I` (Next Review) becomes: `=IF($O3<>"", $O3, IF($F3="","", $F3+$H3))` — the *only* formula change required by the new sync contract.

### 4.4 Target sync contract

```
USER     types F, G                       (or taps "Mark Reviewed" online)
SHEET    computes H (ladder+fuzz), I (F+H or Override), J, K, N
SCRIPT   onEdit(F|G)  → log grade to Review Log → sync THAT ROW's event
         hourly       → full reconcile within SYNC_FUTURE_DAYS window,
                        batched L/M writes, cursor-resume if near 6-min limit
         reschedule   → writes dates into O (never touches F), clears O on next real review
         campaign     → writes O for in-scope lessons until exam date
```

### 4.5 Stable UIDs (fixes W11)

The generator maintains `data/lesson-ids.json`: `{"S8|Maladies de l'appareil locomoteur|Fractures du genou": 932, ...}`. On each run: existing keys keep their ID, new keys get the next integer, removed keys stay in the file (tombstones). Progress and the Review Log key on UID, so **renaming a lesson in `data/` only requires updating one JSON key — or nothing at all if you add an alias entry**. Without the manifest the system behaves exactly as today, so this is a pure-upgrade path.

### 4.6 Future-proofing audit (3000+ lessons, years of history)

| Risk | Verdict | Mitigation |
|---|---|---|
| `COUNTIFS` dashboards over 3000 rows | Fine — linear, ~150 formula cells; both engines handle this easily | Derive `LAST` from data size in the generator (kills the 2000-row cliff) |
| Conditional formatting | Fine — range-level rules, not per-row | Keep it that way |
| Review Log growth (50/day × 3 yr ≈ 55k rows) | Sheets fine (10M-cell limit), but dashboard formulas over raw log get slow | **Daily rollup**: the hourly trigger maintains one row/day (date, reviews, successes, est. minutes) on `_lists`; streaks/trends read the rollup, O(days) not O(reviews) |
| Calendar volume | Today: unbounded (W5). Target: `SYNC_FUTURE_DAYS = 60` window → low hundreds of events | Implement the window; orphan-sweeper monthly trigger |
| Apps Script 6-min/exec, 90-min/day | First-sync and per-edit usage breach today (W4/W5) | Per-row onEdit + batched writes + cursor resume → hourly full sync of 3000 rows ≈ 30–60 s |
| Workbook size | 1310 rows × 13 cols ≈ trivial; charts static | Nothing needed |

---

## 5. Dashboard redesign

Design target: **complete situational awareness in under 10 seconds, top-left to bottom-right, no scrolling for the mission.** Layout (single screen, ~40 rows):

```
┌────────────────────────────────────────────────────────────────────┐
│ 🧠 MEDICAL SRS        🔥 Streak: 6 (best 14)      “quote of day”   │ r1
├──────────┬──────────┬──────────┬──────────┬──────────┬─────────────┤
│ 🔴 OVERDUE│ 🟢 TODAY │ 📖 NEW   │ ⏱ EST    │ 🎯 EXAM  │ ✅ DONE 12% │ r3-6
│    14     │    9     │  5 plan  │  3h10    │  J-23    │ ▓▓░░░░░░░░ │  (big numbers, conditional color)
├──────────┴──────────┴──────────┴──────────┴──────────┴─────────────┤
│ WORKLOAD FORECAST  ▁▂▅▃▂▆▇▂▁▃▂▄▅▂  (14-day column chart)            │ r8-15
│ Tomorrow: 11 · next 7d: 64 · next 30d: 188                          │
├─────────────────────────────────┬──────────────────────────────────┤
│ ⚠ AT RISK (top 8)               │ HEAT MAP  sem × status            │ r17-28
│ overdue>7d or due≤3d at G≤1,   │ 10 rows × {OVD·TOD·WEEK·NEW·DONE} │
│ ranked — INDEX/SMALL, clickable │ COUNTIFS + 3-color scale          │
├─────────────────────────────────┼──────────────────────────────────┤
│ RETENTION (needs log)           │ WEAKEST MODULES (top 5)           │ r30-38
│ 7d success 78% · 30d 81% ·     │ lowest avg G among started,       │
│ reviews/day trend chart         │ in scope — AVERAGEIFS             │
├─────────────────────────────────┴──────────────────────────────────┤
│ (below the fold: existing per-semester table, per-module table,    │ r40+
│  mastery mix — kept as-is, now ×InScope)                           │
└────────────────────────────────────────────────────────────────────┘
```

Key mechanics, all cross-platform:

- **Mission KPIs** are the existing COUNTIFS with an added `,'lesson-database'!$N$3:$N{LAST},1` criterion (scope-aware), formatted as large numbers with `FormulaRule` color flips (overdue > 15 → red fill).
- **Top-N lists without `FILTER`**: `INDEX($E:$E, MATCH(SMALL(IF(in-scope-and-due, $K… )…` is fragile cross-platform; instead use a helper rank column (hidden, in `_lists`) computed with `COUNTIFS`-based ranking, then `INDEX/MATCH` it. Where `FILTER`+`SORT` exist (Sheets/365), show the richer spill list — same graceful-degradation pattern already proven in Module View.
- **Forecast strip**: 14-row helper table `date = TODAY()+n`, `due = COUNTIFS(I,date,N,1)` (+ overdue added to day 0), column chart on top. The 7/30-day KPIs are `COUNTIFS(I,">="&TODAY(),I,"<"&TODAY()+7,N,1)` etc. This replaces the Weekly tab.
- **Heat map**: 10×5 `COUNTIFS(semester, status)` grid + per-column 3-color scale. Zero scripting, instant "where is the fire."
- **Retention block** reads the daily rollup (§4.6): `success rate = SUM(rollup.successes last 7) / SUM(rollup.reviews last 7)`; reviews/day sparkline-style column chart. Cells show `—` until the log has data (offline Excel users simply see `—`; nothing breaks).
- **Weakest modules**: `AVERAGEIFS(G, B,sem, C,mod, F,"<>")` per module, rank ascending, show bottom 5. Answers "which module needs revision first?" directly.

**Curriculum analytics views** (the brief's questions → exact mechanics):

| Question | Mechanic |
|---|---|
| Weakest semester? | `AVERAGEIFS(G per semester, started only)` — row added to the existing semester table |
| Weakest module? | bottom-5 table above |
| Lowest-retention subject? | same `AVERAGEIFS` at Subject granularity, on Module View (per selected module) |
| Most failures? | needs log: `COUNTIFS(log.grade,"<=1", log.module,m)` per module |
| Which semester to revise first? | composite: lowest avg G × % started × days-to-its-exam (Settings) — one ranked row |

---

## 6. Learning algorithm — review and recommendation

### 6.1 Verdicts on the questions asked

**Is 1/3/7/14/30 optimal for medical education?** It's a sound expanding ladder (≈×2.2 growth), consistent with the spacing-effect literature (Cepeda et al. 2006: optimal gap grows with desired retention interval). For *lesson-level* review (not flashcards), coarse steps are appropriate — the per-item precision of FSRS/SM-18 pays off for thousands of micro-cards/day, not ~20 lesson-level reviews/day. **Keep the ladder; fix its edges** (below). What's *not* optimal: no fuzz (W8), no credit for late-but-successful recall, "done" at 30 days (W7), and no exam awareness.

**Does mastery 5 = permanently done make sense?** No (W7). Medical curricula are cumulative; "done" should mean *cheap to maintain*, not *invisible*. Bahrick's 50-year permastore studies show long-lived knowledge needs spaced re-exposure; a 5-minute skim every 60 days is the cheapest insurance in the whole system.

**Should forgetting curves be modeled differently?** Not explicitly. A parametric memory model (e.g., FSRS's difficulty/stability/retrievability) needs per-review history and per-item optimization — wrong cost/benefit in a spreadsheet, and opaque to the user (see §11). The ladder *is* a piecewise approximation of expanding-interval retrieval; with fuzz + delay-credit + maintenance it captures ~90% of the benefit at ~5% of the complexity.

**Should difficult lessons behave differently?** Yes — and they already half-do: a failed lesson restarts at G=0 (1 day). Two cheap additions: (a) **leech detection** — ≥3 logged failures flags the lesson (`🩹` next to status) as "your SRS can't fix this; change the study method" (rewrite notes, make Anki cards, find a video); this is straight from Anki/SuperMemo practice and is where the real learning win lives; (b) an optional **★ high-yield flag** column that boosts queue rank (the owner's own old improvement note asked for exactly this).

**Adaptive intervals?** Lightweight only, online only: when "Mark Reviewed" sees a *successful* review that happened late, schedule from *actual elapsed time* rather than the ladder floor — `next = today + max(ladder[G], round(0.9 × days_elapsed))` written to Override. Rationale: a recall that survived 17 days has demonstrated ≥17 days of stability; rescheduling it at 7 wastes reviews (this is how SM-2 derivatives and FSRS treat delayed reviews). Formula floor stays the plain ladder, so offline behavior is unchanged and strictly safe.

**Exam mode behavior?** Yes — see §6.3. This matters more than any algorithmic refinement: med-school outcomes are decided in exam windows.

### 6.2 The recommended algorithm (v3)

User-visible semantics: **G stops being "where am I on an absolute 0–5 scale" and becomes "the rung of a ladder you climb."** After each review: recalled well → set `G = G+1`; shaky → leave `G`; forgot → `0`. (This is how the system was implicitly meant to be used; making it explicit in How-To + the data-validation prompt removes the single biggest daily judgment cost. Online, the Mark-Reviewed dialog applies the rule automatically — the user just answers *Forgot / Hard / Good*.)

| G | Meaning | Interval | Status |
|---|---|---|---|
| 0 | Forgot | 1 d | as today |
| 1 | Hard | 3 d | as today |
| 2 | OK | 7 d | as today |
| 3 | Good | 14 d | as today |
| 4 | Strong | 30 d | as today |
| 5 | Mastered | **60 d maintenance** (script may extend 60→120→240 via Override on repeated success) | `🟣 MAINTAIN`, **stays on calendar** |

Plus:

- **Fuzz** (kills W8): for `H ≥ 7`, add a deterministic ±10% jitter that needs no RNG and is identical in Excel/Sheets: `H' = H + IF(H>=14, MOD(ROW(),5)-2, IF(H>=7, MOD(ROW(),3)-1, 0))`. Same lesson always gets the same jitter (stable), but a cohort studied together spreads across ±2 days.
- **Delay-credit** (online, Override-mediated) as in §6.1.
- **New-lesson budget** (fixes W3): Settings cell `NEW_PER_DAY` (default 5); the Today tab lists the next N in-scope `⚪ NEW` lessons in curriculum order. The Dashboard shows *recommended* intake: `MAX(0, MIN(NEW_PER_DAY, (DAILY_MINUTES − due×10)/30))` — intake throttles itself when reviews pile up, which is exactly the regulation Anki users discover the hard way.

Implementation complexity: the whole of §6.2 is **one changed H formula, one changed J formula, one changed I formula, and ~40 lines of Apps Script** (Mark Reviewed + delay-credit). Impact: flatter workload, honest long-term retention, fewer wasted reviews, zero change to the two-cell offline contract.

### 6.3 Exam mode

Settings table: one row per semester → exam date (optional per-module dates later). Toggle: `EXAM_MODE = TRUE/FALSE` + the Focus scope defines *what* you're preparing.

Three coordinated behaviors:

1. **Interval cap (formula, works offline):** `H'' = MIN(H', MAX(1, exam_date − TODAY() − 2))` for in-scope rows when exam mode is on. Guarantees every active lesson is seen at least once more before the exam, with a 2-day buffer for final pass. One `IF` wrapped around the existing formula, driven by named Settings cells.
2. **Revision campaign (script, online):** "📚 SRS → 🎯 Generate Exam Campaign" distributes **all in-scope, non-new lessons** (including `G=5` — maintenance items are resurrected naturally since they're scheduled, not dropped) across the days remaining, weakest-first (`G` asc, then most-overdue), respecting `MAX_REVIEWS_PER_DAY`, writing dates to Override. Calendar follows automatically. After the exam (date passes), the script clears stale Overrides and normal SRS resumes — no manual cleanup.
3. **Dashboard shifts:** `🎯 J-23` countdown KPI; the At-Risk list re-weights to in-scope `G ≤ 2`; forecast strip shows campaign load so overload is visible *before* it happens.

This is the highest-leverage *medical-school-specific* feature in the redesign: it converts the SRS from a maintenance tool into a principled exam-prep engine without forking the data model.

---

## 7. Apps Script improvements

Ordered; each independent and shippable alone.

1. **Per-row sync on edit (fixes W4).** `onEditTrigger` gets the edited row from `e.range.getRow()`, calls a new `syncRow_(sheet, row)` that reads/writes only that row, and uses `SpreadsheetApp.flush()` instead of `Utilities.sleep(1500)`. The hourly `syncToCalendar()` remains the reconciler. Effort: ~30 lines. Effect: editing feels instant; trigger runtime drops ~99%.
2. **True batch writes (fixes W5a).** Read L:M for the whole data block once, mutate in memory, write back with **two `setValues` calls**. Replaces the per-cell loop in `applyBatchUpdates_`.
3. **Honor `SYNC_FUTURE_DAYS` (fixes W5b).** Skip event creation for `nextReview > today + 60` (drop the config from 90); delete events that fall out of the window (they'll be recreated when the window reaches them). Calendar stays at ~100–300 events regardless of curriculum size.
4. **Cursor-resume guard.** In the full sync loop: if `Date.now() − t0 > 4.5 min`, persist the row index to `PropertiesService.getScriptProperties()`, return; next hourly run resumes. Makes first-sync and 3000-lesson futures safe inside the 6-minute limit.
5. **Review Log appender (enables R2/W1).** On a G edit, append `[timestamp, UID, oldG, newG, daysLate]` to `Review Log`. ~8 lines inside the existing trigger. The hourly trigger maintains the daily rollup row (date, reviews, successes) — see §4.6.
6. **`markReviewed()` one-tap flow (fixes W2's input half).** Menu items *Reviewed: Good / Hard / Forgot* (and a grade prompt variant): writes `F = today`, `G` per §6.2 rule, appends log, syncs that row. Daily entry cost falls from *locate row + type two cells* to *select row + one menu tap*. (A later HtmlService sidebar listing today's queue with buttons is the deluxe version — see §10.)
7. **Override-based rescheduling (fixes W6).** `smartReschedule` writes target dates to `O` and never touches `F`. Delete the private `[1,3,7,14,30,60]` array — where the script needs an interval it reads the sheet-computed `H`. The dual-maintenance landmine is retired permanently.
8. **Exam campaign generator** (§6.3.2).
9. **Digest → surface, not popup.** `showDailyDigest` writes/activates the `Today` tab (and optionally `MailApp.sendEmail` a 7 a.m. digest — works on any phone with zero UI code; ~15 lines, separate daily trigger).
10. **Dedicated `📚 SRS` calendar.** `CalendarApp.getCalendarsByName('📚 SRS')[0] || CalendarApp.createCalendar('📚 SRS')` — keeps the personal calendar clean, gives one-toggle visibility on mobile, and makes "delete all SRS events" trivial and safe. Migration: run existing `deleteAllSRSEvents_` once, clear sync markers, resync.
11. **Orphan-event sweeper** (monthly trigger): delete prefix-matched events whose ID no row references. Cheap insurance against drift.
12. **Comment hygiene:** fix the off-by-one column comments in `applyBatchUpdates_`/`ensureEventIdColumn_` (`srs-appscript.js:274,288-289`) while touching the file.

Quota math after 1–4: hourly reconcile of 1310 rows ≈ 10–40 s (mostly skips), per-edit work ≈ 1 event call; first sync windowed to ≤ a few hundred events with resume. Comfortably inside 6 min/exec and 90 min/day for 3000+ lessons.

---

## 8. Spreadsheet improvements

1. **`Today` tab (R1)** — layout:
   - Row 1–4: mission KPIs (due, overdue, new planned, est. minutes, streak) — scope-aware COUNTIFS.
   - Rows 6–30: **review queue** — due+overdue, in scope, ranked by (overdue first, then `G` asc, then `K` asc). Sheets/365: one `SORT(FILTER(...))` spill. Older Excel: ranked `INDEX/MATCH` over a hidden helper-rank column (top 25 is plenty — nobody does more).
   - Below: **📖 To Learn today** — first `NEW_PER_DAY` in-scope NEW lessons in curriculum order.
   - This becomes the workbook's active tab on open (`wb.active` in the generator).
2. **`Settings` sheet (R7)** — Focus scope: a 10-row semester table with TRUE/FALSE include cells + a 44-row module table with include flags (pre-filled TRUE); plus `NEW_PER_DAY`, `DAILY_MINUTES`, `EXAM_MODE`, exam-date table, named ranges for each. Scope presets ("All", "Current semester only") as one-tap helper cells.
3. **`InScope` column N**: `=IF(AND(VLOOKUP($B3,sem_table,2,0), VLOOKUP($B3&"|"&$C3,mod_table,2,0)), 1, 0)` — every dashboard/Today/forecast formula adds the `N=1` criterion. This is the entire combined-filter architecture (S3+S4+Cardiologie = tick 3 boxes), works identically in Excel and Sheets, and costs one hidden column.
4. **Formula changes** from §6.2: `H` (fuzz + exam cap), `I` (Override-aware), `J` (`G=5` → `🟣 MAINTAIN`, scheduled statuses apply).
5. **`Review Log` + rollup block on `_lists`** (§4.6) and the **retention/streak/achievement formulas** that read the rollup:
   - Streak: count back from today over rollup rows with `reviews > 0`.
   - Achievements: a static 8-row table (`first 100 reviews`, `7-day streak`, `30-day streak`, `first module 100%`, `semester 100%`, `1000 reviews`, `100 maintained`, `exam campaign completed`) with one ✓ formula each. No popups, no points.
   - Daily quote: 30 quotes in `_lists`, `=INDEX(quotes, MOD(TODAY()-DATE(2026,1,1), 30)+1)` — one cell on Today/Dashboard.
6. **`_lists` sheet (fixes W12):** module/semester reference lists with named ranges; Module View's dropdown re-pointed here; Dashboard tables read from here too, so presentation reorders never break anything.
7. **Generator hygiene:** derive `LAST` from `len(rows)+200`; emit named ranges; write `lesson-ids.json` (§4.5); keep all fixed-row anchors *generated*, never hand-edited (unchanged philosophy).
8. **Heat map + forecast tables** per §5 — pure `COUNTIFS` + conditional color scales.

---

## 9. MVP roadmap (ship in this order; each step is independently shippable and reversible)

| Step | What | Files touched | Size |
|---|---|---|---|
| 1 | Algorithm edges: `G=5 → MAINTAIN` (J formula + remove `isMastered` drop), interval fuzz, derived `LAST` | generator, script | XS |
| 2 | Sync hot path: per-row onEdit, `flush()`, batched L/M writes, live `SYNC_FUTURE_DAYS=60`, cursor-resume | script | S |
| 3 | Override column `O` + `I` formula + reschedule rewrite; delete the JS ladder array | generator, script | S |
| 4 | `Settings` + `InScope` N + scope-aware Dashboard/Weekly KPIs | generator | M |
| 5 | **`Today` tab** (queue + to-learn + mission KPIs) — landing tab | generator | M |
| 6 | `Review Log` + onEdit appender + daily rollup + streak/retention/achievement cells | generator, script | S |
| 7 | Dashboard v2 layout (mission strip, 14-day forecast, heat map, at-risk, weakest modules); retire Weekly | generator | M |
| 8 | `markReviewed()` menu flow + digest-to-Today + UID manifest | script, generator | S |

After step 5 the brief's primary goal is met; after step 8 the daily loop is: *open sheet → Today tab → study → one menu tap per lesson*. Total: well under 30 seconds of overhead per day.

## 10. Advanced roadmap (post-MVP, in rough order of value)

1. **Exam mode** — Settings dates + cap formula + campaign generator + dashboard countdown (§6.3). Do this first when an exam approaches.
2. **Delay-credit scheduling** in `markReviewed()` (§6.1).
3. **Leech detection** from the log (`≥3` failures → 🩹 flag + "change method" list on Dashboard).
4. **High-yield ★ column** + queue boost (owner's old idea #2, still good).
5. **HtmlService "Today" sidebar** — due list with Forgot/Hard/Good buttons; the deluxe one-tap loop on desktop.
6. **Morning email digest** (15 lines; works on every phone, replaces the need for a mobile app).
7. **Dedicated SRS calendar** + orphan sweeper.
8. **Maintenance-interval growth** (60→120→240 on repeated G=5 success, via Override).
9. **Per-module minutes estimate** (Settings column; replaces flat 10 min in Est. time and Weekly load math).
10. **Mobile quick-entry Form** *(only if phone entry proves painful in practice)*: a Form whose dropdown is rebuilt daily by trigger with just *today's due lessons* (~15 items, not 1310), `onFormSubmit` updates the row. Bounded and viable — but try the Sheets mobile app + Today tab first.
11. **Ladder self-tuning report** (yearly): from the log, compute per-rung success rates; if rung 2→3 success is <80%, shorten rung 3 (e.g. 14→12). A *report that suggests*, not an algorithm that silently adapts.

## 11. Features that should NOT be implemented

| Tempting feature | Why not |
|---|---|
| **FSRS / SM-18 / neural scheduling** | Needs per-review history *and* per-item optimization; opaque to the user; the gain over a tuned ladder at ~20 lesson-level reviews/day is marginal. Spreadsheet formulas can't express it; you'd be maintaining a scheduler in Apps Script that the offline sheet can't reproduce — breaking the Excel floor. The log (R2) keeps the door open forever. |
| **Flashcards inside the sheet** | That's Anki's job (as `ai-report.txt` itself concluded). This system schedules *lessons*; pairing with Anki for micro-recall of the hardest items is the right division of labor. |
| **XP / levels / leaderboards / popup celebrations** | Solo user; gamification economies decay into noise and clicks. Streak + quiet achievements is the adherence-without-addiction sweet spot the brief asks for. |
| **Web app / Notion / Airtable / SQLite migration** | Destroys the Excel+Sheets dual mandate, the offline mode, and one-person maintainability for zero scheduling benefit at this scale. |
| **A second write-path for input (general-purpose mobile Form, API, etc.)** | Two write paths drift. (The bounded due-today Form in §10.10 is the only acceptable variant, and only if proven necessary.) |
| **Per-lesson manual metadata campaigns** (difficulty ratings, minute estimates for 1310 rows) | Data-entry debt nobody pays. Use per-module defaults and the ★ flag for the handful of items that matter. |
| **Multiple calendar events per lesson / event-per-overdue-day** | Calendar clutter; the digest/Today tab already carries that information. |
| **Auto-rescheduling that moves dates without being asked** | Trust is the system's currency; all schedule mutation stays behind explicit menu actions (reschedule, campaign). The *formulas* must stay deterministic. |

## 12. Prioritized list — impact × effort

Impact: learning/daily-use value (5 = transformative). Effort: XS < S < M < L. Ratio-ordered:

| # | Item | Impact | Effort | Why it wins |
|---|---|---|---|---|
| 1 | `Today` tab (R1) | 5 | S–M | The brief's primary goal, formulas only |
| 2 | G=5 → maintenance (R5) | 4 | XS | One formula + one script branch; protects years of retention |
| 3 | Per-row sync + batching + window (R3) | 4 | S | System feels instant; quotas safe to 3000+ |
| 4 | Review Log + rollup (R2) | 5 | S | Unlocks every metric the brief asks for; cost is one hidden sheet |
| 5 | Interval fuzz (R6) | 3 | XS | Two `MOD`s; permanently flatter workload |
| 6 | Override column + single-source ladder (R4) | 4 | S | Kills the #1 landmine *and* data corruption |
| 7 | Focus scope + InScope (R7) | 4 | M | The entire combined-filter requirement, one hidden column |
| 8 | Dashboard v2 (R9) | 4 | M | 10-second situational awareness |
| 9 | markReviewed one-tap (part of R1/R2) | 4 | S | Daily friction → near zero (online) |
| 10 | Stable UIDs (R10) | 3 | S | Rename-safe regeneration, log foreign key |
| 11 | Exam mode (R8) | 5 | M–L | Biggest med-school-specific payoff; do at first exam window |
| 12 | Streak/achievements/quote | 3 | S | Adherence, after the log exists |
| 13 | Leech detection | 3 | S | Targets the lessons SRS alone can't fix |
| 14 | Email digest | 2 | XS | Mobile reach for 15 lines of code |
| 15 | Heat map | 2 | XS | Cheap glanceability |
| 16 | Dedicated calendar + sweeper | 2 | S | Hygiene |
| 17 | Delay-credit | 2 | S | Fewer wasted reviews |
| 18 | High-yield ★ | 2 | S | Focus on what examiners ask |
| 19 | Sidebar UI | 3 | L | Deluxe; only after the menu flow proves insufficient |
| 20 | Ladder self-tuning report | 1 | M | Curiosity, yearly at most |

---

### Closing note on philosophy

The current system's greatest asset is its **restraint**: two input cells, plain formulas, one script. Every recommendation above was filtered through that lens — the redesign adds *surfaces* (Today, Settings, Log) and *fixes edges* (maintenance, fuzz, overrides), but the daily contract barely changes: open the sheet, look at Today, study, record. The student should spend their cognition on medicine, not on the machine that schedules it.
