# KRUDE Capture

Phone web app for the two KRUDE co-founders: Youssef (Co-founder · Product, admin) and Ahmad (Co-founder · Market).
Five tabs: **Capture** (voice/text/photo in), **History** (every captured item, search + filters), **Ideas** (my ideas + the shared jam list), **Board** (Tasks: live read of KRUDE Board v2 › Tasks; change status, edit or drop a task, always with a preview first. Week · Products: read-only views of the Week tab and the Pipeline tab), **Suppliers** (v3.9: read and write "20 - Suppliers and Contacts"; add, edit, drop, categories; preview first).
Tap your name (top right) for the **weekly report** and, for Youssef, the **OpenAI cost** per month. The **bell** shows **Recent changes** (Today / This week).

Always write the brand as **KRUDE** in files and UI. Never "Crude".
Write replies to Youssef in ASD-STE100 Simplified Technical English. Keep them short.

## Layout

```
web/index.html          the whole phone app (one file, no build). Cloudflare Workers serves /web
web/_headers            Cache-Control: no-cache (iPhone Home Screen apps must fetch new versions)
wrangler.jsonc          Cloudflare settings: publish /web only. Each push to main deploys
                        Live app: https://krude-capture.youssefshaker36.workers.dev (private link: /#k=<key>)
apps-script/Code.js     backend (Google Apps Script web app). Pushed with clasp
apps-script/appsscript.json
CLAUDE.md               this file
```

## How it works

- `index.html` POSTs JSON (`Content-Type: text/plain`, no preflight) to the Apps Script `/exec` URL in `const API_URL` (about line 180). Never change that URL unless Youssef gives a new one.
- Auth: a private link key (`#k=...`). The backend compares SHA-256 hashes in Script Properties `KEY_YOUSSEF_HASH`, `KEY_AHMAD_HASH`. New key: run `newKeyForYoussef` / `newKeyForAhmad` in the Apps Script editor.
- Actions: `identity`, `upload`, `review`, `confirm`, `list`, `board`, `set_status`, `ideas`, `jam`, `report_draft`, `report_send`, `reports`, `costs`, `task_change` (status / edit / drop; `dry:true` = preview only), `task_request` (Youssef approves or rejects Ahmad's requests), `activity` (Recent changes), `lists` (read-only Week / Products; Suppliers stays for old phones), `suppliers`, `supplier_change` (add / edit; `dry:true` = preview), `category` (add / rename; `dry:true` = preview). `set_status` stays for old phones.
- **Save** = `upload` only (fast). A 5-minute trigger `processPending` transcribes, writes the line in "18 - Capture Inbox", and sorts.
  **Save & Review** = `upload`, then `review` at once, then the user approves / keeps / rejects each item (`confirm`).
- One JSON state file per capture: `cap_<Owner>_<id>.json` in `KRUDE Capture Drop/_Capture State`. Media in a folder per capture in `KRUDE Capture Drop`.
- Approved items: `board_task` → append to Board › Tasks; `task_update` → change that Board task (v3.8); `supplier` → "20 - Suppliers and Contacts"; everything else, and any update to an existing record → a handoff doc in the Handoffs folder (ChatGPT applies it). Items with money, prices, roles, strategy or operating changes need Youssef (`needs_youssef`).
- Index: `_index.json` in `_Capture State` holds the latest 400 captures (copy of the cap files). History, Ideas, the trigger and the report read only this file. Every non-quiet `saveCap_` updates it (inside the lock). If a write fails, `INDEX_STALE` is set and the next read rebuilds it. Manual fix: run `rebuildIndex`.
- Cost: each capture stores `cost` {usd, audio_secs, calls}. Month totals are in Script Property `COST_LEDGER`. Prices are estimates in `CFG.PRICES`; update them if OpenAI prices change.
- Ideas: items with kind `idea`. The owner (or Youssef) can put one on the shared jam list; either founder can mark it discussed. Ahmad sees Youssef's ideas only when they are on the jam list (summary and notes only).
- Weekly report: moved / blocked / needs decision. KRUDE weeks start on Sunday and are named by that date (same as the Board Week column), Cairo time. Due from Thursday until the Sunday 10:00 stand-up; Ahmad sees a banner until it is sent. The draft is filled from `_board_log.json` (status changes made in the app), Blocked tasks and his items that wait for Youssef. Stored in `_reports.json` and appended to "21 - Weekly Reports" (`REPORTS_DOC_ID`).
- Board rules in the app (from 07 / 00 - ChatGPT Project Instructions): a status change to Done writes the Done date (reopen clears it); Dropped counts as closed. New tasks from Capture get Product, Function (fixed list `FUNCTIONS`), Week (Sunday of the due date, else this week) and a "Check: no owner / due date" note when one is missing.
- Task changes (Youssef, 14 Nov; v3.7): every change shows a preview, then writes, then reads the row back. Both founders change any shared Board task directly (`direct` is always true). The request flow (`_task_requests.json`, `task_request`) stays only so Youssef can close old open requests. Drop = status "Dropped" + reason in Notes. Archive and Delete forever: see below (v3.9). A status change also updates the product's Pipeline row: Last updated; a Next step / Blocker that names the task becomes "Check: set the next step" / "None" (old text kept in Pipeline Notes). The open tab reloads when the app comes back to the screen.
- Review (v3.3): no default choice; untouched items stay "To review". Each card has "Goes to" (Board task / Supplier / Stays in Inbox). Items that only stayed in the Inbox (kept, or approved as "Kept in Inbox") can be reviewed again from History ("Review again").
- Sorter (v3.3): anything someone must do = task / board_task, even without owner or date. "Needs Youssef" is decided by the app (`sensitive_`: kinds, destinations, sensitive_topics, money words in field VALUES), never by the sorter's own flag.
- Board view (v3.3): Overdue → Today → This week (to Saturday) → Next week → Later → No due date; filters: owner, status (Open / To do / In progress / Blocked), group by date / product / function, search. Choices are kept in localStorage.
- Recent changes (v3.4): `_activity.json` (last 1000 events, `emit_` inside the lock). Events: sorted items ("captured a task"), approvals ("added to the Board"), task changes, requests, jam list, weekly report, and changes made in the sheet (by hand or ChatGPT) found by comparing the Board with `_board_snap.json` on each Board read (`boardDiff_`; app writes patch the snapshot so they are not reported twice). Visibility: `vis` = 'all' or a list of names; Youssef sees all; Ahmad never sees Youssef's private captures. The phone shows a bell with a badge (changes by others since last opened).
- Review queue (v3.4): History shows "N items to review · Review all" = one screen with every item that waits for me, from all captures. Doc 18 is a log for app captures (they are sorted in the app, not on Sunday in the doc).
- Similar task (v3.4): in review, a Board task that shares key words with an open Board task shows "Similar task on the Board" and a "Same task: do not add" button (= Reject).
- Sorting order (v3.4): sort first, then write the Inbox line (faster). A failed Inbox line is retried by the 5-minute job (`inbox` flag in the index).
- Handoffs (v3.4): one open doc per day, "Handoff - YYYY-MM-DD - KRUDE Capture items", in the KRUDE Operations root (ChatGPT renames it DONE and moves it to Handoffs).
- Sorter context (v3.4): also gets the product names from Board › Pipeline. KRUDE facts and examples live in "19 - Capture Rules" (edited through ChatGPT handoffs).
- Strategy Board (v3.5, Decision Log 13 Nov): an approved Board task with Function "Leadership" from Youssef's capture goes to "KRUDE Strategy Board (Private)" (`CFG.STRATEGY_SHEET_ID`, Tasks tab, same columns). Ahmad's tasks always go to Board v2. Youssef's Board tab has a Shared / Strategy (private) switch (`board` with `which:'strategy'`, `task_change` with `board:'strategy'`; admin only). Strategy changes: no Pipeline update, no sheet-diff, activity visible to Youssef only. Never show the Strategy Board to Ahmad.
- Leadership (v3.7, Youssef 14 Nov): Ahmad never sees Leadership. `boardFor_` removes rows with Function Leadership for him; `planTask_` cannot find them; `fnsFor_` / `fns()` drop Leadership from his function lists and from the sorter schema and context; `validateItem_` clears it on his items; Board events and sheet-diff events for Leadership rows are `vis:['Youssef']`; `activity_` hides older events whose title is a Leadership task (flag `l` in `_board_snap.json`).
- Task changes from Capture (v3.8, Youssef 14 Nov): destination `task_update` = the capture changes an existing Board task. The sorter gets `open_tasks` (ref `B<row>` Board v2, `S<row>` Strategy for Youssef only; Ahmad: no Leadership) and fills `change` {ref, status, due_date, owner, title, notes_add}; the server stores the task title in `change.task`. In review the card shows "Which task" (founder can pick another) and the same dry-run preview as the Board tab. Approve = `planTask_` op edit (status and notes_add allowed) + `applyPlan_`: write, read-back, Pipeline update, Recent changes. Money words still make it wait for Youssef. Drop needs a reason in "Add to notes". No match = founder picks the task; nothing goes to a handoff.
- Archive and Delete forever (v3.9, Youssef 14 Nov; replaces "Board rows are never deleted"): `task_change` ops `archive` / `restore` / `delete` with `from` 'tasks' | 'archive' and `items` [{row,title}] (max 300); `archive_list` reads the Archive. Archive = the row moves from Tasks to an "Archive" tab in the same sheet (Board v2 or Strategy), with "Archived on" / "Archived by"; columns by header name. Both founders archive and restore. "Archive all N" on the Done or dropped list archives the closed tasks in the current view (`all:true`: the server checks each one is closed). Delete forever: Youssef only, from Tasks or Archive; the phone asks him to type DELETE. Each move: preview → copy → check the copy → delete rows from the bottom up → check. Snapshot is patched (no "removed in the sheet" event). Leadership rules stay (Ahmad never sees them; events vis Youssef). Suppliers: rows are still never deleted.
- Suppliers (v3.9, Youssef 14 Nov): "20 - Suppliers and Contacts" (`SUPPLIERS_SHEET_ID`) is the one source of truth. The 9 old columns stay first (Date … Photo link); v3.9 adds `Category | Pin | Price date | Source | Original message | Status | ID` at the end and a `Categories` tab. Columns by header name. Status Active / Historical / Dropped; rows are never deleted; a drop writes the reason and date in Notes. ID `sup_…` on every row (`supEnsure_` fills missing IDs, also for rows added by hand). The app never changes Source, Original message or ID. Price and Price date: Youssef only. Categories: add (no duplicates, compare without case and spaces) and rename (also every supplier with it); never delete; "No category" = empty. Every change: preview → write in the lock → read back by ID. The supplier tab = the first tab that is not "Categories" (`firstTab_`). Sorter: supplier items also carry `category` (only an exact match from the Categories tab) and `pin` (link); a place or pin with a shop name is a supplier with an empty phone. Run once: `setupV39`. One-time import of the WhatsApp backfill: `importBackfillSuppliersOnCopy` (test copy), then `importBackfillSuppliers` (guard: stops if a Source starts with "WhatsApp backfill").
- Phone safety: recordings are written to IndexedDB every second; the outbox keeps a capture until the server confirms it. Do not weaken this.

## Secrets (never in code, never in git)

Script Properties: `OPENAI_API_KEY`, `KEY_YOUSSEF_HASH`, `KEY_AHMAD_HASH`, `DROP_FOLDER_ID`, `STATE_FOLDER_ID`, `RULES_DOC_ID`, `SUPPLIERS_SHEET_ID`, `REPORTS_DOC_ID`.
Written by the app: `COST_LEDGER`, `INDEX_STALE`, `FILE__index.json`, `FILE__reports.json`, `FILE__board_log.json`, `FILE__task_requests.json`.
Optional overrides: `SORT_MODEL`, `TRANSCRIBE_MODEL`.

## Decisions (do not undo without asking Youssef)

- No type selection required before recording; type chips are optional hints.
- Photos are stored as links only. No image analysis (cost).
- Board tab: live read every time. Status, edit and drop in the app, with preview, read-back and Pipeline update (14 Nov; replaces "only the status can change"). Archive (both) and Delete forever (Youssef only) in the app (14 Nov, v3.9). New tasks still come through Capture.
- Board is the only task list for the business. Captured tasks are not Board tasks until approved.
- Leadership tasks live on the private Strategy Board (Youssef only); all other tasks on Board v2 (13 Nov).
- Ahmad has no access to the KRUDE Operations Drive folder. He sees his own captures in History and the Board through the app.
- Mobile first (iPhone Safari). Plain, warm-neutral KRUDE look. No frameworks, no build step.

## Workflow for every change

1. Edit the files. Raise `APP_VERSION` (web/index.html) and `VERSION` (Code.js) together. Youssef sees "App vX · Server vY" in the Me screen and a banner when the server is older than the app.
2. Check syntax: `node --check` on the script part of `web/index.html` and on `apps-script/Code.js`.
3. For UI changes, test in a browser with a mocked API if you can (Playwright), at 390×844.
4. Frontend: `git add -A && git commit -m "<what changed>" && git push` → Cloudflare (Workers Builds, `npx wrangler deploy`) deploys in about 1 minute.
5. Backend (only if `apps-script/` changed):
   `cd apps-script && clasp push -f && clasp deploy -i <DEPLOYMENT_ID> -d "<what changed>"`
   Always use `-i <DEPLOYMENT_ID>` so the `/exec` URL stays the same. Never create a new deployment.
   By hand (Youssef, no clasp): paste Code.js into Code.gs, save, then Deploy › Manage deployments › "capture v3.6" (the only active one) › pencil › New version › Deploy. Each deployment is pinned to one version: updating any other deployment does not change the app. Check: open the `/exec` URL; it must show the new `version`.
   DEPLOYMENT_ID: `AKfycbzkNMBJg5duf7TlBuT155_CryBWkt0g04K_ejtPEIKcv3-QR--JGwada_EEO_B69SSHgA`
6. Tell Youssef in 2–4 short lines what changed and what to test on the phone. If the backend changed, say if a new function must be run once.

## Decided 14 Nov (v3.7)

- D1: both founders see all shared Board tasks, notes included. Leadership is hidden from Ahmad everywhere.
- D2: changes with preview (see Task changes). Both founders change any shared task directly.
- D3: both founders approve their own simple captured tasks straight to the Board. Money, prices, roles and strategy still wait for Youssef.

## Open items

- v3.9 Suppliers tab: code on branch, not live yet. Youssef: merge to main, paste Code.js + New version of "capture v3.6", run `setupV39`, `importBackfillSuppliersOnCopy`, then `importBackfillSuppliers`. Then the phone tests (Step 5 of the v3.9 brief), and archive / restore / delete forever on a test task.
- Records: "Handoff - 2026-11-14 - Close of KRUDE Capture v3.8 session" (KRUDE Operations root) waits for ChatGPT: v3.8 rule in the Decision Log, State to "v3.8 live", two Capture Inbox lines (test + rollback). The v3.6 and v3.7 handoffs are DONE.
- v3.8 live 14 Nov (backend deployed by Youssef, app on main). First real test of a capture that changes a task: not reported yet.
- WhatsApp backfill: Youssef ran it in ChatGPT; he has the .xlsx and a reminder to review it (Capture Inbox 14 Nov).
- Ahmad's test feedback: not received yet.

- Real iPhone test of recovery (close Safari mid-recording) and offline send.
- Compare the cost estimate with the real OpenAI usage page after one month.
