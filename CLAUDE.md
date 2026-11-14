# KRUDE Capture

Phone web app for the two KRUDE co-founders: Youssef (Co-founder · Product, admin) and Ahmad (Co-founder · Market).
Four tabs: **Capture** (voice/text/photo in), **History** (every captured item, search + filters), **Ideas** (my ideas + the shared jam list), **Board** (Tasks: live read of KRUDE Board v2 › Tasks; change status, edit or drop a task, always with a preview first. Week · Products · Suppliers: read-only views of the Week tab, the Pipeline tab and "20 - Suppliers and Contacts").
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
- Actions: `identity`, `upload`, `review`, `confirm`, `list`, `board`, `set_status`, `ideas`, `jam`, `report_draft`, `report_send`, `reports`, `costs`, `task_change` (status / edit / drop; `dry:true` = preview only), `task_request` (Youssef approves or rejects Ahmad's requests), `activity` (Recent changes), `lists` (read-only Week / Products / Suppliers). `set_status` stays for old phones.
- **Save** = `upload` only (fast). A 5-minute trigger `processPending` transcribes, writes the line in "18 - Capture Inbox", and sorts.
  **Save & Review** = `upload`, then `review` at once, then the user approves / keeps / rejects each item (`confirm`).
- One JSON state file per capture: `cap_<Owner>_<id>.json` in `KRUDE Capture Drop/_Capture State`. Media in a folder per capture in `KRUDE Capture Drop`.
- Approved items: `board_task` → append to Board › Tasks; `supplier` → "20 - Suppliers and Contacts"; everything else, and any update to an existing record → a handoff doc in the Handoffs folder (ChatGPT applies it). Items with money, prices, roles, strategy or operating changes need Youssef (`needs_youssef`).
- Index: `_index.json` in `_Capture State` holds the latest 400 captures (copy of the cap files). History, Ideas, the trigger and the report read only this file. Every non-quiet `saveCap_` updates it (inside the lock). If a write fails, `INDEX_STALE` is set and the next read rebuilds it. Manual fix: run `rebuildIndex`.
- Cost: each capture stores `cost` {usd, audio_secs, calls}. Month totals are in Script Property `COST_LEDGER`. Prices are estimates in `CFG.PRICES`; update them if OpenAI prices change.
- Ideas: items with kind `idea`. The owner (or Youssef) can put one on the shared jam list; either founder can mark it discussed. Ahmad sees Youssef's ideas only when they are on the jam list (summary and notes only).
- Weekly report: moved / blocked / needs decision. KRUDE weeks start on Sunday and are named by that date (same as the Board Week column), Cairo time. Due from Thursday until the Sunday 10:00 stand-up; Ahmad sees a banner until it is sent. The draft is filled from `_board_log.json` (status changes made in the app), Blocked tasks and his items that wait for Youssef. Stored in `_reports.json` and appended to "21 - Weekly Reports" (`REPORTS_DOC_ID`).
- Board rules in the app (from 07 / 00 - ChatGPT Project Instructions): a status change to Done writes the Done date (reopen clears it); Dropped counts as closed. New tasks from Capture get Product, Function (fixed list `FUNCTIONS`), Week (Sunday of the due date, else this week) and a "Check: no owner / due date" note when one is missing.
- Task changes (Youssef, 14 Nov): every change shows a preview, then writes, then reads the row back. Youssef changes any task. Ahmad changes his own tasks; a change to Youssef's task, or giving his task away, becomes a request in `_task_requests.json` that waits for Youssef (badge on the Board tab). Drop = status "Dropped" + reason in Notes; rows are never deleted. A status change also updates the product's Pipeline row: Last updated; a Next step / Blocker that names the task becomes "Check: set the next step" / "None" (old text kept in Pipeline Notes). The open tab reloads when the app comes back to the screen.
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
- Phone safety: recordings are written to IndexedDB every second; the outbox keeps a capture until the server confirms it. Do not weaken this.

## Secrets (never in code, never in git)

Script Properties: `OPENAI_API_KEY`, `KEY_YOUSSEF_HASH`, `KEY_AHMAD_HASH`, `DROP_FOLDER_ID`, `STATE_FOLDER_ID`, `RULES_DOC_ID`, `SUPPLIERS_SHEET_ID`, `REPORTS_DOC_ID`.
Written by the app: `COST_LEDGER`, `INDEX_STALE`, `FILE__index.json`, `FILE__reports.json`, `FILE__board_log.json`, `FILE__task_requests.json`.
Optional overrides: `SORT_MODEL`, `TRANSCRIBE_MODEL`.

## Decisions (do not undo without asking Youssef)

- No type selection required before recording; type chips are optional hints.
- Photos are stored as links only. No image analysis (cost).
- Board tab: live read every time. Status, edit and drop in the app, with preview, read-back and Pipeline update (14 Nov; replaces "only the status can change"). New tasks still come through Capture.
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
   DEPLOYMENT_ID: `AKfycbzkNMBJg5duf7TlBuT155_CryBWkt0g04K_ejtPEIKcv3-QR--JGwada_EEO_B69SSHgA`
6. Tell Youssef in 2–4 short lines what changed and what to test on the phone. If the backend changed, say if a new function must be run once.

## Open decisions (Youssef, at the Kickoff)

- D1: Ahmad sees all Board tasks in the app, including notes (decision 12 Nov: he gets his tasks after the stand-up). He can change only his own. Keep / show his own tasks only.
- D2: decided 14 Nov (changes with preview, see above).
- D3: Ahmad can approve his own simple tasks from Capture straight to the Board. Keep / send to the Inbox for Sunday sorting.

## Open items

- Real iPhone test of recovery (close Safari mid-recording) and offline send.
- Compare the cost estimate with the real OpenAI usage page after one month.
