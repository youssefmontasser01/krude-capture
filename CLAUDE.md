# KRUDE Capture

Phone web app for the two KRUDE co-founders: Youssef (Co-founder · Product, admin) and Ahmad (Co-founder · Market).
Four tabs: **Capture** (voice/text/photo in), **History** (every captured item, search + filters), **Ideas** (my ideas + the shared jam list), **Board** (live read of KRUDE Board v2 › Tasks; status is the only edit).
Tap your name (top right) for the **weekly report** and, for Youssef, the **OpenAI cost** per month.

Always write the brand as **KRUDE** in files and UI. Never "Crude".
Write replies to Youssef in ASD-STE100 Simplified Technical English. Keep them short.

## Layout

```
web/index.html          the whole phone app (one file, no build). Cloudflare Workers serves /web
wrangler.jsonc          Cloudflare settings: publish /web only. Each push to main deploys
                        Live app: https://krude-capture.youssefshaker36.workers.dev (private link: /#k=<key>)
apps-script/Code.js     backend (Google Apps Script web app). Pushed with clasp
apps-script/appsscript.json
CLAUDE.md               this file
```

## How it works

- `index.html` POSTs JSON (`Content-Type: text/plain`, no preflight) to the Apps Script `/exec` URL in `const API_URL` (about line 180). Never change that URL unless Youssef gives a new one.
- Auth: a private link key (`#k=...`). The backend compares SHA-256 hashes in Script Properties `KEY_YOUSSEF_HASH`, `KEY_AHMAD_HASH`. New key: run `newKeyForYoussef` / `newKeyForAhmad` in the Apps Script editor.
- Actions: `identity`, `upload`, `review`, `confirm`, `list`, `board`, `set_status`, `ideas`, `jam`, `report_draft`, `report_send`, `reports`, `costs`.
- **Save** = `upload` only (fast). A 5-minute trigger `processPending` transcribes, writes the line in "18 - Capture Inbox", and sorts.
  **Save & Review** = `upload`, then `review` at once, then the user approves / keeps / rejects each item (`confirm`).
- One JSON state file per capture: `cap_<Owner>_<id>.json` in `KRUDE Capture Drop/_Capture State`. Media in a folder per capture in `KRUDE Capture Drop`.
- Approved items: `board_task` → append to Board › Tasks; `supplier` → "20 - Suppliers and Contacts"; everything else, and any update to an existing record → a handoff doc in the Handoffs folder (ChatGPT applies it). Items with money, prices, roles, strategy or operating changes need Youssef (`needs_youssef`).
- Index: `_index.json` in `_Capture State` holds the latest 400 captures (copy of the cap files). History, Ideas, the trigger and the report read only this file. Every non-quiet `saveCap_` updates it (inside the lock). If a write fails, `INDEX_STALE` is set and the next read rebuilds it. Manual fix: run `rebuildIndex`.
- Cost: each capture stores `cost` {usd, audio_secs, calls}. Month totals are in Script Property `COST_LEDGER`. Prices are estimates in `CFG.PRICES`; update them if OpenAI prices change.
- Ideas: items with kind `idea`. The owner (or Youssef) can put one on the shared jam list; either founder can mark it discussed. Ahmad sees Youssef's ideas only when they are on the jam list (summary and notes only).
- Weekly report: moved / blocked / needs decision. KRUDE weeks start on Sunday and are named by that date (same as the Board Week column), Cairo time. Due from Thursday until the Sunday 10:00 stand-up; Ahmad sees a banner until it is sent. The draft is filled from `_board_log.json` (status changes made in the app), Blocked tasks and his items that wait for Youssef. Stored in `_reports.json` and appended to "21 - Weekly Reports" (`REPORTS_DOC_ID`).
- Board rules in the app (from 07 / 00 - ChatGPT Project Instructions): a status change to Done writes the Done date (reopen clears it); Dropped counts as closed. New tasks from Capture get Product, Function (fixed list `FUNCTIONS`), Week (Sunday of the due date, else this week) and a "Check: no owner / due date" note when one is missing.
- Phone safety: recordings are written to IndexedDB every second; the outbox keeps a capture until the server confirms it. Do not weaken this.

## Secrets (never in code, never in git)

Script Properties: `OPENAI_API_KEY`, `KEY_YOUSSEF_HASH`, `KEY_AHMAD_HASH`, `DROP_FOLDER_ID`, `STATE_FOLDER_ID`, `RULES_DOC_ID`, `SUPPLIERS_SHEET_ID`, `REPORTS_DOC_ID`.
Written by the app: `COST_LEDGER`, `INDEX_STALE`, `FILE__index.json`, `FILE__reports.json`, `FILE__board_log.json`.
Optional overrides: `SORT_MODEL`, `TRANSCRIBE_MODEL`.

## Decisions (do not undo without asking Youssef)

- No type selection required before recording; type chips are optional hints.
- Photos are stored as links only. No image analysis (cost).
- Board tab: live read every time; only the status can change in the app. New tasks / owners / dates go through Capture.
- Board is the only task list for the business. Captured tasks are not Board tasks until approved.
- Ahmad has no access to the KRUDE Operations Drive folder. He sees his own captures in History and the Board through the app.
- Mobile first (iPhone Safari). Plain, warm-neutral KRUDE look. No frameworks, no build step.

## Workflow for every change

1. Edit the files.
2. Check syntax: `node --check` on the script part of `web/index.html` and on `apps-script/Code.js`.
3. For UI changes, test in a browser with a mocked API if you can (Playwright), at 390×844.
4. Frontend: `git add -A && git commit -m "<what changed>" && git push` → Cloudflare (Workers Builds, `npx wrangler deploy`) deploys in about 1 minute.
5. Backend (only if `apps-script/` changed):
   `cd apps-script && clasp push -f && clasp deploy -i <DEPLOYMENT_ID> -d "<what changed>"`
   Always use `-i <DEPLOYMENT_ID>` so the `/exec` URL stays the same. Never create a new deployment.
   DEPLOYMENT_ID: `AKfycbzkNMBJg5duf7TlBuT155_CryBWkt0g04K_ejtPEIKcv3-QR--JGwada_EEO_B69SSHgA`
6. Tell Youssef in 2–4 short lines what changed and what to test on the phone. If the backend changed, say if a new function must be run once.

## Open decisions (Youssef, at the Kickoff)

- D1: Ahmad sees all Board tasks in the app (decision 12 Nov: he gets his tasks after the stand-up). Keep / own tasks only / hide.
- D2: Status changes in the app write at once, without the Crude Today preview. Keep / make the Board tab read-only.
- D3: Ahmad can approve his own simple tasks straight to the Board. Keep / send to the Inbox for Sunday sorting.

## Open items

- Real iPhone test of recovery (close Safari mid-recording) and offline send.
- After v3 is deployed: run `setupV3` once in the Apps Script editor.
- Compare the cost estimate with the real OpenAI usage page after one month.
