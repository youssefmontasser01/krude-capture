# KRUDE Capture

Phone web app for the two KRUDE co-founders: Youssef (Co-founder · Product, admin) and Ahmad (Co-founder · Market).
Three tabs: **Capture** (voice/text/photo in), **History** (every captured item, search + filters), **Board** (live read of KRUDE Board v2 › Tasks; status is the only edit).

Always write the brand as **KRUDE** in files and UI. Never "Crude".
Write replies to Youssef in ASD-STE100 Simplified Technical English. Keep them short.

## Layout

```
web/index.html          the whole phone app (one file, no build). Cloudflare Pages serves /web
apps-script/Code.js     backend (Google Apps Script web app). Pushed with clasp
apps-script/appsscript.json
CLAUDE.md               this file
```

## How it works

- `index.html` POSTs JSON (`Content-Type: text/plain`, no preflight) to the Apps Script `/exec` URL in `const API_URL` (about line 150). Never change that URL unless Youssef gives a new one.
- Auth: a private link key (`#k=...`). The backend compares SHA-256 hashes in Script Properties `KEY_YOUSSEF_HASH`, `KEY_AHMAD_HASH`. New key: run `newKeyForYoussef` / `newKeyForAhmad` in the Apps Script editor.
- Actions: `identity`, `upload`, `review`, `confirm`, `list`, `board`, `set_status`.
- **Save** = `upload` only (fast). A 5-minute trigger `processPending` transcribes, writes the line in "18 - Capture Inbox", and sorts.
  **Save & Review** = `upload`, then `review` at once, then the user approves / keeps / rejects each item (`confirm`).
- One JSON state file per capture: `cap_<Owner>_<id>.json` in `KRUDE Capture Drop/_Capture State`. Media in a folder per capture in `KRUDE Capture Drop`.
- Approved items: `board_task` → append to Board › Tasks; `supplier` → "20 - Suppliers and Contacts"; everything else, and any update to an existing record → a handoff doc in the Handoffs folder (ChatGPT applies it). Items with money, prices, roles, strategy or operating changes need Youssef (`needs_youssef`).
- Phone safety: recordings are written to IndexedDB every second; the outbox keeps a capture until the server confirms it. Do not weaken this.

## Secrets (never in code, never in git)

Script Properties: `OPENAI_API_KEY`, `KEY_YOUSSEF_HASH`, `KEY_AHMAD_HASH`, `DROP_FOLDER_ID`, `STATE_FOLDER_ID`, `RULES_DOC_ID`, `SUPPLIERS_SHEET_ID`.
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
4. Frontend: `git add -A && git commit -m "<what changed>" && git push` → Cloudflare deploys in about 1 minute.
5. Backend (only if `apps-script/` changed):
   `cd apps-script && clasp push -f && clasp deploy -i <DEPLOYMENT_ID> -d "<what changed>"`
   Always use `-i <DEPLOYMENT_ID>` so the `/exec` URL stays the same. Never create a new deployment.
   DEPLOYMENT_ID: `<paste here from Deploy › Manage deployments>`
6. Tell Youssef in 2–4 short lines what changed and what to test on the phone. If the backend changed, say if a new function must be run once.

## Open items

- Real iPhone test of recovery (close Safari mid-recording) and offline send.
- Measure real OpenAI cost per capture.
- History reads the last 100 captures; add an index file if it gets slow.
- Ideas list for jam sessions (later).
- Ahmad's weekly 3-line report: moved / blocked / needs decision (later).
