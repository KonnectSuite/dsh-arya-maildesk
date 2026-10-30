# dsh-arya-maildesk — test case set

Two kinds of coverage: suites that run anywhere (no Outlook needed), and checks that need a real classic-Outlook
mailbox. This file says which is which, so a failure can be placed without guessing.

## A. Automated suites — `npm test` (178 assertions)

| Suite | Cases | Covers |
|---|---|---|
| `test/bundle.test.mjs` | 15 | the browser bundle is one hand-written module (`window.__ModuleLoader__`, no imports, no build step); route names and slot ids match between the halves |
| `test/panel-api.test.mjs` | 29 | panel routes through fake `req`/`res`: allowlist, paging and row fields, hostile HTML fixture, mark read, move, delete, confirmation hash (missing → `needs-confirmation`; wrong → refused; matching → sent once), draft reuse, 403/405/400 edges |
| `test/ask-route.test.mjs` | 29 | in-panel assistant: one session reused, panel context in the prompt, ground rules forbidding `ask_user_question`, `running` from turn markers, timestamps, `cancel`, `new`, transcript filtering of harness injections |
| `test/accounts.test.mjs` | 8 | hiding a store: listing, refusal to hide the last one, persistence, mailbox untouched |
| `test/models.test.mjs` | 10 | model catalog passthrough, selection, lazy session creation, 400/405 on bad input |
| `test/drafts-persist.test.mjs` | 11 | drafts survive a restart, reuse on the same thread, discard |
| `test/sample-mail.test.mjs` | 17 | fixture mailbox: labelled data, pill, a COM failure is an error rather than invented mail |
| `test/run-all.ps1` | 59 | the real PowerShell case bodies against a fake COM model: folder resolution, paging with a skipped non-mail item, sorting, preview truncation, unread-only, move returning new entryIds and isolating per-item failures, the `expected` count gate refusing **and deleting nothing**, mark read both ways, folder-tree BFS |

`node test/resident-live.mjs` is not part of `npm test`: it measures the resident bridge against a real mailbox
(cold start, warm `list_mail`, `account`) and is run by hand when the bridge changes.

## B. Live checks — a real Exchange profile (three stores, ~545 folders)

Performed 2026-10-30 on the owner's work laptop, counts restored after each destructive step:

| Check | Result |
|---|---|
| account / folders / list / read | pass |
| mark read → unread | pass (Inbox unread 141 → 140 → 141) |
| move to a folder, then back | pass (Archive 3 → 4 → 3) |
| delete to Deleted Items | pass (504 → 505 → 504) |
| compose → review → send confirmation shows the body | pass (walked `New → Send… → Back`; nothing sent) |
| reply-all with Exchange senders (`EX:/O=…`) | pass — To/Cc resolved to SMTP |
| one real send by the owner | pass (`RE: Asset Center`, present in Sent Items and Inbox) |
| conversation view | pass — four-message thread, three people, current message expanded |
| Focused / Other | pass — first page split 82 / 18 |
| hide an account, then show it | pass — rail and `list_folders` both changed |
| new-arrival count on the Mail row | seeded count rendered; real arrivals not yet observed |
| floating assistant answers a question | pass — timestamps, tool steps, Stop, Markdown links |
| model switch Flash ↔ V4-Pro | pass — chip re-labelled immediately |

## C. Manual or visual checks (no automation yet)

- Drag the assistant card, let it snap, resize from the grip, double-click to reset.
- Press **Stop** mid-answer.
- Keyboard: `Ctrl+J` toggles, `Esc` collapses, `Enter` sends, `Shift+Enter` breaks the line.
- A narrow window (< 900 px) turns the card into a bottom sheet.
- A send still shows the exact body in the confirmation card.

## D. Known boundaries (by design — do not file these)

1. **Focused/Other is AryaCode's rule, not Outlook's.** Outlook keeps that classification server-side; the item's
   property bag has no focus property and `Restrict` on the conversation id returns nothing. Automated senders are
   Other, people are Focused.
2. **`ask_user_question` is not used inside the panel.** Its popup belongs to the chat client; the panel prompt
   forbids the tool and the agent asks in plain text instead.
3. **Choosing a model in the panel changes the deployment default**, exactly as it does from chat.
4. **Windows + classic Outlook only.** The new Outlook for Windows has no COM surface; the panel says so and offers
   the sample mailbox instead of inventing mail.
5. **Search is Outlook's query engine**, stricter than a text search; narrow by folder and date first.

## E. Running it

```sh
npm test                 # everything in the table above
npm run test:api         # the route suites only
npm run test:bundle      # the browser-bundle invariants
```
