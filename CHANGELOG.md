# Changelog

This project's version narrative puts **why** before **what**. Every entry starts with the motivation.

## 1.0.4 — The panel opens in seconds, not tens of seconds

**Why**: on a real Exchange mailbox the Mail panel sat on skeleton rows for ~24 s before the first message appeared. The first paint waited for `list_folders` with `maxDepth: 6`, and a folder walk costs one COM read per folder — 545 folders × 30–90 ms ≈ 16 s — on top of ~1.5 s of PowerShell process start and another 2.5 s before the first page of mail. Nothing was broken; everything was slow, which reads the same to a person who just clicked Mail.

**What changed**:
- `list_folders` answers in two new shapes: `roots: true` (the well-known folders of every store — 0.2 s of COM work, 1.4 s wall including process start) and `parent: <entryId>` (that folder's children — 2.4 s). The whole-tree walk is still there for the tools and the move picker.
- The panel paints the message list first, resolving `Inbox` by **name** inside the bridge, then asks for the rail. The rail is a lazy tree: a twisty on any folder with children fetches one level on demand.
- Folder responses are cached in the host for five minutes, keyed by their arguments, so reopening the panel and re-expanding a folder is free.
- The pager says `Loading…` while a page is in flight instead of `1–0 of 0 · unread count unavailable`, which read as an empty mailbox.
- `test/sample-mail.test.mjs` no longer fails on a machine that **has** classic Outlook: its "no COM" assertion adapts when the live provider answers, so the suite is honest on both kinds of machine.
- **An agent inside the Mail page.** `Ask` opens a drawer under the mailbox: the question goes to `/maildesk/api/ask`, the host runs a real harness session (the `sessionController` service, not a second model client) with the same `maildesk_*` tools and the same human gates, and the panel reads the transcript back and renders it beside the mail. The harness's own context injections are filtered out, so the drawer shows only what was asked and what came back.
- `test/ask-route.test.mjs` — 19 assertions against a stub controller: one session created and reused, the panel context travels with the question, tool calls are named while blank messages are dropped, and the transport edges stay closed.
- **The panel agent got memory and an audit trail.** Its session id is written to `%TEMP%\dsh-arya-maildesk-ask.json` and reused, so it is one `Mail assistant` conversation rather than a new one per restart (a stale id is retired and retried once). The transcript shows the tools it called, and the question carries the selected mail's `entryId`, so "reply to this one" resolves without a search.
- Cached folder counts are dropped whenever a write lands (`mark_read`, `move_mail`, `delete_mail`), so the rail can no longer argue with the list for five minutes.
- **`Move to…` became a picker**, because the lazy rail only knew the folders a person had expanded. It now opens a dialog with a path filter and an on-demand "Load every folder" that walks the whole mailbox once (cached host-side); typing `Materion` finds `\\kphung@frommconnect.com\Inbox\Materion`, `…\Materion Plant 1\Derick Montes` and the rest.
- **Unsent drafts survive a restart.** The draft map is mirrored to `%TEMP%\dsh-arya-maildesk-drafts.json` on every open, save, discard and send, and read back at mount, so reloading no longer eats a half-written reply. `test/drafts-persist.test.mjs` proves it with a second module instance, and that a discard clears the file instead of resurrecting the draft.
- **Search is in the panel**, not only in the tool: the box searches the current folder by subject and then by sender (Outlook's own engine), results take over the list and stay selectable, and `Clear search` puts the folder back. `search_mail` joined the panel's action allowlist for that.

**Measured** (work laptop, three stores, 545 folders): first mail **3 s** (was 24 s), rail **1.5 s**, one expand **2.5 s**, reopen **<0.1 s** from the host cache. In the Mail page, "how many unread mails are in my Inbox?" answered **141** from the tools, and "name the newest unread mail" returned sender, subject and date — both without leaving the panel.

**Verified on the real mailbox**: `mark_read` took the Inbox from 141 unread to 140 and back to 141 with the change confirmed by a fresh folder read; **`move_mail`** carried one mail to Archive and back (Archive 3 → 4 → 3); **`delete_mail`** sent it to Deleted Items (504 → 505) and a move restored it; and the panel's **Reply** opened a prefilled draft (To, `RE:` subject, quoted original) that was discarded without creating an Outlook item. Each step was read back from the folder counts, and the mailbox ended exactly where it started. Send remains unrun there.

**Not published yet** — this build went into the work laptop's profile as a local tarball (`pnpm add file:…`) because this machine has no npm credentials. The next publish takes it to 1.0.4 on npm and the market.

## 1.0.3 — A mailbox you can look at, on a machine with no Outlook

**Why**: installing the plugin on a PC whose Outlook is the new WebView2 client produced a correct panel and nothing to look at — a banner, disabled buttons, and an empty list. That is honest, but it is also the first thing a new user sees, and "nothing is working" is the only reasonable reading. The fixture already existed for exactly this (`MAILDESK_FIXTURE=1`), but an environment variable is not something a person installing from the market can set. The panel now offers the sample mailbox itself.

**What changed**:
- `POST /maildesk/api/fixture` switches the host's provider between live COM and the sample mailbox at runtime; `GET /maildesk/api/health` reports `sample: true` when the switch is what enabled it, so `MAILDESK_FIXTURE=1` and the button stay distinguishable.
- The no-Outlook banner gained **Preview with sample mail**, and the header shows a `SAMPLE MAIL` pill with **Back to live mail** while the sample mailbox is on. The label is deliberately different from `FIXTURE DATA`: one is a developer's environment variable, the other is a person looking at the interface.
- `test/sample-mail.test.mjs` (17 assertions) starts with no fixture, proves the live provider fails loudly instead of returning an empty mailbox, switches on, checks the folders/rows/read path through the same routes the panel uses, switches back, and covers the route's own guards (405/403/400).
- `npm test` now runs the bundle guard, the panel API suite, the sample-mail suite and the PowerShell suites in that order.

The tools share the provider, so an agent asking for mail sees the same sample mailbox the panel shows — and never a real one by accident.

---

## 1.0.2 — The browser bundle named itself by an old name too

**Why**: 1.0.1 fixed the bundle patch and still broke the app. Installing it on a machine with the packaged tarball produced, at web boot, `web boot: 1 entry did not activate @aryacode/dsh-arya-maildesk: import failed: client-modules: could not load "@aryacode/dsh-arya-maildesk"` — the host half registered all 24 tools and the window never appeared. Same mistake as 1.0.0, one layer down: `client.js` called `window.__ModuleLoader__.load({ id: 'dsh-arya-maildesk' })`, unscoped, while the module table keys every client bundle by **package name**. The module therefore registered under a name nothing imports. It passed every test and every development run for the same reason 1.0.0 did: the dev profile links this repository under the unscoped name, so the stale id resolved by luck, and the packed tarball is the first install that ever disagreed with it.

**What changed**:
- `client.js` registers under `id: '@aryacode/dsh-arya-maildesk'`, with a comment naming the failure it prevents.
- `test/bundle.test.mjs` grew to **15 checks**: the bundle's declared module id must equal `package.json`'s `name`. Replaying 1.0.1's `client.js` through it fails on exactly that line.
- Verified against the real thing: the dev profile was switched from the linked repository to the published tarball, which reproduced the failure verbatim, and 1.0.2 boots with the Mail row, the panel and the badge present.

The lesson is now written down twice in the same file: anything the loader reads **by name** — the patch row and the module id — needs an assertion, because neither is exercised by running the code.

---

## 1.0.1 — The bundle named a package that did not exist

**Why**: 1.0.0 installed cleanly on a work laptop and then refused to activate, with `outlook (dsh-arya-maildesk): failed to import: Error [ERR_MODULE_NOT_FOUND]: Cannot find package 'dsh-arya-maildesk'`. The bundle patch still named the package **unscoped** — a leftover of the rename to `@aryacode/dsh-arya-maildesk`. The loader resolves that string with `import()`, so the tarball could be present, correct, digest-verified and still not mount. It never showed up in development because the dev profile links this repository under the *unscoped* name, which made the stale string resolve by luck. Nothing about Outlook, COM or the market was wrong; the package simply called itself by an old name.

**What changed**:
- `cordis.patch.yml` names `@aryacode/dsh-arya-maildesk`, with a comment saying why the scope is not optional there.
- `test/bundle.test.mjs` (13 checks, now part of `npm test`) asserts the patch's row names match `package.json`, that the name stays scoped, that every entry point exists **and** is listed in `files`, and that `dsh.engines.dsh` is declared. Replaying 1.0.0's patch through it fails, which is the point.
- `npm test` runs the bundle guard before the PowerShell suites.

---

## 1.0.0 — Our own product, counted from one

**Why**: this began as someone else's plugin with a good idea and a hard ceiling — 18 tools, a badge, and no way to see more than one mail at a time. Almost everything you can touch now (the panel, the composer, the draft card, the Sidebar tab) was written in this repository, and the version number was still quoting a release line this project never published. A fork's number belongs to the fork. This is the point where the thing has its own name, its own tool prefix (`maildesk_*`), its own tests and its own face, so it starts counting at 1.0.0.

**What it contains**:
- **24 tools** over COM: the read/send/calendar set it inherited, plus `list_folders`, `list_mail`, `move_mail`, `delete_mail` and `mark_read` (the organize half), `compose_draft`/`compose_send` (harness-owned drafts) and `maildesk_compose` (the tool whose call renders as a card).
- **A mail panel** — folder rail, a real selection, reading pane, paging, unread filter, sort pushed to Outlook, and the organize actions.
- **A safe reading pane** — allowlisted HTML, remote images blocked, Proofpoint links unwrapped.
- **Harness-owned drafts** — one per thread, sent from the composer, the conversation card or the Sidebar tab, behind a content-hash confirmation.
- **The draft where the conversation is** — a keyed `tool.call.toolview` card and a right-Sidebar tab.
- **A fixture provider** (`MAILDESK_FIXTURE=1`) so the whole interface is reviewable on a machine with no Outlook, and **59 offline assertions** (`npm test`) that run the real PowerShell case bodies and the real host routes.
- English throughout, with the MIT notice and upstream provenance kept in `NOTICE` and `UPSTREAM.md`.

**What is not proven**: no release of this code has read a real mailbox yet. The COM path has only been exercised against synthetic mail; the first real machine is expected to be a work laptop with classic Outlook.

---

## Fork lineage (before 1.0.0)

This project started from `dsh-outlook` 2.11.0 (MIT). The entries below are that work under its old numbering, newest first, kept because they explain why each decision was made. They are history, not versions that will ever be published under this name.

### 2.17.1 — Naming the Outlook you actually have

**Why**: the first machine this was tried on had Outlook installed — the new one — and the bridge answered "install the classic Win32 Outlook app". That is the wrong sentence for someone who is looking at an Outlook window: it sends them to reinstall software they already have instead of telling them the two clients are different and only one has a COM surface.

**What changed**:
- The bridge's pre-check now distinguishes the two cases. If no `Outlook.Application` class is registered **and** the new Outlook's data folder exists (`%LOCALAPPDATA%\Microsoft\Olk`), it says so by name: the new Outlook is installed, it has no COM automation surface, this bridge cannot drive it, and the choice is a machine with classic Outlook or a non-COM backend.
- The old message stays for a machine with neither client.
- Verified on a host running `Microsoft.OutlookForWindows` 1.2026.1013.100 (signed into a work account and two personal accounts): `POST /maildesk/api/call {"action":"account"}` returns the new text, and the panel shows it under its own "No classic Outlook answered on this machine" explanation with every action disabled and `unread count unavailable` rather than `0`.

### 2.17.0 — The draft can sit beside the conversation

**Why**: the card in chat could show a draft, and `[Edit in panel]` took you to it, but "in panel" meant the chat disappeared. Reading the thread you are answering and writing the answer were still two different screens. The right Sidebar is the harness's own place for "something next to this conversation", so the draft belongs there too.

**What changed**:
- A right-Sidebar tab is two registrations, and the plugin now makes both: a static type via `ctx.sidebarRightTabs.register({ id: 'dsh-arya-maildesk:draft', kind: 'maildesk-draft', title })`, and the body at the keyed slot `sidebar.right.pane.tab` under that same id. The single `rightbar` / `rightbar.session` slots are left alone — they belong to the shipped Sidebar.
- `[Open beside chat]` on the draft card calls `ctx.sidebarRight.openTab('maildesk-draft', { params: { draftId } })`. The body reads the id back from `useTabInfo().tab.navigation.params`, falling back to `pendingRightDraft` because the tab body can mount after the click.
- The tab shows the same host-owned draft as the card and the panel: mode, subject, To/Cc, a 1200-character body excerpt, `[Send…] [Open in panel] [Discard]` and `Close this tab` through `info.tab.actions.close()`. Send keeps the two-step hash gate; a draft edited during the confirmation reports "The draft changed while you were confirming — send it again." rather than sending stale content.
- All three surfaces poll `op: "get"`, so there is still exactly one draft per thread and no second copy of it anywhere.
- `inject` is now `['slots', 'layout', 'sidebarRightTabs', 'sidebarRight']`.

### 2.16.0 — The draft lives in the conversation

**Why**: the draft was reachable only by leaving the chat. You asked the agent for a reply, it answered "draft created", and the thing itself was in another panel — so the conversation about a mail and the mail you were writing could not be on screen together. A tool row saying `{"status":"ok","draftId":"draft-…"}` is not a draft; it is a receipt for one.

**What changed**:
- `maildesk_compose` is a tool of its own (mode, entryId or draftId, to/cc/bcc, subject, body) and the twenty-fourth in the pack. It opens or reuses the host-owned draft and returns what the card needs: mode, addressing, body HTML, whether a quote is attached.
- The card in chat is a keyed `tool.call.toolview` entry for that tool name, not a chat-node kind. A chat node is keyed by the conversation model's own kind, so a plugin cannot add one without changing host data, and occupying `tool-call` would replace every tool row in the product. Registering the tool's own name renders our call and nothing else.
- `POST /maildesk/api/draft` gained `op: "get"`, and the card polls it: this harness has no verified host-to-page push channel, so the card settles on whatever the panel last stored instead of pretending to be live.
- `[Edit in panel]` calls `ctx.layout.selectPanel('mail')` and hands the draft id across a module-level bus — with the id parked in `pendingMailDraft` first, because the panel is unmounted while the chat is showing and a bare emit was silently dropped. The panel drains it on mount.
- `htmlToText` reverses `textToHtml` when a draft is reopened, so the composer shows the stored body instead of an empty box; without it the next save would have written the emptiness back.
- **Fixed**: the agent's tools bypassed the fixture. `maildesk_list_mail` went straight to the bridge and reported "classic desktop Outlook not found" on a machine where the panel was serving fixture mail happily. The fixture check now sits in `olk()` itself, so tools and panel answer from the same mailbox.
- `test/panel-api.test.mjs` is 29 assertions: the tool is registered, a reply draft's body is `<p>…</p>` paragraphs, a second call on the same thread returns the same id, and `op: "get"` 404s for an unknown draft.

### 2.15.0 — Replies that belong to the harness, not to Outlook

**Why**: the agent could draft, and the drafts were the problem. Every attempt left another item in Outlook's Drafts folder, so a thread asked about three times came back as three drafts, and the only way to see any of them was to leave the IDE and open Outlook. Contact lookup was poor enough that replying to the original mail was the reliable way to inherit its recipients. And a reply could not be written where the conversation about it was happening.

**What changed**:
- The draft is a harness object now, keyed `(account, mode, inReplyTo)`. Asking for the same reply twice reuses the same draft (`reused: true`) instead of creating a second, and the store is a `Map` in the host half — Outlook is contacted only when something is actually sent. The reuse is asserted, not assumed.
- Two bridge actions, `compose_draft` and `compose_send`. Both assign `HTMLBody`, so the quoted original keeps its shape; `draft_mail` assigns `.Body`, which flattens it to plain text. `compose_send` keeps the existing `OLK_CONFIRM` gate, captures the addressing *before* `Send()` clears it, and deletes the orphan item if `Send()` throws.
- `POST /maildesk/api/send` is two steps that cannot be confused: without a matching content hash it returns `needs-confirmation` with a preview and a hash over the exact arguments; only a retry carrying that hash sends. Editing the draft after the preview changes the hash.
- The reply is two dedicated same-origin routes (draft and send), separate from the allowlisted `/maildesk/api/call`, which still cannot send.
- Quotes go through `quoteSafe`: script, style, frames, forms and `on*=` handlers removed **with their contents**, because a quote is copied into an outbound message and the reading pane's allowlist is the wrong instrument there.
- The panel gained Reply, Reply all, Forward and New, a compose sheet (To/Cc/Bcc, subject, body, live quote preview) and a confirm card that lists the addressing before anything leaves.
- `read_mail` now also returns `senderSmtp`, with a note that on corporate Exchange it can be an `EX:/O=…` directory name, so callers must test it for `@` before offering it as a recipient.
- The fixture grew `compose_draft` and `compose_send`, stores the HTML body, and a sent reply reads back through `read_mail` exactly as it was sent — asserted in `test/panel-api.test.mjs`, which is now 26 assertions.

### 2.14.0 — Mail you can read, not just list

**Why**: the panel listed mail and rendered the body as one flat paragraph, so an invoice, a trip sheet and a vendor notice all looked alike, and the quote you were about to answer lost its shape before you saw it. The plain-text path also erased a distinction that matters: *this mail has no HTML* and *this mail has HTML we refuse to render* both came out as text. A Proofpoint-rewritten link could not be checked before clicking it, and a tracking pixel was never a decision the reader got to make.

**What changed**:
- `read_mail` accepts `format: 'html'` and returns `htmlBody` with `htmlTruncated` (cap 200,000 chars by default, `maxHtmlChars` to raise it). The bridge returns the HTML raw, and the comment in `lib/olk.ps1` says why: the renderer owns the document parser, so that is where the decision belongs.
- The reading pane sanitizes before rendering: allowlisted tags, attributes and style properties; `<script>`, `<style>`, frames, forms, objects and SVG dropped **with their contents**; every `<img>` replaced by a *remote image blocked* chip so no pixel is fetched; links limited to `http:`, `https:`, `mailto:` and `tel:` and stamped `rel="noopener noreferrer" target="_blank"`.
- Proofpoint unwrapping: both the `urldefense.proofpoint.com/v2/url?u=…` parameter and the `/v3/__…__;` segment form resolve to the real target before the allowlist test, so the link you inspect is the link you visit.
- The fixture's `read_mail` now serves deliberately hostile HTML — a table, a wrapped vendor link, a `javascript:` link, a tracking pixel, an inline `<script>` and an `onclick` — so the sanitizer is exercised through the same route the panel uses, not only in a unit test.
- `test/panel-api.test.mjs` gained the assertion that `htmlBody` is absent unless html is asked for, that it contains `<table>` when it is, and that the raw bridge body still contains `<script` — the bridge must not quietly rewrite someone's mail.

### 2.13.0 — Named dsh-arya-maildesk, and the plugin gets a face

**Why**: two reasons, one of them embarrassing.

`dsh-outlook` is taken on npm, and every tool carried an `outlook_*` prefix, so installing this release beside the upstream one put two plugins on the same 23 tool names and let whichever registered second win silently. A fork that cannot be installed next to its parent, and cannot be addressed by a distinct name, is not a fork yet.

The second reason is the point of the fork. Every complaint the agent produced — no delete, no move, a reply that only exists inside Outlook, having to ask twice for the same thing — has the same shape: the plugin could *read* a mailbox and nothing else, so the human did the organizing in another application. No tool surface fixes that. It takes a panel: folders, a list you can select in, a reading pane, and a composer that keeps the draft in the harness instead of stacking another item in Outlook's Drafts folder.

**What changed**:
- Identity: package and plugin `dsh-arya-maildesk`; all 23 tools `maildesk_*`; slot registrations, the storage id and the style element carry the new name; the watcher log is `%TEMP%\dsh-arya-maildesk-watch-debug.log`. Upstream keeps its own name inside `NOTICE`, `UPSTREAM.md` and `LICENSE` — that is provenance, not identity.
- `repository`, `bugs` and `homepage` dropped from `package.json` until this fork has its own remote: they pointed at upstream, which would send bug reports to people who did not write this code.
- **A mailbox panel**, not a bigger badge: folder rail, list with a real selection (click, shift-range, ctrl-toggle, arrows), reading pane, and mark read/unread, move and delete on the selection. The badge stays a doorbell for new arrivals.
- One allowlisted host route, `POST /maildesk/api/call`, is the panel's only outlet. `send_draft`, `reply_mail` and every calendar write are absent from the allowlist, so nothing the panel can do leaves the machine.
- `MAILDESK_FIXTURE=1` serves a synthetic mailbox so the interface is buildable and reviewable on a PC with no classic Outlook. Opt-in only: a COM failure is still reported as an error, never answered with fake mail.
- `test/panel-api.test.mjs` drives the real route against the fixture provider: 16 assertions covering paging, sorting, unread filtering, the read-state and move/delete contracts, the count gate on delete, the send-path refusal, the cross-origin refusal and a malformed body.

### 2.12.0 — Mail management, and the plugin speaks English (first fork increment)

*(This repository is a fork of `dsh-outlook`, base `5a2ec7d` / v2.11.0 — see `NOTICE` and `UPSTREAM.md`.)*

**Why**: two problems, one release.

Upstream's 18 tools do "read" exhaustively and "change" not at all — no delete, no move or archive, no read-state, no folder tree, no paging. So the agent could look and search inside a mailbox but never tidy it up, and tidying is the part that actually costs time every day. The concrete blocker is the GUI: `search_mail` caps `limit` at 50 and has no offset, so a list view that scrolls cannot be built on it. The first brick of the panel was missing exactly here.

And the plugin talked to its user in Chinese: the badge, the popover, the toasts, and — worse — the text of the confirmation popups that decide whether a mail goes out. On a desk where nobody reads it that way, every confirmation was a wall of text the reader had to trust instead of read.

**What changed**:
- Bridge (`lib/olk.ps1`): 5 new actions, **all additive** — not one character of the existing 18 branches was touched: `list_folders` (BFS over every store's folder tree, returning entryId/path/depth/child count/unread; `counts:false` skips the two count reads per folder, which is one COM round trip each on a large mailbox), `list_mail` (folder + sort + offset/limit ≤ 200 + optional unread-only + optional 160-char preview, returning total/offset/returned — browse, not search), `move_mail` (`Item.Move()`, **returning the new entryId**, because a move changes the ID and any caller holding the old one holds nothing), `delete_mail` (soft delete only, `Item.Delete()` into Deleted Items, `recoverable:true`), `mark_read` (`UnRead` + `Save()`, reversible).
- New `Resolve-Folder`: a folder may be a path (`Inbox/Project/2026`, `/` and `\` both accepted), an English well-known name (Inbox / Sent Items / Drafts / Deleted Items), or an entryId. **CJK aliases were deliberately dropped everywhere** — folder names, the GAL match, the room-list match — so the source stays pure ASCII and no non-UTF-8 transport can quietly turn a literal into one that never matches. Consequence: a mailbox whose folders are named in another language must be addressed by entryId. `search_mail` still carries its own inline copy of the old resolver; folding it into `Resolve-Folder` is a `TODO(fork)` once the new actions have run against a real Outlook.
- Batch cap of 500 per call, each item in its own try/catch, failures collected into `failed[]` instead of failing the batch — dragging 40 mails must not die because the 17th is a meeting invite.
- `delete_mail` carries an `expected` count gate: the caller must state "I believe there are N here", and a mismatch **refuses outright** and asks for a re-read, instead of deleting a selection it never saw clearly. Deletion stays soft; nothing in the plugin purges Deleted Items.
- Host: 5 new tools — `maildesk_list_folders`, `maildesk_list_mail`, `maildesk_move_mail`, `maildesk_delete_mail`, `maildesk_mark_read` (23 in total).
- **No new confirmation gate, on purpose.** The two-call confirmation, the content hash and the `OLK_CONFIRM` token exist for actions whose output *leaves the machine* (`send_draft`/`reply_mail`/`calendar_create` with attendees/`respond_meeting`/`meeting_update`/`meeting_cancel`). Deleting to Deleted Items and moving between folders do not leave it, and a gate there would only teach the user to approve without reading. The real hazard is covered by the count gate instead.
- Full English sweep: every user-facing string in `index.js` (confirmation previews, tool descriptions, the chat link labels) and every string in `client.js` (badge, aria labels, popover, toasts) is now English; the CJK folder/GAL/room aliases are gone from `lib/olk.ps1`; `README.md`, `TESTCASES.md` and this file are rewritten in English. The two `assets/*.png` badge screenshots still show the old Chinese tooltip and need a re-shoot.

- Bridge bugs the offline harness caught, both of which would have failed on the first real mailbox. `Resolve-Folder` had been written **inside** `Out-Json`, so it existed only while that function ran: `list_mail` and `move_mail` with a `destFolder` would have died with "term not recognized", and no parse check can see this — only an extent check can. And any folder path of 60 characters or more was assumed to be an entryId and rejected, so `Inbox/Projects/2026/Q4/...` could never resolve; the id attempt now falls through to the path walk, and the error names the segment that is missing.
- `list_mail` no longer fails when a store does not expose `UnReadItemCount` (public folders, some archive add-ins): the field reports `-1`, meaning the store did not tell us, instead of aborting the whole listing.
- `test/` — an offline harness that needs no Outlook at all: `folder-resolution.test.ps1` (extent/nesting checks plus 8 resolution cases, including a 78-character path and a 71-character entryId containing `/`) and `mail-actions.test.ps1`, which runs the **real** case bodies out of `lib/olk.ps1` against a fake COM model (59 assertions: paging, sort, unread-only, previews, move returning the new ids, the delete count gate refusing, both directions of read state, the folder walk). `npm test` runs both in child processes. This is what makes the COM layer provable on a machine with no Outlook installed.

### 2.11.0 — Chat "open" links no longer leave dead tabs

**Why**: the `[Open](…/olk-notify/open?entryId=…)` markdown link the agent posts in a conversation is a GET navigation — every click left behind a "opened in Outlook / this page can be closed" tab. The most common action ("take a look at this mail") was the one that littered the browser. The same link also hard-coded `127.0.0.1:3080`, so on a custom-port deployment the click failed outright (README listed it as a known boundary).

**What changed**:
- The client half gained capture-phase click interception: on a plain left click of an anchor whose pathname is `/olk-notify/open` carrying an `entryId`, **matched by pathname rather than by full URL** (so any port hits), the click becomes a same-origin POST — the exact channel the sidebar popover already used — and Outlook opens the mail with **no new tab**, plus a 2.6s toast (subject on success, error plus beacon on failure). This also fixed the custom-port dead link: once intercepted, it always targets the page's own origin.
- **The GET navigation fallback was removed entirely** (a decision, not an oversight): `/olk-notify/open` is now a POST-only JSON route. The result page ("this page can be closed") and its close button are gone — the normal path no longer creates a tab, so there was no reason to maintain a second shape. A modifier-click or middle-click on such a link now gets 405, deliberately. The GET 500 on an invalid entryId disappeared with the branch.
- Both the interception listener and the toast are registered inside a fiber effect, so disabling or updating the plugin removes them (listener, timer and DOM node).
- Pending-queue cleanup and the badge decrement are unchanged: the sidebar and the chat link now go through exactly the same POST handler.

### 2.10.1 — Portability hardening before the public release

**Why**: the plugin ships to arbitrary machines. It cannot assume its install directory is writable, that Outlook was ever installed, or that the README is current.

**What changed**:
- Watcher diagnostics moved from the plugin install directory (which may be read-only for a global install) to `%TEMP%\dsh-outlook-watch-debug.log`.
- The bridge prechecks `HKCR` before creating any COM object: with no classic Outlook installed (or only the UWP "New Outlook") it returns an actionable message instead of an opaque ActiveX error.
- Bridge and watcher are invoked via `Get-Content -Raw -Encoding UTF8`, so they no longer depend on a file BOM (the "an editor strips the BOM" trap has now reproduced twice).
- README refreshed to the actual 18 tools / 43 cases, the four meeting tools and `maildesk_open_mail` documented, and the known-boundaries section updated (port assumptions, the `OLK_WATCH` switch, log location).
- `meeting_update`/`meeting_cancel` descriptions aligned with what the code does (uniform two-call confirmation).

### 2.10.0 — Click a mail in the conversation to open it, and a bridge encoding hazard fixed

**Why**: search and new-mail results returned only an entryId, so a user who wanted to read the original had to go find it in Outlook — the most common action ("take a look at this mail") was the most roundabout. Testing also exposed something worse: `olk.ps1`/`watch.ps1` were BOM-less UTF-8, and PowerShell 5.1's `Get-Content -Raw` decodes with the system code page (GBK) — every Chinese literal came out garbled and the whole file failed to parse. Production had only survived because the system code page happened not to be GBK.

**What changed**:
- Bridge action `open_mail`: `GetItemFromID().Display()` opens the mail in a desktop Outlook inspector and explicitly sets `UnRead=$false` (opening counts as reading; `wasUnread` is returned). Purely local, never leaves the machine, no confirmation gate.
- Host route `/olk-notify/open`: GET (`?entryId=`, returning a minimal HTML result page the chat markdown link could navigate to) plus POST (JSON, for the sidebar popover). Both keep the same-origin guard (a browser navigation sends no Origin header, so it is trusted at the same level as the existing routes) and an entryId allowlist regex (`[A-Za-z0-9+/=_-]{20,512}`). On success the host drops that entryId from `notify.pending` and decrements `unread`, so the badge follows.
- New tool `maildesk_open_mail` (18 in total); the `maildesk_search_mail`/`maildesk_check_new` descriptions gained a hint to attach a clickable open link to each mail.
- Every mail in the sidebar popover became a clickable button: POST to the open route, removed from the list on success (badge decrements immediately, not at the next 20s poll), beacon on failure.
- `olk.ps1`/`watch.ps1` got a UTF-8 BOM, killing the PS 5.1 ANSI misread at the root.

### 2.6.1 — The badge stays visible when the sidebar is collapsed

**Why**: in collapsed mode the badge protruded past the button's edge (`top:-4px`) and was either clipped by the sidebar's collapse-animation container or covered by the row above — so the user could not see new mail, and the notification feature failed exactly in the quiet rail form where it mattered most.

**What changed**: in collapsed mode a 14px dot is drawn inside the 36px circle's top-right corner; the expanded outside badge is unchanged; the badge got `z-index:1`.

### 2.6.0 — From DOM injection to an official slot

**Why**: the old version injected its button straight into the shell's footArea container and defended itself with two MutationObservers. That bought positional freedom at the price of three real incidents: before v2.5.1 an observer microtask storm froze the entire page ("Loading plugins…"); substring matching against hashed CSS class names broke silently whenever the shell was upgraded; and hover feedback never rendered at all because inline styles outranked the stylesheet.

**What changed**: the whole button now registers through the `sidebar.footer.action` slot (the same official channel dsh-dock's Whale Bay menu uses), and React owns its lifecycle — observer, self-healing and class-name probing are all deleted (about 150 lines of defensive code gone). `wide`/rail geometry comes from the slot props; the popover became a React Portal (click-outside still closes it). `order:10` keeps the original visual position (above Whale Bay).

### 2.5.3 — Hover feedback actually renders now

**Why**: v2.5.2 fixed the hover colour variable (`--dsh-` → `--dsw-`) and the user still saw no background — an inline `background:'transparent'` always outranks a stylesheet's `:hover` rule.

**What changed**: the transparent background moved out of the inline style and into the injected stylesheet so `:hover` can override it.

### 2.5.2 — Aligned with the native settings button, item by item

**Why**: hover used the wrong token prefix (`--dsh-`, a variable that does not exist), and the collapsed margin did not match the native settings button (4px vs 8px/10px), which is visibly crooked side by side.

**What changed**: hover token aligned to `--dsw-alias-interactive-bg-hover`; rail margin aligned to `8px 0 10px`; the remaining geometry parameters were checked one by one and were already correct.

### 2.6.0 batch — 14 fixes from a deep review

Deterministic bugs and hardening found by a security review (all covered by cases in TESTCASES.md):

- **Content-hash binding** (hardening): the content hash on a `confirmed:true` retry must match the last thing shown to the user, or confirmation is forced again — "show A, send B" is blocked at the host layer (both attack scenarios verified as blocked).
- `search_mail`'s `unread` field had inverted semantics (`-not $it.UnRead` → `[bool]$it.UnRead`).
- An unknown action returned `status:'ok'` together with an error; it now throws.
- Attachment saving: `GetFileName` against path traversal, plus automatic `(n)` numbering instead of overwriting.
- COM calls serialized (Outlook is single-threaded; concurrent spawns trampled each other).
- `client-log` route capped at 64KB (413); the debug log rotates at 1MB.
- The watcher's `received` now uses the mail's real `ReceivedTime`.
- A failed `ResolveAll()` lists the unresolved attendees and aborts; room lookup prefers a unique SMTP resolution.
- `account` matches by the default Inbox's DeliveryStore.
- DASL LIKE strips `%`/`_` wildcards; a failed send cleans up the orphan draft; the popover got proper ARIA semantics (dialog).

### 2.5.1 and earlier

(No changelog was kept for these versions; the quirks and the field lessons are encoded in the `lib/olk.ps1` header comment.)
