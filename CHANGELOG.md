# Changelog

This project's version narrative puts **why** before **what**. Every entry starts with the motivation.

## 2.14.0 — Mail you can read, not just list

**Why**: the panel listed mail and rendered the body as one flat paragraph, so an invoice, a trip sheet and a vendor notice all looked alike, and the quote you were about to answer lost its shape before you saw it. The plain-text path also erased a distinction that matters: *this mail has no HTML* and *this mail has HTML we refuse to render* both came out as text. A Proofpoint-rewritten link could not be checked before clicking it, and a tracking pixel was never a decision the reader got to make.

**What changed**:
- `read_mail` accepts `format: 'html'` and returns `htmlBody` with `htmlTruncated` (cap 200,000 chars by default, `maxHtmlChars` to raise it). The bridge returns the HTML raw, and the comment in `lib/olk.ps1` says why: the renderer owns the document parser, so that is where the decision belongs.
- The reading pane sanitizes before rendering: allowlisted tags, attributes and style properties; `<script>`, `<style>`, frames, forms, objects and SVG dropped **with their contents**; every `<img>` replaced by a *remote image blocked* chip so no pixel is fetched; links limited to `http:`, `https:`, `mailto:` and `tel:` and stamped `rel="noopener noreferrer" target="_blank"`.
- Proofpoint unwrapping: both the `urldefense.proofpoint.com/v2/url?u=…` parameter and the `/v3/__…__;` segment form resolve to the real target before the allowlist test, so the link you inspect is the link you visit.
- The fixture's `read_mail` now serves deliberately hostile HTML — a table, a wrapped vendor link, a `javascript:` link, a tracking pixel, an inline `<script>` and an `onclick` — so the sanitizer is exercised through the same route the panel uses, not only in a unit test.
- `test/panel-api.test.mjs` gained the assertion that `htmlBody` is absent unless html is asked for, that it contains `<table>` when it is, and that the raw bridge body still contains `<script` — the bridge must not quietly rewrite someone's mail.

## 2.13.0 — Named dsh-arya-maildesk, and the plugin gets a face

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

## 2.12.0 — Mail management, and the plugin speaks English (first fork increment)

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

## 2.11.0 — Chat "open" links no longer leave dead tabs

**Why**: the `[Open](…/olk-notify/open?entryId=…)` markdown link the agent posts in a conversation is a GET navigation — every click left behind a "opened in Outlook / this page can be closed" tab. The most common action ("take a look at this mail") was the one that littered the browser. The same link also hard-coded `127.0.0.1:3080`, so on a custom-port deployment the click failed outright (README listed it as a known boundary).

**What changed**:
- The client half gained capture-phase click interception: on a plain left click of an anchor whose pathname is `/olk-notify/open` carrying an `entryId`, **matched by pathname rather than by full URL** (so any port hits), the click becomes a same-origin POST — the exact channel the sidebar popover already used — and Outlook opens the mail with **no new tab**, plus a 2.6s toast (subject on success, error plus beacon on failure). This also fixed the custom-port dead link: once intercepted, it always targets the page's own origin.
- **The GET navigation fallback was removed entirely** (a decision, not an oversight): `/olk-notify/open` is now a POST-only JSON route. The result page ("this page can be closed") and its close button are gone — the normal path no longer creates a tab, so there was no reason to maintain a second shape. A modifier-click or middle-click on such a link now gets 405, deliberately. The GET 500 on an invalid entryId disappeared with the branch.
- Both the interception listener and the toast are registered inside a fiber effect, so disabling or updating the plugin removes them (listener, timer and DOM node).
- Pending-queue cleanup and the badge decrement are unchanged: the sidebar and the chat link now go through exactly the same POST handler.

## 2.10.1 — Portability hardening before the public release

**Why**: the plugin ships to arbitrary machines. It cannot assume its install directory is writable, that Outlook was ever installed, or that the README is current.

**What changed**:
- Watcher diagnostics moved from the plugin install directory (which may be read-only for a global install) to `%TEMP%\dsh-outlook-watch-debug.log`.
- The bridge prechecks `HKCR` before creating any COM object: with no classic Outlook installed (or only the UWP "New Outlook") it returns an actionable message instead of an opaque ActiveX error.
- Bridge and watcher are invoked via `Get-Content -Raw -Encoding UTF8`, so they no longer depend on a file BOM (the "an editor strips the BOM" trap has now reproduced twice).
- README refreshed to the actual 18 tools / 43 cases, the four meeting tools and `maildesk_open_mail` documented, and the known-boundaries section updated (port assumptions, the `OLK_WATCH` switch, log location).
- `meeting_update`/`meeting_cancel` descriptions aligned with what the code does (uniform two-call confirmation).

## 2.10.0 — Click a mail in the conversation to open it, and a bridge encoding hazard fixed

**Why**: search and new-mail results returned only an entryId, so a user who wanted to read the original had to go find it in Outlook — the most common action ("take a look at this mail") was the most roundabout. Testing also exposed something worse: `olk.ps1`/`watch.ps1` were BOM-less UTF-8, and PowerShell 5.1's `Get-Content -Raw` decodes with the system code page (GBK) — every Chinese literal came out garbled and the whole file failed to parse. Production had only survived because the system code page happened not to be GBK.

**What changed**:
- Bridge action `open_mail`: `GetItemFromID().Display()` opens the mail in a desktop Outlook inspector and explicitly sets `UnRead=$false` (opening counts as reading; `wasUnread` is returned). Purely local, never leaves the machine, no confirmation gate.
- Host route `/olk-notify/open`: GET (`?entryId=`, returning a minimal HTML result page the chat markdown link could navigate to) plus POST (JSON, for the sidebar popover). Both keep the same-origin guard (a browser navigation sends no Origin header, so it is trusted at the same level as the existing routes) and an entryId allowlist regex (`[A-Za-z0-9+/=_-]{20,512}`). On success the host drops that entryId from `notify.pending` and decrements `unread`, so the badge follows.
- New tool `maildesk_open_mail` (18 in total); the `maildesk_search_mail`/`maildesk_check_new` descriptions gained a hint to attach a clickable open link to each mail.
- Every mail in the sidebar popover became a clickable button: POST to the open route, removed from the list on success (badge decrements immediately, not at the next 20s poll), beacon on failure.
- `olk.ps1`/`watch.ps1` got a UTF-8 BOM, killing the PS 5.1 ANSI misread at the root.

## 2.6.1 — The badge stays visible when the sidebar is collapsed

**Why**: in collapsed mode the badge protruded past the button's edge (`top:-4px`) and was either clipped by the sidebar's collapse-animation container or covered by the row above — so the user could not see new mail, and the notification feature failed exactly in the quiet rail form where it mattered most.

**What changed**: in collapsed mode a 14px dot is drawn inside the 36px circle's top-right corner; the expanded outside badge is unchanged; the badge got `z-index:1`.

## 2.6.0 — From DOM injection to an official slot

**Why**: the old version injected its button straight into the shell's footArea container and defended itself with two MutationObservers. That bought positional freedom at the price of three real incidents: before v2.5.1 an observer microtask storm froze the entire page ("Loading plugins…"); substring matching against hashed CSS class names broke silently whenever the shell was upgraded; and hover feedback never rendered at all because inline styles outranked the stylesheet.

**What changed**: the whole button now registers through the `sidebar.footer.action` slot (the same official channel dsh-dock's Whale Bay menu uses), and React owns its lifecycle — observer, self-healing and class-name probing are all deleted (about 150 lines of defensive code gone). `wide`/rail geometry comes from the slot props; the popover became a React Portal (click-outside still closes it). `order:10` keeps the original visual position (above Whale Bay).

## 2.5.3 — Hover feedback actually renders now

**Why**: v2.5.2 fixed the hover colour variable (`--dsh-` → `--dsw-`) and the user still saw no background — an inline `background:'transparent'` always outranks a stylesheet's `:hover` rule.

**What changed**: the transparent background moved out of the inline style and into the injected stylesheet so `:hover` can override it.

## 2.5.2 — Aligned with the native settings button, item by item

**Why**: hover used the wrong token prefix (`--dsh-`, a variable that does not exist), and the collapsed margin did not match the native settings button (4px vs 8px/10px), which is visibly crooked side by side.

**What changed**: hover token aligned to `--dsw-alias-interactive-bg-hover`; rail margin aligned to `8px 0 10px`; the remaining geometry parameters were checked one by one and were already correct.

## 2.6.0 batch — 14 fixes from a deep review

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

## 2.5.1 and earlier

(No changelog was kept for these versions; the quirks and the field lessons are encoded in the `lib/olk.ps1` header comment.)
