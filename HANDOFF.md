# HANDOFF — dsh-arya-maildesk

State of the published release, what was verified, what is still open, and how to continue on a machine
that actually has classic Outlook. Keep this current when a version ships.

## Current version

| | |
|---|---|
| Package | `@aryacode/dsh-arya-maildesk` |
| Version | **1.0.4** — built and installed on the work laptop from a local tarball; **not on npm yet** |
| Repository | `KonnectSuite/dsh-arya-maildesk` (public) |
| Built from commit | working tree (1.0.4 is uncommitted as this is written) |
| Last published | **1.0.3** — `https://registry.npmjs.org/@aryacode/dsh-arya-maildesk/-/dsh-arya-maildesk-1.0.3.tgz` |
| 1.0.3 sha256 | `1b1be6b2667410ce7bc3bf6440b50f9233dc18c87ff7b3562a466d01736b2f5d` (85,872 bytes) |
| 1.0.3 npm shasum | `426103d47383a11976a81bc6b118bf97410c75da` |
| GitHub release | `https://github.com/KonnectSuite/dsh-arya-maildesk/releases/tag/v1.0.3` |
| Host requirement | `dsh.engines.dsh`: `^0.2.1-rc.3 \|\| >=0.2.2-alpha.2` |
| Licence | MIT (fork of `UnknowCao/dsh-outlook` @ `5a2ec7d`; see `NOTICE`) |

`1.0.0` and `1.0.1` are **deprecated on npm**: each installed and then failed to activate (see *Two bugs the tests could not see* below).
The unscoped `dsh-arya-maildesk` name is deprecated too; its `0.0.0-stage` stub is inert and can only be
removed in the npm UI.

## What this is

A Cordis plugin that gives AryaCode/DSH a real mail client: a three-pane panel, harness-owned reply
drafts, a draft card inside the conversation, and a draft tab beside it. It drives the **local classic
Outlook** over COM through one resident-free PowerShell dispatcher — no Graph app registration, no SMTP
password, no admin consent.

Two halves, one package:

| File | Half | What it does |
|---|---|---|
| `index.js` | host | 24 `maildesk_*` tools, the COM bridge spawn, the draft store, the panel + compose HTTP routes, the fixture provider |
| `lib/olk.ps1` | bridge | one-shot dispatcher: `OLK_ACTION` + one JSON in, one JSON out; 25 actions over COM |
| `lib/watch.ps1` | bridge | `NewMailEx` listener for the badge (self-healing, 30 s) |
| `client.js` | browser | one `window.__ModuleLoader__` bundle: badge, panel, composer, chat card, Sidebar tab |
| `cordis.patch.yml` | bundle | the host-plane row that mounts `index.js` |

Registration surface (all published APIs, no harness edits): `main` key `mail`, `sidebar.panellist` id
`mail`, `sidebar.footer.action` id `newmail-badge`, `tool.call.toolview` key `maildesk_compose`,
`sidebarRightTabs` type `maildesk-draft` + `sidebar.right.pane.tab` key `dsh-arya-maildesk:draft`.
Routes: `GET /maildesk/api/health`, `GET|POST /maildesk/api/ask`; `POST /maildesk/api/call`, `/maildesk/api/draft`,
`/maildesk/api/send`, `/maildesk/api/fixture` — all same-origin, request bodies capped at 256 KB.

`/maildesk/api/ask` is the in-panel agent: POST runs a real harness session (`ctx.get('sessionController')`) with
the question plus the panel's context, GET reads that session's transcript back as `{role, text}`. Nothing in the
client talks to a model directly, and the agent there is the same agent the chat gets, tools and gates included.

The browser half **cannot** use `host.call` (that provider is for model-authored dynamic plugins, not for
bundles), which is why the panel talks over routes.

## Where things stand

Verified in **Dev AryaCode** (`%LOCALAPPDATA%\AryaCode-dev\home`, profile `web` — that tree is not on the work
laptop, see *the dev lane* below), and — from 1.0.2 on — in the **installed** application on this PC: Mail row in
the sidebar, 24 tools registered live.

On **2026-10-30**, on this PC (the work laptop), classic Outlook was signed in and **rungs 1–5 plus 8 ran against a
real mailbox**: `maildesk_account` → `kphung@frommconnect.com`; `maildesk_list_folders` → three stores (one Inbox
1,573 items / 141 unread with 88 child folders; another 2,728 / 2,095); `list_mail` → 100 rows of 1,573; `read_mail`
→ a real HTML body in the reading pane; the panel opening on real mail; `mark_read` → Inbox unread 141 → 140 → 141,
each state read back from a fresh folder walk; and **Reply** → a prefilled draft (To, `RE:` subject, quoted
original) discarded without writing an Outlook item. Move, delete and send are still unrun there.

The same session added the thing the owner asked for: **the agent is in the Mail page**. `Ask` opens a drawer over
the mailbox; the host runs a real harness session with the same `maildesk_*` tools and the same gates, and the
panel renders the exchange. Verified live: "how many unread mails are in my Inbox?" → **141** (the tool answer, not
a guess), and "name the newest unread mail" → sender, subject and date, with the conversation sat under the message
list. The host route is `/maildesk/api/ask`; the client never talks to a model itself. Its session id is persisted
so the conversation survives a reload, the transcript names the tools it calls, and the question carries the
selected mail's `entryId`, so "reply to this one" resolves without a search.

`Move to…` is a picker of its own now, because the lazy rail only knew the folders a person had expanded: it
filters by path and can load the whole tree on demand (one slow walk, cached).

That session also found and fixed the reason the panel looked broken: the first paint waited for a whole-mailbox
folder walk (545 folders, 16 s) before it asked for any mail, so "opening Mail" meant ~24 s of skeleton rows. The
panel now paints the message list first and grows the folder rail lazily — **first mail in 3 s, rail in 1.5 s, one
expand 2.5 s** on that same mailbox. 1.0.4 is that build; it is installed on this PC as a local tarball because the
machine has no npm credentials, so nothing after 1.0.3 exists on the registry yet.

The interface milestones below were captured on synthetic mail; the panel and reading pane have also now been seen
on the real mailbox, the rest have not:

| Milestone | Evidence |
|---|---|
| Panel | folder rail, list, reading pane, paging, selection, mark read/move/delete round-trips through the real route (`docs/panel.png`) |
| Safe reading pane | allowlisted HTML, `<script>`/`on*` dropped, remote image → `🖼 (remote image blocked)` chip, Proofpoint `urldefense` link resolved (`docs/panel-html.png`) |
| Composer | Reply all prefilled To/Cc/Subject, quoted original, two-step confirm, `Sent` toast, Drafts unchanged + Sent Items +1 (`docs/compose.png`, `docs/compose-confirm.png`) |
| Draft card | keyed `tool.call.toolview` row on the host draft, `[Edit in panel] [Send…] [Discard]` (`docs/draft-card.png`) |
| Draft beside chat | real right-Sidebar tab, conversation still visible (`docs/panel-beside-chat.png`) |
| No classic Outlook | the panel names the reason instead of reporting zero mail (`docs/panel-new-outlook.png`) |
| Sample mailbox | `SAMPLE MAIL` pill, **Preview with sample mail** / **Back to live mail**, whole mailbox on a PC with no Outlook (`docs/panel-sample-mail.png`) |

Offline:

- `npm test` — bundle guard (**15** checks) + **59** PowerShell assertions: 14 in
  `test/folder-resolution.test.ps1`, 45 in `test/mail-actions.test.ps1` against a fake COM object model.
- `node test/panel-api.test.mjs` — **29** assertions driving the real host routes with fake `req`/`res`.
- `node test/sample-mail.test.mjs` — **17** assertions: no fixture → the live provider fails loudly (no
  invented empty mailbox) → toggle on → folders, rows, read → toggle off → fails again, plus route guards.
- `node --check index.js` / `client.js` — the only build gate; there is no bundler and no TypeScript.

`npm test` also needs permission to spawn `powershell.exe`. Under a restricted harness — this session's file
sandbox, for one — `test/sample-mail.test.mjs` stops at the first live-provider check with `spawn EPERM`: the route
is honest about the failure, but the text is the raw spawn error, which that assertion's regex does not accept. Run
the suites from a normal shell before treating a red run as a regression.

**Still not verified: nothing has been written or sent from a real mailbox.** Rungs 1–5 pass on the work laptop —
COM activation, the folder walk, paging, body retrieval and the panel itself. Everything that *writes* (move,
delete, mark read), the composer and send have only ever run on synthetic mail from `MAILDESK_FIXTURE=1`, and the
send gate has never seen a real recipient. Closing that is the rest of the ladder.

## Continue on the work laptop (classic Outlook)

You are picking this up cold. Read this section top to bottom before touching anything.

### 1. Preconditions

- Classic Win32 Outlook installed, one work account signed in. Outlook may be closed; COM starts it.
- AryaCode installed, host version satisfying `^0.2.1-rc.3 || >=0.2.2-alpha.2`. Check before blaming the
  plugin: an older host refuses to mount and the market marks the install failed.
- Nothing else. **No IT ticket, no app registration, no Graph consent** — that is the whole point of the
  COM path.

### 2. Install

The market card installs a version into a pnpm profile, and pnpm holds **any version younger than 24 h**
(`minimumReleaseAge`, default 1440 minutes). For 1.0.3 that hold ends `2026-10-31T04:10Z` (**00:10 local, EDT, on
Oct 31**). 1.0.3 is installed on this PC, so the order below is the one that actually worked:

1. **Publish, then refresh the catalogue** — `npm publish --access public`, then the market scripts in the
   *Release procedure* below. The catalogue is CDN-cached for up to 5 minutes, so a just-pushed version can take a
   moment to appear.
2. **Install from the market.** It is refused on the age gate: red `!`, `added 0`, and a `Done in Ns` tail that is
   **not** success. That refusal is expected, not an outage.
3. **Exempt the exact version** in `%USERPROFILE%\.aryacode\profiles\desktop\pnpm-workspace.yaml`:

    ```yaml
    minimumReleaseAgeExclude:
      - '@aryacode/dsh-arya-maildesk@1.0.3'     # REPLACE any older @1.0.x line, do not append
    ```

    pnpm honours only the **first** rule per package name, so a leftover `@1.0.0` line shadows the new one — and
    each future publish needs its own line in place of the last.
4. **Install again.** The failed row is not a successful install: retry it, or remove it and add it back. On success
   the market copies the tarball to
   `<profile>\.dsh-market\releases\aryacode-dsh-arya-maildesk-<version>-<sha256>.tgz` and rewrites the profile
   dependency to `file:` that path — which is why the digest in the table above is also the file you can hash.
5. **Restart AryaCode.** Bundles mount at startup.

If the market is unavailable, the equivalent without the hold: copy the tarball over and
`pnpm add file:<path>.tgz` in the profile with the app's own pnpm
(`…\AryaCode\resources\runtime\primary-runtime\dependencies\pnpm\bin\pnpm.cjs`), then add
`@aryacode/dsh-arya-maildesk` to `dsh.profile.bundles` in that profile's `package.json`.

### 3. Activation ladder — run these in order, one at a time

Each rung isolates one layer; stop at the first failure and read the signature table.

1. **`maildesk_account`** — must return your display name and SMTP address. This is the gate: it proves
   COM activation, the Outlook profile and the store. Everything below depends on it.
2. **`maildesk_list_folders`** — folder tree, ids, item counts. Proves the folder walk and the count reads.
3. **`maildesk_list_mail`** with `limit: 5` — proves `Restrict`/`Sort` against the real store.
4. **`maildesk_read_mail`** on one id, `format: 'html'` — proves body retrieval and the sanitizer against a
   real corporate mail (Proofpoint links, blocked images).
5. **Panel check** — Mail row, rail, list, reading pane. Then select a row and use *Mark unread → Refresh*:
   proves a write round-trip through the route. If the header reads `Classic Outlook · COM` while the list is empty
   and the footer says `unread count unavailable`, the bridge returned `unreadInFolder: -1` — that is the ladder
   diverging, not a slow load.
6. **`maildesk_move_mail`** one mail to another folder — note the entryId changes; re-read the folder.
7. **`maildesk_delete_mail`** on one mail with the right `expected` count — soft delete to Deleted Items.
   The `expected` gate must refuse a wrong count; try it once deliberately.
8. **`maildesk_compose`** with `mode: 'reply'`, no send — the card must appear in chat and
   `[Edit in panel]` must open the same draft, not a second one.
9. **Send** — the confirm card must list the exact content; the hash binds it. After sending: Drafts
   unchanged, Sent Items +1.

### 4. Failure signatures

| Signature | Meaning |
|---|---|
| `80040154 Class not registered` | no COM Outlook on this machine — new Outlook is the handler, or Office has no Outlook |
| `classic desktop Outlook not found, but the new Outlook for Windows is installed` | our own message: the WebView2 client is present, it has no COM surface |
| `ERR_MODULE_NOT_FOUND: Cannot find package '<name>'` | the bundle patch names a package that does not exist — this was 1.0.0's bug; check `cordis.patch.yml`'s `name:` equals `package.json`'s `name`, scope included |
| `SEND BLOCKED: no human confirmation token` | the bridge refused a send because the host did not mint `OLK_CONFIRM`; normal for any direct bridge call |
| "a program is trying to send email on your behalf" | Outlook's own programmatic-access guard; drafts still save, sending may prompt once |
| `spawn EPERM` / powershell spawn denied | EDR/AppLocker — or a harness sandbox — blocks `powershell.exe`; that is the blocker, not the plugin. The text is raw, not one of our sentences |
| panel shows "No classic Outlook answered", all actions disabled | the bridge is alive and reported the reason — read its sentence, it is accurate |

### 5. What to report back

For each rung: pass/fail and, on failure, the verbatim error. If rung 1 passes, the backend is real and the
rest is polish. If it fails, the string above identifies the layer. Record which rung first diverged — that
is the whole value of the session.

### 6. Working from this laptop (GitHub CLI is available here now)

Earlier sessions were crippled on this machine because `gh` was missing. **It is installed and authenticated
as `KonnectSuite` on the work laptop now**, so the full loop can run there instead of bouncing back to the
other PC: fix, commit, push, publish, release, catalogue, verify.

```powershell
gh auth status                      # expect: logged in as KonnectSuite, scopes gist, read:org, repo
gh repo view KonnectSuite/dsh-arya-maildesk
git -c http.sslBackend=openssl push origin main    # schannel fails on this network without -c http.sslBackend=openssl
```

Rules that still apply there:

- **Always pass `--repo KonnectSuite/dsh-arya-maildesk` to `gh release`.** `gh` resolves the repository from the
  remotes, and this clone has carried an `upstream` remote pointing at the original project at times — without
  `--repo` it once tried to push a tag to `UnknowCao/dsh-outlook` and failed. The current clone has only `origin`;
  keep the flag anyway, so a re-added upstream cannot bite.
- Never republish a version; bump it. `npm deprecate` retires a bad one, `npm unpublish` is refused for
  tokens that skip 2FA.
- `git push --tags` also carries the fork's inherited tags; push `v<version>` explicitly and delete strays
  with `git push origin --delete <tag>`.
- Publish order is fixed: `npm publish --access public` → `dsh-market` `--pull` → `--release --npm <v>` →
  `--push` → verify the **live raw URL** with a cache-buster, not the local file.

## Two bugs the tests could not see (why the guard exists)

1.0.0 installed cleanly on a work laptop and then refused to activate:

```
outlook (dsh-arya-maildesk): failed to import: Error [ERR_MODULE_NOT_FOUND]:
Cannot find package 'dsh-arya-maildesk' imported from ~\.aryacode\profiles\desktop\
```

`cordis.patch.yml` named the package **unscoped** — a leftover of the rename to
`@aryacode/dsh-arya-maildesk`. The loader resolves that string with `import()`, so a correct, digest
verified tarball could sit in `node_modules` and still not mount. Development never caught it because the
dev profile links this repository under the *unscoped* name, so the stale string resolved by luck.

`test/bundle.test.mjs` now fails on that patch shape: row names must equal `package.json`, the name must
stay scoped, every entry point must exist **and** be in `files`, and `dsh.engines.dsh` must be declared.

**Lesson:** tests that exercise code do not prove a package mounts. Anything the loader reads by name needs
its own assertion.

### 1.0.1 — same bug, one layer down

1.0.1 fixed the patch row and the installed app still refused to boot:

```
web boot: 1 entry did not activate
@aryacode/dsh-arya-maildesk: import failed (see console for the import error)
```

The console held the real line: `client-modules: could not load "@aryacode/dsh-arya-maildesk"`. `client.js`
called `window.__ModuleLoader__.load({ id: 'dsh-arya-maildesk' })` — **unscoped** — while the boot module
table keys every client bundle by *package name*. The host half was fine (24 tools registered, the bridge
answered); the browser half registered itself where nothing looks for it. Development hid it exactly the
same way: the dev profile links the repo under the unscoped name.

1.0.2 fixes the id, and `test/bundle.test.mjs` grew a 15th check: the module id must equal `package.json`'s
`name`. Replaying 1.0.1's `client.js` through the guard fails on that line, which is the point.

**Reproduce a bug like this by installing the packed tarball.** The dev lane's `file:` link resolves stale
names by luck. The fatal signal is `entry did not activate` — never `Done in 4.1s`.

Verified after the fix (this PC, installed AryaCode, desktop profile): the sidebar shows **Mail**, the 24
`maildesk_*` tools are registered, and `maildesk_account` answers from the signed-in profile instead of failing
silently.


## Known gaps

1. **No send has touched a real mailbox.** Reads, mark-read, move, delete and the reply composer pass (rungs 1–8);
   `send` is still synthetic-only, and the send gate has never seen a real recipient.
2. **User-visible strings are English only.** No locale layer, no `locale/*.json`. Deliberate deviation from
   the market's localisation gate, recorded rather than hidden; first item for 1.1.0.
3. **`TESTCASES.md`'s 49 cases were inherited** and run by upstream's author. Post-fork actions have no case
   of their own; the ladder above is the substitute until they do.
4. **Classic Outlook only.** No Graph or IMAP provider. The seam exists (`fixture` vs bridge); a second
   provider is not implemented.
5. **Soft delete only.** Moves to Deleted Items, reports `recoverable: true`; no permanent delete.
6. **English folder names** in `Resolve-Folder` (`Inbox`, `Sent Items`, `Drafts`, `Deleted Items`); anything
   else must be addressed by entryId from `maildesk_list_folders`.
7. **Drafts are in memory.** A host restart loses unsent drafts.
8. **`move_mail` changes entryIds** — re-read after a move.
9. **Search covers one folder at a time.** The panel's search box asks the bridge for the *current* folder, subject
   first and sender second; there is no whole-mailbox or free-text search, and `maildesk_search_mail` still does
   the richer queries (dates, recipient, other folders) from chat.
10. **The panel agent is new and thin.** It keeps ONE session per host process (`Mail assistant`, titled on
    creation) and creates a fresh one after a restart; it polls the transcript every 1.5 s instead of streaming;
    it shows the conversation but not the tool calls the agent made; and its session is a normal workspace
    session, so it also appears in the sidebar. Those are the next four cuts.
11. **The Fromm mailbox needs IT** if Graph is ever wanted: the tenant demands admin consent and has not
    enabled the self-service approval flow. Not on the critical path for classic Outlook.

## Release procedure

1. `node scripts/set-version.mjs <x.y.z>` (never `JSON.stringify` the manifest — it is 4-space formatted).
2. `npm test` and `node test/panel-api.test.mjs` green; `node --check` both halves.
3. `npm publish --access public`. The registry holds new packages under "being processed" for a few minutes;
   confirm `npm view … dist-tags.latest`.
4. `git tag vX.Y.Z && git push origin main vX.Y.Z`, then
   `gh release create vX.Y.Z --repo KonnectSuite/dsh-arya-maildesk …` — **pass `--repo`**, because `gh`
   resolves the repository from the remotes and this clone has carried an `upstream` at times.
5. Catalogue, from `KonnectSuite/dsh-market`: `node scripts/publish-arya-catalog.mjs --pull`, then
   `node scripts/publish-arya-release.mjs --npm <x.y.z> --source KonnectSuite/dsh-arya-maildesk
   --name '@aryacode/dsh-arya-maildesk' --description … --description-zh …`, then `--push`. Verify the live
   raw URL, not the local file.
6. Retire a bad version with `npm deprecate`, never unpublish.

## Operating rules

- **Never call it "works" until it has read real mail.** Say "synthetic" when it is the fixture. This rule
  exists because the word was used loosely once and cost a session.
- **Chinese never crosses `pwsh -Command`** — it arrives mojibake'd. Stage text in a file, read it with
  `-Encoding UTF8`, pass variables; or write files with the `write`/`edit` tools.
- **The bundled client face is hand-written** (`client.js`, no build step). `node --check` is the only gate.
- **`minimumReleaseAgeExclude` honours only the first rule per package name** — replace the line, do not append.
- **Don't edit installed app bundles.** Releases go out through npm and the market.
- **The dev lane**: dev profile `web`, repo linked by dependency `dsh-arya-maildesk: file:…` plus a
  `mklink /J` junction; server with
  `node --import tsx/esm apps/cli/src/bin.ts --profile web --port 58231 --host 127.0.0.1 --no-open`
  (a new `?token=` on every restart); `MAILDESK_FIXTURE=1` for synthetic mail, `OLK_WATCH=0` to silence the
  watcher. `--profile desktop` refuses to run outside Electron, so the desktop profile cannot be driven
  from a shell. This tree lives on the original development PC: on the work laptop
  `%LOCALAPPDATA%\AryaCode-dev\home` does not exist and `%USERPROFILE%\.aryacode\profiles\` holds only
  `desktop`, so there the installed package is the only lane.

## Next, in priority order

1. **Rungs 6–9** (move, delete, compose, send) on the work laptop; record where it diverges. Reads already pass.
   **(blocks the release claim, not the UI work)**
2. **Finish the panel agent** (it exists and works): persist its session id across restarts, stream instead of the
   1.5 s poll, show the tool calls it makes, and let it act on the selected mail without being told the entryId.
3. Panel search box over `maildesk_search_mail` (small, high daily-use value).
4. Persistent drafts (write the draft store to `storage`), so a restart stops eating replies.
5. Publish 1.0.4 (npm credentials needed on this machine — none are present), then catalogue + GitHub release.
6. A morning-brief prompt/tool built on `list_mail` + `read_mail` — daily value with no new UI.
7. 1.1.0: localisation layer, then the localisation gate can be honoured.
8. Optional, decision pending: a `graph` provider for the new Outlook, which needs IT consent on a
   corporate tenant and no consent at all on personal accounts.
