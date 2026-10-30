# HANDOFF — dsh-arya-maildesk

State of the published release, what was verified, what is still open, and how to continue on a machine
that actually has classic Outlook. Keep this current when a version ships.

## Current version

| | |
|---|---|
| Package | `@aryacode/dsh-arya-maildesk` |
| Version | **1.0.1** |
| Repository | `KonnectSuite/dsh-arya-maildesk` (public) |
| Built from commit | `d41b047` — `fix: the bundle patch named an unscoped package, so 1.0.0 could install but never activate (1.0.1)` |
| Published tarball | `https://registry.npmjs.org/@aryacode/dsh-arya-maildesk/-/dsh-arya-maildesk-1.0.1.tgz` |
| sha256 | `7186fe0abe81b8cf3254166663fafaff71dc2a8b3e8feaeefbd338119385d4da` (83,353 bytes) |
| npm shasum | `3b4764aef530d5f22e92880b871321acf0d2b0d5` |
| GitHub release | `https://github.com/KonnectSuite/dsh-arya-maildesk/releases/tag/v1.0.1` |
| Host requirement | `dsh.engines.dsh`: `^0.2.1-rc.3 \|\| >=0.2.2-alpha.2` |
| Licence | MIT (fork of `UnknowCao/dsh-outlook` @ `5a2ec7d`; see `NOTICE`) |

`1.0.0` is **deprecated on npm**: it installed and then failed to activate (see *The 1.0.0 bug* below).
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
| `index.js` | host | 25 `maildesk_*` tools, the COM bridge spawn, the draft store, the panel + compose HTTP routes, the fixture provider |
| `lib/olk.ps1` | bridge | one-shot dispatcher: `OLK_ACTION` + one JSON in, one JSON out; 25 actions over COM |
| `lib/watch.ps1` | bridge | `NewMailEx` listener for the badge (self-healing, 30 s) |
| `client.js` | browser | one `window.__ModuleLoader__` bundle: badge, panel, composer, chat card, Sidebar tab |
| `cordis.patch.yml` | bundle | the host-plane row that mounts `index.js` |

Registration surface (all published APIs, no harness edits): `main` key `mail`, `sidebar.panellist` id
`mail`, `sidebar.footer.action` id `newmail-badge`, `tool.call.toolview` key `maildesk_compose`,
`sidebarRightTabs` type `maildesk-draft` + `sidebar.right.pane.tab` key `dsh-arya-maildesk:draft`.
Routes: `GET /maildesk/api/health`, `POST /maildesk/api/call`, `/maildesk/api/draft`,
`/maildesk/api/send` — all same-origin, POST-only, 256 KB cap.

The browser half **cannot** use `host.call` (that provider is for model-authored dynamic plugins, not for
bundles), which is why the panel talks over routes.

## Where things stand

Verified on **Dev AryaCode** (`%LOCALAPPDATA%\AryaCode-dev\home`, profile `web`, repo linked), never on an
installed application:

| Milestone | Evidence |
|---|---|
| Panel | folder rail, list, reading pane, paging, selection, mark read/move/delete round-trips through the real route (`docs/panel.png`) |
| Safe reading pane | allowlisted HTML, `<script>`/`on*` dropped, remote image → `🖼 (remote image blocked)` chip, Proofpoint `urldefense` link resolved (`docs/panel-html.png`) |
| Composer | Reply all prefilled To/Cc/Subject, quoted original, two-step confirm, `Sent` toast, Drafts unchanged + Sent Items +1 (`docs/compose.png`, `docs/compose-confirm.png`) |
| Draft card | keyed `tool.call.toolview` row on the host draft, `[Edit in panel] [Send…] [Discard]` (`docs/draft-card.png`) |
| Draft beside chat | real right-Sidebar tab, conversation still visible (`docs/panel-beside-chat.png`) |
| No classic Outlook | the panel names the reason instead of reporting zero mail (`docs/panel-new-outlook.png`) |

Offline:

- `npm test` — bundle guard (**13** checks) + **59** assertions: 14 in `test/folder-resolution.test.ps1`,
  45 in `test/mail-actions.test.ps1` against a fake COM object model.
- `node test/panel-api.test.mjs` — **29** assertions driving the real host routes with fake `req`/`res`.
- `node --check index.js` / `client.js` — the only build gate; there is no bundler and no TypeScript.

**Not verified: nothing here has read a real mailbox.** Everything above runs on synthetic mail from
`MAILDESK_FIXTURE=1`. The COM path is proven only in that its PowerShell parses and its JSON contract is
exercised. Closing that is the job below.

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
(`minimumReleaseAge`, default 1440 minutes). 1.0.1 was published `2026-10-30T03:38Z`, so it is held until
`2026-10-31T03:38Z` (23:38 local, EDT) unless exempted; before then the install is refused and the market
marks it failed.

Two edits, in this order, in `%USERPROFILE%\.aryacode\profiles\desktop\pnpm-workspace.yaml`:

```yaml
minimumReleaseAgeExclude:
  - '@aryacode/dsh-arya-maildesk@1.0.1'     # REPLACE any @1.0.0 line, do not append
```

pnpm honours only the **first** rule per package name, so a leftover `@1.0.0` line shadows the new one.
Then: Plugins → *Arya Plugin Market* → search `maildesk` → **Install**. The catalogue is CDN-cached for up
to 5 minutes, so a just-pushed version can take a moment to appear.

If the market is unavailable, the equivalent without the hold: copy the tarball over and
`pnpm add file:<path>.tgz` in the profile with the app's own pnpm
(`…\AryaCode\resources\runtime\primary-runtime\dependencies\pnpm\bin\pnpm.cjs`), then add
`@aryacode/dsh-arya-maildesk` to `dsh.profile.bundles` in that profile's `package.json`.

Then **restart AryaCode**. Bundles mount at startup.

### 3. Activation ladder — run these in order, one at a time

Each rung isolates one layer; stop at the first failure and read the signature table.

1. **`maildesk_account`** — must return your display name and SMTP address. This is the gate: it proves
   COM activation, the Outlook profile and the store. Everything below depends on it.
2. **`maildesk_list_folders`** — folder tree, ids, item counts. Proves the folder walk and the count reads.
3. **`maildesk_list_mail`** with `limit: 5` — proves `Restrict`/`Sort` against the real store.
4. **`maildesk_read_mail`** on one id, `format: 'html'` — proves body retrieval and the sanitizer against a
   real corporate mail (Proofpoint links, blocked images).
5. **Panel check** — Mail row, rail, list, reading pane. Then select a row and use *Mark unread → Refresh*:
   proves a write round-trip through the route.
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
| powershell spawn denied | EDR/AppLocker blocks `powershell.exe`; that is the blocker, not the plugin |
| panel shows "No classic Outlook answered", all actions disabled | the bridge is alive and reported the reason — read its sentence, it is accurate |

### 5. What to report back

For each rung: pass/fail and, on failure, the verbatim error. If rung 1 passes, the backend is real and the
rest is polish. If it fails, the string above identifies the layer. Record which rung first diverged — that
is the whole value of the session.

## The 1.0.0 bug (why the guard exists)

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

## Known gaps

1. **No release has read a real mailbox.** The next session closes this.
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
9. **No search box in the panel.** `maildesk_search_mail` exists as a tool only.
10. **The Fromm mailbox needs IT** if Graph is ever wanted: the tenant demands admin consent and has not
    enabled the self-service approval flow. Not on the critical path for classic Outlook.

## Release procedure

1. `node scripts/set-version.mjs <x.y.z>` (never `JSON.stringify` the manifest — it is 4-space formatted).
2. `npm test` and `node test/panel-api.test.mjs` green; `node --check` both halves.
3. `npm publish --access public`. The registry holds new packages under "being processed" for a few minutes;
   confirm `npm view … dist-tags.latest`.
4. `git tag vX.Y.Z && git push origin main vX.Y.Z`, then
   `gh release create vX.Y.Z --repo KonnectSuite/dsh-arya-maildesk …` — **pass `--repo`**, because `gh`
   otherwise resolves the repository from the `upstream` remote.
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
  from a shell.

## Next, in priority order

1. Run the ladder on the work laptop; record where it diverges. **(blocks everything else)**
2. If rung 1 passes: fix whatever the first failing rung exposes, then re-run the ladder to completion.
3. Panel search box over `maildesk_search_mail` (small, high daily-use value).
4. Persistent drafts (write the draft store to `storage`), so a restart stops eating replies.
5. A morning-brief prompt/tool built on `list_mail` + `read_mail` — daily value with no new UI.
6. 1.1.0: localisation layer, then the localisation gate can be honoured.
7. Optional, decision pending: a `graph` provider for the new Outlook, which needs IT consent on a
   corporate tenant and no consent at all on personal accounts.
