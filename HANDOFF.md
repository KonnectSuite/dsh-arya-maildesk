# HANDOFF — dsh-arya-maildesk

State of the published release, what was verified, and what is still open. Keep this current when a
version ships.

## Current version

| | |
|---|---|
| Package | `@aryacode/dsh-arya-maildesk` |
| Version | **1.0.0** |
| Repository | `KonnectSuite/dsh-arya-maildesk` (public) |
| Built from commit | `66f5838` — `release: publish as @aryacode/dsh-arya-maildesk with a host requirement and an icon` |
| Published tarball | `https://registry.npmjs.org/@aryacode/dsh-arya-maildesk/-/dsh-arya-maildesk-1.0.0.tgz` |
| sha256 | `656d66c80d9e403675179e03b9f53cdd47eb606c466b42d969f743e20909a014` (82,665 bytes) |
| npm shasum | `186a14e5a75452b111ba89d1696356414830b5a7` |
| Host requirement | `dsh.engines.dsh`: `^0.2.1-rc.3 \|\| >=0.2.2-alpha.2` |
| Licence | MIT (fork of `UnknowCao/dsh-outlook` @ `5a2ec7d`; see `NOTICE`) |

The earlier unscoped `dsh-arya-maildesk@1.0.0` is deprecated in favour of this package — same code,
scoped name. npm also holds a `0.0.0-stage` placeholder for the unscoped name, an artifact of the
first publish; it is inert and can only be removed in the npm UI.

## What was verified, and where

All of it on **Dev AryaCode** (`%LOCALAPPDATA%\AryaCode-dev\home`, profile `web`, plugin linked from
this repository), never on the installed application.

| Milestone | Evidence |
|---|---|
| Panel | folder rail, message list, reading pane, paging, selection, mark read/move/delete round-trips through the real route (`docs/panel.png`) |
| Safe reading pane | allowlisted HTML, `<script>` and `on*` dropped, remote image → `🖼 (remote image blocked)` chip, Proofpoint `urldefense` link resolved to the vendor URL (`docs/panel-html.png`) |
| Composer | Reply all prefilled To/Cc/Subject, quoted original below the body, two-step confirm card, `Sent` toast, Drafts unchanged and Sent Items +1 (`docs/compose.png`, `docs/compose-confirm.png`) |
| Draft card in chat | keyed `tool.call.toolview` row rendering the host draft, `[Edit in panel] [Send…] [Discard]` (`docs/draft-card.png`) |
| Draft beside chat | real right-Sidebar tab on the same draft, conversation still visible (`docs/panel-beside-chat.png`) |
| New Outlook on a machine without classic Outlook | the panel explains the reason instead of reporting zero mail (`docs/panel-new-outlook.png`) |

Offline, on this machine (no Outlook installed):

- `npm test` — **59** assertions: 14 in `test/folder-resolution.test.ps1`, 45 in `test/mail-actions.test.ps1` against a fake COM object model.
- `node test/panel-api.test.mjs` — **29** assertions driving the real host routes with fake `req`/`res`.

## Known gaps

1. **No release of this code has read a real mailbox.** Everything above runs against synthetic mail. The COM path is proven only in that its PowerShell parses and its JSON contract is exercised. The first real mailbox — a work laptop with classic Outlook — is what closes this.
2. **User-visible strings are English only.** The panel, composer, card and tab have no locale layer; there is no `locale/*.json`. The catalogue carries an English and a Chinese description, but the in-app text does not switch. This is a deliberate deviation from the market's localisation gate, recorded here rather than hidden — the fix is a `locale` injection plus a string table, and it is the first item for 1.1.0.
3. **`TESTCASES.md`'s 49 cases were inherited** and run against Exchange by upstream's author, not by this release. The actions added after the fork have no case of their own.
4. **Classic Outlook only.** The new Outlook for Windows has no COM surface; there is no Graph or IMAP backend in this build. The provider seam (`fixture` vs the PowerShell bridge) is in place, but a second provider is not implemented.
5. **Soft delete only.** `maildesk_delete_mail` moves to Deleted Items and reports `recoverable: true`; there is no permanent-delete path.
6. **English folder names.** `Resolve-Folder` matches `Inbox`, `Sent Items`, `Drafts`, `Deleted Items`; any other folder must be addressed by entryId (get one from `maildesk_list_folders`).
7. **Drafts are in memory.** A host restart loses unsent drafts. Nothing is written to disk.
8. **`move_mail` changes entryIds**, so a stale selection has to be re-read after a move.

## How it was released

1. `npm publish --access public` from this repository (version bumped with `node scripts/set-version.mjs`).
2. Registry confirmed: `dist-tags.latest` = `1.0.0`, tarball answers 200.
3. Catalogue: `node scripts/publish-arya-catalog.mjs --pull`, then `node scripts/publish-arya-release.mjs --npm 1.0.0 --source KonnectSuite/dsh-arya-maildesk --name '@aryacode/dsh-arya-maildesk' --description … --description-zh …`, then `--push` from `KonnectSuite/dsh-market`.
4. Live document verified at `https://raw.githubusercontent.com/KonnectSuite/dsh-arya-market/catalog/data/arya-plugins.json`.
