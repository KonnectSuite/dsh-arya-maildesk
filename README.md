# dsh-outlook

> **Let an AI drive the Outlook you are already logged into, on a work laptop where IT will not give you a Graph API permission.**
> No credentials, no admin rights, and nothing leaves the machine without asking you first.

[![DSH Plugin](https://img.shields.io/badge/DSH-Plugin-blue)](https://www.npmjs.com/package/@deepseek-ai/dsh) [![Platform](https://img.shields.io/badge/Platform-Windows%20%2B%20classic%20Outlook-0078D4)](https://support.microsoft.com/en-us/office) [![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](./LICENSE) [![Tools](https://img.shields.io/badge/tools-23-8A2BE2)](#what-it-delivers)

> **This is a fork.** Upstream is [UnknowCao/dsh-outlook](https://github.com/UnknowCao/dsh-outlook) at `5a2ec7d` (v2.11.0), MIT. This fork adds delete, move/archive, read-state and a paged folder listing — the "read" half of the plugin was complete, the "organize" half was missing. `NOTICE` records what came from upstream; `UPSTREAM.md` records how to diff against it. The package name here still says `dsh-outlook` and must be renamed before it is installed next to the upstream release.

*Zero-credential Outlook automation for AI agents on Windows: drives your local desktop Outlook client via COM. No Graph API app registration, no SMTP passwords, no tenant admin rights — everything that leaves the machine passes a three-layer human approval gate.*

## When do you need this?

1. **Your work laptop has classic Outlook, but IT will not let you register an Azure app.** Every Graph API route (outlook-mcp and friends) dies at app registration; this plugin never asks, because it drives the desktop client you are already signed into.
2. **You want the AI to be your mail secretary.** Ask "anything new?" when you sit down; say "reply to this one, Wednesday works" and it drafts, quotes the original, and shows you the result before anything is sent.
3. **You need a meeting and do not know who is free.** One sentence looks up a colleague's free/busy and finds an open room, and you still approve the invite before it goes out.

## What it delivers

**23 AI tools** — reads run on their own, anything that leaves the machine goes through a confirmation popup first:

| Tool | What it does | Gate |
|---|---|---|
| `outlook_check_new` | New-mail arrivals pushed by the watcher (keeps the sidebar badge in sync) | read-only |
| `outlook_search_mail` | Search mail (days / folder / sender / subject / date range / recipient side); results can carry a click-to-open link | read-only |
| `outlook_read_mail` | Full body, recipients and attachment names, truncated when long | read-only |
| `outlook_open_mail` | Open a mail in a desktop Outlook inspector window (marks it read, badge decrements) | local action |
| `outlook_list_folders` | **The folder tree** of every store: path, depth, item and unread counts | read-only |
| `outlook_list_mail` | **One page of a folder**, offset/limit, sortable, unread-only — browse, not search | read-only |
| `outlook_move_mail` | **Move or archive** one mail or up to 500 at once; returns the new entryIds | local action |
| `outlook_delete_mail` | **Delete** to Deleted Items (soft, undoable), with a selection-count guard | local action |
| `outlook_mark_read` | **Mark read or unread**, in bulk | local action |
| `outlook_account` | Account name and SMTP address | read-only |
| `outlook_draft_mail` | Create a draft (**never sends**, attachments allowed) | safe |
| `outlook_send_mail` | Send a mail | three-layer gate |
| `outlook_reply_mail` | Reply / reply-all / forward (quotes the original, forward carries its attachments) | three-layer gate |
| `outlook_calendar_query` | The next N days (recurring items expanded) | read-only |
| `outlook_calendar_create` | Personal appointment directly; with attendees it becomes an invite | invites need approval |
| `outlook_meeting_requests` | Pending meeting invitations | read-only |
| `outlook_respond_meeting` | Accept / tentative / decline | approval (notifies the organizer) |
| `outlook_meeting_update` | Change an appointment or meeting (organizer notifies attendees) | approval |
| `outlook_meeting_cancel` | Cancel a meeting (notifies attendees) or drop a personal appointment | approval |
| `outlook_search_people` | Address book (GAL) by name, email or alias | read-only |
| `outlook_freebusy` | A colleague's busy/free blocks, 30-minute resolution | read-only |
| `outlook_search_rooms` | Room search plus availability for a given window | read-only |
| `outlook_save_attachment` | Save an attachment to disk (no path traversal, never overwrites) | local write |

**Plus a sidebar badge.** New mail shows a red count on the Outlook button; clicking it lists the latest arrivals and **clicking one opens it in Outlook** (marked read, count decrements). Hover previews the list. A `[Open](…)` markdown link in the chat works the same way: the plugin intercepts the left click in-page, so there is no stray tab, and a toast confirms at the bottom.

| Expanded (outside badge) | Collapsed (dot inside the circle) |
|---|---|
| ![expanded sidebar badge](assets/badge-wide.png) | ![collapsed rail badge](assets/badge-rail.png) |

## Quick start

```sh
# short name from npm (recommended)
dsh plugin --profile web add dsh-outlook
# or straight from GitHub
dsh plugin --profile web add https://github.com/UnknowCao/dsh-outlook
```

Then restart the profile (or rely on `patchReload: live`). There is no second step — no credential, no app registration, no Outlook setting to change.

**Prerequisites**: Windows + classic desktop Outlook (Win32 `OUTLOOK.EXE`) signed in + Windows PowerShell 5.1 (ships with the OS). No tenant admin rights.

> This fork is not published yet; the commands above install the upstream release. See `UPSTREAM.md` for the release checklist.

## Try saying

- "Anything new?"
- "Find the mail from the vendor last week about the deliverables"
- "Forward this to Dana and ask her to confirm"
- "Set up a meeting with Wei next Wednesday — check when he is free first"
- "Find a room free tomorrow from two to three"

## How it differs from the alternatives

| | this plugin | Graph API (outlook-mcp etc.) | SMTP (mcp-server-email etc.) |
|---|---|---|---|
| Needs an Azure app registration | ❌ no | ✅ required (usually locked by IT) | ❌ |
| Needs mailbox credentials | ❌ | OAuth token | ✅ SMTP password |
| Sees the state of the client **you are using** (unread, rules, local folders) | ✅ | partial | ❌ |
| Human confirmation before anything is sent | three gates (below) | implementation-dependent | implementation-dependent |
| Platform | Windows, classic Outlook only | anything | anything |

## Security boundary

**Anything that leaves the machine — sending mail, dispatching an invite — passes three independent gates:**

1. **Two-call confirmation**: the first call only returns "here is what would be sent"; the AI must show it to you verbatim, and the call executes only on a retry carrying `confirmed:true` after you allow it.
2. **Content-hash binding**: the content shown at confirmation and the content actually sent are compared by SHA-256 — showing you A and sending B is rejected at the host layer and forced back through confirmation.
3. **Bridge token**: the PowerShell bridge independently refuses every send path without a one-time `OLK_CONFIRM` token. Even if a future bug or a new tool bypassed the first two gates, no token means no mail.

**What it will not do**: send anything silently; store or expose your mailbox credentials (it never sees them); auto-send a draft; write an attachment over an existing file or outside the chosen directory.

Known boundaries: the `/olk-notify/*` routes are reachable only from this host and same-origin checked (other local processes and non-browser requests can still read new-mail subject metadata); the port inside a chat "open" link does not matter for a plain left click, because the client intercepts by pathname and posts to the page's own origin; `/olk-notify/open` is POST-only, so a modifier-click or middle-click on a chat link gets 405 (deliberate; the `outlook_open_mail` tool is unaffected); several tabs share one pending queue; `OLK_WATCH=0` turns the new-mail watcher off. This fork: folder names are matched in English only (`Inbox`, `Sent Items`, `Drafts`, `Deleted Items`) — a mailbox whose folders are named in another language must be addressed by entryId; `delete_mail` is a soft delete to Deleted Items and there is no permanent-delete path; `move_mail` changes entryIds, so re-read the folder afterwards.

## Layout

```
index.js          host half: 23 tool registrations, two-call confirmation, content hash,
                  watcher management, state/open routes
client.js         browser half: sidebar.footer.action slot button, badge, popover
                  (React Portal, click-to-open)
lib/olk.ps1       the COM bridge: every Outlook operation, the token gate, and the
                  localization quirks (see the file header)
lib/watch.ps1     NewMailEx resident listener: pushes arrivals, self-heals after 30s
cordis.patch.yml  host row declaration
NOTICE            upstream attribution (MIT) and the pinned base commit
UPSTREAM.md       how this fork relates to upstream, and how to diff it
TESTCASES.md      49 test cases (pipeline / rendering / coexistence / routes /
                  click-to-open / link interception / known boundaries)
CHANGELOG.md      version narrative (why before what)
```

Diagnostic log: `%TEMP%\dsh-outlook-watch-debug.log` (rotated at 1 MB).

## Verification and tests

- 49 cases cover the whole chain ([TESTCASES.md](./TESTCASES.md)); the automated ones are all verified against a real Exchange mailbox: new mail → badge → clear, watcher killed and back within 30s, both confirmation-swap attacks blocked, route 405/413, log rotation, open-marks-read, the GBK code page regression, and chat link interception (v2.11.0 group H).
- The new mail-management actions (v2.12.0) are **not yet covered by a case** and have not run against a live Outlook — they were added on a machine without Outlook, so only their PowerShell parse and the tool registration are proven. Run `outlook_list_folders` first on a real machine.
- COM quirks (two-stage Jet/DASL filtering, the `Restrict` date format, `IncludeRecurrences` reporting MaxInt, PowerShell 5.1 stdin encoding) are all encoded in `lib/olk.ps1` and verified against a real Exchange mailbox.
