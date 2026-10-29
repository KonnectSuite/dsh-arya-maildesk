# dsh-outlook sidebar badge — test case set

Chain under test: **NewMailEx watcher → index.js state route → client 20s poll → badge render → popover → `outlook_check_new` clears the queue**.

Labels: `[auto]` can be checked with a script or tool; `[manual]` needs eyes on the browser; `[regression]` corresponds to a defect that really happened.

---

## A. Data path (watcher → state API)

| # | Case | Steps | Expected | Type |
|---|---|---|---|---|
| A1 | Single new mail | Send yourself one mail (through the `send_mail` confirmation flow) | Within 20s `GET /olk-notify/api/state` returns `pendingCount:1`, and pending holds the right entryId/subject/sender/received | [auto] |
| A2 | Batch arrival | One mail cc'ing you, or anything that fires NewMailEx with several EntryIds | One pending entry per EntryId, none lost | [auto] |
| A3 | `received` is the real time | Look at what A1 returns | Equals the mail's `ReceivedTime`, not the moment the event was handled (regression #7: it once used `Get-Date`) | [auto][regression] |
| A4 | Watcher self-heals | `taskkill` the watcher's powershell process | The debug log records the exit, a respawn happens within 30s, `hello` comes back, and the state route keeps answering meanwhile (with the stale unread value) | [auto] |
| A5 | Outlook not running | Quit Outlook, restart DSH | The watcher waits without crashing; once Outlook starts, `hello` arrives with the correct unread count | [manual] |
| A6 | Pending cap | Simulate more than 50 unconsumed arrivals | The queue keeps the newest 50 (`shift()` drops the oldest) and the count reads 50 | [auto] |
| A7 | `check_new` clears | With a non-empty queue, call `outlook_check_new` | Returns every pending entry with a matching `count`; a second call returns `count:0`; the badge disappears on the next poll | [auto] |
| A8 | Watcher off switch | Set `OLK_WATCH=0`, restart DSH | The watcher does not start, `state` returns `pendingCount:0` and `unread:null`, every other tool still works | [manual] |

## B. Badge rendering (client.js NewMailRow)

| # | Case | Steps | Expected | Type |
|---|---|---|---|---|
| B1 | Zero state hidden | `pendingCount:0` | Badge hidden; the button shows only the icon and "Outlook" | [manual] |
| B2 | Digit centred | pending=1 | "1" is centred both ways inside the 16px red circle (regression: `text-align`/`line-height` once pushed it off centre) | [manual][regression] |
| B3 | Two digits | pending≥10 | `min-width:16px` plus `padding:0 4px` widens the pill and the number stays centred | [manual] |
| B4 | Badge not clipped | Look at the badge while expanded | The part sticking out of the button (`top:-4px;right:-4px`) renders in full (regression: the button's `overflow:hidden` once cut the corner off) | [manual][regression] |
| B5 | Poll period | Send a mail, time it | The badge appears within 20s (`POLL_MS=20000`) | [manual] |
| B6 | Collapsed rail | Collapse the sidebar | The badge still shows at the top-right of the 36px circle and does not cover the icon | [manual] |
| B7 | Tooltip | Hover the button | The title lists the newest 8 as `[sender] subject`; with nothing new it reads "Outlook mail (nothing new)" | [manual] |
| B8 | ARIA | Inspect the button | `aria-label` is "Outlook, N new" (or "Outlook mail" at 0), and `aria-haspopup:dialog` is set | [manual] |
| B9 | Polling across a server restart | Watch while DSH restarts | A failed fetch is swallowed silently; the next poll after recovery works, with no error spam | [manual] |

## C. Popover

| # | Case | Steps | Expected | Type |
|---|---|---|---|---|
| C1 | Open | Click the button | The popover appears 8px to the right of the button, titled "📬 New mail (N)", listing the newest 10 (sender / subject / time) | [manual] |
| C2 | Close by outside click | Click anywhere outside | The popover is removed at once and the document listener is unbound | [manual] |
| C3 | Close by toggle | Click the button again | The popover closes (the button click does not also trigger C2) | [manual] |
| C4 | Empty list | Click with pending=0 | Shows "New mail (0)" and the hint "Click a mail to open it in Outlook; ask "anything new?" in chat to read and clear the queue" | [manual] |
| C5 | Unload cleanup | Disable the plugin or reload the page | Popover, button and the `dsh-outlook-styles` tag are all gone, with no DOM left behind | [manual] |

## D. Coexistence and style regressions (slot rewrite, v2.6.0)

| # | Case | Steps | Expected | Type |
|---|---|---|---|---|
| D1 | Ordering | Expand the sidebar | Outlook sits above Whale Bay, both above the settings area (`order` 10 < 90) | [manual] |
| D2 | Same height | Compare side by side with the settings button | Both are 42px tall with 12px radius and identical padding and font (regression: `flex:1` in the vertical flex column once stretched it) | [manual][regression] |
| D3 | Same hover | Hover all three buttons | The same `--dsw-alias-interactive-bg-hover` background (regression: once used a non-existent `--dsh-` variable plus an inline background that overrode the stylesheet) | [manual][regression] |
| D4 | With dsh-dock disabled | Disable dsh-dock, reload | The plugin's own `[data-slot]` stacking CSS keeps the entries vertical instead of collapsing them into a row | [manual] |
| D5 | Rail spacing | Compare while collapsed | The circle's margin is `8px 0 10px`, matching the settings button | [manual] |

## E. Server routes

| # | Case | Steps | Expected | Type |
|---|---|---|---|---|
| E1 | Method restriction | POST `/olk-notify/api/state` | 405 | [auto] |
| E2 | Auth | Request both routes without a cookie | 401 (the web layer intercepts uniformly) | [auto] |
| E3 | Beacon cap | POST a 200KB body to client-log | 413, connection closed, nothing written to the log | [auto] |
| E4 | Beacon noise | POST a non-JSON body | 204, ignored silently | [auto] |
| E5 | Log rotation | Write past 1MB and look | The debug log is truncated and rewritten; it does not grow without bound | [auto] |

## F. Known boundaries (by design — do not file these)

- At pending ≥ 100 the badge shows the real number (no "99+" clamp); a clamp would be a reasonable future change.
- Several browser tabs poll independently and share one server-side queue: once any session calls `check_new`, every page's badge goes to zero.
- A mail you send to yourself has a 10-20s NewMailEx delivery delay. That is Exchange behaviour, not a plugin defect.

## G. Click-to-open (v2.10.0)

| # | Case | Steps | Expected | Type |
|---|---|---|---|---|
| G1 | Bridge `open_mail` | Get an entryId from `search_mail`, run the bridge with `OLK_ACTION=open_mail` | A desktop Outlook inspector opens and the bridge returns `opened:true` with subject/sender (verified live) | [auto] |
| G2 | Tool | Call `outlook_open_mail` | Same as G1; purely local, no confirmation gate | [auto] |
| G3 | POST-only route | GET `/olk-notify/open?entryId=<valid id>` (simulating modifier-click navigation) | 405, the bridge is not touched (v2.11.0 removed the GET fallback; deliberate) | [auto] |
| G4 | Popover click | Click a mail inside the sidebar popover | POST succeeds, the entry leaves the list immediately and the badge decrements at once; the desktop opens the mail | [manual] |
| G5 | entryId allowlist | POST `{"entryId":"<script>"}` to `/olk-notify/open` | 400 / `{"ok":false,"error":"invalid entryId"}`, bridge untouched | [auto] |
| G6 | POST body cap | POST a body larger than 8KB | 413, connection closed | [auto] |
| G7 | Cross-origin | Send a foreign Origin header to all three olk-notify routes | 403 | [auto] |
| G8 | GBK code page regression | Run the bridge from a shell whose system code page is GBK | Parses with no mojibake (regression: a BOM-less UTF-8 file was once read as ANSI by PS 5.1 and the whole file failed to parse; fixed with a BOM and verified live) | [auto][regression] |
| G9 | Opening marks read | Open a mail whose `unread:true`, then re-check it with `search_mail` | The bridge returns `wasUnread:true`; the mail now reads `unread:false` (verified live). Note: an editor that rewrites the ps1 strips the BOM — restore it after any edit | [auto] |
| G10 | Badge decrements | Open one entry that is in pending (popover or chat link) | The host removes that entryId from `notify.pending` and decrements unread; the popover list drops it at once; other tabs agree within 20s | [auto][manual] |

## H. Chat link interception (v2.11.0)

| # | Case | Steps | Expected | Type |
|---|---|---|---|---|
| H1 | Plain left click | Click `[Open](…/olk-notify/open?entryId=…)` in chat | No new tab; a same-origin POST fires, the desktop opens the mail, and a 2.6s toast at the bottom says "Opened in Outlook: " plus the subject | [manual] |
| H2 | Any port | The link names a port other than the current one (e.g. :9999); plain left click | Still intercepted and posted to the page's own origin, exactly as H1 (this fixed dead links on custom-port deployments) | [manual] |
| H3 | Modifier navigation | Ctrl/⌘/Shift-click or middle-click the same link | Not intercepted; the browser navigates and gets 405 (POST-only; deliberate) | [manual] |
| H4 | Failure feedback | Click a link whose entryId no longer resolves (mail deleted) | The toast says "Could not open: " plus the error, and a beacon reports to client-log | [manual] |
| H5 | Cleanup is reversible | Disable or update the plugin (or reload) | The click listener, the toast timer and any leftover toast node are all removed | [auto] |
| H6 | Other links unaffected | Click any other markdown link in a chat message | Default behaviour; no `preventDefault` | [manual] |

---

## Helper scripts (for group A)

```powershell
# State assertion: expect pendingCount = N
$v = '<cookie>'; (Invoke-RestMethod -Headers @{Cookie=$v} 'http://127.0.0.1:3080/olk-notify/api/state').pendingCount

# Watcher liveness
(Get-Process powershell | Measure-Object).Count -ge 2  # bridge + watcher can coexist

# Trigger a real new mail: call outlook_send_mail in a session (needs user confirmation), then run the A1 assertion
```
