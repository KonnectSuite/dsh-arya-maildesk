/**
 * dsh-arya-maildesk — Outlook COM automation for DSH Harness (Windows).
 *
 * Bridges the local desktop Outlook client via a PowerShell COM dispatcher
 * (lib/olk.ps1). Read operations (account, search/read mail, calendar query,
 * people/room search, free/busy) run automatically.
 *
 * Write operations that leave the machine (sending mail, dispatching meeting
 * invites) use a TWO-CALL CONFIRMATION protocol: this host-plane plugin's
 * context cannot reach the session-scoped popup answerers, so an in-plugin
 * popup would fail closed every time. Instead the first call returns
 * `needs-confirmation` with the exact facts to confirm; the calling AGENT
 * asks the user through its own in-session popup (ask_user_question), and
 * only retries with `confirmed: true` after the user allows. That retry
 * mints the one-time OLK_CONFIRM token the bridge demands before sending.
 *
 * The bridge is invoked as
 *   powershell -NoProfile -Command "Invoke-Expression (Get-Content -Raw <olk.ps1>)"
 * because group policy commonly blocks running .ps1 files directly; the
 * action name travels in the OLK_ACTION environment variable, arguments
 * travel as ONE JSON object on STDIN (env vars cap near 32 KB), and the
 * result comes back as one JSON object on stdout.
 * @module dsh-arya-maildesk
 */

import { spawn } from 'node:child_process'
import { createHash, randomBytes, randomUUID } from 'node:crypto'
import { statSync, appendFileSync, writeFileSync, readFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { fileURLToPath } from 'node:url'
import { join, dirname } from 'node:path'

export const name = 'dsh-arya-maildesk'
export const inject = ['tools']

const HERE = dirname(fileURLToPath(import.meta.url))
const BRIDGE = join(HERE, 'lib', 'olk.ps1')
const TIMEOUT_MS = 120000

/** File-backed lifecycle log for the watcher (console.warn/error are not
 *  captured by the launcher log, which hid earlier failures). Lives in the
 *  OS temp dir, NOT the plugin install dir — a global/DSH install location
 *  can be read-only (then the log silently no-ops, which is fine for
 *  diagnostics). Truncates at 1 MB so it cannot grow unbounded. */
const WATCH_LOG = join(tmpdir(), 'dsh-arya-maildesk-watch-debug.log')
const WATCH_LOG_MAX = 1024 * 1024
function wlog(message) {
  try {
    let size = 0
    try { size = statSync(WATCH_LOG).size } catch { /* not created yet */ }
    if (size > WATCH_LOG_MAX) writeFileSync(WATCH_LOG, '')
    appendFileSync(WATCH_LOG, new Date().toISOString() + ' ' + message + '\n')
  } catch { /* diagnostics must never break the plugin */ }
}

/**
 * Serialize every bridge invocation: Outlook COM is single-threaded, and
 * concurrent PowerShell bridges racing the same COM apartment can throw
 * mid-operation. Each olk() call waits for the previous one to settle.
 */
const olkChain = { tail: Promise.resolve() }

/**
 * Run one bridge action and parse its JSON result.
 * Arguments travel over STDIN (not env vars) because Windows caps a single
 * environment variable near 32 KB — a large mail body silently overflowed
 * that cap in env transport, leaving the previous call's stale value behind.
 *
 * HUMAN-APPROVAL GATE (policy, enforced end to end):
 * Nothing that leaves the machine — sending mail, dispatching meeting
 * invites — may run without an explicit in-session human approval first.
 * `opts.confirmed` is set ONLY when the tool's caller passed confirmed:true
 * after the user allowed the two-call confirmation, and it mints a one-time
 * random OLK_CONFIRM token for the child process. The bridge independently
 * refuses every send path unless that token is present, so a future tool
 * (or a bug) that reaches the bridge without a human "allow" cannot send
 * anything.
 */
function olkOne(action, args = {}, opts = {}) {
  return new Promise((resolve) => {
    const env = {
      ...process.env,
      OLK_ACTION: action,
    }
    if (opts.confirmed) env.OLK_CONFIRM = randomBytes(24).toString('hex')
    const child = spawn('powershell', [
      '-NoProfile', '-Command',
      `Invoke-Expression (Get-Content -Raw -Encoding UTF8 '${BRIDGE.replace(/'/g, "''")}')`,
    ], {
      env,
      windowsHide: true,
      stdio: ['pipe', 'pipe', 'pipe'],
    })
    let stdout = ''
    let stderr = ''
    let settled = false
    const timer = setTimeout(() => {
      if (!settled) {
        settled = true
        child.kill()
        resolve({ status: 'error', error: 'bridge timed out after ' + TIMEOUT_MS + 'ms (is Outlook running?)' })
      }
    }, TIMEOUT_MS)
    child.stdin.on('error', () => { /* bridge died early; close handler reports */ })
    child.stdin.write(JSON.stringify(args))
    child.stdin.end()
    child.stdout.on('data', (d) => { stdout += d })
    child.stderr.on('data', (d) => { stderr += d })
    child.on('error', (e) => {
      if (settled) return
      settled = true
      clearTimeout(timer)
      resolve({ status: 'error', error: 'failed to start powershell: ' + e.message })
    })
    child.on('close', (code) => {
      if (settled) return
      settled = true
      clearTimeout(timer)
      try {
        resolve(JSON.parse(stdout))
      } catch {
        resolve({
          status: 'error',
          error: 'bridge output not JSON: ' + stdout.slice(0, 400),
          exitCode: code,
          stderr: stderr.slice(0, 400),
        })
      }
    })
  })
}

/** Queued bridge entry point (see olkChain above). */
function olk(action, args = {}, opts = {}) {
  // The fixture is the whole mailbox when MAILDESK_FIXTURE=1: the agent's
  // tools have to land on the same data the panel shows, or a host without
  // Outlook can render a mailbox but not answer a single mail question.
  if (fixture !== null) return Promise.resolve(fixtureCall(action, args))
  const run = () => olkOne(action, args, opts)
  const next = olkChain.tail.then(run, run)
  olkChain.tail = next.catch(() => { /* keep the queue alive after failures */ })
  return next
}

/** Stable-content SHA-256: keys sorted at every level so semantically equal
 *  argument objects hash identically regardless of property order. */
function contentHash(value) {
  const stable = (v) => {
    if (Array.isArray(v)) return v.map(stable)
    if (v !== null && typeof v === 'object') {
      const o = {}
      for (const k of Object.keys(v).sort()) o[k] = stable(v[k])
      return o
    }
    return v
  }
  return createHash('sha256').update(JSON.stringify(stable(value))).digest('hex')
}

/**
 * CONTENT-BINDING for the two-call confirmation (security hardening):
 * remembers the hash of the exact argument content carried by the LAST
 * needs-confirmation result per kind. A confirmed:true retry whose content
 * hash differs from what was last displayed is forced back through
 * needs-confirmation — a caller can never silently swap content between
 * what the user approved and what is dispatched. A successful dispatch
 * consumes the hash (single-use), so the same approval cannot be replayed
 * into a second send.
 */
const confirmations = new Map()

/** Run the confirmed dispatch and consume the approval on success, so a
 *  replayed confirmed:true call must pass through a fresh popup. */
async function dispatchConfirmed(kind, ...olkArgs) {
  const result = await olk(...olkArgs)
  if (result && result.status === 'ok') confirmations.delete(kind)
  return result
}

/** Validate local attachment paths and render them (name + human size) for
 *  the confirmation popup. Returns { error } when a file is missing. */
function attachmentLines(paths) {
  if (!Array.isArray(paths) || paths.length === 0) return { lines: '' }
  const lines = []
  for (const p of paths) {
    let st
    try { st = statSync(p) } catch { return { error: 'attachment not found: ' + p } }
    const kb = Math.max(1, Math.round(st.size / 1024))
    const base = String(p).replace(/^.*[\\/]/, '')
    lines.push(base + ' (' + kb + ' KB)')
  }
  return { lines: lines.join(', ') }
}

/** The `needs-confirmation` first-call result for the two-call protocol.
 *  The agent must relay `detail` to the user verbatim via an in-session
 *  popup and retry with confirmed:true only on an explicit allow. The
 *  popup is EDITABLE: the user may return modified fields; the agent then
 *  retries the call with the edited arguments (which re-enters
 *  needs-confirmation showing the edited content for one final allow), and
 *  only confirmed:true on content the user has seen dispatches. */
function needsConfirmation(kind, detail, editable = null) {
  return {
    status: 'needs-confirmation',
    kind,
    detail,
    editable,
    instruction: 'Show this detail to the user via the in-session popup (ask_user_question) and offer three choices: (a) send as-is, (b) edit — for each editable field give a "keep as-is" option AND let the user type a replacement in the free-text answer, (c) reject. If the user edited any field, retry the call with the edited arguments (without confirmed) so the edited content gets one final confirmation popup; then retry with confirmed:true only after the user allows the content they saw. Never set confirmed:true for content the user has not explicitly seen and allowed.',
  }
}

/* ---------------------------------------------------------------------------
 * PANEL DATA PATH (the GUI's half of the bridge)
 *
 * The panel lives in the client bundle, and a bundle's only outlet is HTTP:
 * the module loader hands `factory(require)` the kernel modules (react,
 * react-dom) and nothing that reaches this half, so `host.call` — a provider
 * for model-authored dynamic plugins — is not available there. A same-origin
 * POST route is the path the shipped client faces take, dsh-llamacpp's
 * control center included, and it is the one proven on this host.
 *
 * One route, an allowlist, and the same origin guard as the badge routes.
 * The panel may read and organize mail. It may NEVER send: send_draft,
 * reply_mail and every calendar write are absent from the allowlist, so the
 * only way mail leaves this machine is still the agent-side two-call
 * confirmation that mints OLK_CONFIRM.
 * ------------------------------------------------------------------------- */
const PANEL_ACTIONS = new Set([
  'account', 'list_folders', 'list_mail', 'read_mail',
  'mark_read', 'move_mail', 'delete_mail', 'open_mail',
])

/**
 * MAILDESK_FIXTURE=1 serves a synthetic mailbox instead of COM.
 *
 * It exists so the panel can be built and reviewed on a machine with no
 * Outlook, and so UI work never depends on a live mailbox. It is strictly
 * opt-in: a COM failure is surfaced as an error, never silently answered
 * with fake mail, because a panel showing invented messages would be worse
 * than a panel showing none.
 */
const FIXTURE_ON = process.env.MAILDESK_FIXTURE === '1'

function escapeHtml(value) {
  return String(value).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c])
}

function makeFixture() {
  let seed = 0x9e3779b9
  const rnd = () => {
    seed = (seed + 0x6d2b79f5) >>> 0
    let t = seed
    t = Math.imul(t ^ (t >>> 15), t | 1)
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61)
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296
  }
  const pick = (list) => list[Math.floor(rnd() * list.length)]
  const makeId = (kind, n) => (kind + '-' + n + '-' + Math.floor(rnd() * 1e9).toString(36) + '-fixture').padEnd(90, 'Q')

  const folders = []
  const folder = (name, parent) => {
    const path = parent === null ? name : parent.path + '/' + name
    const f = { entryId: makeId('folder', folders.length), name, path, store: 'Fromm Connect', depth: parent === null ? 0 : parent.depth + 1, childCount: 0, items: [] }
    if (parent !== null) parent.childCount++
    folders.push(f)
    return f
  }
  const inbox = folder('Inbox', null)
  const projects = folder('Projects', inbox)
  const vendors = folder('Vendors', inbox)
  const sent = folder('Sent Items', null)
  const drafts = folder('Drafts', null)
  const archive = folder('Archive', null)
  const a2026 = folder('2026', archive)
  const deleted = folder('Deleted Items', null)

  const people = [
    ['Dana Whitfield', 'dwhitfield@frommconnect.com'],
    ['Marcus Lee', 'mlee@frommconnect.com'],
    ['Sofia Alvarez', 'salvarez@frommconnect.com'],
    ['Nathan Brooks', 'nbrooks@frommconnect.com'],
    ['Melissa Chen', 'mchen@frommconnect.com'],
    ['Rockwell Automation Support', 'support@rockwellautomation.com'],
    ['APEX Industrial Supply', 'orders@apexindustrial.example'],
    ['Digi-Key Order Status', 'noreply@digikey.example'],
    ['UPS Freight', 'billing@ups.example'],
    ['Proofpoint Quarantine', 'quarantine@ppe-hosted.example'],
  ]
  const subjects = [
    'RE: PO 45-88213 - PanelView Plus 7 lead time',
    'Weekly service dispatch - 3 tickets closed',
    'RFQ: 1734-IB8S point I/O modules',
    'Invoice 77821 from APEX Industrial Supply',
    'Tracking update: shipment 1Z999AA10123456784',
    'Quarantine digest for frommconnect.com',
    'Line 4 retrofit - PLC schedule and cutoff',
    'Firmware 35.011 available for CompactLogix 5380',
    'Vendor onboarding: W-9 and insurance certificate',
    'RE: Overtime approval for week of the 26th',
    'Safety audit findings - action items due Friday',
    'Contractor badge paperwork - need a start date',
  ]
  const bodies = [
    'Morning,\n\nLead time on the PanelView Plus 7 replacement came back at 14 weeks from the distributor, so we need a decision this week if we want to hold the March window.\n\nThe alternate 2711P-T10C22D9P is in stock but it is a 10 inch unit and the cutout changes. I can put together a comparison this afternoon.\n\nDana',
    'The attached quote covers the 1734-IB8S modules plus the adapter. Pricing holds for 30 days.\n\nLet me know if you want me to split the shipment so the first two points can go out next week.\n\nMarcus',
    'Your invoice is attached. Payment terms are net 30 from the invoice date.',
    'Schedule for the retrofit:\n\n- Thursday: electrical walkdown, lockout tags\n- Friday: panel swap, 6 hour window\n- Monday: I/O checkout and point-to-point\n\nI need the line down by 5am Friday to make this work.',
    'This message was held for review. Release it from the quarantine digest if you recognize the sender.',
  ]
  const makeMail = (n, f) => {
    const [sender, smtp] = pick(people)
    const subject = pick(subjects)
    const body = pick(bodies)
    const unread = f.name === 'Inbox' || f.name === 'Projects' ? rnd() < 0.45 : rnd() < 0.15
    const m = {
      entryId: makeId('mail', n),
      subject,
      sender,
      senderSmtp: smtp,
      received: new Date(Date.now() - Math.round(n * 47 + rnd() * 30) * 60000).toISOString().slice(0, 16).replace('T', ' '),
      unread,
      hasAttach: rnd() < 0.3,
      body: body + '\n\n--\n' + sender + '\nFromm Electric | Fromm Connect',
      to: 'K Phung <kphung@frommconnect.com>',
      cc: rnd() < 0.25 ? 'Dana Whitfield <dwhitfield@frommconnect.com>' : '',
      bcc: '',
      folder: f.entryId,
      preview: body.replace(/\s+/g, ' ').slice(0, 160),
      attachments: rnd() < 0.3 ? ['quote-45-88213.pdf'] : [],
    }
    f.items.push(m)
    return m
  }
  let n = 0
  for (const f of [inbox, projects, vendors, sent, drafts, a2026, deleted]) {
    const count = f === inbox ? 48 : f === sent ? 12 : f === a2026 ? 22 : f === drafts ? 4 : f === deleted ? 8 : 9
    for (let k = 0; k < count; k++) makeMail(n++, f)
  }
  return { folders, folderById: new Map(folders.map((f) => [f.entryId, f])) }
}

// The sample mailbox. `MAILDESK_FIXTURE=1` makes it the provider for the whole
// session, and the panel's own "Preview with sample mail" button can also switch
// to it at runtime — a machine whose Outlook is the new WebView2 client has no
// COM surface at all, and without this the panel is an error message with nothing
// to look at. Either way it is never mistaken for a real mailbox: the header says
// SAMPLE MAIL and the tools answer from the same data.
let fixture = FIXTURE_ON ? makeFixture() : null
let sampleMail = fixture

/** Switch the provider to the sample mailbox (on) or back to live COM (off). */
function setSampleMail(on) {
  if (on) {
    if (sampleMail === null) sampleMail = makeFixture()
    fixture = sampleMail
  } else {
    fixture = FIXTURE_ON ? sampleMail : null
  }
  return fixture === null ? 'com' : 'fixture'
}

function fixtureCall(action, args = {}) {
  const ok = (extra) => ({ status: 'ok', ...extra })
  const resolve = (spec) => {
    if (!spec) return fixture.folders[0]
    return fixture.folderById.get(spec) || fixture.folders.find((f) => f.path === spec) || fixture.folders.find((f) => f.name.toLowerCase() === String(spec).toLowerCase()) || null
  }
  switch (action) {
    case 'account':
      return ok({ name: 'K Phung', smtp: 'kphung@frommconnect.com', fixture: true })
    case 'list_folders':
      return ok({
        folders: fixture.folders.map((f) => ({
          entryId: f.entryId, name: f.name, path: f.path, store: f.store, depth: f.depth,
          childCount: f.childCount, itemCount: f.items.length, unread: f.items.filter((m) => m.unread).length,
        })),
      })
    case 'list_mail': {
      const f = resolve(args.folder)
      if (f === null) return { status: 'error', error: "no folder matches '" + args.folder + "'" }
      let items = f.items.filter(() => true)
      if (args.unreadOnly === true) items = items.filter((m) => m.unread)
      const sort = String(args.sort || 'received').toLowerCase()
      const key = sort === 'subject' ? 'subject' : sort === 'sender' ? 'sender' : 'received'
      const desc = args.descending !== false
      items = items.slice().sort((a, b) => (a[key] < b[key] ? -1 : a[key] > b[key] ? 1 : 0) * (desc ? -1 : 1))
      const offset = Math.max(0, Number(args.offset) || 0)
      const limit = Math.min(200, Math.max(1, Number(args.limit) || 50))
      const page = items.slice(offset, offset + limit)
      return ok({
        total: items.length, offset, returned: page.length, folderPath: f.path,
        unreadInFolder: f.items.filter((m) => m.unread).length,
        mails: page.map((m) => ({ entryId: m.entryId, subject: m.subject, sender: m.sender, senderSmtp: m.senderSmtp, received: m.received, unread: m.unread, hasAttach: m.hasAttach, preview: args.preview ? m.preview : '' })),
      })
    }
    case 'read_mail': {
      for (const f of fixture.folders) {
        const m = f.items.find((x) => x.entryId === args.entryId)
        if (m) {
          const out = { subject: m.subject, sender: m.sender, senderSmtp: m.senderSmtp, received: m.received, to: m.to, cc: m.cc, bcc: m.bcc, body: m.body, attachments: m.attachments, unread: m.unread }
          if (args.format === 'html') {
            if (typeof m.htmlBody === 'string' && m.htmlBody !== '') {
              // A message we composed ourselves: return exactly what was sent, so
              // the reading pane is checked against the real body, not a mock.
              out.htmlBody = m.htmlBody
            } else {
              // Deliberately hostile fixture HTML: a tracking pixel, a script, an
              // onclick, a javascript: link and a Proofpoint-wrapped link. The
              // renderer must drop or rewrite every one of them.
              const paras = String(m.body).split('\n').filter(Boolean).map((line) => '<p>' + escapeHtml(line) + '</p>').join('')
              out.htmlBody = '<html><body><div style="font-family: Calibri">' + paras
                + '<p style="color:#666">-- <br>' + escapeHtml(m.sender) + ' | Fromm Connect</p>'
                + '<table border="1" cellpadding="4"><tr><th>PO</th><th>Qty</th></tr><tr><td>45-88213</td><td>12</td></tr></table>'
                + '<p><a href="https://urldefense.proofpoint.com/v2/url?u=https%3A%2F%2Frockwellautomation%2Ecom%2Fsupport&d=DwMF&c=abc&r=def&m=ghi&s=jkl&e=">Track this order</a></p>'
                + '<p><a href="javascript:alert(1)">unsubscribe</a></p>'
                + '<img src="https://tracker.example/pixel.gif?id=1" width="1" height="1">'
                + '<script>window.__pwned = true</script>'
                + '<div onclick="window.__clicked = true">click me</div>'
                + '</div></body></html>'
            }
            out.htmlTruncated = false
          }
          return ok(out)
        }
      }
      return { status: 'error', error: 'item not found: ' + args.entryId }
    }
    case 'mark_read': {
      const ids = new Set(args.entryIds || (args.entryId ? [args.entryId] : []))
      const read = args.read !== false
      let updated = 0
      for (const f of fixture.folders) for (const m of f.items) if (ids.has(m.entryId)) { m.unread = !read; updated++ }
      return ok({ updated, read, failed: [] })
    }
    case 'move_mail': {
      const dest = resolve(args.destFolder)
      if (dest === null) return { status: 'error', error: "no folder matches '" + args.destFolder + "'" }
      const ids = new Set(args.entryIds || (args.entryId ? [args.entryId] : []))
      const newEntryIds = []
      const failed = []
      for (const f of fixture.folders) {
        for (let i = f.items.length - 1; i >= 0; i--) {
          const m = f.items[i]
          if (!ids.has(m.entryId)) continue
          f.items.splice(i, 1)
          m.folder = dest.entryId
          // An EntryID does not survive Move(): the new item gets a new id,
          // so the caller's selection goes stale and it must re-read.
          m.entryId = ('mail-moved-' + newEntryIds.length + '-' + Math.random().toString(36).slice(2)).padEnd(90, 'Q')
          dest.items.push(m)
          newEntryIds.push(m.entryId)
        }
      }
      return ok({ moved: newEntryIds.length, destPath: dest.path, newEntryIds, failed })
    }
    case 'delete_mail': {
      const ids = new Set(args.entryIds || (args.entryId ? [args.entryId] : []))
      if (args.expected !== undefined && Number(args.expected) !== ids.size) {
        return { status: 'error', error: 'refusing to delete: caller expected ' + args.expected + ' items but supplied ' + ids.size + '. Re-read the folder and retry.' }
      }
      const deletedIds = []
      for (const f of fixture.folders) {
        for (let i = f.items.length - 1; i >= 0; i--) {
          if (ids.has(f.items[i].entryId)) { deletedIds.push(f.items[i].entryId); f.items.splice(i, 1) }
        }
      }
      return ok({ deleted: deletedIds.length, recoverable: true, deletedEntryIds: deletedIds, failed: [] })
    }
    case 'compose_draft': {
      const stamp = () => new Date().toISOString().slice(0, 16).replace('T', ' ')
      const folder = fixture.folders.find((x) => x.name === 'Drafts') || fixture.folders[0]
      const entryId = ('draft-' + Date.now().toString(36)).padEnd(90, 'Q')
      folder.items.unshift({
        entryId, subject: String(args.subject || '(no subject)'), sender: 'K Phung',
        senderSmtp: 'kphung@frommconnect.com', received: stamp(), unread: false, hasAttach: false,
        preview: '', to: String(args.to || ''), cc: String(args.cc || ''), bcc: '', body: String(args.htmlBody || ''), htmlBody: String(args.htmlBody || ''), attachments: [],
      })
      return ok({ draftEntryId: entryId, mode: String(args.mode || 'new'), message: 'Draft saved to the Drafts folder.' })
    }
    case 'compose_send': {
      const stamp = () => new Date().toISOString().slice(0, 16).replace('T', ' ')
      const folder = fixture.folders.find((x) => x.name === 'Sent Items') || fixture.folders[0]
      const entryId = ('sent-' + Date.now().toString(36)).padEnd(90, 'Q')
      fixture.sent = fixture.sent || []
      fixture.sent.push({
        entryId, to: String(args.to || ''), cc: String(args.cc || ''), subject: String(args.subject || ''),
        htmlBody: String(args.htmlBody || ''), mode: String(args.mode || 'new'), inReplyToEntryId: String(args.inReplyToEntryId || ''),
      })
      folder.items.unshift({
        entryId, subject: String(args.subject || '(no subject)'), sender: 'K Phung',
        senderSmtp: 'kphung@frommconnect.com', received: stamp(), unread: false, hasAttach: false,
        preview: '', to: String(args.to || ''), cc: String(args.cc || ''), bcc: '', body: String(args.htmlBody || ''), htmlBody: String(args.htmlBody || ''), attachments: [],
      })
      return ok({ sent: true, subject: String(args.subject || ''), to: String(args.to || ''), mode: String(args.mode || 'new') })
    }
    case 'open_mail':
      return ok({ opened: true, subject: '(fixture)' })
    default:
      return { status: 'error', error: 'action not available to the panel: ' + action }
  }
}

/**
 * The folder walk is the slowest call in the plugin: one COM read per folder,
 * so a real Exchange mailbox with a few hundred folders takes tens of seconds.
 * A tree changes when a person files a folder, not between two panel opens, so
 * a short cache keeps reopen instant without hiding a change for long. Keyed by
 * the arguments, because `counts` and `maxDepth` change the payload.
 */
const FOLDER_CACHE_MS = 5 * 60 * 1000
const folderCache = new Map()
// Anything that changes read state, a folder count or a move makes the cached
// tree stale; a person who just marked something read must not wait five
// minutes for the rail to agree.
const FOLDER_MUTATIONS = new Set(['mark_read', 'move_mail', 'delete_mail'])

/** One panel call: the fixture when explicitly enabled, COM otherwise. */
async function panelCall(action, args) {
  if (!PANEL_ACTIONS.has(action)) return { status: 'error', error: 'action not allowed: ' + action }
  if (fixture !== null) return fixtureCall(action, args)
  if (action !== 'list_folders') {
    const value = await olk(action, args)
    if (value !== null && value.status === 'ok' && FOLDER_MUTATIONS.has(action)) folderCache.clear()
    return value
  }
  const key = JSON.stringify(args || {})
  const hit = folderCache.get(key)
  if (hit !== undefined && Date.now() - hit.at < FOLDER_CACHE_MS) return hit.value
  const value = await olk(action, args)
  if (value !== null && value.status === 'ok') {
    folderCache.set(key, { at: Date.now(), value })
    if (folderCache.size > 64) folderCache.delete(folderCache.keys().next().value)
  }
  return value
}

/** NewMailEx watcher state shared by the badge route and check tool. */
const notify = {
  pending: [],   // new-mail events not yet fetched by the agent
  unread: null,  // last reported Inbox unread count
}

/**
 * Spawn (and baby-sit) the persistent PowerShell watcher (lib/watch.ps1).
 * Emits JSON lines on stdout; crashes restart with a 30s backoff. Killing
 * the child on fiber dispose is the only cleanup the process needs.
 */
function startWatcher(ctx) {
  if (process.env.OLK_WATCH === '0') return
  let child = null
  let stopped = false
  let restartTimer = null
  const WATCH = join(HERE, 'lib', 'watch.ps1')

  function spawnWatcher() {
    if (stopped) return
    child = spawn('powershell', [
      '-NoProfile', '-Command',
      `Invoke-Expression (Get-Content -Raw -Encoding UTF8 '${WATCH.replace(/'/g, "''")}')`,
    ], { windowsHide: true, stdio: ['ignore', 'pipe', 'pipe'] })
    wlog('spawned pid=' + child.pid)
    console.info('[dsh-arya-maildesk] watcher spawned (pid ' + child.pid + ')')
    child.on('error', (e) => wlog('spawn error: ' + e.message))
    let buf = ''
    child.stdout.on('data', (d) => {
      wlog('stdout chunk: ' + String(d).trim().slice(0, 200))
      buf += d
      let nl
      while ((nl = buf.indexOf('\n')) >= 0) {
        const line = buf.slice(0, nl).trim()
        buf = buf.slice(nl + 1)
        if (line === '') continue
        try {
          const ev = JSON.parse(line)
          if (ev.type === 'hello' || ev.type === 'tick') {
            notify.unread = ev.unread
            if (ev.type === 'hello') console.info('[dsh-arya-maildesk] watcher online, unread=' + ev.unread)
          } else if (ev.type === 'newmail') {
            notify.pending.push({ entryId: ev.entryId, subject: ev.subject, sender: ev.sender, received: ev.received })
            if (notify.pending.length > 50) notify.pending.shift()
            console.info('[dsh-arya-maildesk] new mail: [' + ev.sender + '] ' + ev.subject)
          } else if (ev.type === 'no-outlook') {
            // Classic Outlook not installed on this machine: the watcher
            // exits with code 3 right after this; do not restart it.
            console.info('[dsh-arya-maildesk] ' + (ev.message || 'classic Outlook not installed') + ' — badge watcher disabled (tools still available if Outlook appears later via a profile restart)')
          }
        } catch { /* partial line or non-JSON noise */ }
      }
    })
    child.stderr.on('data', (d) => { const s = String(d).trim(); if (s) { wlog('stderr: ' + s.slice(0, 200)); console.error('[dsh-arya-maildesk] watcher stderr: ' + s.slice(0, 200)) } })
    child.on('close', (code, signal) => {
      wlog('closed code=' + code + ' signal=' + signal + ' stopped=' + stopped)
      if (stopped) return
      if (code === 3) {
        // Watcher's "classic Outlook not installed" exit: no restart loop
        // on machines without desktop Outlook (public-distribution safety).
        console.warn('[dsh-arya-maildesk] watcher disabled: classic desktop Outlook not installed')
        return
      }
      console.warn('[dsh-arya-maildesk] watcher exited (code ' + code + '), restarting in 30s')
      restartTimer = setTimeout(spawnWatcher, 30000)
    })
  }

  spawnWatcher()
  ctx.effect(() => {
    // ctx.effect(fn): fn runs IMMEDIATELY (as setup) and its RETURN VALUE is
    // the disposer — the teardown must be returned, not written as the body
    // (writing it as the body killed the watcher 1ms after spawn).
    return () => {
      wlog('fiber dispose: stopping watcher')
      stopped = true
      if (restartTimer !== null) clearTimeout(restartTimer)
      if (child !== null) { try { child.kill() } catch { /* already gone */ } }
    }
  }, 'dsh-arya-maildesk: newmail watcher')
}

/**
 * Register the ten Outlook tools on `ctx.tools`.
 * @param {import('@deepseek-ai/cordis').Context} ctx - registrant context.
 */
export function apply(ctx) {
  if (process.platform !== 'win32') {
    console.info('[dsh-arya-maildesk] Windows only — skipping tool registration')
    return
  }

  // NewMailEx push watcher + badge state route (opt out with OLK_WATCH=0).
  startWatcher(ctx)
  // The draft store lives inside the route injection, but the compose tool is
  // registered outside it: hand the API across once the routes are mounted.
  let composeApi = null
  ctx.inject(['webServer'], (webCtx) => {
    // Same-origin guard for the two badge routes (M3): mail metadata must
    // not be readable by random pages probing localhost. Non-browser
    // clients (no Origin header) stay allowed — same trust as before.
    const sameOrigin = (req) => {
      const origin = req.headers.origin
      if (origin === undefined) return true
      try {
        return new URL(origin).host === req.headers.host
      } catch {
        return false
      }
    }
    ctx.effect(() => webCtx.webServer.register({
      kind: 'exact',
      path: '/olk-notify/api/state',
      handler: async (req, res) => {
        if (req.method !== 'GET') { res.statusCode = 405; res.end(); return }
        if (!sameOrigin(req)) { res.statusCode = 403; res.end(); return }
        res.setHeader('Content-Type', 'application/json; charset=utf-8')
        res.end(JSON.stringify({
          ok: true,
          pendingCount: notify.pending.length,
          pending: notify.pending,
          unread: notify.unread,
        }))
      },
    }), 'dsh-arya-maildesk: badge state route')
    // Click-to-open route: POST-only JSON, serving the sidebar popover and
    // the chat-link click interception (both fetch same-origin). The GET
    // navigation fallback / result page was REMOVED in v2.11.0 — chat
    // markdown links still point at this URL because the client intercepts
    // clicks by pathname; anything that navigates here directly
    // (modifier/middle click) gets a 405.
    const ENTRY_ID_RE = /^[A-Za-z0-9+/=_-]{20,512}$/
    const readEntryId = (q) => {
      const id = String(q || '')
      return ENTRY_ID_RE.test(id) ? id : null
    }
    ctx.effect(() => webCtx.webServer.register({
      kind: 'exact',
      path: '/olk-notify/open',
      handler: async (req, res) => {
        if (req.method !== 'POST') { res.statusCode = 405; res.end(); return }
        if (!sameOrigin(req)) { res.statusCode = 403; res.end(); return }
        let entryId = null
        let body = ''
        let overflow = false
        for await (const chunk of req) {
          body += chunk
          if (body.length > 8192) { overflow = true; break }
        }
        if (overflow) { res.statusCode = 413; res.end(); return }
        try { entryId = readEntryId(JSON.parse(body).entryId) } catch { entryId = null }
        if (entryId === null) {
          res.statusCode = 400
          res.setHeader('Content-Type', 'application/json; charset=utf-8')
          res.end(JSON.stringify({ ok: false, error: 'invalid entryId' }))
          return
        }
        const r = await olk('open_mail', { entryId })
        if (r.status === 'ok') {
          // Reading by opening: drop the mail from the pending queue (badge
          // count = pending length, so the sidebar chip decrements on the
          // next poll) and keep the unread total in step when the mail was
          // previously unread.
          const i = notify.pending.findIndex((m) => m.entryId === entryId)
          if (i >= 0) notify.pending.splice(i, 1)
          if (r.wasUnread && typeof notify.unread === 'number' && notify.unread > 0) notify.unread--
        }
        res.setHeader('Content-Type', 'application/json; charset=utf-8')
        res.statusCode = r.status === 'ok' ? 200 : 500
        res.end(JSON.stringify({ ok: r.status === 'ok', subject: r.subject, sender: r.sender, error: r.error }))
      },
    }), 'dsh-arya-maildesk: open-mail route')
    ctx.effect(() => webCtx.webServer.register({
      kind: 'exact',
      path: '/olk-notify/api/client-log',
      handler: async (req, res) => {
        if (req.method !== 'POST') { res.statusCode = 405; res.end(); return }
        if (!sameOrigin(req)) { res.statusCode = 403; res.end(); return }
        let body = ''
        let overflow = false
        for await (const chunk of req) {
          body += chunk
          if (body.length > 65536) { overflow = true; break } // beacon cap: 64 KB
        }
        if (overflow) { res.statusCode = 413; res.end(); return }
        try {
          const data = JSON.parse(body)
          wlog('client: ' + String(data.message || '').slice(0, 300))
        } catch { /* ignore malformed beacons */ }
        res.statusCode = 204
        res.end()
      },
    }), 'dsh-arya-maildesk: client beacon route')

    // ---- panel routes (M1) ---------------------------------------------
    // Health tells the panel which provider is answering, so fixture data is
    // labelled in the UI instead of being mistaken for a real mailbox.
    ctx.effect(() => webCtx.webServer.register({
      kind: 'exact',
      path: '/maildesk/api/health',
      handler: async (req, res) => {
        if (req.method !== 'GET') { res.statusCode = 405; res.end(); return }
        if (!sameOrigin(req)) { res.statusCode = 403; res.end(); return }
        res.setHeader('Content-Type', 'application/json; charset=utf-8')
        res.end(JSON.stringify({
          ok: true,
          provider: fixture !== null ? 'fixture' : 'com',
          // `sample` is true only when this session's provider is the sample
          // mailbox, so the panel can offer "back to live mail" instead of
          // pretending MAILDESK_FIXTURE=1 was set.
          sample: fixture !== null && !FIXTURE_ON,
        }))
      },
    }), 'dsh-arya-maildesk: panel health route')

    // Switch provider between the sample mailbox and live COM.
    ctx.effect(() => webCtx.webServer.register({
      kind: 'exact',
      path: '/maildesk/api/fixture',
      handler: async (req, res) => {
        if (req.method !== 'POST') { res.statusCode = 405; res.end(); return }
        if (!sameOrigin(req)) { res.statusCode = 403; res.end(); return }
        const parsed = await readJsonBody(req, 4096)
        if (parsed.error !== undefined) {
          res.statusCode = 400
          res.setHeader('Content-Type', 'application/json; charset=utf-8')
          res.end(JSON.stringify({ ok: false, error: parsed.error }))
          return
        }
        const provider = setSampleMail(parsed.value !== null && typeof parsed.value === 'object' && parsed.value.on === true)
        res.setHeader('Content-Type', 'application/json; charset=utf-8')
        res.end(JSON.stringify({ ok: true, provider, sample: provider === 'fixture' && !FIXTURE_ON }))
      },
    }), 'dsh-arya-maildesk: sample mailbox route')

    const readJsonBody = async (req, cap) => {
      let body = ''
      for await (const chunk of req) {
        body += chunk
        if (body.length > cap) return { error: 'request body too large' }
      }
      try {
        return { value: JSON.parse(body) }
      } catch {
        return { error: 'invalid JSON body' }
      }
    }

    ctx.effect(() => webCtx.webServer.register({
      kind: 'exact',
      path: '/maildesk/api/call',
      handler: async (req, res) => {
        res.setHeader('Content-Type', 'application/json; charset=utf-8')
        if (req.method !== 'POST') {
          res.statusCode = 405
          res.end(JSON.stringify({ ok: false, error: 'POST only' }))
          return
        }
        if (!sameOrigin(req)) {
          res.statusCode = 403
          res.end(JSON.stringify({ ok: false, error: 'cross-origin request refused' }))
          return
        }
        const body = await readJsonBody(req, 256 * 1024)
        if (body.error !== undefined) {
          res.statusCode = 400
          res.end(JSON.stringify({ ok: false, error: body.error }))
          return
        }
        const payload = body.value !== null && typeof body.value === 'object' ? body.value : {}
        const action = String(payload.action || '')
        if (!PANEL_ACTIONS.has(action)) {
          res.statusCode = 400
          res.end(JSON.stringify({ ok: false, error: 'action not allowed: ' + action }))
          return
        }
        try {
          const args = payload.args !== null && typeof payload.args === 'object' ? payload.args : {}
          const result = await panelCall(action, args)
          // Always 200: the bridge reports its own failures as status:'error',
          // and an HTTP status would make "no Outlook installed" look like a
          // transport bug in the panel.
          res.end(JSON.stringify({ ok: true, result }))
        } catch (error) {
          res.end(JSON.stringify({ ok: false, error: (error && error.message) || String(error) }))
        }
      },
    }), 'dsh-arya-maildesk: panel call route')

    // ---- in-panel assistant --------------------------------------------
    // Ask runs a REAL agent turn in a session of its own: the harness stays the
    // only thing that talks to a model, the maildesk_* tools are already in its
    // registry, and anything that leaves the machine still passes the same human
    // gates as the chat. The panel only needs "send this" and "what came back".
    // The assistant's session id outlives a reload but not a fresh machine.
    // The env override exists so tests never read the user's own state.
    const ASK_STATE_FILE = process.env.MAILDESK_STATE_FILE || join(tmpdir(), 'dsh-arya-maildesk-ask.json')

    const readAskSession = () => {
      try {
        const saved = JSON.parse(readFileSync(ASK_STATE_FILE, 'utf8'))
        return saved !== null && typeof saved.sessionId === 'string' && saved.sessionId !== '' ? saved.sessionId : null
      } catch {
        return null
      }
    }

    const saveAskSession = (sessionId) => {
      try {
        writeFileSync(ASK_STATE_FILE, JSON.stringify({ sessionId }), 'utf8')
      } catch { /* the assistant still works for this process */ }
    }

    const askState = { sessionId: readAskSession(), asked: new Set() }

    const createAskSession = async (controller) => {
      const created = await controller.create({})
      askState.sessionId = created.sessionId
      saveAskSession(created.sessionId)
      // A recognizable name in the sidebar beats five rows called "desktop";
      // a host without renaming still works.
      try {
        await controller.rename({ sessionId: created.sessionId, title: 'Mail assistant' })
      } catch { /* cosmetic */ }
    }

    // A session's log also holds the harness's own context injections (runtime
    // context, skill catalogue, clock) as user-role messages. The panel must
    // show the person's words and the answer, nothing else.
    const looksLikeHarnessContext = (text) =>
      /^Current runtime context\./.test(text) ||
      text.includes('<system-reminder>') ||
      /^Time sampled while preparing turn/.test(text)

    const textOfMessage = (message) => {
      if (message === null || message === undefined) return ''
      if (typeof message === 'string') return message.trim()
      const blocks = Array.isArray(message.content) ? message.content : []
      return blocks
        .filter((b) => b !== null && typeof b === 'object' && b.type === 'text' && typeof b.text === 'string')
        .map((b) => b.text)
        .join('\n')
        .trim()
    }

    ctx.effect(() => webCtx.webServer.register({
      kind: 'exact',
      path: '/maildesk/api/ask',
      handler: async (req, res) => {
        res.setHeader('Content-Type', 'application/json; charset=utf-8')
        if (req.method !== 'POST' && req.method !== 'GET') {
          res.statusCode = 405
          res.end(JSON.stringify({ ok: false, error: 'GET or POST only' }))
          return
        }
        if (!sameOrigin(req)) {
          res.statusCode = 403
          res.end(JSON.stringify({ ok: false, error: 'cross-origin request refused' }))
          return
        }
        const controller = ctx.get('sessionController')
        if (controller === undefined) {
          res.statusCode = 501
          res.end(JSON.stringify({ ok: false, error: 'this host mounts no session controller, so the in-panel assistant is unavailable' }))
          return
        }
        try {
          if (req.method === 'GET') {
            const sessionId = askState.sessionId
            if (sessionId === null) {
              res.end(JSON.stringify({ ok: true, sessionId: null, messages: [] }))
              return
            }
            // History is addressed by a Session address, not a bare id: the
            // events resolver reads address.sessionId.
            const page = await controller.events(
              { address: { kind: 'session', sessionId }, maxMessages: 80 },
              new AbortController().signal,
            )
            const records = page !== null && page !== undefined && Array.isArray(page.records) ? page.records : []
            const messages = []
            for (const record of records) {
              const event = record !== null && record !== undefined ? record.event : null
              if (event === null || event === undefined) continue
              // Showing the tools makes the answer auditable: a person can see
              // it read the mailbox instead of trusting a sentence.
              if (event.type === 'tool/call') {
                const data = event.data !== null && typeof event.data === 'object' ? event.data : {}
                const name = typeof data.name === 'string' ? data.name
                  : (typeof data.tool === 'string' ? data.tool
                    : (typeof data.toolName === 'string' ? data.toolName : ''))
                if (name !== '') messages.push({ seq: event.seq, role: 'tool', text: name })
                continue
              }
              if (event.type !== 'user/message' && event.type !== 'assistant/message') continue
              // assistant/message carries data.message; user/message carries the
              // content on the event data itself in some builds, so accept both.
              const text = textOfMessage(event.data && (event.data.message !== undefined ? event.data.message : event.data))
              if (text === '') continue
              if (event.type === 'user/message') {
                if (looksLikeHarnessContext(text)) continue
                const source = event.data && event.data.source
                if (source !== null && typeof source === 'object' && typeof source.rpcId === 'string' && !askState.asked.has(source.rpcId)) continue
              }
              messages.push({ seq: event.seq, role: event.type === 'user/message' ? 'user' : 'assistant', text })
            }
            res.end(JSON.stringify({ ok: true, sessionId, messages }))
            return
          }
          const body = await readJsonBody(req, 64 * 1024)
          if (body.error !== undefined) {
            res.statusCode = 400
            res.end(JSON.stringify({ ok: false, error: body.error }))
            return
          }
          const payload = body.value !== null && typeof body.value === 'object' ? body.value : {}
          const text = String(payload.text || '').trim()
          if (text === '') {
            res.statusCode = 400
            res.end(JSON.stringify({ ok: false, error: 'text is required' }))
            return
          }
          if (askState.sessionId === null) {
            await createAskSession(controller)
          }
          // "this one" is ambiguous to a tool that wants an entryId: hand the
          // agent the panel's own context along with the words.
          const context = [
            payload.folderPath ? 'folder ' + String(payload.folderPath) : '',
            payload.selected ? 'selected mail "' + String(payload.selected) + '"' : '',
          ].filter(Boolean).join(', ')
          const composed = context === '' ? text : text + '\n\n(mail panel · ' + context + ')'
          const requestId = randomUUID()
          askState.asked.add(requestId)
          // The controller owns the durable source; its second argument is a
          // cancellation signal, not prompt metadata.
          const send = () => controller.prompt(
            { sessionId: askState.sessionId, requestId, content: [{ type: 'text', text: composed }], mode: 'queue' },
            new AbortController().signal,
          )
          try {
            await send()
          } catch (error) {
            // A session id that outlived its log (a restart, a cleared profile)
            // is worth exactly one retry on a fresh session.
            if (!/not.?found|unknown session|no such/i.test(String(error && error.message))) throw error
            askState.sessionId = null
            await createAskSession(controller)
            await send()
          }
          res.end(JSON.stringify({ ok: true, sessionId: askState.sessionId }))
        } catch (error) {
          res.statusCode = 500
          res.end(JSON.stringify({ ok: false, error: (error && error.message) || String(error) }))
        }
      },
    }), 'dsh-arya-maildesk: in-panel assistant route')

    // ---- compose routes (M3) -------------------------------------------
    // A draft is HARNESS DATA keyed by (account, mode, thread), so opening a
    // reply twice returns the draft you already wrote instead of stacking a
    // second item in Outlook's Drafts folder. Outlook is the transport at send
    // time and the source of truth for received mail — never where the draft
    // lives. "Save to Outlook Drafts" is the explicit escape hatch, and it
    // updates that same item on every later save.
    const drafts = new Map()
    const draftKeyOf = (account, mode, thread) => [account || 'default', mode || 'new', thread || ''].join('|')
    const asList = (v) => (Array.isArray(v) ? v : [])
    const publicDraft = (d) => ({
      id: d.id, mode: d.mode, inReplyTo: d.inReplyTo, inReplyToSubject: d.inReplyToSubject,
      to: d.to, cc: d.cc, bcc: d.bcc, subject: d.subject, bodyHtml: d.bodyHtml,
      quoteHtml: d.quoteHtml, outlookEntryId: d.outlookEntryId,
      createdAt: d.createdAt, updatedAt: d.updatedAt,
    })
    const findDraft = (id) => {
      for (const d of drafts.values()) if (d.id === id) return d
      return null
    }
    const parseAddrList = (text) => String(text || '').split(';').map((x) => x.trim()).filter(Boolean).map((x) => {
      const m = /^(.*)<([^>]+)>$/.exec(x)
      if (m) return { name: m[1].trim().replace(/^["']|["']$/g, ''), email: m[2].trim() }
      return { name: '', email: x }
    })
    const formatAddrList = (list) => asList(list)
      .map((a) => (a && a.email ? (a.name ? '"' + a.name + '" <' + a.email + '>' : a.email) : ''))
      .filter(Boolean).join('; ')
    /** A quote is copied into an OUTBOUND message, so the reading pane's
     *  allowlist is not enough here: the active parts are stripped before the
     *  text can leave the machine. */
    function quoteSafe(html) {
      return String(html || '')
        .replace(/<(script|style|iframe|object|embed|link|meta|form|title)\b[\s\S]*?<\/\1>/gi, '')
        .replace(/<(script|style|iframe|object|embed|link|meta|form|title|img|input|button)\b[^>]*>/gi, '')
        .replace(/\son[a-z]+\s*=\s*(?:"[^"]*"|'[^']*'|[^\s>]+)/gi, '')
        .replace(/(href|src)\s*=\s*(?:"|')?\s*javascript:[^"'>\s]*(?:"|')?/gi, '')
    }
    async function openDraft(args) {
      const mode = ['new', 'reply', 'replyAll', 'forward'].includes(String(args.mode)) ? String(args.mode) : 'new'
      const thread = args.inReplyTo ? String(args.inReplyTo) : ''
      const account = String(args.account || '')
      const key = draftKeyOf(account, mode, thread)
      const hit = drafts.get(key)
      if (hit) return { draft: publicDraft(hit), reused: true }
      const draft = {
        id: 'draft-' + Date.now().toString(36) + '-' + Math.random().toString(36).slice(2, 8),
        account, mode, inReplyTo: thread, inReplyToSubject: '', to: [], cc: [], bcc: [],
        subject: '', bodyHtml: '', quoteHtml: '', outlookEntryId: null,
        createdAt: Date.now(), updatedAt: Date.now(),
      }
      if (thread !== '') {
        let src = null
        try { src = await panelCall('read_mail', { entryId: thread, maxChars: 20000, format: 'html' }) } catch { src = null }
        if (src && src.status === 'ok') {
          draft.inReplyToSubject = String(src.subject || '')
          const prefix = mode === 'forward' ? 'Fwd: ' : 'Re: '
          const base = draft.inReplyToSubject
          draft.subject = base.toLowerCase().startsWith(prefix.toLowerCase()) ? base : prefix + base
          // SenderEmailAddress is a usable SMTP address on some mailboxes and an
          // EX:/O=... directory name on a corporate Exchange one. Only an
          // address can become a recipient chip; without one the reply is left
          // for Outlook's own Reply() to address, which is why an empty To is
          // allowed for a reply (see sendArgs/sendDraft).
          const senderSmtp = String(src.senderSmtp || '')
          const sender = senderSmtp.includes('@')
            ? { name: String(src.sender || ''), email: senderSmtp }
            : null
          let self = ''
          try {
            const acct = await panelCall('account', {})
            if (acct && acct.status === 'ok') self = String(acct.smtp || '').toLowerCase()
          } catch { self = '' }
          if (mode === 'reply' && sender) draft.to = [sender]
          if (mode === 'replyAll') {
            const seen = new Set([self])
            const push = (a) => {
              if (!a || !a.email) return
              const k = String(a.email).toLowerCase()
              if (seen.has(k)) return
              seen.add(k)
              draft.to.push(a)
            }
            push(sender)
            for (const a of parseAddrList(src.to)) push(a)
            for (const a of parseAddrList(src.cc)) {
              if (!a.email) continue
              const k = String(a.email).toLowerCase()
              if (seen.has(k)) continue
              seen.add(k)
              draft.cc.push(a)
            }
          }
          const html = typeof src.htmlBody === 'string' && src.htmlBody !== ''
            ? src.htmlBody
            : '<p>' + String(src.body || '').replace(/&/g, '&amp;').replace(/</g, '&lt;')
                .split(/\n{2,}/).map((p) => p.replace(/\n/g, '<br>')).join('</p><p>') + '</p>'
          const who = sender
            ? (sender.name ? sender.name + ' &lt;' + sender.email + '&gt;' : sender.email)
            : (String(src.sender || '') || 'the sender')
          draft.quoteHtml = '<div class="maildesk-quote"><p>On ' + String(src.received || '') + ', ' + who + ' wrote:</p><blockquote>' + quoteSafe(html) + '</blockquote></div>'
        }
      }
      drafts.set(key, draft)
      return { draft: publicDraft(draft), reused: false }
    }
    function patchDraft(draft, patch) {
      for (const f of ['to', 'cc', 'bcc']) {
        if (patch[f] !== undefined) {
          draft[f] = asList(patch[f])
            .map((a) => ({ name: String((a && a.name) || ''), email: String((a && a.email) || '') }))
            .filter((a) => a.email !== '')
        }
      }
      if (patch.subject !== undefined) draft.subject = String(patch.subject)
      if (patch.bodyHtml !== undefined) draft.bodyHtml = String(patch.bodyHtml)
      draft.updatedAt = Date.now()
      return draft
    }
    const sendArgs = (d) => ({
      mode: d.mode,
      inReplyToEntryId: d.inReplyTo || undefined,
      to: formatAddrList(d.to),
      cc: formatAddrList(d.cc),
      bcc: formatAddrList(d.bcc),
      subject: d.subject,
      htmlBody: String(d.bodyHtml || '') + String(d.quoteHtml || ''),
    })
    async function sendDraft(id, confirmHash) {
      const d = findDraft(String(id || ''))
      if (d === null) return { status: 'error', error: 'unknown draft: ' + id }
      const args = sendArgs(d)
      // A reply may legitimately carry no explicit recipient list: Outlook's
      // own Reply()/ReplyAll() already addressed it from the source item, and
      // an empty To must not overwrite that.
      const replyMode = args.mode === 'reply' || args.mode === 'replyAll' || args.mode === 'forward'
      if (!args.to && !replyMode) return { status: 'error', error: 'a send needs at least one recipient' }
      const hash = contentHash(args)
      if (String(confirmHash || '') !== hash) {
        return {
          status: 'needs-confirmation',
          hash,
          preview: { mode: args.mode, to: args.to, cc: args.cc, bcc: args.bcc, subject: args.subject, htmlBody: args.htmlBody },
          instruction: 'Show the recipient, subject and body to the user and retry with the same hash only after they allow this exact content.',
        }
      }
      let result
      if (fixture !== null) {
        result = fixtureCall('compose_send', args)
      } else {
        // The human clicked Send and then allowed the exact content above, so
        // the host mints the one-time token the bridge demands.
        result = await olk('compose_send', args, { confirmed: true })
      }
      if (result && result.status === 'ok') {
        drafts.delete(draftKeyOf(d.account, d.mode, d.inReplyTo))
        result.draftId = d.id
      }
      return result
    }
    const draftRoute = async (req, res) => {
      res.setHeader('Content-Type', 'application/json; charset=utf-8')
      if (req.method !== 'POST') { res.statusCode = 405; res.end(JSON.stringify({ ok: false, error: 'POST only' })); return }
      if (!sameOrigin(req)) { res.statusCode = 403; res.end(JSON.stringify({ ok: false, error: 'cross-origin request refused' })); return }
      const body = await readJsonBody(req, 256 * 1024)
      if (body.error !== undefined) { res.statusCode = 400; res.end(JSON.stringify({ ok: false, error: body.error })); return }
      const payload = body.value !== null && typeof body.value === 'object' ? body.value : {}
      const op = String(payload.op || '')
      try {
        if (op === 'open') { res.end(JSON.stringify({ ok: true, ...(await openDraft(payload)) })); return }
        if (op === 'get') {
          const d = findDraft(String(payload.draftId || ''))
          if (d === null) { res.statusCode = 404; res.end(JSON.stringify({ ok: false, error: 'unknown draft: ' + payload.draftId })); return }
          res.end(JSON.stringify({ ok: true, draft: publicDraft(d) }))
          return
        }
        if (op === 'save') {
          const d = findDraft(String(payload.draftId || ''))
          if (d === null) { res.statusCode = 404; res.end(JSON.stringify({ ok: false, error: 'unknown draft: ' + payload.draftId })); return }
          patchDraft(d, payload.patch !== null && typeof payload.patch === 'object' ? payload.patch : {})
          res.end(JSON.stringify({ ok: true, draft: publicDraft(d) }))
          return
        }
        if (op === 'discard') {
          const d = findDraft(String(payload.draftId || ''))
          if (d !== null) drafts.delete(draftKeyOf(d.account, d.mode, d.inReplyTo))
          res.end(JSON.stringify({ ok: true, discarded: d !== null }))
          return
        }
        if (op === 'list') {
          res.end(JSON.stringify({ ok: true, drafts: [...drafts.values()].map(publicDraft) }))
          return
        }
        res.statusCode = 400
        res.end(JSON.stringify({ ok: false, error: 'unknown draft op: ' + op }))
      } catch (error) {
        res.end(JSON.stringify({ ok: false, error: (error && error.message) || String(error) }))
      }
    }
    const sendRoute = async (req, res) => {
      res.setHeader('Content-Type', 'application/json; charset=utf-8')
      if (req.method !== 'POST') { res.statusCode = 405; res.end(JSON.stringify({ ok: false, error: 'POST only' })); return }
      if (!sameOrigin(req)) { res.statusCode = 403; res.end(JSON.stringify({ ok: false, error: 'cross-origin request refused' })); return }
      const body = await readJsonBody(req, 256 * 1024)
      if (body.error !== undefined) { res.statusCode = 400; res.end(JSON.stringify({ ok: false, error: body.error })); return }
      const payload = body.value !== null && typeof body.value === 'object' ? body.value : {}
      try {
        const result = await sendDraft(payload.draftId, payload.confirmHash)
        // 200 either way: needs-confirmation is a step in a flow, not a fault.
        res.end(JSON.stringify({ ok: result.status === 'ok', result }))
      } catch (error) {
        res.end(JSON.stringify({ ok: false, error: (error && error.message) || String(error) }))
      }
    }
    ctx.effect(() => webCtx.webServer.register({ kind: 'exact', path: '/maildesk/api/draft', handler: draftRoute }), 'dsh-arya-maildesk: draft route')
    ctx.effect(() => webCtx.webServer.register({ kind: 'exact', path: '/maildesk/api/send', handler: sendRoute }), 'dsh-arya-maildesk: send route')
    composeApi = { drafts, draftKeyOf, findDraft, publicDraft, openDraft, patchDraft, parseAddrList, formatAddrList, sendArgs }
  })

  ctx.tools.register({
    name: 'maildesk_account',
    description: 'Get the currently configured primary Outlook account name and SMTP address. Read-only.',
    parameters: { type: 'object', additionalProperties: true, properties: {} },
    output: {
      schema: { type: 'object', additionalProperties: true },
      render: (_a, v) => [{ type: 'text', text: JSON.stringify(v) }],
    },
    async execute() { return olk('account') },
    presentCall: () => ({ card: 'generic', title: 'Outlook account', kind: 'other', rawInput: {} }),
  })

  ctx.tools.register({
    name: 'maildesk_open_mail',
    description: 'Open one email in a desktop Outlook inspector window by entryId (from maildesk_search_mail / maildesk_check_new). Local action: nothing leaves the machine, no confirmation needed. When presenting search results in chat, prefer composing a clickable markdown link per mail, e.g. [Open in Outlook](http://127.0.0.1:3080/olk-notify/open?entryId=<entryId>) — the user can click it to open the mail directly.',
    parameters: {
      type: 'object',
      additionalProperties: true,
      properties: {
        entryId: { type: 'string', description: 'entryId from maildesk_search_mail or maildesk_check_new.' },
      },
      required: ['entryId'],
    },
    output: {
      schema: { type: 'object', additionalProperties: true },
      render: (_a, v) => [{ type: 'text', text: JSON.stringify(v) }],
    },
    async execute(a) { return olk('open_mail', a) },
    presentCall: (a) => ({ card: 'generic', title: 'Open Outlook mail', kind: 'other', rawInput: { entryId: a.entryId } }),
  })

  ctx.tools.register({
    name: 'maildesk_search_mail',
    description: 'Search emails in the local Outlook mailbox (read-only). Default: last 14 days of Inbox, newest first. Supports arbitrary folder paths ("Inbox/ProjectA", "Sent Items"), recipient-side matching (to), and exact date ranges (dateFrom/dateTo override days). Returns entryId values usable with maildesk_read_mail. Tip: when listing results to the user, append a clickable open link per mail: [Open](http://127.0.0.1:3080/olk-notify/open?entryId=<entryId>).',
    parameters: {
      type: 'object',
      additionalProperties: true,
      properties: {
        days: { type: 'number', description: 'How many days back to search. Default 14. Ignored when dateFrom/dateTo is given.' },
        folder: { type: 'string', description: 'Folder to search: "Inbox" (default), "Sent Items", "Drafts", or any path like "Inbox/ProjectA" (segments separated by / or \\). Paths are walked from the primary mailbox root.' },
        from: { type: 'string', description: 'Filter: sender name contains this text.' },
        to: { type: 'string', description: 'Filter: recipient display name contains this text (recipient-side search).' },
        subjectContains: { type: 'string', description: 'Filter: subject contains this text.' },
        dateFrom: { type: 'string', description: 'Exact range start, "yyyy-MM-dd" or "yyyy-MM-dd HH:mm". Overrides days.' },
        dateTo: { type: 'string', description: 'Exact range end, "yyyy-MM-dd" (extends to 23:59) or "yyyy-MM-dd HH:mm". Overrides days.' },
        limit: { type: 'number', description: 'Max results. Default 10.' },
      },
    },
    output: {
      schema: { type: 'object', additionalProperties: true },
      render: (_a, v) => [{ type: 'text', text: JSON.stringify(v) }],
    },
    async execute(a) { return olk('search_mail', a) },
    presentCall: (a) => ({ card: 'generic', title: 'Search Outlook mail', kind: 'other', rawInput: a }),
  })

  ctx.tools.register({
    name: 'maildesk_read_mail',
    description: 'Read one full email by its entryId (obtained from maildesk_search_mail). Returns subject, sender, recipients, body text and attachment names. Read-only.',
    parameters: {
      type: 'object',
      additionalProperties: true,
      properties: {
        entryId: { type: 'string', description: 'entryId from maildesk_search_mail.' },
        maxChars: { type: 'number', description: 'Max body characters. Default 8000.' },
        format: { type: 'string', description: "Pass 'html' to also return htmlBody (raw, unsanitized: a renderer must sanitize it). Default text only." },
        maxHtmlChars: { type: 'number', description: 'Max htmlBody characters when format is html. Default 200000.' },
      },
      required: ['entryId'],
    },
    output: {
      schema: { type: 'object', additionalProperties: true },
      render: (_a, v) => [{ type: 'text', text: JSON.stringify(v) }],
    },
    async execute(a) { return olk('read_mail', a) },
    presentCall: (a) => ({ card: 'generic', title: 'Read Outlook mail', kind: 'other', rawInput: { entryId: a.entryId } }),
  })

  ctx.tools.register({
    name: 'maildesk_draft_mail',
    description: 'Create an email DRAFT in the Outlook Drafts folder. Nothing is sent; the user can review and send it manually in Outlook, or ask for maildesk_send_mail which confirms before sending.',
    parameters: {
      type: 'object',
      additionalProperties: true,
      properties: {
        to: { type: 'string', description: 'To recipients, semicolon-separated.' },
        cc: { type: 'string', description: 'Cc recipients, semicolon-separated.' },
        bcc: { type: 'string', description: 'Bcc (blind carbon copy) recipients, semicolon-separated.' },
        subject: { type: 'string', description: 'Email subject.' },
        body: { type: 'string', description: 'Plain-text body.' },
        attachments: { type: 'array', items: { type: 'string' }, description: 'Optional local file paths to attach.' },
      },
      required: ['to', 'subject', 'body'],
    },
    output: {
      schema: { type: 'object', additionalProperties: true },
      render: (_a, v) => [{ type: 'text', text: JSON.stringify(v) }],
    },
    async execute(a) { return olk('draft_mail', a) },
    presentCall: (a) => ({ card: 'generic', title: 'Draft Outlook mail', kind: 'other', rawInput: { to: a.to, subject: a.subject } }),
  })

  ctx.tools.register({
    name: 'maildesk_send_mail',
    description: 'Send an email through the local Outlook client. TWO-CALL CONFIRMATION with EDITABLE popup: without confirmed:true the call returns needs-confirmation with the exact content and the editable field list; show it to the user in-session — they may allow as-is, edit any field (to/cc/subject/body) via free-text answers, or reject. If edited, retry with the edited arguments (one more confirmation of the edited content), then confirmed:true only after the user allows. Nothing is sent on the first call.',
    parameters: {
      type: 'object',
      additionalProperties: true,
      properties: {
        to: { type: 'string', description: 'To recipients, semicolon-separated.' },
        cc: { type: 'string', description: 'Cc recipients, semicolon-separated.' },
        bcc: { type: 'string', description: 'Bcc (blind carbon copy) recipients, semicolon-separated; hidden from other recipients.' },
        subject: { type: 'string', description: 'Email subject.' },
        body: { type: 'string', description: 'Plain-text body.' },
        attachments: { type: 'array', items: { type: 'string' }, description: 'Optional local file paths to attach; shown in the confirmation popup.' },
        confirmed: { type: 'boolean', description: 'Set true ONLY after the user approved this exact content via an in-session popup.' },
      },
      required: ['to', 'subject', 'body'],
    },
    output: {
      schema: { type: 'object', additionalProperties: true },
      render: (_a, v) => [{ type: 'text', text: JSON.stringify(v) }],
    },
    async execute(a) {
      // Content binding: the confirmed retry must carry EXACTLY the content
      // of the last needs-confirmation the user saw (see `confirmations`).
      const content = {
        to: a.to, cc: a.cc || '', bcc: a.bcc || '', subject: a.subject, body: String(a.body),
        attachments: Array.isArray(a.attachments) ? a.attachments : [],
      }
      const h = contentHash(content)
      if (a.confirmed !== true || confirmations.get('send-mail') !== h) {
        // Full content, verbatim: the user must see exactly what would be sent.
        const att = attachmentLines(a.attachments)
        if (att.error) return { status: 'error', error: att.error }
        confirmations.set('send-mail', h)
        return needsConfirmation('send-mail',
          'To: ' + a.to + (a.cc ? '\nCc: ' + a.cc : '') + (a.bcc ? '\nBcc: ' + a.bcc : '') + '\nSubject: ' + a.subject
          + (att.lines ? '\nAttachments: ' + att.lines : '')
          + '\n\nBody:\n' + String(a.body),
          ['to', 'cc', 'bcc', 'subject', 'body'])
      }
      const draft = await olk('draft_mail', a)
      if (draft.status !== 'ok') return draft
      return dispatchConfirmed('send-mail', 'send_draft', { entryId: draft.draftEntryId }, { confirmed: true })
    },
    presentCall: (a) => ({ card: 'generic', title: 'Send Outlook mail (asks first)', kind: 'other', rawInput: { to: a.to, subject: a.subject } }),
  })

  ctx.tools.register({
    name: 'maildesk_calendar_query',
    description: 'List calendar events for the next N days (default 7) from the local Outlook client. Each event carries an entryId usable with maildesk_meeting_update / maildesk_meeting_cancel, plus an isMeeting flag. Read-only.',
    parameters: {
      type: 'object',
      additionalProperties: true,
      properties: {
        days: { type: 'number', description: 'How many days ahead to query. Default 7.' },
      },
    },
    output: {
      schema: { type: 'object', additionalProperties: true },
      render: (_a, v) => [{ type: 'text', text: JSON.stringify(v) }],
    },
    async execute(a) { return olk('calendar_query', a) },
    presentCall: (a) => ({ card: 'generic', title: 'Query Outlook calendar', kind: 'other', rawInput: a }),
  })

  ctx.tools.register({
    name: 'maildesk_calendar_create',
    description: 'Create a calendar entry in Outlook. Personal entries (no attendees) are created directly. With attendees (required recipients or optionalRecipients) it becomes a meeting invite and uses TWO-CALL CONFIRMATION with EDITABLE popup: the first call returns needs-confirmation; the user may allow as-is, edit fields (subject/location/body), or reject. Supports recurring meetings via the recurrence object (daily/weekly/monthly).',
    parameters: {
      type: 'object',
      additionalProperties: true,
      properties: {
        subject: { type: 'string', description: 'Event subject.' },
        start: { type: 'string', description: 'Start time, "yyyy-MM-dd HH:mm" (24h).' },
        end: { type: 'string', description: 'End time, "yyyy-MM-dd HH:mm" (24h).' },
        location: { type: 'string', description: 'Location text.' },
        body: { type: 'string', description: 'Plain-text notes/body.' },
        recipients: { type: 'array', items: { type: 'string' }, description: 'Required attendee SMTP addresses; adding any (or optionalRecipients) turns this into a meeting invite (requires user confirmation).' },
        optionalRecipients: { type: 'array', items: { type: 'string' }, description: 'Optional attendees (marked optional in the invite).' },
        recurrence: { type: 'object', additionalProperties: true, description: 'Optional recurrence: { type: "daily"|"weekly"|"monthly", interval: n, daysOfWeek: ["Mon","Wed"...] (weekly), count: n | until: "yyyy-MM-dd" }. Defaults to 10 occurrences if neither count nor until is given.' },
        confirmed: { type: 'boolean', description: 'Set true ONLY after the user approved the invite via an in-session popup. Ignored for personal entries.' },
      },
      required: ['subject', 'start', 'end'],
    },
    output: {
      schema: { type: 'object', additionalProperties: true },
      render: (_a, v) => [{ type: 'text', text: JSON.stringify(v) }],
    },
    async execute(a) {
      const required = Array.isArray(a.recipients) ? a.recipients : []
      const optional = Array.isArray(a.optionalRecipients) ? a.optionalRecipients : []
      const hasRecipients = required.length > 0 || optional.length > 0
      const rec = a.recurrence || null
      const content = {
        subject: a.subject, start: a.start, end: a.end,
        location: a.location || '', body: String(a.body || ''),
        recipients: required, optionalRecipients: optional,
        recurrence: rec ? JSON.stringify(rec) : '',
      }
      const h = contentHash(content)
      if ((hasRecipients && a.confirmed !== true) || (hasRecipients && confirmations.get('calendar-invite') !== h)) {
        // Full content, verbatim: the user must see exactly what would be sent.
        confirmations.set('calendar-invite', h)
        let recText = ''
        if (rec) {
          recText = '\nRecurrence: ' + rec.type + (rec.interval ? ' x' + rec.interval : '')
          if (rec.daysOfWeek && rec.daysOfWeek.length) recText += ' on ' + rec.daysOfWeek.join(',')
          if (rec.count) recText += ', ' + rec.count + ' times'
          else if (rec.until) recText += ', until ' + rec.until
        }
        return needsConfirmation('calendar-invite',
          'Subject: ' + a.subject + '\nWhen: ' + a.start + ' ~ ' + a.end
            + (a.location ? '\nLocation: ' + a.location : '')
            + (required.length ? '\nRequired: ' + required.join('; ') : '')
            + (optional.length ? '\nOptional: ' + optional.join('; ') : '')
            + (hasRecipients ? ' (meeting invites will be sent)' : '')
            + recText
            + (a.body ? '\n\nBody:\n' + String(a.body) : ''),
          ['subject', 'location', 'body'])
      }
      return dispatchConfirmed('calendar-invite', 'calendar_create', a, { confirmed: hasRecipients })
    },
    presentCall: (a) => ({ card: 'generic', title: 'Create Outlook calendar entry (asks first)', kind: 'other', rawInput: { subject: a.subject, start: a.start, end: a.end } }),
  })

  ctx.tools.register({
    name: 'maildesk_meeting_requests',
    description: 'List meeting requests waiting for your response (invites others sent you, last 30 days), with organizer, proposed time and location. Read-only; respond with maildesk_respond_meeting.',
    parameters: { type: 'object', additionalProperties: true, properties: {} },
    output: {
      schema: { type: 'object', additionalProperties: true },
      render: (_a, v) => [{ type: 'text', text: JSON.stringify(v) }],
    },
    async execute() { return olk('meeting_requests') },
    presentCall: () => ({ card: 'generic', title: 'List Outlook meeting requests', kind: 'other', rawInput: {} }),
  })

  ctx.tools.register({
    name: 'maildesk_respond_meeting',
    description: 'Accept, tentatively accept, or decline a meeting request (entryId from maildesk_meeting_requests). TWO-CALL CONFIRMATION: the first call returns needs-confirmation; retry with confirmed:true after the user allows. Accepting/declining sends a response to the organizer.',
    parameters: {
      type: 'object',
      additionalProperties: true,
      properties: {
        entryId: { type: 'string', description: 'entryId of the meeting request (from maildesk_meeting_requests).' },
        response: { type: 'string', description: '"accept", "tentative", or "decline".' },
        confirmed: { type: 'boolean', description: 'Set true ONLY after the user approved via an in-session popup.' },
      },
      required: ['entryId', 'response'],
    },
    output: {
      schema: { type: 'object', additionalProperties: true },
      render: (_a, v) => [{ type: 'text', text: JSON.stringify(v) }],
    },
    async execute(a) {
      const h = contentHash({ entryId: a.entryId, response: a.response })
      if (a.confirmed !== true || confirmations.get('meeting-respond') !== h) {
        confirmations.set('meeting-respond', h)
        const label = a.response === 'accept' ? 'accept' : a.response === 'tentative' ? 'tentative' : a.response === 'decline' ? 'decline' : String(a.response)
        return needsConfirmation('meeting-respond',
          'Meeting response: ' + label + '\nRequest entryId: ' + a.entryId + '\n(the response goes to the organizer; call maildesk_meeting_requests first for details)')
      }
      return dispatchConfirmed('meeting-respond', 'respond_meeting', a, { confirmed: true })
    },
    presentCall: (a) => ({ card: 'generic', title: 'Respond to Outlook meeting (asks first)', kind: 'other', rawInput: { entryId: a.entryId, response: a.response } }),
  })

  ctx.tools.register({
    name: 'maildesk_meeting_update',
    description: 'Update an existing appointment/meeting on the calendar (entryId from maildesk_calendar_query — use the outlook entryId if available). Changed fields are applied; if the meeting has attendees, Outlook re-notifies them. TWO-CALL CONFIRMATION: first call returns needs-confirmation; retry with confirmed:true after the user allows. Applies to personal appointments too (uniform gate).',
    parameters: {
      type: 'object',
      additionalProperties: true,
      properties: {
        entryId: { type: 'string', description: 'entryId of the appointment/meeting to update.' },
        subject: { type: 'string', description: 'New subject (optional).' },
        start: { type: 'string', description: 'New start, "yyyy-MM-dd HH:mm" (optional).' },
        end: { type: 'string', description: 'New end, "yyyy-MM-dd HH:mm" (optional).' },
        location: { type: 'string', description: 'New location (optional).' },
        body: { type: 'string', description: 'New body (optional).' },
        confirmed: { type: 'boolean', description: 'Set true ONLY after the user approved; required only when attendees will be notified.' },
      },
      required: ['entryId'],
    },
    output: {
      schema: { type: 'object', additionalProperties: true },
      render: (_a, v) => [{ type: 'text', text: JSON.stringify(v) }],
    },
    async execute(a) {
      const h = contentHash({ entryId: a.entryId, subject: a.subject || '', start: a.start || '', end: a.end || '', location: a.location || '', body: a.body || '' })
      if (a.confirmed !== true || confirmations.get('meeting-update') !== h) {
        confirmations.set('meeting-update', h)
        const changes = []
        if (a.subject) changes.push('subject -> ' + a.subject)
        if (a.start) changes.push('start -> ' + a.start)
        if (a.end) changes.push('end -> ' + a.end)
        if (a.location) changes.push('location -> ' + a.location)
        if (a.body) changes.push('body -> ' + String(a.body).slice(0, 300))
        if (changes.length === 0) return { status: 'error', error: 'nothing to update: provide at least one changed field.' }
        return needsConfirmation('meeting-update',
          'Update meeting entryId: ' + a.entryId + '\nchanges:\n- ' + changes.join('\n- ')
          + '\n(attendees are notified automatically if this is a group meeting)')
      }
      return dispatchConfirmed('meeting-update', 'meeting_update', a, { confirmed: true })
    },
    presentCall: (a) => ({ card: 'generic', title: 'Update Outlook meeting (asks first)', kind: 'other', rawInput: { entryId: a.entryId } }),
  })

  ctx.tools.register({
    name: 'maildesk_meeting_cancel',
    description: 'Cancel a meeting (notifies all attendees with a cancellation) or delete a personal appointment. TWO-CALL CONFIRMATION for both paths: the first call returns needs-confirmation; retry with confirmed:true after the user allows.',
    parameters: {
      type: 'object',
      additionalProperties: true,
      properties: {
        entryId: { type: 'string', description: 'entryId of the appointment/meeting to cancel.' },
        confirmed: { type: 'boolean', description: 'Set true ONLY after the user approved via an in-session popup.' },
      },
      required: ['entryId'],
    },
    output: {
      schema: { type: 'object', additionalProperties: true },
      render: (_a, v) => [{ type: 'text', text: JSON.stringify(v) }],
    },
    async execute(a) {
      const h = contentHash({ entryId: a.entryId })
      if (a.confirmed !== true || confirmations.get('meeting-cancel') !== h) {
        confirmations.set('meeting-cancel', h)
        return needsConfirmation('meeting-cancel',
          'Cancel meeting entryId: ' + a.entryId + '\n(all attendees get a cancellation notice if this is a group meeting)')
      }
      return dispatchConfirmed('meeting-cancel', 'meeting_cancel', a, { confirmed: true })
    },
    presentCall: (a) => ({ card: 'generic', title: 'Cancel Outlook meeting (asks first)', kind: 'other', rawInput: { entryId: a.entryId } }),
  })

  ctx.tools.register({
    name: 'maildesk_search_people',
    description: 'Search the Outlook address book (GAL) by display name, SMTP address or alias. Returns matching people with name and smtp. Read-only.',
    parameters: {
      type: 'object',
      additionalProperties: true,
      properties: {
        query: { type: 'string', description: 'Search text: partial display name (e.g. Wei Liu), email, or alias.' },
        limit: { type: 'number', description: 'Max results. Default 20.' },
      },
      required: ['query'],
    },
    output: {
      schema: { type: 'object', additionalProperties: true },
      render: (_a, v) => [{ type: 'text', text: JSON.stringify(v) }],
    },
    async execute(a) { return olk('search_people', a) },
    presentCall: (a) => ({ card: 'generic', title: 'Search Outlook address book', kind: 'other', rawInput: { query: a.query } }),
  })

  ctx.tools.register({
    name: 'maildesk_freebusy',
    description: 'Query a colleague busy/free status from Outlook. Returns busy slots (status: tentative/busy/out-of-office) per day for the next N days starting today. Read-only.',
    parameters: {
      type: 'object',
      additionalProperties: true,
      properties: {
        email: { type: 'string', description: 'SMTP address of the person (from maildesk_search_people).' },
        days: { type: 'number', description: 'How many days starting today. Default 3.' },
      },
      required: ['email'],
    },
    output: {
      schema: { type: 'object', additionalProperties: true },
      render: (_a, v) => [{ type: 'text', text: JSON.stringify(v) }],
    },
    async execute(a) { return olk('freebusy', a) },
    presentCall: (a) => ({ card: 'generic', title: 'Query Outlook free/busy', kind: 'other', rawInput: { email: a.email, days: a.days } }),
  })

  ctx.tools.register({
    name: 'maildesk_search_rooms',
    description: 'Search Outlook meeting rooms by name and optionally check their availability for a time window (date + start/end time). Returns rooms with name, smtp, availability and conflicting slots. Read-only.',
    parameters: {
      type: 'object',
      additionalProperties: true,
      properties: {
        query: { type: 'string', description: 'Optional name filter, e.g. "Ningbo" or "09.3".' },
        date: { type: 'string', description: 'Check availability on this date, "yyyy-MM-dd". Skip to list rooms only.' },
        start: { type: 'string', description: 'Window start "HH:mm". Default 09:00. Used with date.' },
        end: { type: 'string', description: 'Window end "HH:mm". Default 18:00. Used with date.' },
        limit: { type: 'number', description: 'Max rooms. Default 15.' },
      },
    },
    output: {
      schema: { type: 'object', additionalProperties: true },
      render: (_a, v) => [{ type: 'text', text: JSON.stringify(v) }],
    },
    async execute(a) { return olk('search_rooms', a) },
    presentCall: (a) => ({ card: 'generic', title: 'Search Outlook meeting rooms', kind: 'other', rawInput: { query: a.query, date: a.date } }),
  })

  ctx.tools.register({
    name: 'maildesk_reply_mail',
    description: 'Reply to, reply-all to, or forward an existing email (entryId from maildesk_search_mail). Auto-quotes the original message below your new text; forward carries the original attachments and needs a "to" recipient. TWO-CALL CONFIRMATION with EDITABLE popup: the first call returns needs-confirmation showing the original mail plus your reply; the user may allow as-is, edit (body/to), or reject. Only confirmed:true dispatches.',
    parameters: {
      type: 'object',
      additionalProperties: true,
      properties: {
        entryId: { type: 'string', description: 'entryId of the mail to reply to / forward (from maildesk_search_mail).' },
        mode: { type: 'string', description: '"reply" (sender only), "reply-all" (sender + all recipients), or "forward" (new recipient, keeps attachments).' },
        body: { type: 'string', description: 'New text to write above the auto-quoted original.' },
        to: { type: 'string', description: 'Forward recipient(s), semicolon-separated. Required for forward; ignored for reply/reply-all.' },
        cc: { type: 'string', description: 'Optional Cc recipients, semicolon-separated. Overwrites the reply-all preset when given.' },
        bcc: { type: 'string', description: 'Optional Bcc recipients, semicolon-separated; hidden from other recipients.' },
        attachments: { type: 'array', items: { type: 'string' }, description: 'Optional extra local file paths to attach (forward already carries the original attachments).' },
        confirmed: { type: 'boolean', description: 'Set true ONLY after the user approved this exact reply via an in-session popup.' },
      },
      required: ['entryId', 'mode'],
    },
    output: {
      schema: { type: 'object', additionalProperties: true },
      render: (_a, v) => [{ type: 'text', text: JSON.stringify(v) }],
    },
    async execute(a) {
      const content = {
        entryId: a.entryId, mode: a.mode, body: String(a.body || ''),
        to: a.to || '', cc: a.cc || '', bcc: a.bcc || '',
        attachments: Array.isArray(a.attachments) ? a.attachments : [],
      }
      const h = contentHash(content)
      if (a.confirmed !== true || confirmations.get('reply-mail') !== h) {
        // Fetch the original mail so the confirmation shows exactly which
        // mail the reply goes to — fail closed if it cannot be read.
        const prev = await olk('reply_preview', { entryId: a.entryId })
        if (prev.status !== 'ok') return prev
        const att = attachmentLines(a.attachments)
        if (att.error) return { status: 'error', error: att.error }
        const modeText = a.mode === 'forward' ? 'Forward to: ' + (a.to || '(empty)')
          : a.mode === 'reply-all' ? 'Reply all' : 'Reply'
        confirmations.set('reply-mail', h)
        return needsConfirmation('reply-mail',
          modeText + '\nOriginal: [' + prev.sender + '] ' + prev.subject
          + (a.to && a.mode === 'forward' ? '\nForward to: ' + a.to : '')
          + (a.cc ? '\nCc: ' + a.cc : '')
          + (a.bcc ? '\nBcc: ' + a.bcc : '')
          + (att.lines ? '\nAttach: ' + att.lines : '')
          + (a.mode === 'forward' ? '\n(a forward carries the original attachments automatically)' : '')
          + '\n\nYour content:\n' + String(a.body || '')
          + '\n\n--- original message quoted below ---\nFrom: ' + prev.sender
          + '\nSent: ' + (prev.received || '')
          + '\nTo: ' + prev.to
          + '\nSubject: ' + prev.subject,
          a.mode === 'forward' ? ['to', 'cc', 'bcc', 'body'] : ['cc', 'bcc', 'body'])
      }
      return dispatchConfirmed('reply-mail', 'reply_mail', a, { confirmed: true })
    },
    presentCall: (a) => ({ card: 'generic', title: 'Reply/forward Outlook mail (asks user first)', kind: 'other', rawInput: { entryId: a.entryId, mode: a.mode, to: a.to } }),
  })

  ctx.tools.register({
    name: 'maildesk_save_attachment',
    description: 'Save one attachment of an email to a local directory, by mail entryId and attachment display name (names come from maildesk_read_mail). Returns the saved file path and size. Writes a local file only; nothing leaves the machine.',
    parameters: {
      type: 'object',
      additionalProperties: true,
      properties: {
        entryId: { type: 'string', description: 'entryId of the mail holding the attachment (from maildesk_search_mail).' },
        name: { type: 'string', description: 'Attachment display name exactly as listed by maildesk_read_mail.' },
        dir: { type: 'string', description: 'Target directory for the saved file (must exist).' },
      },
      required: ['entryId', 'name', 'dir'],
    },
    output: {
      schema: { type: 'object', additionalProperties: true },
      render: (_a, v) => [{ type: 'text', text: JSON.stringify(v) }],
    },
    async execute(a) { return olk('save_attachment', a) },
    presentCall: (a) => ({ card: 'generic', title: 'Save Outlook attachment', kind: 'other', rawInput: { entryId: a.entryId, name: a.name } }),
  })

  ctx.tools.register({
    name: 'maildesk_check_new',
    description: 'Fetch new-mail arrivals pushed by the Outlook NewMailEx watcher since the last check, as a list (entryId, subject, sender, time). Fetching clears the pending queue and resets the sidebar badge. Returns unread totals too. Read-only. Tip: when presenting arrivals to the user, append a clickable open link per mail: [Open](http://127.0.0.1:3080/olk-notify/open?entryId=<entryId>).',
    parameters: {
      type: 'object',
      additionalProperties: true,
      properties: {},
    },
    output: {
      schema: { type: 'object', additionalProperties: true },
      render: (_a, v) => [{ type: 'text', text: JSON.stringify(v) }],
    },
    async execute() {
      const pending = notify.pending.splice(0, notify.pending.length)
      return { status: 'ok', count: pending.length, newMails: pending, unread: notify.unread }
    },
    presentCall: () => ({ card: 'generic', title: 'Check new Outlook mail', kind: 'other', rawInput: {} }),
  })


  ctx.tools.register({
    name: 'maildesk_list_folders',
    description: 'List the Outlook folder tree for every mailbox store: name, path, entryId, depth, item and unread counts. Read-only. Use a returned entryId or path as the folder argument of the mail tools.',
    parameters: {
      type: 'object',
      additionalProperties: true,
      properties: {
        maxDepth: { type: 'number', description: 'Maximum folder depth to walk. Default 6.' },
        counts: { type: 'boolean', description: 'Include item/unread counts (one COM read per folder). Default true; pass false on very large mailboxes.' },
      },
    },
    output: {
      schema: { type: 'object', additionalProperties: true },
      render: (_a, v) => [{ type: 'text', text: JSON.stringify(v) }],
    },
    async execute(a) { return olk('list_folders', a) },
    presentCall: () => ({ card: 'generic', title: 'List Outlook folders', kind: 'other', rawInput: {} }),
  })

  ctx.tools.register({
    name: 'maildesk_list_mail',
    description: 'List one page of a folder, newest first by default (read-only). Unlike maildesk_search_mail this is not a search: it answers what is in this folder, with offset/limit paging and a total count. Use it to browse a mailbox.',
    parameters: {
      type: 'object',
      additionalProperties: true,
      properties: {
        folder: { type: 'string', description: 'Folder name, path ("Inbox/ProjectA"), or entryId from maildesk_list_folders. Default Inbox.' },
        offset: { type: 'number', description: 'Items to skip. Default 0.' },
        limit: { type: 'number', description: 'Max items. Default 50, max 200.' },
        sort: { type: 'string', enum: ['received', 'subject', 'sender'], description: 'Sort field. Default received.' },
        descending: { type: 'boolean', description: 'Sort direction. Default true (newest first).' },
        unreadOnly: { type: 'boolean', description: 'Only unread items.' },
        preview: { type: 'boolean', description: 'Include a 160-char body preview per item (one COM read per item). Default false.' },
      },
    },
    output: {
      schema: { type: 'object', additionalProperties: true },
      render: (_a, v) => [{ type: 'text', text: JSON.stringify(v) }],
    },
    async execute(a) { return olk('list_mail', a) },
    presentCall: (a) => ({ card: 'generic', title: 'List Outlook mail', kind: 'other', rawInput: { folder: a.folder, offset: a.offset, limit: a.limit } }),
  })

  ctx.tools.register({
    name: 'maildesk_move_mail',
    description: 'Move or archive emails to another folder: one entryId or an array of entryIds, plus a destination folder (name, path, or folder entryId). Returns the NEW entryIds, because moving changes them. Recoverable; nothing leaves the machine.',
    parameters: {
      type: 'object',
      additionalProperties: true,
      properties: {
        entryId: { type: 'string', description: 'Single mail entryId.' },
        entryIds: { type: 'array', items: { type: 'string' }, description: 'Several mail entryIds (max 500).' },
        destFolder: { type: 'string', description: 'Destination folder: "Archive", "Inbox/ProjectA", or an entryId from maildesk_list_folders.' },
      },
      required: ['destFolder'],
    },
    output: {
      schema: { type: 'object', additionalProperties: true },
      render: (_a, v) => [{ type: 'text', text: JSON.stringify(v) }],
    },
    async execute(a) { return olk('move_mail', a) },
    presentCall: (a) => ({ card: 'generic', title: 'Move Outlook mail', kind: 'other', rawInput: { destFolder: a.destFolder, entryIds: a.entryIds } }),
  })

  ctx.tools.register({
    name: 'maildesk_delete_mail',
    description: 'Delete emails by entryId or entryIds. SOFT delete only: items go to Deleted Items and the user can undo them in Outlook. Pass expected equal to the number of entryIds you are deleting; the bridge aborts on a mismatch, so a stale selection cannot delete more than the user saw. No permanent delete exists.',
    parameters: {
      type: 'object',
      additionalProperties: true,
      properties: {
        entryId: { type: 'string', description: 'Single mail entryId.' },
        entryIds: { type: 'array', items: { type: 'string' }, description: 'Several mail entryIds (max 500).' },
        expected: { type: 'number', description: 'How many items the human selected. The bridge aborts when it does not match the number of entryIds supplied.' },
      },
    },
    output: {
      schema: { type: 'object', additionalProperties: true },
      render: (_a, v) => [{ type: 'text', text: JSON.stringify(v) }],
    },
    async execute(a) { return olk('delete_mail', a) },
    presentCall: (a) => ({ card: 'generic', title: 'Delete Outlook mail', kind: 'other', rawInput: { entryIds: a.entryIds, expected: a.expected } }),
  })

  ctx.tools.register({
    name: 'maildesk_mark_read',
    description: 'Mark emails read or unread by entryId or entryIds. Local state only; nothing leaves the machine.',
    parameters: {
      type: 'object',
      additionalProperties: true,
      properties: {
        entryId: { type: 'string', description: 'Single mail entryId.' },
        entryIds: { type: 'array', items: { type: 'string' }, description: 'Several mail entryIds (max 500).' },
        read: { type: 'boolean', description: 'true marks read (default); false marks unread.' },
      },
    },
    output: {
      schema: { type: 'object', additionalProperties: true },
      render: (_a, v) => [{ type: 'text', text: JSON.stringify(v) }],
    },
    async execute(a) { return olk('mark_read', a) },
    presentCall: (a) => ({ card: 'generic', title: 'Mark Outlook mail read/unread', kind: 'other', rawInput: { entryIds: a.entryIds, read: a.read } }),
  })
  ctx.tools.register({
    name: 'maildesk_compose',
    description:
      'Open or reopen a mail draft that belongs to this session instead of to Outlook, and show it as a draft card in the chat. '
      + 'Use it for every reply, forward or new message: the user sees the recipients, subject and body, can send it from the card, '
      + 'and re-opening the same thread returns the SAME draft instead of piling up a new one in Outlook\'s Drafts folder. '
      + 'Nothing is sent and nothing is written to Outlook by this call. Sending happens only when the human presses Send in the panel or the card.',
    parameters: {
      type: 'object',
      additionalProperties: true,
      properties: {
        mode: { type: 'string', enum: ['new', 'reply', 'replyAll', 'forward'], description: 'Draft kind. Default new.' },
        entryId: { type: 'string', description: 'The mail being answered (reply/replyAll/forward). Recipients and the quote are filled from it.' },
        draftId: { type: 'string', description: 'Reopen this exact draft instead of deriving one from mode+entryId.' },
        to: { type: 'string', description: 'Recipients as a ; separated list of "Name <addr>" or bare addresses. Omit to keep what the reply derived.' },
        cc: { type: 'string', description: 'Cc as a ; separated list.' },
        bcc: { type: 'string', description: 'Bcc as a ; separated list.' },
        subject: { type: 'string', description: 'Subject line. Omit to keep the derived "Re: …".' },
        body: { type: 'string', description: 'Plain-text body. Blank lines become paragraphs.' },
      },
    },
    output: {
      schema: { type: 'object', additionalProperties: true },
      render: (_a, v) => [{ type: 'text', text: JSON.stringify(v) }],
    },
    async execute(a) {
      if (composeApi === null) return { status: 'error', error: 'the draft store is not mounted yet' }
      const api = composeApi
      let draft = null
      let reused = false
      if (a.draftId) {
        draft = api.findDraft(String(a.draftId))
        if (draft === null) return { status: 'error', error: 'unknown draft: ' + a.draftId }
        reused = true
      } else {
        const opened = await api.openDraft({ mode: a.mode, inReplyTo: a.entryId })
        draft = api.findDraft(opened.draft.id)
        reused = opened.reused
      }
      const patch = {}
      if (a.to !== undefined) patch.to = api.parseAddrList(a.to)
      if (a.cc !== undefined) patch.cc = api.parseAddrList(a.cc)
      if (a.bcc !== undefined) patch.bcc = api.parseAddrList(a.bcc)
      if (a.subject !== undefined) patch.subject = String(a.subject)
      if (a.body !== undefined) {
        // Plain text in, minimal HTML out: the panel and Outlook both render
        // HTML, and a draft that silently degrades to one paragraph reads as a
        // different message than the one the user approved.
        const html = escapeHtml(String(a.body))
          .split(/\n{2,}/)
          .map((p) => '<p>' + p.replace(/\n/g, '<br>') + '</p>')
          .join('')
        patch.bodyHtml = html
      }
      if (Object.keys(patch).length > 0) api.patchDraft(draft, patch)
      const shown = api.publicDraft(draft)
      return {
        status: 'ok',
        draftId: shown.id,
        mode: shown.mode,
        reused,
        to: api.formatAddrList(shown.to),
        cc: api.formatAddrList(shown.cc),
        subject: shown.subject,
        bodyHtml: shown.bodyHtml,
        hasQuote: shown.quoteHtml !== '',
        instruction: 'The draft is now on screen. The user can edit it in the mail panel, or send it — sending asks for an explicit confirmation first.',
      }
    },
    presentCall: (a) => ({ card: 'generic', title: 'Compose mail', kind: 'other', rawInput: { mode: a.mode || 'new', entryId: a.entryId, subject: a.subject } }),
  })
  console.info('[dsh-arya-maildesk] registered 24 Outlook COM tools')
}
