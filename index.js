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
import { createHash, randomBytes } from 'node:crypto'
import { statSync, appendFileSync, writeFileSync } from 'node:fs'
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

const fixture = FIXTURE_ON ? makeFixture() : null

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
        if (m) return ok({ subject: m.subject, sender: m.sender, received: m.received, to: m.to, cc: m.cc, bcc: m.bcc, body: m.body, attachments: m.attachments, unread: m.unread })
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
    case 'open_mail':
      return ok({ opened: true, subject: '(fixture)' })
    default:
      return { status: 'error', error: 'action not available to the panel: ' + action }
  }
}

/** One panel call: the fixture when explicitly enabled, COM otherwise. */
async function panelCall(action, args) {
  if (!PANEL_ACTIONS.has(action)) return { status: 'error', error: 'action not allowed: ' + action }
  if (fixture !== null) return fixtureCall(action, args)
  return olk(action, args)
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
        res.end(JSON.stringify({ ok: true, provider: fixture !== null ? 'fixture' : 'com' }))
      },
    }), 'dsh-arya-maildesk: panel health route')

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
  console.info('[dsh-arya-maildesk] registered 23 Outlook COM tools')
}
