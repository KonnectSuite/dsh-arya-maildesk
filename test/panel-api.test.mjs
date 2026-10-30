/**
 * Panel API smoke test — runs the real host half's HTTP route against the
 * fixture provider, with no Outlook and no harness involved.
 *
 * What it can prove here: the route, the action allowlist, the same-origin
 * guard, the JSON contracts the panel consumes, and that organize actions
 * mutate the mailbox. What it cannot prove: anything about COM.
 *
 *   node test/panel-api.test.mjs
 */
import assert from 'node:assert/strict'
import { rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

process.env.MAILDESK_FIXTURE = '1'
process.env.OLK_WATCH = '0'
// Drafts now persist to disk, so a run must not inherit the previous run's
// drafts (or the app's): start every suite from an empty store.
process.env.MAILDESK_DRAFTS_FILE = join(tmpdir(), 'maildesk-panel-drafts-' + process.pid + '.json')
process.env.MAILDESK_STATE_FILE = join(tmpdir(), 'maildesk-panel-ask-' + process.pid + '.json')
rmSync(process.env.MAILDESK_DRAFTS_FILE, { force: true })

const plugin = await import('../index.js')

const routes = new Map()
const toolSpecs = new Map()
const ctx = {
  tools: { register(spec) { toolSpecs.set(spec.name, spec) } },
  effect(fn) { const dispose = fn(); return typeof dispose === 'function' ? dispose : () => {} },
  inject(names, callback) {
    if (names.includes('webServer')) {
      callback({
        webServer: {
          register(spec) { routes.set(spec.path, spec.handler); return () => routes.delete(spec.path) },
        },
      })
    }
  },
}
plugin.apply(ctx)

const call = routes.get('/maildesk/api/call')
const health = routes.get('/maildesk/api/health')
const draftRoute = routes.get('/maildesk/api/draft')
const sendRoute = routes.get('/maildesk/api/send')
assert.ok(typeof call === 'function', 'call route registered')
assert.ok(typeof health === 'function', 'health route registered')
assert.ok(typeof draftRoute === 'function', 'draft route registered')
assert.ok(typeof sendRoute === 'function', 'send route registered')

const HOST = '127.0.0.1:57019'

function request(method, body, headers) {
  const text = body === undefined ? '' : JSON.stringify(body)
  const req = {
    method,
    headers: { host: HOST, origin: 'http://' + HOST, 'content-type': 'application/json', ...(headers || {}) },
    async *[Symbol.asyncIterator]() {
      if (text !== '') yield Buffer.from(text)
    },
  }
  const res = {
    statusCode: 200,
    setHeader() {},
    end(payload) { this.payload = payload === undefined ? '' : String(payload) },
  }
  return { req, res }
}

async function post(action, args) {
  const { req, res } = request('POST', { action, args: args || {} })
  await call(req, res)
  const parsed = JSON.parse(res.payload)
  return { status: res.statusCode, ...parsed }
}

/** Hit any of the panel's routes, not just the allowlisted call route. */
async function hit(path, body) {
  const handler = routes.get(path)
  const { req, res } = request('POST', body)
  await handler(req, res)
  return { status: res.statusCode, ...JSON.parse(res.payload || '{}') }
}

let passed = 0
let draftId = ''
async function check(name, fn) {
  try {
    await fn()
    passed++
    console.log('  ok   ' + name)
  } catch (error) {
    console.log('  FAIL ' + name + ' :: ' + error.message)
    process.exitCode = 1
  }
}

console.log('panel api (fixture provider)')

let folders = []
await check('health reports the fixture provider', async () => {
  const { req, res } = request('GET')
  await health(req, res)
  const data = JSON.parse(res.payload)
  assert.equal(data.ok, true)
  assert.equal(data.provider, 'fixture')
})

await check('list_folders returns a tree with counts', async () => {
  const r = await post('list_folders', { maxDepth: 6 })
  assert.equal(r.ok, true)
  assert.equal(r.result.status, 'ok')
  folders = r.result.folders
  assert.ok(folders.length >= 8, 'expected at least 8 folders, got ' + folders.length)
  const inbox = folders.find((f) => f.name === 'Inbox')
  assert.ok(inbox, 'Inbox present')
  assert.equal(typeof inbox.itemCount, 'number')
  assert.equal(typeof inbox.unread, 'number')
  assert.ok(folders.some((f) => f.depth > 0), 'nested folder present')
  assert.ok(folders.every((f) => f.entryId.length >= 60), 'every entryId is long enough to take the id path')
})

const inboxId = folders.find((f) => f.name === 'Inbox').entryId

await check('list_mail pages a folder and returns the panel fields', async () => {
  const r = await post('list_mail', { folder: inboxId, limit: 5, preview: true })
  assert.equal(r.ok, true)
  assert.equal(r.result.returned, 5)
  assert.ok(r.result.total >= r.result.returned)
  const m = r.result.mails[0]
  for (const key of ['entryId', 'subject', 'sender', 'senderSmtp', 'received', 'unread', 'hasAttach', 'preview']) {
    assert.ok(key in m, 'mail row is missing ' + key)
  }
  assert.match(m.received, /^\d{4}-\d{2}-\d{2} \d{2}:\d{2}$/, 'received is yyyy-MM-dd HH:mm')
})

await check('list_mail honours unreadOnly', async () => {
  const r = await post('list_mail', { folder: inboxId, limit: 200, unreadOnly: true })
  assert.equal(r.ok, true)
  assert.ok(r.result.mails.length > 0, 'the fixture has unread mail')
  assert.ok(r.result.mails.every((m) => m.unread === true), 'every row is unread')
})

await check('list_mail sorts by sender', async () => {
  const r = await post('list_mail', { folder: inboxId, limit: 200, sort: 'sender' })
  const names = r.result.mails.map((m) => m.sender)
  const sorted = names.slice().sort((a, b) => (a < b ? 1 : a > b ? -1 : 0))
  assert.deepEqual(names, sorted)
})

await check('list_mail accepts a folder path as well as an entryId', async () => {
  const r = await post('list_mail', { folder: 'Inbox', limit: 2 })
  assert.equal(r.ok, true)
  assert.match(r.result.folderPath, /Inbox/)
})

await check('an unknown folder is an error the panel can show', async () => {
  const r = await post('list_mail', { folder: 'Nope/Missing', limit: 2 })
  assert.equal(r.ok, true)
  assert.equal(r.result.status, 'error')
  assert.match(r.result.error, /no folder matches/)
})

await check('read_mail returns a body and attachments array', async () => {
  const list = await post('list_mail', { folder: inboxId, limit: 1 })
  const id = list.result.mails[0].entryId
  const r = await post('read_mail', { entryId: id })
  assert.equal(r.ok, true)
  assert.ok(r.result.body.length > 20)
  assert.ok(Array.isArray(r.result.attachments))
  assert.ok(r.result.to.includes('kphung@frommconnect.com'))
})

await check('read_mail returns htmlBody only when html is asked for', async () => {
  const list = await post('list_mail', { folder: inboxId, limit: 1 })
  const id = list.result.mails[0].entryId
  const plain = await post('read_mail', { entryId: id })
  assert.equal(plain.result.htmlBody, undefined, 'htmlBody must stay out of the default payload')
  const rich = await post('read_mail', { entryId: id, format: 'html' })
  assert.equal(rich.ok, true)
  assert.equal(typeof rich.result.htmlBody, 'string')
  assert.ok(rich.result.htmlBody.includes('<table'), 'the fixture body carries a table')
  // The fixture body is deliberately hostile; the renderer, not the bridge,
  // is responsible for removing these - the bridge must not silently rewrite
  // the mail it reports.
  assert.ok(rich.result.htmlBody.includes('<script'), 'raw body is untrusted input')
  assert.equal(rich.result.htmlTruncated, false)
})

await check('mark_read flips only the given mail', async () => {
  const list = await post('list_mail', { folder: inboxId, limit: 2 })
  const [a, b] = list.result.mails
  await post('mark_read', { entryIds: [a.entryId], read: true })
  const after = await post('list_mail', { folder: inboxId, limit: 200 })
  const reloaded = after.result.mails.find((m) => m.entryId === a.entryId)
  assert.equal(reloaded.unread, false)
  const other = after.result.mails.find((m) => m.entryId === b.entryId)
  assert.equal(other.unread, b.unread)
})

await check('move_mail returns a NEW entryId and a destination path', async () => {
  const list = await post('list_mail', { folder: inboxId, limit: 1 })
  const victim = list.result.mails[0].entryId
  const dest = folders.find((f) => f.name === '2026')
  const r = await post('move_mail', { entryIds: [victim], destFolder: dest.entryId })
  assert.equal(r.ok, true)
  assert.equal(r.result.moved, 1)
  assert.ok(r.result.destPath.endsWith('2026'))
  assert.notEqual(r.result.newEntryIds[0], victim, 'EntryID must not survive a move')
  const after = await post('list_mail', { folder: inboxId, limit: 200 })
  assert.ok(!after.result.mails.some((m) => m.entryId === victim), 'the mail left the inbox')
})

await check('delete_mail refuses a count mismatch and deletes nothing', async () => {
  const list = await post('list_mail', { folder: inboxId, limit: 3 })
  const ids = list.result.mails.map((m) => m.entryId)
  const r = await post('delete_mail', { entryIds: ids, expected: 2 })
  assert.equal(r.ok, true)
  assert.equal(r.result.status, 'error')
  assert.match(r.result.error, /refusing to delete/)
  const after = await post('list_mail', { folder: inboxId, limit: 200 })
  for (const id of ids) assert.ok(after.result.mails.some((m) => m.entryId === id), 'nothing was deleted on refusal')
})

await check('delete_mail soft-deletes what was asked for', async () => {
  const list = await post('list_mail', { folder: inboxId, limit: 2 })
  const ids = list.result.mails.map((m) => m.entryId)
  const r = await post('delete_mail', { entryIds: ids, expected: ids.length })
  assert.equal(r.ok, true)
  assert.equal(r.result.deleted, ids.length)
  assert.equal(r.result.recoverable, true)
  const after = await post('list_mail', { folder: inboxId, limit: 200 })
  for (const id of ids) assert.ok(!after.result.mails.some((m) => m.entryId === id))
})

await check('the send paths are not reachable from the panel', async () => {
  for (const action of ['send_draft', 'reply_mail', 'draft_mail', 'calendar_create']) {
    const r = await post(action, {})
    assert.equal(r.status, 400, action + ' must be refused by the allowlist')
    assert.equal(r.ok, false)
  }
})

await check('a cross-origin caller is refused', async () => {
  const { req, res } = request('POST', { action: 'list_folders', args: {} }, { origin: 'http://evil.example' })
  await call(req, res)
  assert.equal(res.statusCode, 403)
})

await check('a GET on the call route is refused', async () => {
  const { req, res } = request('GET')
  await call(req, res)
  assert.equal(res.statusCode, 405)
})

await check('a malformed body is refused', async () => {
  const req = {
    method: 'POST',
    headers: { host: HOST, origin: 'http://' + HOST },
    async *[Symbol.asyncIterator]() { yield Buffer.from('{not json') },
  }
  const res = { statusCode: 200, setHeader() {}, end(p) { this.payload = String(p) } }
  await call(req, res)
  assert.equal(res.statusCode, 400)
  assert.equal(JSON.parse(res.payload).ok, false)
})

// ---- compose flow (M3) --------------------------------------------------
await check('a reply draft is created once and reused for the same thread', async () => {
  const list = await post('list_mail', { folder: 'Inbox', limit: 1 })
  const mail = list.result.mails[0]
  const first = await hit('/maildesk/api/draft', { op: 'open', mode: 'reply', inReplyTo: mail.entryId })
  assert.equal(first.ok, true)
  assert.equal(first.reused, false)
  assert.ok(first.draft.subject.toLowerCase().startsWith('re: '), 'subject carries the reply prefix')
  assert.ok(first.draft.quoteHtml.includes('<blockquote>'), 'the quote is carried')
  assert.ok(!/<script/i.test(first.draft.quoteHtml), 'the quote must not carry a script')
  assert.ok(!/\son[a-z]+\s*=/i.test(first.draft.quoteHtml), 'the quote must not carry an event handler')
  const again = await hit('/maildesk/api/draft', { op: 'open', mode: 'reply', inReplyTo: mail.entryId })
  assert.equal(again.reused, true, 're-opening a reply reuses the draft instead of stacking another')
  assert.equal(again.draft.id, first.draft.id)
  draftId = first.draft.id
})

await check('a draft patch is stored, and a list returns it', async () => {
  const r = await hit('/maildesk/api/draft', {
    op: 'save', draftId,
    patch: {
      to: [{ name: 'Melissa Chen', email: 'melissa@example.com' }],
      subject: 'Re: Line 4 retrofit',
      bodyHtml: '<p>Yes, Thursday works.</p>',
    },
  })
  assert.equal(r.ok, true)
  assert.equal(r.draft.to.length, 1)
  assert.equal(r.draft.to[0].email, 'melissa@example.com')
  const back = await hit('/maildesk/api/draft', { op: 'list' })
  assert.equal(back.drafts.length, 1)
  assert.equal(back.drafts[0].bodyHtml, '<p>Yes, Thursday works.</p>')
})

await check('send refuses until it is given the hash of the content shown', async () => {
  const r = await hit('/maildesk/api/send', { draftId })
  assert.equal(r.ok, false)
  assert.equal(r.result.status, 'needs-confirmation')
  assert.equal(r.result.hash.length, 64)
  assert.ok(r.result.preview.htmlBody.includes('Thursday works'), 'the preview is the exact content that would be sent')
  const sent = await post('list_mail', { folder: 'Sent Items', limit: 200 })
  assert.ok(!sent.result.mails.some((m) => m.subject === 'Re: Line 4 retrofit'), 'nothing may leave yet')
})

await check('send refuses a hash that does not match the draft', async () => {
  const r = await hit('/maildesk/api/send', { draftId, confirmHash: 'f'.repeat(64) })
  assert.equal(r.result.status, 'needs-confirmation')
})

await check('send with the matching hash sends it and consumes the draft', async () => {
  const pre = await hit('/maildesk/api/send', { draftId })
  const r = await hit('/maildesk/api/send', { draftId, confirmHash: pre.result.hash })
  assert.equal(r.ok, true)
  assert.equal(r.result.sent, true)
  const sent = await post('list_mail', { folder: 'Sent Items', limit: 200 })
  assert.ok(sent.result.mails.some((m) => m.subject === 'Re: Line 4 retrofit'), 'the sent mail is in Sent Items')
  const left = await hit('/maildesk/api/draft', { op: 'list' })
  assert.equal(left.drafts.length, 0, 'a sent draft is consumed, not left to pile up')
})

await check('a sent reply reads back as the very HTML that was sent', async () => {
  const sentList = await hit('/maildesk/api/call', { action: 'list_mail', args: { folder: 'Sent Items', limit: 5 } })
  const top = sentList.result.mails[0]
  assert.ok(top !== undefined, 'the reply landed in Sent Items')
  const read = await hit('/maildesk/api/call', { action: 'read_mail', args: { entryId: top.entryId, format: 'html' } })
  assert.ok(typeof read.result.htmlBody === 'string', 'the sent copy carries an HTML body')
  assert.ok(read.result.htmlBody.indexOf('maildesk-quote') !== -1, 'the quoted original is still under the reply')
  assert.equal(read.result.htmlBody.indexOf('<script'), -1, 'the outbound quote never carries the original script')
})

await check('a consumed draft cannot be sent twice', async () => {
  const r = await hit('/maildesk/api/send', { draftId })
  assert.equal(r.result.status, 'error')
})

await check('discard drops a draft without sending it', async () => {
  const opened = await hit('/maildesk/api/draft', { op: 'open', mode: 'new' })
  assert.equal(opened.ok, true)
  const dropped = await hit('/maildesk/api/draft', { op: 'discard', draftId: opened.draft.id })
  assert.equal(dropped.ok, true)
  assert.equal(dropped.discarded, true)
  const left = await hit('/maildesk/api/draft', { op: 'list' })
  assert.equal(left.drafts.length, 0)
})

await check('the compose routes refuse a cross-origin caller', async () => {
  for (const path of ['/maildesk/api/draft', '/maildesk/api/send']) {
    const handler = routes.get(path)
    const { req, res } = request('POST', { op: 'list' }, { origin: 'http://evil.example' })
    await handler(req, res)
    assert.equal(res.statusCode, 403, path + ' must refuse a foreign origin')
  }
})

// M5: the draft card. The chat renders a keyed tool.call.toolview entry for
// maildesk_compose, so the tool has to hand back a draft the card can poll.
const compose = toolSpecs.get('maildesk_compose')
assert.ok(compose !== undefined, 'maildesk_compose is registered')
let cardDraftId = ''

await check('maildesk_compose opens a reply draft the card can poll', async () => {
  const inbox = await post('list_mail', { folder: 'Inbox', limit: 1 })
  const entryId = inbox.result.mails[0].entryId
  const out = await compose.execute({ mode: 'reply', entryId, body: 'First line.\n\nSecond paragraph.' })
  assert.equal(out.status, 'ok')
  assert.equal(out.mode, 'reply')
  assert.equal(out.reused, false)
  assert.match(out.draftId, /^draft-/)
  assert.match(out.subject, /^Re: /)
  assert.match(out.bodyHtml, /<p>First line\.<\/p><p>Second paragraph\.<\/p>/)
  cardDraftId = out.draftId
  const got = await hit('/maildesk/api/draft', { op: 'get', draftId: cardDraftId })
  assert.equal(got.ok, true)
  assert.equal(got.draft.id, cardDraftId)
  assert.equal(got.draft.subject, out.subject)
})

await check('a second call on the same thread returns the same draft', async () => {
  const inbox = await post('list_mail', { folder: 'Inbox', limit: 1 })
  const entryId = inbox.result.mails[0].entryId
  const again = await compose.execute({ mode: 'reply', entryId })
  assert.equal(again.draftId, cardDraftId)
  assert.equal(again.reused, true)
  const draftsNow = await hit('/maildesk/api/draft', { op: 'list' })
  assert.equal(draftsNow.drafts.length, 1, 'one thread must hold exactly one draft')
})

await check('op get 404s for a draft that does not exist', async () => {
  const gone = await hit('/maildesk/api/draft', { op: 'get', draftId: 'draft-nope' })
  assert.equal(gone.status, 404)
})

console.log((passed) + ' assertions passed' + (process.exitCode ? ' — with failures' : ''))
