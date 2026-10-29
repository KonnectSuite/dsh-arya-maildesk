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

process.env.MAILDESK_FIXTURE = '1'
process.env.OLK_WATCH = '0'

const plugin = await import('../index.js')

const routes = new Map()
const ctx = {
  tools: { register() { /* tool registration is covered elsewhere */ } },
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
assert.ok(typeof call === 'function', 'call route registered')
assert.ok(typeof health === 'function', 'health route registered')

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

let passed = 0
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

console.log((passed) + ' assertions passed' + (process.exitCode ? ' — with failures' : ''))
