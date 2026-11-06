/**
 * Attachments picked in the Mail panel.
 *
 * The browser can read a file's bytes but not its path, so the panel posts the
 * bytes here and the host writes them to one temp folder, handing the agent a
 * real absolute path. This suite pins the contract: what comes back, that a
 * hostile name cannot escape the folder, and that the endpoints refuse junk.
 *
 *   node test/attachments.test.mjs
 */
import assert from 'node:assert/strict'
import { existsSync, readFileSync, rmSync, mkdirSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join, dirname, isAbsolute } from 'node:path'

const DIR = join(tmpdir(), 'maildesk-attach-test-' + process.pid)
process.env.MAILDESK_ATTACH_DIR = DIR
process.env.OLK_WATCH = '0'
process.env.MAILDESK_NO_RESIDENT = '1'
process.env.MAILDESK_DRAFTS_FILE = join(tmpdir(), 'maildesk-attach-drafts-' + process.pid + '.json')
process.env.MAILDESK_STATE_FILE = join(tmpdir(), 'maildesk-attach-ask-' + process.pid + '.json')
rmSync(DIR, { recursive: true, force: true })

const plugin = await import('../index.js')

const routes = new Map()
const ctx = {
  tools: { register() {} },
  effect(fn) { const dispose = fn(); return typeof dispose === 'function' ? dispose : () => {} },
  get() { return undefined },
  inject(names, callback) {
    if (names.includes('webServer')) {
      callback({ webServer: { register(spec) { routes.set(spec.path, spec.handler); return () => routes.delete(spec.path) } } })
    }
  },
}
plugin.apply(ctx)

const upload = routes.get('/maildesk/api/attachments')
assert.ok(typeof upload === 'function', 'attachments route registered')

const HOST = '127.0.0.1:58011'
async function post(handler, body) {
  const text = body === undefined ? '' : JSON.stringify(body)
  const req = {
    method: 'POST',
    headers: { host: HOST, origin: 'http://' + HOST, 'content-type': 'application/json' },
    async *[Symbol.asyncIterator]() { if (text !== '') yield Buffer.from(text) },
  }
  const res = {
    statusCode: 200,
    headers: {},
    setHeader(name, value) { this.headers[String(name).toLowerCase()] = value },
    end(payload) { this.payload = payload === undefined ? '' : String(payload) },
  }
  await handler(req, res)
  let json = null
  try { json = JSON.parse(res.payload) } catch { /* not JSON */ }
  return { status: res.statusCode, json }
}

let checks = 0
function check(label, ok, detail) {
  checks += 1
  if (!ok) throw new Error(`FAIL ${label}${detail === undefined ? '' : ' — ' + detail}`)
  console.log(`ok   ${label}`)
}

const bytes = Buffer.from('the quick brown fox\n', 'utf8')

// 1. A normal upload lands on disk, byte for byte, at an absolute path.
{
  const r = await post(upload, { name: 'punchlist.txt', data: bytes.toString('base64') })
  check('the upload is accepted', r.json.ok === true, JSON.stringify(r.json).slice(0, 140))
  check('an absolute path comes back', typeof r.json.path === 'string' && isAbsolute(r.json.path), String(r.json.path))
  check('the file really exists', existsSync(r.json.path), String(r.json.path))
  check('the bytes survive the round trip', readFileSync(r.json.path).equals(bytes), 'content differs')
  check('the name and size are reported', r.json.name === 'punchlist.txt' && r.json.size === bytes.length, JSON.stringify({ name: r.json.name, size: r.json.size }))
  check('the file sits in the attachment folder', dirname(r.json.path) === DIR, dirname(r.json.path))
}

// 2. A hostile name is reduced to its base name: nothing escapes the folder.
{
  const r = await post(upload, { name: '../../evil.txt', data: bytes.toString('base64') })
  check('a traversal name is accepted but flattened', r.json.ok === true && r.json.name === 'evil.txt', JSON.stringify(r.json).slice(0, 140))
  check('it still lands inside the attachment folder', dirname(r.json.path) === DIR, dirname(r.json.path))
  check('nothing was written above the folder', !existsSync(join(DIR, '..', 'evil.txt')))
}

// 3. Windows-reserved characters are replaced rather than passed through.
{
  const r = await post(upload, { name: 'a<b>c|d?.txt', data: bytes.toString('base64') })
  check('reserved characters are replaced', r.json.ok === true && !/[<>|?]/.test(r.json.name), String(r.json.name))
}

// 4. Empty payloads are refused instead of writing a zero-byte attachment.
{
  const r = await post(upload, { name: 'empty.txt', data: '' })
  check('an empty file is refused', r.status === 400 && r.json.ok === false, JSON.stringify(r.json))
}

// 5. Transport edges.
{
  const req = { method: 'GET', headers: { host: HOST, origin: 'http://' + HOST }, async *[Symbol.asyncIterator]() {} }
  const res = { statusCode: 200, setHeader() {}, end(p) { this.payload = String(p) } }
  await upload(req, res)
  check('GET is refused', res.statusCode === 405, String(res.statusCode))

  const foreign = await post(upload, { name: 'x.txt', data: bytes.toString('base64') })
  check('a same-origin upload works for the comparison', foreign.json.ok === true)

  const req2 = { method: 'POST', headers: { host: HOST, origin: 'https://evil.example', 'content-type': 'application/json' }, async *[Symbol.asyncIterator]() { yield Buffer.from('{}') } }
  const res2 = { statusCode: 200, setHeader() {}, end(p) { this.payload = String(p) } }
  await upload(req2, res2)
  check('another origin is refused', res2.statusCode === 403, String(res2.statusCode))
}

console.log(`\n${checks} assertions passed`)
