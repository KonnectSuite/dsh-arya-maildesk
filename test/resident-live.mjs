/**
 * Live probe for the resident bridge — real Outlook, real host routes.
 * Not part of `npm test`: it needs COM and a signed-in profile.
 *
 *   node test/resident-live.mjs
 */
import { rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

process.env.OLK_WATCH = '0'
delete process.env.MAILDESK_FIXTURE
process.env.MAILDESK_DRAFTS_FILE = join(tmpdir(), 'maildesk-resident-drafts-' + process.pid + '.json')
process.env.MAILDESK_STATE_FILE = join(tmpdir(), 'maildesk-resident-ask-' + process.pid + '.json')
rmSync(process.env.MAILDESK_DRAFTS_FILE, { force: true })

const plugin = await import('../index.js?resident=1')

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

const call = routes.get('/maildesk/api/call')
const HOST = '127.0.0.1:57019'

async function hit(action, args) {
  const text = JSON.stringify({ action, args })
  const req = {
    method: 'POST',
    headers: { host: HOST, origin: 'http://' + HOST, 'content-type': 'application/json' },
    async *[Symbol.asyncIterator]() { yield Buffer.from(text) },
  }
  const res = {
    statusCode: 200,
    headers: {},
    setHeader(name, value) { this.headers[String(name).toLowerCase()] = value },
    end(payload) { this.payload = payload === undefined ? '' : String(payload) },
  }
  await call(req, res)
  return JSON.parse(res.payload)
}

async function timed(label, action, args) {
  const started = Date.now()
  const out = await hit(action, args)
  const ms = Date.now() - started
  const result = out.result || {}
  console.log(`${label.padEnd(28)} ${String(ms).padStart(5)} ms  status=${result.status} ${result.error ? 'error=' + result.error : ''}`)
  return { ms, result }
}

const cold = await timed('cold (starts the worker)', 'list_folders', { roots: true })
const warm1 = await timed('warm list_mail x20', 'list_mail', { folder: 'Inbox', limit: 20, preview: true })
const warm2 = await timed('warm list_mail x20 again', 'list_mail', { folder: 'Inbox', limit: 20, preview: true })
const warm3 = await timed('warm account', 'account', {})

const ok = cold.result.status === 'ok' && warm1.result.status === 'ok' && warm2.result.status === 'ok' && warm3.result.status === 'ok'
console.log(`\n${ok ? 'PASS' : 'FAIL'}: cold ${cold.ms} ms, warm ${Math.round((warm1.ms + warm2.ms + warm3.ms) / 3)} ms on average`)
if (!ok) process.exit(1)
process.exit(0)
