/**
 * Drafts survive a restart — the one loss that costs a person real work.
 *
 * The host mirrors its draft map to a JSON file on every change and reads it
 * back at mount. This drives the real compose route in one module instance,
 * then loads a second instance (cache-busted import) pointed at the same file
 * and asks it to list drafts.
 *
 *   node test/drafts-persist.test.mjs
 */
import assert from 'node:assert/strict'
import { readFileSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

process.env.MAILDESK_FIXTURE = '1'
process.env.OLK_WATCH = '0'
const DRAFTS = join(tmpdir(), 'maildesk-drafts-test-' + process.pid + '.json')
process.env.MAILDESK_DRAFTS_FILE = DRAFTS
process.env.MAILDESK_STATE_FILE = join(tmpdir(), 'maildesk-drafts-ask-' + process.pid + '.json')
rmSync(DRAFTS, { force: true })

const HOST = '127.0.0.1:57019'

function harness() {
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
  return { routes, ctx }
}

function request(method, body) {
  const text = body === undefined ? '' : JSON.stringify(body)
  return {
    req: {
      method,
      headers: { host: HOST, origin: 'http://' + HOST, 'content-type': 'application/json' },
      async *[Symbol.asyncIterator]() { if (text !== '') yield Buffer.from(text) },
    },
    res: {
      statusCode: 200,
      headers: {},
      setHeader(name, value) { this.headers[String(name).toLowerCase()] = value },
      end(payload) { this.payload = payload === undefined ? '' : String(payload) },
    },
  }
}

async function hit(handler, method, body) {
  const { req, res } = request(method, body)
  await handler(req, res)
  let json = null
  try { json = JSON.parse(res.payload) } catch { /* not JSON */ }
  return { status: res.statusCode, json, text: res.payload }
}

let checks = 0
function check(label, ok, detail) {
  checks += 1
  if (!ok) throw new Error(`FAIL ${label}${detail === undefined ? '' : ' — ' + detail}`)
  console.log(`ok   ${label}`)
}

// 1. First instance: open a draft and write something into it.
let draftId = null
let subject = null
{
  const { routes, ctx } = harness()
  const first = await import('../index.js?instance=1')
  first.apply(ctx)
  const draftRoute = routes.get('/maildesk/api/draft')
  check('the draft route is mounted', typeof draftRoute === 'function')

  const opened = await hit(draftRoute, 'POST', { op: 'open', mode: 'new' })
  check('a draft opens', opened.json.ok === true && opened.json.draft && typeof opened.json.draft.id === 'string', JSON.stringify(opened.json).slice(0, 120))
  draftId = opened.json.draft.id
  const saved = await hit(draftRoute, 'POST', {
    op: 'save',
    draftId,
    patch: {
      to: [{ name: 'Someone', email: 'someone@example.com' }],
      subject: 'Draft that must survive',
      bodyHtml: '<p>First line.</p><p>Second paragraph.</p>',
    },
  })
  check('the draft saves', saved.json.ok === true && saved.json.draft.subject === 'Draft that must survive', JSON.stringify(saved.json).slice(0, 140))
  subject = saved.json.draft.subject
  check('the draft kept its subject', subject === 'Draft that must survive', String(subject))

  const onDisk = JSON.parse(readFileSync(DRAFTS, 'utf8'))
  check('the draft was written to disk immediately', Array.isArray(onDisk) && onDisk.some((d) => d.id === draftId), JSON.stringify(onDisk).slice(0, 120))
  check('the body text made it', onDisk.some((d) => String(d.bodyHtml).includes('Second paragraph')), 'body missing')
}

// 2. A fresh instance pointed at the same file: a restart, in miniature.
{
  const { routes, ctx } = harness()
  const second = await import('../index.js?instance=2')
  second.apply(ctx)
  const draftRoute = routes.get('/maildesk/api/draft')
  const listed = await hit(draftRoute, 'POST', { op: 'list' })
  const drafts = listed.json.drafts || []
  check('the restarted host lists the draft', drafts.some((d) => d.id === draftId), JSON.stringify(drafts).slice(0, 160))
  const restored = drafts.find((d) => d.id === draftId)
  check('the restored draft keeps its recipients', restored && String(restored.to?.[0]?.email || '').includes('someone@example.com'), JSON.stringify(restored && restored.to))
  check('the restored draft keeps its body', restored && String(restored.bodyHtml).includes('Second paragraph'))

  // Discarding it must clear the file too, or the next restart resurrects it.
  const discarded = await hit(draftRoute, 'POST', { op: 'discard', draftId })
  check('discard removes it', discarded.json.ok === true && discarded.json.discarded === true)
  const after = JSON.parse(readFileSync(DRAFTS, 'utf8'))
  check('the file no longer holds the discarded draft', Array.isArray(after) && !after.some((d) => d.id === draftId), JSON.stringify(after).slice(0, 120))
}

console.log(`\n${checks} assertions passed`)
