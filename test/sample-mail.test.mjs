/**
 * The sample-mailbox toggle — the one button that makes the panel useful on a
 * machine whose Outlook is the new WebView2 client (no COM surface at all).
 *
 * This test starts with NO fixture: the provider must answer `com`, fail the way
 * a COM-less machine fails, switch to the sample mailbox on request, serve that
 * mailbox through the same routes, and switch back.
 *
 *   node test/sample-mail.test.mjs
 */
import assert from 'node:assert/strict'

delete process.env.MAILDESK_FIXTURE
process.env.OLK_WATCH = '0'
// This suite imports the host with live COM, and the resident bridge would keep
// the process alive after the assertions; it exercises the one-shot path.
process.env.MAILDESK_NO_RESIDENT = '1'

const plugin = await import('../index.js')

const routes = new Map()
const ctx = {
  tools: { register() {} },
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

const health = routes.get('/maildesk/api/health')
const call = routes.get('/maildesk/api/call')
const fixtureRoute = routes.get('/maildesk/api/fixture')
assert.ok(typeof health === 'function', 'health route registered')
assert.ok(typeof call === 'function', 'call route registered')
assert.ok(typeof fixtureRoute === 'function', 'sample mailbox route registered')

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
    headers: {},
    setHeader(name, value) { this.headers[String(name).toLowerCase()] = value },
    end(payload) { this.payload = payload === undefined ? '' : String(payload) },
  }
  return { req, res }
}

async function hit(handler, method, body, headers) {
  const { req, res } = request(method, body, headers)
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

// 1. Live provider first: with no fixture env the panel must be told 'com', and
//    the mailbox call must either fail with the bridge's own explanation (a
//    machine with no classic Outlook) or answer from the real mailbox (a machine
//    that has one). Both are legitimate; an EMPTY SUCCESS is not, and that is
//    what this asserts on both kinds of machine.
let liveMailbox = false
{
  const h = await hit(health, 'GET')
  check('health reports the live provider', h.json.ok === true && h.json.provider === 'com', JSON.stringify(h.json))
  check('health says this is not sample mail', h.json.sample === false)
  const r = await hit(call, 'POST', { action: 'list_folders', args: { roots: true } })
  // The route answers 200 for bridge-level failures (a missing Outlook is not a
  // transport fault), so the signal is result.status, not the HTTP code.
  const verdict = JSON.stringify(r.json)
  liveMailbox = r.json.ok === true && r.json.result.status === 'ok'
  if (liveMailbox) {
    // This machine can be verified for real; the loud-failure path is not
    // reachable here, so say so instead of failing the suite.
    check(
      'this machine has classic Outlook, so the live provider answers with folders',
      Array.isArray(r.json.result.folders) && r.json.result.folders.length > 0,
      verdict.slice(0, 160),
    )
  } else {
    check(
      'with no COM the call fails loudly instead of returning an empty mailbox',
      r.json.ok === true &&
        r.json.result.status === 'error' &&
        !Array.isArray(r.json.result.folders) &&
        /classic desktop Outlook not found|Classic Outlook|not registered|80040154|spawn|EPERM/i.test(verdict),
      verdict.slice(0, 160),
    )
  }
}

// 2. Switch to the sample mailbox. Everything after this must behave exactly like
//    a real mailbox through the same routes.
{
  const s = await hit(fixtureRoute, 'POST', { on: true })
  check('the toggle accepts the switch', s.json.ok === true && s.json.provider === 'fixture', JSON.stringify(s.json))
  check('the toggle reports that this is sample mail', s.json.sample === true)

  const h = await hit(health, 'GET')
  check('health follows the switch', h.json.provider === 'fixture' && h.json.sample === true, JSON.stringify(h.json))

  const folders = await hit(call, 'POST', { action: 'list_folders', args: { maxDepth: 6 } })
  const list = folders.json.result.folders
  check('the sample mailbox has folders', Array.isArray(list) && list.length >= 8, String(list && list.length))
  check('the sample folders are labelled by path', list.some((f) => f.path === 'Inbox/Projects'))

  const inbox = list.find((f) => /^inbox$/i.test(f.name))
  const page = await hit(call, 'POST', { action: 'list_mail', args: { folder: inbox.entryId, limit: 5, preview: true } })
  const rows = page.json.result.mails
  check('the sample mailbox lists mail', Array.isArray(rows) && rows.length === 5, String(rows && rows.length))
  check('the rows look like real mail', typeof rows[0].subject === 'string' && rows[0].subject.length > 0 && /^\d{4}-\d{2}-\d{2} \d{2}:\d{2}$/.test(rows[0].received))

  // The tools share the provider, so an agent sees the same mailbox the panel does.
  const read = await hit(call, 'POST', { action: 'read_mail', args: { entryId: rows[0].entryId, format: 'html' } })
  check('reading a sample mail returns a body', read.json.result.status === 'ok' && String(read.json.result.body).length > 0)
}

// 3. Switch back. The panel must return to the honest failure, not a stale inbox.
{
  const s = await hit(fixtureRoute, 'POST', { on: false })
  check('the toggle switches back to live', s.json.ok === true && s.json.provider === 'com', JSON.stringify(s.json))
  const h = await hit(health, 'GET')
  check('health is back to the live provider', h.json.provider === 'com' && h.json.sample === false)
  const r = await hit(call, 'POST', { action: 'list_folders', args: { roots: true } })
  check(
    liveMailbox ? 'the real mailbox answers again after switching off' : 'the live provider fails again after switching back',
    liveMailbox ? r.json.result.status === 'ok' : r.json.result.status === 'error',
  )
}

// 4. The route's own guards: POST only, same origin.
{
  const get = await hit(fixtureRoute, 'GET')
  check('the toggle refuses GET', get.status === 405)
  const foreign = await hit(fixtureRoute, 'POST', { on: true }, { origin: 'https://example.com' })
  check('the toggle refuses another origin', foreign.status === 403)
  const bad = await hit(fixtureRoute, 'POST', undefined, { 'content-type': 'application/json' })
  check('the toggle rejects a body that is not JSON', bad.status === 400 && bad.json.ok === false, JSON.stringify(bad.json))
}

console.log(`\n${checks} assertions passed`)
