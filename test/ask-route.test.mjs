/**
 * In-panel assistant route — the host half of "talk to the agent inside Mail".
 *
 * This drives the real `/maildesk/api/ask` route with a stub session
 * controller, so it proves the contract the panel depends on: one session is
 * created and reused, the prompt carries the panel's context and a user source,
 * only user/assistant text comes back, and the transport edges stay closed.
 * It cannot prove the harness's own session behaviour; that needs the GUI.
 *
 *   node test/ask-route.test.mjs
 */
import assert from 'node:assert/strict'
import { rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

process.env.MAILDESK_FIXTURE = '1'
process.env.OLK_WATCH = '0'
// A saved assistant session from the running app must not leak into the test:
// point the state file at a throwaway path and start clean.
process.env.MAILDESK_STATE_FILE = join(tmpdir(), 'maildesk-ask-test-' + process.pid + '.json')
rmSync(process.env.MAILDESK_STATE_FILE, { force: true })

const plugin = await import('../index.js')

const routes = new Map()
const seen = { created: 0, prompts: [], events: 0, cancels: [] }

// The transcript the stub returns; tests swap the turn markers to exercise
// "running" without touching anything else.
let records = [
  { event: { seq: 1, type: 'user/message', time: 1793340000000, data: { source: { kind: 'user', rpcId: 'r-1' }, message: { content: [{ type: 'text', text: 'anything new?' }] } } } },
  { event: { seq: 2, type: 'tool/call', time: 1793340001000, data: { name: 'maildesk_list_mail' } } },
  { event: { seq: 3, type: 'assistant/message', time: 1793340002000, data: { message: { content: [{ type: 'text', text: 'Five unread.' }, { type: 'tool-call', id: 'x' }] } } } },
  { event: { seq: 4, type: 'assistant/message', time: 1793340003000, data: { message: { content: [{ type: 'text', text: '   ' }] } } } },
  // A harness context injection: a user-role message with no rpc id.
  { event: { seq: 5, type: 'user/message', time: 1793340004000, data: { message: { content: [{ type: 'text', text: 'Current runtime context. This snapshot supersedes earlier ones.' }] } } } },
  { event: { seq: 6, type: 'user/message', time: 1793340005000, data: { message: { content: [{ type: 'text', text: '<system-reminder>skills</system-reminder>' }] } } } },
]

const controller = {
  async create() {
    seen.created += 1
    return { sessionId: 'sess-' + seen.created }
  },
  async prompt(request, signal) {
    seen.prompts.push({ request, signal })
    return { accepted: true }
  },
  async cancel(request) {
    seen.cancels.push(request)
    return { cancelled: true }
  },
  async events() {
    seen.events += 1
    return { cursor: 3, hasMore: false, records };
  },
}

const ctx = {
  tools: { register() {} },
  effect(fn) { const dispose = fn(); return typeof dispose === 'function' ? dispose : () => {} },
  get(name) { return name === 'sessionController' ? controller : undefined },
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

const ask = routes.get('/maildesk/api/ask')
assert.ok(typeof ask === 'function', 'ask route registered')

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

async function hit(method, body, headers) {
  const { req, res } = request(method, body, headers)
  await ask(req, res)
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

// 1. Before anything is asked there is no session and no invented transcript.
{
  const r = await hit('GET')
  check('GET before the first question is an empty transcript', r.json.ok === true && r.json.sessionId === null && Array.isArray(r.json.messages) && r.json.messages.length === 0)
  check('no session is created just by opening the pane', seen.created === 0, String(seen.created))
}

// 2. A question creates one session, reuses it afterwards, and carries the
//    panel's context plus a user-owned prompt source.
{
  const r = await hit('POST', { text: 'anything new?', folderPath: '\\kphung@frommconnect.com\\Inbox', selected: 'Schedule AN-X4' })
  check('POST accepts the question', r.json.ok === true && r.json.sessionId === 'sess-1', JSON.stringify(r.json))
  check('exactly one session is created', seen.created === 1, String(seen.created))
  check('exactly one prompt reached the session', seen.prompts.length === 1, String(seen.prompts.length))
  const { request: p, signal } = seen.prompts[0]
  check('the prompt names the session', p.sessionId === 'sess-1')
  check('the prompt carries the panel context', String(p.content[0].text).includes('folder \\kphung@frommconnect.com\\Inbox') && String(p.content[0].text).includes('Schedule AN-X4'), String(p.content[0].text))
  check('the prompt opens with the panel ground rules', String(p.content[0].text).startsWith('[mail assistant] ') && String(p.content[0].text).includes('Never call that tool'), String(p.content[0].text).slice(0, 80))
  check('the words follow the rules block', String(p.content[0].text).includes('\n\nanything new?'), String(p.content[0].text).slice(0, 160))
  check('the rules tell the agent to ask in plain text', String(p.content[0].text).includes('ask it in your reply'), String(p.content[0].text).slice(0, 200))
  check('the prompt carries a request id and a cancellation signal', typeof p.requestId === 'string' && p.requestId.length > 0 && signal !== null && typeof signal.throwIfAborted === 'function', String(p.requestId))

  const again = await hit('POST', { text: 'and now?' })
  check('a second question reuses the same session', again.json.sessionId === 'sess-1' && seen.created === 1, String(seen.created))
  check('a question without context sends the words alone', seen.prompts[1].request.content[0].text.endsWith('\n\nand now?') && !seen.prompts[1].request.content[0].text.includes('(mail panel'), String(seen.prompts[1].request.content[0].text).slice(-40))
}

// 3. The reply is read back as readable text only: the tool call is named,
//    blank assistant messages are dropped, order is preserved.
{
  const r = await hit('GET')
  check('GET returns the transcript', r.json.ok === true && Array.isArray(r.json.messages), JSON.stringify(r.json).slice(0, 120))
  check('user, tool and assistant rows survive', r.json.messages.length === 3, JSON.stringify(r.json.messages))
  check('roles are in order', r.json.messages[0].role === 'user' && r.json.messages[1].role === 'tool' && r.json.messages[2].role === 'assistant', JSON.stringify(r.json.messages.map((m) => m.role)))
  check('the tool row names the tool', r.json.messages[1].text === 'maildesk_list_mail', r.json.messages[1].text)
  check('the assistant text is the text block', r.json.messages[2].text === 'Five unread.', r.json.messages[2].text)
}

// 5. What the drawer adds: a time on every message, a running flag derived from
//    the turn markers, and the two control actions.
{
  const r = await hit('GET')
  check('GET reports the turn as finished', r.json.running === false, JSON.stringify(r.json.running))
  check('every message carries a millisecond timestamp', r.json.messages.every((m) => typeof m.at === 'number' && m.at > 1e12), JSON.stringify(r.json.messages.map((m) => m.at)))

  records = [{ event: { seq: 7, type: 'turn/start', time: 1793340010000, data: {} } }].concat(records)
  const running = await hit('GET')
  check('an open turn reads as running', running.json.running === true, JSON.stringify(running.json.running))
  records = records.concat([{ event: { seq: 8, type: 'turn/end', time: 1793340011000, data: {} } }])
  const finished = await hit('GET')
  check('closing the turn clears the flag', finished.json.running === false)

  const created = await hit('POST', { action: 'new' })
  check('new starts a second session', created.json.ok === true && created.json.sessionId === 'sess-2' && seen.created === 2, JSON.stringify(created.json))
  const afterNew = await hit('GET')
  check('the transcript belongs to the new session', afterNew.json.sessionId === 'sess-2')

  await hit('POST', { text: 'stop me' })
  const cancelled = await hit('POST', { action: 'cancel' })
  check('cancel reaches the controller for this session', cancelled.json.ok === true && seen.cancels.length === 1 && seen.cancels[0].sessionId === 'sess-2', JSON.stringify(seen.cancels))

  const unknown = await hit('POST', { action: 'explode' })
  check('an unknown action is refused', unknown.status === 400 && unknown.json.ok === false, JSON.stringify(unknown.json))
}

// 4. Transport edges: empty text, other methods, another origin.
{
  const empty = await hit('POST', { text: '   ' })
  check('empty text is refused', empty.status === 400 && empty.json.ok === false)
  const put = await hit('PUT', { text: 'x' })
  check('PUT is refused', put.status === 405)
  const foreign = await hit('POST', { text: 'x' }, { origin: 'https://example.com' })
  check('another origin is refused', foreign.status === 403)
}

console.log(`\n${checks} assertions passed`)
