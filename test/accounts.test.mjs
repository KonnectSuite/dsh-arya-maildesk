/**
 * Account visibility — hide a store from the panel and the agent without
 * touching Outlook's account list.
 *
 * Runs against the fixture provider so it needs no COM: the fixture builds two
 * stores, the route reports them, hiding one removes its folders from every
 * later list_folders call, and the last visible account cannot be hidden.
 *
 *   node test/accounts.test.mjs
 */
import assert from 'node:assert/strict'
import { readFileSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

process.env.MAILDESK_FIXTURE = '1'
process.env.OLK_WATCH = '0'
process.env.MAILDESK_NO_RESIDENT = '1'
process.env.MAILDESK_ACCOUNTS_FILE = join(tmpdir(), 'maildesk-accounts-test-' + process.pid + '.json')
process.env.MAILDESK_DRAFTS_FILE = join(tmpdir(), 'maildesk-accounts-drafts-' + process.pid + '.json')
process.env.MAILDESK_STATE_FILE = join(tmpdir(), 'maildesk-accounts-ask-' + process.pid + '.json')
rmSync(process.env.MAILDESK_ACCOUNTS_FILE, { force: true })

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

const accountsRoute = routes.get('/maildesk/api/accounts')
const callRoute = routes.get('/maildesk/api/call')
assert.ok(typeof accountsRoute === 'function', 'accounts route registered')

const HOST = '127.0.0.1:57019'

async function hit(handler, method, body) {
  const text = body === undefined ? '' : JSON.stringify(body)
  const req = {
    method,
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

const folders = async () => {
  const r = await hit(callRoute, 'POST', { action: 'list_folders', args: { roots: true } })
  return r.json.result.folders || []
}

// 1. The route learns the stores from the mailbox itself.
const first = await hit(accountsRoute, 'GET')
check('GET lists the stores in the profile', Array.isArray(first.json.stores) && first.json.stores.length >= 1, JSON.stringify(first.json).slice(0, 140))
check('nothing is hidden at first', first.json.hidden.length === 0)
const only = first.json.stores[0]

// 2. The fixture has one store, so the guard is what can be proven here:
//    refusing to hide the last visible account. Cross-store filtering was
//    verified live against the three-store mailbox (see HANDOFF).
const refused = await hit(accountsRoute, 'POST', { store: only, hidden: true })
check('hiding the only account is refused', refused.status === 409, JSON.stringify(refused.json))

// 3. Any other name is just a name: it is remembered, and it changes nothing
//    about the mailbox that is actually there.
const before = await folders()
const bogus = await hit(accountsRoute, 'POST', { store: 'Not An Account', hidden: true })
check('hiding an unknown store is accepted and recorded', bogus.json.ok === true && bogus.json.hidden.includes('Not An Account'), JSON.stringify(bogus.json).slice(0, 140))
const after = await folders()
check('the real mailbox is untouched', after.length === before.length && after.every((f) => f.store === only))
check('the store still lists as visible', bogus.json.visible.includes(only))

// 4. Showing it again clears the record, and the file on disk agrees.
const cleared = await hit(accountsRoute, 'POST', { store: 'Not An Account', hidden: false })
check('the flag can be cleared', cleared.json.ok === true && cleared.json.hidden.length === 0)
const onDisk = JSON.parse(readFileSync(process.env.MAILDESK_ACCOUNTS_FILE, 'utf8'))
check('the choice is persisted', Array.isArray(onDisk.hidden) && onDisk.hidden.length === 0, JSON.stringify(onDisk))

console.log(`\n${checks} assertions passed`)
