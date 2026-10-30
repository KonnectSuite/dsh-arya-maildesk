/**
 * Model selection for the in-panel assistant.
 *
 * The card can only name or change the model if the host exposes the harness
 * catalog; this proves the route maps it, validates input, and creates the
 * session when a model is chosen before the first question.
 *
 *   node test/models.test.mjs
 */
import assert from 'node:assert/strict'
import { rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

process.env.OLK_WATCH = '0'
process.env.MAILDESK_NO_RESIDENT = '1'
process.env.MAILDESK_STATE_FILE = join(tmpdir(), 'maildesk-models-ask-' + process.pid + '.json')
process.env.MAILDESK_DRAFTS_FILE = join(tmpdir(), 'maildesk-models-drafts-' + process.pid + '.json')
process.env.MAILDESK_ACCOUNTS_FILE = join(tmpdir(), 'maildesk-models-accounts-' + process.pid + '.json')
rmSync(process.env.MAILDESK_STATE_FILE, { force: true })

const seen = { created: 0, selected: [] }

const controller = {
  async create() {
    seen.created += 1
    return { sessionId: 'sess-' + seen.created }
  },
  async rename() { return { ok: true } },
  async modelCatalog() {
    return {
      default: { provider: 'deepseek-account', model: 'deepseek-flash', reasoningEffort: 'high' },
      groups: [
        {
          id: 'deepseek-account',
          name: 'DeepSeek Account',
          models: [
            { id: 'deepseek-flash', name: 'DeepSeek-V4.1-Flash', reasoning: { efforts: [{ id: 'off', name: 'Off' }, { id: 'high', name: 'High' }], defaultEffort: 'high' } },
            { id: 'deepseek-v4-pro', name: 'DeepSeek-V4-Pro' },
          ],
        },
      ],
      failures: [],
    }
  },
  async selectModel(request) {
    seen.selected.push(request)
    return { selected: { provider: request.provider, model: request.model, ...(request.reasoningEffort === undefined ? {} : { reasoningEffort: request.reasoningEffort }) } }
  },
}

const plugin = await import('../index.js')

const routes = new Map()
const ctx = {
  tools: { register() {} },
  effect(fn) { const dispose = fn(); return typeof dispose === 'function' ? dispose : () => {} },
  get(name) { return name === 'sessionController' ? controller : undefined },
  inject(names, callback) {
    if (names.includes('webServer')) {
      callback({ webServer: { register(spec) { routes.set(spec.path, spec.handler); return () => routes.delete(spec.path) } } })
    }
  },
}
plugin.apply(ctx)

const modelsRoute = routes.get('/maildesk/api/models')
assert.ok(typeof modelsRoute === 'function', 'models route registered')

const HOST = '127.0.0.1:57031'
async function hit(method, body) {
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
  await modelsRoute(req, res)
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

// 1. The catalog is passed through with what the picker needs.
{
  const r = await hit('GET')
  check('GET answers with the catalog', r.json.ok === true && Array.isArray(r.json.groups), JSON.stringify(r.json).slice(0, 140))
  check('the selected model is named', r.json.selected.provider === 'deepseek-account' && r.json.selected.model === 'deepseek-flash', JSON.stringify(r.json.selected))
  const group = r.json.groups[0]
  check('the provider group survives', group.id === 'deepseek-account' && group.name === 'DeepSeek Account', JSON.stringify(group).slice(0, 120))
  check('model names and efforts survive', group.models[0].name === 'DeepSeek-V4.1-Flash' && group.models[0].reasoning.efforts.length === 2, JSON.stringify(group.models[0]).slice(0, 160))
  check('a model without efforts still lists', group.models[1].id === 'deepseek-v4-pro' && group.models[1].reasoning === undefined)
}

// 2. Choosing a model creates the session first, then selects on it.
{
  const r = await hit('POST', { provider: 'deepseek-account', model: 'deepseek-v4-pro', reasoningEffort: 'max' })
  check('POST selects the model', r.json.ok === true && r.json.selected.model === 'deepseek-v4-pro' && r.json.selected.reasoningEffort === 'max', JSON.stringify(r.json))
  check('the session existed before the selection', seen.created === 1, String(seen.created))
  check('the controller saw the request', seen.selected.length === 1 && seen.selected[0].sessionId === 'sess-1', JSON.stringify(seen.selected))
}

// 3. Bad input is refused rather than guessed.
{
  const missing = await hit('POST', { provider: 'deepseek-account' })
  check('a model is required', missing.status === 400 && missing.json.ok === false, JSON.stringify(missing.json))
  const method = await hit('DELETE')
  check('other methods are refused', method.status === 405)
}

console.log(`\n${checks} assertions passed`)
