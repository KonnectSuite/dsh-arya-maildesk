// Guards the two mistakes that only show up after publishing: the bundle patch
// naming a package that does not exist, and a shipped file that package.json
// forgets to list. Both are invisible to every other test here, because they
// exercise the code rather than the package.
//
// Run with: node test/bundle.test.mjs
import { readFileSync, existsSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

const root = dirname(dirname(fileURLToPath(import.meta.url)))
const pkg = JSON.parse(readFileSync(join(root, 'package.json'), 'utf8'))

let failures = 0
let checks = 0
function check(label, ok, detail) {
  checks += 1
  if (ok) {
    console.log(`ok   ${label}`)
  } else {
    failures += 1
    console.log(`FAIL ${label}${detail === undefined ? '' : ' — ' + detail}`)
  }
}

// 1. The patch's rows must name THIS package, scope included. The loader resolves
//    the string with import(); 'dsh-arya-maildesk' is not our name any more, so a
//    stale value installs fine and then fails to activate (v1.0.0 shipped that bug).
const patchPath = join(root, String(pkg.dsh?.bundle?.patch ?? 'cordis.patch.yml'))
check('bundle patch exists', existsSync(patchPath), patchPath)

if (existsSync(patchPath)) {
  const yml = readFileSync(patchPath, 'utf8')
  const names = [...yml.matchAll(/^\s+name:\s*['"]?([^'"\n]+?)['"]?\s*$/gm)].map((m) => m[1].trim())
  check('bundle patch declares at least one row name', names.length > 0)
  for (const name of names) {
    check(`row name is the package name (${name})`, name === pkg.name, `expected ${pkg.name}`)
  }
}

// 2. The published manifest must be scoped: first-party Arya packages are
//    @aryacode/dsh-<name>, and the market requires it.
check('package name is scoped', /^@[a-z0-9-]+\/dsh-/.test(pkg.name), pkg.name)

// 3. Everything the manifest points at has to exist AND ship, or the installed
//    copy is a package that references files it does not contain.
const entryPoints = [pkg.main, pkg.dsh?.bundle?.patch, 'lib/olk.ps1', 'client.js'].filter(Boolean)
for (const rel of entryPoints) {
  const relPath = String(rel).replace(/^\.\//, '')
  check(`entry point exists (${relPath})`, existsSync(join(root, relPath)), relPath)
  const shipped = Array.isArray(pkg.files) && pkg.files.some((f) => String(f).replace(/^\.\//, '') === relPath)
  check(`entry point is published (${relPath})`, shipped, 'missing from package.json files[]')
}

// 4. The host floor has to be declared, or the market card can neither confirm
//    nor block and installs onto hosts we never meant to support.
check('dsh.engines.dsh is declared', typeof pkg.dsh?.engines?.dsh === 'string' && pkg.dsh.engines.dsh.length > 0)

// 5. The browser bundle registers itself under an id, and the module table keys
//    every client bundle by package name. An unscoped id therefore registers
//    this module where nothing looks for it: the host half loads, the panel never
//    appears, and the web boot fails with
//    `client-modules: could not load "<name>"`. That shipped in 1.0.1 because the
//    dev profile linked the repository under the unscoped name, so the stale id
//    resolved by luck. Same class of bug as the patch row above, one layer down.
const clientPath = join(root, 'client.js')
if (existsSync(clientPath)) {
  const bundle = readFileSync(clientPath, 'utf8')
  const ids = [...bundle.matchAll(/__ModuleLoader__\.load\(\s*\{[^}]*?\bid:\s*['"]([^'"]+)['"]/g)].map((m) => m[1])
  check('client bundle declares a module id', ids.length > 0)
  for (const id of ids) {
    check(`client module id is the package name (${id})`, id === pkg.name, `expected ${pkg.name}`)
  }
}

console.log('')
if (failures > 0) {
  console.log(`FAILED: ${failures} of ${checks} checks`)
  process.exit(1)
}
console.log(`OK: ${checks} checks`)
