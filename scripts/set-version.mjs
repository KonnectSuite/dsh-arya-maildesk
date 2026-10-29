// Rewrites only the "version" value in package.json.
//
// Upstream's package.json is written by ConvertTo-Json and carries its odd
// `":  "` spacing, so it is not a file to round-trip through JSON.parse /
// JSON.stringify — that would reformat every line of a file we do not own.
import { readFileSync, writeFileSync } from 'node:fs'

const wanted = process.argv[2]
if (wanted === undefined || !/^\d+\.\d+\.\d+$/.test(wanted)) {
  console.error('usage: node scripts/set-version.mjs <major.minor.patch>')
  process.exit(1)
}

const path = new URL('../package.json', import.meta.url)
const text = readFileSync(path, 'utf8')
const next = text.replace(/("version":\s*)"[^"]*"/, `$1"${wanted}"`)
if (next === text) {
  console.error('package.json has no "version" field to replace')
  process.exit(1)
}
writeFileSync(path, next)
console.log('version ->', JSON.parse(next).version)
