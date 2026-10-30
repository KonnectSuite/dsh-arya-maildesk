// Split lib/olk.ps1 into "library + resident loop" without touching the switch.
//   node scripts/split-bridge.mjs lib/olk.ps1
import { readFileSync, writeFileSync } from 'node:fs'

const path = process.argv[2]
const src = readFileSync(path, 'utf8').split(/\r?\n/)

const find = (needle, from = 0) => {
  for (let i = from; i < src.length; i++) if (src[i].trim() === needle) return i
  return -1
}

const entry = find('try {', 120)            // the one-shot entry block
const connect = find('$ns = $outlook.GetNamespace(\'MAPI\')')
const sw = find('switch ($action) {')
if (entry < 0 || connect < 0 || sw < 0) { console.error('markers missing', { entry, connect, sw }); process.exit(1) }

// 1. entry .. just before the COM try  ->  function header
const head = src.slice(0, entry)
const comTry = src.findIndex((l, i) => i > entry && l.trim() === 'try {')
if (comTry < 0) { console.error('COM try not found'); process.exit(1) }

// 2. the switch stays; the region between the COM try and the switch is the connect code
const connectBody = src.slice(comTry, connect)          // includes the final `  }` of the catch chain
const switchBody = src.slice(sw, src.length)

// drop the old tail after the switch's closing brace is handled by keeping switchBody,
// then trimming the old entry tail and re-adding functions.
const tailStart = switchBody.findIndex((l) => l.trim() === '$out.status = \'ok\'')
if (tailStart < 0) { console.error('tail marker missing'); process.exit(1) }
const cases = switchBody.slice(0, tailStart)             // switch (...) { ...cases...  (no closing brace yet)

const out = [
  ...head,
  'function Connect-Outlook {',
  ...connectBody,
  '  $ns = $outlook.GetNamespace(\'MAPI\')',
  '  return $ns',
  '}',
  '',
  'function Invoke-OlkAction($ns, $action, $args_) {',
  '  $out = [ordered]@{}',
  '  try {',
  ...cases,
  '    $out.status = \'ok\'',
  '  } catch {',
  '    $out.status = \'error\'',
  '    $out.error = $_.Exception.Message',
  '  }',
  '  return $out',
  '}',
  '',
  'function Read-StdinAll {',
  '  $reader = New-Object System.IO.StreamReader([System.Console]::OpenStandardInput(), (New-Object System.Text.UTF8Encoding($false)))',
  '  return $reader.ReadToEnd()',
  '}',
  '',
  'if ($env:OLK_RESIDENT -eq \'1\') {',
  '  # Resident mode: connect once, then one JSON request per line on stdin and one',
  '  # JSON answer per line on stdout. The host keeps this process alive, so a folder',
  '  # click costs COM work instead of PowerShell startup.',
  '  try { $ns = Connect-Outlook } catch {',
  '    [Console]::Out.WriteLine(([ordered]@{ status = \'error\'; error = $_.Exception.Message } | ConvertTo-Json -Compress))',
  '    [Console]::Out.Flush()',
  '    exit 1',
  '  }',
  '  $stdin = New-Object System.IO.StreamReader([System.Console]::OpenStandardInput(), (New-Object System.Text.UTF8Encoding($false)))',
  '  $stdout = New-Object System.IO.StreamWriter([System.Console]::OpenStandardOutput(), (New-Object System.Text.UTF8Encoding($false)))',
  '  $stdout.AutoFlush = $true',
  '  while ($true) {',
  '    $line = $stdin.ReadLine()',
  '    if ($null -eq $line) { break }',
  '    if ($line.Trim() -eq \'\') { continue }',
  '    $req = $null',
  '    try { $req = $line | ConvertFrom-Json } catch { $req = $null }',
  '    if ($null -eq $req) { $stdout.WriteLine(\'{"status":"error","error":"request is not JSON"}\'); continue }',
  '    $res = Invoke-OlkAction $ns ([string]$req.action) ($req.args)',
  '    $stdout.WriteLine(($res | ConvertTo-Json -Depth 6 -Compress))',
  '  }',
  '  exit 0',
  '}',
  '',
  '# One-shot mode: OLK_ACTION + one JSON object on stdin, one JSON object out.',
  '$action = $env:OLK_ACTION',
  '$args_ = @{}',
  '$raw = Read-StdinAll',
  'if ($raw) { $args_ = $raw | ConvertFrom-Json }',
  'try {',
  '  $ns = Connect-Outlook',
  '  Out-Json (Invoke-OlkAction $ns $action $args_)',
  '} catch {',
  '  Out-Json ([ordered]@{ status = \'error\'; error = $_.Exception.Message })',
  '}',
]

writeFileSync(path, out.join('\n'))
console.log(`rewrote ${path}: ${src.length} -> ${out.length} lines`)
