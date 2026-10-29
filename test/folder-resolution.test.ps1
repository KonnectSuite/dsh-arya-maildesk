# Offline tests for lib/olk.ps1: structure + Resolve-Folder behaviour.
# Runs with no Outlook installed: the folder object model is faked.
$ErrorActionPreference = 'Stop'
$bridge = Join-Path (Split-Path -Parent $PSScriptRoot) 'lib\olk.ps1'
$src = Get-Content -LiteralPath $bridge -Raw -Encoding UTF8
$fail = 0
function Ok($cond, $msg) { if ($cond) { "PASS  $msg" } else { "FAIL  $msg"; $script:fail++ } }

# ---- 1. structural: every function must be defined at script scope
$errs = $null; $toks = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($src, [ref]$toks, [ref]$errs)
Ok ($errs.Count -eq 0) "parse clean ($($errs.Count) errors)"
$fns = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true)
$names = @($fns | ForEach-Object { $_.Name })
Ok ($names -contains 'Resolve-Folder') "Resolve-Folder exists as a function"
$nested = @()
foreach ($f in $fns) {
  foreach ($g in $fns) {
    if ($f.Name -ne $g.Name -and $g.Extent.StartLineNumber -gt $f.Extent.StartLineNumber -and $g.Extent.StartLineNumber -le $f.Extent.EndLineNumber) {
      $nested += ("{0} is nested inside {1}" -f $g.Name, $f.Name)
    }
  }
}
$detail = if ($nested.Count -gt 0) { ' -> ' + ($nested -join '; ') } else { '' }
Ok ($nested.Count -eq 0) ("no function is nested inside another$detail")

# ---- 2. extract Resolve-Folder alone (brace-matched) and load it
$start = $src.IndexOf('function Resolve-Folder')
if ($start -lt 0) { "FAIL  cannot locate Resolve-Folder"; exit 1 }
$i = $src.IndexOf('{', $start); $d = 0; $end = -1
for ($j = $i; $j -lt $src.Length; $j++) {
  $c = $src[$j]
  if ($c -eq '{') { $d++ } elseif ($c -eq '}') { $d--; if ($d -eq 0) { $end = $j; break } }
}
$fn = $src.Substring($start, $end - $start + 1)
Invoke-Expression $fn

# ---- 3. fake folder object model
function New-Coll([object[]]$items) {
  $o = New-Object PSObject
  Add-Member -InputObject $o -MemberType NoteProperty -Name Items -Value @($items)
  Add-Member -InputObject $o -MemberType ScriptMethod -Name Item -Value {
    param($k)
    if ($k -is [int]) { return $this.Items[$k - 1] }
    foreach ($x in $this.Items) { if ($x.Name -ieq [string]$k) { return $x } }
    throw "no folder named '$k'"
  }
  Add-Member -InputObject $o -MemberType ScriptMethod -Name GetEnumerator -Value { $this.Items.GetEnumerator() }
  return $o
}
function New-Folder([string]$name, [object[]]$children = @()) {
  $f = New-Object PSObject
  Add-Member -InputObject $f -MemberType NoteProperty -Name Name -Value $name
  Add-Member -InputObject $f -MemberType NoteProperty -Name Folders -Value (New-Coll $children)
  return $f
}

$leaf = New-Folder 'Invoices for the October 2026 review cycle'
$abc = New-Folder 'Customer ABC' @($leaf)
$q4 = New-Folder 'Q4' @($abc)
$y26 = New-Folder '2026' @($q4)
$projects = New-Folder 'Projects' @($y26)
$inbox = New-Folder 'Inbox' @($projects)
$sent = New-Folder 'Sent Items'
$deleted = New-Folder 'Deleted Items'
$drafts = New-Folder 'Drafts'
$idFolder = New-Folder 'IdFolder'
$root = New-Folder 'mailbox@example.com' @($inbox, $sent, $deleted, $drafts)

$deepId = ('A' * 30) + '/' + ('B' * 40)   # >= 60 chars and contains '/', like base64
Ok ($deepId.Length -ge 60) "test id is >= 60 chars ($($deepId.Length))"

$ns = New-Object PSObject
Add-Member -InputObject $ns -MemberType NoteProperty -Name Folders -Value (New-Coll @($root))
Add-Member -InputObject $ns -MemberType NoteProperty -Name Defaults -Value @{ 6 = $inbox; 5 = $sent; 3 = $deleted; 16 = $drafts }
Add-Member -InputObject $ns -MemberType NoteProperty -Name Ids -Value @{ $deepId = $idFolder }
Add-Member -InputObject $ns -MemberType ScriptMethod -Name GetDefaultFolder -Value { param($n) return $this.Defaults[[int]$n] }
Add-Member -InputObject $ns -MemberType ScriptMethod -Name GetFolderFromID -Value {
  param($id)
  if ($this.Ids.ContainsKey([string]$id)) { return $this.Ids[[string]$id] }
  throw "id not found"
}

# ---- 4. behaviour
function Try-Name([string]$spec) {
  try { (Resolve-Folder $ns $spec).Name } catch { 'ERROR: ' + $_.Exception.Message }
}
function Expect([string]$spec, [string]$want, [string]$label) {
  $got = Try-Name $spec
  Ok ($got -eq $want) ("{0}: '{1}' -> '{2}'" -f $label, $spec, $got)
}

Expect '' 'Inbox' 'empty means Inbox'
Expect 'Inbox' 'Inbox' 'well-known Inbox'
Expect 'Sent' 'Sent Items' 'alias Sent'
Expect 'Deleted Items' 'Deleted Items' 'well-known Deleted Items'
Expect 'Inbox/Projects' 'Projects' 'one path segment'
Expect 'Inbox\Projects\2026' '2026' 'backslash path'
$deepPath = 'Inbox/Projects/2026/Q4/Customer ABC/Invoices for the October 2026 review cycle'
Ok ($deepPath.Length -ge 60) "deep path is >= 60 chars ($($deepPath.Length))"
Expect $deepPath 'Invoices for the October 2026 review cycle' 'DEEP PATH (>=60 chars) must not be mistaken for an entryId'
Expect $deepId 'IdFolder' 'entryId containing / resolves by id'

$bad = Try-Name 'Inbox/NoSuchFolder'
Ok ($bad -like 'ERROR:*') "missing folder throws (got: $bad)"

"---"
"failures: $fail"
if ($fail -gt 0) { exit 1 }
