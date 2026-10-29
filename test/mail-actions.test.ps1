# Offline behaviour tests for the five new bridge actions.
# No Outlook: the COM object model is faked. The REAL case bodies are extracted
# from lib/olk.ps1 and executed, so the tests exercise shipped code.
$ErrorActionPreference = 'Stop'
$bridge = Join-Path (Split-Path -Parent $PSScriptRoot) 'lib\olk.ps1'
$src = Get-Content -LiteralPath $bridge -Raw -Encoding UTF8
$script:fail = 0
function Ok($cond, $msg) { if ($cond) { "PASS  $msg" } else { "FAIL  $msg"; $script:fail++ } }

function Get-BraceText([int]$start) {
  $i = $src.IndexOf('{', $start); $d = 0
  for ($j = $i; $j -lt $src.Length; $j++) {
    $c = $src[$j]
    if ($c -eq '{') { $d++ } elseif ($c -eq '}') { $d--; if ($d -eq 0) { return $src.Substring($start, $j - $start + 1) } }
  }
  throw 'unbalanced braces'
}
function Get-CaseText([string]$name) {
  $start = $src.IndexOf("    '$name' {")
  if ($start -lt 0) { throw "case '$name' not found" }
  return (Get-BraceText $start)
}
Invoke-Expression (Get-BraceText ($src.IndexOf('function Resolve-Folder')))

# ---------------- fake Outlook object model ----------------
# A folder collection must be BOTH enumerable (list_folders walks it) and
# indexable by name (Resolve-Folder uses .Item('Name')), which a plain PSObject
# cannot be: foreach needs a real IEnumerable. So define one.
if (-not ('FakeColl' -as [type])) {
  Add-Type -ReferencedAssemblies ([System.Management.Automation.PSObject].Assembly.Location) -TypeDefinition @"
using System; using System.Collections;
public class FakeColl : IEnumerable {
  public ArrayList Items = new ArrayList();
  public FakeColl(ArrayList items) { Items = items; }
  public int Count { get { return Items.Count; } }
  public object Item(object k) {
    if (k is int) {
      int i = (int)k;
      if (i < 1 || i > Items.Count) throw new Exception("index " + i + " out of range");
      return Items[i - 1];
    }
    string name = Convert.ToString(k);
    foreach (object x in Items) {
      var p = System.Management.Automation.PSObject.AsPSObject(x).Properties["Name"];
      if (p != null && string.Equals(Convert.ToString(p.Value), name, StringComparison.OrdinalIgnoreCase)) return x;
    }
    throw new Exception("no folder named '" + name + "'");
  }
  public IEnumerator GetEnumerator() { return Items.GetEnumerator(); }
}
"@
}
function New-Coll([System.Collections.ArrayList]$list) {
  # The unary comma prevents New-Object from enumerating the ArrayList into
  # several constructor arguments: a 2-folder collection would otherwise fail.
  # -NoEnumerate: a function's output pipeline unrolls IEnumerable, so a plain
  # return would hand the caller the FOLDER, not the collection.
  return (Write-Output -NoEnumerate (New-Object FakeColl -ArgumentList (, $list)))
}
function New-CollOld([System.Collections.ArrayList]$list) {
  $c = New-Object PSObject
  Add-Member -InputObject $c -MemberType NoteProperty -Name Items -Value $list
  Add-Member -InputObject $c -MemberType ScriptProperty -Name Count -Value { return $this.Items.Count }
  Add-Member -InputObject $c -MemberType ScriptMethod -Name Item -Value {
    param($k)
    if ($k -is [int]) { if ($k -lt 1 -or $k -gt $this.Items.Count) { throw "index $k out of range" }; return $this.Items[$k - 1] }
    foreach ($x in $this.Items) { if ($x.Name -ieq [string]$k) { return $x } }
    throw "no folder named '$k'"
  }
  Add-Member -InputObject $c -MemberType ScriptMethod -Name GetEnumerator -Value { return $this.Items.GetEnumerator() }
  return $c
}

function New-Items($folder) {
  $o = New-Object PSObject
  Add-Member -InputObject $o -MemberType NoteProperty -Name Owner -Value $folder
  Add-Member -InputObject $o -MemberType ScriptProperty -Name Count -Value { return $this.Owner.List.Count }
  Add-Member -InputObject $o -MemberType ScriptMethod -Name Item -Value {
    param($i)
    if ($i -lt 1 -or $i -gt $this.Owner.List.Count) { throw "item index $i out of range" }
    return $this.Owner.List[$i - 1]
  }
  Add-Member -InputObject $o -MemberType ScriptMethod -Name Sort -Value {
    param($field, $descending)
    $arr = @($this.Owner.List)
    $sorted = $arr | Sort-Object -Property @{ Expression = { $_.$field } } -Descending:([bool]$descending)
    $this.Owner.List.Clear()
    foreach ($x in $sorted) { [void]$this.Owner.List.Add($x) }
  }
  Add-Member -InputObject $o -MemberType ScriptMethod -Name Restrict -Value {
    param($filter)
    if ([string]$filter -notmatch 'UnRead') { throw "fake Restrict does not understand '$filter'" }
    $keep = New-Object System.Collections.ArrayList
    foreach ($x in $this.Owner.List) { if ($x.UnRead) { [void]$keep.Add($x) } }
    $view = New-Object PSObject
    Add-Member -InputObject $view -MemberType NoteProperty -Name Owner -Value $this.Owner
    Add-Member -InputObject $view -MemberType NoteProperty -Name Subset -Value $keep
    Add-Member -InputObject $view -MemberType ScriptProperty -Name Count -Value { return $this.Subset.Count }
    Add-Member -InputObject $view -MemberType ScriptMethod -Name Item -Value {
      param($i)
      if ($i -lt 1 -or $i -gt $this.Subset.Count) { throw "item index $i out of range" }
      return $this.Subset[$i - 1]
    }
    Add-Member -InputObject $view -MemberType ScriptMethod -Name Sort -Value {
      param($field, $descending)
      $sorted = @($this.Subset) | Sort-Object -Property @{ Expression = { $_.$field } } -Descending:([bool]$descending)
      $this.Subset.Clear()
      foreach ($x in $sorted) { [void]$this.Subset.Add($x) }
    }
    Add-Member -InputObject $view -MemberType ScriptMethod -Name Restrict -Value { throw 'double Restrict not faked' }
    return $view
  }
  return $o
}

function New-Folder([string]$name, [string]$path) {
  $f = New-Object PSObject
  Add-Member -InputObject $f -MemberType NoteProperty -Name Name -Value $name
  Add-Member -InputObject $f -MemberType NoteProperty -Name FolderPath -Value $path
  Add-Member -InputObject $f -MemberType NoteProperty -Name EntryID -Value ("fid" + $path)
  Add-Member -InputObject $f -MemberType NoteProperty -Name List -Value (New-Object System.Collections.ArrayList)
  $children = New-Object System.Collections.ArrayList
  Add-Member -InputObject $f -MemberType NoteProperty -Name Children -Value $children
  Add-Member -InputObject $f -MemberType ScriptProperty -Name Items -Value { return (New-Items $this) }
  # A stored NoteProperty, not a ScriptProperty: a value returned through a
  # scriptblock output pipeline is unrolled when it implements IEnumerable,
  # which turns a one-child collection into that child. FakeColl holds the
  # same ArrayList, so children added later are still visible.
  Add-Member -InputObject $f -MemberType NoteProperty -Name Folders -Value (New-Object FakeColl -ArgumentList (, $children))
  Add-Member -InputObject $f -MemberType ScriptProperty -Name UnReadItemCount -Value {
    $n = 0; foreach ($x in $this.List) { if ($x.UnRead) { $n++ } }; return $n
  }
  return $f
}

function New-Mail([string]$id, [string]$subj, [string]$sender, [string]$received, [bool]$unread, [int]$cls, [string]$body, [int]$att) {
  $m = New-Object PSObject
  Add-Member -InputObject $m -MemberType NoteProperty -Name EntryID -Value $id
  Add-Member -InputObject $m -MemberType NoteProperty -Name Subject -Value $subj
  Add-Member -InputObject $m -MemberType NoteProperty -Name SenderName -Value $sender
  Add-Member -InputObject $m -MemberType NoteProperty -Name SenderEmailAddress -Value ($sender + '@example.com')
  Add-Member -InputObject $m -MemberType NoteProperty -Name ReceivedTime -Value ([datetime]$received)
  Add-Member -InputObject $m -MemberType NoteProperty -Name UnRead -Value $unread
  Add-Member -InputObject $m -MemberType NoteProperty -Name Class -Value $cls
  Add-Member -InputObject $m -MemberType NoteProperty -Name Body -Value $body
  Add-Member -InputObject $m -MemberType NoteProperty -Name Attachments -Value ([pscustomobject]@{ Count = $att })
  Add-Member -InputObject $m -MemberType NoteProperty -Name Deleted -Value $false
  Add-Member -InputObject $m -MemberType NoteProperty -Name Saved -Value $false
  Add-Member -InputObject $m -MemberType NoteProperty -Name Owner -Value $null
  Add-Member -InputObject $m -MemberType ScriptMethod -Name Save -Value { $this.Saved = $true }
  Add-Member -InputObject $m -MemberType ScriptMethod -Name Delete -Value {
    $this.Deleted = $true
    if ($this.Owner) { [void]$this.Owner.List.Remove($this) }
  }
  Add-Member -InputObject $m -MemberType ScriptMethod -Name Move -Value {
    param($dest)
    $new = New-Object PSObject
    Add-Member -InputObject $new -MemberType NoteProperty -Name EntryID -Value ($this.EntryID + '+moved')
    $this.Deleted = $true
    if ($this.Owner) { [void]$this.Owner.List.Remove($this) }
    [void]$dest.List.Add($new)
    return $new
  }
  return $m
}

function New-Mailbox {
  $inbox = New-Folder 'Inbox' '\Inbox'
  $sent = New-Folder 'Sent Items' '\Sent Items'
  $deleted = New-Folder 'Deleted Items' '\Deleted Items'
  $drafts = New-Folder 'Drafts' '\Drafts'
  $projects = New-Folder 'Projects' '\Inbox\Projects'
  [void]$inbox.Children.Add($projects)
  $root = New-Folder 'mailbox@example.com' '\'
  foreach ($f in @($inbox, $sent, $deleted, $drafts)) { [void]$root.Children.Add($f) }

  $mails = @(
    (New-Mail 'm1' 'Alpha report'  'Avery'  '2026-10-20 09:00' $true  43 ("line one`r`nline two " + ('x' * 300)) 0),
    (New-Mail 'm2' 'Beta invoice'  'Blake'  '2026-10-21 10:30' $false 43 'body two' 2),
    (New-Mail 'm3' 'Gamma meeting' 'Casey'  '2026-10-22 11:00' $true  53 'body three' 0),
    (New-Mail 'm4' 'Delta note'    'Dana'   '2026-10-23 12:00' $false 1  'body four' 0),
    (New-Mail 'm5' 'Epsilon'       'Emery'  '2026-10-24 13:00' $false 43 'body five' 0)
  )
  foreach ($m in $mails) { $m.Owner = $inbox; [void]$inbox.List.Add($m) }
  [void]$projects.List.Add((New-Mail 'p1' 'Project item' 'Flynn' '2026-10-19 08:00' $false 43 'project body' 0))
  $projects.List[0].Owner = $projects

  $ns = New-Object PSObject
  Add-Member -InputObject $ns -MemberType NoteProperty -Name Mailbox -Value @{ Inbox = $inbox; Sent = $sent; Deleted = $deleted; Drafts = $drafts; Projects = $projects; Root = $root }
  $rootColl = [System.Collections.ArrayList]@($root)
  Add-Member -InputObject $ns -MemberType NoteProperty -Name Folders -Value (New-Object FakeColl -ArgumentList (, $rootColl))
  Add-Member -InputObject $ns -MemberType ScriptMethod -Name GetDefaultFolder -Value {
    param($n)
    switch ([int]$n) { 6 { return $this.Mailbox.Inbox } 5 { return $this.Mailbox.Sent } 3 { return $this.Mailbox.Deleted } 16 { return $this.Mailbox.Drafts } }
    throw "default folder $n not faked"
  }
  Add-Member -InputObject $ns -MemberType ScriptMethod -Name GetFolderFromID -Value { throw 'entryId lookup not faked' }
  Add-Member -InputObject $ns -MemberType ScriptMethod -Name GetItemFromID -Value {
    param($id)
    $all = @($this.Mailbox.Inbox) + @($this.Mailbox.Projects) + @($this.Mailbox.Sent) + @($this.Mailbox.Deleted) + @($this.Mailbox.Drafts)
    foreach ($f in $all) { foreach ($x in $f.List) { if ([string]$x.EntryID -eq [string]$id) { return $x } } }
    throw "item not found: $id"
  }
  return $ns
}

function Run-Case([string]$caseName, [hashtable]$a, $nsObj) {
  $ns = $nsObj
  $args_ = $a
  $out = [ordered]@{}
  $text = Get-CaseText $caseName
  try {
    Invoke-Expression ("switch ('" + $caseName + "') {`n" + $text + "`n}")
    return @{ ok = $true; out = $out }
  } catch {
    Write-Host ("    [case " + $caseName + " threw] " + $_.Exception.Message)
    return @{ ok = $false; error = $_.Exception.Message; out = $out }
  }
}

function J($r) { return ($r.out | ConvertTo-Json -Depth 6 -Compress) }

# ---------------- list_mail ----------------
$ns = New-Mailbox; $r = Run-Case 'list_mail' @{} $ns
Ok $r.ok "list_mail runs"
Ok ($r.out.total -eq 5) "list_mail total counts every item in the folder (5), incl. the non-mail one"
Ok ($r.out.returned -eq 4) "list_mail returns only mail/meeting items (4 of 5)"
Ok ($r.out.mails[0].entryId -eq 'm5') "list_mail sorts newest first by default"
Ok ($r.out.mails[0].preview -eq '') "no preview unless asked"
Ok ($r.out.unreadInFolder -eq 2) "list_mail reports unread count for the folder (2)"
Ok ($r.out.folderPath -eq '\Inbox') "list_mail reports the folder path"

$ns = New-Mailbox; $r = Run-Case 'list_mail' @{ limit = 2 } $ns
Ok ($r.out.returned -eq 2 -and $r.out.mails[0].entryId -eq 'm5' -and $r.out.mails[1].entryId -eq 'm3') "limit=2 returns the two newest mail items (m5, m3)"

$ns = New-Mailbox; $r = Run-Case 'list_mail' @{ limit = 2; offset = 2 } $ns
Ok ($r.out.returned -eq 2 -and $r.out.mails[0].entryId -eq 'm3' -and $r.out.mails[1].entryId -eq 'm2') "offset=2 limit=2 lands on m3, m2 (m4 is skipped as non-mail)"

$ns = New-Mailbox; $r = Run-Case 'list_mail' @{ descending = $false } $ns
Ok ($r.out.mails[0].entryId -eq 'm1') "descending=false puts the oldest first"

$ns = New-Mailbox; $r = Run-Case 'list_mail' @{ sort = 'subject'; descending = $false } $ns
Ok ($r.out.mails[0].subject -eq 'Alpha report') "sort=subject ascending starts at 'Alpha report'"

$ns = New-Mailbox; $r = Run-Case 'list_mail' @{ unreadOnly = $true } $ns
Ok ($r.out.returned -eq 2 -and ($r.out.mails | ForEach-Object { $_.entryId }) -join ',' -eq 'm3,m1') "unreadOnly returns the 2 unread mails"

$ns = New-Mailbox; $r = Run-Case 'list_mail' @{ preview = $true } $ns
Ok ($r.out.mails[0].preview.Length -le 160) "preview is capped at 160 chars ($($r.out.mails[0].preview.Length))"
Ok ($r.out.mails[3].preview -notmatch "`n" -and $r.out.mails[3].preview -notmatch "`r") "preview collapses newlines"

$ns = New-Mailbox; $r = Run-Case 'list_mail' @{ folder = 'Inbox/Projects' } $ns
Ok ($r.ok -and $r.out.returned -eq 1 -and $r.out.mails[0].entryId -eq 'p1') "folder path 'Inbox/Projects' resolves"

$ns = New-Mailbox; $r = Run-Case 'list_mail' @{ limit = 999 } $ns
Ok ($r.ok -and $r.out.returned -eq 4) "limit is clamped, not rejected"

# ---------------- move_mail ----------------
$ns = New-Mailbox; $r = Run-Case 'move_mail' @{ entryIds = @('m1', 'm2'); destFolder = 'Deleted Items' } $ns
Ok ($r.out.moved -eq 2) "move_mail moved 2"
Ok ($r.out.newEntryIds.Count -eq 2 -and $r.out.newEntryIds[0] -ne 'm1') "move_mail returns the NEW entryIds"
Ok ($ns.Mailbox.Inbox.List.Count -eq 3) "moved items left the source folder (3 remain of 5)"
Ok ($ns.Mailbox.Deleted.List.Count -eq 2) "moved items landed in the destination (2)"
Ok ($r.out.failed.Count -eq 0) "move_mail reports no failures"

$ns = New-Mailbox; $r = Run-Case 'move_mail' @{ entryIds = @('m1', 'nope'); destFolder = 'Drafts' } $ns
Ok ($r.out.moved -eq 1 -and $r.out.failed.Count -eq 1) "one bad id fails alone instead of failing the batch"
Ok ($r.out.failed[0].entryId -eq 'nope') "failed[] names the id"

$ns = New-Mailbox; $r = Run-Case 'move_mail' @{ destFolder = 'Drafts' } $ns
Ok ((-not $r.ok) -and $r.error -like '*entryId or entryIds is required*') "move_mail without ids is rejected"

# ---------------- delete_mail ----------------
$ns = New-Mailbox; $m1 = $ns.Mailbox.Inbox.List[0]
Ok ($m1.EntryID -eq 'm1') "test fixture: first inbox item is m1"
$r = Run-Case 'delete_mail' @{ entryIds = @('m1', 'm2'); expected = 2 } $ns
Ok ($r.out.deleted -eq 2 -and $r.out.recoverable -eq $true) "delete_mail deletes 2 and says it is recoverable"
Ok ($r.out.deletedEntryIds -contains 'm1') "delete_mail reports what it deleted"
Ok ($m1.Deleted -eq $true) "the item is flagged deleted (soft)"

$ns = New-Mailbox; $r = Run-Case 'delete_mail' @{ entryIds = @('m1', 'm2'); expected = 3 } $ns
Ok ((-not $r.ok) -and $r.error -like '*caller expected 3 items but supplied 2*') "expected-count mismatch refuses to delete"
Ok ($ns.Mailbox.Inbox.List.Count -eq 5) "nothing was deleted when the gate refused"

$ns = New-Mailbox; $r = Run-Case 'delete_mail' @{} $ns
Ok ((-not $r.ok) -and $r.error -like '*entryId or entryIds is required*') "delete_mail without ids is rejected"

# ---------------- mark_read ----------------
$ns = New-Mailbox; $r = Run-Case 'mark_read' @{ entryIds = @('m1'); read = $true } $ns
Ok ($r.out.updated -eq 1 -and $r.out.read -eq $true) "mark_read updated 1 as read"
Ok ($ns.Mailbox.Inbox.List[0].UnRead -eq $false) "UnRead is false after marking read"
Ok ($ns.Mailbox.Inbox.List[0].Saved -eq $true) "the item was saved"

$ns = New-Mailbox; $r = Run-Case 'mark_read' @{ entryIds = @('m2'); read = $false } $ns
Ok ($ns.Mailbox.Inbox.List[1].UnRead -eq $true) "read=false marks it unread"

# ---------------- list_folders ----------------
$ns = New-Mailbox; $r = Run-Case 'list_folders' @{} $ns
Ok ($r.ok) "list_folders runs"
Ok ($r.out.folders.Count -eq 6) "list_folders walks the whole tree breadth-first (6 nodes)"
$byPath = @{}; foreach ($f in $r.out.folders) { $byPath[$f.path] = $f }
Ok ($byPath.ContainsKey('\Inbox')) "folder paths use backslashes"
Ok ($byPath['\Inbox'].depth -eq 1) "Inbox sits at depth 1 under the store root"
Ok ($byPath['\Inbox\Projects'].depth -eq 2) "nested folder sits at depth 2"
Ok ($byPath['\Inbox'].itemCount -eq 5) "itemCount is reported"
Ok ($byPath['\Inbox'].unread -eq 2) "unread is reported"
Ok ($byPath['\Inbox'].childCount -eq 1) "childCount is reported"

$ns = New-Mailbox; $r = Run-Case 'list_folders' @{ counts = $false } $ns
Ok ($r.out.folders[1].itemCount -eq 0) "counts=false skips the per-folder item count"

$ns = New-Mailbox; $r = Run-Case 'list_folders' @{ maxDepth = 1 } $ns
Ok ($r.out.folders.Count -eq 5) "maxDepth=1 stops before the nested folder (5 nodes)"

"---"
"failures: $script:fail"
if ($script:fail -gt 0) { exit 1 }
