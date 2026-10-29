# dsh-arya-maildesk Outlook COM bridge dispatcher.
# Invoked as: powershell -NoProfile -Command "Invoke-Expression (Get-Content -Raw <this file>)"
# Input:  $env:OLK_ACTION = action name; arguments as ONE JSON OBJECT on STDIN
#         (stdin transport — Windows env vars cap near 32 KB and a large mail
#          body silently overflowed that cap, leaving stale values behind)
# Output: single JSON object on stdout.
#
# Quirks baked in (validated against classic Outlook 16.0.x / Exchange):
#  - execution policy blocks .ps1 files -> caller uses Invoke-Expression
#  - Restrict filters need "yyyy-MM-dd HH:mm" date STRINGS, not OADate
#  - with IncludeRecurrences, Items.Count reports MaxInt -> bound loops yourself
#  - recurring-meeting exceptions can have null Start/End -> skip them
#  - PS 5.1 forbids try/catch inline inside hash literals

$ErrorActionPreference = 'Stop'
$out = [ordered]@{}

function Out-Json($obj) {
  [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
  $obj | ConvertTo-Json -Depth 6
}

function Resolve-Folder($ns, $spec) {
  # Resolve a folder to a Folder object. Accepted forms, in order:
  #   - a folder entryId from list_folders (>= 60 chars)
  #   - a well-known English name: Inbox, Sent Items / Sent, Drafts, Deleted Items
  #   - a path from the primary store root: "Inbox\ProjectA", "Projects/2026"
  # An empty spec means Inbox. Segments use / or \; the first segment may
  # name a well-known folder so "Inbox/ProjectA" resolves the same way.
  #
  # Localized (for example CJK) folder names are NOT matched here on purpose:
  # this source must stay ASCII so that no non-UTF-8 transport can corrupt a
  # literal into one that silently never matches. A localized profile should
  # pass a path or a folder entryId from list_folders instead.
  $s = [string]$spec
  if ($s -eq '') { return $ns.GetDefaultFolder(6) }
  # Neither length nor separators can decide id-vs-path on their own: a deep
  # folder path easily exceeds 60 chars, and a real entryId may contain '/'
  # (base64). Try the id first and fall through to the path walk instead of
  # giving up, so a long path resolves and a stale id still reports clearly.
  $idTried = $false
  if ($s.Length -ge 60) {
    $idTried = $true
    try { return $ns.GetFolderFromID($s) } catch { }
  }
  $s = $s -replace '/', '\'
  $segs = @($s -split '\\' | Where-Object { $_ -ne '' })
  $f = $ns.Folders.Item(1)
  if ($segs.Count -gt 0) {
    $first = $segs[0]
    if ($first -ieq 'Inbox') { $f = $ns.GetDefaultFolder(6); $segs = @($segs | Select-Object -Skip 1) }
    elseif ($first -ieq 'Sent Items' -or $first -ieq 'Sent') { $f = $ns.GetDefaultFolder(5); $segs = @($segs | Select-Object -Skip 1) }
    elseif ($first -ieq 'Drafts') { $f = $ns.GetDefaultFolder(16); $segs = @($segs | Select-Object -Skip 1) }
    elseif ($first -ieq 'Deleted Items' -or $first -ieq 'Deleted') { $f = $ns.GetDefaultFolder(3); $segs = @($segs | Select-Object -Skip 1) }
  }
  foreach ($seg in $segs) {
    try { $f = $f.Folders.Item($seg) } catch {
      if ($idTried) { throw "no folder matches '${spec}': it is not a live folder entryId, and '${seg}' is missing from path '${s}'. Re-run list_folders." }
      throw "folder not found: '${seg}' (in path '${s}')"
    }
  }
  return $f
}

try {
  $args_ = @{}
  # Read stdin as UTF-8 explicitly: the Node caller writes UTF-8 bytes, but
  # [Console]::In on PS 5.1 decodes with the OEM/ANSI codepage (GBK here),
  # turning every non-ASCII argument into mojibake.
  $stdinReader = New-Object System.IO.StreamReader([System.Console]::OpenStandardInput(), (New-Object System.Text.UTF8Encoding($false)))
  $raw = $stdinReader.ReadToEnd()
  if ($raw) { $args_ = $raw | ConvertFrom-Json }
  $action = $env:OLK_ACTION

  try {
    $outlook = [Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application')
  } catch {
    # Pre-check before New-Object: on machines with only the UWP "New
    # Outlook" (or nothing installed), ComObject creation fails with an
    # opaque ActiveX error — give the user an actionable message instead.
    if (-not (Test-Path 'Registry::HKEY_CLASSES_ROOT\Outlook.Application')) {
      # Name the case that is now the common one. The new Outlook for Windows
      # (an MSIX app, olk.exe) installs mail, calendar and even Microsoft's own
      # Outlook MCP Server, but registers no COM classes, so this bridge can
      # never drive it — telling someone with Outlook already installed to
      # "install Outlook" only sends them in a circle.
      $newOutlook = $false
      if ($env:LOCALAPPDATA) {
        $newOutlook = Test-Path (Join-Path $env:LOCALAPPDATA 'Microsoft\Olk')
      }
      if ($newOutlook) {
        throw 'classic desktop Outlook not found, but the new Outlook for Windows is installed. The new Outlook has no COM automation surface, so this bridge cannot drive it. Use a machine with classic (Win32) Outlook, or add a non-COM mail backend.'
      }
      throw 'classic desktop Outlook not found (HKCR\Outlook.Application missing). Install the classic Win32 Outlook app and sign in; the UWP "New Outlook" does not support COM automation.'
    }
    $outlook = New-Object -ComObject Outlook.Application
  }
  $ns = $outlook.GetNamespace('MAPI')

  switch ($action) {

    'account' {
      # Prefer the account whose delivery store owns the default Inbox —
      # Accounts order is not guaranteed to put the primary account first.
      $a = $null
      try {
        $defStore = $ns.GetDefaultFolder(6).Store
        foreach ($acc in $ns.Accounts) {
          if ($acc.DeliveryStore -and $acc.DeliveryStore.EntryID -eq $defStore.EntryID) { $a = $acc; break }
        }
      } catch {}
      if (-not $a) { $a = $ns.Accounts | Select-Object -First 1 }
      $out.account = $a.DisplayName
      $out.smtp = $a.SmtpAddress
    }

    'search_mail' {
      $days = if ($args_.days) { [int]$args_.days } else { 14 }
      if ($days -lt 1) { $days = 1 }
      $limit = if ($args_.limit) { [int]$args_.limit } else { 10 }
      if ($limit -lt 1) { $limit = 1 }
      if ($limit -gt 50) { $limit = 50 }
      # Folder resolution: well-known defaults by name, otherwise a path
      # ("Inbox/ProjectA", "Sent Items/2026") walked segment by segment from
      # the primary store root. A missing segment throws with the segment
      # name, so a typo fails loudly instead of searching the wrong folder.
      # TODO(fork): list_mail/move_mail call Resolve-Folder; fold this inline copy
      # into it once the new actions are verified against a real Outlook.
      $folderSpec = if ($args_.folder) { ([string]$args_.folder) -replace '/', '\' } else { 'Inbox' }
      if ($folderSpec -ieq 'Inbox') {
        $folder = $ns.GetDefaultFolder(6)
      } elseif ($folderSpec -ieq 'Sent Items' -or $folderSpec -ieq 'Sent') {
        $folder = $ns.GetDefaultFolder(5)
      } elseif ($folderSpec -ieq 'Drafts') {
        $folder = $ns.GetDefaultFolder(16)
      } else {
        $f = $ns.Folders.Item(1)
        $segs = @($folderSpec -split '\\' | Where-Object { $_ -ne '' })
        if ($segs.Count -gt 0) {
          # First segment may name a well-known default folder; start the
          # walk there so "Inbox/ProjectA" works on localized profiles too.
          if ($segs[0] -ieq 'Inbox') { $f = $ns.GetDefaultFolder(6); $segs = @($segs | Select-Object -Skip 1) }
          elseif ($segs[0] -ieq 'Sent Items' -or $segs[0] -ieq 'Sent') { $f = $ns.GetDefaultFolder(5); $segs = @($segs | Select-Object -Skip 1) }
          elseif ($segs[0] -ieq 'Drafts') { $f = $ns.GetDefaultFolder(16); $segs = @($segs | Select-Object -Skip 1) }
        }
        foreach ($seg in $segs) {
          try { $f = $f.Folders.Item($seg) } catch { throw "folder not found: '$seg' (in path '$folderSpec')" }
        }
        $folder = $f
      }
      # Two-stage Restrict: Jet syntax filters ReceivedTime reliably, but this
      # Outlook/locale build rejects Jet "AND ... LIKE '%x%'" with "Condition
      # is not valid". DASL LIKE works — and a Restrict result can be
      # Restrict-ed again — so text filters run as a second DASL pass.
      # (DASL datereceived string literals silently match everything, so the
      # time bound must stay in the Jet stage.)
      $fmtDate = 'yyyy-MM-dd'
      $fmtFull = 'yyyy-MM-dd HH:mm'
      if ($args_.dateFrom -or $args_.dateTo) {
        # Exact date range; a date-only dateTo extends to end of day.
        $fromTxt = ''; $toTxt = ''
        if ($args_.dateFrom) {
          # Try formats one by one: the array-format TryParseExact overload
          # misbehaves under this PowerShell build (always $false).
          $d = [datetime]::MinValue
          if (-not ([datetime]::TryParseExact([string]$args_.dateFrom, $fmtFull, $null, [System.Globalization.DateTimeStyles]::None, [ref]$d) -or [datetime]::TryParseExact([string]$args_.dateFrom, $fmtDate, $null, [System.Globalization.DateTimeStyles]::None, [ref]$d))) { throw "dateFrom must be yyyy-MM-dd or yyyy-MM-dd HH:mm" }
          $fromTxt = $d.ToString($fmtFull)
        } else { $fromTxt = '2000-01-01 00:00' }
        if ($args_.dateTo) {
          $d = [datetime]::MinValue
          if (-not ([datetime]::TryParseExact([string]$args_.dateTo, $fmtFull, $null, [System.Globalization.DateTimeStyles]::None, [ref]$d) -or [datetime]::TryParseExact([string]$args_.dateTo, $fmtDate, $null, [System.Globalization.DateTimeStyles]::None, [ref]$d))) { throw "dateTo must be yyyy-MM-dd or yyyy-MM-dd HH:mm" }
          if (([string]$args_.dateTo).Length -eq 10) { $d = $d.Date.AddHours(23).AddMinutes(59) }
          $toTxt = $d.ToString($fmtFull)
        } else { $toTxt = (Get-Date).ToString($fmtFull) }
        $items = $folder.Items.Restrict("[ReceivedTime] >= '$fromTxt' AND [ReceivedTime] <= '$toTxt'")
      } else {
        $cutoff = (Get-Date).AddDays(-$days).ToString('yyyy-MM-dd HH:mm')
        $items = $folder.Items.Restrict("[ReceivedTime] >= '$cutoff'")
      }
      if ($args_.from) {
        $v = ([string]$args_.from).Replace([string][char]34, '').Replace('%', '').Replace('_', '')
        $items = $items.Restrict("@SQL=""urn:schemas:httpmail:fromname"" LIKE '%$v%'")
      }
      if ($args_.to) {
        # Recipient-side match on the display-to field (all To recipients as
        # one string), same DASL LIKE pass as fromname.
        $v = ([string]$args_.to).Replace([string][char]34, '').Replace('%', '').Replace('_', '')
        $items = $items.Restrict("@SQL=""urn:schemas:httpmail:displayto"" LIKE '%$v%'")
      }
      if ($args_.subjectContains) {
        $v = ([string]$args_.subjectContains).Replace([string][char]34, '').Replace('%', '').Replace('_', '')
        $items = $items.Restrict("@SQL=""urn:schemas:httpmail:subject"" LIKE '%$v%'")
      }
      $items.Sort('ReceivedTime', $true)
      $mails = @()
      $i = 1
      while ($mails.Count -lt $limit -and $i -le $items.Count) {
        $it = $items.Item($i)
        if ($it.Class -eq 43 -and $null -ne $it.ReceivedTime) {
          $smtp = $null
          try { $smtp = [string]$it.SenderEmailAddress } catch {}
          $mails += [ordered]@{
            entryId   = $it.EntryID
            subject   = [string]$it.Subject
            sender    = [string]$it.SenderName
            senderSmtp = $smtp
            received  = $it.ReceivedTime.ToString('yyyy-MM-dd HH:mm')
            unread    = [bool]$it.UnRead
          }
        }
        $i++
      }
      $out.totalInUnread = $folder.UnReadItemCount
      $out.mails = $mails
    }

    'read_mail' {
      $it = $ns.GetItemFromID([string]$args_.entryId)
      $out.subject = [string]$it.Subject
      $out.sender = [string]$it.SenderName
      # SenderEmailAddress is an EX:/O=... directory name on a corporate
      # Exchange mailbox, so the caller must test it for '@' before offering it
      # as a reply recipient; the display name is always usable for display.
      $senderSmtp = $null
      try { $senderSmtp = [string]$it.SenderEmailAddress } catch {}
      $out.senderSmtp = $senderSmtp
      $out.received = if ($it.ReceivedTime) { $it.ReceivedTime.ToString('yyyy-MM-dd HH:mm') } else { $null }
      $out.to = [string]$it.To
      $out.cc = [string]$it.CC
      $out.bcc = [string]$it.BCC
      $body = [string]$it.Body
      $maxChars = if ($args_.maxChars) { [int]$args_.maxChars } else { 8000 }
      if ($body.Length -gt $maxChars) {
        $out.body = $body.Substring(0, $maxChars) + "`n...[truncated, total $($body.Length) chars]"
        $out.truncated = $true
      } else {
        $out.body = $body
      }
      # HTMLBody is opt-in: it is much larger than Body and most callers only
      # want the text. The panel asks for it so it can render the formatting.
      # It is returned raw - sanitizing belongs in the renderer, which has the
      # document parser, not here.
      if ([string]$args_.format -eq 'html') {
        $html = [string]$it.HTMLBody
        $htmlMax = if ($args_.maxHtmlChars) { [int]$args_.maxHtmlChars } else { 200000 }
        if ($html.Length -gt $htmlMax) {
          $out.htmlBody = $html.Substring(0, $htmlMax)
          $out.htmlTruncated = $true
        } else {
          $out.htmlBody = $html
          $out.htmlTruncated = $false
        }
      }
      $out.attachments = @($it.Attachments | ForEach-Object { $_.DisplayName })
    }

    'draft_mail' {
      $draft = $outlook.CreateItem(0)
      if ($args_.to) { $draft.To = [string]$args_.to }
      if ($args_.cc) { $draft.CC = [string]$args_.cc }
      if ($args_.bcc) { $draft.BCC = [string]$args_.bcc }
      $draft.Subject = [string]$args_.subject
      $draft.Body = [string]$args_.body
      if ($args_.attachments) {
        foreach ($p_ in @($args_.attachments)) {
          if (-not (Test-Path -LiteralPath $p_)) { throw "attachment not found: $p_" }
          $null = $draft.Attachments.Add($p_)
        }
        $out.attachments = @($draft.Attachments | ForEach-Object { $_.DisplayName })
      }
      $draft.Save()
      $out.draftEntryId = $draft.EntryID
      $out.message = 'Draft saved to Drafts folder.'
    }

    'send_draft' {
      # HUMAN-APPROVAL GATE: sending anything requires a one-time OLK_CONFIRM
      # token that index.js mints ONLY after the in-session user popup
      # accepted. Fail closed with a loud error.
      if (-not $env:OLK_CONFIRM -or $env:OLK_CONFIRM.Length -lt 16) {
        throw 'SEND BLOCKED: no human confirmation token (OLK_CONFIRM) — sending requires an approved in-session confirmation first.'
      }
      $draft = $ns.GetItemFromID([string]$args_.entryId)
      # Capture before Send(): Outlook clears these fields once the item
      # moves to Sent Items, so reading after Send() returned empty strings.
      $out.subject = [string]$draft.Subject
      $out.to = [string]$draft.To
      $out.cc = [string]$draft.CC
      $out.bcc = [string]$draft.BCC
      try {
        $draft.Send()
      } catch {
        # Send failed: remove the orphaned draft instead of leaving it to
        # pile up in the Drafts folder, then surface the original failure.
        try { $draft.Delete() } catch { /* best effort */ }
        throw
      }
      $out.sent = $true
    }

    'compose_draft' {
      # A draft composed HERE keeps its HTML. draft_mail assigns $draft.Body
      # (plain text), which flattens a formatted reply, so this action sets
      # HTMLBody instead. When inReplyToEntryId is given, Outlook itself opens
      # the reply/forward so the item joins the right conversation.
      $item = $null
      $mode = [string]$args_.mode
      if (-not $mode) { $mode = 'new' }
      if ($args_.inReplyToEntryId) {
        $base = $ns.GetItemFromID([string]$args_.inReplyToEntryId)
        if ($mode -eq 'replyAll') { $item = $base.ReplyAll() }
        elseif ($mode -eq 'forward') { $item = $base.Forward() }
        else { $item = $base.Reply() }
      } else {
        $item = $outlook.CreateItem(0)
      }
      if ($args_.to) { $item.To = [string]$args_.to }
      if ($args_.cc) { $item.CC = [string]$args_.cc }
      if ($args_.bcc) { $item.BCC = [string]$args_.bcc }
      if ($args_.subject) { $item.Subject = [string]$args_.subject }
      if ($null -ne $args_.htmlBody -and [string]$args_.htmlBody -ne '') {
        $item.HTMLBody = [string]$args_.htmlBody
      }
      if ($args_.attachments) {
        foreach ($p_ in @($args_.attachments)) {
          if (-not (Test-Path -LiteralPath $p_)) { throw "attachment not found: $p_" }
          $null = $item.Attachments.Add($p_)
        }
      }
      $item.Save()
      $out.draftEntryId = $item.EntryID
      $out.mode = $mode
      $out.message = 'Draft saved to the Drafts folder.'
    }

    'compose_send' {
      # Same HUMAN-APPROVAL GATE as send_draft: index.js mints the token only
      # after a human allowed this exact content. Fail closed.
      if (-not $env:OLK_CONFIRM -or $env:OLK_CONFIRM.Length -lt 16) {
        throw 'SEND BLOCKED: no human confirmation token (OLK_CONFIRM) — sending requires an approved in-session confirmation first.'
      }
      $item = $null
      $mode = [string]$args_.mode
      if (-not $mode) { $mode = 'new' }
      if ($args_.inReplyToEntryId) {
        $base = $ns.GetItemFromID([string]$args_.inReplyToEntryId)
        if ($mode -eq 'replyAll') { $item = $base.ReplyAll() }
        elseif ($mode -eq 'forward') { $item = $base.Forward() }
        else { $item = $base.Reply() }
      } else {
        $item = $outlook.CreateItem(0)
      }
      if ($args_.to) { $item.To = [string]$args_.to }
      if ($args_.cc) { $item.CC = [string]$args_.cc }
      if ($args_.bcc) { $item.BCC = [string]$args_.bcc }
      if ($args_.subject) { $item.Subject = [string]$args_.subject }
      if ($null -ne $args_.htmlBody -and [string]$args_.htmlBody -ne '') {
        $item.HTMLBody = [string]$args_.htmlBody
      }
      if ($args_.attachments) {
        foreach ($p_ in @($args_.attachments)) {
          if (-not (Test-Path -LiteralPath $p_)) { throw "attachment not found: $p_" }
          $null = $item.Attachments.Add($p_)
        }
      }
      # Capture before Send(): Outlook clears these once the item is in Sent Items.
      $out.subject = [string]$item.Subject
      $out.to = [string]$item.To
      $out.cc = [string]$item.CC
      $out.bcc = [string]$item.BCC
      try {
        $item.Send()
      } catch {
        try { $item.Delete() } catch { /* best effort */ }
        throw
      }
      $out.sent = $true
      $out.mode = $mode
    }

    'calendar_query' {
      $days = if ($args_.days) { [int]$args_.days } else { 7 }
      if ($days -lt 1) { $days = 1 }
      $cal = $ns.GetDefaultFolder(9)
      $ci = $cal.Items
      $ci.IncludeRecurrences = $true
      $ci.Sort('[Start]')
      $s = (Get-Date).ToString('yyyy-MM-dd HH:mm')
      $e = (Get-Date).AddDays($days).ToString('yyyy-MM-dd HH:mm')
      $r = $ci.Restrict("[Start] >= '$s' AND [Start] <= '$e'")
      $events = @()
      $i = 1
      while ($events.Count -lt 20) {
        $ev = $null
        try { $ev = $r.Item($i) } catch { break }  # empty/boundary collection throws instead of returning null
        if ($null -eq $ev) { break }
        if ($null -ne $ev.Start) {
          $events += [ordered]@{
            entryId   = $ev.EntryID
            subject   = [string]$ev.Subject
            start     = $ev.Start.ToString('yyyy-MM-dd HH:mm')
            end       = if ($ev.End) { $ev.End.ToString('yyyy-MM-dd HH:mm') } else { $null }
            organizer = [string]$ev.Organizer
            location  = [string]$ev.Location
            busy      = [string]$ev.BusyStatus
            isMeeting = ($ev.Recipients.Count -gt 0)
          }
        }
        $i++
      }
      $out.events = $events
    }

    'calendar_create' {
      # Pre-validate the exact format COM expects, so a bad date fails with a
      # clear message instead of COM's opaque "object does not support this method".
      $fmt = 'yyyy-MM-dd HH:mm'
      $startOk = $false; $endOk = $false
      $parsedStart = [datetime]::MinValue; $parsedEnd = [datetime]::MinValue
      if ($args_.start) { $startOk = [datetime]::TryParseExact([string]$args_.start, $fmt, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::None, [ref]$parsedStart) }
      if ($args_.end)   { $endOk   = [datetime]::TryParseExact([string]$args_.end,   $fmt, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::None, [ref]$parsedEnd) }
      if (-not $startOk -or -not $endOk) {
        throw "start/end must use the exact format 'yyyy-MM-dd HH:mm' (24h), e.g. '2026-09-08 15:30'."
      }
      if ($parsedEnd -le $parsedStart) {
        throw 'end must be later than start.'
      }
      $appt = $outlook.CreateItem(1)  # olAppointmentItem
      $appt.Subject = [string]$args_.subject
      $appt.Start = $parsedStart
      $appt.End = $parsedEnd
      if ($args_.location) { $appt.Location = [string]$args_.location }
      if ($args_.body) { $appt.Body = [string]$args_.body }
      # Recurrence: { type: daily|weekly|monthly, interval, daysOfWeek (weekly),
      # count or until ("yyyy-MM-dd") }. Applied before Save/Send.
      $recSummary = ''
      if ($args_.recurrence) {
        $rec = $args_.recurrence
        $type = [string]$rec.type
        $pat = $appt.GetRecurrencePattern()
        switch ($type) {
          'daily'    { $pat.RecurrenceType = 0 }
          'weekly'   { $pat.RecurrenceType = 1 }
          'monthly'  { $pat.RecurrenceType = 2 }
          default    { throw "recurrence.type must be daily, weekly or monthly (got '$type')." }
        }
        $pat.Interval = if ($rec.interval) { [int]$rec.interval } else { 1 }
        if ($type -eq 'weekly') {
          if (-not $rec.daysOfWeek -or @($rec.daysOfWeek).Count -eq 0) { throw 'weekly recurrence requires daysOfWeek, e.g. ["Mon","Wed"].' }
          $mask = 0
          foreach ($d_ in @($rec.daysOfWeek)) {
            switch ([string]$d_) {
              'Sun' { $mask += 1 } 'Mon' { $mask += 2 } 'Tue' { $mask += 4 }
              'Wed' { $mask += 8 } 'Thu' { $mask += 16 } 'Fri' { $mask += 32 } 'Sat' { $mask += 64 }
              default { throw "daysOfWeek entry must be Sun/Mon/Tue/Wed/Thu/Fri/Sat (got '$d_')." }
            }
          }
          $pat.DayOfWeekMask = $mask
        }
        if ($rec.count) { $pat.Occurrences = [int]$rec.count }
        elseif ($rec.until) {
          $ud = [datetime]::MinValue
          if (-not ([datetime]::TryParseExact([string]$rec.until, 'yyyy-MM-dd', $null, [System.Globalization.DateTimeStyles]::None, [ref]$ud))) { throw 'recurrence.until must be yyyy-MM-dd.' }
          $pat.PatternEndDate = $ud.Date
        } else { $pat.Occurrences = 10 }  # sensible default, keeps patterns finite
        # PS 5.1 has no if-expressions; build the summary with statements.
        $recSummary = $type + ' x' + $pat.Interval
        if ($type -eq 'weekly') { $recSummary += ' on ' + (@($rec.daysOfWeek) -join ',') }
        if ($rec.count) { $recSummary += ', ' + $rec.count + 'x' }
        elseif ($rec.until) { $recSummary += ', until ' + $rec.until }
        else { $recSummary += ', 10x' }
        $out.recurrence = $recSummary
      }
      # Capture before Send()/Save(): fields may be cleared once the item
      # leaves the drafts/calendar window.
      $sentSubject = [string]$appt.Subject
      $hasAttendees = ($args_.recipients -and @($args_.recipients).Count -gt 0) -or ($args_.optionalRecipients -and @($args_.optionalRecipients).Count -gt 0)
      if ($hasAttendees) {
        # HUMAN-APPROVAL GATE: dispatching a meeting invite leaves the
        # machine — same OLK_CONFIRM token rule as send_draft.
        if (-not $env:OLK_CONFIRM -or $env:OLK_CONFIRM.Length -lt 16) {
          throw 'SEND BLOCKED: no human confirmation token (OLK_CONFIRM) — dispatching a meeting invite requires an approved in-session confirmation first.'
        }
        foreach ($r_ in @($args_.recipients)) { $null = $appt.Recipients.Add([string]$r_); $appt.Recipients.Item($appt.Recipients.Count).Type = 1 }        # olRequired
        foreach ($r_ in @($args_.optionalRecipients)) { $null = $appt.Recipients.Add([string]$r_); $appt.Recipients.Item($appt.Recipients.Count).Type = 2 }  # olOptional
        if (-not $appt.Recipients.ResolveAll()) {
          $unresolved = @($appt.Recipients | Where-Object { -not $_.Resolved } | ForEach-Object { $_.Name })
          throw ('cannot resolve recipient(s): ' + ($unresolved -join '; ') + ' — invite NOT sent.')
        }
        $appt.MeetingStatus = 1  # olMeeting
        $appt.Send()  # dispatch the invite (caller already obtained user confirmation)
        $out.invited = $true
        $out.message = 'Meeting created and invite dispatched.'
      } else {
        $appt.Save()
        $out.message = 'Personal calendar entry created (no attendees).'
      }
      $out.created = $true
      $out.subject = $sentSubject
      $out.start = $parsedStart.ToString($fmt)
    }

    'meeting_requests' {
      # List pending meeting-request items (Class 53) in the Inbox: invites
      # others sent you that are still awaiting your response.
      $folder = $ns.GetDefaultFolder(6)
      $items = $folder.Items.Restrict("[ReceivedTime] >= '" + (Get-Date).AddDays(-30).ToString('yyyy-MM-dd HH:mm') + "'")
      $items.Sort('ReceivedTime', $true)
      $reqs = @()
      $i = 1
      while ($reqs.Count -lt 20 -and $i -le $items.Count) {
        $it = $null
        try { $it = $items.Item($i) } catch { break }
        if ($null -eq $it) { break }
        if ($it.Class -eq 53) {
          $assoc = $null
          try { $assoc = $it.GetAssociatedAppointment($false) } catch {}
          $reqs += [ordered]@{
            entryId   = $it.EntryID
            subject   = [string]$it.Subject
            organizer = [string]$it.SenderName
            received  = $it.ReceivedTime.ToString('yyyy-MM-dd HH:mm')
            when      = if ($assoc -and $assoc.Start) { $assoc.Start.ToString('yyyy-MM-dd HH:mm') + ' ~ ' + $assoc.End.ToString('yyyy-MM-dd HH:mm') } else { $null }
            location  = if ($assoc) { [string]$assoc.Location } else { '' }
          }
        }
        $i++
      }
      $out.meetingRequests = $reqs
      $out.count = $reqs.Count
    }

    'respond_meeting' {
      # HUMAN-APPROVAL GATE: responding sends a reply to the organizer.
      if (-not $env:OLK_CONFIRM -or $env:OLK_CONFIRM.Length -lt 16) {
        throw 'SEND BLOCKED: no human confirmation token (OLK_CONFIRM) — responding to a meeting request requires an approved in-session confirmation first.'
      }
      $resp = [string]$args_.response
      $code = 0
      switch ($resp) {
        'accept'     { $code = 3 }  # olMeetingAccepted
        'tentative'  { $code = 2 }  # olMeetingTentative
        'decline'    { $code = 4 }  # olMeetingDeclined
        default { throw "response must be accept, tentative or decline (got '$resp')." }
      }
      $it = $ns.GetItemFromID([string]$args_.entryId)
      if ($it.Class -ne 53) { throw 'item is not a meeting request (Class ' + $it.Class + ').' }
      $null = $it.Respond($code, $true, $true)  # (response, NoUI, sendResponse)
      $out.responded = $resp
      $out.subject = [string]$it.Subject
      $out.message = 'Meeting ' + $resp + 'ed; response sent to organizer.'
    }

    'meeting_update' {
      # HUMAN-APPROVAL GATE: updating an organized meeting re-notifies attendees.
      if (-not $env:OLK_CONFIRM -or $env:OLK_CONFIRM.Length -lt 16) {
        throw 'SEND BLOCKED: no human confirmation token (OLK_CONFIRM) — updating a meeting requires an approved in-session confirmation first.'
      }
      $appt = $ns.GetItemFromID([string]$args_.entryId)
      if ($appt.Class -ne 26) { throw 'item is not an appointment (Class ' + $appt.Class + ').' }
      # Organizer check: compare against the CURRENT user, not just "has an
      # Organizer value" — an attendee's calendar copy also carries Organizer,
      # and calling Send() on it would throw mid-operation (H2).
      $me = ''
      try { $me = [string]$ns.CurrentUser.Name } catch {}
      $isOrganizer = ($me -ne '' -and $appt.Organizer -and (([string]$appt.Organizer).Trim() -ieq $me.Trim()) -and ($appt.Recipients.Count -gt 0))
      $fmt = 'yyyy-MM-dd HH:mm'
      if ($args_.start) {
        $d = [datetime]::MinValue
        if (-not ([datetime]::TryParseExact([string]$args_.start, $fmt, $null, [System.Globalization.DateTimeStyles]::None, [ref]$d))) { throw 'start must be yyyy-MM-dd HH:mm' }
        $appt.Start = $d
      }
      if ($args_.end) {
        $d = [datetime]::MinValue
        if (-not ([datetime]::TryParseExact([string]$args_.end, $fmt, $null, [System.Globalization.DateTimeStyles]::None, [ref]$d))) { throw 'end must be yyyy-MM-dd HH:mm' }
        $appt.End = $d
      }
      if ($args_.subject) { $appt.Subject = [string]$args_.subject }
      if ($args_.location) { $appt.Location = [string]$args_.location }
      if ($args_.body) { $appt.Body = [string]$args_.body }
      $out.subject = [string]$appt.Subject
      $out.start = $appt.Start.ToString($fmt)
      $out.end = $appt.End.ToString($fmt)
      $out.updated = $true
      if ($isOrganizer) {
        $appt.Send()  # re-notify attendees with the update
        $out.notified = $true
        $out.message = 'Meeting updated; attendees notified.'
      } else {
        $appt.Save()
        $out.message = 'Appointment updated (no attendees to notify).'
      }
    }

    'meeting_cancel' {
      # HUMAN-APPROVAL GATE: cancelling notifies every attendee.
      if (-not $env:OLK_CONFIRM -or $env:OLK_CONFIRM.Length -lt 16) {
        throw 'SEND BLOCKED: no human confirmation token (OLK_CONFIRM) — cancelling a meeting requires an approved in-session confirmation first.'
      }
      $appt = $ns.GetItemFromID([string]$args_.entryId)
      if ($appt.Class -ne 26) { throw 'item is not an appointment (Class ' + $appt.Class + ').' }
      $out.subject = [string]$appt.Subject
      $out.start = $appt.Start.ToString('yyyy-MM-dd HH:mm')
      # Same organizer check as meeting_update: only the ORGANIZER may cancel
      # with notifications; an attendee cancelling just deletes their own copy
      # (H2 — Send() on an attendee copy throws mid-operation).
      $me = ''
      try { $me = [string]$ns.CurrentUser.Name } catch {}
      $isOrganizer = ($me -ne '' -and $appt.Organizer -and (([string]$appt.Organizer).Trim() -ieq $me.Trim()) -and ($appt.Recipients.Count -gt 0))
      if ($isOrganizer) {
        $appt.MeetingStatus = 5  # olMeetingCanceled
        $appt.Send()             # dispatch cancellation to attendees
        $out.notified = $true
        $out.message = 'Meeting cancelled; attendees notified.'
      } else {
        $appt.Delete()           # personal entry, or attendee removing their own copy
        $out.message = 'Appointment deleted (no notifications sent).'
      }
      $out.cancelled = $true
    }

    'search_people' {
      # GAL search: display name always; SMTP address too when the query is
      # plain ASCII (a PropertyAccessor call per entry is too slow for 20k+
      # entries, so the expensive path runs only for ascii queries and stops
      # at the first hit that matches).
      $q = [string]$args_.query
      $limit = if ($args_.limit) { [int]$args_.limit } else { 20 }
      if ($limit -lt 1) { $limit = 1 }
      if ($limit -gt 100) { $limit = 100 }
      $ascii = $q -match '^[A-Za-z0-9.@_\-]+$'
      $rx = [regex]::Escape($q)
      $gal = $null
      foreach ($al in $ns.AddressLists) { if ($al.Name -match 'Global') { $gal = $al; break } }
      if (-not $gal) { $gal = @($ns.AddressLists)[0] }
      $people = @()
      foreach ($e in $gal.AddressEntries) {
        $hit = $false; $smtp = ''
        if ($e.Name -match $rx) { $hit = $true }
        if (-not $hit -and $ascii) {
          try { $p = $e.PropertyAccessor.GetProperty('http://schemas.microsoft.com/mapi/proptag/0x39FE001E'); if ($p) { $smtp = [string]$p; if ($smtp -match $rx) { $hit = $true } } } catch {}
        }
        if ($hit) {
          if ($smtp -eq '') { try { $eu = $e.GetExchangeUser(); if ($eu) { $smtp = $eu.PrimarySmtpAddress } } catch {} }
          $people += [ordered]@{ name = [string]$e.Name; smtp = $smtp }
          if ($people.Count -ge $limit) { break }
        }
      }
      $out.people = $people
      $out.count = $people.Count
    }

    'freebusy' {
      # Free/Busy per day starting today, 30-min grid, merged consecutive
      # same-status slots. Status chars: 0 free, 1 tentative, 2 busy, 3 OOF.
      # FreeBusy() returns one month-long string from the passed start; slots
      # after 22:00 are dropped as noise.
      $mail = [string]$args_.email
      $days = if ($args_.days) { [int]$args_.days } else { 3 }
      if ($days -lt 1) { $days = 1 }
      if ($days -gt 14) { $days = 14 }
      $r = $ns.CreateRecipient($mail)
      [void]$r.Resolve()
      if (-not $r.Resolved) { throw "cannot resolve recipient: $mail" }
      $base = (Get-Date).Date
      $daysOut = @()
      for ($d = 0; $d -lt $days; $d++) {
        $day = $base.AddDays($d)
        $dayStart = $day.AddHours(8)
        $s = $r.FreeBusy($dayStart, 30, $true)
        $slots = @(); $i = 0
        while ($i -lt $s.Length) {
          $c = [string]$s[$i]
          if ($c -ne '0') {
            $j = $i
            while ($j -lt $s.Length -and [string]$s[$j] -eq $c) { $j++ }
            $st = $dayStart.AddMinutes(30 * $i)
            $en = $dayStart.AddMinutes(30 * $j)
            # FreeBusy returns one month from the passed start; keep only the
            # queried day, working-relevant hours (08:00-22:00).
            if ($st.Date -eq $day -and $st.Hour -lt 22) {
              $statusName = 'busy'
              if ($c -eq '1') { $statusName = 'tentative' } elseif ($c -eq '3') { $statusName = 'out-of-office' }
              $slots += [ordered]@{
                start  = $st.ToString('yyyy-MM-dd HH:mm')
                end    = $en.ToString('yyyy-MM-dd HH:mm')
                status = $statusName
              }
            }
            $i = $j
          } else { $i++ }
        }
        if ($slots.Count -gt 0) {
          $daysOut += [ordered]@{ date = $day.ToString('yyyy-MM-dd'); busySlots = $slots }
        }
      }
      $out.email = $mail
      $out.days = $daysOut
    }

    'search_rooms' {
      # Room search over the "All Rooms"-style address list, with optional
      # availability check for one date + time window (30-min grid).
      $q = ''
      if ($args_.query) { $q = [string]$args_.query }
      $limit = if ($args_.limit) { [int]$args_.limit } else { 15 }
      if ($limit -lt 1) { $limit = 1 }
      if ($limit -gt 50) { $limit = 50 }
      $dt = ''
      if ($args_.date) { $dt = [string]$args_.date }
      $t1 = if ($args_.start) { [string]$args_.start } else { '09:00' }
      $t2 = if ($args_.end) { [string]$args_.end } else { '18:00' }
      if ($dt -ne '' -and $dt -notmatch '^\d{4}-\d{2}-\d{2}$') { throw "date must be yyyy-MM-dd" }
      $roomList = $null
      foreach ($al in $ns.AddressLists) { if ($al.Name -match 'Room') { $roomList = $al; break } }
      if (-not $roomList) { throw 'no room address list found (looked for an address list named Room)' }
      $rx = $null
      if ($q -ne '') { $rx = [regex]::Escape($q) }
      $rooms = @()
      foreach ($e in $roomList.AddressEntries) {
        if ($rx -and $e.Name -notmatch $rx) { continue }
        $smtp = ''
        try { $p = $e.PropertyAccessor.GetProperty('http://schemas.microsoft.com/mapi/proptag/0x39FE001E'); if ($p) { $smtp = [string]$p } } catch {}
        if ($smtp -eq '') { try { $eu = $e.GetExchangeUser(); if ($eu) { $smtp = $eu.PrimarySmtpAddress } } catch {} }
        $entry = [ordered]@{ name = [string]$e.Name; smtp = $smtp; available = $null; conflicts = @() }
        if ($dt -ne '') {
          $ws = [datetime]::ParseExact(($dt + ' ' + $t1), 'yyyy-MM-dd HH:mm', $null)
          $we = [datetime]::ParseExact(($dt + ' ' + $t2), 'yyyy-MM-dd HH:mm', $null)
          # Resolve by SMTP when available (unique); fall back to the display
          # name, which can ambiguously match a same-named person.
          $resolveKey = $e.Name
          if ($smtp -ne '') { $resolveKey = $smtp }
          $rr = $ns.CreateRecipient($resolveKey)
          [void]$rr.Resolve()
          if ($rr.Resolved) {
            $fs = $rr.FreeBusy($ws, 30, $true)
            $conf = @(); $avail = $true
            $n = [math]::Floor(($we - $ws).TotalMinutes / 30)
            for ($k = 0; $k -lt $n; $k++) {
              $cc = [string]$fs[$k]
              if ($cc -ne '0') {
                $avail = $false
                $cs = $ws.AddMinutes(30 * $k)
                $ce = $cs.AddMinutes(30)
                $conf += ($cs.ToString('HH:mm') + '-' + $ce.ToString('HH:mm'))
              }
            }
            $entry.available = $avail
            $entry.conflicts = $conf
          }
        }
        $rooms += $entry
        if ($rooms.Count -ge $limit) { break }
      }
      $out.rooms = $rooms
      $out.count = $rooms.Count
    }

    'reply_preview' {
      # Read-only preview of the original mail for the two-call confirmation.
      # Returns the headers the popup needs so the user sees exactly which
      # mail they are replying to / forwarding.
      $it = $ns.GetItemFromID([string]$args_.entryId)
      $out.subject = [string]$it.Subject
      $out.sender = [string]$it.SenderName
      $out.to = [string]$it.To
      $out.cc = [string]$it.CC
      $out.bcc = [string]$it.BCC
      $out.received = if ($it.ReceivedTime) { $it.ReceivedTime.ToString('yyyy-MM-dd HH:mm') } else { $null }
      $out.messageClass = [string]$it.MessageClass
    }

    'reply_mail' {
      # HUMAN-APPROVAL GATE: reply / reply-all / forward all leave the
      # machine — same OLK_CONFIRM token rule as send_draft.
      if (-not $env:OLK_CONFIRM -or $env:OLK_CONFIRM.Length -lt 16) {
        throw 'SEND BLOCKED: no human confirmation token (OLK_CONFIRM) — replying or forwarding requires an approved in-session confirmation first.'
      }
      $mode = [string]$args_.mode
      if ($mode -ne 'reply' -and $mode -ne 'reply-all' -and $mode -ne 'forward') {
        throw "mode must be one of: reply, reply-all, forward (got '$mode')."
      }
      $it = $ns.GetItemFromID([string]$args_.entryId)
      # Reply()/ReplyAll()/Forward() return a NEW unsent MailItem with the
      # recipients (reply/reply-all) and attachments (forward) preset.
      switch ($mode) {
        'reply'      { $r = $it.Reply() }
        'reply-all'  { $r = $it.ReplyAll() }
        'forward'    {
          $r = $it.Forward()
          if ($args_.to) { $r.To = [string]$args_.to }  # forward needs an explicit recipient
        }
      }
      if ($mode -eq 'forward' -and -not $r.To) {
        throw 'forward requires a "to" recipient.'
      }
      # Optional cc/bcc on every mode (reply-all presets cc from the
      # original; an explicit cc/bcc overwrites the preset).
      if ($args_.cc) { $r.CC = [string]$args_.cc }
      if ($args_.bcc) { $r.BCC = [string]$args_.bcc }
      # Auto-quote the original below the new text (Outlook classic style).
      $origSent = if ($it.ReceivedTime) { $it.ReceivedTime.ToString('yyyy-MM-dd HH:mm') } else { '' }
      $quote = "-----Original Message-----`r`n" +
        "From: $([string]$it.SenderName)`r`n" +
        "Sent: $origSent`r`n" +
        "To: $([string]$it.To)`r`n" +
        "Subject: $([string]$it.Subject)`r`n`r`n" +
        "$([string]$it.Body)"
      $newBody = [string]$args_.body
      if ($newBody) { $r.Body = $newBody + "`r`n`r`n" + $quote } else { $r.Body = $quote }
      if ($args_.attachments) {
        foreach ($p_ in @($args_.attachments)) {
          if (-not (Test-Path -LiteralPath $p_)) { throw "attachment not found: $p_" }
          $null = $r.Attachments.Add($p_)
        }
      }
      # Capture before Send(): Outlook clears these fields once the item
      # moves to Sent Items.
      $out.subject = [string]$r.Subject
      $out.to = [string]$r.To
      $out.cc = [string]$r.CC
      $out.bcc = [string]$r.BCC
      $out.mode = $mode
      $r.Send()
      $out.sent = $true
      $out.message = "Mail $mode dispatched."
    }

    'save_attachment' {
      # Save one attachment of a mail to a local directory by display name.
      $it = $ns.GetItemFromID([string]$args_.entryId)
      $name = [string]$args_.name
      $dir = [string]$args_.dir
      if ($dir -eq '' ) { throw 'dir is required (target directory for the saved file).' }
      if (-not (Test-Path -LiteralPath $dir -PathType Container)) { throw "target directory not found: $dir" }
      $found = $null
      foreach ($a_ in $it.Attachments) { if ($a_.DisplayName -eq $name) { $found = $a_; break } }
      if (-not $found) { throw "attachment not found on mail: $name" }
      # Sanitize: keep only the file NAME (no path traversal from a crafted
      # display name), and never silently overwrite an existing file.
      $safe = [System.IO.Path]::GetFileName($name)
      if ($safe -eq '' -or $safe -eq '.' -or $safe -eq '..') { throw "attachment has an unusable file name: $name" }
      $target = Join-Path $dir $safe
      if (Test-Path -LiteralPath $target) {
        $base = [System.IO.Path]::GetFileNameWithoutExtension($safe)
        $ext = [System.IO.Path]::GetExtension($safe)
        $n = 1
        while (Test-Path -LiteralPath (Join-Path $dir "$base($n)$ext")) { $n++ }
        $target = Join-Path $dir "$base($n)$ext"
      }
      $found.SaveAsFile($target)
      $out.savedPath = $target
      $out.sizeBytes = $found.Size
    }

    'open_mail' {
      # Open one mail in a desktop Outlook inspector window. Local-only
      # (nothing leaves the machine), so no OLK_CONFIRM gate is needed.
      # Displaying also marks it READ: the user has now seen this mail, so
      # badge / unread counters should stop counting it.
      $it = $ns.GetItemFromID([string]$args_.entryId)
      $out.subject = [string]$it.Subject
      $out.sender = [string]$it.SenderName
      $wasUnread = [bool]$it.UnRead
      $it.UnRead = $false
      $it.Display()
      $out.opened = $true
      $out.wasUnread = $wasUnread
    }


    # Mail management (fork addition). Everything below is ADDITIVE: no
    # existing case is touched. The machine that wrote this has no Outlook
    # installed, so nothing here could be regression-tested against the
    # working read paths - keep it that way until M0 validates on a real
    # profile.
    # ------------------------------------------------------------------

    'list_folders' {
      # Folder tree of every store in the profile, for the GUI folder pane.
      # Counts cost one COM read per folder; pass counts=false on very large
      # mailboxes and fetch them lazily per folder instead.
      $maxDepth = if ($args_.maxDepth) { [int]$args_.maxDepth } else { 6 }
      $withCounts = $true
      if ($null -ne $args_.counts) { $withCounts = [bool]$args_.counts }
      $folders = @()
      $q = New-Object System.Collections.Queue
      foreach ($st in $ns.Folders) {
        $q.Enqueue((New-Object psobject -Property @{ F = $st; D = 0; S = [string]$st.Name }))
      }
      while ($q.Count -gt 0) {
        $node = $q.Dequeue()
        $f = $node.F
        $d = [int]$node.D
        $cnt = 0
        $unc = 0
        if ($withCounts) {
          try { $cnt = [int]$f.Items.Count } catch {}
          try { $unc = [int]$f.UnReadItemCount } catch {}
        }
        $folders += [ordered]@{
          entryId    = [string]$f.EntryID
          name       = [string]$f.Name
          path       = ([string]$f.FolderPath -replace '/', '\')
          store      = [string]$node.S
          depth      = $d
          childCount = [int]$f.Folders.Count
          itemCount  = $cnt
          unread     = $unc
        }
        if ($d -lt $maxDepth) {
          foreach ($c in $f.Folders) {
            $q.Enqueue((New-Object psobject -Property @{ F = $c; D = ($d + 1); S = [string]$node.S }))
          }
        }
      }
      $out.folders = $folders
    }

    'list_mail' {
      # One page of one folder, for a GUI list pane. search_mail answers
      # "find mail matching X"; this answers "show me folder F, page N".
      # Paging is offset-based on a sorted view, so a new arrival shifts
      # later pages - a GUI should re-read page one when it regains focus.
      $folder = Resolve-Folder $ns ([string]$args_.folder)
      $limit = if ($args_.limit) { [int]$args_.limit } else { 50 }
      if ($limit -lt 1) { $limit = 1 }
      if ($limit -gt 200) { $limit = 200 }
      $offset = if ($null -ne $args_.offset) { [int]$args_.offset } else { 0 }
      if ($offset -lt 0) { $offset = 0 }
      $desc = $true
      if ($null -ne $args_.descending) { $desc = [bool]$args_.descending }
      $sortField = 'ReceivedTime'
      $sf = [string]$args_.sort
      if ($sf -ieq 'subject') { $sortField = 'Subject' }
      elseif ($sf -ieq 'sender') { $sortField = 'SenderName' }
      $withPreview = $false
      if ($null -ne $args_.preview) { $withPreview = [bool]$args_.preview }
      $items = $folder.Items
      if ($args_.unreadOnly -eq $true) { $items = $items.Restrict('[UnRead] = True') }
      $items.Sort($sortField, $desc)
      $total = [int]$items.Count
      $mails = @()
      $i = $offset + 1
      $seen = 0
      # Item(i) throws past the boundary instead of returning null, so the
      # loop is bounded by the count read above (upstream quirk, preserved).
      while ($seen -lt $limit -and $i -le $total) {
        $it = $items.Item($i)
        # 43 = olMail, 53 = olMeetingRequest: invites belong in an inbox list.
        if ($it.Class -eq 43 -or $it.Class -eq 53) {
          $smtp = $null
          try { $smtp = [string]$it.SenderEmailAddress } catch {}
          $attach = 0
          try { $attach = [int]$it.Attachments.Count } catch {}
          $preview = ''
          if ($withPreview) {
            try {
              $b = [string]$it.Body
              if ($b.Length -gt 160) { $b = $b.Substring(0, 160) }
              $preview = ($b -replace '\s+', ' ').Trim()
            } catch {}
          }
          $mails += [ordered]@{
            entryId    = $it.EntryID
            subject    = [string]$it.Subject
            sender     = [string]$it.SenderName
            senderSmtp = $smtp
            received   = $it.ReceivedTime.ToString('yyyy-MM-dd HH:mm')
            unread     = [bool]$it.UnRead
            hasAttach  = ($attach -gt 0)
            preview    = $preview
          }
          $seen++
        }
        $i++
      }
      $out.total = $total
      $out.offset = $offset
      $out.returned = $mails.Count
      $out.folderPath = [string]$folder.FolderPath
      # Some stores (public folders, some archive add-ins) do not expose
      # UnReadItemCount; -1 means "the store did not tell us".
      $unc = -1
      try { $unc = [int]$folder.UnReadItemCount } catch { }
      $out.unreadInFolder = $unc
      $out.mails = $mails
    }

    'move_mail' {
      # Bulk move/archive. Move() returns the NEW item and the EntryID
      # changes with it, so the caller gets the new ids back or its
      # selection goes stale.
      $dest = Resolve-Folder $ns ([string]$args_.destFolder)
      $ids = @()
      if ($args_.entryIds) { $ids = @($args_.entryIds) } elseif ($args_.entryId) { $ids = @([string]$args_.entryId) }
      if ($ids.Count -eq 0) { throw 'entryId or entryIds is required.' }
      if ($ids.Count -gt 500) { throw "refusing to move $($ids.Count) items in one call (max 500)." }
      $movedIds = @()
      $failed = @()
      foreach ($id in $ids) {
        try {
          $it = $ns.GetItemFromID([string]$id)
          $newItem = $it.Move($dest)
          $movedIds += [string]$newItem.EntryID
        } catch {
          $failed += [ordered]@{ entryId = [string]$id; error = $_.Exception.Message }
        }
      }
      $out.moved = $movedIds.Count
      $out.destPath = [string]$dest.FolderPath
      $out.newEntryIds = $movedIds
      $out.failed = $failed
    }

    'delete_mail' {
      # SOFT delete only: Item.Delete() moves the item to Deleted Items, so
      # the user can undo it from Outlook. There is deliberately no
      # permanent-delete path: purging Deleted Items defeats the retention
      # policy on a managed mailbox and cannot be undone.
      # expected is a safety gate: the caller states how many items the
      # human actually selected, and a mismatch aborts instead of deleting
      # a wider set than was shown.
      $ids = @()
      if ($args_.entryIds) { $ids = @($args_.entryIds) } elseif ($args_.entryId) { $ids = @([string]$args_.entryId) }
      if ($ids.Count -eq 0) { throw 'entryId or entryIds is required.' }
      if ($ids.Count -gt 500) { throw "refusing to delete $($ids.Count) items in one call (max 500)." }
      if ($null -ne $args_.expected) {
        $expect = [int]$args_.expected
        if ($expect -ne $ids.Count) { throw "refusing to delete: caller expected $expect items but supplied $($ids.Count). Re-read the folder and retry." }
      }
      $deletedIds = @()
      $failed = @()
      foreach ($id in $ids) {
        try {
          $it = $ns.GetItemFromID([string]$id)
          $eid = [string]$it.EntryID
          $it.Delete()
          $deletedIds += $eid
        } catch {
          $failed += [ordered]@{ entryId = [string]$id; error = $_.Exception.Message }
        }
      }
      $out.deleted = $deletedIds.Count
      $out.recoverable = $true
      $out.deletedEntryIds = $deletedIds
      $out.failed = $failed
    }

    'mark_read' {
      # Set or clear the read state on one or many items.
      $read = $true
      if ($null -ne $args_.read) { $read = [bool]$args_.read }
      $ids = @()
      if ($args_.entryIds) { $ids = @($args_.entryIds) } elseif ($args_.entryId) { $ids = @([string]$args_.entryId) }
      if ($ids.Count -eq 0) { throw 'entryId or entryIds is required.' }
      if ($ids.Count -gt 500) { throw "refusing to touch $($ids.Count) items in one call (max 500)." }
      $n = 0
      $failed = @()
      foreach ($id in $ids) {
        try {
          $it = $ns.GetItemFromID([string]$id)
          $it.UnRead = (-not $read)
          $it.Save()
          $n++
        } catch {
          $failed += [ordered]@{ entryId = [string]$id; error = $_.Exception.Message }
        }
      }
      $out.updated = $n
      $out.read = $read
      $out.failed = $failed
    }
    default {
      throw "Unknown action: $action"
    }
  }

  $out.status = 'ok'
} catch {
  $out.status = 'error'
  $out.error = $_.Exception.Message
}

Out-Json $out
