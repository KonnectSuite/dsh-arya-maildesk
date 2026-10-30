# Guards against the class of bug that made every send fail silently on
# 2026-10-30: Connect-Outlook returned the MAPI namespace but the Application
# object it also created stayed local to that function, so `$outlook.CreateItem`
# was a method call on $null. Reads kept working (they only need $ns), which is
# why nothing caught it until a person tried to send.
#
# These are static checks over the bridge source: cheap, and they fail on the
# edit that would reintroduce the bug rather than on the send that follows it.

$ErrorActionPreference = 'Stop'
$bridge = Join-Path (Split-Path -Parent $PSScriptRoot) 'lib/olk.ps1'
if (-not (Test-Path $bridge)) { Write-Host "missing $bridge"; exit 1 }
$raw = Get-Content -LiteralPath $bridge -Raw -Encoding UTF8

$fail = 0
function Check($label, $ok, $detail) {
  if ($ok) { Write-Host ('ok   ' + $label) }
  else { Write-Host ('FAIL ' + $label + $(if ($detail) { ' :: ' + $detail } else { '' })); $script:fail++ }
}

# 1. The application object is parked where the action bodies can see it.
Check 'Connect-Outlook parks the Application object at script scope' ($raw -match '\$script:outlookApp\s*=\s*\$outlook') 'missing $script:outlookApp assignment'

# 2. Nothing calls a method on $outlook outside that function: the variable does
#    not exist there, and PowerShell reports it as a null method call.
$fnStart = $raw.IndexOf('function Connect-Outlook')
$fnEnd = $raw.IndexOf('function Invoke-OlkAction')
Check 'the bridge still has its two functions' ($fnStart -ge 0 -and $fnEnd -gt $fnStart) 'Connect-Outlook / Invoke-OlkAction not found as expected'
$outside = if ($fnStart -ge 0 -and $fnEnd -gt $fnStart) { $raw.Substring(0, $fnStart) + $raw.Substring($fnEnd) } else { $raw }
$uses = [regex]::Matches($outside, '\$outlook\.')
Check 'no $outlook.<method> calls remain in the action bodies' ($uses.Count -eq 0) ("found " + $uses.Count + " use(s)")

# 3. Every CreateItem goes through the parked reference.
$creates = [regex]::Matches($raw, 'CreateItem\(')
$good = [regex]::Matches($raw, '\$script:outlookApp\.CreateItem\(')
Check 'every CreateItem uses the parked Application object' ($creates.Count -eq $good.Count) ("CreateItem=" + $creates.Count + " via parked object=" + $good.Count)

# 4. The send paths still refuse to send without the host's one-time token.
Check 'send_draft refuses without OLK_CONFIRM' ($raw -match "(?s)'send_draft'\s*\{.*?SEND BLOCKED") 'send_draft gate missing'
Check 'compose_send refuses without OLK_CONFIRM' ($raw -match "(?s)'compose_send'\s*\{.*?SEND BLOCKED") 'compose_send gate missing'

if ($fail -gt 0) { Write-Host ('FAILED: ' + $fail); exit 1 }
Write-Host 'OK: bridge source invariants'
exit 0
