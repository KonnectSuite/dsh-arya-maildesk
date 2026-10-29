# Runs every *.test.ps1 in this folder, each in its own powershell process so a
# failing file cannot take the runner down with it. Requires Windows PowerShell
# 5.1 and no Outlook: the tests drive the real case bodies against a fake COM model.
$files = @(Get-ChildItem -LiteralPath $PSScriptRoot -Filter '*.test.ps1' | Sort-Object Name)
if ($files.Count -eq 0) { Write-Host 'no test scripts found'; exit 1 }
$fail = 0
foreach ($f in $files) {
  Write-Host ('=== ' + $f.Name)
  & powershell -NoProfile -ExecutionPolicy Bypass -File $f.FullName
  if ($LASTEXITCODE -ne 0) { $fail++ }
}
Write-Host ''
if ($fail -gt 0) { Write-Host ('FAILED: ' + $fail + ' of ' + $files.Count + ' test scripts'); exit 1 }
Write-Host ('OK: ' + $files.Count + ' test scripts passed')