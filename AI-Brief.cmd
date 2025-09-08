@echo off
setlocal
set "PS=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
set "ROOT=%~dp0"

"%PS%" -NoLogo -NoProfile -Command ^
  "$repo = '%ROOT%';" ^
  "$exp  = Join-Path $repo 'export-onglets-csv';" ^
  "if (-not (Test-Path -LiteralPath $exp)) { Write-Host 'Aucun dossier export-onglets-csv.'; exit 1 }" ^
  "$snap = Get-ChildItem -LiteralPath $exp -Directory | Where-Object { $_.Name -like 'SNAPSHOT_*' } | Sort-Object LastWriteTime -Descending | Select-Object -First 1;" ^
  "if (-not $snap) { Write-Host 'Aucun snapshot trouvé.'; exit 1 }" ^
  "$brief = Join-Path $snap.FullName 'docs\00_SESSION_BRIEF.md';" ^
  "if (-not (Test-Path -LiteralPath $brief)) { Write-Host 'Brief introuvable. Lance Snapshot.cmd puis réessayez.'; exit 1 }" ^
  "Set-Clipboard (Get-Content -Raw -LiteralPath $brief);" ^
  "Start-Process notepad $brief;" ^
  "Write-Host ('Brief copié dans le presse-papiers : ' + $brief)"
