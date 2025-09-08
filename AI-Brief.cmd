@echo off
setlocal
rem --- PowerShell 7 si dispo, sinon Windows PowerShell ---
set "ROOT=%~dp0"
set "PS7=%ProgramFiles%\PowerShell\7\pwsh.exe"
if exist "%PS7%" (set "PS=%PS7%") else set "PS=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"

"%PS%" -NoLogo -NoProfile -ExecutionPolicy Bypass -Command ^
  "$ErrorActionPreference='Stop';" ^
  "$repo=(Resolve-Path -LiteralPath '%ROOT%').Path;" ^
  "$exp=Join-Path $repo 'export-onglets-csv'; if(-not(Test-Path $exp)){throw 'export-onglets-csv introuvable'}" ^
  "$snap=Get-ChildItem $exp -Directory ^| ? Name -like 'SNAPSHOT_*' ^| Sort LastWriteTime -Descending ^| Select -First 1; if(-not $snap){throw 'Aucun snapshot'}" ^
  "$docs=Join-Path $snap.FullName 'docs'; New-Item -ItemType Directory -Force -Path $docs ^| Out-Null" ^
  "$brief=Join-Path $docs '00_SESSION_BRIEF.md'; $man=Join-Path $snap.FullName 'manifest.json'" ^
  "$resume='_manifest.json indisponible pour ce snapshot._'; if(Test-Path $man){$m=Get-Content $man -Raw ^| ConvertFrom-Json; $resume='- **Fichiers total** : {0}`n- **Taille totale (octets)** : {1}' -f $m.summary.counts.total,$m.summary.totalSize}" ^
  "$content='# BRIEF SESSION — à coller au début de la conversation`r`n`r`n'+" ^
  "'> **Snapshot** : '+$snap.Name+'  **Généré** : '+(Get-Date -Format ''yyyy-MM-dd HH:mm'')+\"`r`n\"+" ^
  "'> **Chemins utiles** :`r`n> - docs/etat/etat_projet.md`r`n> - diff.md`r`n> - docs/ai/*.md`r`n> - scripts_*.txt`r`n`r`n'+" ^
  "'## Résumé rapide`r`n'+$resume+\"`r`n\";" ^
  "$utf8NoBom=New-Object System.Text.UTF8Encoding($false); [IO.File]::WriteAllText($brief,$content,$utf8NoBom);" ^
  "Get-Content -Raw -LiteralPath $brief -Encoding UTF8 ^| Set-Clipboard;" ^
  "Start-Process notepad $brief;" ^
  "Write-Host ('[BRIEF] Copié et ouvert : '+$brief)"

