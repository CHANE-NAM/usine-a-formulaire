# Tools\AI-Brief.ps1 — brief enrichi + IA_kit
[CmdletBinding()]param()
$ErrorActionPreference='Stop'; [Console]::OutputEncoding=[System.Text.Encoding]::UTF8
$ScriptRoot = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
$Repo=(Resolve-Path (Join-Path $ScriptRoot '..')).Path
$ExportDir=Join-Path $Repo 'export-onglets-csv'
if (-not (Test-Path -LiteralPath $ExportDir)) { throw "export-onglets-csv introuvable : $ExportDir" }
$Snap = Get-ChildItem -LiteralPath $ExportDir -Directory | Where-Object { $_.Name -like 'SNAPSHOT_*' } | Sort-Object LastWriteTime -Descending | Select-Object -First 1
if (-not $Snap) { throw "Aucun snapshot trouvé dans $ExportDir" }
$DocsDir=Join-Path $Snap.FullName 'docs'; New-Item -ItemType Directory -Force -Path $DocsDir | Out-Null
$BriefPath=Join-Path $DocsDir '00_SESSION_BRIEF.md'
$ManPath=Join-Path $Snap.FullName 'manifest.json'
$DiffPath=Join-Path $Snap.FullName 'diff.md'

# Résumé manifest
$resume = '_manifest.json indisponible pour ce snapshot._'
if (Test-Path $ManPath) {
  try {
    $m = Get-Content -LiteralPath $ManPath -Raw -Encoding UTF8 | ConvertFrom-Json
    $byType=''; if ($m.summary.counts.byType) {
      $props=$m.summary.counts.byType.PSObject.Properties | Sort-Object Name
      $lines = $props | Select-Object -First 8 | ForEach-Object { "  - **$($_.Name)** : $($_.Value)" }
      if ($lines) { $byType = "- **Par type** :`r`n" + ($lines -join "`r`n") }
    }
    $resume = "- **Fichiers total** : $($m.summary.counts.total)`r`n- **Taille totale (octets)** : $($m.summary.totalSize)"
    if ($byType) { $resume += "`r`n$byType" }
  } catch { $resume = "_Échec lecture manifest.json : $($_.Exception.Message)_" }
}

# Extrait du diff
$diffBlock=''
if (Test-Path $DiffPath) {
  $d = Get-Content -LiteralPath $DiffPath -Raw -Encoding UTF8
  $curr=''; $add=@(); $rem=@(); $chg=@()
  foreach($line in ($d -split "`r?`n")){
    if ($line -match '^\#\#\s+(Ajouts|Suppressions|Modifications)\b'){ $curr=$Matches[1]; continue }
    if ($line -match '^\*\s+') { switch ($curr) { 'Ajouts'{$add+=$line}; 'Suppressions'{$rem+=$line}; 'Modifications'{$chg+=$line} } }
  }
  $add=$add | Select-Object -First 8; $rem=$rem | Select-Object -First 8; $chg=$chg | Select-Object -First 8
  if ($add.Count -or $rem.Count -or $chg.Count) {
    $sb = New-Object System.Text.StringBuilder
    if ($add.Count){ [void]$sb.AppendLine('**Ajouts**');         $add | ForEach-Object { [void]$sb.AppendLine($_) } }
    if ($rem.Count){ [void]$sb.AppendLine(''); [void]$sb.AppendLine('**Suppressions**');  $rem | ForEach-Object { [void]$sb.AppendLine($_) } }
    if ($chg.Count){ [void]$sb.AppendLine(''); [void]$sb.AppendLine('**Modifications**'); $chg | ForEach-Object { [void]$sb.AppendLine($_) } }
    $diffBlock = $sb.ToString()
  }
}

# IA_kit (appel optionnel)
$KitDir = Join-Path $Snap.FullName 'IA_kit'
$KitScript1 = Join-Path $Repo 'Build-IA-Kit.ps1'
$KitScript2 = Join-Path $Repo 'Tools\Build-IA-Kit.ps1'
$kitBuilt=$false
try {
  if (Test-Path $KitScript1) { & $KitScript1; $kitBuilt=$true }
  elseif (Test-Path $KitScript2) { & $KitScript2; $kitBuilt=$true }
} catch { Write-Warning ("Build-IA-Kit.ps1 a échoué : {0}" -f $_.Exception.Message) }
if (-not $kitBuilt) {
  New-Item -ItemType Directory -Force -Path $KitDir | Out-Null
  $toCopy=@()
  $EtatPath = Join-Path $DocsDir 'etat\etat_projet.md'
  if (Test-Path $BriefPath) { $toCopy += $BriefPath }
  if (Test-Path $EtatPath)  { $toCopy += $EtatPath }
  if (Test-Path $DiffPath)  { $toCopy += $DiffPath }
  foreach($f in $toCopy){ Copy-Item -LiteralPath $f -Destination $KitDir -Force }
}

# Liste IA_kit
$kitList=''; if (Test-Path $KitDir) {
  $files = Get-ChildItem -LiteralPath $KitDir -File | Select-Object -Expand Name
  if ($files) { $kitList = '- ' + ($files -join "`r`n- ") }
}

# Brief enrichi
$content  = "# BRIEF SESSION — à coller au début de la conversation`r`n`r`n"
$content += "> **Snapshot** : $($Snap.Name)  **Généré** : $(Get-Date -Format 'yyyy-MM-dd HH:mm')`r`n"
$content += "> **Chemins utiles** :`r`n> - docs/etat/etat_projet.md`r`n> - diff.md`r`n> - docs/ai/*.md`r`n> - scripts_*.txt`r`n`r`n"
$content += "## Résumé rapide`r`n$resume`r`n`r`n"
if ($diffBlock) { $content += "## Changements clés (extrait du diff)`r`n$diffBlock`r`n`r`n" }
if ($kitList)   { $content += "## Fichiers joints (IA_kit)`r`n$kitList`r`n`r`n" }

$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
[System.IO.File]::WriteAllText($BriefPath,$content,$utf8NoBom)

try { Get-Content -Raw -LiteralPath $BriefPath -Encoding UTF8 | Set-Clipboard } catch { Write-Warning "Set-Clipboard indisponible : $($_.Exception.Message)" }
Start-Process notepad $BriefPath
if (Test-Path $KitDir) { Start-Process explorer $KitDir }
Write-Host "[BRIEF] Généré, copié et ouvert : $BriefPath"