# Tools\AI-Brief.ps1 — brief riche (ASCII-safe)
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# ====== Reglages ======
$MaxCommits   = 8
$MaxDiffItems = 10
$MaxAiList    = 8
# ======================

# Racine du projet (ce script est dans Tools\)
$Root   = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$Export = Join-Path $Root 'export-onglets-csv'
if (-not (Test-Path -LiteralPath $Export)) { throw "export-onglets-csv introuvable: $Export" }

# Dernier snapshot
$snap = Get-ChildItem -LiteralPath $Export -Directory |
        Where-Object { $_.Name -like 'SNAPSHOT_*' } |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1
if (-not $snap) { throw "Aucun snapshot trouve dans $Export" }

# Dossiers/fichiers utiles
$DocsDir   = Join-Path $snap.FullName 'docs'
New-Item -ItemType Directory -Force -Path $DocsDir | Out-Null
$BriefPath = Join-Path $DocsDir '00_SESSION_BRIEF.md'
$ManPath   = Join-Path $snap.FullName 'manifest.json'
$DiffPath  = Join-Path $snap.FullName 'diff.md'
$AiDir     = Join-Path $DocsDir 'ai'

# ====== Manifest: resume + par type ======
$Resume = '_manifest.json indisponible pour ce snapshot._'
$ByTypeLines = @()
if (Test-Path -LiteralPath $ManPath) {
  try {
    $m = Get-Content -LiteralPath $ManPath -Raw | ConvertFrom-Json
    $total = [string]$m.summary.counts.total
    $size  = [string]$m.summary.totalSize
    $Resume = "- **Fichiers total** : $total`n- **Taille totale (octets)** : $size"

    if ($m.summary.counts.byType) {
      $props = $m.summary.counts.byType.PSObject.Properties | Sort-Object Name
      foreach ($p in $props) { $ByTypeLines += ('  - **{0}** : {1}' -f $p.Name, $p.Value) }
    }
  } catch { }
}

# ====== Compte CSV ======
$CsvCount = (Get-ChildItem -LiteralPath $snap.FullName -Recurse -File -Filter '*.csv' -ErrorAction SilentlyContinue).Count

# ====== Diff condense (Ajouts/Suppressions/Modifications) ======
$TopAdded=@(); $TopRemoved=@(); $TopChanged=@()
if (Test-Path -LiteralPath $DiffPath) {
  $d = Get-Content -LiteralPath $DiffPath -Raw -Encoding UTF8
  $curr = ''
  foreach ($line in ($d -split "`r?`n")) {
    if ($line -match '^\#\#\s+(Ajouts|Suppressions|Modifications)\b') { $curr = $matches[1]; continue }
    if ($line -match '^\*\s+') {
      switch ($curr) {
        'Ajouts'        { $TopAdded   += $line }
        'Suppressions'  { $TopRemoved += $line }
        'Modifications' { $TopChanged += $line }
      }
    }
  }
  $TopAdded   = $TopAdded   | Select-Object -First $MaxDiffItems
  $TopRemoved = $TopRemoved | Select-Object -First $MaxDiffItems
  $TopChanged = $TopChanged | Select-Object -First $MaxDiffItems
}

# ====== Liste courte des docs AI ======
$AiList = @()
if (Test-Path -LiteralPath $AiDir) {
  $files = Get-ChildItem -LiteralPath $AiDir -File -Filter *.md -ErrorAction SilentlyContinue | Sort-Object Name
  foreach ($f in ($files | Select-Object -First $MaxAiList)) { $AiList += ('- docs/ai/' + $f.Name) }
}

# ====== Commits recents (optionnel) ======
$Commits = @()
try {
  Push-Location -LiteralPath $Root
  $Commits = git log --oneline -n $MaxCommits 2>$null
} catch {} finally { Pop-Location }

# ====== Construction du brief (ASCII uniquement) ======
$content = New-Object System.Text.StringBuilder
[void]$content.AppendLine('# BRIEF SESSION - a coller au debut de la conversation')
[void]$content.AppendLine('')
[void]$content.AppendLine('> **Snapshot** : ' + $snap.Name + '  **Genere** : ' + (Get-Date -Format 'yyyy-MM-dd HH:mm'))
[void]$content.AppendLine('> **Chemins utiles** :')
[void]$content.AppendLine('> - docs/etat/etat_projet.md')
[void]$content.AppendLine('> - diff.md')
[void]$content.AppendLine('> - docs/ai/*.md')
[void]$content.AppendLine('> - scripts_*.txt')
[void]$content.AppendLine('')
[void]$content.AppendLine('## Resume rapide')
[void]$content.AppendLine($Resume)
if ($ByTypeLines.Count) {
  [void]$content.AppendLine('')
  [void]$content.AppendLine('- **Par type** :')
  foreach ($l in $ByTypeLines) { [void]$content.AppendLine($l) }
}
[void]$content.AppendLine('')
[void]$content.AppendLine('- **CSV exportes** : ' + $CsvCount)
[void]$content.AppendLine('')
[void]$content.AppendLine('## Commits recents')
if ($Commits -and $Commits.Count) { foreach ($c in $Commits) { [void]$content.AppendLine('* ' + $c) } }
else { [void]$content.AppendLine('_(git log indisponible)_') }

[void]$content.AppendLine('')
[void]$content.AppendLine('## Changements cles (diff condense)')
if ($TopAdded.Count)   { [void]$content.AppendLine('**Ajouts**');         foreach ($l in $TopAdded)   { [void]$content.AppendLine($l) } }
if ($TopRemoved.Count) { [void]$content.AppendLine(''); [void]$content.AppendLine('**Suppressions**'); foreach ($l in $TopRemoved) { [void]$content.AppendLine($l) } }
if ($TopChanged.Count) { [void]$content.AppendLine(''); [void]$content.AppendLine('**Modifications**'); foreach ($l in $TopChanged) { [void]$content.AppendLine($l) } }
if (-not ($TopAdded.Count -or $TopRemoved.Count -or $TopChanged.Count)) { [void]$content.AppendLine('_Aucun diff detecte._') }

[void]$content.AppendLine('')
[void]$content.AppendLine('## Docs AI a me fournir si besoin')
if ($AiList.Count) { foreach ($l in $AiList) { [void]$content.AppendLine($l) } } else { [void]$content.AppendLine('- (aucun fichier dans docs/ai)') }

# Ecrit en UTF-8 sans BOM
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
[IO.File]::WriteAllText($BriefPath, $content.ToString(), $utf8NoBom)

# Copie dans presse-papiers si dispo et ouvre Notepad
if (Get-Command Set-Clipboard -ErrorAction SilentlyContinue) {
  Get-Content -LiteralPath $BriefPath -Raw | Set-Clipboard
}
Start-Process notepad $BriefPath
Write-Host "[BRIEF] Genere, copie et ouvert : $BriefPath"
