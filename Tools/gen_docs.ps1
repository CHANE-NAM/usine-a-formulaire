# Tools\gen_docs.ps1 — génère docs AI/État/Users + brief IA, appels nommés/positionnels/auto
#requires -Version 5.1
[CmdletBinding()]
param(
  [string]$RepoRoot,
  [string]$SnapshotDir,
  [string]$ExportDir,
  [string]$Timestamp
)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# 0) Déductions si manquants (rend l'appel ultra tolérant)
$ScriptRoot = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }

if (-not $RepoRoot -or -not (Test-Path -LiteralPath $RepoRoot)) {
  $RepoRoot = (Resolve-Path (Join-Path $ScriptRoot '..')).Path
}
if (-not $ExportDir -or -not (Test-Path -LiteralPath $ExportDir)) {
  $ExportDir = Join-Path $RepoRoot 'export-onglets-csv'
}
if (-not (Test-Path -LiteralPath $ExportDir)) {
  throw "export-onglets-csv introuvable : $ExportDir"
}
if (-not $SnapshotDir -or -not (Test-Path -LiteralPath $SnapshotDir)) {
  $snap = Get-ChildItem -LiteralPath $ExportDir -Directory |
          Where-Object { $_.Name -like 'SNAPSHOT_*' } |
          Sort-Object LastWriteTime -Descending |
          Select-Object -First 1
  if (-not $snap) { throw "Aucun snapshot trouvé dans $ExportDir" }
  $SnapshotDir = $snap.FullName
}

# Helpers
function New-Dir([string]$p) { New-Item -ItemType Directory -Force -Path $p | Out-Null }
function Write-FileUtf8([string]$Path,[string]$Content) {
  $enc = New-Object System.Text.UTF8Encoding($false)  # UTF-8 sans BOM
  [System.IO.File]::WriteAllText($Path, $Content, $enc)
}
function Guess-CodeFence([string]$filePath) {
  switch ([System.IO.Path]::GetExtension($filePath).ToLowerInvariant()) {
    '.js'   { 'javascript' }
    '.gs'   { 'javascript' }
    '.ts'   { 'typescript' }
    '.html' { 'html' }
    '.json' { 'json' }
    default { '' }
  }
}
function Make-Tree([string]$root) {
  $sb = New-Object System.Text.StringBuilder
  if (-not (Test-Path -LiteralPath $root)) { return '' }
  $rootLen = $root.Length
  $items = Get-ChildItem -LiteralPath $root -Recurse -File -ErrorAction SilentlyContinue
  foreach ($it in $items) {
    $rel = $it.FullName.Substring($rootLen).TrimStart('\')
    [void]$sb.AppendLine('* ' + $rel)
  }
  $sb.ToString()
}

# 1) Dossiers de sortie
$DocsRoot = Join-Path $SnapshotDir 'docs'
$DirAI    = Join-Path $DocsRoot 'ai'
$DirEtat  = Join-Path $DocsRoot 'etat'
$DirUsers = Join-Path $DocsRoot 'users'
New-Dir $DocsRoot; New-Dir $DirAI; New-Dir $DirEtat; New-Dir $DirUsers

# 2) AI-friendly (scripts_*.txt -> docs/ai/*.md)
$concatFiles = Get-ChildItem -LiteralPath $SnapshotDir -File -Filter 'scripts_*.txt' -ErrorAction SilentlyContinue
foreach ($cf in $concatFiles) {
  $projectName = ([System.IO.Path]::GetFileNameWithoutExtension($cf.Name) -replace '^scripts_','')
  $outPath = Join-Path $DirAI ($projectName + '.md')

  $whole = Get-Content -LiteralPath $cf.FullName -Raw -Encoding UTF8
  if (-not $whole) { continue }

  $parts = [System.Text.RegularExpressions.Regex]::Split($whole, '(?m)^\s*---\s*FILE:\s*(.+?)\s*---\s*$')

  $sb = New-Object System.Text.StringBuilder
  [void]$sb.AppendLine('# ' + $projectName)
  [void]$sb.AppendLine('')
  [void]$sb.AppendLine('> Généré automatiquement depuis **' + $cf.Name + '** — snapshot: **' + (Split-Path -Leaf $SnapshotDir) + '**.')
  [void]$sb.AppendLine('')

  for ($i=1; $i -lt $parts.Count; $i+=2) {
    $filePath = ($parts[$i] -as [string]).Trim()
    $content  = if ($i + 1 -lt $parts.Count) { $parts[$i+1] } else { '' }
    $lang = Guess-CodeFence $filePath

    [void]$sb.AppendLine('## ' + $filePath)
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('```' + $lang)
    [void]$sb.Append($content)
    [void]$sb.AppendLine('```')
    [void]$sb.AppendLine('')
  }

  $csvs = Get-ChildItem -LiteralPath $SnapshotDir -Recurse -File -Filter '*.csv' -ErrorAction SilentlyContinue
  if ($csvs.Count -gt 0) {
    [void]$sb.AppendLine('---')
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('### Fichiers CSV exportés (aperçu)')
    foreach ($c in ($csvs | Select-Object -First 30)) {
      $rel = $c.FullName.Replace($SnapshotDir, '').TrimStart('\')
      [void]$sb.AppendLine('* ' + $rel)
    }
    if ($csvs.Count -gt 30) { [void]$sb.AppendLine(('* ... ({0} de plus)' -f ($csvs.Count - 30))) }
    [void]$sb.AppendLine('')
  }

  Write-FileUtf8 $outPath $sb.ToString()
}

# 3) État du projet
$EtatPath     = Join-Path $DirEtat 'etat_projet.md'
$manifestPath = Join-Path $SnapshotDir 'manifest.json'
$man = $null
if (Test-Path -LiteralPath $manifestPath) {
  try { $man = Get-Content -LiteralPath $manifestPath -Raw -Encoding UTF8 | ConvertFrom-Json } catch {}
}

$gitLog = @()
try { Push-Location -LiteralPath $RepoRoot; $gitLog = git log --oneline -n 12 2>$null } finally { Pop-Location }

$aiFiles  = Get-ChildItem -LiteralPath $DirAI -File -Filter '*.md' -ErrorAction SilentlyContinue | Sort-Object Name
$csvCount = (Get-ChildItem -LiteralPath $SnapshotDir -Recurse -File -Filter '*.csv' -ErrorAction SilentlyContinue).Count

$sbEtat = New-Object System.Text.StringBuilder
[void]$sbEtat.AppendLine('# État du projet — ' + (Split-Path -Leaf $SnapshotDir))
[void]$sbEtat.AppendLine('')
[void]$sbEtat.AppendLine('- **Généré** : ' + (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))
[void]$sbEtat.AppendLine('- **Snapshot** : ' + (Split-Path -Leaf $SnapshotDir))
[void]$sbEtat.AppendLine('- **CSV exportés** : ' + $csvCount)
[void]$sbEtat.AppendLine('- **Racine repo** : ' + $RepoRoot)
[void]$sbEtat.AppendLine('')
[void]$sbEtat.AppendLine('## Résumé (manifest)')
if ($man) {
  [void]$sbEtat.AppendLine('- **fichiersTotal** : ' + $man.summary.counts.total)
  [void]$sbEtat.AppendLine('- **tailleTotale** : ' + $man.summary.totalSize + ' octets')
  [void]$sbEtat.AppendLine('- **par type** :')
  if ($man.summary.counts.byType) {
    $props = $man.summary.counts.byType.PSObject.Properties
    foreach ($p in $props) { [void]$sbEtat.AppendLine('  - **' + $p.Name + '** : ' + $p.Value) }
  } else { [void]$sbEtat.AppendLine('  - (non disponible)') }
} else {
  [void]$sbEtat.AppendLine('_Aucun manifest.json disponible pour ce snapshot._')
}

[void]$sbEtat.AppendLine('')
[void]$sbEtat.AppendLine('## Derniers commits')
if ($gitLog -and $gitLog.Count) { foreach ($l in $gitLog) { [void]$sbEtat.AppendLine('* ' + $l) } }
else { [void]$sbEtat.AppendLine('_(git log indisponible ou vide)_') }

[void]$sbEtat.AppendLine('')
[void]$sbEtat.AppendLine('## Index documents AI-friendly')
if ($aiFiles.Count -gt 0) { foreach ($f in $aiFiles) { [void]$sbEtat.AppendLine('* [' + $f.BaseName + '](' + $f.Name + ')') } }
else { [void]$sbEtat.AppendLine('_Aucun document AI généré (pas de scripts_*.txt trouvés)._') }

[void]$sbEtat.AppendLine('')
[void]$sbEtat.AppendLine('## Fichiers utiles dans le snapshot')
[void]$sbEtat.AppendLine('- `manifest.json` : ' + [System.IO.File]::Exists($manifestPath))
[void]$sbEtat.AppendLine('- `diff.md` : ' + [System.IO.File]::Exists((Join-Path $SnapshotDir 'diff.md')))
[void]$sbEtat.AppendLine('- `brief.md` : ' + [System.IO.File]::Exists((Join-Path $SnapshotDir 'brief.md')))
[void]$sbEtat.AppendLine('- `zip` : ' + [System.IO.File]::Exists((Join-Path $ExportDir ((Split-Path -Leaf $SnapshotDir) + '.zip'))))
Write-FileUtf8 $EtatPath $sbEtat.ToString()

# 4) Brief minimal pour IA (00_SESSION_BRIEF.md)
$DocsBrief = Join-Path $DocsRoot '00_SESSION_BRIEF.md'

# Résumé manifest (par type)
$countsByTypeLines = @()
if ($man -and $man.summary -and $man.summary.counts -and $man.summary.counts.byType) {
  $props = $man.summary.counts.byType.PSObject.Properties | Sort-Object Name
  foreach ($p in ($props | Select-Object -First 8)) {
    $countsByTypeLines += ('  - **{0}** : {1}' -f $p.Name, $p.Value)
  }
}

# Diff condensé
$diffPath = Join-Path $SnapshotDir 'diff.md'
$topAdded=@(); $topRemoved=@(); $topChanged=@()
if (Test-Path -LiteralPath $diffPath) {
  $d = Get-Content -LiteralPath $diffPath -Raw -Encoding UTF8
  $curr = ''
  foreach ($line in ($d -split "`r?`n")) {
    if ($line -match '^\#\#\s+(Ajouts|Suppressions|Modifications)\b') { $curr = $matches[1]; continue }
    if ($line -match '^\*\s+') {
      switch ($curr) {
        'Ajouts'        { $topAdded   += $line }
        'Suppressions'  { $topRemoved += $line }
        'Modifications' { $topChanged += $line }
      }
    }
  }
  $topAdded   = $topAdded   | Select-Object -First 10
  $topRemoved = $topRemoved | Select-Object -First 10
  $topChanged = $topChanged | Select-Object -First 10
}

# Liste courte des fichiers AI
$aiIndexLines = @()
foreach ($f in ($aiFiles | Select-Object -First 6)) { $aiIndexLines += ('- ' + $f.Name) }

# Construction du brief
$sbBrief = New-Object System.Text.StringBuilder
[void]$sbBrief.AppendLine('# BRIEF SESSION — à coller au début de la conversation')
[void]$sbBrief.AppendLine('')
[void]$sbBrief.AppendLine('> **Snapshot** : ' + (Split-Path -Leaf $SnapshotDir) + '  **Généré** : ' + (Get-Date).ToString('yyyy-MM-dd HH:mm'))
[void]$sbBrief.AppendLine('> **Chemins utiles** :')
[void]$sbBrief.AppendLine('> - docs/etat/etat_projet.md')
[void]$sbBrief.AppendLine('> - diff.md')
[void]$sbBrief.AppendLine('> - docs/ai/*.md')
[void]$sbBrief.AppendLine('> - scripts_*.txt')
[void]$sbBrief.AppendLine('')
[void]$sbBrief.AppendLine('## Résumé rapide')
if ($man) {
  [void]$sbBrief.AppendLine( ("- **{0}** : {1}" -f 'Fichiers total', $man.summary.counts.total) )
  [void]$sbBrief.AppendLine( ("- **{0}** : {1}" -f 'Taille totale (octets)', $man.summary.totalSize) )
  if ($countsByTypeLines.Count) {
    [void]$sbBrief.AppendLine('- **Par type** :')
    foreach ($l in $countsByTypeLines) { [void]$sbBrief.AppendLine($l) }
  }
} else {
  [void]$sbBrief.AppendLine('_manifest.json indisponible pour ce snapshot._')
}
[void]$sbBrief.AppendLine('')
[void]$sbBrief.AppendLine('## Commits récents')
if ($gitLog -and $gitLog.Count) { foreach ($l in ($gitLog | Select-Object -First 8)) { [void]$sbBrief.AppendLine('* ' + $l) } }
else { [void]$sbBrief.AppendLine('_(git log indisponible)_') }
[void]$sbBrief.AppendLine('')
[void]$sbBrief.AppendLine('## Changements clés (diff condensé)')
if ($topAdded.Count)   { [void]$sbBrief.AppendLine('**Ajouts**');         foreach ($l in $topAdded)   { [void]$sbBrief.AppendLine($l) } }
if ($topRemoved.Count) { [void]$sbBrief.AppendLine(''); [void]$sbBrief.AppendLine('**Suppressions**'); foreach ($l in $topRemoved) { [void]$sbBrief.AppendLine($l) } }
if ($topChanged.Count) { [void]$sbBrief.AppendLine(''); [void]$sbBrief.AppendLine('**Modifications**'); foreach ($l in $topChanged) { [void]$sbBrief.AppendLine($l) } }
if (-not ($topAdded.Count -or $topRemoved.Count -or $topChanged.Count)) { [void]$sbBrief.AppendLine('_Aucun diff détecté._') }
[void]$sbBrief.AppendLine('')
[void]$sbBrief.AppendLine('## Docs à me demander au besoin (pointeurs)')
if ($aiIndexLines.Count) { foreach ($l in $aiIndexLines) { [void]$sbBrief.AppendLine($l) } } else { [void]$sbBrief.AppendLine('- (Aucun fichier dans docs/ai)') }
[void]$sbBrief.AppendLine('')
[void]$sbBrief.AppendLine('---')
[void]$sbBrief.AppendLine('### Prompt suggéré (copier/coller après le brief)')
[void]$sbBrief.AppendLine('> Lis le brief ci-dessus. Propose-moi les points d''attention et liste les documents complémentaires qu''il te faudrait (nom exact dans ce snapshot). Dis-moi dans quel ordre les lire. Puis pose tes questions de clarification.')
[void]$sbBrief.AppendLine('')
Write-FileUtf8 $DocsBrief $sbBrief.ToString()
# === Bloc correctif pour générer manifest.json, brief.md, diff.md ===
. (Join-Path $ScriptRoot 'snapshot_helpers.ps1')

$manifestPath = Write-Manifest -SnapshotDir $SnapshotDir -RepoRoot $RepoRoot
$briefPath = Write-BriefMd -SnapshotDir $SnapshotDir -Manifest $manifestPath

# Diff avec le snapshot précédent
$snapDirs = Get-ChildItem -LiteralPath $ExportDir -Directory | Where-Object { $_.Name -like 'SNAPSHOT_*' } | Sort-Object Name
$prevSnap = $null
foreach ($d in $snapDirs) {
  if ($d.FullName -eq $SnapshotDir) { break }
  $prevSnap = $d
}
if ($prevSnap) {
  $prevManifest = Join-Path $prevSnap.FullName 'manifest.json'
  if (Test-Path $prevManifest) {
    $diffPath = Join-Path $SnapshotDir 'diff.md'
    Write-DiffMd -PrevManifestPath $prevManifest -CurrManifestPath $manifestPath -OutPath $diffPath
  }
}

Write-Host '[DOCS] Génération terminée : ' $DocsRoot
exit 0
