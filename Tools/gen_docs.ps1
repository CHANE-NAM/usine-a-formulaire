# Tools\gen_docs.ps1 — génère docs AI/État/Users + brief IA
# Utilisation :
#   1) Sans paramètre (auto-détection) :  powershell -NoLogo -ExecutionPolicy Bypass -File .\Tools\gen_docs.ps1
#   2) Avec paramètres (nommés ou positionnels) : -RepoRoot <repo> -SnapshotDir <snap> -ExportDir <export>

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

# ---------- Helpers sûrs ----------
function As-Scalar([object]$v) {
  if ($null -eq $v) { return '' }
  if ($v -is [System.Array]) { foreach ($e in $v) { if ($e) { return [string]$e } }; return '' }
  return [string]$v
}
function SafeJoin {
  param([Parameter(Mandatory=$true)]$Base,[Parameter(Mandatory=$true)]$Child)
  $b = As-Scalar $Base
  if ($Child -is [System.Array]) {
    $acc = $b; foreach ($seg in $Child) { $acc = [System.IO.Path]::Combine($acc, [string]$seg) }; return $acc
  } else {
    return (Join-Path -Path $b -ChildPath (As-Scalar $Child))
  }
}
function New-Dir([string]$p) { New-Item -ItemType Directory -Force -Path $p | Out-Null }
function Write-FileUtf8([string]$Path,[string]$Content) {
  $enc = New-Object System.Text.UTF8Encoding($false)  # UTF-8 sans BOM
  [System.IO.File]::WriteAllText((As-Scalar $Path), $Content, $enc)
}
function Guess-CodeFence([string]$filePath) {
  switch ([System.IO.Path]::GetExtension($filePath).ToLowerInvariant()) {
    '.js' { 'javascript' } '.gs' { 'javascript' } '.ts' { 'typescript' }
    '.html' { 'html' } '.json' { 'json' } default { '' }
  }
}
function Make-Tree([string]$root) {
  $root = As-Scalar $root
  $sb = New-Object System.Text.StringBuilder
  if (-not (Test-Path -LiteralPath $root)) { return '' }
  $rootLen = $root.Length
  $items = Get-ChildItem -LiteralPath $root -Recurse -File -ErrorAction SilentlyContinue
  foreach ($it in $items) { $rel = $it.FullName.Substring($rootLen).TrimStart('\'); [void]$sb.AppendLine('* ' + $rel) }
  $sb.ToString()
}

# ---------- Auto-détection des chemins si non fournis ----------
# 1) Essaie les $args nus (positionnels) si le binder n’a pas rempli les variables
if (-not $RepoRoot -or -not $SnapshotDir -or -not $ExportDir) {
  if ($args.Count -ge 3) {
    if (-not $RepoRoot)    { $RepoRoot    = [string]$args[0] }
    if (-not $SnapshotDir) { $SnapshotDir = [string]$args[1] }
    if (-not $ExportDir)   { $ExportDir   = [string]$args[2] }
    if (-not $Timestamp -and $args.Count -ge 4) { $Timestamp = [string]$args[3] }
  }
}

# 2) Si toujours vide, auto-détecte par rapport à l’emplacement du script (Tools\… -> repo = ..)
$ScriptRoot = if ($PSScriptRoot) { $PSScriptRoot }
elseif ($MyInvocation.MyCommand.Path) { Split-Path -Parent $MyInvocation.MyCommand.Path }
else { (Get-Location).Path }

if (-not $RepoRoot)    { $RepoRoot  = (Resolve-Path (Join-Path $ScriptRoot '..')).Path }
if (-not $ExportDir)   { $ExportDir = SafeJoin $RepoRoot 'export-onglets-csv' }
if (-not $SnapshotDir) {
  $Latest = (Get-ChildItem -LiteralPath $ExportDir -Directory |
             Where-Object { $_.Name -like 'SNAPSHOT_*' } |
             Sort-Object LastWriteTime -Descending | Select-Object -First 1)
  if ($Latest) { $SnapshotDir = $Latest.FullName }
}

# 3) Dernier filet : variables d’environnement
if (-not $RepoRoot -and $env:REPO_ROOT)      { $RepoRoot = $env:REPO_ROOT }
if (-not $ExportDir -and $env:EXPORT_DIR)    { $ExportDir = $env:EXPORT_DIR }
if (-not $SnapshotDir -and $env:SNAPSHOT_DIR){ $SnapshotDir = $env:SNAPSHOT_DIR }

# Normalise en chaînes simples
$RepoRoot    = As-Scalar $RepoRoot
$SnapshotDir = As-Scalar $SnapshotDir
$ExportDir   = As-Scalar $ExportDir
$Timestamp   = As-Scalar $Timestamp

if ([string]::IsNullOrWhiteSpace($RepoRoot) -or
    [string]::IsNullOrWhiteSpace($ExportDir) -or
    [string]::IsNullOrWhiteSpace($SnapshotDir)) {
  throw "Chemins insuffisants. RepoRoot='$RepoRoot'  ExportDir='$ExportDir'  SnapshotDir='$SnapshotDir'"
}

# ---------- Dossiers de sortie ----------
$DocsRoot = SafeJoin $SnapshotDir 'docs'
$DirAI    = SafeJoin $DocsRoot 'ai'
$DirEtat  = SafeJoin $DocsRoot 'etat'
$DirUsers = SafeJoin $DocsRoot 'users'
New-Dir $DocsRoot; New-Dir $DirAI; New-Dir $DirEtat; New-Dir $DirUsers

# ========================= 1) AI-friendly (scripts_*.txt) =========================
$concatFiles = Get-ChildItem -LiteralPath $SnapshotDir -File -Filter 'scripts_*.txt' -ErrorAction SilentlyContinue
foreach ($cf in $concatFiles) {
  $projectName = ([System.IO.Path]::GetFileNameWithoutExtension($cf.Name) -replace '^scripts_','')
  $outPath = SafeJoin $DirAI ($projectName + '.md')

  $whole = Get-Content -LiteralPath $cf.FullName -Raw -Encoding UTF8
  if ([string]::IsNullOrEmpty($whole)) { continue }

  $parts = [System.Text.RegularExpressions.Regex]::Split($whole, '(?m)^\s*---\s*FILE:\s*(.+?)\s*---\s*$')

  $sb = New-Object System.Text.StringBuilder
  [void]$sb.AppendLine('# ' + $projectName)
  [void]$sb.AppendLine('')
  [void]$sb.AppendLine('> Généré automatiquement depuis **' + $cf.Name + '** — snapshot: **' + (Split-Path -Leaf $SnapshotDir) + '**.')
  [void]$sb.AppendLine('')

  for ($i=1; $i -lt $parts.Count; $i+=2) {
    $filePath = (As-Scalar $parts[$i]).Trim()
    $content  = if ($i + 1 -lt $parts.Count) { As-Scalar $parts[$i+1] } else { '' }
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
    [void]$sb.AppendLine('---'); [void]$sb.AppendLine('')
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

# ========================= 2) État du projet =========================
$EtatPath     = SafeJoin $DirEtat 'etat_projet.md'
$manifestPath = SafeJoin $SnapshotDir 'manifest.json'
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
[void]$sbEtat.AppendLine('- `diff.md` : ' + [System.IO.File]::Exists((SafeJoin $SnapshotDir 'diff.md')))
[void]$sbEtat.AppendLine('- `brief.md` : ' + [System.IO.File]::Exists((SafeJoin $SnapshotDir 'brief.md')))
[void]$sbEtat.AppendLine('- `zip` : ' + [System.IO.File]::Exists((SafeJoin $ExportDir ((Split-Path -Leaf $SnapshotDir) + '.zip'))))
Write-FileUtf8 $EtatPath $sbEtat.ToString()

# ========================= 3) Docs utilisateurs =========================
$gasRoots = @(
  SafeJoin $RepoRoot '01_Moteur',
  SafeJoin $RepoRoot '02_configuration',
  SafeJoin $RepoRoot '03_BaseDeDonnées',
  SafeJoin $RepoRoot '03_BaseDeDonnees',
  SafeJoin $RepoRoot '04_Templates'
) | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -Unique

$treeSb = New-Object System.Text.StringBuilder
foreach ($gr in $gasRoots) {
  $title = Split-Path -Leaf $gr
  [void]$treeSb.AppendLine('# ' + $title)
  [void]$treeSb.AppendLine((Make-Tree $gr))
  [void]$treeSb.AppendLine('')
}
$treeMd = $treeSb.ToString()

$GuidePath = SafeJoin $DirUsers 'guide_utilisateur.md'
$guideSb = New-Object System.Text.StringBuilder
[void]$guideSb.AppendLine('# Guide utilisateur — Usine à Formulaire (snapshot ' + (Split-Path -Leaf $SnapshotDir) + ')')
[void]$guideSb.AppendLine('')
[void]$guideSb.AppendLine('## Objet')
[void]$guideSb.AppendLine('Ce document explique comment utiliser l''application (création/édition de formulaires, exécutions, récupération des résultats).')
[void]$guideSb.AppendLine('')
[void]$guideSb.AppendLine('## Pré-requis')
[void]$guideSb.AppendLine('- Compte Google avec accès aux classeurs et Apps Script.')
[void]$guideSb.AppendLine('- Autorisations OAuth accordées (Drive/Sheets lecture).')
[void]$guideSb.AppendLine('')
[void]$guideSb.AppendLine('## Démarrage rapide')
[void]$guideSb.AppendLine('1. Ouvrir le classeur [CONFIG] V2 Usine à Tests.')
[void]$guideSb.AppendLine('2. Lancer le menu Usine > Générer ...')
[void]$guideSb.AppendLine('3. Vérifier la génération côté [TEMPLATE] / [MOTEUR] si applicable.')
[void]$guideSb.AppendLine('')
[void]$guideSb.AppendLine('## Points d''attention')
[void]$guideSb.AppendLine('- Les exports CSV sont disponibles dans `export-onglets-csv\' + (Split-Path -Leaf $SnapshotDir) + '\`.')
[void]$guideSb.AppendLine('- Les documents AI-friendly sont dans `docs\ai\`.')
Write-FileUtf8 $GuidePath $guideSb.ToString()

$ArchPath = SafeJoin $DirUsers 'architecture.md'
$archSb = New-Object System.Text.StringBuilder
[void]$archSb.AppendLine('# Architecture — Vue d''ensemble')
[void]$archSb.AppendLine('')
[void]$archSb.AppendLine('## Modules Google Apps Script par domaine')
[void]$archSb.AppendLine('- 01_Moteur : moteur d''orchestration, traitement, envoi d''emails, etc.')
[void]$archSb.AppendLine('- 02_configuration : UI de configuration, validations, conversion, etc.')
[void]$archSb.AppendLine('- 03_BaseDeDonnees : logique partagée ou wrappers (selon projet).')
[void]$archSb.AppendLine('- 04_Templates : moteurs de scoring / rendu / PDF / scénarios.')
[void]$archSb.AppendLine('')
[void]$archSb.AppendLine('## Arborescence GAS (extrait)')
[void]$archSb.AppendLine($treeMd)
Write-FileUtf8 $ArchPath $archSb.ToString()

$GlossPath = SafeJoin $DirUsers 'glossaire.md'
$glossSb = New-Object System.Text.StringBuilder
[void]$glossSb.AppendLine('# Glossaire (squelette)')
[void]$glossSb.AppendLine('- Formulaire : ...')
[void]$glossSb.AppendLine('- Profil : ...')
[void]$glossSb.AppendLine('- Scénario : ...')
[void]$glossSb.AppendLine('- Gabarit : ...')
Write-FileUtf8 $GlossPath $glossSb.ToString()

$IndexPath = SafeJoin $DocsRoot 'README.md'
$indexSb = New-Object System.Text.StringBuilder
[void]$indexSb.AppendLine('# Paquet de documentation (snapshot ' + (Split-Path -Leaf $SnapshotDir) + ')')
[void]$indexSb.AppendLine('')
[void]$indexSb.AppendLine('- AI-friendly : docs/ai/ — 1 fichier par projet (code segmenté par source).')
[void]$indexSb.AppendLine('- État du projet : docs/etat/etat_projet.md')
[void]$indexSb.AppendLine('- Doc utilisateurs : docs/users/guide_utilisateur.md, architecture.md, glossaire.md')
[void]$indexSb.AppendLine('')
[void]$indexSb.AppendLine('Ces fichiers sont régénérés à chaque snapshot.')
Write-FileUtf8 $IndexPath $indexSb.ToString()

# ========================= 4) Brief minimal pour IA =========================
$BriefPath = SafeJoin $DocsRoot '00_SESSION_BRIEF.md'

$countsByTypeLines = @()
if ($man -and $man.summary -and $man.summary.counts -and $man.summary.counts.byType) {
  $props = $man.summary.counts.byType.PSObject.Properties | Sort-Object Name
  foreach ($p in ($props | Select-Object -First 8)) { $countsByTypeLines += ('  - **{0}** : {1}' -f $p.Name, $p.Value) }
}

$diffPath = SafeJoin $SnapshotDir 'diff.md'
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

$aiIndexLines = @()
$aiMdFiles = Get-ChildItem -LiteralPath $DirAI -File -Filter '*.md' -ErrorAction SilentlyContinue | Sort-Object Name
foreach ($f in ($aiMdFiles | Select-Object -First 6)) { $aiIndexLines += ('- ' + $f.Name) }

$sbBrief = New-Object System.Text.StringBuilder
[void]$sbBrief.AppendLine('# BRIEF SESSION — à coller au début de la conversation')
[void]$sbBrief.AppendLine('')
[void]$sbBrief.AppendLine('> **Snapshot** : ' + (Split-Path -Leaf $SnapshotDir) + ' — **Généré** : ' + (Get-Date).ToString('yyyy-MM-dd HH:mm'))
[void]$sbBrief.AppendLine('> **Chemins utiles** :')
[void]$sbBrief.AppendLine('> - docs/etat/etat_projet.md (commits, métriques)')
[void]$sbBrief.AppendLine('> - diff.md (ajouts/suppressions/modifications)')
[void]$sbBrief.AppendLine('> - docs/ai/*.md (code par projet)')
[void]$sbBrief.AppendLine('> - scripts_*.txt (concat code brut)')
[void]$sbBrief.AppendLine('')
[void]$sbBrief.AppendLine('## Résumé rapide')
if ($man) {
  [void]$sbBrief.AppendLine( ("- **{0}** : {1}" -f 'Fichiers total', $man.summary.counts.total) )
  [void]$sbBrief.AppendLine( ("- **{0}** : {1}" -f 'Taille totale (octets)', $man.summary.totalSize) )
  if ($countsByTypeLines.Count) { [void]$sbBrief.AppendLine('- **Par type** :'); foreach ($l in $countsByTypeLines) { [void]$sbBrief.AppendLine($l) } }
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
if ($aiIndexLines.Count) { foreach ($l in $aiIndexLines) { [void]$sbBrief.AppendLine($l) } }
else { [void]$sbBrief.AppendLine('- (Aucun fichier dans docs/ai)') }
[void]$sbBrief.AppendLine('')
[void]$sbBrief.AppendLine('---')
[void]$sbBrief.AppendLine('### Prompt suggéré (copier/coller après le brief)')
[void]$sbBrief.AppendLine('> Lis le brief ci-dessus. Propose-moi les points d''attention et liste les documents complémentaires qu''il te faudrait (nom exact dans ce snapshot). Dis-moi dans quel ordre les lire. Puis pose tes questions de clarification.')
[void]$sbBrief.AppendLine('')
Write-FileUtf8 $BriefPath $sbBrief.ToString()

Write-Host '[DOCS] Génération terminée : ' $DocsRoot
exit 0
