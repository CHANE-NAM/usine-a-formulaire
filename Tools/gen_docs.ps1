param(
  [Parameter(Mandatory=$true)] [string]$RepoRoot,
  [Parameter(Mandatory=$true)] [string]$SnapshotDir,
  [Parameter(Mandatory=$true)] [string]$ExportDir,
  [Parameter(Mandatory=$false)] [string]$Timestamp
)

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

function New-Dir([string]$p) { New-Item -ItemType Directory -Force -Path $p | Out-Null }

function Write-FileUtf8([string]$Path,[string]$Content) {
  $enc = New-Object System.Text.UTF8Encoding($false)
  [System.IO.File]::WriteAllText($Path, $Content, $enc)
}

function Guess-CodeFence([string]$filePath) {
  $ext = [System.IO.Path]::GetExtension($filePath).ToLowerInvariant()
  switch ($ext) {
    ".js"   { "javascript" }
    ".gs"   { "javascript" }
    ".ts"   { "typescript" }
    ".html" { "html" }
    ".json" { "json" }
    default { "" }
  }
}

# === Dossiers de sortie ===
$DocsRoot = Join-Path $SnapshotDir "docs"
$DirAI    = Join-Path $DocsRoot "ai"
$DirEtat  = Join-Path $DocsRoot "etat"
$DirUsers = Join-Path $DocsRoot "users"
New-Dir $DocsRoot; New-Dir $DirAI; New-Dir $DirEtat; New-Dir $DirUsers

# === 1) AI-friendly ===
$concatFiles = Get-ChildItem -LiteralPath $SnapshotDir -File -Filter 'scripts__*.txt' -ErrorAction SilentlyContinue
foreach ($cf in $concatFiles) {
  $projectName = [System.IO.Path]::GetFileNameWithoutExtension($cf.Name) -replace '^scripts__',''
  $outPath = Join-Path $DirAI ($projectName + ".md")

  $whole = Get-Content -LiteralPath $cf.FullName -Raw -Encoding UTF8
  $parts = [System.Text.RegularExpressions.Regex]::Split($whole, '(?m)^\s*---\s*FILE:\s*(.+?)\s*---\s*$')

  $sb = New-Object System.Text.StringBuilder
  [void]$sb.AppendLine("# " + $projectName)
  [void]$sb.AppendLine("")
  [void]$sb.AppendLine("> Genere automatiquement depuis **" + $cf.Name + "** — snapshot: **" + ([System.IO.Path]::GetFileName($SnapshotDir)) + "**.")
  [void]$sb.AppendLine("")

  for ($i=1; $i -lt $parts.Count; $i+=2) {
    $filePath = $parts[$i].Trim()
    $content  = if ($i + 1 -lt $parts.Count) { $parts[$i+1] } else { "" }
    $lang = Guess-CodeFence $filePath

    [void]$sb.AppendLine("## " + $filePath)
    [void]$sb.AppendLine("")
    [void]$sb.AppendLine("```" + $lang)
    [void]$sb.Append($content)
    [void]$sb.AppendLine("```")
    [void]$sb.AppendLine("")
  }

  $csvs = Get-ChildItem -LiteralPath $SnapshotDir -Recurse -File -Filter '*.csv' -ErrorAction SilentlyContinue
  if ($csvs.Count -gt 0) {
    [void]$sb.AppendLine("---")
    [void]$sb.AppendLine("")
    [void]$sb.AppendLine("### Fichiers CSV exportes (aperçu)")
    $preview = $csvs | Select-Object -First 30
    foreach ($c in $preview) {
      $rel = $c.FullName.Replace($SnapshotDir, '').TrimStart('\')
      [void]$sb.AppendLine("* " + $rel)
    }
    if ($csvs.Count -gt 30) {
      $more = $csvs.Count - 30
      [void]$sb.AppendLine(("* ... ({0} de plus)" -f $more))
    }
    [void]$sb.AppendLine("")
  }

  Write-FileUtf8 $outPath $sb.ToString()
}

# === 2) Etat du projet ===
$EtatPath = Join-Path $DirEtat "etat_projet.md"

$manifestPath = Join-Path $SnapshotDir "manifest.json"
$man = $null
if (Test-Path -LiteralPath $manifestPath) {
  try { $man = Get-Content -LiteralPath $manifestPath -Raw -Encoding UTF8 | ConvertFrom-Json } catch {}
}

$gitLog = @()
try { Push-Location -LiteralPath $RepoRoot; $gitLog = git log --oneline -n 12 2>$null } finally { Pop-Location }

$aiFiles = Get-ChildItem -LiteralPath $DirAI -File -Filter '*.md' -ErrorAction SilentlyContinue | Sort-Object Name
$csvCount = (Get-ChildItem -LiteralPath $SnapshotDir -Recurse -File -Filter '*.csv' -ErrorAction SilentlyContinue).Count

$sbEtat = New-Object System.Text.StringBuilder
[void]$sbEtat.AppendLine("# Etat du projet — " + (Split-Path -Leaf $SnapshotDir))
[void]$sbEtat.AppendLine("")
[void]$sbEtat.AppendLine("- **Genere** : " + (Get-Date).ToString("yyyy-MM-dd HH:mm:ss"))
[void]$sbEtat.AppendLine("- **Snapshot** : " + (Split-Path -Leaf $SnapshotDir))
[void]$sbEtat.AppendLine("- **CSV exportes** : " + $csvCount)
[void]$sbEtat.AppendLine("- **Racine repo** : " + $RepoRoot)
[void]$sbEtat.AppendLine("")
[void]$sbEtat.AppendLine("## Resume (manifest)")

if ($man) {
  $filesTotal = $man.summary.filesTotal
  $totalBytes = $man.summary.totalBytes
  [void]$sbEtat.AppendLine("- **fichiersTotal** : " + $filesTotal)
  [void]$sbEtat.AppendLine("- **tailleTotale** : " + $totalBytes + " octets")
  [void]$sbEtat.AppendLine("- **par type** :")
  if ($man.summary.counts.byType) {
    $props = $man.summary.counts.byType.PSObject.Properties
    foreach ($p in $props) { [void]$sbEtat.AppendLine("  - **" + $p.Name + "** : " + $p.Value) }
  } else {
    [void]$sbEtat.AppendLine("  - (non disponible)")
  }
} else {
  [void]$sbEtat.AppendLine("_Aucun manifest.json disponible pour ce snapshot._")
}

[void]$sbEtat.AppendLine("")
[void]$sbEtat.AppendLine("## Derniers commits")
if ($gitLog -and $gitLog.Count) {
  foreach ($l in $gitLog) { [void]$sbEtat.AppendLine("* " + $l) }
} else {
  [void]$sbEtat.AppendLine("_(git log indisponible ou vide)_")
}

[void]$sbEtat.AppendLine("")
[void]$sbEtat.AppendLine("## Index documents AI-friendly")
if ($aiFiles.Count -gt 0) {
  foreach ($f in $aiFiles) { [void]$sbEtat.AppendLine("* [" + $f.BaseName + "](" + $f.Name + ")") }
} else {
  [void]$sbEtat.AppendLine("_Aucun document AI genere (pas de scripts__*.txt trouves)._")
}

[void]$sbEtat.AppendLine("")
[void]$sbEtat.AppendLine("## Fichiers utiles dans le snapshot")
[void]$sbEtat.AppendLine("- `manifest.json` : " + [System.IO.File]::Exists($manifestPath))
[void]$sbEtat.AppendLine("- `diff.md` : " + [System.IO.File]::Exists((Join-Path $SnapshotDir "diff.md")))
[void]$sbEtat.AppendLine("- `brief.md` : " + [System.IO.File]::Exists((Join-Path $SnapshotDir "brief.md")))
[void]$sbEtat.AppendLine("- `zip` : " + [System.IO.File]::Exists((Join-Path $ExportDir ((Split-Path -Leaf $SnapshotDir) + ".zip"))))

Write-FileUtf8 $EtatPath $sbEtat.ToString()

# === 3) Doc utilisateurs ===
function Make-Tree([string]$root) {
  $sb = New-Object System.Text.StringBuilder
  if (-not (Test-Path -LiteralPath $root)) { return "" }
  $rootLen = $root.Length
  $items = Get-ChildItem -LiteralPath $root -Recurse -File -ErrorAction SilentlyContinue
  foreach ($it in $items) {
    $rel = $it.FullName.Substring($rootLen).TrimStart('\')
    [void]$sb.AppendLine("* " + $rel)
  }
  $sb.ToString()
}

$gasRoots = @(
  Join-Path $RepoRoot "01_Moteur",
  Join-Path $RepoRoot "02_configuration",
  Join-Path $RepoRoot "03_BaseDeDonnées",
  Join-Path $RepoRoot "03_BaseDeDonnees",
  Join-Path $RepoRoot "04_Templates"
) | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -Unique

$treeSb = New-Object System.Text.StringBuilder
foreach ($gr in $gasRoots) {
  $title = Split-Path -Leaf $gr
  [void]$treeSb.AppendLine("# " + $title)
  [void]$treeSb.AppendLine((Make-Tree $gr))
  [void]$treeSb.AppendLine("")
}
$treeMd = $treeSb.ToString()

$GuidePath = Join-Path $DirUsers "guide_utilisateur.md"
$guideSb = New-Object System.Text.StringBuilder
[void]$guideSb.AppendLine("# Guide utilisateur — Usine a Formulaire (snapshot " + (Split-Path -Leaf $SnapshotDir) + ")")
[void]$guideSb.AppendLine("")
[void]$guideSb.AppendLine("## Objet")
[void]$guideSb.AppendLine("Ce document explique comment utiliser l'application (creation/edition de formulaires, executions, recuperation des resultats).")
[void]$guideSb.AppendLine("")
[void]$guideSb.AppendLine("## Pre-requis")
[void]$guideSb.AppendLine("- Compte Google avec acces aux classeurs et Apps Script.")
[void]$guideSb.AppendLine("- Autorisations OAuth accordees (Drive/Sheets lecture).")
[void]$guideSb.AppendLine("")
[void]$guideSb.AppendLine("## Demarrage rapide")
[void]$guideSb.AppendLine("1. Ouvrir le classeur [CONFIG] V2 Usine a Tests.")
[void]$guideSb.AppendLine("2. Lancer le menu Usine > Generer ...")
[void]$guideSb.AppendLine("3. Verifier la generation cote [TEMPLATE] / [MOTEUR] si applicable.")
[void]$guideSb.AppendLine("")
[void]$guideSb.AppendLine("## Points d'attention")
[void]$guideSb.AppendLine("- Les exports CSV sont disponibles dans export-onglets-csv\" + (Split-Path -Leaf $SnapshotDir) + "\ .")
[void]$guideSb.AppendLine("- Les documents AI-friendly sont dans docs\ai\ .")
Write-FileUtf8 $GuidePath $guideSb.ToString()

$ArchPath = Join-Path $DirUsers "architecture.md"
$archSb = New-Object System.Text.StringBuilder
[void]$archSb.AppendLine("# Architecture — Vue d'ensemble")
[void]$archSb.AppendLine("")
[void]$archSb.AppendLine("## Modules Google Apps Script par domaine")
[void]$archSb.AppendLine("- 01_Moteur : moteur d'orchestration, traitement, envoi d'emails, etc.")
[void]$archSb.AppendLine("- 02_configuration : UI de configuration, validations, conversion, etc.")
[void]$archSb.AppendLine("- 03_BaseDeDonnees : logique partagee ou wrappers (selon projet).")
[void]$archSb.AppendLine("- 04_Templates : moteurs de scoring / rendu / PDF / scenarios.")
[void]$archSb.AppendLine("")
[void]$archSb.AppendLine("## Arborescence GAS (extrait)")
[void]$archSb.AppendLine($treeMd)
Write-FileUtf8 $ArchPath $archSb.ToString()

$GlossPath = Join-Path $DirUsers "glossaire.md"
$glossSb = New-Object System.Text.StringBuilder
[void]$glossSb.AppendLine("# Glossaire (squelette)")
[void]$glossSb.AppendLine("- Formulaire : ...")
[void]$glossSb.AppendLine("- Profil : ...")
[void]$glossSb.AppendLine("- Scenario : ...")
[void]$glossSb.AppendLine("- Gabarit : ...")
Write-FileUtf8 $GlossPath $glossSb.ToString()

$IndexPath = Join-Path $DocsRoot "README.md"
$indexSb = New-Object System.Text.StringBuilder
[void]$indexSb.AppendLine("# Paquet de documentation (snapshot " + (Split-Path -Leaf $SnapshotDir) + ")")
[void]$indexSb.AppendLine("")
[void]$indexSb.AppendLine("- AI-friendly : docs/ai/ — 1 fichier par projet (code segmente par source).")
[void]$indexSb.AppendLine("- Etat du projet : docs/etat/etat_projet.md")
[void]$indexSb.AppendLine("- Doc utilisateurs : docs/users/guide_utilisateur.md, architecture.md, glossaire.md")
[void]$indexSb.AppendLine("")
[void]$indexSb.AppendLine("Ces fichiers sont regeneres a chaque snapshot.")
Write-FileUtf8 $IndexPath $indexSb.ToString()

Write-Host "[DOCS] Generation terminee : $DocsRoot"
